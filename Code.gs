const SPREADSHEET_ID = '1I8Nanf7nlJyp_-Y1CsZs9yhsQKFHL_wg5WEfC1mDTe0';
const APP_NAME = 'myPI KSSM';
const TINGKATAN_AKTIF = [1, 2, 3, 4, 5];

function assertTingkatan_(tingkatan) {
  const value = Number(tingkatan);
  if (!TINGKATAN_AKTIF.includes(value)) {
    throw new Error('Pilih Tingkatan 1 hingga 5.');
  }
  return value;
}

function getPortalUser_() {
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email) return { email: '', nama: '', peranan: '', status: '' };

  const sh = getSS_().getSheetByName('Pengguna');
  if (!sh || sh.getLastRow() < 2) return { email: email, nama: '', peranan: '', status: '' };
  const row = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getDisplayValues()
    .find(r => String(r[0] || '').trim().toLowerCase() === email);
  return row ? {
    email: email,
    nama: String(row[1] || '').trim(),
    peranan: String(row[2] || '').trim(),
    status: String(row[3] || '').trim()
  } : { email: email, nama: '', peranan: '', status: '' };
}

function assertGuru_() {
  const user = getPortalUser_();
  if (!user.email) {
    throw new Error('Sila buka portal menggunakan akaun Google sekolah yang telah didaftarkan.');
  }
  if (user.status !== 'Aktif' || !['Admin', 'Guru'].includes(user.peranan)) {
    throw new Error('Akses guru diperlukan. Minta pentadbir menambah emel anda dalam tab Pengguna.');
  }
  return user;
}

function assertAdmin_() {
  const user = assertGuru_();
  if (user.peranan !== 'Admin') {
    throw new Error('Akses pentadbir diperlukan.');
  }
  return user;
}

function logAudit_(tindakan, butiran, status) {
  const sh = getSS_().getSheetByName('Audit_Log');
  if (!sh) return;
  const user = getPortalUser_();
  sh.appendRow([
    Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Kuala_Lumpur', 'yyyy-MM-dd HH:mm:ss'),
    user.email || 'sistem', user.peranan || 'Sistem', tindakan,
    String(butiran || '').slice(0, 1000), status || 'Berjaya'
  ]);
}

function getSS_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSetting_(key, fallback) {
  const sh = getSS_().getSheetByName('Tetapan');
  const values = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 2).getValues();
  const found = values.find(r => String(r[0]).trim() === key);
  return found && found[1] !== '' ? String(found[1]).trim() : fallback;
}

function getGeminiApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('GEMINI_API_KEY belum ditetapkan di Script Properties.');
  return key;
}

function senaraiPilihanJanaSoalan() {
  assertGuru_();
  const ss = getSS_();
  const out = {};
  TINGKATAN_AKTIF.forEach(t => {
    const sh = ss.getSheetByName('DSKP_T' + t);
    const rows = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 9).getValues()
      .filter(r => r[0] && String(r[7]).trim() === 'Disemak');
    out[t] = rows.map(r => ({
      id: r[0],
      tingkatan: r[1],
      bidang: r[2],
      tajuk: r[3],
      sk: r[4],
      sp: r[5],
      sumber: r[6]
    }));
  });

  return {
    dskp: out,
    tahap: getSetting_('TAHAP_SOALAN', 'Mudah,Sederhana,Tinggi').split(',').map(s => s.trim()),
    jenis: getSetting_('JENIS_SOALAN', 'Objektif,Struktur,Esei').split(',').map(s => s.trim()),
    model: getSetting_('GEMINI_MODEL', 'gemini-3.8-flash')
  };
}

function cariDSKP_(tingkatan, bidang, tajuk) {
  const sh = getSS_().getSheetByName('DSKP_T' + Number(tingkatan));
  if (!sh) throw new Error('Tab DSKP tidak dijumpai.');

  const rows = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 9).getValues();
  const row = rows.find(r =>
    String(r[2]).trim().toLowerCase() === String(bidang).trim().toLowerCase() &&
    String(r[3]).trim().toLowerCase() === String(tajuk).trim().toLowerCase() &&
    String(r[7]).trim() === 'Disemak'
  );

  if (!row) {
    throw new Error('Tajuk DSKP yang telah disemak tidak dijumpai.');
  }

  return {
    id: row[0],
    tingkatan: row[1],
    bidang: row[2],
    tajuk: row[3],
    sk: row[4],
    sp: row[5],
    sumber: row[6]
  };
}

function janaSoalanGemini(input) {
  assertGuru_();
  input = input || {};

  const tingkatan = assertTingkatan_(input.tingkatan);
  const bidang = String(input.bidang || '').trim();
  const tajuk = String(input.tajuk || '').trim();
  const tahap = String(input.tahap || '').trim();
  const jenis = String(input.jenis || '').trim();
  const bilangan = Math.max(1, Math.min(Number(input.bilangan || 5), 20));

  if (!bidang || !tajuk || !tahap || !jenis) throw new Error('Lengkapkan semua pilihan janaan.');

  const dskp = cariDSKP_(tingkatan, bidang, tajuk);
  const model = getSetting_('GEMINI_MODEL', 'gemini-3.8-flash');
  const apiKey = getGeminiApiKey_();

  const prompt = [
    'Anda ialah pembina item Pendidikan Islam KSSM Malaysia bagi Tingkatan 1 hingga 5.',
    'Gunakan HANYA konteks DSKP yang diberikan. Jangan cipta fakta agama di luar konteks.',
    '',
    'TINGKATAN: ' + tingkatan,
    'BIDANG: ' + bidang,
    'TAJUK: ' + tajuk,
    'TAHAP: ' + tahap,
    'JENIS: ' + jenis,
    'BILANGAN: ' + bilangan,
    '',
    'STANDARD KANDUNGAN:',
    dskp.sk,
    '',
    'STANDARD PEMBELAJARAN:',
    dskp.sp,
    '',
    'Keperluan:',
    '- Bahasa Melayu baku, jelas dan sesuai murid Tingkatan ' + tingkatan + '.',
    '- Pastikan aras ' + tahap + ' benar-benar tercermin pada item.',
    '- Untuk Objektif: sediakan 4 pilihan A-D, satu jawapan betul sahaja.',
    '- Untuk Struktur/Esei: biarkan pilihan A-D sebagai string kosong.',
    '- Jawapan mesti padan dengan soalan.',
    '- Huraian jawapan ringkas tetapi membantu guru menyemak.',
    '- Elakkan item kabur, terlalu umum atau bergantung pada fakta yang tiada dalam SK/SP.',
    '- Jangan keluarkan markdown. Pulangkan JSON sahaja.'
  ].join('\n');

  const schema = {
    type: 'object',
    properties: {
      soalan: {
        type: 'array',
        minItems: bilangan,
        maxItems: bilangan,
        items: {
          type: 'object',
          properties: {
            soalan: { type: 'string' },
            pilihanA: { type: 'string' },
            pilihanB: { type: 'string' },
            pilihanC: { type: 'string' },
            pilihanD: { type: 'string' },
            jawapan: { type: 'string' },
            huraian: { type: 'string' }
          },
          required: ['soalan','pilihanA','pilihanB','pilihanC','pilihanD','jawapan','huraian']
        }
      }
    },
    required: ['soalan']
  };

  const parsed = callGeminiStructuredInteraction_(model, apiKey, prompt, schema);

  if (!parsed.soalan || !Array.isArray(parsed.soalan)) {
    throw new Error('Format jawapan Gemini tidak sah.');
  }

  return {
    meta: {
      model: model,
      dskpId: dskp.id,
      tingkatan: tingkatan,
      bidang: bidang,
      tajuk: tajuk,
      tahap: tahap,
      jenis: jenis,
      sumber: dskp.sumber
    },
    soalan: parsed.soalan
  };
}

function extractGeminiText_(json) {
  const candidates = json && json.candidates;
  if (!candidates || !candidates.length) {
    throw new Error('Gemini tidak memulangkan calon jawapan.');
  }
  const parts = candidates[0].content && candidates[0].content.parts;
  if (!parts || !parts.length) throw new Error('Respons Gemini kosong.');

  const text = parts.map(p => p.text || '').join('').trim();
  if (!text) throw new Error('Teks respons Gemini kosong.');
  return text;
}

function simpanSoalanKeBank(payload) {
  assertGuru_();
  if (!payload || !payload.meta || !Array.isArray(payload.soalan)) {
    throw new Error('Payload soalan tidak sah.');
  }

  const sh = getSS_().getSheetByName('Bank_Soalan');
  const meta = payload.meta;
  const now = new Date();
  const tz = Session.getScriptTimeZone() || 'Asia/Kuala_Lumpur';
  const stamp = Utilities.formatDate(now, tz, 'yyyyMMdd-HHmmss');

  const rows = payload.soalan.map((q, i) => {
    const id = 'Q-' + stamp + '-' + String(i + 1).padStart(2, '0');
    return [
      id,
      Number(meta.tingkatan),
      meta.bidang,
      meta.tajuk,
      meta.tahap,
      meta.jenis,
      q.soalan || '',
      q.pilihanA || '',
      q.pilihanB || '',
      q.pilihanC || '',
      q.pilihanD || '',
      q.jawapan || '',
      q.huraian || '',
      meta.sumber || '',
      'Belum Disemak',
      defaultMarkahPenuh_(meta.jenis)
    ];
  });

  if (!rows.length) return { saved: 0 };

  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  logAudit_('JANA_SOALAN', rows.length + ' soalan baharu disimpan ke Bank_Soalan.', 'Berjaya');
  return { saved: rows.length, ids: rows.map(r => r[0]) };
}

function janaDanSimpanSoalan(input) {
  assertGuru_();
  const hasil = janaSoalanGemini(input);
  const simpan = simpanSoalanKeBank(hasil);
  return { hasil: hasil, simpan: simpan };
}



/* =========================
   MODUL KUIZ myPI KSSM
   ========================= */

function getApprovedQuestionOptions() {
  assertGuru_();
  const sh = getSS_().getSheetByName('Bank_Soalan');
  const last = sh.getLastRow();
  if (last < 2) return [];

  return sh.getRange(2, 1, last - 1, 16).getValues()
    .filter(r => r[0] && String(r[14]).trim() === 'Lulus')
    .map(r => ({
      id: String(r[0]),
      tingkatan: Number(r[1]),
      bidang: String(r[2] || ''),
      tajuk: String(r[3] || ''),
      tahap: String(r[4] || ''),
      jenis: String(r[5] || ''),
      soalan: String(r[6] || ''),
      markahPenuh: Number(r[15] || defaultMarkahPenuh_(r[5]))
    }));
}

function createQuiz(input) {
  assertGuru_();
  input = input || {};
  const tajukKuiz = String(input.tajukKuiz || '').trim();
  const tingkatan = Number(input.tingkatan);
  const bidang = String(input.bidang || '').trim();
  const tajuk = String(input.tajuk || '').trim();
  const tahap = String(input.tahap || '').trim();
  const jenis = String(input.jenis || '').trim();
  const bilangan = Math.max(1, Math.min(Number(input.bilangan || 10), 50));

  if (!tajukKuiz) throw new Error('Masukkan Tajuk Kuiz.');
  assertTingkatan_(tingkatan);

  const bank = getApprovedQuestionOptions().filter(q =>
    q.tingkatan === tingkatan &&
    (!bidang || q.bidang === bidang) &&
    (!tajuk || q.tajuk === tajuk) &&
    (!tahap || q.tahap === tahap) &&
    (!jenis || q.jenis === jenis)
  );

  if (bank.length < bilangan) {
    throw new Error('Soalan Lulus yang sepadan hanya ' + bank.length + '. Kurangkan bilangan atau longgarkan penapis.');
  }

  shuffle_(bank);
  const selected = bank.slice(0, bilangan);
  const idSoalan = selected.map(q => q.id).join(',');
  const jenisSoalan = [...new Set(selected.map(q => q.jenis))].join(', ');
  const code = uniqueQuizCode_();
  const quizId = 'K-' + Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'yyyyMMdd-HHmmss');

  const sh = getSS_().getSheetByName('Kuiz');
  const row = sh.getLastRow() + 1;
  const now = new Date();

  sh.getRange(row, 1, 1, 15).setValues([[
    quizId,
    tajukKuiz,
    tingkatan,
    bidang,
    tajuk,
    tahap || 'Campuran',
    selected.length,
    code,
    now,
    '',
    'Aktif',
    idSoalan,
    jenisSoalan,
    '',
    'Dijana daripada Bank_Soalan berstatus Lulus'
  ]]);
  logAudit_('CIPTA_KUIZ', quizId + ' · ' + tajukKuiz + ' · Tingkatan ' + tingkatan, 'Berjaya');

  return {
    idKuiz: quizId,
    kodKuiz: code,
    bilangan: selected.length,
    urlSuffix: '?page=kuiz&code=' + encodeURIComponent(code)
  };
}

function uniqueQuizCode_() {
  const sh = getSS_().getSheetByName('Kuiz');
  const last = sh.getLastRow();
  const existing = last > 1 ? new Set(sh.getRange(2, 8, last - 1, 1).getDisplayValues().flat()) : new Set();

  for (let i = 0; i < 30; i++) {
    const code = Math.random().toString(36).slice(2, 8).toUpperCase();
    if (!existing.has(code)) return code;
  }
  throw new Error('Gagal menjana kod kuiz unik.');
}

function shuffle_(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

function loadQuizByCode(code) {
  code = String(code || '').trim().toUpperCase();
  if (!code) throw new Error('Masukkan Kod Kuiz.');

  const ss = getSS_();
  const qz = ss.getSheetByName('Kuiz');
  const last = qz.getLastRow();
  if (last < 2) throw new Error('Tiada kuiz tersedia.');

  const rows = qz.getRange(2, 1, last - 1, 15).getValues();
  const row = rows.find(r => String(r[7]).trim().toUpperCase() === code);
  if (!row) throw new Error('Kod Kuiz tidak sah.');
  if (String(row[10]).trim() !== 'Aktif') throw new Error('Kuiz ini tidak aktif.');

  const now = new Date();
  if (row[8] instanceof Date && now < row[8]) throw new Error('Kuiz belum bermula.');
  if (row[9] instanceof Date && now > row[9]) throw new Error('Kuiz telah tamat.');

  const ids = String(row[11] || '').split(',').map(s => s.trim()).filter(Boolean);
  if (!ids.length) throw new Error('Kuiz belum mempunyai set soalan.');

  const bank = ss.getSheetByName('Bank_Soalan');
  const bRows = bank.getRange(2, 1, Math.max(bank.getLastRow()-1,1), 15).getValues();
  const map = new Map(bRows.filter(r => r[0]).map(r => [String(r[0]), r]));

  const soalan = ids.map((id, index) => {
    const r = map.get(id);
    if (!r) return null;
    return {
      id: id,
      no: index + 1,
      jenis: String(r[5] || ''),
      soalan: String(r[6] || ''),
      pilihan: [
        String(r[7] || ''),
        String(r[8] || ''),
        String(r[9] || ''),
        String(r[10] || '')
      ].filter(Boolean)
    };
  }).filter(Boolean);

  return {
    idKuiz: String(row[0]),
    tajukKuiz: String(row[1]),
    tingkatan: Number(row[2]),
    bidang: String(row[3] || ''),
    tajuk: String(row[4] || ''),
    tahap: String(row[5] || ''),
    bilangan: soalan.length,
    kodKuiz: code,
    soalan: soalan
  };
}

function submitQuiz(payload) {
  payload = payload || {};
  const code = String(payload.kodKuiz || '').trim().toUpperCase();
  const quiz = loadQuizByCode(code);
  const nama = String(payload.nama || '').trim();
  const idMurid = String(payload.idMurid || '').trim();
  const kelas = String(payload.kelas || '').trim();
  const answers = payload.answers || {};

  if (!nama) throw new Error('Masukkan nama murid.');
  if (!kelas) throw new Error('Masukkan kelas.');

  const ss = getSS_();
  const bank = ss.getSheetByName('Bank_Soalan');
  const rows = bank.getRange(2, 1, Math.max(bank.getLastRow()-1,1), 16).getValues();
  const map = new Map(rows.filter(r => r[0]).map(r => [String(r[0]), r]));

  const jawapanSheet = ss.getSheetByName('Jawapan_Murid');

  // Elak penghantaran berganda bagi ID murid atau nama+kelas yang sama.
  if (jawapanSheet.getLastRow() > 1) {
    const existing = jawapanSheet.getRange(2, 2, jawapanSheet.getLastRow()-1, 4).getDisplayValues();
    const dup = existing.some(r =>
      r[0] === quiz.idKuiz &&
      ((idMurid && r[1] === idMurid) ||
       (!idMurid && r[2].toLowerCase() === nama.toLowerCase() && r[3].toLowerCase() === kelas.toLowerCase()))
    );
    if (dup) throw new Error('Jawapan untuk murid ini telah dihantar bagi kuiz ini.');
  }

  let autoScore = 0;
  let autoMax = 0;
  const outRows = [];

  quiz.soalan.forEach((q, idx) => {
    const r = map.get(q.id);
    if (!r) return;

    const jenis = String(r[5] || '');
    const murid = String(answers[q.id] || '').trim();
    const betul = String(r[11] || '').trim();
    const markahPenuh = Number(r[15] || defaultMarkahPenuh_(jenis)) || defaultMarkahPenuh_(jenis);

    let markah = '';
    let status = 'Semakan Guru';

    if (jenis.toLowerCase() === 'objektif') {
      const isCorrect = checkObjective_(murid, betul, [r[7], r[8], r[9], r[10]]);
      markah = isCorrect ? markahPenuh : 0;
      status = isCorrect ? 'Betul' : 'Salah';
      autoScore += Number(markah);
      autoMax += markahPenuh;
    }

    outRows.push([
      new Date(),            // A Timestamp
      quiz.idKuiz,           // B ID Kuiz
      idMurid,               // C ID Murid
      nama,                  // D Nama Murid
      kelas,                 // E Kelas
      idx + 1,               // F No. Soalan
      murid,                 // G Jawapan Murid
      betul,                 // H Jawapan Betul
      markah,                // I Markah
      status,                // J Betul/Salah / Semakan Guru
      markahPenuh,           // K Markah Penuh
      ''                     // L Catatan Guru
    ]);
  });

  if (!outRows.length) throw new Error('Tiada jawapan diterima.');
  jawapanSheet.getRange(jawapanSheet.getLastRow()+1, 1, outRows.length, 12).setValues(outRows);

  return {
    saved: outRows.length,
    autoScore: autoScore,
    autoMax: autoMax,
    perluSemakan: outRows.filter(r => r[9] === 'Semakan Guru').length
  };
}

function checkObjective_(studentAnswer, correctAnswer, options) {
  const norm = s => String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
  const sa = norm(studentAnswer);
  const ca = norm(correctAnswer);
  if (!sa) return false;
  if (sa === ca) return true;

  const letters = ['a','b','c','d'];
  const idx = letters.indexOf(ca.replace(/[\.\)\s]/g,''));
  if (idx >= 0 && options[idx] !== undefined) {
    return sa === norm(options[idx]);
  }
  return false;
}



/* =========================
   MODUL JANA RPH
   ========================= */

function getRPHOptions() {
  assertGuru_();
  const ss = getSS_();
  const result = {};
  TINGKATAN_AKTIF.forEach(t => {
    const sh = ss.getSheetByName('RPT_T' + t);
    const last = sh.getLastRow();
    const rows = last > 1 ? sh.getRange(2,1,last-1,10).getDisplayValues() : [];
    result[t] = rows.filter(r => r[0]).map(r => ({
      id: r[0],
      tahun: r[1],
      tingkatan: Number(r[2]),
      minggu: Number(r[3]),
      tarikhMula: r[4],
      tarikhAkhir: r[5],
      bidang: r[6],
      tajuk: r[7],
      spRPT: r[8],
      catatan: r[9]
    }));
  });
  return result;
}

function findDSKPForRPH_(tingkatan, rpt) {
  const sh = getSS_().getSheetByName('DSKP_T' + tingkatan);
  const last = sh.getLastRow();
  if (last < 2) return null;

  const rows = sh.getRange(2,1,last-1,9).getDisplayValues()
    .filter(r => r[0] && ['Disemak','Lulus'].includes(String(r[7]).trim()));

  const clean = s => String(s || '')
    .toLowerCase()
    .replace(/\(sambungan\)/g,'')
    .replace(/\s*\/.*$/g,'')
    .replace(/[’‘]/g,"'")
    .replace(/\s+/g,' ')
    .trim();

  const targetTitle = clean(rpt.tajuk);
  const targetField = clean(rpt.bidang);

  let row = rows.find(r => clean(r[3]) === targetTitle);
  if (!row) {
    row = rows.find(r => clean(r[3]).includes(targetTitle) || targetTitle.includes(clean(r[3])));
  }
  if (!row && targetField && !['ulang kaji','pentaksiran','pengukuhan','pengayaan','spm','umum'].includes(targetField)) {
    const candidates = rows.filter(r => clean(r[2]) === targetField);
    if (candidates.length === 1) row = candidates[0];
  }

  if (!row) return null;
  return {
    id: row[0],
    bidang: row[2],
    tajuk: row[3],
    sk: row[4],
    sp: row[5],
    sumber: row[6],
    status: row[7]
  };
}

function extractInteractionText_(json) {
  if (!json) throw new Error('Respons Gemini kosong.');

  if (typeof json.output_text === 'string' && json.output_text.trim()) {
    return json.output_text.trim();
  }

  const steps = Array.isArray(json.steps) ? json.steps : [];
  for (let i = steps.length - 1; i >= 0; i--) {
    const step = steps[i];
    if (step && step.type === 'model_output' && Array.isArray(step.content)) {
      const parts = step.content
        .filter(x => x && x.type === 'text' && typeof x.text === 'string')
        .map(x => x.text);
      if (parts.length) return parts.join('').trim();
    }
  }

  throw new Error('Gemini tidak memulangkan output teks yang boleh dibaca.');
}

function callGeminiStructuredInteraction_(model, apiKey, prompt, schema) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/interactions';

  const payload = {
    model: model,
    input: prompt,
    response_format: {
      type: 'text',
      mime_type: 'application/json',
      schema: schema
    }
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-goog-api-key': apiKey,
      'Api-Revision': '2026-05-20'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = res.getResponseCode();
  const body = res.getContentText();
  if (status < 200 || status >= 300) {
    throw new Error('Gemini API gagal (' + status + '): ' + body.slice(0, 700));
  }

  let json;
  try {
    json = JSON.parse(body);
  } catch (e) {
    throw new Error('Respons Gemini bukan JSON API yang sah: ' + body.slice(0, 300));
  }

  const text = extractInteractionText_(json);
  try {
    return JSON.parse(text);
  } catch (e) {
    throw new Error('Output structured JSON Gemini tidak sah: ' + text.slice(0, 500));
  }
}

function generateStructuredGemini_(prompt, schema) {
  const model = getSetting_('GEMINI_MODEL', 'gemini-3.8-flash');
  const apiKey = getGeminiApiKey_();
  return callGeminiStructuredInteraction_(model, apiKey, prompt, schema);
}

function janaRPH(input) {
  assertGuru_();
  input = input || {};
  const tingkatan = assertTingkatan_(input.tingkatan);
  const rptId = String(input.rptId || '').trim();
  const kelas = String(input.kelas || '').trim();
  const tarikh = String(input.tarikh || '').trim();
  const tempoh = Math.max(30, Math.min(Number(input.tempoh || 60), 120));
  const catatanGuru = String(input.catatanGuru || '').trim();

  if (!rptId) throw new Error('Pilih minggu/tajuk RPT.');
  if (!kelas) throw new Error('Masukkan kelas.');
  if (!tarikh) throw new Error('Pilih tarikh PdP.');

  const options = getRPHOptions();
  const rpt = (options[tingkatan] || []).find(x => x.id === rptId);
  if (!rpt) throw new Error('Rekod RPT tidak dijumpai.');

  const dskp = findDSKPForRPH_(tingkatan, rpt);
  const sk = dskp ? dskp.sk : '';
  const sp = dskp ? dskp.sp : (rpt.spRPT || '');

  const context = [
    'TINGKATAN: ' + tingkatan,
    'KELAS: ' + kelas,
    'TARIKH PDP: ' + tarikh,
    'TEMPOH: ' + tempoh + ' minit',
    'MINGGU RPT: ' + rpt.minggu,
    'BIDANG: ' + rpt.bidang,
    'TAJUK RPT: ' + rpt.tajuk,
    'SP RPT: ' + rpt.spRPT,
    'CATATAN RPT: ' + rpt.catatan,
    'STANDARD KANDUNGAN DSKP: ' + (sk || '[Tiada padanan khusus DSKP]'),
    'STANDARD PEMBELAJARAN DSKP: ' + (dskp ? dskp.sp : '[Gunakan konteks RPT sahaja]'),
    'CATATAN GURU: ' + (catatanGuru || '-')
  ].join('\n');

  const prompt = [
    'Anda ialah pembantu guru Pendidikan Islam KSSM Malaysia.',
    'Sediakan RPH ringkas, praktikal, berpusatkan murid dan sesuai untuk rekod guru.',
    'Gunakan HANYA maklumat kurikulum yang diberi. Jangan mencipta nombor SK/SP baharu.',
    'Jika tiada padanan khusus DSKP kerana minggu ulang kaji, pengayaan, pentaksiran atau SPM, jangan cipta SK/SP; bina aktiviti berdasarkan tajuk RPT sahaja.',
    'Objektif mesti boleh diukur dan munasabah untuk satu sesi.',
    'Kriteria kejayaan mesti sepadan terus dengan objektif.',
    'Aktiviti PdP hendaklah mempunyai Set Induksi, Langkah 1, Langkah 2, Langkah 3 dan Penutup dalam satu teks padat.',
    'Pentaksiran mesti nyata, contohnya pemerhatian, lisan, lembaran kerja, kuiz atau exit ticket.',
    'Refleksi ialah templat yang boleh diedit guru selepas PdP, bukan dakwaan bahawa murid telah berjaya.',
    '',
    context
  ].join('\n');

  const schema = {
    type: 'object',
    properties: {
      objektif: { type: 'string' },
      kriteriaKejayaan: { type: 'string' },
      aktivitiPdP: { type: 'string' },
      bbm: { type: 'string' },
      pentaksiran: { type: 'string' },
      refleksi: { type: 'string' }
    },
    required: ['objektif','kriteriaKejayaan','aktivitiPdP','bbm','pentaksiran','refleksi']
  };

  const ai = generateStructuredGemini_(prompt, schema);

  return {
    meta: {
      rptId: rpt.id,
      minggu: rpt.minggu,
      tarikh: tarikh,
      tingkatan: tingkatan,
      kelas: kelas,
      bidang: dskp ? dskp.bidang : rpt.bidang,
      tajuk: dskp ? dskp.tajuk : rpt.tajuk,
      sk: sk,
      sp: sp,
      dskpId: dskp ? dskp.id : '',
      sumber: dskp ? dskp.sumber : '',
      model: getSetting_('GEMINI_MODEL', 'gemini-3.8-flash')
    },
    rph: ai
  };
}

function simpanRPH(payload) {
  assertGuru_();
  if (!payload || !payload.meta || !payload.rph) throw new Error('Data RPH tidak sah.');
  const m = payload.meta, r = payload.rph;
  const sh = getSS_().getSheetByName('RPH');
  const id = 'RPH-' + Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'yyyyMMdd-HHmmss');

  sh.getRange(sh.getLastRow()+1,1,1,16).setValues([[
    id,
    m.tarikh,
    Number(m.tingkatan),
    m.kelas,
    m.bidang,
    m.tajuk,
    m.sk || '',
    m.sp || '',
    r.objektif || '',
    r.kriteriaKejayaan || '',
    r.aktivitiPdP || '',
    r.bbm || '',
    r.pentaksiran || '',
    r.refleksi || '',
    'Draf',
    'Gemini AI · ' + m.model
  ]]);
  logAudit_('JANA_RPH', id + ' · Tingkatan ' + m.tingkatan + ' · ' + m.kelas, 'Berjaya');

  return { id: id, saved: true };
}

function janaDanSimpanRPH(input) {
  assertGuru_();
  const hasil = janaRPH(input);
  const simpan = simpanRPH(hasil);
  return { hasil: hasil, simpan: simpan };
}


/* =========================
   DASHBOARD GURU
   ========================= */
function getDashboardData(filters) {
  assertGuru_();
  filters = filters || {};
  const tingkatanFilter = String(filters.tingkatan || '').trim();
  const kelasFilter = String(filters.kelas || '').trim().toLowerCase();
  const ss = getSS_();

  const readRows = (sheetName, cols) => {
    const sh = ss.getSheetByName(sheetName);
    const last = sh.getLastRow();
    if (last < 2) return [];
    return sh.getRange(2, 1, last - 1, cols).getValues().filter(r => r[0] !== '');
  };

  const bank = readRows('Bank_Soalan', 16).filter(r => !tingkatanFilter || String(r[1]) === tingkatanFilter);
  const kuiz = readRows('Kuiz', 15).filter(r => !tingkatanFilter || String(r[2]) === tingkatanFilter);
  const rph = readRows('RPH', 16).filter(r =>
    (!tingkatanFilter || String(r[2]) === tingkatanFilter) &&
    (!kelasFilter || String(r[3] || '').toLowerCase() === kelasFilter)
  );
  const murid = readRows('Murid', 6).filter(r =>
    (!tingkatanFilter || String(r[2]) === tingkatanFilter) &&
    (!kelasFilter || String(r[3] || '').toLowerCase() === kelasFilter)
  );
  const quizById = new Map(kuiz.map(r => [String(r[0]), r]));
  const jawapan = readRows('Jawapan_Murid', 12).filter(r => {
    const quiz = quizById.get(String(r[1] || ''));
    return !!quiz && (!kelasFilter || String(r[4] || '').toLowerCase() === kelasFilter);
  });

  const totalSoalan = bank.length;
  const soalanLulus = bank.filter(r => String(r[14]).trim() === 'Lulus').length;
  const belumSemak = bank.filter(r => String(r[14]).trim() === 'Belum Disemak').length;
  const kuizAktif = kuiz.filter(r => String(r[10]).trim() === 'Aktif').length;
  const muridAktif = murid.filter(r => String(r[5]).trim() === 'Aktif').length;
  const menungguSemakan = jawapan.filter(r => String(r[9]).trim() === 'Semakan Guru').length;

  const groups = {};
  jawapan.forEach(r => {
    const quizId = String(r[1] || '');
    const studentKey = String(r[2] || '') ||
      (String(r[3] || '').trim().toLowerCase() + '|' + String(r[4] || '').trim().toLowerCase());
    const key = quizId + '|' + studentKey;
    if (!groups[key]) groups[key] = { score: 0, max: 0 };
    if (typeof r[8] === 'number' && isFinite(r[8])) {
      groups[key].score += Number(r[8]);
      groups[key].max += Number(r[10] || 1);
    }
  });

  const attempts = Object.values(groups);
  const gradedAttempts = attempts.filter(x => x.max > 0);
  const avgPercent = gradedAttempts.length
    ? gradedAttempts.reduce((sum, x) => sum + (x.score / x.max * 100), 0) / gradedAttempts.length
    : 0;

  const recentRph = rph.slice(-5).reverse().map(r => ({
    id: String(r[0] || ''),
    tarikh: formatDashboardDate_(r[1]),
    tingkatan: String(r[2] || ''),
    kelas: String(r[3] || ''),
    bidang: String(r[4] || ''),
    tajuk: String(r[5] || ''),
    status: String(r[14] || '')
  }));

  const recentQuiz = kuiz.slice(-5).reverse().map(r => ({
    id: String(r[0] || ''),
    tajuk: String(r[1] || ''),
    tingkatan: String(r[2] || ''),
    kod: String(r[7] || ''),
    status: String(r[10] || '')
  }));

  return {
    totalSoalan,
    soalanLulus,
    belumSemak,
    kuizAktif,
    jumlahRespons: attempts.length,
    purataMarkah: Math.round(avgPercent),
    jumlahRPH: rph.length,
    muridAktif,
    menungguSemakan,
    recentRph,
    recentQuiz
  };
}

function formatDashboardDate_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'Asia/Kuala_Lumpur', 'dd/MM/yyyy');
  }
  return String(value);
}


/* =========================
   ROUTER WEB APP
   ========================= */


/* ==========================================================
   MODUL AKHIR: SEMAKAN, MURID, ANALISIS & REKOD
   ========================================================== */

function defaultMarkahPenuh_(jenis) {
  const j = String(jenis || '').trim().toLowerCase();
  if (j === 'objektif') return 1;
  if (j === 'struktur') return 2;
  if (j === 'esei') return 4;
  return 1;
}

function findRowById_(sheet, id, col) {
  col = col || 1;
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const finder = sheet.getRange(2, col, last - 1, 1)
    .createTextFinder(String(id))
    .matchEntireCell(true)
    .findNext();
  return finder ? finder.getRow() : 0;
}

/* ---------- Bank Soalan: semak, edit, lulus/tolak ---------- */

function getBankSemakanData(filters) {
  assertGuru_();
  filters = filters || {};
  const sh = getSS_().getSheetByName('Bank_Soalan');
  const last = sh.getLastRow();
  if (last < 2) return [];

  const rows = sh.getRange(2,1,last-1,16).getValues();
  const query = String(filters.query || '').trim().toLowerCase();
  const t = String(filters.tingkatan || '').trim();
  const status = String(filters.status || '').trim();
  const jenis = String(filters.jenis || '').trim();

  return rows.filter(r => {
    if (!r[0]) return false;
    if (t && String(r[1]) !== t) return false;
    if (status && String(r[14]) !== status) return false;
    if (jenis && String(r[5]) !== jenis) return false;
    if (query) {
      const hay = [r[0],r[2],r[3],r[6],r[11],r[12]].join(' ').toLowerCase();
      if (!hay.includes(query)) return false;
    }
    return true;
  }).slice(0,300).map(r => ({
    id: String(r[0]),
    tingkatan: Number(r[1]),
    bidang: String(r[2] || ''),
    tajuk: String(r[3] || ''),
    tahap: String(r[4] || ''),
    jenis: String(r[5] || ''),
    soalan: String(r[6] || ''),
    pilihanA: String(r[7] || ''),
    pilihanB: String(r[8] || ''),
    pilihanC: String(r[9] || ''),
    pilihanD: String(r[10] || ''),
    jawapan: String(r[11] || ''),
    huraian: String(r[12] || ''),
    sumber: String(r[13] || ''),
    status: String(r[14] || ''),
    markahPenuh: Number(r[15] || defaultMarkahPenuh_(r[5]))
  }));
}

function saveBankQuestion(payload) {
  assertGuru_();
  payload = payload || {};
  const id = String(payload.id || '').trim();
  if (!id) throw new Error('ID soalan tiada.');

  const sh = getSS_().getSheetByName('Bank_Soalan');
  const row = findRowById_(sh, id, 1);
  if (!row) throw new Error('Soalan tidak dijumpai.');

  const allowedStatus = ['Belum Disemak','Disemak','Lulus','Tolak'];
  const status = allowedStatus.includes(String(payload.status)) ? String(payload.status) : 'Belum Disemak';
  const markah = Math.max(1, Math.min(Number(payload.markahPenuh || 1), 20));

  sh.getRange(row, 7, 1, 10).setValues([[
    String(payload.soalan || ''),
    String(payload.pilihanA || ''),
    String(payload.pilihanB || ''),
    String(payload.pilihanC || ''),
    String(payload.pilihanD || ''),
    String(payload.jawapan || ''),
    String(payload.huraian || ''),
    String(payload.sumber || ''),
    status,
    markah
  ]]);
  logAudit_('SEMAK_SOALAN', id + ' · Status: ' + status, 'Berjaya');

  return { saved: true, id: id, status: status };
}

/* ---------- Pengurusan Murid ---------- */

function getMuridData(filters) {
  assertGuru_();
  filters = filters || {};
  const sh = getSS_().getSheetByName('Murid');
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2,1,last-1,6).getValues();
  const t = String(filters.tingkatan || '').trim();
  const query = String(filters.query || '').trim().toLowerCase();

  return rows.filter(r => {
    if (!r[0]) return false;
    if (t && String(r[2]) !== t) return false;
    if (query && ![r[0],r[1],r[3],r[4]].join(' ').toLowerCase().includes(query)) return false;
    return true;
  }).map(r => ({
    id: String(r[0]),
    nama: String(r[1] || ''),
    tingkatan: Number(r[2] || 0),
    kelas: String(r[3] || ''),
    kontak: String(r[4] || ''),
    status: String(r[5] || 'Aktif')
  }));
}

function saveMurid(payload) {
  assertGuru_();
  payload = payload || {};
  const sh = getSS_().getSheetByName('Murid');
  let id = String(payload.id || '').trim();
  const nama = String(payload.nama || '').trim();
  const tingkatan = Number(payload.tingkatan);
  const kelas = String(payload.kelas || '').trim();
  const kontak = String(payload.kontak || '').trim();
  const status = ['Aktif','Tidak Aktif'].includes(String(payload.status)) ? String(payload.status) : 'Aktif';

  if (!nama) throw new Error('Masukkan nama murid.');
  assertTingkatan_(tingkatan);
  if (!kelas) throw new Error('Masukkan kelas.');

  if (id) {
    const row = findRowById_(sh, id, 1);
    if (!row) throw new Error('ID murid tidak dijumpai.');
    sh.getRange(row,1,1,6).setValues([[id,nama,tingkatan,kelas,kontak,status]]);
  } else {
    id = nextMuridId_();
    sh.getRange(sh.getLastRow()+1,1,1,6).setValues([[id,nama,tingkatan,kelas,kontak,status]]);
  }
  logAudit_('SIMPAN_MURID', id + ' · Tingkatan ' + tingkatan + ' · ' + kelas, 'Berjaya');
  return { saved:true, id:id };
}

function nextMuridId_() {
  const sh = getSS_().getSheetByName('Murid');
  const last = sh.getLastRow();
  const ids = last > 1 ? sh.getRange(2,1,last-1,1).getDisplayValues().flat() : [];
  let max = 0;
  ids.forEach(id => {
    const m = String(id).match(/^M-(\d+)$/i);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return 'M-' + String(max + 1).padStart(3,'0');
}

/* ---------- Pengurusan Kuiz ---------- */

function getQuizManagerData() {
  assertGuru_();
  const sh = getSS_().getSheetByName('Kuiz');
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2,1,last-1,15).getValues()
    .filter(r => r[0])
    .slice(-30).reverse().map(r => ({
      id: String(r[0]),
      tajuk: String(r[1] || ''),
      tingkatan: Number(r[2] || 0),
      bidang: String(r[3] || ''),
      topik: String(r[4] || ''),
      tahap: String(r[5] || ''),
      bilangan: Number(r[6] || 0),
      kod: String(r[7] || ''),
      mula: formatDashboardDate_(r[8]),
      tamat: formatDashboardDate_(r[9]),
      status: String(r[10] || '')
    }));
}

function updateQuizStatus(id, status) {
  assertGuru_();
  if (!['Draf','Aktif','Tamat'].includes(String(status))) throw new Error('Status kuiz tidak sah.');
  const sh = getSS_().getSheetByName('Kuiz');
  const row = findRowById_(sh, id, 1);
  if (!row) throw new Error('Kuiz tidak dijumpai.');
  sh.getRange(row,11).setValue(String(status));
  logAudit_('STATUS_KUIZ', String(id) + ' · ' + String(status), 'Berjaya');
  return { saved:true, id:id, status:status };
}

/* ---------- Semakan jawapan Struktur/Esei ---------- */

function getManualReviewData(filters) {
  assertGuru_();
  filters = filters || {};
  const ss = getSS_();
  const js = ss.getSheetByName('Jawapan_Murid');
  const kuiz = ss.getSheetByName('Kuiz');
  const bank = ss.getSheetByName('Bank_Soalan');

  const jLast = js.getLastRow();
  if (jLast < 2) return [];

  const jRows = js.getRange(2,1,jLast-1,12).getValues();
  const qRows = kuiz.getLastRow() > 1 ? kuiz.getRange(2,1,kuiz.getLastRow()-1,15).getValues() : [];
  const bRows = bank.getLastRow() > 1 ? bank.getRange(2,1,bank.getLastRow()-1,16).getValues() : [];

  const qMap = new Map(qRows.filter(r=>r[0]).map(r=>[String(r[0]),r]));
  const bMap = new Map(bRows.filter(r=>r[0]).map(r=>[String(r[0]),r]));

  const query = String(filters.query || '').trim().toLowerCase();
  const quizFilter = String(filters.idKuiz || '').trim();

  return jRows.map((r, i) => ({r:r, sheetRow:i+2}))
    .filter(x => String(x.r[9] || '') === 'Semakan Guru')
    .filter(x => !quizFilter || String(x.r[1]) === quizFilter)
    .map(x => {
      const r = x.r;
      const qz = qMap.get(String(r[1]));
      let questionId = '';
      let questionText = '';
      let jenis = '';
      if (qz) {
        const ids = String(qz[11] || '').split(',').map(s=>s.trim()).filter(Boolean);
        questionId = ids[Number(r[5])-1] || '';
        const b = bMap.get(questionId);
        if (b) {
          questionText = String(b[6] || '');
          jenis = String(b[5] || '');
        }
      }
      return {
        sheetRow: x.sheetRow,
        idKuiz: String(r[1] || ''),
        tajukKuiz: qz ? String(qz[1] || '') : '',
        idMurid: String(r[2] || ''),
        nama: String(r[3] || ''),
        kelas: String(r[4] || ''),
        noSoalan: Number(r[5] || 0),
        questionId: questionId,
        jenis: jenis,
        soalan: questionText,
        jawapanMurid: String(r[6] || ''),
        skema: String(r[7] || ''),
        markah: r[8] === '' ? '' : Number(r[8]),
        markahPenuh: Number(r[10] || defaultMarkahPenuh_(jenis)),
        catatan: String(r[11] || '')
      };
    })
    .filter(x => !query || [x.nama,x.kelas,x.tajukKuiz,x.soalan,x.jawapanMurid].join(' ').toLowerCase().includes(query))
    .slice(0,300);
}

function saveManualMark(payload) {
  assertGuru_();
  payload = payload || {};
  const row = Number(payload.sheetRow);
  if (!row || row < 2) throw new Error('Baris jawapan tidak sah.');
  const sh = getSS_().getSheetByName('Jawapan_Murid');
  const max = Number(sh.getRange(row,11).getValue() || 1);
  let markah = Number(payload.markah);
  if (!isFinite(markah)) throw new Error('Markah tidak sah.');
  markah = Math.max(0, Math.min(markah, max));
  const catatan = String(payload.catatan || '').trim();
  const status = markah >= max ? 'Betul' : (markah <= 0 ? 'Salah' : 'Disemak');

  sh.getRange(row,9,1,4).setValues([[markah,status,max,catatan]]);
  logAudit_('SEMAK_JAWAPAN', 'Baris ' + row + ' · ' + markah + '/' + max, 'Berjaya');
  return { saved:true, row:row, markah:markah, max:max };
}

/* ---------- Analisis Prestasi ---------- */

function getAnalisisData(filters) {
  assertGuru_();
  filters = filters || {};
  const tingkatanFilter = String(filters.tingkatan || '').trim();
  const kelasFilter = String(filters.kelas || '').trim().toLowerCase();
  const ss = getSS_();
  const qz = ss.getSheetByName('Kuiz');
  const js = ss.getSheetByName('Jawapan_Murid');
  const bank = ss.getSheetByName('Bank_Soalan');

  const qRows = qz.getLastRow() > 1 ? qz.getRange(2,1,qz.getLastRow()-1,15).getValues() : [];
  const jRows = js.getLastRow() > 1 ? js.getRange(2,1,js.getLastRow()-1,12).getValues() : [];
  const bRows = bank.getLastRow() > 1 ? bank.getRange(2,1,bank.getLastRow()-1,16).getValues() : [];

  const qMap = new Map(qRows
    .filter(r => r[0] && (!tingkatanFilter || String(r[2]) === tingkatanFilter))
    .map(r=>[String(r[0]),r]));
  const bMap = new Map(bRows.filter(r=>r[0]).map(r=>[String(r[0]),r]));
  const threshold = Number(getSetting_('INTERVENSI_MARKAH','50')) || 50;

  const attempts = {};
  const topicAgg = {};
  const studentAgg = {};

  jRows.forEach(r => {
    if (!r[1]) return;
    const qid = String(r[1]);
    const q = qMap.get(qid);
    if (!q || (kelasFilter && String(r[4] || '').trim().toLowerCase() !== kelasFilter)) return;
    const sid = String(r[2] || '') || (String(r[3]||'')+'|'+String(r[4]||''));
    const attemptKey = qid+'|'+sid;

    if (!attempts[attemptKey]) {
      attempts[attemptKey] = {
        idKuiz: qid,
        tajukKuiz: q ? String(q[1]||'') : qid,
        tingkatan: q ? Number(q[2]||0) : 0,
        nama: String(r[3]||''),
        kelas: String(r[4]||''),
        idMurid: String(r[2]||''),
        score:0,max:0,pending:0
      };
    }

    const hasMark = typeof r[8] === 'number' && isFinite(r[8]);
    const max = Number(r[10] || 0);
    if (hasMark && max > 0) {
      attempts[attemptKey].score += Number(r[8]);
      attempts[attemptKey].max += max;
    } else if (String(r[9]||'') === 'Semakan Guru') {
      attempts[attemptKey].pending++;
    }

    // Resolve the exact question topic using the Kuiz ID-soalan list + No. Soalan.
    let b = null;
    if (q) {
      const ids = String(q[11]||'').split(',').map(s=>s.trim()).filter(Boolean);
      const qidBank = ids[Number(r[5])-1] || '';
      b = bMap.get(qidBank);
    }
    const topic = b ? String(b[3]||'') : (q ? String(q[4]||q[1]||'') : 'Tidak diketahui');
    const field = b ? String(b[2]||'') : (q ? String(q[3]||'') : '');
    const tkey = field+'|'+topic;
    if (!topicAgg[tkey]) topicAgg[tkey] = {bidang:field,tajuk:topic,score:0,max:0,items:0};
    if (hasMark && max > 0) {
      topicAgg[tkey].score += Number(r[8]);
      topicAgg[tkey].max += max;
      topicAgg[tkey].items++;
    }
  });

  Object.values(attempts).forEach(a => {
    a.percent = a.max > 0 ? Math.round(a.score/a.max*100) : 0;
    const skey = a.idMurid || (a.nama+'|'+a.kelas);
    if (!studentAgg[skey]) studentAgg[skey] = {
      idMurid:a.idMurid,nama:a.nama,kelas:a.kelas,totalPercent:0,count:0,pending:0
    };
    if (a.max > 0) {
      studentAgg[skey].totalPercent += a.percent;
      studentAgg[skey].count++;
    }
    studentAgg[skey].pending += a.pending;
  });

  const quizAgg = {};
  Object.values(attempts).forEach(a => {
    if (!quizAgg[a.idKuiz]) quizAgg[a.idKuiz] = {
      idKuiz:a.idKuiz,tajuk:a.tajukKuiz,tingkatan:a.tingkatan,attempts:0,total:0,graded:0,pending:0
    };
    quizAgg[a.idKuiz].attempts++;
    if (a.max > 0) {
      quizAgg[a.idKuiz].total += a.percent;
      quizAgg[a.idKuiz].graded++;
    }
    quizAgg[a.idKuiz].pending += a.pending;
  });

  const byQuiz = Object.values(quizAgg).map(x => ({
    idKuiz:x.idKuiz,
    tajuk:x.tajuk,
    tingkatan:x.tingkatan,
    respons:x.attempts,
    purata:x.graded ? Math.round(x.total/x.graded) : 0,
    pending:x.pending
  })).sort((a,b)=>b.respons-a.respons);

  const topics = Object.values(topicAgg)
    .filter(x=>x.max>0)
    .map(x=>({...x,percent:Math.round(x.score/x.max*100)}))
    .sort((a,b)=>a.percent-b.percent);

  const students = Object.values(studentAgg).map(x => ({
    idMurid:x.idMurid,
    nama:x.nama,
    kelas:x.kelas,
    purata:x.count ? Math.round(x.totalPercent/x.count) : 0,
    bilKuiz:x.count,
    pending:x.pending
  })).sort((a,b)=>a.purata-b.purata);

  const intervention = students.filter(x=>x.bilKuiz>0 && x.purata < threshold);

  const graded = Object.values(attempts).filter(x=>x.max>0);
  const overall = graded.length ? Math.round(graded.reduce((s,x)=>s+x.percent,0)/graded.length) : 0;

  return {
    threshold,
    jumlahRespons:Object.keys(attempts).length,
    purataKeseluruhan:overall,
    menungguSemakan:jRows.filter(r => {
      const q = qMap.get(String(r[1] || ''));
      return q && (!kelasFilter || String(r[4] || '').trim().toLowerCase() === kelasFilter) && String(r[9] || '') === 'Semakan Guru';
    }).length,
    bilIntervensi:intervention.length,
    byQuiz:byQuiz.slice(0,50),
    weakTopics:topics.slice(0,12),
    intervention:intervention.slice(0,100),
    students:students.slice(0,200)
  };
}

/* ---------- Rekod RPH ---------- */

function getRPHRecords(filters) {
  assertGuru_();
  filters = filters || {};
  const sh = getSS_().getSheetByName('RPH');
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2,1,last-1,16).getValues();
  const t = String(filters.tingkatan || '').trim();
  const status = String(filters.status || '').trim();
  const query = String(filters.query || '').trim().toLowerCase();

  return rows.filter(r => {
    if (!r[0]) return false;
    if (t && String(r[2]) !== t) return false;
    if (status && String(r[14]) !== status) return false;
    if (query && ![r[0],r[3],r[4],r[5],r[8],r[10]].join(' ').toLowerCase().includes(query)) return false;
    return true;
  }).slice(-300).reverse().map(r => ({
    id:String(r[0]),
    tarikh:formatDashboardDate_(r[1]),
    tingkatan:Number(r[2]||0),
    kelas:String(r[3]||''),
    bidang:String(r[4]||''),
    tajuk:String(r[5]||''),
    sk:String(r[6]||''),
    sp:String(r[7]||''),
    objektif:String(r[8]||''),
    kriteria:String(r[9]||''),
    aktiviti:String(r[10]||''),
    bbm:String(r[11]||''),
    pentaksiran:String(r[12]||''),
    refleksi:String(r[13]||''),
    status:String(r[14]||'Draf'),
    dijanaOleh:String(r[15]||'')
  }));
}

function updateRPHRecord(payload) {
  assertGuru_();
  payload = payload || {};
  const id = String(payload.id || '').trim();
  const sh = getSS_().getSheetByName('RPH');
  const row = findRowById_(sh,id,1);
  if (!row) throw new Error('RPH tidak dijumpai.');

  const allowed = ['Draf','Disemak','Lulus'];
  const status = allowed.includes(String(payload.status)) ? String(payload.status) : 'Draf';
  sh.getRange(row,9,1,7).setValues([[
    String(payload.objektif || ''),
    String(payload.kriteria || ''),
    String(payload.aktiviti || ''),
    String(payload.bbm || ''),
    String(payload.pentaksiran || ''),
    String(payload.refleksi || ''),
    status
  ]]);
  logAudit_('KEMAS_KINI_RPH', id + ' · Status: ' + status, 'Berjaya');
  return {saved:true,id:id,status:status};
}

/* ---------- Import, eksport dan sandaran ---------- */

function getPenggunaData() {
  assertAdmin_();
  const sh = getSS_().getSheetByName('Pengguna');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 6).getDisplayValues()
    .filter(r => String(r[0] || '').trim())
    .map(r => ({
      email: String(r[0] || '').trim(),
      nama: String(r[1] || '').trim(),
      peranan: String(r[2] || '').trim(),
      status: String(r[3] || '').trim(),
      catatan: String(r[4] || '').trim(),
      dikemasKini: String(r[5] || '').trim()
    }));
}

function savePengguna(payload) {
  const admin = assertAdmin_();
  payload = payload || {};
  const email = String(payload.email || '').trim().toLowerCase();
  const nama = String(payload.nama || '').trim();
  const peranan = ['Admin', 'Guru'].includes(String(payload.peranan)) ? String(payload.peranan) : 'Guru';
  const status = ['Aktif', 'Tidak Aktif'].includes(String(payload.status)) ? String(payload.status) : 'Aktif';
  const catatan = String(payload.catatan || '').trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Emel pengguna tidak sah.');
  if (!nama) throw new Error('Masukkan nama pengguna.');

  const sh = getSS_().getSheetByName('Pengguna');
  const values = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().flat() : [];
  const index = values.findIndex(v => String(v || '').trim().toLowerCase() === email);
  const record = [email, nama, peranan, status, catatan, Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'yyyy-MM-dd HH:mm:ss')];
  if (index >= 0) sh.getRange(index + 2, 1, 1, 6).setValues([record]);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, 6).setValues([record]);
  logAudit_('SIMPAN_PENGGUNA', email + ' · ' + peranan + ' · oleh ' + admin.email, 'Berjaya');
  return { saved: true, email: email };
}

function normaliseHeader_(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function findHeaderIndex_(headers, aliases) {
  const normalised = headers.map(normaliseHeader_);
  for (let i = 0; i < aliases.length; i++) {
    const index = normalised.indexOf(normaliseHeader_(aliases[i]));
    if (index !== -1) return index;
  }
  return -1;
}

function getCell_(row, index, fallback) {
  const value = index >= 0 && row[index] !== undefined && row[index] !== null
    ? String(row[index]).trim()
    : '';
  return value || (fallback || '');
}

function buildKurikulumRows_(input) {
  const jenis = String(input.jenis || '').trim().toUpperCase();
  const tingkatan = assertTingkatan_(input.tingkatan);
  const rawRows = Array.isArray(input.rows) ? input.rows : [];
  const rows = rawRows.filter(r => Array.isArray(r) && r.some(v => String(v || '').trim() !== ''));
  if (rows.length < 2) throw new Error('Fail atau Sheet sumber mesti mengandungi tajuk lajur dan sekurang-kurangnya satu rekod.');

  const headers = rows[0];
  const data = rows.slice(1);
  const indexes = {
    id: findHeaderIndex_(headers, ['ID', 'Kod']),
    bidang: findHeaderIndex_(headers, ['Bidang']),
    tajuk: findHeaderIndex_(headers, ['Tajuk', 'Topik']),
    sumber: findHeaderIndex_(headers, ['Sumber URL', 'URL', 'Sumber']),
    status: findHeaderIndex_(headers, ['Status Semakan', 'Status']),
    catatan: findHeaderIndex_(headers, ['Catatan', 'Nota'])
  };

  if (jenis === 'DSKP') {
    indexes.sk = findHeaderIndex_(headers, ['Standard Kandungan', 'SK']);
    indexes.sp = findHeaderIndex_(headers, ['Standard Pembelajaran', 'SP']);
    if ([indexes.bidang, indexes.tajuk, indexes.sk, indexes.sp].some(i => i < 0)) {
      throw new Error('Header DSKP diperlukan: Bidang, Tajuk, Standard Kandungan dan Standard Pembelajaran.');
    }
    return data.map((r, i) => [
      getCell_(r, indexes.id, 'D' + tingkatan + '-' + String(i + 1).padStart(3, '0')),
      tingkatan,
      getCell_(r, indexes.bidang),
      getCell_(r, indexes.tajuk),
      getCell_(r, indexes.sk),
      getCell_(r, indexes.sp),
      getCell_(r, indexes.sumber, String(input.sumberUrl || '').trim()),
      ['Belum Disemak','Disemak','Lulus','Tolak'].includes(getCell_(r, indexes.status)) ? getCell_(r, indexes.status) : 'Belum Disemak',
      getCell_(r, indexes.catatan)
    ]).filter(r => r[2] && r[3] && r[4] && r[5]);
  }

  if (jenis === 'RPT') {
    indexes.tahun = findHeaderIndex_(headers, ['Tahun']);
    indexes.minggu = findHeaderIndex_(headers, ['Minggu']);
    indexes.tarikhMula = findHeaderIndex_(headers, ['Tarikh Mula', 'Mula']);
    indexes.tarikhAkhir = findHeaderIndex_(headers, ['Tarikh Akhir', 'Akhir']);
    indexes.sp = findHeaderIndex_(headers, ['Standard Pembelajaran', 'SP']);
    if ([indexes.minggu, indexes.bidang, indexes.tajuk].some(i => i < 0)) {
      throw new Error('Header RPT diperlukan: Minggu, Bidang dan Tajuk.');
    }
    const tahun = getSetting_('TAHUN_RPT', '2026');
    return data.map((r, i) => [
      getCell_(r, indexes.id, 'RPT' + tingkatan + '-' + String(i + 1).padStart(3, '0')),
      getCell_(r, indexes.tahun, tahun),
      tingkatan,
      getCell_(r, indexes.minggu),
      getCell_(r, indexes.tarikhMula),
      getCell_(r, indexes.tarikhAkhir),
      getCell_(r, indexes.bidang),
      getCell_(r, indexes.tajuk),
      getCell_(r, indexes.sp),
      getCell_(r, indexes.catatan)
    ]).filter(r => r[3] && r[6] && r[7]);
  }

  throw new Error('Jenis import mestilah DSKP atau RPT.');
}

function saveImportedKurikulum_(input) {
  const user = assertGuru_();
  const jenis = String(input.jenis || '').trim().toUpperCase();
  const tingkatan = assertTingkatan_(input.tingkatan);
  const output = buildKurikulumRows_(input);
  if (!output.length) throw new Error('Tiada rekod yang lengkap untuk diimport.');

  const sh = getSS_().getSheetByName(jenis + '_T' + tingkatan);
  if (!sh) throw new Error('Tab sasaran tidak dijumpai.');
  const width = jenis === 'DSKP' ? 9 : 10;
  const mode = String(input.mode || 'Tambah').trim();
  let saved = output;

  if (mode === 'Ganti') {
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, width).clearContent();
  } else {
    const existing = sh.getLastRow() > 1
      ? new Set(sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().flat().map(String))
      : new Set();
    saved = output.filter(r => !existing.has(String(r[0])));
  }
  if (saved.length) sh.getRange(sh.getLastRow() + 1, 1, saved.length, width).setValues(saved);
  logAudit_('IMPORT_' + jenis, 'Tingkatan ' + tingkatan + ' · ' + saved.length + ' rekod · ' + mode + ' · ' + user.email, 'Berjaya');
  return { saved: saved.length, skipped: output.length - saved.length, jenis: jenis, tingkatan: tingkatan, mode: mode };
}

function importKurikulumRows(input) {
  return saveImportedKurikulum_(input || {});
}

function spreadsheetIdFromUrl_(value) {
  const match = String(value || '').match(/[a-zA-Z0-9_-]{20,}/);
  if (!match) throw new Error('Pautan Google Sheet tidak sah.');
  return match[0];
}

function importKurikulumFromSheet(input) {
  assertGuru_();
  input = input || {};
  const tab = String(input.tabSumber || '').trim();
  if (!tab) throw new Error('Masukkan nama tab sumber.');
  const sourceUrl = String(input.urlSumber || '').trim();
  const source = SpreadsheetApp.openById(spreadsheetIdFromUrl_(sourceUrl));
  const sh = source.getSheetByName(tab);
  if (!sh) throw new Error('Tab sumber tidak dijumpai: ' + tab);
  input.rows = sh.getDataRange().getDisplayValues();
  input.sumberUrl = sourceUrl;
  return saveImportedKurikulum_(input);
}

function getExportData(input) {
  assertGuru_();
  input = input || {};
  const jenis = String(input.jenis || '').trim().toUpperCase();
  const tingkatan = String(input.tingkatan || '').trim();
  const kelas = String(input.kelas || '').trim().toLowerCase();
  let sheetName = '';
  let tingkatanCol = -1;
  let kelasCol = -1;

  if (jenis === 'DSKP' || jenis === 'RPT') {
    assertTingkatan_(tingkatan);
    sheetName = jenis + '_T' + tingkatan;
  } else if (jenis === 'BANK_SOALAN') {
    sheetName = 'Bank_Soalan'; tingkatanCol = 1;
  } else if (jenis === 'RPH') {
    sheetName = 'RPH'; tingkatanCol = 2; kelasCol = 3;
  } else if (jenis === 'MURID') {
    sheetName = 'Murid'; tingkatanCol = 2; kelasCol = 3;
  } else {
    throw new Error('Pilih data DSKP, RPT, Bank Soalan, RPH atau Murid.');
  }

  const sh = getSS_().getSheetByName(sheetName);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const values = sh.getRange(1, 1, Math.max(lastRow, 1), lastCol).getDisplayValues();
  const headers = values.shift() || [];
  const rows = values.filter(r => r.some(v => String(v || '').trim() !== ''))
    .filter(r => (!tingkatan || tingkatanCol < 0 || String(r[tingkatanCol]) === tingkatan))
    .filter(r => (!kelas || kelasCol < 0 || String(r[kelasCol] || '').trim().toLowerCase() === kelas));
  logAudit_('EKSPORT_DATA', sheetName + ' · ' + rows.length + ' rekod', 'Berjaya');
  return { sheetName: sheetName, headers: headers, rows: rows, filename: APP_NAME.replace(/[^a-z0-9]+/gi, '_') + '_' + sheetName + '_' + Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'yyyyMMdd-HHmmss') + '.xlsx' };
}

function getBackupFolder_() {
  const folderId = getSetting_('BACKUP_FOLDER_ID', '');
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch (e) {}
  }
  const name = 'Sandaran ' + APP_NAME;
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function backupPortal_(jenis) {
  const stamp = Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'yyyyMMdd-HHmmss');
  const copy = DriveApp.getFileById(SPREADSHEET_ID).makeCopy(APP_NAME + ' · Sandaran ' + stamp, getBackupFolder_());
  logAudit_('SANDARAN_' + String(jenis || 'MANUAL').toUpperCase(), copy.getName(), 'Berjaya');
  return { success: true, name: copy.getName(), url: copy.getUrl() };
}

function createBackupNow() {
  assertGuru_();
  return backupPortal_('manual');
}

function backupMingguan_() {
  return backupPortal_('mingguan');
}

function setupWeeklyBackup() {
  assertGuru_();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'backupMingguan_')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('backupMingguan_').timeBased().everyWeeks(1).atHour(1).create();
  logAudit_('TETAP_SANDARAN_MINGGUAN', 'Pencetus mingguan pada sekitar 1 pagi.', 'Berjaya');
  return { success: true, message: 'Sandaran mingguan telah dijadualkan.' };
}

/* ---------- Web App: satu Portal SPA sahaja ---------- */

function doGet(e) {
  const tpl = HtmlService.createTemplateFromFile('Portal');
  tpl.initialPage = (e && e.parameter && e.parameter.page) ? String(e.parameter.page) : 'dashboard';
  tpl.initialCode = (e && e.parameter && e.parameter.code) ? String(e.parameter.code) : '';
  return tpl.evaluate()
    .setTitle(APP_NAME + ' · Portal Pendidikan Islam')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
