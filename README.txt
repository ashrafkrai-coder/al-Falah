myPI KSSM PWA — DEPLOY READY

LANGKAH:
1. Dalam projek Google Apps Script, kemas kini Code.gs dan HTML Portal menggunakan Portal-AppsScript.html.
2. Pastikan tab "Pengguna" Sheet mengandungi emel pentadbir/guru dan berstatus Aktif.
3. Deploy Apps Script sebagai Web App: jalankan sebagai pengguna yang mengakses aplikasi dan hadkan akses kepada akaun sekolah yang dibenarkan.
4. Tetapkan GEMINI_API_KEY dalam Script Properties jika fungsi AI mahu digunakan.
5. Deploy folder PWA ini ke Vercel.
6. Buka URL Vercel dan masukkan URL Apps Script yang berakhir /exec.
7. Install melalui browser: Install app / Add to Home Screen.

FUNGSI KSSM:
- Sokongan Tingkatan 1 hingga 5 untuk bank soalan, kuiz, RPH, murid dan analisis.
- Import DSKP/RPT daripada Excel/CSV atau Google Sheet, serta eksport Excel/PDF.
- Akses guru/admin, Audit_Log dan sandaran manual atau mingguan.
- Tab DSKP Tingkatan 1 hingga 3 mengandungi indeks enam bidang dengan pautan dokumen rasmi. Import butiran SK/SP yang telah disemak sebelum janaan AI diaktifkan.

PWA shell boleh dibuka offline.
Google Sheet, Gemini AI, kuiz dan analisis memerlukan internet.
API key kekal di Apps Script Properties — tidak berada dalam PWA.
