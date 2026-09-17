const CACHE='mypi-spm-v1';
const APP_SHELL=['./','./index.html','./manifest.json','./offline.html','./icons/icon-192.png','./icons/icon-512.png'];
self.addEventListener('install',event=>{event.waitUntil(caches.open(CACHE).then(c=>c.addAll(APP_SHELL)));self.skipWaiting()});
self.addEventListener('activate',event=>{event.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))));self.clients.claim()});
self.addEventListener('fetch',event=>{
 const req=event.request,url=new URL(req.url);
 if(url.hostname.includes('script.google.com')||url.hostname.includes('googleusercontent.com')||url.hostname.includes('googleapis.com')){
   event.respondWith(fetch(req).catch(()=>caches.match('./offline.html')));return;
 }
 if(req.mode==='navigate'){
   event.respondWith(fetch(req).then(res=>{const copy=res.clone();caches.open(CACHE).then(c=>c.put(req,copy));return res}).catch(()=>caches.match(req).then(r=>r||caches.match('./index.html'))));return;
 }
 event.respondWith(caches.match(req).then(cached=>cached||fetch(req).then(res=>{if(req.method==='GET'&&res.ok){const copy=res.clone();caches.open(CACHE).then(c=>c.put(req,copy))}return res})));
});