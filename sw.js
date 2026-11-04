// Keeps the app itself available offline. Data offline is handled by Firestore's own cache.
const CACHE = 'workbench-v8';
const CORE = ['./', 'index.html', 'boot.js', 'styles.css', 'app.js', 'store.js', 'i18n.js', 'firebase-config.js', 'manifest.webmanifest',
  'xlsx.js', 'tracker.js', 'icons/icon-192.png', 'assets/mccann-valley.png', 'fonts/rubik-hebrew-400-normal.woff2', 'fonts/rubik-latin-400-normal.woff2', 'vendor/firebase/firebase-app.js', 'vendor/firebase/firebase-auth.js', 'vendor/firebase/firebase-firestore.js'];
self.addEventListener('install', e=>{ e.waitUntil(caches.open(CACHE).then(c=>c.addAll(CORE)).then(()=>self.skipWaiting())); });
self.addEventListener('activate', e=>{ e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim())); });
self.addEventListener('fetch', e=>{
  const req = e.request, url = new URL(req.url);
  if (req.method !== 'GET') return;
  const sameOrigin = url.origin === location.origin;
  if (!sameOrigin) return; // Firebase traffic goes straight to the network
  // Network first for our own files (so updates show up), cache as fallback when offline
  e.respondWith(fetch(req).then(res=>{
    if (res.ok && res.type === 'basic'){ const copy = res.clone(); caches.open(CACHE).then(c=>c.put(req, copy)); }
    return res;
  }).catch(()=>caches.match(req).then(r=>r || (req.mode === 'navigate' ? caches.match('index.html') : Response.error()))));
});
