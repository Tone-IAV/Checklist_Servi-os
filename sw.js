self.addEventListener('install', e=>{
  e.waitUntil(caches.open('os-cache').then(cache=>cache.addAll(['./'])));
});
self.addEventListener('fetch', e=>{
  e.respondWith(fetch(e.request).catch(()=>caches.match(e.request)));
});
