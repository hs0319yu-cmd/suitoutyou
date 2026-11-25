const CACHE_NAME='suitoutyou-shell-v16-icon-isolated';
const OWN_CACHE_PREFIX='suitoutyou-';
const SHELL = ['./', './index.html', './api.js', './manifest.json', './icon.svg', './apple-touch-icon.png'];
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(SHELL))
      .then(() => self.skipWaiting())
  );
});
self.addEventListener('activate', event => {
  // このアプリ自身のキャッシュだけを更新。他アプリのキャッシュは削除しない。
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(key => key.startsWith(OWN_CACHE_PREFIX) && key !== CACHE_NAME)
            .map(key => caches.delete(key))
      ))
      .then(() => self.clients.claim())
  );
});
self.addEventListener('fetch', event => {
  const req = event.request;
  const url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== self.location.origin) return;
  if (req.mode === 'navigate') {
    event.respondWith(fetch(req).then(res => {
      const copy = res.clone();
      caches.open(CACHE_NAME).then(cache => cache.put('./index.html', copy));
      return res;
    }).catch(() => caches.match('./index.html')));
    return;
  }
  event.respondWith(caches.match(req).then(cached => cached || fetch(req).then(res => {
    if (res.ok) {
      const copy = res.clone();
      caches.open(CACHE_NAME).then(cache => cache.put(req, copy));
    }
    return res;
  })));
});
