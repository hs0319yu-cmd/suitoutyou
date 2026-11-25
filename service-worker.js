const CACHE_NAME='suitoutyou-shell-v17-isolated';
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
  // suitoutyou 自身のキャッシュだけを更新。他アプリのキャッシュには触れない。
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys
          .filter(key => key.startsWith(OWN_CACHE_PREFIX) && key !== CACHE_NAME)
          .map(key => caches.delete(key))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const req = event.request;
  const url = new URL(req.url);

  if (req.method !== 'GET' || url.origin !== self.location.origin) return;

  // この Service Worker の公開フォルダ内だけを扱う。
  const appBase = new URL(self.registration.scope);
  if (!url.href.startsWith(appBase.href)) return;

  // 必ず suitoutyou 専用キャッシュだけを参照する。
  event.respondWith((async () => {
    const cache = await caches.open(CACHE_NAME);

    if (req.mode === 'navigate') {
      try {
        const response = await fetch(req);
        if (response && response.ok) {
          await cache.put(new URL('./index.html', appBase).href, response.clone());
        }
        return response;
      } catch (error) {
        const fallback = await cache.match(new URL('./index.html', appBase).href);
        if (fallback) return fallback;
        throw error;
      }
    }

    const cached = await cache.match(req);
    if (cached) return cached;

    const response = await fetch(req);
    if (response && response.ok) {
      await cache.put(req, response.clone());
    }
    return response;
  })());
});
