/* Service Worker – Class Visit PWA */
const APP_CACHE = 'class-visit-app-v1';
const CDN_CACHE = 'class-visit-cdn-v1';

const APP_ASSETS = [
  './index.html',
  './styles.css',
  './app.js',
  './manifest.json',
  './icons/icon.svg',
  './icons/icon-maskable.svg'
];

const CDN_ASSETS = [
  'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js'
];

self.addEventListener('install', event => {
  event.waitUntil(
    (async () => {
      const appCache = await caches.open(APP_CACHE);
      await appCache.addAll(APP_ASSETS);

      try {
        const cdnCache = await caches.open(CDN_CACHE);
        await cdnCache.addAll(CDN_ASSETS);
      } catch {
        /* Offline on first install – CDN cached when online */
      }

      await self.skipWaiting();
    })()
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    (async () => {
      const keys = await caches.keys();
      await Promise.all(
        keys
          .filter(k => k !== APP_CACHE && k !== CDN_CACHE)
          .map(k => caches.delete(k))
      );
      await self.clients.claim();
    })()
  );
});

self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;

  const url = new URL(event.request.url);

  /* CDN resources – cache-first, network fallback */
  if (url.origin !== self.location.origin) {
    event.respondWith(
      (async () => {
        const cdnCache = await caches.open(CDN_CACHE);
        const cached = await cdnCache.match(event.request);
        if (cached) return cached;

        try {
          const response = await fetch(event.request);
          if (response.ok) cdnCache.put(event.request, response.clone());
          return response;
        } catch {
          return new Response('/* offline */', {
            headers: { 'Content-Type': 'application/javascript' }
          });
        }
      })()
    );
    return;
  }

  /* App assets – cache-first, network fallback, then update cache */
  event.respondWith(
    (async () => {
      const appCache = await caches.open(APP_CACHE);
      const cached = await appCache.match(event.request);
      if (cached) return cached;

      try {
        const response = await fetch(event.request);
        if (response.ok) appCache.put(event.request, response.clone());
        return response;
      } catch {
        return new Response('Offline', { status: 503 });
      }
    })()
  );
});
