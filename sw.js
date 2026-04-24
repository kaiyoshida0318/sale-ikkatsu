const CACHE_NAME = 'sale-ikkatsu-v0.6.0';
const PYODIDE_CACHE = 'pyodide-v0.27.0';

const APP_FILES = [
  './',
  './index.html',
  './pyapp/merger.py',
  './pyapp/excel_builder.py',
  './pyapp/runner.py',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_FILES))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(
        names
          .filter((n) => n !== CACHE_NAME && n !== PYODIDE_CACHE)
          .map((n) => caches.delete(n))
      )
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  if (url.hostname === 'cdn.jsdelivr.net' && url.pathname.includes('/pyodide/')) {
    event.respondWith(cacheFirst(event.request, PYODIDE_CACHE));
    return;
  }

  if (url.origin === location.origin) {
    if (url.pathname.includes('/pyapp/') || url.pathname.endsWith('.html') || url.pathname.endsWith('/')) {
      event.respondWith(networkFirst(event.request, CACHE_NAME));
      return;
    }
  }
});

async function cacheFirst(request, cacheName) {
  const cache = await caches.open(cacheName);
  const cached = await cache.match(request);
  if (cached) return cached;
  const response = await fetch(request);
  if (response.ok) cache.put(request, response.clone());
  return response;
}

async function networkFirst(request, cacheName) {
  const cache = await caches.open(cacheName);
  try {
    const response = await fetch(request);
    if (response.ok) cache.put(request, response.clone());
    return response;
  } catch (e) {
    const cached = await cache.match(request);
    if (cached) return cached;
    throw e;
  }
}
