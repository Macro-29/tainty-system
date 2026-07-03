/* Tainty PWA · Service Worker
   Estrategia:
   - Páginas HTML (navegación): NETWORK-FIRST -> siempre datos frescos; si no hay
     internet, se muestra la última versión cacheada de esa página.
   - Assets estáticos (css/js/imágenes/íconos): CACHE-FIRST con revalidación en
     segundo plano -> carga instantánea y se actualiza solo.
   - Llamadas al backend de Apps Script (script.google.com): NUNCA se cachean.
   Sube CACHE_VERSION cada vez que quieras forzar la actualización del cache. */

const CACHE_VERSION = 'tainty-v2';
const STATIC_CACHE = `${CACHE_VERSION}-static`;
const PAGES_CACHE = `${CACHE_VERSION}-pages`;

// Recursos base para que la app abra aunque no haya internet.
const PRECACHE = [
  './',
  './index.html',
  './assets/common.css',
  './assets/common.js',
  './assets/auth.js',
  './assets/config.js',
  './assets/icons/icon-192.png',
  './assets/icons/icon-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(STATIC_CACHE)
      .then((cache) => cache.addAll(PRECACHE).catch(() => {}))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((k) => !k.startsWith(CACHE_VERSION))
          .map((k) => caches.delete(k))
      )
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;

  // Solo manejamos GET del mismo origen. El backend (POST a Apps Script) pasa directo.
  if (req.method !== 'GET') return;

  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return; // fuentes, Apps Script, etc.

  // Navegación / documentos HTML -> network-first.
  const isPage = req.mode === 'navigate' ||
    (req.headers.get('accept') || '').includes('text/html');

  if (isPage) {
    event.respondWith(
      fetch(req)
        .then((res) => {
          const copy = res.clone();
          caches.open(PAGES_CACHE).then((c) => c.put(req, copy));
          return res;
        })
        .catch(() =>
          caches.match(req).then((cached) => cached || caches.match('./index.html'))
        )
    );
    return;
  }

  // Assets estáticos -> cache-first + revalidación en segundo plano.
  event.respondWith(
    caches.match(req).then((cached) => {
      const network = fetch(req)
        .then((res) => {
          if (res && res.status === 200) {
            const copy = res.clone();
            caches.open(STATIC_CACHE).then((c) => c.put(req, copy));
          }
          return res;
        })
        .catch(() => cached);
      return cached || network;
    })
  );
});
