// BSU Schedule - Service Worker
// Стратегия: Network-First для schedule_data.json (всегда свежие данные с сервера),
//             Stale-While-Revalidate для index.html (мгновенная загрузка офлайн).

const CACHE_NAME = 'bsu-schedule-v2';
const SHELL_ASSETS = [
  './',
  './index.html',
  './schedule_data.json'
];

// ─── Install: предзагрузка оболочки приложения ──────────────────────────────
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache =>
      cache.addAll(SHELL_ASSETS.map(url => new Request(url, { cache: 'reload' })))
    ).then(() => self.skipWaiting())
  );
});

// ─── Activate: очистка устаревших версий кэша ─────────────────────────────────
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
      )
    ).then(() => self.clients.claim())
  );
});

// ─── Fetch ────────────────────────────────────────────────────────────────────
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);
  const isLocal = url.origin === self.location.origin;
  if (event.request.method !== 'GET' || !isLocal) return;

  const path = url.pathname.replace(/\/$/, '') || '/';

  // 1. schedule_data.json: Network-First (всегда пытаемся загрузить актуальное расписание)
  if (path.endsWith('schedule_data.json')) {
    event.respondWith(
      fetch(event.request)
        .then(response => {
          if (response.ok) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then(cache => {
              // Сохраняем и с query-параметрами, и под чистым ключом для гарантированного офлайн-доступа
              cache.put('./schedule_data.json', clone.clone());
              cache.put(event.request, clone);
            });
          }
          return response;
        })
        .catch(async () => {
          const cache = await caches.open(CACHE_NAME);
          const cached = (await cache.match(event.request, { ignoreSearch: true })) ||
                         (await cache.match('./schedule_data.json')) ||
                         (await cache.match('/schedule_data.json')) ||
                         (await cache.match('schedule_data.json'));
          if (cached) return cached;
          return new Response(JSON.stringify({}), {
            status: 503,
            statusText: 'Service Unavailable',
            headers: { 'Content-Type': 'application/json' }
          });
        })
    );
    return;
  }

  // 2. index.html и корень: Stale-While-Revalidate (мгновенный старт из кэша, обновление в фоне)
  if (path.endsWith('.html') || path === '/' || path.endsWith('/')) {
    event.respondWith(
      caches.open(CACHE_NAME).then(async cache => {
        const cached = (await cache.match(event.request)) ||
                       (await cache.match('./index.html')) ||
                       (await cache.match('./'));

        const networkFetch = fetch(event.request).then(response => {
          if (response.ok) {
            cache.put(event.request, response.clone());
            cache.put('./index.html', response.clone());
          }
          return response;
        }).catch(() => null);

        return cached || networkFetch;
      })
    );
    return;
  }
});
