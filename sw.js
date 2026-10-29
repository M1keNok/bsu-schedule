// BSU Schedule - Service Worker
// Стратегия: Network-First для данных и html-оболочки (всегда свежий интерфейс при наличии сети, мгновенный откат на кэш в офлайне),
//             Cache-First для внешних CDN библиотек (Tailwind CSS, daisyUI, Telegram SDK).

const CACHE_NAME = 'bsu-schedule-v16';
const SHELL_ASSETS = [
  './',
  './index.html',
  './motion.js',
  './schedule_data.json'
];

const CDN_ASSETS = [
  'https://cdn.tailwindcss.com',
  'https://cdn.jsdelivr.net/npm/daisyui@latest/dist/full.min.css',
  'https://telegram.org/js/telegram-web-app.js'
];

// ─── Install: предзагрузка оболочки приложения ──────────────────────────────
self.addEventListener('install', event => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(async cache => {
      await cache.addAll(SHELL_ASSETS.map(url => new Request(url, { cache: 'reload' })));
      // Предзагружаем CDN-ресурсы, не блокируя установку в случае сбоя сети
      await Promise.allSettled(
        CDN_ASSETS.map(url =>
          fetch(url)
            .then(res => {
              if (res.ok || res.type === 'opaque') {
                return cache.put(url, res);
              }
            })
            .catch(() => null)
        )
      );
    })
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

// Мгновенная активация по команде от страницы
self.addEventListener('message', event => {
  if (event.data && event.data.action === 'skipWaiting') {
    self.skipWaiting();
  }
});

// ─── Fetch ────────────────────────────────────────────────────────────────────
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);
  const isLocal = url.origin === self.location.origin;
  const isCdn = url.hostname === 'cdn.tailwindcss.com' ||
                url.hostname === 'cdn.jsdelivr.net' ||
                url.hostname === 'telegram.org';

  if (event.request.method !== 'GET') return;
  if (!isLocal && !isCdn) return;

  // 1. Внешние CDN-библиотеки (Tailwind CSS, daisyUI, Telegram SDK): Cache-First
  if (isCdn) {
    event.respondWith(
      caches.match(event.request).then(cached => {
        if (cached) return cached;
        return fetch(event.request).then(networkResponse => {
          if (networkResponse && (networkResponse.status === 200 || networkResponse.type === 'opaque')) {
            const clone = networkResponse.clone();
            caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
          }
          return networkResponse;
        }).catch(() => cached || new Response('', { status: 408, statusText: 'Offline' }));
      })
    );
    return;
  }

  const path = url.pathname.replace(/\/$/, '') || '/';

  // 2. JSON-данные (расписание): Network-First
  if (path.endsWith('schedule_data.json')) {
    event.respondWith(
      fetch(event.request)
        .then(response => {
          if (response.ok) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then(cache => {
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

  // 3. index.html и корень: Network-First с быстрым откатом на кэш для 100% офлайна
  if (path.endsWith('.html') || path === '/' || path.endsWith('/')) {
    event.respondWith(
      fetch(new Request(event.request, { cache: 'reload' }))
        .then(response => {
          if (response.ok) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then(cache => {
              cache.put(event.request, clone.clone());
              cache.put('./index.html', clone);
            });
          }
          return response;
        })
        .catch(async () => {
          const cache = await caches.open(CACHE_NAME);
          return (await cache.match(event.request)) ||
                 (await cache.match('./index.html')) ||
                 (await cache.match('./')) ||
                 (await cache.match('/'));
        })
    );
    return;
  }
});
