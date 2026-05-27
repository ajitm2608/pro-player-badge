// Pro Player Badge — Service Worker
// =================================================================
// Strategy:
//   • Same-origin requests (the HTML shell + any future static asset):
//       stale-while-revalidate — serve from cache instantly, refresh
//       in the background, so repeat visits are sub-100ms paint.
//   • Apps Script API (script.google.com / script.googleusercontent.com):
//       NEVER cache here. The IndexedDB layer in index.html already
//       handles host + cohort caching with TTLs and a refresh button.
//       Double-caching at the SW layer would just produce stale data
//       that's harder to invalidate.
//   • Cross-origin everything else: passthrough.
// =================================================================

// Bump the version segment whenever you push a meaningful HTML change.
// On activation, all old caches with a different version are deleted,
// forcing a fresh fetch of the shell.
const CACHE_VERSION = 'ppb-shell-v8';

// Resources prefetched at install time. Keep this small — anything else
// gets cached lazily on first hit by the stale-while-revalidate handler.
const PRECACHE_URLS = [
  './',
  './index.html',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION)
      .then((cache) => cache.addAll(PRECACHE_URLS))
      // Activate this SW immediately even if another version is in control.
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(
        keys.filter((k) => k !== CACHE_VERSION).map((k) => caches.delete(k))
      ))
      // Take control of pages already loaded so they start using this SW now.
      .then(() => self.clients.claim())
  );
});

function isApiRequest(url) {
  return url.host.indexOf('script.google.com') !== -1
      || url.host.indexOf('script.googleusercontent.com') !== -1
      || url.host.indexOf('docs.google.com') !== -1;
}

self.addEventListener('fetch', (event) => {
  // Only GETs. POST/PUT/etc. go straight to network.
  if (event.request.method !== 'GET') return;

  const url = new URL(event.request.url);

  // Never proxy the Apps Script API through cache — let IDB handle it.
  if (isApiRequest(url)) return;

  // Only serve same-origin (the GitHub Pages site itself).
  if (url.origin !== self.location.origin) return;

  event.respondWith(staleWhileRevalidate(event.request));
});

async function staleWhileRevalidate(request) {
  const cache = await caches.open(CACHE_VERSION);
  const cached = await cache.match(request);

  // Kick off the network fetch in parallel. If it succeeds AND looks like a
  // cacheable response (200 OK, basic origin), write it back to the cache for
  // the NEXT visit. Failures silently fall back to cached.
  const networkPromise = fetch(request).then((response) => {
    if (response && response.status === 200 && response.type === 'basic') {
      cache.put(request, response.clone()).catch(() => {});
    }
    return response;
  }).catch(() => null);

  // Return cached immediately if we have it; otherwise wait for network.
  return cached || (await networkPromise) || new Response('', { status: 504 });
}
