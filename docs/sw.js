const CACHE_NAME = "storyboard-pwa-v1";
const CORE_FILES = [
  "./",
  "./wechat-guide.html",
  "./storyboard-standalone.html",
  "./index.html",
  "./styles.css",
  "./app.js",
  "./manifest.webmanifest",
  "./icon.svg",
];

self.addEventListener("install", event => {
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(CORE_FILES)));
  self.skipWaiting();
});

self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", event => {
  const { request } = event;
  if (request.method !== "GET") return;

  event.respondWith(
    caches.match(request).then(cached => {
      if (cached) return cached;
      return fetch(request)
        .then(response => {
          const copied = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(request, copied));
          return response;
        })
        .catch(() => caches.match("./storyboard-standalone.html"));
    })
  );
});
