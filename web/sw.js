/* Ateş İdare service worker — cache-first shell, network-first data */
const CACHE = "ates-idare-v4";
const SHELL = [
  "./",
  "./index.html",
  "./css/app.css",
  "./js/app.js",
  "./js/calc.js",
  "./js/geometry.js",
  "./js/interpolate.js",
  "./js/metro.js",
  "./js/elevation.js",
  "./js/vendor/mgrs.js",
  "./manifest.webmanifest",
  "./icons/icon.svg",
  "./data/shared.json",
  "./data/charge_5.json",
  "./data/charge_6.json",
  "./data/charge_7.json",
  "./data/METRAP.txt",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE).then((cache) => cache.addAll(SHELL)).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

function isDataRequest(url) {
  return url.pathname.includes("/data/");
}

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);

  if (isDataRequest(url)) {
    // data: network-first, cache fallback
    event.respondWith(
      fetch(req)
        .then((res) => {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(req, copy));
          return res;
        })
        .catch(() => caches.match(req))
    );
    return;
  }

  // static: cache-first
  event.respondWith(
    caches.match(req).then((cached) => {
      if (cached) return cached;
      return fetch(req).then((res) => {
        if (res.ok && url.origin === self.location.origin) {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(req, copy));
        }
        return res;
      });
    })
  );
});
