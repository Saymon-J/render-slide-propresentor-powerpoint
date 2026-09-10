// Service worker: precache всего приложения → полная работа оффлайн.
// Версия кэша = версия приложения; при релизе поднять вместе с app.py __version__.
const V = "1.2.0";

const ASSETS = [
  "./",
  "./index.html",
  "./manifest.webmanifest",
  "./favicon.ico",
  "./vendor/jszip.min.js",
  "./vendor/protobuf.min.js",
  "./vendor/pptxgen.bundle.js",
  "./js/books.js",
  "./js/core.js",
  "./js/pro.js",
  "./js/pptx.js",
  "./js/docx.js",
  "./js/api.js",
  "./js/pro-schema.json",
  "./assets/theme.pro",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./icons/maskable-512.png",
];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(V).then((c) => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== V).map((k) => caches.delete(k))))
      .then(() => self.clients.claim()),
  );
});

self.addEventListener("fetch", (e) => {
  if (e.request.method !== "GET") return;
  e.respondWith(
    caches.match(e.request, { ignoreSearch: true }).then((hit) => hit || fetch(e.request).then((r) => {
      const copy = r.clone();
      if (r.ok && new URL(e.request.url).origin === location.origin) {
        caches.open(V).then((c) => c.put(e.request, copy));
      }
      return r;
    }).catch(() => e.request.mode === "navigate" ? caches.match("./index.html") : undefined)),
  );
});
