const CACHE = "royal-agro-v2";
const ASSETS = [
  "./", "./index.html", "./manifest.webmanifest",
  "https://cdn.tailwindcss.com",
  "https://unpkg.com/react@18/umd/react.production.min.js",
  "https://unpkg.com/react-dom@18/umd/react-dom.production.min.js",
  "https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js",
  "https://unpkg.com/papaparse@5.4.1/papaparse.min.js",
  "https://unpkg.com/lucide@latest/dist/umd/lucide.min.js"
];
self.addEventListener("install", function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) {
    return Promise.all(ASSETS.map(function (u) { return c.add(u).catch(function(){}); }));
  }).then(function () { return self.skipWaiting(); }));
});
self.addEventListener("activate", function (e) {
  e.waitUntil(caches.keys().then(function (ks) {
    return Promise.all(ks.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});
// Network-first: always try the network when online (avoids stale/broken cache), fall back to cache offline.
self.addEventListener("fetch", function (e) {
  if (e.request.method !== "GET") return;
  e.respondWith(
    fetch(e.request).then(function (resp) {
      var cp = resp.clone();
      caches.open(CACHE).then(function (c) { c.put(e.request, cp).catch(function(){}); });
      return resp;
    }).catch(function () {
      return caches.match(e.request).then(function (r) { return r || caches.match("./index.html"); });
    })
  );
});
