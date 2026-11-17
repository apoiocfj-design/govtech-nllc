/**
 * SERVICE WORKER - GOVTECH NLLC PWA
 * Cache Offline e Resiliência Operacional para Redes Restritas ou Remotas
 */

const CACHE_NAME = "govtech-nllc-v2.0";
const OFFLINE_ASSETS = [
  "/",
  "/index.html",
  "/preview_app.html",
  "/manifest.json",
  "/favicon.ico",
  "/icon.png",
  "/icon_transparent.png",
  "/brasao_oficial.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      console.log("[PWA SW] Pre-caching offline assets");
      return cache.addAll(OFFLINE_ASSETS).catch((e) => {
        console.warn("[PWA SW] Pre-cache partial failure:", e);
      });
    })
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) {
            console.log("[PWA SW] Removing old cache:", key);
            return caches.delete(key);
          }
        })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  // Ignora requisições de API de backend para não cachear dados desatualizados
  if (event.request.url.includes("/api/")) {
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      if (cachedResponse) {
        return cachedResponse;
      }
      return fetch(event.request).catch(() => {
        // Fallback offline para navegação principal
        if (event.request.mode === "navigate") {
          return caches.match("/index.html");
        }
      });
    })
  );
});
