const CACHE_NAME = "contabtech-shell-v3";
const SHELL_FILES = [
  "./",
  "./manifest.webmanifest",
  "./icon.svg",
  "./icon-192.png",
  "./icon-512.png",
  "./apple-touch-icon.png",
];

// Pré-carrega também os arquivos JS/CSS referenciados pela página, para abrir offline desde o primeiro uso.
async function precacheShell() {
  const cache = await caches.open(CACHE_NAME);
  await cache.addAll(SHELL_FILES);
  const html = await (await fetch("./", { cache: "no-cache" })).text();
  const assets = [...html.matchAll(/(?:src|href)="(\.?\/?assets\/[^"]+)"/g)].map((match) =>
    match[1].startsWith("./") ? match[1] : `./${match[1].replace(/^\//, "")}`,
  );
  await cache.addAll(assets);
}

self.addEventListener("install", (event) => {
  event.waitUntil(precacheShell());
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((keys) =>
        Promise.all(
          keys
            .filter((key) => key.startsWith("contabtech-") && key !== CACHE_NAME)
            .map((key) => caches.delete(key)),
        ),
      ),
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const request = event.request;
  const url = new URL(request.url);
  if (request.method !== "GET" || url.origin !== self.location.origin) return;

  if (request.mode === "navigate") {
    event.respondWith(
      fetch(request).catch(async () => {
        const cachedPage = await caches.match(request, { ignoreSearch: true, ignoreVary: true });
        return cachedPage || caches.match("./", { ignoreVary: true });
      }),
    );
    return;
  }

  event.respondWith(
    caches.match(request, { ignoreVary: true }).then(
      (cached) =>
        cached ||
        fetch(request).then(async (response) => {
          if (response.ok) {
            const cache = await caches.open(CACHE_NAME);
            await cache.put(request, response.clone());
          }
          return response;
        }),
    ),
  );
});
