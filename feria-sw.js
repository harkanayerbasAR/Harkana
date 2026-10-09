/* Guarda la pantalla de venta de ferias en el celular para que abra aunque no haya señal.
   Solo toca /feria: el resto del sitio funciona como siempre. */
const CACHE = "harkana-feria-v1";
const PAGINAS = ["/feria"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(PAGINAS)).then(() => self.skipWaiting())); });
self.addEventListener("activate", e => { e.waitUntil(self.clients.claim()); });
self.addEventListener("fetch", e => {
  const u = new URL(e.request.url);
  if (e.request.method !== "GET" || u.origin !== location.origin) return;
  if (u.pathname !== "/feria" && u.pathname !== "/feria.html") return;
  // Primero lo guardado (abre al instante, con o sin señal); en segundo plano se actualiza si hay conexión.
  e.respondWith(caches.open(CACHE).then(c => c.match("/feria").then(guardada => {
    const red = fetch(e.request).then(r => { if (r && r.ok) c.put("/feria", r.clone()); return r; }).catch(() => guardada);
    return guardada || red;
  })));
});
