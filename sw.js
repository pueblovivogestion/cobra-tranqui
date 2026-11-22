// Service worker: guarda SOLO los archivos de la app (para que abra rápido y se pueda instalar).
// Los cobros NUNCA salen de caché: siempre se piden en vivo al servidor.
const VERSION = 'cobra-tranqui-v15';
const ARCHIVOS = ['./', 'index.html', 'caja.html', 'config.js', 'manifest-caja.webmanifest', 'manifest-panel.webmanifest',
  'iconos/icono-192.png', 'iconos/icono-512.png', 'iconos/apple-touch-icon.png', 'iconos/favicon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(ARCHIVOS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  // Solo archivos propios con GET; todo lo demás (API, Supabase, CDN) va directo a la red.
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  // Red primero (para tener siempre la última versión); si no hay internet, lo guardado.
  e.respondWith(fetch(e.request).then(r => {
    if (r.ok) { const copia = r.clone(); caches.open(VERSION).then(c => c.put(e.request, copia)); }
    return r;
  }).catch(() => caches.match(e.request, { ignoreSearch: true })));
});
