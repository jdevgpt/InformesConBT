/* Informe BT · funcionamiento sin conexión (JDev) */
const VERSION = 'informesconbt-v1';
const LOCALES = ["./", "./index.html", "./apple-touch-icon.png", "./direcciones.json", "./icon-192.png", "./icon-512.png", "./icon-maskable-512.png", "./icon.svg", "./manifest.webmanifest", "./motor.js", "./plantilla.docx"];
const EXTERNOS = ["https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js", "https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js", "https://fonts.googleapis.com/css2?family=Barlow:wght@500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap"];

self.addEventListener('install', e => {
  e.waitUntil((async () => {
    const c = await caches.open(VERSION);
    await c.addAll(LOCALES);
    await Promise.all(EXTERNOS.map(u => fetch(u, { mode: 'cors' }).then(r => r.ok && c.put(u, r)).catch(() => {})));
    self.skipWaiting();
  })());
});

self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    for (const k of await caches.keys()) if (k !== VERSION) await caches.delete(k);
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // Página: primero la red (para recibir actualizaciones); sin cobertura, la copia guardada
  if (req.mode === 'navigate') {
    e.respondWith((async () => {
      const c = await caches.open(VERSION);
      try {
        const ctl = new AbortController(); const t = setTimeout(() => ctl.abort(), 5000);
        const r = await fetch(req, { signal: ctl.signal }); clearTimeout(t);
        if (r.ok) c.put('./index.html', r.clone());
        return r;
      } catch (err) {
        return (await c.match('./index.html')) || (await c.match('./')) || new Response('Sin conexión', { status: 504 });
      }
    })());
    return;
  }

  // Archivos de la app: se sirven al momento desde lo guardado y se actualizan por detrás
  if (url.origin === self.location.origin) {
    e.respondWith((async () => {
      const c = await caches.open(VERSION);
      const g = await c.match(req, { ignoreSearch: true });
      const red = fetch(req).then(r => { if (r.ok) c.put(req, r.clone()); return r; }).catch(() => null);
      if (g) { e.waitUntil(red); return g; }
      return (await red) || new Response('', { status: 504, statusText: 'Sin conexión' });
    })());
    return;
  }

  // Librerías y fuentes externas: lo guardado primero
  if (/cdnjs\.cloudflare\.com|fonts\.googleapis\.com|fonts\.gstatic\.com|cdn\.jsdelivr\.net|unpkg\.com/.test(url.hostname)) {
    e.respondWith((async () => {
      const c = await caches.open(VERSION);
      const g = await c.match(req.url);
      if (g) return g;
      try { const r = await fetch(req); if (r.ok || r.type === 'opaque') c.put(req.url, r.clone()); return r; }
      catch (err) { return new Response('', { status: 504 }); }
    })());
  }
});
