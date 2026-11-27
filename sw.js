/* Orbit service worker: makes the app installable and opens instantly.
   Pages load fresh from the network when online (so updates arrive right away),
   with the cached copy as a fallback. Data calls to Google are never cached. */
const CACHE = 'orbit-v61';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png', './icons/apple-touch-icon.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE && k !== 'orbit-push').map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;
  e.respondWith(fetch(req).then(res => {
    const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); return res;
  }).catch(() => caches.match(req).then(r => r || caches.match('./index.html'))));
});

/* Daily summary notifications. The push carries no data: on arrival we ask Orbit for this
   device's summary (using the device key saved when notifications were turned on) and show it. */
self.addEventListener('push', e => {
  e.waitUntil((async () => {
    let n = { title: 'Orbit', body: 'Open Orbit to see today’s tasks.', url: '#/mytasks' };
    try {
      const c = await caches.open('orbit-push'), r = await c.match('./__orbit_push'), cfg = r ? await r.json() : null;
      if (cfg && cfg.api && cfg.key) {
        const res = await fetch(cfg.api, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ action: 'pushSummary', key: cfg.key }) });
        const j = await res.json(); if (j && j.ok && j.data) n = j.data;
      }
    } catch (err) { }
    await self.registration.showNotification(n.title, { body: n.body, tag: 'orbit-daily', renotify: true, icon: './icons/icon-192.png', badge: './icons/icon-192.png', data: { url: n.url || '#/mytasks' } });
  })());
});
self.addEventListener('notificationclick', e => {
  e.notification.close();
  const hash = (e.notification.data && e.notification.data.url) || '#/mytasks', target = new URL('./' + hash, self.registration.scope).href;
  e.waitUntil((async () => {
    const wins = await clients.matchAll({ type: 'window', includeUncontrolled: true });
    for (const w of wins) { if (w.url.startsWith(self.registration.scope)) { await w.focus(); w.postMessage({ type: 'orbit-open', hash }); return; } }
    await clients.openWindow(target);
  })());
});
