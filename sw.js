// 経費ノート Service Worker — 圏外でもアプリの画面を開けるようにする
// ・アプリの画面（HTML・アイコン）だけを端末に保存します。データ（スプレッドシート）には触れません。
// ・電波があるときは常に最新の画面を取りに行き、圏外のときだけ保存した画面を使います。
const CACHE = 'keiji-note-v3.2.5';
const CORE = ['keiji-kanri-sheets.html', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => Promise.all(CORE.map(u => c.add(u).catch(() => null)))).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k.startsWith('keiji-note-') && k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;                       // GASへの送信などはそのまま通す
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;        // 他のサイト（Google等）には関与しない
  const isPage = req.mode === 'navigate' || url.pathname.endsWith('.html');
  if (isPage) {
    // ネット優先（必ずサーバーに最新か確認する。6秒で応答がなければ保存した画面）
    e.respondWith((async () => {
      const cache = await caches.open(CACHE);
      try {
        const net = await Promise.race([fetch(req, { cache: 'no-cache' }), new Promise((_, rej) => setTimeout(() => rej(new Error('timeout')), 6000))]);
        if (net && net.ok) cache.put(req, net.clone()).catch(() => {});
        return net;
      } catch (err) {
        return (await cache.match(req, { ignoreSearch: true })) || (await cache.match('keiji-kanri-sheets.html')) || new Response('<h1>圏外です</h1><p>一度電波のある場所でアプリを開くと、次からは圏外でも開けます。</p>', { headers: { 'Content-Type': 'text/html; charset=utf-8' } });
      }
    })());
    return;
  }
  // アイコンなど：保存したものを優先
  e.respondWith(caches.match(req, { ignoreSearch: true }).then(hit => hit || fetch(req).then(res => { if (res.ok) caches.open(CACHE).then(c => c.put(req, res.clone())); return res; })));
});
