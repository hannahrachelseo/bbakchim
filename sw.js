/* 한나T 빡침모드 ON — 자동 업데이트용 서비스 워커
   화면(index.html)은 항상 인터넷에서 먼저 받아오고, 인터넷이 없을 때만 저장해 둔 화면을 보여줘요. */
const CACHE = 'bbak-v4';
const SHELL = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== location.origin) return; // 구글 서버·글꼴·엑셀 도구는 건드리지 않음
  e.respondWith(
    fetch(req, { cache: 'no-store' }).then(res => {
      if (res && res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
      return res;
    }).catch(() => caches.match(req).then(hit => hit || caches.match('./index.html')))
  );
});
