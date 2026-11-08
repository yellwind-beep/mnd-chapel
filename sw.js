/* 국방부 성당 앱 - 오프라인 대비용 서비스 워커
 * - 홈페이지 파일: 항상 인터넷에서 새로 받고, 끊겼을 때만 저장본을 보여줌
 * - 글꼴: 한 번 받으면 저장본 사용
 * - 공지·미사 시간(구글 스크립트)은 건드리지 않음 (홈페이지가 따로 저장해 둠)
 */
const CACHE = 'chapel-v25';
const CORE = ['/', '/index.html', '/manifest.webmanifest', '/icons/icon-192.png', '/icons/favicon.svg'];
const FONT_HOSTS = ['fonts.googleapis.com', 'fonts.gstatic.com', 'cdn.jsdelivr.net'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(CORE)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // 같은 사이트 파일: 인터넷 먼저, 실패하면 저장본
  if (url.origin === self.location.origin) {
    e.respondWith(
      fetch(req)
        .then((res) => {
          if (res.ok) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(req, copy)); }
          return res;
        })
        .catch(() => caches.match(req, { ignoreSearch: true })
          .then((hit) => hit || (req.mode === 'navigate' ? caches.match('/') : undefined)))
    );
    return;
  }

  // 글꼴: 저장본 먼저
  if (FONT_HOSTS.includes(url.hostname)) {
    e.respondWith(
      caches.match(req).then((hit) => hit || fetch(req).then((res) => {
        if (res.ok || res.type === 'opaque') { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(req, copy)); }
        return res;
      }))
    );
  }
});
