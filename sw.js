// ═══════════════════════════════════════════════════════════
//  Service Worker — ניהול תיקים ר.ר שמאות
//  אסטרטגיה: Cache-First לקבצים סטטיים, Network-First לשאר
// ═══════════════════════════════════════════════════════════

const CACHE_NAME = 'rr-v21';

// קבצים שיישמרו בcache בעת ההתקנה
const STATIC_ASSETS = [
  '/rrr/',
  '/rrr/index.html',
  '/rrr/Logo_smol.png',
  '/rrr/Logo_smol_512.png',
  '/rrr/Logo.png',
  '/rrr/manifest.json',
];

// דומיינים שתמיד יעברו ברשת (Google APIs)
const NETWORK_ONLY_ORIGINS = [
  'googleapis.com',
  'accounts.google.com',
  'fonts.googleapis.com',
  'fonts.gstatic.com',
  'cdnjs.cloudflare.com',
  'script.google.com',
  'drive.google.com',
];

// ─── התקנה: שמור קבצים סטטיים ───────────────────────────
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) =>
      // כל קובץ בנפרד — קובץ חסר לא ישבור את כל ההתקנה
      Promise.all(STATIC_ASSETS.map((u) => cache.add(u).catch(() => {})))
    )
  );
  self.skipWaiting();
});

// ─── הפעלה: מחק caches ישנים ─────────────────────────────
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key !== CACHE_NAME)
          .map((key) => caches.delete(key))
      )
    )
  );
  self.clients.claim();
});

// ─── בקשות: Network-First לAPI, Cache-First לסטטי ────────
self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // בקשות לGoogle APIs — תמיד ברשת, לא נגע בהן
  const isNetworkOnly = NETWORK_ONLY_ORIGINS.some((origin) =>
    url.hostname.includes(origin)
  );
  if (isNetworkOnly) return;

  // קבצים סטטיים — Cache-First (אם אין cache, שלוף מרשת)
  const isStatic = STATIC_ASSETS.some(
    (path) => url.pathname === path || url.pathname.endsWith('.png') ||
              url.pathname.endsWith('.js') || url.pathname.endsWith('.css')
  );

  if (isStatic) {
    event.respondWith(
      caches.match(event.request).then((cached) => {
        return cached || fetch(event.request).then((response) => {
          if (response.ok) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
          }
          return response;
        });
      })
    );
    return;
  }

  // כל השאר — Network-First עם fallback לcache
  event.respondWith(
    fetch(event.request)
      .then((response) => {
        if (response.ok) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
        }
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});
