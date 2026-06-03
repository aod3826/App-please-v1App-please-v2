/**
 * sw.js — Service Worker สำหรับคลังรูปภาพกิจกรรม
 * Strategy: Cache-First สำหรับ Shell assets, Network-First สำหรับ GAS API
 */

const CACHE_NAME    = "gallery-v1";
const GAS_ORIGIN    = "script.google.com";

// ── Assets ที่ Cache ไว้ตั้งแต่ติดตั้ง (App Shell) ──────────────
const SHELL_ASSETS = [
  "/",
  "/index.html",
  "/manifest.json",
  "/icon-192x192.png",
  "/icon-512x512.png",
];

// ── Offline Fallback Page ────────────────────────────────────────
const OFFLINE_HTML = `<!DOCTYPE html>
<html lang="th">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1"/>
  <title>ไม่มีการเชื่อมต่ออินเทอร์เน็ต</title>
  <style>
    body { margin:0; background:#0d1117; color:#e6edf3; font-family:'Prompt',sans-serif;
           display:flex; align-items:center; justify-content:center; min-height:100vh;
           flex-direction:column; gap:1rem; text-align:center; padding:2rem; }
    .icon { font-size:4rem; margin-bottom:.5rem; }
    h1 { font-size:1.25rem; font-weight:600; margin:0; color:#fff; }
    p  { font-size:.875rem; color:#8b949e; margin:0; max-width:280px; }
    button { margin-top:.5rem; padding:.6rem 1.5rem; border-radius:8px;
             background:#22c55e; color:#fff; border:none; font-size:.875rem;
             font-family:'Prompt',sans-serif; cursor:pointer; font-weight:500; }
    button:hover { background:#16a34a; }
  </style>
</head>
<body>
  <div class="icon">📡</div>
  <h1>ไม่มีการเชื่อมต่ออินเทอร์เน็ต</h1>
  <p>กรุณาตรวจสอบการเชื่อมต่อแล้วลองใหม่อีกครั้ง</p>
  <button onclick="location.reload()">ลองใหม่</button>
</body>
</html>`;

// ══════════════════════════════════════════════
//  INSTALL — Cache App Shell
// ══════════════════════════════════════════════
self.addEventListener("install", event => {
  console.log("[SW] Installing v" + CACHE_NAME);
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      console.log("[SW] Caching App Shell");
      return cache.addAll(SHELL_ASSETS);
    }).then(() => self.skipWaiting())
  );
});

// ══════════════════════════════════════════════
//  ACTIVATE — Clean old caches
// ══════════════════════════════════════════════
self.addEventListener("activate", event => {
  console.log("[SW] Activating v" + CACHE_NAME);
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys
          .filter(key => key !== CACHE_NAME)
          .map(key => {
            console.log("[SW] Deleting old cache:", key);
            return caches.delete(key);
          })
      )
    ).then(() => self.clients.claim())
  );
});

// ══════════════════════════════════════════════
//  FETCH — Routing Strategy
// ══════════════════════════════════════════════
self.addEventListener("fetch", event => {
  const { request } = event;
  const url = new URL(request.url);

  // ── 1. GAS API → Network-First (ข้อมูลต้อง fresh เสมอ) ──────
  if (url.hostname.includes(GAS_ORIGIN)) {
    event.respondWith(networkFirst(request));
    return;
  }

  // ── 2. Google Drive Images → Cache-First ─────────────────────
  if (url.hostname.includes("drive.google.com") ||
      url.hostname.includes("lh3.googleusercontent.com")) {
    event.respondWith(cacheFirst(request));
    return;
  }

  // ── 3. Google Fonts → Cache-First ────────────────────────────
  if (url.hostname.includes("fonts.googleapis.com") ||
      url.hostname.includes("fonts.gstatic.com")) {
    event.respondWith(cacheFirst(request));
    return;
  }

  // ── 4. Tailwind CDN → Cache-First ────────────────────────────
  if (url.hostname.includes("cdn.tailwindcss.com")) {
    event.respondWith(cacheFirst(request));
    return;
  }

  // ── 5. App Shell (HTML/Icons/Manifest) → Cache-First ─────────
  if (request.destination === "document" ||
      request.destination === "image"    ||
      SHELL_ASSETS.some(a => url.pathname.endsWith(a))) {
    event.respondWith(cacheFirst(request, true));
    return;
  }

  // ── 6. Default → Network with Cache fallback ─────────────────
  event.respondWith(networkFirst(request));
});

// ══════════════════════════════════════════════
//  STRATEGY: Cache-First
//  ดึงจาก Cache → ถ้าไม่มีดึง Network แล้ว Cache ไว้
// ══════════════════════════════════════════════
async function cacheFirst(request, offlineFallback = false) {
  const cached = await caches.match(request);
  if (cached) return cached;

  try {
    const response = await fetch(request);
    if (response.ok && response.status === 200) {
      const cache = await caches.open(CACHE_NAME);
      cache.put(request, response.clone());  // clone เพราะ response อ่านได้ครั้งเดียว
    }
    return response;
  } catch (err) {
    if (offlineFallback) {
      return new Response(OFFLINE_HTML, {
        headers: { "Content-Type": "text/html; charset=utf-8" }
      });
    }
    throw err;
  }
}

// ══════════════════════════════════════════════
//  STRATEGY: Network-First
//  ดึง Network ก่อน → ถ้า fail ดึงจาก Cache
// ══════════════════════════════════════════════
async function networkFirst(request) {
  const cache = await caches.open(CACHE_NAME);
  try {
    const response = await fetch(request);
    if (response.ok) {
      cache.put(request, response.clone());
    }
    return response;
  } catch (err) {
    const cached = await cache.match(request);
    if (cached) return cached;
    // ถ้าเป็น navigation request ให้แสดง offline page
    if (request.mode === "navigate") {
      return new Response(OFFLINE_HTML, {
        headers: { "Content-Type": "text/html; charset=utf-8" }
      });
    }
    throw err;
  }
}

// ══════════════════════════════════════════════
//  MESSAGE — รับคำสั่งจาก Frontend
// ══════════════════════════════════════════════
self.addEventListener("message", event => {
  // SKIP_WAITING: บังคับ activate SW เวอร์ชันใหม่ทันที
  if (event.data && event.data.type === "SKIP_WAITING") {
    self.skipWaiting();
  }
  // CLEAR_CACHE: ล้าง cache ทั้งหมด (ใช้ตอน debug)
  if (event.data && event.data.type === "CLEAR_CACHE") {
    caches.delete(CACHE_NAME).then(() => {
      event.ports[0].postMessage({ cleared: true });
    });
  }
});
