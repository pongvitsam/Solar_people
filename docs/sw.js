/* Service worker — ลงทะเบียนเข้างาน กฟภ.
 *
 * กลยุทธ์:
 * - หน้าเว็บ/ไฟล์ของแอป (same-origin): ดึงจากเครือข่ายก่อนเสมอ (ได้เวอร์ชันล่าสุด) ถ้าออฟไลน์/ช้าเกินไปค่อยใช้ของที่เก็บไว้
 * - Google Fonts: ใช้ของที่เก็บไว้ก่อน (ไม่ค่อยเปลี่ยน)
 * - API (script.google.com) และทุกอย่างอื่น: ไม่แตะ ไม่ cache — ข้อมูลต้องสดเสมอ
 * เปลี่ยน VERSION เมื่อเพิ่ม/ลบไฟล์ใน SHELL
 */
const VERSION = 'v10';
const SHELL_CACHE = 'shell-' + VERSION;
const FONT_CACHE = 'fonts-' + VERSION;
const NETWORK_TIMEOUT_MS = 4000;

const SHELL = [
  './', './index.html', './checkin.html', './register.html', './admin.html', './staff-register.html',
  './styles.css', './js/config.js', './js/common.js', './js/export.js', './manifest.json',
  './img/pea-logo.png', './img/pea-emblem.png', './img/icon-192.png', './img/apple-touch-icon.png', './img/favicon-32.png',
];

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(SHELL_CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== SHELL_CACHE && k !== FONT_CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

/** เครือข่ายก่อน ถ้าล้มเหลวหรือเกินเวลา ใช้ของใน cache (เทียบโดยไม่สน ?v=...) */
async function networkFirst(req) {
  const cache = await caches.open(SHELL_CACHE);
  const url = new URL(req.url);
  const key = url.origin + url.pathname; // เก็บตาม path อย่างเดียว ไม่ให้ cache บวมตาม ?v= / ?source=
  const net = fetch(req).then((res) => {
    if (res && res.status === 200 && res.type === 'basic') cache.put(key, res.clone()).catch(() => {});
    return res;
  });
  net.catch(() => {}); // กัน unhandled rejection เมื่อเลิกรอแล้ว
  try {
    return await Promise.race([
      net,
      new Promise((_, reject) => setTimeout(() => reject(new Error('timeout')), NETWORK_TIMEOUT_MS)),
    ]);
  } catch (err) {
    const hit = await cache.match(key, { ignoreSearch: true });
    if (hit) return hit;
    // ไม่มีของใน cache (เปิดครั้งแรกบนเน็ตช้า): รอเครือข่ายต่อ ดีกว่าให้หน้าโหลดไม่ขึ้น
    try { return await net; } catch (e) { /* ออฟไลน์จริง */ }
    if (req.mode === 'navigate') {
      const fallback = await cache.match('./checkin.html', { ignoreSearch: true });
      if (fallback) return fallback;
    }
    throw err;
  }
}

async function cacheFirst(req) {
  const cache = await caches.open(FONT_CACHE);
  const hit = await cache.match(req);
  if (hit) return hit;
  const res = await fetch(req);
  if (res && (res.ok || res.type === 'opaque')) cache.put(req, res.clone());
  return res;
}

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return; // POST (API) ผ่านตรง
  const url = new URL(req.url);
  if (url.origin === self.location.origin) {
    event.respondWith(networkFirst(req));
  } else if (url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com') {
    event.respondWith(cacheFirst(req));
  }
  // อื่น ๆ (API, CDN สร้าง QR): ให้เบราว์เซอร์จัดการเอง ไม่ cache
});
