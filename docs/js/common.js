/* ฟังก์ชันใช้ร่วมกันทุกหน้า: เรียก API, session, ค้นหา, toast */
(function () {
  const cfg = window.APP_CONFIG;
  const SESSION_KEY = 'solar_people_session';

  class ApiError extends Error {
    constructor(code, message) {
      super(message);
      this.code = code;
    }
  }

  function assertConfigured() {
    if (!cfg.API_URL || cfg.API_URL.includes('PASTE_')) {
      throw new ApiError('NO_CONFIG', 'ยังไม่ได้ตั้งค่า API_URL ในไฟล์ js/config.js');
    }
  }

  async function parse(res) {
    let j;
    try {
      j = await res.json();
    } catch (e) {
      throw new ApiError('BAD_RESPONSE', 'เซิร์ฟเวอร์ตอบกลับไม่ถูกต้อง (ตรวจสอบการ Deploy Web App)');
    }
    if (!j.ok) throw new ApiError(j.code, j.error || 'เกิดข้อผิดพลาด');
    return j.data;
  }

  const api = {
    async get(action, params = {}) {
      assertConfigured();
      const url = new URL(cfg.API_URL);
      url.searchParams.set('action', action);
      Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));
      url.searchParams.set('_', Date.now());
      let res;
      try {
        res = await fetch(url.toString(), { cache: 'no-store' });
      } catch (e) {
        throw new ApiError('NETWORK', 'เชื่อมต่อเซิร์ฟเวอร์ไม่ได้ ตรวจสอบอินเทอร์เน็ต');
      }
      return parse(res);
    },
    // ใช้ text/plain เพื่อไม่ให้เบราว์เซอร์ส่ง CORS preflight (Apps Script ไม่รองรับ OPTIONS)
    async post(action, payload = {}) {
      assertConfigured();
      let res;
      try {
        res = await fetch(cfg.API_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify(Object.assign({ action }, payload)),
        });
      } catch (e) {
        throw new ApiError('NETWORK', 'เชื่อมต่อเซิร์ฟเวอร์ไม่ได้ ตรวจสอบอินเทอร์เน็ต');
      }
      return parse(res);
    },
  };

  const session = {
    get() {
      try {
        return JSON.parse(localStorage.getItem(SESSION_KEY) || 'null');
      } catch (e) {
        return null;
      }
    },
    set(s) {
      try { localStorage.setItem(SESSION_KEY, JSON.stringify(s)); } catch (e) { /* private mode */ }
    },
    clear() {
      try { localStorage.removeItem(SESSION_KEY); } catch (e) { /* ignore */ }
    },
  };

  function escapeHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
    }[c]));
  }

  // ค้นหาแบบไม่สนตัวพิมพ์/ช่องว่าง และไม่สนคำนำหน้า เช่น บริษัท, จำกัด, (มหาชน)
  function normalize(s) {
    return String(s || '')
      .toLowerCase()
      .replace(/บริษัท|บจก\.?|จำกัด|\(มหาชน\)|มหาชน/g, '')
      .replace(/[\s.,()\-]/g, '');
  }

  // ค้นหาจากชื่อบริษัท หรือ PEA No. (พิมพ์ "P00024", "00024" หรือ "24" ก็ได้)
  function matches(company, query) {
    const raw = String(query || '').trim().toLowerCase().replace(/\s+/g, '');
    if (!raw) return true;
    const id = String(company.id).toLowerCase();
    if (id.includes(raw)) return true;
    if (/^\d+$/.test(raw) && id.replace(/\D/g, '').replace(/^0+/, '').includes(raw.replace(/^0+/, '') || '0')) return true;
    const q = normalize(query);
    if (!q) return String(company.name).toLowerCase().replace(/\s+/g, '').includes(raw); // เช่น พิมพ์แค่ "บริษัท"
    return normalize(company.name).includes(q);
  }

  const timeFmt = new Intl.DateTimeFormat('th-TH', { hour: '2-digit', minute: '2-digit', hour12: false, timeZone: 'Asia/Bangkok' });
  const fullFmt = new Intl.DateTimeFormat('th-TH', { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false, timeZone: 'Asia/Bangkok' });

  function fmtTime(iso) {
    if (!iso) return '';
    const d = new Date(iso);
    return isNaN(d) ? String(iso) : timeFmt.format(d) + ' น.';
  }
  function fmtClock(d) {
    return fullFmt.format(d || new Date());
  }

  function toast(message, type = 'info') {
    let wrap = document.querySelector('.toasts');
    if (!wrap) {
      wrap = document.createElement('div');
      wrap.className = 'toasts';
      wrap.setAttribute('role', 'status');
      wrap.setAttribute('aria-live', 'polite');
      document.body.appendChild(wrap);
    }
    const el = document.createElement('div');
    el.className = 'toast toast-' + type;
    el.textContent = message;
    wrap.appendChild(el);
    setTimeout(() => el.classList.add('out'), 3200);
    setTimeout(() => el.remove(), 3600);
  }

  // ดึงข้อมูลซ้ำเป็นช่วง ๆ หยุดเมื่อแท็บถูกซ่อน และดึงทันทีเมื่อกลับมา
  function poll(fn, ms) {
    let timer = null;
    const tick = async () => {
      clearTimeout(timer);
      if (!document.hidden) {
        try { await fn(); } catch (e) { /* แสดงผลใน fn แล้ว */ }
      }
      timer = setTimeout(tick, ms);
    };
    document.addEventListener('visibilitychange', () => { if (!document.hidden) tick(); });
    tick();
    return { now: tick };
  }

  // ใส่ชื่อหน่วยงาน/ชื่องานบนแถบหัวเว็บทุกหน้า
  document.addEventListener('DOMContentLoaded', () => {
    const org = document.getElementById('orgName');
    if (org) org.textContent = cfg.ORG_NAME || '';
    const ev = document.getElementById('eventName');
    if (ev && !ev.dataset.fixed) {
      // แต่ละวลีไม่ตัดกลางคำ ตัดบรรทัดได้เฉพาะที่เว้นวรรค
      ev.innerHTML = String(cfg.EVENT_NAME).split(/\s+/).map((w) => `<b class="nw">${escapeHtml(w)}</b>`).join(' ');
    }
  });

  /* ---------- รายชื่อบริษัท ----------
   * - ส่งเวอร์ชันล่าสุด (v) ไปด้วย ถ้าข้อมูลไม่เปลี่ยน เซิร์ฟเวอร์ตอบสั้น ๆ ไม่ต้องโหลดทั้งก้อน
   * - เก็บ snapshot ไว้ในเครื่อง เปิดหน้ามาแสดงผลได้ทันทีระหว่างรอข้อมูลใหม่
   */
  const SNAP_KEY = 'solar_people_snapshot_v2';

  function expand(r) {
    return {
      v: r.v,
      companies: r.d.c.map((x) => ({
        id: x[0], name: x[1], mt: x[2], arrived: !!x[3], count: x[4], time: x[5] || 0, by: x[6],
        staff: x[7] == null ? !!x[3] : !!x[7], // เจ้าหน้าที่เช็คอินให้
        self: x[8] || 0, // จำนวนคนที่ลงทะเบียนเอง
      })),
      summary: r.d.s,
    };
  }

  function summarize(companies) {
    const arrived = companies.filter((c) => c.arrived);
    return {
      total: companies.length,
      arrived: arrived.length,
      remaining: companies.length - arrived.length,
      people: arrived.reduce((s, c) => s + (c.count || 0), 0),
    };
  }

  const list = {
    current: null,
    snapshot() {
      try {
        const raw = localStorage.getItem(SNAP_KEY);
        if (raw) this.current = expand(JSON.parse(raw));
      } catch (e) { this.current = null; }
      return this.current;
    },
    /** รับผลจาก API (list/checkin/undo) คืนค่า true ถ้าข้อมูลเปลี่ยน */
    accept(r) {
      if (r.same && this.current) return false;
      this.current = expand(r);
      try { localStorage.setItem(SNAP_KEY, JSON.stringify(r)); } catch (e) { /* ignore */ }
      return true;
    },
    async refresh() {
      return this.accept(await api.get('list', this.current ? { f: 2, v: this.current.v } : { f: 2 }));
    },
  };

  window.App = { cfg, api, session, list, summarize, ApiError, escapeHtml, matches, fmtTime, fmtClock, toast, poll };
})();
