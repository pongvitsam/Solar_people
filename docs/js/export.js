/* ส่งออกรายชื่อเป็น Excel (.xlsx) / PDF — ใช้ร่วมกันหน้าเช็คอิน (เจ้าหน้าที่) และหน้าจัดการผู้ใช้
 * ปุ่มใดมี data-export="people|arrived|pending" และ data-fmt="xlsx|pdf" จะถูกผูกให้อัตโนมัติ
 * ต้องโหลด xlsx.full.min.js (Excel) และ common.js ก่อน */
(function () {
  const { api, session, escapeHtml, toast } = window.App;

  const expFmt = new Intl.DateTimeFormat('th-TH', { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit', timeZone: 'Asia/Bangkok' });
  const stamp = () => new Date().toLocaleDateString('sv-SE', { timeZone: 'Asia/Bangkok' });
  const fmtT = (ms) => (ms ? expFmt.format(new Date(ms)) : '');

  async function companies() {
    await window.App.list.refresh();
    return window.App.list.current;
  }

  const SPECS = {
    // รายคน: มีข้อมูลส่วนบุคคล ดึงผ่านคำสั่งที่ต้องล็อกอิน
    people: {
      title: 'รายชื่อผู้เข้างาน', sheet: 'ผู้เข้างาน', unit: 'คน', empty: 'ยังไม่มีรายชื่อผู้เข้างาน',
      head: ['ลำดับ', 'บริษัท', 'PEA No.', 'ชื่อ-นามสกุล', 'ตำแหน่ง', 'อีเมล', 'เบอร์โทร', 'ช่องทาง', 'เวลาลงทะเบียน'],
      widths: [6, 38, 12, 26, 22, 28, 14, 20, 22],
      async load(token) {
        const r = await api.post('exportPeople', { token });
        const rows = r.people.slice().sort((a, b) => a.company.localeCompare(b.company, 'th') || a.time - b.time);
        return rows.map((p, i) => [i + 1, p.company, p.id.indexOf('X-') === 0 ? 'นอกรายชื่อ' : p.id, p.name, p.position, p.email, p.phone, p.channel, fmtT(p.time)]);
      },
    },
    // บริษัทที่ลงทะเบียนแล้ว (รวมนอกรายชื่อ) — ไม่มีข้อมูลส่วนบุคคล
    arrived: {
      title: 'รายชื่อบริษัทที่ลงทะเบียนแล้ว', sheet: 'บริษัทที่ลงทะเบียน', unit: 'บริษัท', empty: 'ยังไม่มีบริษัทลงทะเบียน',
      head: ['ลำดับ', 'PEA No.', 'ชื่อบริษัท', 'เลขที่ มท.', 'จำนวนคน', 'เวลาเข้างาน', 'ผู้บันทึก'],
      widths: [6, 12, 46, 18, 10, 22, 30],
      extra: (rows) => `รวม ${rows.reduce((n, r) => n + (Number(r[4]) || 0), 0).toLocaleString('th-TH')} คน`,
      async load() {
        const d = await companies();
        const all = d.companies.filter((c) => c.arrived).concat(d.offlist || []).sort((a, b) => (a.time || 0) - (b.time || 0));
        return all.map((c, i) => [i + 1, c.offlist ? 'นอกรายชื่อ' : c.id, c.name, c.mt || '', c.count, fmtT(c.time), c.by || '']);
      },
    },
    // บริษัทที่ยังไม่มา
    pending: {
      title: 'รายชื่อบริษัทที่ยังไม่มา', sheet: 'บริษัทที่ยังไม่มา', unit: 'บริษัท', empty: 'ลงทะเบียนครบทุกบริษัทแล้ว',
      head: ['ลำดับ', 'PEA No.', 'ชื่อบริษัท', 'เลขที่ มท.'],
      widths: [6, 12, 54, 18],
      async load() {
        const d = await companies();
        return d.companies.filter((c) => !c.arrived).map((c, i) => [i + 1, c.id, c.name, c.mt || '']);
      },
    },
  };

  function toXlsx(spec, rows) {
    if (!window.XLSX) { toast('โหลดตัวสร้างไฟล์ Excel ไม่สำเร็จ ลองรีเฟรชหน้า', 'error'); return; }
    const ws = XLSX.utils.aoa_to_sheet([spec.head].concat(rows)); // ค่าทุกช่องเป็นข้อความ/ตัวเลข ไม่ถูกตีเป็นสูตร
    ws['!cols'] = spec.widths.map((w) => ({ wch: w }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, spec.sheet);
    XLSX.writeFile(wb, `${spec.title}_${stamp()}.xlsx`);
    toast(`ส่งออก Excel ${rows.length} รายการแล้ว`, 'ok');
  }

  // PDF: เปิดหน้าสำหรับพิมพ์ แล้วให้เลือก "บันทึกเป็น PDF" (ฟอนต์ไทยแสดงถูกต้องโดยเบราว์เซอร์)
  function toPdf(spec, rows) {
    const w = window.open('', '_blank');
    if (!w) { toast('เบราว์เซอร์บล็อกหน้าต่างใหม่ กรุณาอนุญาต pop-up แล้วลองอีกครั้ง', 'error'); return; }
    const head = spec.head.map((h) => `<th>${escapeHtml(h)}</th>`).join('');
    const body = rows.map((r) => `<tr>${r.map((c) => `<td>${escapeHtml(c)}</td>`).join('')}</tr>`).join('');
    const sub = `ทั้งหมด ${rows.length.toLocaleString('th-TH')} ${spec.unit}${spec.extra ? ' · ' + spec.extra(rows) : ''} · ส่งออกเมื่อ ${expFmt.format(new Date())}`;
    w.document.open();
    w.document.write(`<!doctype html><html lang="th"><head><meta charset="utf-8"><title>${escapeHtml(spec.title)} ${stamp()}</title>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@400;600&display=swap" rel="stylesheet">
<style>
@page { size: A4 landscape; margin: 10mm; }
body { font-family: 'IBM Plex Sans Thai', sans-serif; font-size: 10px; color: #111; margin: 0; }
h1 { font-size: 15px; margin: 0 0 2px; } p { margin: 0 0 8px; color: #555; }
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid #999; padding: 3px 5px; text-align: left; vertical-align: top; word-break: break-word; }
th { background: #eee; } thead { display: table-header-group; } tr { break-inside: avoid; }
</style></head><body>
<h1>${escapeHtml(spec.title)} · โซลาร์ภาคประชาชน</h1>
<p>${escapeHtml(sub)}</p>
<table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>
</body></html>`);
    w.document.close();
    const go = () => { w.focus(); w.print(); };
    if (w.document.fonts && w.document.fonts.ready) w.document.fonts.ready.then(() => setTimeout(go, 300)); else setTimeout(go, 1000);
  }

  /** onAuthFail(err): เรียกเมื่อเซสชันหมดอายุ (แต่ละหน้าจัดการเอง เช่น กลับไปหน้าล็อกอิน) */
  function bind(onAuthFail) {
    document.querySelectorAll('[data-export]').forEach((btn) => {
      btn.addEventListener('click', async () => {
        const spec = SPECS[btn.dataset.export];
        if (!spec) return;
        const s = session.get();
        if (!s || !s.token) { onAuthFail && onAuthFail({ code: 'SESSION_EXPIRED', message: 'กรุณาเข้าสู่ระบบใหม่' }); return; }
        const label = btn.textContent;
        btn.disabled = true;
        btn.textContent = 'กำลังเตรียมข้อมูล…';
        try {
          const rows = await spec.load(s.token);
          if (!rows.length) { toast(spec.empty, 'error'); return; }
          (btn.dataset.fmt === 'pdf' ? toPdf : toXlsx)(spec, rows);
        } catch (err) {
          if (err.code === 'SESSION_EXPIRED' && onAuthFail) onAuthFail(err);
          else toast(err.message || 'ส่งออกไม่สำเร็จ', 'error');
        } finally {
          btn.disabled = false;
          btn.textContent = label;
        }
      });
    });
  }

  window.App.exports = { bind };
})();
