/**
 * นับยอดในชีต — ใส่ยอดนับจริงทีละหลายตัวบนคอมพิวเตอร์
 *
 *   1) เมนู คลังใน → สร้างแท็บนับยอด  : ได้แท็บ "นับยอด" ยาที่ใช้งานทุกตัว เรียงตามหมวดแล้วตามชื่อ
 *   2) กรอกจำนวนในช่อง "นับได้" (สีเหลือง) เฉพาะตัวที่นับ ตัวที่ว่างไว้ระบบจะข้าม
 *      ช่อง "วันหมดอายุ" ไม่บังคับ ว่างไว้ระบบใช้วันหมดอายุจากคลังยาใหญ่ (ใช้เฉพาะยาที่ยังไม่เคยมีล็อต)
 *   3) เมนู คลังใน → บันทึกยอดนับ      : ลงเป็น "นับจริง" ให้ทุกตัว เหมือนกดในแอป มีบันทึกในแท็บประวัติ
 *
 * ไม่แตะระบบคลังยาใหญ่ · กดบันทึกซ้ำได้ ตัวที่บันทึกแล้วจะถูกล้างช่อง "นับได้" และเขียนผลไว้ช่องสุดท้าย
 */
const COUNT_SHEET = 'นับยอด';
const COUNT_HEAD = ['รหัส', 'ชื่อยา', 'หน่วย', 'หมวด/ชั้น', 'ยอดในระบบ', 'นับได้', 'วันหมดอายุ (ไม่บังคับ)', 'ผล'];

function makeCountSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const db = loadDb();
  const snap = Core.snapshot(db, new Date());
  const bal = {}; snap.items.forEach(i => { bal[i.c] = i.has ? i.b : ''; });
  const rows = db.items.filter(i => i.active)
    .map(i => [i.code, i.name, i.unit, i.shelf || '', bal[i.code], '', '', ''])
    .sort((a, b) => String(a[3]).localeCompare(String(b[3]), 'th') || String(a[1]).localeCompare(String(b[1]), 'en'));
  let sh = ss.getSheetByName(COUNT_SHEET);
  if (sh) {
    const n = sh.getLastRow() - 1;
    const filled = n > 0 ? sh.getRange(2, 6, n, 1).getValues().filter(r => r[0] !== '' && r[0] != null).length : 0;
    if (filled) {
      const ui = SpreadsheetApp.getUi();
      const ans = ui.alert('แท็บนับยอดมีตัวเลขที่ยังไม่บันทึก ' + filled + ' ช่อง', 'สร้างใหม่จะลบตัวเลขเหล่านั้น ต้องการสร้างใหม่ไหม\n(ถ้าจะเก็บไว้ ให้กด ไม่ แล้วใช้เมนู "บันทึกยอดนับ" ก่อน)', ui.ButtonSet.YES_NO);
      if (ans !== ui.Button.YES) return;
    }
    sh.clear();
  } else sh = ss.insertSheet(COUNT_SHEET);
  sh.getRange(1, 1, 1, COUNT_HEAD.length).setValues([COUNT_HEAD]).setFontWeight('bold').setBackground('#e0f1ec');
  if (rows.length) sh.getRange(2, 1, rows.length, COUNT_HEAD.length).setValues(rows);
  sh.setFrozenRows(1);
  const n = Math.max(rows.length, 1);
  sh.getRange(2, 6, n, 1).setBackground('#fff7cc').setNumberFormat('0');
  sh.getRange(2, 7, n, 1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(2, 5, n, 1).setFontColor('#6b7f7a');
  sh.setColumnWidth(1, 70); sh.setColumnWidth(2, 340); sh.setColumnWidth(3, 70); sh.setColumnWidth(4, 150);
  sh.setColumnWidth(5, 90); sh.setColumnWidth(6, 90); sh.setColumnWidth(7, 150); sh.setColumnWidth(8, 260);
  ss.setActiveSheet(sh);
  try { SpreadsheetApp.getUi().alert('สร้างแท็บ "นับยอด" แล้ว ' + rows.length + ' รายการ\n\nกรอกจำนวนในช่อง "นับได้" (สีเหลือง) เฉพาะยาที่นับ แล้วกดเมนู คลังใน → บันทึกยอดนับ'); } catch (_) {}
}

/* แปลงค่าวันที่ในเซลล์ → yyyy-MM-dd (รับทั้งค่าวันที่ และข้อความ วว/ดด/ปปปป ค.ศ. หรือ พ.ศ.) */
function countIso_(v) {
  if (v === '' || v == null) return '';
  if (v instanceof Date) return isNaN(v) ? '' : Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(v).trim());
  if (!m) return null;
  let y = Number(m[3]); if (y > 2400) y -= 543;
  return y + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
}

function saveCountSheet(byName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(COUNT_SHEET);
  if (!sh || sh.getLastRow() < 2) throw new Error('ยังไม่มีแท็บ "นับยอด" ให้กดเมนู คลังใน → สร้างแท็บนับยอด ก่อน');
  const n = sh.getLastRow() - 1;
  const vals = sh.getRange(2, 1, n, COUNT_HEAD.length).getValues();
  const todo = [];
  const bad = [];
  vals.forEach((r, i) => {
    if (r[5] === '' || r[5] == null) return;
    const q = Number(String(r[5]).replace(/,/g, ''));
    const exp = countIso_(r[6]);
    if (!(q >= 0) || Math.round(q) !== q) bad.push([i, 'จำนวนต้องเป็นเลขจำนวนเต็ม 0 ขึ้นไป']);
    else if (exp === null) bad.push([i, 'วันหมดอายุไม่ถูกต้อง ใช้แบบ วว/ดด/ปปปป']);
    else todo.push({ i, c: String(r[0]).trim(), q, exp });
  });
  const ui = (() => { try { return SpreadsheetApp.getUi(); } catch (_) { return null; } })();
  if (!todo.length && !bad.length) { if (ui) ui.alert('ยังไม่มีตัวเลขในช่อง "นับได้"'); return { saved: 0 }; }

  let by = byName;
  if (by == null && ui) {
    const r = ui.prompt('บันทึกยอดนับ ' + todo.length + ' รายการ', 'ชื่อผู้นับ (จะขึ้นในประวัติ)', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return { saved: 0 };
    by = r.getResponseText().trim();
  }
  by = String(by || 'ผู้ดูแล (ชีต)');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('ระบบกำลังบันทึกของคนอื่น ลองใหม่อีกครั้ง');
  const out = vals.map(r => [r[4], r[5], r[6], r[7]]);
  let saved = 0;
  try {
    const db = loadDb();
    /* วันหมดอายุจากคลังยาใหญ่: อ่านครั้งเดียว ใช้ทุกตัว */
    let hubMap = null;
    db.hubExp = name => {
      if (!hubMap) {
        hubMap = {};
        try {
          hubLots_(hubSS_()).filter(l => l.exp).sort(hubByExp_).forEach(l => {
            const k = l.name; if (!hubMap[k] || (l.qty > 0 && !hubMap[k].pos)) hubMap[k] = { exp: l.exp, pos: l.qty > 0 };
          });
        } catch (_) {}
      }
      return hubMap[name] ? hubMap[name].exp : '';
    };
    const now = new Date();
    const stamp = Utilities.formatDate(now, TZ, 'dd/MM/yyyy HH:mm');
    todo.forEach(t => {
      try {
        const res = Core.count(db, { c: t.c, q: t.q, exp: t.exp, by, dev: 'sheet' }, now);
        const d = res.diff;
        out[t.i] = [res.bal, '', '', `บันทึกแล้ว ${stamp} · ${d === 0 ? 'ตรงกับระบบ' : (d > 0 ? 'เกิน ' : 'ขาด ') + Math.abs(d)}`];
        saved++;
      } catch (e) { out[t.i][3] = '⚠ ' + (e && e.message || e); }
    });
    bad.forEach(([i, msg]) => { out[i][3] = '⚠ ' + msg; });
    if (saved) saveDb(db);
  } finally { lock.releaseLock(); }
  sh.getRange(2, 5, n, 4).setValues(out);
  const msg = `บันทึกยอดนับแล้ว ${saved} รายการ` + (bad.length + todo.length - saved ? `\nมีปัญหา ${bad.length + todo.length - saved} รายการ ดูช่อง "ผล" ที่ขึ้น ⚠` : '');
  Logger.log(msg);
  if (ui) ui.alert(msg);
  return { saved, problems: bad.length + todo.length - saved };
}
