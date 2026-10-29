/**
 * คลังยาใน รพ.สต. — สต๊อกการ์ดออนไลน์
 * ใช้คู่กับหน้าเว็บ https://focus2541.github.io/pharmacy-scanner/stock/
 *
 * แท็บในไฟล์นี้
 *   ตั้งค่า     ชื่อหน่วยบริการ, PIN ผู้ดูแลสต๊อก, รหัสไฟล์คลังบวกหมื้อ
 *   รายการยา   รหัส | ชื่อยา | ชื่อไทย | หน่วยนับ | ชั้น | จุดสั่งซื้อ | ใช้งาน
 *   ล็อต        แต่ละครั้งที่รับเข้า พร้อมวันหมดอายุและคงเหลือ (ตัดแบบหมดอายุก่อนออกก่อน)
 *   ประวัติ     ทุกการหยิบ / รับเข้า / นับ / ยกเลิก
 *   สรุปคงเหลือ สูตรรวมยอดต่อยา (เปิดดูในชีตได้)
 */

const SHEET = { settings: 'ตั้งค่า', items: 'รายการยา', lots: 'ล็อต', moves: 'ประวัติ', summary: 'สรุปคงเหลือ' };
const DEFAULTS = {
  FACILITY: 'รพ.สต.บ้านเป้า',
  PIN: '1234',
  HUB_ID: '1GTSOjQmz3LiAY5UgrxiDEFvU15IPWLkvF2NclxEFW0k',
};
const TZ = 'Asia/Bangkok';
const DENY = { ok: false, code: 'unapproved', error: 'เครื่องนี้ยังไม่ได้รับอนุญาตให้ใช้งาน' };

/* ===================== ติดตั้งครั้งแรก ===================== */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mk = (name, header) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(header);
      sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#e0f1ec');
      sh.setFrozenRows(1);
    }
    return sh;
  };
  const st = mk(SHEET.settings, ['key', 'value', 'คำอธิบาย']);
  if (st.getLastRow() === 1) {
    st.getRange(2, 1, 5, 3).setValues([
      ['FACILITY', DEFAULTS.FACILITY, 'ชื่อหน่วยบริการที่แสดงในหน้าเว็บ'],
      ['PIN', DEFAULTS.PIN, 'รหัสผู้ดูแลสต๊อก ใช้ตอนรับเข้าและนับ (เปลี่ยนได้)'],
      ['HUB_ID', DEFAULTS.HUB_ID, 'รหัสไฟล์ระบบคลังยาบวกหมื้อ ใช้ดึงวันหมดอายุ'],
      ['LINE_TOKEN', '', '(ไม่บังคับ) Channel access token ของ LINE OA สำหรับแจ้งเตือนรายวัน'],
      ['LINE_TO', '', '(ไม่บังคับ) รหัสกลุ่มไลน์ของหน่วยบริการที่จะรับแจ้งเตือน'],
    ]);
    st.getRange('B:B').setNumberFormat('@');
    st.getRange('B3').setValue(DEFAULTS.PIN);
  }
  const it = mk(SHEET.items, ITEM_HEADER);
  if (it.getLastRow() === 1) writeItems_(it);
  try { accessSettings_(); devSheet_(); } catch (_) {}
  mk(SHEET.lots, ['lot_id', 'รหัส', 'ชื่อยา', 'วันหมดอายุ', 'รับเข้า', 'คงเหลือ', 'วันที่รับ']);
  mk(SHEET.moves, ['move_id', 'เวลา', 'ประเภท', 'รหัส', 'ชื่อยา', 'จำนวน', 'ก่อน', 'หลัง', 'วันหมดอายุ', 'ผู้ทำ', 'เครื่อง', 'หมายเหตุ', 'ตัดจากล็อต']);
  const sm = mk(SHEET.summary, ['รหัส', 'ชื่อยา', 'คงเหลือ', 'หมดอายุเร็วสุด']);
  if (sm.getLastRow() === 1) {
    sm.getRange('A2').setFormula(
      "=IFERROR(QUERY({'ล็อต'!B2:B,'ล็อต'!C2:C,'ล็อต'!F2:F,'ล็อต'!D2:D},\"select Col1, Col2, sum(Col3), min(Col4) where Col1 is not null group by Col1, Col2 order by Col2 label sum(Col3) '', min(Col4) ''\",0),\"ยังไม่มีข้อมูล\")"
    );
    sm.getRange('D:D').setNumberFormat('dd/mm/yyyy');
  }
  ss.getSheetByName(SHEET.lots).getRange('D:D').setNumberFormat('dd/mm/yyyy');
  ss.getSheetByName(SHEET.lots).getRange('G:G').setNumberFormat('dd/mm/yyyy');
  const s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  SpreadsheetApp.getUi && Logger.log('ติดตั้งเสร็จ');
}

const ITEM_HEADER = ['รหัส', 'ชื่อยา', 'ชื่อไทย', 'หน่วยนับ', 'หมวด/ชั้น', 'จุดสั่งซื้อ', 'ใช้งาน', 'ชื่อในคลังบวกหมื้อ'];
function writeItems_(sh) {
  const rows = SEED_ITEMS.map(r => [r[0], r[1], r[2], r[3], r[4] || '', r[5] || '', r[6] !== false, r[7] || '']);
  sh.getRange(1, 1, 1, ITEM_HEADER.length).setValues([ITEM_HEADER]).setFontWeight('bold').setBackground('#e0f1ec');
  sh.getRange(2, 1, rows.length, ITEM_HEADER.length).setValues(rows);
  sh.getRange(2, 7, rows.length, 1).insertCheckboxes();
  sh.autoResizeColumns(1, 5);
}

/* แทนที่แท็บรายการยาด้วยรายการใน Items.gs (ใช้ตอนตั้งค่าครั้งแรกเท่านั้น) */
function resetItems() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lots = ss.getSheetByName(SHEET.lots);
  if (lots && lots.getLastRow() > 1) throw new Error('มีข้อมูลรับเข้าในแท็บล็อตแล้ว ถ้าต้องการเริ่มใหม่จริง ให้ลบแถวข้อมูลในแท็บล็อตและประวัติก่อน แล้วรันอีกครั้ง');
  let sh = ss.getSheetByName(SHEET.items);
  if (!sh) { setup(); return; }
  sh.clear();
  writeItems_(sh);
  Logger.log('ตั้งรายการยาใหม่ ' + SEED_ITEMS.length + ' รายการแล้ว');
}

/* ===================== API ===================== */

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.a === 'approve' || p.a === 'reject') return decide_(p);
    if (p.a === 'status') return json({ ok: true, status: devStatus_(p.tok) });
    if (p.a && !isApproved_(p.tok)) return json(DENY);
    if (p.a === 'all') return json(getAll());
    if (p.a === 'hubexp') return json({ ok: true, exp: hubExpiry(p.n || '') });
    if (p.a === 'usage') return json(getUsage(Number(p.days) || 90));
    return json({ ok: true, msg: 'stock api' });
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents || '{}'); } catch (_) { return json({ ok: false, error: 'ข้อมูลไม่ถูกต้อง' }); }
  if (b.a === 'register') {
    const lk = LockService.getScriptLock();
    if (!lk.tryLock(15000)) return json({ ok: false, error: 'ระบบไม่ว่าง ลองกดอีกครั้ง' });
    try { return json(register_(b)); } catch (err) { return json({ ok: false, error: String(err && err.message || err) }); } finally { lk.releaseLock(); }
  }
  if (!isApproved_(b.tok)) return json(DENY);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return json({ ok: false, error: 'ระบบกำลังบันทึกของคนอื่น ลองกดอีกครั้ง' });
  try {
    const db = loadDb();
    let res;
    if (b.a === 'take') res = Core.take(db, b, new Date());
    else if (b.a === 'receive') { checkPin(db, b.pin); res = Core.receive(db, b, new Date()); }
    else if (b.a === 'count') { checkPin(db, b.pin); res = Core.count(db, b, new Date()); }
    else if (b.a === 'undo') res = Core.undo(db, b, new Date());
    else if (b.a === 'dispose') { checkPin(db, b.pin); res = Core.dispose(db, b, new Date()); }
    else if (b.a === 'pin') { checkPin(db, b.pin); res = { ok: true }; }
    else return json({ ok: false, error: 'ไม่รู้จักคำสั่ง' });
    saveDb(db);
    res.state = Core.snapshot(db, new Date());
    return json(res);
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* PIN อ่านจาก "พร็อพเพอร์ตี้ของสคริปต์" ก่อน (คนเปิดชีตมองไม่เห็น) ถ้าไม่มีจึงใช้ค่าในแท็บตั้งค่า */
function checkPin(db, pin) {
  const real = String(PropertiesService.getScriptProperties().getProperty('PIN') || db.settings.PIN || DEFAULTS.PIN).trim();
  if (String(pin || '').trim() !== real) throw new Error('รหัสผู้ดูแลไม่ถูกต้อง');
}

function getUsage(days) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.moves);
  const n = sh.getLastRow() - 1;
  const rows = n > 0 ? sh.getRange(2, 2, n, 11).getValues().map(r => [r[0] instanceof Date ? r[0].toISOString() : r[0], String(r[1]), String(r[2]), r[4], String(r[10])]) : [];
  return Object.assign({ ok: true }, Core.usage(rows, days, new Date()));
}

function getAll() {
  const db = loadDb();
  return Object.assign({ ok: true }, Core.snapshot(db, new Date()));
}

/* ===================== อ่าน / เขียนชีต ===================== */

function loadDb() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = {};
  ss.getSheetByName(SHEET.settings).getDataRange().getValues().slice(1).forEach(r => { if (r[0]) settings[r[0]] = String(r[1]); });
  const items = ss.getSheetByName(SHEET.items).getDataRange().getValues().slice(1)
    .filter(r => r[0])
    .map(r => ({ code: String(r[0]), name: String(r[1]), th: String(r[2] || ''), unit: String(r[3] || ''), shelf: String(r[4] || ''), min: Number(r[5]) || 0, active: r[6] === true || String(r[6]).toUpperCase() === 'TRUE', hub: String(r[7] || '').trim() }));
  const lotSh = ss.getSheetByName(SHEET.lots);
  const lots = lotSh.getLastRow() > 1 ? lotSh.getRange(2, 1, lotSh.getLastRow() - 1, 7).getValues().map((r, i) => ({
    row: i + 2, id: String(r[0]), code: String(r[1]), name: String(r[2]), exp: toIso(r[3]), qin: Number(r[4]) || 0, left: Number(r[5]) || 0, date: toIso(r[6]),
  })) : [];
  const mvSh = ss.getSheetByName(SHEET.moves);
  const n = mvSh.getLastRow() - 1;
  const take = Math.min(n, 400);
  const moves = take > 0 ? mvSh.getRange(mvSh.getLastRow() - take + 1, 1, take, 13).getValues().map((r, i) => ({
    row: mvSh.getLastRow() - take + 1 + i, id: String(r[0]), time: r[1] instanceof Date ? r[1].toISOString() : String(r[1]), type: String(r[2]), code: String(r[3]), name: String(r[4]),
    qty: Number(r[5]) || 0, before: Number(r[6]) || 0, after: Number(r[7]) || 0, exp: toIso(r[8]), by: String(r[9]), dev: String(r[10]), note: String(r[11]), from: String(r[12] || ''),
  })) : [];
  return { settings, items, lots, moves, newLots: [], newMoves: [], dirtyLots: {}, dirtyMoves: {}, hubExp: hubExpiry };
}

function saveDb(db) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lotSh = ss.getSheetByName(SHEET.lots), mvSh = ss.getSheetByName(SHEET.moves);
  Object.keys(db.dirtyLots).forEach(id => {
    const l = db.lots.find(x => x.id === id);
    if (l && l.row) lotSh.getRange(l.row, 6).setValue(l.left);
  });
  if (db.newLots.length) {
    lotSh.getRange(lotSh.getLastRow() + 1, 1, db.newLots.length, 7).setValues(db.newLots.map(l => [l.id, l.code, l.name, l.exp ? isoToDate(l.exp) : '', l.qin, l.left, isoToDate(l.date)]));
  }
  Object.keys(db.dirtyMoves).forEach(id => {
    const m = db.moves.find(x => x.id === id);
    if (m && m.row) mvSh.getRange(m.row, 12).setValue(m.note);
  });
  if (db.newMoves.length) {
    mvSh.getRange(mvSh.getLastRow() + 1, 1, db.newMoves.length, 13).setValues(db.newMoves.map(m => [
      m.id, new Date(m.time), m.type, m.code, m.name, m.qty, m.before, m.after, m.exp ? isoToDate(m.exp) : '', m.by, m.dev, m.note, m.from,
    ]));
  }
}

function toIso(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = String(v).trim();
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0, 10) : '';
}
function isoToDate(s) { const p = s.split('-').map(Number); return new Date(p[0], p[1] - 1, p[2]); }

/* วันหมดอายุจากคลังบวกหมื้อ: ล็อตที่หมดอายุเร็วสุดที่ยังมีของ */
function hubExpiry(name) {
  name = String(name || '').trim();
  if (!name) return '';
  const cache = CacheService.getScriptCache();
  let map = null;
  const c = cache.get('hubexp');
  if (c) map = JSON.parse(c);
  if (!map) {
    map = {};
    const id = (loadSettingsOnly().HUB_ID) || DEFAULTS.HUB_ID;
    const sh = SpreadsheetApp.openById(id).getSheetByName('StockLots');
    const vals = sh.getDataRange().getValues();
    const h = vals[0].map(String);
    const iN = h.indexOf('drug_name'), iE = h.indexOf('expiry_date'), iQ = h.indexOf('qty_on_hand');
    vals.slice(1).forEach(r => {
      const n = String(r[iN] || '').trim(), e = toIso(r[iE]), q = Number(r[iQ]) || 0;
      if (!n || !e) return;
      const cur = map[n];
      const better = !cur || (q > 0 && cur.q <= 0) || ((q > 0) === (cur.q > 0) && e < cur.e);
      if (better) map[n] = { e, q };
    });
    try { cache.put('hubexp', JSON.stringify(map), 21600); } catch (_) {}
  }
  return map[name] ? map[name].e : '';
}
function loadSettingsOnly() {
  const out = {};
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.settings).getDataRange().getValues().slice(1).forEach(r => { if (r[0]) out[r[0]] = String(r[1]); });
  return out;
}

/* ล้างแคชวันหมดอายุ (กดรันเมื่อบวกหมื้อเพิ่มล็อตใหม่แล้วอยากเห็นทันที) */
function refreshHubExpiry() { CacheService.getScriptCache().remove('hubexp'); Logger.log(hubExpiry('Paracetamol tab 500 mg')); }

/* ===================== แจ้งเตือนทางไลน์ (ไม่บังคับ) =====================
 * ใส่ LINE_TOKEN (Channel access token ของ LINE OA) และ LINE_TO (รหัสกลุ่มไลน์ของบ้านเป้า)
 * ในแท็บตั้งค่า แล้วกดรัน installDailyAlert() หนึ่งครั้ง ระบบจะส่งทุกวันตอน 8 โมงเช้า
 */
function dailyAlert() {
  const db = loadDb();
  const token = db.settings.LINE_TOKEN, to = db.settings.LINE_TO;
  if (!token || !to) { Logger.log('ยังไม่ได้ตั้ง LINE_TOKEN / LINE_TO'); return; }
  const s = Core.snapshot(db, new Date());
  const day = iso => Math.round((new Date(iso) - new Date(s.today)) / 864e5);
  const exp = [], low = [];
  s.items.forEach(it => {
    const name = it.t || it.n, u = it.u || 'ชิ้น';
    it.lots.forEach(([e, q]) => {
      if (!e || q <= 0) return;
      const d = day(e);
      if (d <= 180) exp.push([d, `${d < 0 ? '⚫ หมดอายุแล้ว' : d <= 90 ? '🔴 ' + d + ' วัน' : '🟡 ' + Math.round(d / 30) + ' เดือน'} · ${name} ${q} ${u} (${e.split('-').reverse().join('/')})`]);
    });
    if (it.has && (it.b <= 0 || (it.min > 0 && it.b <= it.min))) low.push(`${it.b <= 0 ? '⚪ หมด' : '🟠 เหลือ ' + it.b + ' ' + u} · ${name}`);
  });
  if (!exp.length && !low.length) return;
  exp.sort((a, b) => a[0] - b[0]);
  let msg = `📦 คลังใน ${s.facility}\nประจำวันที่ ${s.today.split('-').reverse().join('/')}`;
  if (exp.length) msg += `\n\nยาใกล้หมดอายุ (${exp.length})\n` + exp.slice(0, 25).map(x => x[1]).join('\n') + (exp.length > 25 ? `\n…และอีก ${exp.length - 25} รายการ` : '');
  if (low.length) msg += `\n\nยาใกล้หมด / หมด (${low.length})\n` + low.slice(0, 20).join('\n') + (low.length > 20 ? `\n…และอีก ${low.length - 20} รายการ` : '');
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ to, messages: [{ type: 'text', text: msg.slice(0, 4900) }] }), muteHttpExceptions: true,
  });
}
function installDailyAlert() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailyAlert').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyAlert').timeBased().everyDays(1).atHour(8).inTimezone(TZ).create();
  dailyAlert();
}
