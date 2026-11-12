/**
 * รายการยาของคลังใน รพ.สต.บ้านบวกหมื้อ — ดึงจาก DrugMaster ของระบบคลังยาใหญ่ (ไม่มีรายการฝังในโค้ด)
 *
 * เพิ่มยาใหม่ภายหลัง ทำได้ 3 ทาง
 *   1) เพิ่มยาใน DrugMaster ของระบบใหญ่ตามปกติ แล้วกดเมนู คลังใน → ดึงยาใหม่จากคลังยาใหญ่
 *      (ระบบดึงให้เองทุกคืนด้วย ถ้ากดรัน installHubSync() ไว้หนึ่งครั้ง)
 *   2) เมนู คลังใน → เพิ่มยาใหม่ + สร้าง QR
 *   3) พิมพ์ชื่อยาในแถวว่างท้ายแท็บ "รายการยา" ระบบใส่รหัสให้เอง
 * ยาที่มีอยู่แล้วจะไม่ถูกแก้ชื่อ หมวด จุดสั่งซื้อ หรือติ๊กใช้งาน ระบบอัปเดตให้แค่ช่อง "ทุน/หน่วย"
 * รหัส B0001, B0002, … ไม่เปลี่ยนเมื่อดึงซ้ำ สติกเกอร์ QR ที่พิมพ์แล้วจึงใช้ได้ตลอด
 */
const ITEM_HEADER = ['รหัส', 'ชื่อยา', 'ชื่อไทย', 'หน่วยนับ', 'หมวด/ชั้น', 'จุดสั่งซื้อ', 'ใช้งาน', 'ชื่อในคลังบวกหมื้อ', 'ทุน/หน่วย'];
const CODE_PREFIX = 'B', CODE_DIGITS = 4;

/* เดาหมวดจากชื่อยา (แก้ได้ในชีต) */
function guessShelf_(name) {
  const n = ' ' + String(name || '').toLowerCase() + ' ';
  const has = re => re.test(n);
  if (has(/\b(inj|injection|vaccine|vial|amp)\b/)) return '1 ยาฉีด';
  if (has(/\b(tab|tablet|cap|capsule|chewable)\b/)) return '2 ยาเม็ด แคปซูล';
  if (has(/\b(syr|syrup|susp|suspension|elixir|mixt|mixture|drop|drops|oral sol|solution for oral|ors|powder|sachet)\b/)) return '3 ยาน้ำ ยาผง';
  if (has(/\b(cream|oint|ointment|gel|lotion|eye|ear|nasal|spray|supp|suppository|inhaler|mdi|neb|nebule|paint|solution|sol|scrub|balm)\b/)) return '4 ยาใช้ภายนอก';
  return '5 เวชภัณฑ์ วัสดุ';
}

/* ดึงยาที่ยังไม่มีในแท็บรายการยาจาก DrugMaster และอัปเดตทุน/หน่วยจาก StockLots */
function syncFromHub(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET.items);
  if (!sh) sh = ss.insertSheet(SHEET.items);
  sh.getRange(1, 1, 1, ITEM_HEADER.length).setValues([ITEM_HEADER]).setFontWeight('bold').setBackground('#e0f1ec');
  sh.setFrozenRows(1);
  const hub = hubSS_();
  const M = hubTable_(hub, HUB.master);
  const cN = M.col('name'), cU = M.has('unit') ? M.col('unit') : -1, cA = M.has('is_active') ? M.col('is_active') : -1;
  const cost = {};
  hubLots_(hub).forEach(l => { if (l.cost && (l.qty > 0 || !(l.name in cost))) cost[l.name] = l.cost; });
  const n = sh.getLastRow() - 1;
  const cur = n > 0 ? sh.getRange(2, 1, n, ITEM_HEADER.length).getValues() : [];
  const key = s => String(s || '').trim().toLowerCase();
  const have = {};
  let max = 0;
  cur.forEach(r => {
    if (r[7]) have[key(r[7])] = 1; if (r[1]) have[key(r[1])] = 1;
    const m = new RegExp('^' + CODE_PREFIX + '(\\d+)$').exec(String(r[0]).trim()); if (m) max = Math.max(max, Number(m[1]));
  });
  const add = [];
  M.vals.slice(1).forEach(r => {
    const name = String(r[cN] || '').trim();
    if (!name || have[key(name)]) return;
    have[key(name)] = 1;
    const on = cA < 0 ? true : !(r[cA] === false || String(r[cA]).toUpperCase() === 'FALSE');
    add.push([CODE_PREFIX + String(++max).padStart(CODE_DIGITS, '0'), name, '', cU < 0 ? '' : String(r[cU] || '').trim(), guessShelf_(name), '', on, name, cost[name] || '']);
  });
  if (cur.length) {
    const col = cur.map(r => { const c = cost[String(r[7] || r[1] || '').trim()]; return [c || r[8] || '']; });
    sh.getRange(2, 9, col.length, 1).setValues(col);
  }
  if (add.length) {
    const r0 = sh.getLastRow() + 1;
    sh.getRange(r0, 1, add.length, ITEM_HEADER.length).setValues(add);
    sh.getRange(r0, 7, add.length, 1).insertCheckboxes();
  }
  sh.getRange('I:I').setNumberFormat('#,##0.00');
  const msg = add.length
    ? `เพิ่มยาใหม่จากคลังยาใหญ่ ${add.length} รายการ (${add[0][0]} ถึง ${add[add.length - 1][0]})\n\nพิมพ์สติกเกอร์: เลือกแถวยาใหม่ในแท็บรายการยา แล้วกด คลังใน → พิมพ์สติกเกอร์ QR`
    : 'รายการยาตรงกับคลังยาใหญ่แล้ว ไม่มียาใหม่ (อัปเดตทุน/หน่วยให้แล้ว)';
  Logger.log(msg);
  if (silent !== true) { try { SpreadsheetApp.getUi().alert(msg); } catch (_) {} }
  return add.length;
}

/* ตัวตั้งเวลา: ดึงยาใหม่ทุกคืนตอนตี 1 (กดรันครั้งเดียว) */
function nightlyHubSync() { syncFromHub(true); }
function installHubSync() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'nightlyHubSync').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('nightlyHubSync').timeBased().everyDays(1).atHour(1).inTimezone(TZ).create();
  Logger.log('ตั้งเวลาดึงยาใหม่จากคลังยาใหญ่ทุกคืนแล้ว');
}
