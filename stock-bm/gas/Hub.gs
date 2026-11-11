/**
 * เชื่อมกับระบบคลังยาใหญ่ (CUP บวกหมื้อ) — ไฟล์ HUB_ID
 *   อ่าน:   DrugMaster (รายการยา) · StockLots (ล็อต วันหมดอายุ ทุน/หน่วย ยอดคงเหลือ)
 *   เขียน:  เฉพาะตอนกด "รับเข้า" ในแอปคลังใน และ HUB_DEDUCT ไม่ได้ตั้งเป็น OFF
 *           1) ลด qty_on_hand ของล็อตใน StockLots (ล็อตที่ยังไม่หมดอายุ หมดอายุเร็วสุดก่อน)
 *           2) เพิ่มบิลเบิกใน Dispenses + DispenseItems (ยกเลิกบิลได้จากหน้า "บิลย้อนหลัง" ของระบบใหญ่)
 *           3) เพิ่มแถวใน Movements
 *   ถ้าคลังยาใหญ่มีไม่พอ จะไม่บันทึกทั้งสองฝั่ง
 *   ลองดูก่อนโดยไม่เขียนอะไร: กดรัน previewHubCut()
 *
 * ค่าที่เขียนลงระบบใหญ่ ให้ตรงกับที่โค้ดระบบใหญ่ใช้ (ดูหมายเหตุในคู่มือ)
 */
const HUB = {
  lots: 'StockLots', master: 'DrugMaster', moves: 'Movements', disp: 'Dispenses', dispItems: 'DispenseItems',
  moveType: 'DISPENSE',     // ค่าในคอลัมน์ type ของ Movements สำหรับการเบิกออก
  status: 'ACTIVE',         // ค่าในคอลัมน์ status ของ Dispenses สำหรับบิลปกติ
};

function hubSS_() { return SpreadsheetApp.openById(settings_().HUB_ID); }
function hubDeduct_(st) { return String((st || settings_()).HUB_DEDUCT || 'ON').trim().toUpperCase() !== 'OFF'; }

/* วันที่จากชีตระบบใหญ่ → yyyy-MM-dd (รับทั้งค่าวันที่ และข้อความ วัน/เดือน/ปี ทั้ง ค.ศ. และ พ.ศ.) */
function hubIso_(v) {
  if (!v) return '';
  if (v instanceof Date) return isNaN(v) ? '' : Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = String(v).trim();
  let m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  if (m) return m[0];
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(s);
  if (!m) return '';
  let y = Number(m[3]); if (y > 2400) y -= 543;
  return y + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
}

/* อ่านแท็บของระบบใหญ่ พร้อมตัวหาคอลัมน์จากชื่อหัวตาราง */
function hubTable_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('ไม่พบแท็บ ' + name + ' ในระบบคลังยาใหญ่');
  const vals = sh.getDataRange().getValues();
  const head = (vals[0] || []).map(x => String(x).trim());
  const col = k => { const i = head.indexOf(k); if (i < 0) throw new Error('แท็บ ' + name + ' ของระบบคลังยาใหญ่ไม่มีคอลัมน์ ' + k); return i; };
  return { sh, vals, head, col, has: k => head.indexOf(k) >= 0 };
}

function hubLots_(ss) {
  const T = hubTable_(ss, HUB.lots);
  const c = { id: T.col('lot_id'), n: T.col('drug_name'), e: T.col('expiry_date'), q: T.col('qty_on_hand') };
  const opt = k => T.has(k) ? T.col(k) : -1;
  const cNo = opt('lot_no'), cU = opt('unit'), cC = opt('unit_cost');
  return T.vals.slice(1).map((r, i) => ({
    row: i + 2, id: String(r[c.id] || ''), name: String(r[c.n] || '').trim(), expRaw: r[c.e], exp: hubIso_(r[c.e]),
    lotNo: cNo < 0 ? '' : r[cNo], unit: cU < 0 ? '' : String(r[cU] || ''), cost: cC < 0 ? 0 : Number(r[cC]) || 0, qty: Number(r[c.q]) || 0,
  })).filter(l => l.id && l.name);
}

const hubByExp_ = (a, b) => (a.exp || '9999') < (b.exp || '9999') ? -1 : (a.exp || '9999') > (b.exp || '9999') ? 1 : 0;
function hubUsable_(lots, name, today) {
  return lots.filter(l => l.name === name && l.qty > 0 && (!l.exp || l.exp >= today)).sort(hubByExp_);
}

/* ข้อมูลให้หน้ารับเข้า: วันหมดอายุของล็อตที่จะถูกตัดก่อน และยอดที่คลังยาใหญ่มี */
function hubInfo_(name) {
  name = String(name || '').trim();
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const use = name ? hubUsable_(hubLots_(hubSS_()), name, today) : [];
  return { exp: use.length ? use[0].exp : '', avail: use.reduce((s, l) => s + l.qty, 0), unit: use.length ? use[0].unit : '', deduct: hubDeduct_() };
}

/* วันหมดอายุเร็วสุดที่ยังมีของ (ใช้ตอน "นับจริง" ยาที่ยังไม่เคยมีล็อต) */
function hubExpiry(name) {
  name = String(name || '').trim();
  if (!name) return '';
  const ls = hubLots_(hubSS_()).filter(l => l.name === name && l.exp);
  const pos = ls.filter(l => l.qty > 0).sort(hubByExp_);
  const any = ls.sort(hubByExp_);
  return pos.length ? pos[0].exp : any.length ? any[0].exp : '';
}

/* เลือกล็อตของคลังยาใหญ่ที่จะตัด: ยังไม่หมดอายุ หมดอายุเร็วสุดก่อน */
function hubPlan_(lots, name, q, today) {
  if (!lots.some(l => l.name === name)) throw new Error('ไม่พบ "' + name + '" ในคลังยาใหญ่ ตรวจช่อง "ชื่อในคลังบวกหมื้อ" ในแท็บรายการยา หรือถ้าของไม่ได้มาจากคลังยาใหญ่ให้ใช้ "นับจริง"');
  const use = hubUsable_(lots, name, today);
  const total = use.reduce((s, l) => s + l.qty, 0);
  if (total < q) throw new Error('คลังยาใหญ่มีของที่ยังไม่หมดอายุ ' + total + ' ' + (use.length ? use[0].unit : '') + ' รับเข้าได้ไม่เกินนี้');
  let rest = q; const plan = [];
  use.forEach(l => { if (!rest) return; const m = Math.min(l.qty, rest); rest -= m; plan.push({ lot: l, q: m, exp: l.exp }); });
  return plan;
}

/* รับเข้าคลังใน = เบิกออกจากคลังยาใหญ่ (ยอดสองฝั่งขยับพร้อมกัน) */
function receiveFromHub_(db, b, now) {
  if (!hubDeduct_(db.settings)) return Core.receive(db, b, now);
  const it = db.items.find(i => i.code === String(b.c));
  if (!it) throw new Error('ไม่พบรหัสยา ' + b.c);
  const q = Math.round(Number(b.q));
  if (!(q >= 1)) throw new Error('จำนวนต้องเป็น 1 ขึ้นไป');
  const ss = hubSS_();
  const today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
  const plan = hubPlan_(hubLots_(ss), it.hub || it.name, q, today);
  const ref = 'DP-' + Utilities.getUuid().replace(/-/g, '').slice(0, 12);
  let res = null;
  plan.forEach(a => { res = Core.receive(db, { c: it.code, q: a.q, exp: a.exp, by: b.by, dev: b.dev, note: 'จากคลังยาใหญ่ ' + ref }, now); });
  hubCommit_(ss, plan, { ref, by: b.by || '', facility: db.settings.FACILITY || DEFAULTS.FACILITY, now });
  res.exp = plan[0].exp;
  res.hub = { ref, lots: plan.map(a => [a.exp, a.q]) };
  return res;
}

/* เขียนลงระบบใหญ่: ตรวจทุกล็อตก่อน แล้วค่อยเขียน */
function hubCommit_(ss, plan, meta) {
  const L = hubTable_(ss, HUB.lots);
  const cId = L.col('lot_id') + 1, cQ = L.col('qty_on_hand') + 1;
  plan.forEach(a => {
    const id = String(L.sh.getRange(a.lot.row, cId).getValue());
    const cur = Number(L.sh.getRange(a.lot.row, cQ).getValue()) || 0;
    if (id !== a.lot.id || cur < a.q) throw new Error('ยอดในคลังยาใหญ่เพิ่งเปลี่ยน ลองกดรับเข้าอีกครั้ง');
    a.after = cur - a.q;
  });
  let user = ''; try { user = Session.getEffectiveUser().getEmail(); } catch (_) {}
  const uid = p => p + Utilities.getUuid().replace(/-/g, '').slice(0, 12);
  const append = (name, objs) => {
    const T = hubTable_(ss, name);
    const rows = objs.map(o => T.head.map(h => (h in o ? o[h] : '')));
    T.sh.getRange(T.sh.getLastRow() + 1, 1, rows.length, T.head.length).setValues(rows);
  };
  plan.forEach(a => L.sh.getRange(a.lot.row, cQ).setValue(a.after));
  append(HUB.disp, [{ dispense_id: meta.ref, dispense_date: meta.now, location: meta.facility,
    note: 'รับเข้าคลังใน (บันทึกจากแอปคลังใน' + (meta.by ? ' โดย ' + meta.by : '') + ')', status: HUB.status, created_by: user, created_at: meta.now }]);
  append(HUB.dispItems, plan.map(a => ({ item_id: uid('DI-'), dispense_id: meta.ref, drug_name: a.lot.name, lot_id: a.lot.id, lot_no: a.lot.lotNo,
    expiry_date: a.lot.expRaw, unit: a.lot.unit, unit_cost: a.lot.cost, qty: a.q, line_cost: Math.round(a.q * a.lot.cost * 100) / 100 })));
  append(HUB.moves, plan.map(a => ({ movement_id: uid('MV-'), timestamp: meta.now, lot_id: a.lot.id, drug_name: a.lot.name, type: HUB.moveType,
    qty_change: -a.q, ref: meta.ref, balance_after: a.after, user })));
}

/* ลองดูว่าถ้ารับเข้าจะตัดล็อตไหนของคลังยาใหญ่ (ไม่เขียนอะไร) — แก้ชื่อยาและจำนวนด้านล่างแล้วกดรัน */
function previewHubCut() {
  const name = 'Paracetamol tab 500 mg', q = 100;
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const plan = hubPlan_(hubLots_(hubSS_()), name, q, today);
  plan.forEach(a => Logger.log(`${a.lot.id} · หมดอายุ ${a.exp || '-'} · มี ${a.lot.qty} · จะตัด ${a.q} · เหลือ ${a.lot.qty - a.q}`));
  Logger.log('ตัดยอดระบบใหญ่ตอนรับเข้า: ' + (hubDeduct_() ? 'เปิด' : 'ปิด (HUB_DEDUCT = OFF)'));
}
