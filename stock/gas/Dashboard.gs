/**
 * แท็บ "แดชบอร์ด" ในชีต: ภาพรวมคลังใน
 * อัปเดตเองทุกครั้งที่เปิดชีต และทุก 1 ชั่วโมง (กดรัน installDashboard ครั้งเดียว)
 * หรือกดเมนู คลังใน → รีเฟรชแดชบอร์ด
 */
const DASH = 'แดชบอร์ด';

/* คำนวณข้อมูลทั้งหมดของแดชบอร์ด (ไม่แตะชีต ใช้ทดสอบได้) */
function dashData_(db, mv, now) {
  const s = Core.snapshot(db, now);
  const today = s.today, dayMs = 864e5;
  const days = iso => Math.round((new Date(iso) - new Date(today)) / dayMs);
  const price = {};
  (typeof REQ_META !== 'undefined' ? REQ_META : []).forEach(m => { if (m[6] != null) price[m[0]] = Number(m[6]) || 0; });
  const live = mv.filter(m => !/ยกเลิกแล้ว/.test(m.note));
  const takes = live.filter(m => m.type === 'หยิบ');
  const since = d => now.getTime() - d * dayMs;
  const use30 = {};
  takes.filter(m => new Date(m.time).getTime() >= since(30)).forEach(m => { use30[m.code] = (use30[m.code] || 0) + m.qty; });
  const exp = [], low = [];
  let stocked = 0, value = 0;
  s.items.forEach(it => {
    if (it.has && it.b > 0) { stocked++; value += it.b * (price[it.c] || 0); }
    it.lots.forEach(([e, q]) => { if (e && q > 0 && days(e) <= 180) exp.push({ name: it.t || it.n, en: it.n, exp: e, q, u: it.u, d: days(e) }); });
    if (it.has && (it.b <= 0 || (it.min > 0 && it.b <= it.min))) low.push({ name: it.t || it.n, en: it.n, b: it.b, u: it.u, min: it.min, pm: use30[it.c] || 0 });
  });
  exp.sort((a, b) => a.d - b.d);
  low.sort((a, b) => a.b - b.b || b.pm - a.pm);
  const isToday = t => Utilities.formatDate(new Date(t), TZ, 'yyyy-MM-dd') === today;
  const daily = [];
  for (let i = 13; i >= 0; i--) {
    const d = Utilities.formatDate(new Date(now.getTime() - i * dayMs), TZ, 'yyyy-MM-dd');
    daily.push([d.slice(8) + '/' + d.slice(5, 7), takes.filter(m => Utilities.formatDate(new Date(m.time), TZ, 'yyyy-MM-dd') === d).length]);
  }
  const nameOf = {}; s.items.forEach(it => { nameOf[it.c] = it.t || it.n; });
  const top = Object.keys(use30).map(c => [nameOf[c] || c, use30[c], (s.items.find(i => i.c === c) || {}).u || '']).sort((a, b) => b[1] - a[1]).slice(0, 10);
  const who = {};
  takes.filter(m => new Date(m.time).getTime() >= since(7)).forEach(m => { const k = m.by || '(ไม่ระบุชื่อ)'; who[k] = (who[k] || 0) + 1; });
  const people = Object.keys(who).map(k => [k, who[k]]).sort((a, b) => b[1] - a[1]).slice(0, 8);
  const recent = mv.slice(-10).reverse().map(m => [m.time, m.type, nameOf[m.code] || m.name, m.qty, m.by, m.note]);
  return {
    facility: s.facility, today,
    kpi: {
      stocked, expired: exp.filter(x => x.d < 0).length, exp90: exp.filter(x => x.d >= 0 && x.d <= 90).length, exp180: exp.filter(x => x.d > 90).length,
      low: low.length, takesToday: takes.filter(m => isToday(m.time)).length, value,
    },
    exp, low, daily, top, people, recent,
  };
}

function readMoves_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.moves);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const k = Math.min(n, 5000);
  return sh.getRange(sh.getLastRow() - k + 1, 1, k, 12).getValues().map(r => ({
    time: r[1] instanceof Date ? r[1].toISOString() : String(r[1]), type: String(r[2]), code: String(r[3]), name: String(r[4]),
    qty: Number(r[5]) || 0, by: String(r[9] || ''), note: String(r[11] || ''),
  }));
}

function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const D = dashData_(loadDb(), readMoves_(), now);
  let sh = ss.getSheetByName(DASH);
  if (!sh) { sh = ss.insertSheet(DASH, 0); }
  sh.clear(); sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.getCharts().forEach(c => sh.removeChart(c));
  sh.setHiddenGridlines(true);
  if (sh.getMaxColumns() < 18) sh.insertColumnsAfter(sh.getMaxColumns(), 18 - sh.getMaxColumns());
  [16, 150, 80, 70, 70, 16, 150, 70, 70, 70, 16, 120, 60].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  const ink = '#15302b', muted = '#5d7470', green = '#0f7a68';
  const fmtDate = iso => iso ? iso.slice(8, 10) + '/' + iso.slice(5, 7) + '/' + (Number(iso.slice(0, 4)) + 543 - 2500) : '';
  sh.getRange('B1').setValue(`แดชบอร์ดคลังใน ${D.facility}`).setFontSize(18).setFontWeight('bold').setFontColor(ink);
  sh.getRange('B2').setValue(`อัปเดต ${Utilities.formatDate(now, TZ, 'dd/MM/yyyy HH:mm')} น. · อัปเดตเองเมื่อเปิดชีตและทุก 1 ชั่วโมง`).setFontColor(muted).setFontSize(9);
  sh.setRowHeight(1, 34);
  // การ์ดตัวเลข 6 ใบ: ป้าย (แถว 4) ตัวเลข (แถว 5)
  const tiles = [
    ['ยาที่มีของในคลัง', D.kpi.stocked, 'รายการ', '#e0f1ec', green],
    ['หมดอายุแล้ว', D.kpi.expired, 'ล็อต', D.kpi.expired ? '#f9e3e1' : '#f3f6f5', D.kpi.expired ? '#b3302b' : muted],
    ['หมดอายุใน 3 เดือน', D.kpi.exp90, 'ล็อต', D.kpi.exp90 ? '#f9e3e1' : '#f3f6f5', D.kpi.exp90 ? '#b3302b' : muted],
    ['หมดอายุใน 4–6 เดือน', D.kpi.exp180, 'ล็อต', D.kpi.exp180 ? '#fbeedd' : '#f3f6f5', D.kpi.exp180 ? '#8a4a06' : muted],
    ['ใกล้หมด / หมด', D.kpi.low, 'รายการ', D.kpi.low ? '#fbeedd' : '#f3f6f5', D.kpi.low ? '#8a4a06' : muted],
    ['หยิบวันนี้', D.kpi.takesToday, 'ครั้ง', '#f3f6f5', ink],
  ];
  const tcol = [2, 3, 5, 7, 8, 10];   // แต่ละใบกว้าง 1–2 คอลัมน์
  const span = [1, 2, 1, 1, 2, 3];
  tiles.forEach((t, i) => {
    const lab = sh.getRange(4, tcol[i], 1, span[i]), num = sh.getRange(5, tcol[i], 1, span[i]);
    if (span[i] > 1) { lab.merge(); num.merge(); }
    lab.setValue(t[0]).setBackground(t[3]).setFontColor(muted).setFontSize(9).setHorizontalAlignment('left');
    num.setValue(t[1]).setBackground(t[3]).setFontColor(t[4]).setFontSize(22).setFontWeight('bold').setHorizontalAlignment('left').setNumberFormat(`0" ${t[2]}"`);
  });
  sh.setRowHeight(5, 40);
  sh.getRange('B6').setValue(`มูลค่ายาคงคลังโดยประมาณ ${D.kpi.value.toLocaleString('th-TH', { maximumFractionDigits: 0 })} บาท (ราคาตามใบเบิก)`).setFontColor(muted).setFontSize(9);
  // ตาราง helper
  const table = (r, c, title, head, rows, empty) => {
    sh.getRange(r, c).setValue(title).setFontWeight('bold').setFontSize(12).setFontColor(ink);
    sh.getRange(r + 1, c, 1, head.length).setValues([head]).setFontWeight('bold').setFontColor(muted).setFontSize(9).setBorder(false, false, true, false, false, false, '#c9d6d2', SpreadsheetApp.BorderStyle.SOLID);
    if (!rows.length) { sh.getRange(r + 2, c).setValue(empty).setFontColor(muted); return r + 3; }
    sh.getRange(r + 2, c, rows.length, head.length).setValues(rows).setFontColor(ink).setVerticalAlignment('middle')
      .setBorder(null, null, true, null, null, true, '#e3ebe9', SpreadsheetApp.BorderStyle.SOLID);
    return r + 2 + rows.length;
  };
  const R = 8;
  const expRows = D.exp.slice(0, 15).map(x => [x.name, fmtDate(x.exp), `${x.q} ${x.u}`, x.d < 0 ? 'หมดอายุแล้ว' : x.d <= 90 ? `อีก ${x.d} วัน` : `อีก ${Math.round(x.d / 30)} เดือน`]);
  const e1 = table(R, 2, `ยาใกล้หมดอายุ (${D.exp.length})`, ['ยา', 'วันหมดอายุ', 'เหลือ', 'สถานะ'], expRows, 'ไม่มียาหมดอายุภายใน 6 เดือน');
  D.exp.slice(0, 15).forEach((x, i) => sh.getRange(R + 2 + i, 5).setFontColor(x.d <= 90 ? '#b3302b' : '#8a4a06').setFontWeight('bold'));
  const lowRows = D.low.slice(0, 15).map(x => [x.name, x.b <= 0 ? 'หมด' : `${x.b} ${x.u}`, x.min || '-', x.pm ? `${x.pm} ${x.u}` : '-']);
  const e2 = table(R, 7, `ใกล้หมด / หมด (${D.low.length})`, ['ยา', 'คงเหลือ', 'จุดสั่งซื้อ', 'ใช้ 30 วัน'], lowRows, 'ยังไม่มียาที่ถึงจุดสั่งซื้อ');
  D.low.slice(0, 15).forEach((x, i) => sh.getRange(R + 2 + i, 8).setFontColor(x.b <= 0 ? '#b3302b' : '#8a4a06').setFontWeight('bold'));
  sh.getRange(R + 2, 2, 15, 1).setWrap(true); sh.getRange(R + 2, 7, 15, 1).setWrap(true);
  // ใครหยิบบ้าง 7 วัน (ขวาสุด)
  table(R, 12, 'ใครหยิบ 7 วัน', ['ชื่อ', 'ครั้ง'], D.people, 'ยังไม่มี');
  // แถวล่าง: หยิบมากสุด 30 วัน + กราฟหยิบรายวัน
  const R2 = Math.max(e1, e2) + 2;
  const topRows = D.top.map(x => [x[0], `${x[1]} ${x[2]}`, '']);
  const e3 = table(R2, 2, 'หยิบมากสุด 30 วัน', ['ยา', 'จำนวน', ''], topRows, 'ยังไม่มีการหยิบ');
  if (D.top.length) {
    const max = D.top[0][1];
    D.top.forEach((x, i) => sh.getRange(R2 + 2 + i, 4, 1, 2).merge().setFormula(`=SPARKLINE(${x[1]},{"charttype","bar";"max",${max};"color1","${green}"})`));
  }
  sh.getRange(R2, 7).setValue('จำนวนครั้งที่หยิบ 14 วันล่าสุด').setFontWeight('bold').setFontSize(12).setFontColor(ink);
  // ข้อมูลกราฟเก็บไว้คอลัมน์ซ่อน Q:R
  sh.getRange(1, 17, D.daily.length + 1, 2).setValues([['วัน', 'ครั้ง']].concat(D.daily));
  sh.hideColumns(17, 2);
  const chart = sh.newChart().asColumnChart().addRange(sh.getRange(1, 17, D.daily.length + 1, 2))
    .setPosition(R2 + 1, 7, 0, 4).setOption('legend', { position: 'none' }).setOption('colors', [green])
    .setOption('width', 470).setOption('height', 220).setOption('vAxis', { minValue: 0, format: '0' }).setOption('hAxis', { textStyle: { fontSize: 9 } })
    .setOption('chartArea', { left: 30, top: 10, width: '92%', height: '75%' }).build();
  sh.insertChart(chart);
  // รายการล่าสุด
  const R3 = Math.max(e3, R2 + 13) + 2;
  const recentRows = D.recent.map(x => [Utilities.formatDate(new Date(x[0]), TZ, 'dd/MM HH:mm'), x[1], x[2], x[3], x[4], x[5]]);
  table(R3, 2, 'รายการล่าสุด', ['เวลา', 'ประเภท', 'ยา', 'จำนวน', 'ผู้ทำ', 'หมายเหตุ'], recentRows, 'ยังไม่มีรายการ');
  sh.getRange(1, 1, sh.getMaxRows(), 13).setFontFamily('Sarabun');
  sh.setFrozenRows(0);
  ss.setActiveSheet(sh);
}

/* กดรันครั้งเดียว: อัปเดตแดชบอร์ดทุก 1 ชั่วโมง */
function installDashboard() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'refreshDashboard').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('refreshDashboard').timeBased().everyHours(1).create();
  refreshDashboard();
}
