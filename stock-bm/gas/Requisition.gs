/**
 * ใบเบิกในชีต: เมนู "คลังใน" ด้านบนของชีต
 *   สร้างใบเบิกเดือนนี้          → สร้างแท็บ "ใบเบิกยา" และ "ใบเบิกวัสดุ" ใส่คงเหลือและจำนวนเบิกแนะนำให้
 *   ซ่อนรายการที่ไม่เบิก / แสดงทุกรายการ
 *   ดาวน์โหลด PDF                 → ได้ลิงก์ PDF ขนาด A4 ของทั้งสองแท็บ
 * ช่อง "เบิก" (สีเหลือง) แก้ได้ มูลค่ารวมคำนวณใหม่เอง
 * จำนวนแนะนำ = ยอดใช้เฉลี่ยต่อเดือน × จำนวนเดือน + จุดสั่งซื้อ − คงเหลือ
 *   (ตั้งจำนวนเดือนได้ที่พร็อพเพอร์ตี้ REQ_MONTHS ค่าเริ่มต้น 1, ยอดใช้คิดย้อนหลัง REQ_DAYS วัน ค่าเริ่มต้น 90)
 */
const REQ_SHEETS = { d: 'ใบเบิกยา', s: 'ใบเบิกวัสดุ' };
const REQ_SIGN = {
  d: { title: 'รายการขอเบิกยาจาก หน่วยบริการประจำศูนย์สาธารณสุขและการแพทย์บวกหมื้อ', approver: '( นายไพโรจน์  เลาหาง )', approverPos: 'ผู้อำนวยการโรงพยาบาลส่งเสริมสุขภาพตำบลบ้านบวกหมื้อ' },
  s: { title: 'รายการขอเบิกวัสดุการแพทย์ จากศูนย์สาธารณสุขและการแพทย์บวกหมื้อ', approver: '( นายไพโรจน์  เลาหาง )', approverPos: 'ผู้อำนวยการโรงพยาบาลส่งเสริมสุขภาพตำบลบ้านบวกหมื้อ' },
};
const TH_MONTH = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('คลังใน')
    .addItem('เปิดแดชบอร์ด', 'openDashboard')
    .addSeparator()
    .addItem('สร้างใบเบิกเดือนนี้', 'makeRequisition')
    .addSeparator()
    .addItem('ซ่อนรายการที่ไม่เบิก', 'reqHideZero')
    .addItem('แสดงทุกรายการ', 'reqShowAll')
    .addSeparator()
    .addItem('ดาวน์โหลด PDF ใบเบิก', 'reqPdf')
    .addSeparator()
    .addItem('สร้างแท็บนับยอด', 'makeCountSheet')
    .addItem('บันทึกยอดนับ', 'saveCountSheet')
    .addSeparator()
    .addItem('ดึงยาใหม่จากคลังยาใหญ่', 'syncFromHub')
    .addItem('เพิ่มยาใหม่ + สร้าง QR', 'addItemDialog')
    .addItem('พิมพ์สติกเกอร์ QR', 'labelsDialog')
    .addToUi();
}

function makeRequisition() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const st = settings_();
  const months = Number(st.REQ_MONTHS) || 1, days = Number(st.REQ_DAYS) || 90;
  const db = loadDb();
  const META = reqMeta_(db);
  const snap = Core.snapshot(db, new Date());
  const byCode = {};
  snap.items.forEach(i => { byCode[i.c] = i; });
  const u = getUsage(days);
  const now = new Date();
  const y = now.getFullYear() + 543, fy = now.getMonth() >= 9 ? y + 1 : y;
  const monthText = `${TH_MONTH[now.getMonth()]} ${String(y).slice(2)} ปีงบประมาณ ${fy}`;
  const fac = String(st.FACILITY || '').replace(/^รพ\.?\s*สต\.?\s*/, '');
  const suggest = (c, it) => {
    if (!it) return '';
    const min = it.min || 0, bal = Math.max(0, it.b);
    if (u.days < 14) return min && bal < min ? min - bal : '';
    const pm = (u.use[c] || 0) / u.days * 30;
    const need = Math.ceil(pm * months + min - bal);
    return need > 0 ? need : '';
  };
  ['d', 's'].forEach(kind => {
    const S = REQ_SIGN[kind], drug = kind === 'd';
    const W = drug ? 8 : 7;                       // จำนวนคอลัมน์ของตาราง
    let sh = ss.getSheetByName(REQ_SHEETS[kind]);
    if (sh) { sh.showRows(1, sh.getMaxRows()); sh.clear(); sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); }
    else sh = ss.insertSheet(REQ_SHEETS[kind]);
    sh.setHiddenGridlines(true);
    const head = drug ? ['ลำดับ', 'รายการ', 'ขนาดบรรจุ', 'หน่วยเบิก', 'ราคาต้นทุน', 'คงเหลือ', 'เบิก', 'หมายเหตุ'] : ['ลำดับ', 'รายการ', 'หน่วยเบิก', 'ราคาต้นทุน', 'คงเหลือ', 'เบิก', 'หมายเหตุ'];
    const qCol = drug ? 7 : 6, pCol = drug ? 5 : 4;
    // หัวกระดาษ
    const top = [[S.title], [`โรงพยาบาลส่งเสริมสุขภาพตำบล${fac || '…………………'} หมู่ที่……… ตำบล………………… อำเภอแม่แตง จังหวัดเชียงใหม่`], [`ประจำเดือน ${monthText}`]];
    top.forEach((t, i) => sh.getRange(i + 1, 1, 1, W).merge().setValue(t[0]).setHorizontalAlignment('center'));
    sh.getRange(1, 1).setFontWeight('bold').setFontSize(13);
    sh.getRange(4, 1, 1, W).setValues([head]).setFontWeight('bold').setBackground('#eef3f1').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    // รายการ
    const out = [], secRows = [], hasRows = [];
    let sec = '', n = 0;
    META.filter(m => m[1] === kind).forEach(m => {
      const [c, , s, name, pack, unit, price] = m;
      if (s !== sec) { sec = s; secRows.push(out.length); out.push([s].concat(Array(W - 1).fill(''))); }
      n++;
      const it = byCode[c], bal = it && it.has ? it.b : '', q = suggest(c, it);
      if (q) hasRows.push(out.length);
      out.push(drug ? [n, name, pack, unit, price == null ? '' : price, bal, q, ''] : [n, name, unit, price == null ? '' : price, bal, q, '']);
    });
    if (!out.length) { secRows.push(0); out.push(['ไม่มีรายการในหมวดนี้'].concat(Array(W - 1).fill(''))); }
    const r0 = 5, last = r0 + out.length - 1;
    sh.getRange(r0, 1, out.length, W).setValues(out).setVerticalAlignment('middle');
    sh.getRange(r0, 1, out.length, 1).setHorizontalAlignment('center');
    if (drug) sh.getRange(r0, 3, out.length, 2).setHorizontalAlignment('center'); else sh.getRange(r0, 3, out.length, 1).setHorizontalAlignment('center');
    sh.getRange(r0, pCol, out.length, 1).setNumberFormat('#,##0.00');
    sh.getRange(r0, pCol + 1, out.length, 2).setNumberFormat('#,##0').setHorizontalAlignment('right');
    sh.getRange(r0, qCol, out.length, 1).setBackground('#fffbea');
    secRows.forEach(i => sh.getRange(r0 + i, 1, 1, W).merge().setFontWeight('bold').setBackground('#f6f9f8').setHorizontalAlignment('left'));
    sh.getRange(4, 1, out.length + 1, W).setBorder(true, true, true, true, true, true, '#444444', SpreadsheetApp.BorderStyle.SOLID);
    // รวม (คำนวณใหม่เองเมื่อแก้ช่องเบิก)
    const qL = colLetter_(qCol), pL = colLetter_(pCol);
    const sumRow = last + 1;
    sh.getRange(sumRow, 1, 1, W).merge().setFormula(`="รวมรายการที่เบิก "&COUNTIF(${qL}${r0}:${qL}${last},">0")&" รายการ   มูลค่ารวม "&TEXT(SUMPRODUCT(IFERROR(${pL}${r0}:${pL}${last}*1,0),IFERROR(${qL}${r0}:${qL}${last}*1,0)),"#,##0.00")&" บาท"`);
    sh.getRange(sumRow, 1).setFontWeight('bold');
    // ลายเซ็น 2 × 2
    const half = Math.ceil(W / 2);
    const sig = [
      ['ลงชื่อ ........................................ ผู้เบิก\n( ........................................ )\nตำแหน่ง ........................................\nวันที่ ......../......../........',
       'ลงชื่อ ........................................ ผู้จ่าย\n(ภญ. กาญจนา ขอดแก้ว)\nตำแหน่ง เภสัชกรปฏิบัติการ\nวันที่ ......../......../........'],
      [`ลงชื่อ ........................................ ผู้อนุมัติจ่าย\n${S.approver}\nตำแหน่ง ${S.approverPos}\nวันที่ ......../......../........`,
       'ลงชื่อ ........................................ ผู้รับ\n( ........................................ )\nตำแหน่ง ........................................\nวันที่ ......../......../........'],
    ];
    sig.forEach((pair, i) => {
      const r = sumRow + 2 + i;
      sh.getRange(r, 1, 1, half).merge().setValue(pair[0]);
      sh.getRange(r, half + 1, 1, W - half).merge().setValue(pair[1]);
      sh.getRange(r, 1, 1, W).setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('top');
      sh.setRowHeight(r, 92);
    });
    // ขนาดคอลัมน์ให้พอดี A4 แนวตั้ง
    const widths = drug ? [42, 250, 62, 56, 72, 62, 56, 70] : [42, 300, 62, 72, 62, 56, 76];
    widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.getRange(1, 1, sumRow + 3, W).setFontFamily('Sarabun').setFontSize(11);
    sh.getRange(1, 1).setFontSize(13);
    sh.getRange(r0, 2, out.length, 1).setWrap(true);
    if (sh.getMaxColumns() > W) sh.deleteColumns(W + 1, sh.getMaxColumns() - W);
    if (sh.getMaxRows() > sumRow + 3) sh.deleteRows(sumRow + 4, sh.getMaxRows() - sumRow - 3);
    sh.setFrozenRows(4);
    sh.getRange(1, 1, 1, 1).setNote(`สร้างเมื่อ ${Utilities.formatDate(now, TZ, 'dd/MM/yyyy HH:mm')} · คงเหลือ ณ เวลานั้น · ยอดใช้คิดจาก ${u.days} วัน · เบิกให้พอ ${months} เดือน`);
  });
  ss.setActiveSheet(ss.getSheetByName(REQ_SHEETS.d));
  const msg = u.days < 14
    ? `สร้างใบเบิกแล้ว\n\nประวัติการหยิบมี ${u.days} วัน ยังน้อยเกินไปที่จะคำนวณยอดใช้ จำนวนเบิกแนะนำจึงคิดจากจุดสั่งซื้ออย่างเดียว ให้ตรวจและกรอกช่อง "เบิก" (สีเหลือง) เอง`
    : `สร้างใบเบิกแล้ว ตรวจและแก้ช่อง "เบิก" (สีเหลือง) ได้เลย\n\nเสร็จแล้วกด คลังใน → ดาวน์โหลด PDF ใบเบิก`;
  try { SpreadsheetApp.getUi().alert(msg); } catch (_) { Logger.log(msg); }
}

function colLetter_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

/* ซ่อนแถวที่ช่องเบิกว่าง (หัวหมวดยังอยู่) ก่อนดาวน์โหลด PDF ถ้าต้องการพิมพ์เฉพาะที่เบิก */
function reqHideZero() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['d', 's'].forEach(kind => {
    const sh = ss.getSheetByName(REQ_SHEETS[kind]);
    if (!sh) return;
    const qCol = kind === 'd' ? 7 : 6;
    const vals = sh.getRange(5, 1, sh.getLastRow() - 4, qCol).getValues();
    sh.showRows(1, sh.getMaxRows());
    let secAt = -1, secHas = false;
    const endSec = () => { if (secAt >= 0 && !secHas) sh.hideRows(5 + secAt); };
    vals.forEach((r, i) => {
      if (typeof r[0] !== 'number') { if (r[0] === '' || /^รวมรายการ|^ลงชื่อ/.test(String(r[0]))) return; endSec(); secAt = i; secHas = false; return; }   // หัวหมวด
      if (Number(r[qCol - 1]) > 0) secHas = true; else sh.hideRows(5 + i);
    });
    endSec();
  });
}
function reqShowAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.values(REQ_SHEETS).forEach(n => { const sh = ss.getSheetByName(n); if (sh) sh.showRows(1, sh.getMaxRows()); });
}

/* ลิงก์ดาวน์โหลด PDF A4 ของแต่ละแท็บ (เปิดด้วยบัญชีที่มีสิทธิ์ชีตนี้) */
function reqPdf() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=CENTER&fzr=false&top_margin=0.4&bottom_margin=0.5&left_margin=0.4&right_margin=0.4&horizontal_alignment=CENTER&gid=`;
  const links = Object.values(REQ_SHEETS).map(n => ss.getSheetByName(n)).filter(Boolean)
    .map(sh => `<p><a href="${base}${sh.getSheetId()}" target="_blank" style="display:inline-block;padding:10px 16px;border-radius:8px;background:#0f7a68;color:#fff;text-decoration:none;font-weight:bold">ดาวน์โหลด ${sh.getName()}.pdf</a></p>`);
  if (!links.length) { SpreadsheetApp.getUi().alert('ยังไม่มีใบเบิก ให้กด คลังใน → สร้างใบเบิกเดือนนี้ ก่อน'); return; }
  const html = HtmlService.createHtmlOutput(`<div style="font-family:sans-serif">${links.join('')}<p style="color:#666;font-size:13px">ถ้าซ่อนรายการที่ไม่เบิกไว้ PDF จะมีเฉพาะรายการที่เบิก</p></div>`).setWidth(340).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'ใบเบิก PDF');
}

/* รายการในใบเบิก สร้างจากแท็บรายการยา (ยาที่ติ๊กใช้งาน) เรียงตามหมวดแล้วตามชื่อ
   [รหัส, ประเภท (d ยา / s วัสดุ), หมวด, ชื่อ, ขนาดบรรจุ, หน่วยเบิก, ราคาต้นทุน] */
function reqMeta_(db) {
  return db.items.filter(i => i.active)
    .map(i => [i.code, /เวชภัณฑ์|วัสดุ|lab/i.test(i.shelf) ? 's' : 'd', String(i.shelf || 'ไม่ระบุหมวด').replace(/^\d+\s*/, ''), i.name, '', i.unit, i.cost || null, i.shelf || 'ฮ'])
    .sort((a, b) => a[7].localeCompare(b[7], 'th') || a[3].localeCompare(b[3], 'en'))
    .map(m => m.slice(0, 7));
}
