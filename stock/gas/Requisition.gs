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
  try { refreshDashboard(); } catch (_) {}      // แดชบอร์ดอัปเดตทุกครั้งที่เปิดชีต
  SpreadsheetApp.getUi().createMenu('คลังใน')
    .addItem('รีเฟรชแดชบอร์ด', 'refreshDashboard')
    .addSeparator()
    .addItem('สร้างใบเบิกเดือนนี้', 'makeRequisition')
    .addSeparator()
    .addItem('ซ่อนรายการที่ไม่เบิก', 'reqHideZero')
    .addItem('แสดงทุกรายการ', 'reqShowAll')
    .addSeparator()
    .addItem('ดาวน์โหลด PDF ใบเบิก', 'reqPdf')
    .addSeparator()
    .addItem('เพิ่มยาใหม่ + สร้าง QR', 'addItemDialog')
    .addItem('พิมพ์สติกเกอร์ QR', 'labelsDialog')
    .addToUi();
}

function makeRequisition() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const st = settings_();
  const months = Number(st.REQ_MONTHS) || 1, days = Number(st.REQ_DAYS) || 90;
  const db = loadDb();
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
    const known = new Set(REQ_META.map(m => m[0]));
    const extra = db.items.filter(i => i.active && !known.has(i.code))           // ยาที่เพิ่มในชีตทีหลัง
      .map(i => [i.code, /เวชภัณฑ์|lab|วัสดุ/i.test(i.shelf) ? 's' : 'd', 'รายการเพิ่มใหม่', i.name, '', i.unit, null]);
    REQ_META.concat(extra).filter(m => m[1] === kind).forEach(m => {
      const [c, , s, name, pack, unit, price] = m;
      if (s !== sec) { sec = s; secRows.push(out.length); out.push([s].concat(Array(W - 1).fill(''))); }
      n++;
      const it = byCode[c], bal = it && it.has ? it.b : '', q = suggest(c, it);
      if (q) hasRows.push(out.length);
      out.push(drug ? [n, name, pack, unit, price == null ? '' : price, bal, q, ''] : [n, name, unit, price == null ? '' : price, bal, q, '']);
    });
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

/* รหัส · ประเภท (d ยา / s วัสดุ) · หมวด · ชื่อในใบเบิก · ขนาดบรรจุ · หน่วยเบิก · ราคาต้นทุน */
const REQ_META = [
  ["P001", "d", "ยาฉีด", "Adrenaline 1 mg/ml inj", "1", "Amp", 6.21],
  ["P002", "d", "ยาฉีด", "Chlorpheniramine 10 mg/ml inj", "1", "Amp", 5.35],
  ["P003", "d", "ยาฉีด", "Dexamethasone 4 mg/ml inj", "1", "Amp", 5.3],
  ["P004", "d", "ยาฉีด", "Dextrose 50% inj 50 ml", "1", "Vial", 15],
  ["P005", "d", "ยาฉีด", "Dimenhydrinate 50 mg/ml inj", "1", "Amp", 4],
  ["P006", "d", "ยาฉีด", "dT 0.5 ml inj", "1", "Amp", 19.2],
  ["P007", "d", "ยาฉีด", "Hyoscine 20 mg/ml inj", "1", "Amp", 10.25],
  ["P008", "d", "ยาฉีด", "Insulin 70/30 10 ml inj", "1", "Vial", 125],
  ["P009", "d", "ยาฉีด", "Insulin NPH 100 IU/ml 10 ml inj", "1", "Vial", 75],
  ["P010", "d", "ยาฉีด", "Lidocaine 2% 20 ml inj", "1", "Vial", 41.2],
  ["P011", "d", "ยาฉีด", "Medroxyprogesterone 150 mg/3 ml inj", "1", "Vial", 10.7],
  ["P012", "d", "ยาฉีด", "Metoclopramide 10 mg/2 ml inj", "1", "Amp", 6.42],
  ["P013", "d", "ยาฉีด", "NSS 0.9% 5 ml inj", "50 nebule", "box", 100],
  ["P014", "d", "ยาฉีด", "NSS 0.9% 100 ml inj", "1", "Bag", 13],
  ["P015", "d", "ยาฉีด", "NSS 0.9% 1000 ml inj", "1", "Bag", 28],
  ["P016", "d", "ยาฉีด", "Omeprazole 40 mg inj", "1", "Vial", 57],
  ["P017", "d", "ยาฉีด", "Rabies vaccine", "1", "Amp", 240.97],
  ["P018", "d", "ยาฉีด", "Sterile water for INJ 10 ml", "50 ampule", "box", 99],
  ["P019", "d", "ยาใช้ภายนอก", "Aromatic ammonia spirit 450 ml", "1", "Bottle", 152.47],
  ["P020", "d", "ยาใช้ภายนอก", "Benzyl benzoate 25% 60 ml", "1", "Bottle", 18.2],
  ["P021", "d", "ยาใช้ภายนอก", "Betamethasone cream 5 g", "1", "Bottle", 8.56],
  ["P022", "d", "ยาใช้ภายนอก", "Berodual Nebule 0.5 mg + 1.25 mg/4 ml", "1", "Nebule", 7.16],
  ["P023", "d", "ยาใช้ภายนอก", "Berodual MDI 200 doses", "1", "Tube", 100.58],
  ["P024", "d", "ยาใช้ภายนอก", "Budesonide MDI 200 mcg", "1", "Tube", 57.6],
  ["P025", "d", "ยาใช้ภายนอก", "Calamine lotion 60 ml", "1", "Bottle", 14.98],
  ["P026", "d", "ยาใช้ภายนอก", "Chlorhexidine scrub 450 ml", "1", "Bottle", 90.95],
  ["P027", "d", "ยาใช้ภายนอก", "Clotrimazole cream 1% 15 g", "1", "Tube", 13],
  ["P028", "d", "ยาใช้ภายนอก", "Clotrimazole VT 500 mg", "1", "Tablet", 10],
  ["P029", "d", "ยาใช้ภายนอก", "Dexa-oph eye drop 4 ml", "1", "Bottle", 31],
  ["P030", "d", "ยาใช้ภายนอก", "Ethanol 70% 450 ml", "1", "Bottle", 27.82],
  ["P031", "d", "ยาใช้ภายนอก", "Hydrogen peroxide 450 ml", "1", "Bottle", 25.78],
  ["P032", "d", "ยาใช้ภายนอก", "Hista-oph eye drop 10 ml", "1", "Bottle", 15],
  ["P033", "d", "ยาใช้ภายนอก", "Methyl salicylate cream 15 g", "1", "Tube", 13.91],
  ["P034", "d", "ยาใช้ภายนอก", "NSS irrigation 1000 ml", "1", "Bottle", 27],
  ["P035", "d", "ยาใช้ภายนอก", "Poly-oph eye drop 5 ml", "1", "Bottle", 33],
  ["P036", "d", "ยาใช้ภายนอก", "Povidine 10% 450 ml", "1", "Bottle", 107],
  ["P037", "d", "ยาใช้ภายนอก", "Povidine scrub 450 ml", "1", "Bottle", 97],
  ["P038", "d", "ยาใช้ภายนอก", "Salbutamol Nebule 2.5 mg/2.5 ml", "1", "Nebule", 7.5],
  ["P039", "d", "ยาใช้ภายนอก", "Salbutamol MDI 0.1 mg/dose inhaler", "1", "Tube", 30],
  ["P040", "d", "ยาใช้ภายนอก", "Salmeterol + Fluticasone 25/125 evohaler", "1", "Tube", 116.63],
  ["P041", "d", "ยาใช้ภายนอก", "Silver sulfadiazine cream 25 g", "1", "Bottle", 25],
  ["P042", "d", "ยาใช้ภายนอก", "T.A. cream 0.02% 5 g", "1", "Tube", 5.7],
  ["P043", "d", "ยาใช้ภายนอก", "T.A. cream 0.1% 5 g", "1", "Tube", 10],
  ["P044", "d", "ยาใช้ภายนอก", "T.A. oral paste 1 g", "50 Sachet", "box", 148],
  ["P045", "d", "ยาใช้ภายนอก", "Water for irrigation 1000 ml", "1", "Bottle", 29],
  ["P046", "d", "ยาน้ำ", "Albendazole 400 mg/20 ml Syr", "1", "Bottle", 14.45],
  ["P047", "d", "ยาน้ำ", "Amoxicillin 125 mg/5 ml Syr", "1", "Bottle", 12.8],
  ["P048", "d", "ยาน้ำ", "Amoxicillin/Clavulanic acid 457 mg/5 ml Syr", "1", "Bottle", 93.09],
  ["P049", "d", "ยาน้ำ", "Antacid 240 ml Syr", "1", "Bottle", 19],
  ["P050", "d", "ยาน้ำ", "Brown mixture 60 ml", "1", "Bottle", 12.84],
  ["P051", "d", "ยาน้ำ", "Brompheniramine + Phenylephrine 2 mg +5 mg/5 ml Syr", "1", "Bottle", 12.5],
  ["P052", "d", "ยาน้ำ", "Carminative mixture 180 ml", "1", "Bottle", 16.05],
  ["P053", "d", "ยาน้ำ", "Cetirizine 1 mg/ml Syr", "1", "Bottle", 12.1],
  ["P054", "d", "ยาน้ำ", "Chlorpheniramine 2 mg/5ml Syr", "1", "Bottle", 8.56],
  ["P055", "d", "ยาน้ำ", "Dicloxacillin 62.5 mg/ml Syr", "1", "Bottle", 16.5],
  ["P056", "d", "ยาน้ำ", "Domperidone 5 mg/5 ml Syr", "1", "Bottle", 6],
  ["P057", "d", "ยาน้ำ", "Erythromycin 125 mg/5ml Syr", "1", "Bottle", 16.98],
  ["P058", "d", "ยาน้ำ", "Ferrous fumarate 15 mg/0.6 ml drop", "1", "Bottle", 25],
  ["P059", "d", "ยาน้ำ", "Glyceryl guaiacolate 100 mg/5ml Syr", "1", "Bottle", 13.91],
  ["P060", "d", "ยาน้ำ", "Hyoscine 5 mg/5ml Syr", "1", "Bottle", 13.5],
  ["P061", "d", "ยาน้ำ", "Ibuprofen 100 mg/5 ml Syr", "1", "Bottle", 14.98],
  ["P062", "d", "ยาน้ำ", "Milk of magnesia (MOM) 240 ml", "1", "Bottle", 29.96],
  ["P063", "d", "ยาน้ำ", "Multivitamins Syr", "1", "Bottle", 17.78],
  ["P064", "d", "ยาน้ำ", "Paracetamol 120 ml/5 mg Syr", "1", "Bottle", 8],
  ["P065", "d", "ยาน้ำ", "Paracetamol 60 mg/0.6 ml Drop", "1", "Bottle", 13],
  ["P066", "d", "ยาน้ำ", "Salbutamol 2 mg/5 ml Syr", "1", "Bottle", 11.77],
  ["P067", "d", "ยาน้ำ", "Sulfamethoxazole + Trimethoprim Syrup", "1", "Bottle", 10],
  ["P068", "d", "ยาเม็ด", "Acyclovir 400 mg", "70's", "box", 211],
  ["P069", "d", "ยาเม็ด", "Albendozole 200 mg", "100's", "box", 165],
  ["P070", "d", "ยาเม็ด", "Amitriptyline 10 mg", "500's", "box", 166.92],
  ["P071", "d", "ยาเม็ด", "Amoxicillin 500 mg", "500's", "box", 855],
  ["P072", "d", "ยาเม็ด", "Amoxicillin/Clavulanic acid 875/125 mg", "100's", "box", 293],
  ["P073", "d", "ยาเม็ด", "Betahistine mesilate 6 mg", "500's", "box", 130],
  ["P074", "d", "ยาเม็ด", "Brompheniramine 4 mg + Phenylephrine 7.5 mg", "1000's", "Bottle", 220],
  ["P075", "d", "ยาเม็ด", "Calcium carbonate 1500 mg", "100's", "box", 40.66],
  ["P076", "d", "ยาเม็ด", "Cetirizine 10 mg", "1000's", "box", 135.89],
  ["P077", "d", "ยาเม็ด", "Chlorpheniramine 4 mg", "1000's", "box", 90.95],
  ["P078", "d", "ยาเม็ด", "Clindamycin 300 mg", "100's", "box", 260],
  ["P079", "d", "ยาเม็ด", "Dextromethorphan 15 mg", "1000's", "box", 316],
  ["P080", "d", "ยาเม็ด", "Diazepam 5 mg", "1000's", "Bottle", 246.1],
  ["P081", "d", "ยาเม็ด", "Diclofenac 25 mg", "1000's", "box", 440],
  ["P082", "d", "ยาเม็ด", "Dicloxacillin 250 mg", "500's", "box", 555],
  ["P083", "d", "ยาเม็ด", "Dimenhydrinate 50 mg", "1000's", "box", 180],
  ["P084", "d", "ยาเม็ด", "Domperidone 10 mg", "500's", "box", 95],
  ["P085", "d", "ยาเม็ด", "Ethinylestradiol 30 mcg + Levonorgestrel 0.15 mg", "50 blister", "box", 270],
  ["P086", "d", "ยาเม็ด", "Fluconazole 200 mg", "50's", "box", 190],
  ["P087", "d", "ยาเม็ด", "Glyceryl guaiacolate 100 mg", "1000's", "Bottle", 150],
  ["P088", "d", "ยาเม็ด", "Hydroxyzine 10 mg", "1000's", "box", 120],
  ["P089", "d", "ยาเม็ด", "Hyoscine 10 mg", "500's", "box", 405],
  ["P090", "d", "ยาเม็ด", "Ibuprofen 400 mg", "1000's", "box", 350],
  ["P091", "d", "ยาเม็ด", "Lorazepam 1 mg", "1000's", "box", 270],
  ["P092", "d", "ยาเม็ด", "Metronidazole 400 mg", "500's", "box", 300],
  ["P093", "d", "ยาเม็ด", "Multivitamins tab", "1000's", "Bottle", 256.8],
  ["P094", "d", "ยาเม็ด", "Norfloxacin 400 mg", "500's", "box", 400],
  ["P095", "d", "ยาเม็ด", "Omeprazole 20 mg", "100's", "box", 55],
  ["P096", "d", "ยาเม็ด", "Orphenadrine 35 + paracetamol 500 mg", "1000's", "Bottle", 147.66],
  ["P097", "d", "ยาเม็ด", "ORS เด็ก", "100 Sachet", "box", 80],
  ["P098", "d", "ยาเม็ด", "ORS ผู้ใหญ่", "100 Sachet", "box", 128],
  ["P099", "d", "ยาเม็ด", "Paracetamol 325 mg", "1000's", "Bottle", 240],
  ["P100", "d", "ยาเม็ด", "Paracetamol 500 mg", "500's", "box", 225],
  ["P101", "d", "ยาเม็ด", "Prednisolone 5 mg", "1000's", "Bottle", 175.58],
  ["P102", "d", "ยาเม็ด", "Roxithromycin 150 mg", "100's", "box", 90],
  ["P103", "d", "ยาเม็ด", "Sennoside 7.5 mg", "60's", "box", 53.43],
  ["P104", "d", "ยาเม็ด", "Simethicone 80 mg", "100's", "box", 34],
  ["P105", "d", "ยาเม็ด", "Tramadol 50 mg", "500's", "box", 160],
  ["P106", "d", "ยาเม็ด", "Triferdine", "30's", "Bottle", 26.75],
  ["P107", "d", "ยาเม็ด", "Vitamin B1 100 mg", "1000's", "Bottle", 310],
  ["P108", "d", "ยาเม็ด", "Vitamin B complex", "1000's", "Bottle", 149.8],
  ["P109", "d", "ยาเม็ด", "Vitamin C 100 mg", "1000's", "Bottle", 134.75],
  ["P110", "d", "สมุนไพร", "ขมิ้นชัน 500 mg", "100's", "box", 78.11],
  ["P111", "d", "สมุนไพร", "เถาวัลย์เปรียง 500 mg", "100's", "box", 80],
  ["P112", "d", "สมุนไพร", "ฟ้าทะลายโจร 500 mg", "100's", "box", 58],
  ["P113", "d", "สมุนไพร", "สหัสธาราแคปซูล 500 mg", "100's", "box", 80],
  ["P114", "d", "สมุนไพร", "ยาแก้ไอมะขามป้อม 120 ml", "1", "Bottle", 18],
  ["P115", "d", "สมุนไพร", "ยาธาตุอบเชย 120 ml", "1", "Bottle", 18],
  ["P116", "d", "สมุนไพร", "ยาอมมะแว้ง 1 ซอง 20's", "20 Sachet", "box", 88.8],
  ["P117", "d", "สมุนไพร", "ยาไพล 30 g", "1", "Tube", 34],
  ["P118", "d", "สมุนไพร", "ยารางจืด 400 mg", "100's", "box", 80],
  ["P119", "d", "สมุนไพร", "เพชรสังฆาต 500 mg", "100's", "box", 60],
  ["P120", "d", "สมุนไพร", "หญ้าหนวดแมว 2 g", "20 ซองชา", "Sachet", 50],
  ["P121", "d", "สมุนไพร", "ยาหอมเทพจิตร 200 mg 1 ซอง 20's", "20 Sachet", "box", 178],
  ["P122", "d", "สมุนไพร", "ยามะขามแขก 400 mg", "100's", "box", 60],
  ["P123", "d", "สมุนไพร", "ครีมพญายอ 5 g", "1", "Tube", 20],
  ["P124", "d", "สมุนไพร", "ลูกประคบ 200 g", "1", "Unit", 42],
  ["P125", "d", "ยาโรคเรื้อรัง", "Aspirin EC 81 mg", "1000's", "box", 160],
  ["P126", "d", "ยาโรคเรื้อรัง", "Allopurinol 100 mg", "500's", "box", 235.4],
  ["P127", "d", "ยาโรคเรื้อรัง", "Amiloride 5 mg + HCTZ 50 mg", "250's", "box", 112.5],
  ["P128", "d", "ยาโรคเรื้อรัง", "Amlodipine 5 mg", "100's", "box", 38.52],
  ["P129", "d", "ยาโรคเรื้อรัง", "Atenolol 50 mg", "500's", "box", 116],
  ["P130", "d", "ยาโรคเรื้อรัง", "Atorvastatin 40 mg", "30's", "box", 34.2],
  ["P131", "d", "ยาโรคเรื้อรัง", "Colchicine 0.6 mg", "500's", "box", 180],
  ["P132", "d", "ยาโรคเรื้อรัง", "Digoxin 0.25 mg", "500's", "box", 400.65],
  ["P133", "d", "ยาโรคเรื้อรัง", "Doxazosin 2 mg", "500's", "box", 93],
  ["P134", "d", "ยาโรคเรื้อรัง", "Enalapril 5 mg", "1000's", "box", 143],
  ["P135", "d", "ยาโรคเรื้อรัง", "Enalapril 20 mg", "1000's", "box", 247],
  ["P136", "d", "ยาโรคเรื้อรัง", "Fenofibrate 100 mg", "100's", "box", 200],
  ["P137", "d", "ยาโรคเรื้อรัง", "Ferrous fumarate 200 mg", "1000's", "Bottle", 120],
  ["P138", "d", "ยาโรคเรื้อรัง", "Folic 5 mg", "1000's", "Bottle", 224.7],
  ["P139", "d", "ยาโรคเรื้อรัง", "Furosemide 40 mg", "500's", "Bottle", 139.1],
  ["P140", "d", "ยาโรคเรื้อรัง", "Gemfibrozil 600 mg", "500's", "box", 558],
  ["P141", "d", "ยาโรคเรื้อรัง", "Glipizide 5 mg", "500's", "box", 69],
  ["P142", "d", "ยาโรคเรื้อรัง", "Hydralazine 25 mg", "500's", "box", 169],
  ["P143", "d", "ยาโรคเรื้อรัง", "Hydrochlorothiazide 50 mg", "500's", "box", 190],
  ["P144", "d", "ยาโรคเรื้อรัง", "Isosorbide dinitrate SL 5 mg", "100's", "box", 75],
  ["P145", "d", "ยาโรคเรื้อรัง", "Isosorbide dinitrate 10 mg", "500's", "box", 270],
  ["P146", "d", "ยาโรคเรื้อรัง", "Losartan 50 mg", "300's", "box", 214],
  ["P147", "d", "ยาโรคเรื้อรัง", "Manidipine 20 mg", "100's", "box", 110],
  ["P148", "d", "ยาโรคเรื้อรัง", "Metformin 500 mg", "500's", "box", 130],
  ["P149", "d", "ยาโรคเรื้อรัง", "Methyldopa 250 mg", "500's", "box", 716],
  ["P150", "d", "ยาโรคเรื้อรัง", "Metoprolol tartrate 100 mg", "500's", "box", 190],
  ["P151", "d", "ยาโรคเรื้อรัง", "Pioglitazone 30 mg", "30's", "box", 52],
  ["P152", "d", "ยาโรคเรื้อรัง", "Phenyltoin 100 mg", "1000's", "Bottle", 490],
  ["P153", "d", "ยาโรคเรื้อรัง", "Propranolol 10 mg", "500's", "box", 125],
  ["P154", "d", "ยาโรคเรื้อรัง", "Salbutamol 2 mg", "1000's", "Bottle", 130],
  ["P155", "d", "ยาโรคเรื้อรัง", "Simvastatin 20 mg", "100's", "box", 48.15],
  ["P156", "d", "ยาโรคเรื้อรัง", "Spironolactone 25 mg", "500's", "box", 298],
  ["P157", "d", "ยาโรคเรื้อรัง", "Theophylline 200 mg", "100's", "box", 80],
  ["P158", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Autoclave 0.5\"", "", "ม้วน", 200],
  ["P159", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Blade No.11 (กล่องละ 100 ชิ้น)", "", "ชิ้น", 3.6],
  ["P160", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Blade No.15 (กล่องละ 100 ชิ้น)", "", "ชิ้น", 3.2],
  ["P161", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Bottle ขวดน้ำพลาสติก 30 ml", "", "ขวด", 1.5],
  ["P162", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Condom 49 mm (กล่องละ 100 ชิ้น)", "", "กล่อง", 230],
  ["P163", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Condom 52 mm (กล่องละ 100 ชิ้น)", "", "กล่อง", 440],
  ["P164", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "cotton 0.35 g.", "", "ถุง", 110],
  ["P165", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Cotton Stick NO.S (100 ก้าน)", "", "กล่อง", 120],
  ["P166", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Cotton Stick NO.L (100 ก้าน)", "", "กล่อง", 200],
  ["P167", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Disp. Glove No.S", "", "กล่อง", 130],
  ["P168", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Disp. Glove No.M", "", "กล่อง", 130],
  ["P169", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Disp. Mask 50 pcs.", "", "กล่อง", 60],
  ["P170", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Disp. Needle 18 x 1.5\"", "", "กล่อง", 60],
  ["P171", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Disp. Needle 21 x 1.5\"", "", "กล่อง", 57],
  ["P172", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Disp. Needle 23 x 1.5\"", "", "กล่อง", 57],
  ["P173", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Disp. Needle 24 x 1.5\"", "", "กล่อง", 57],
  ["P174", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Disp. Needle 25 x 1\"", "", "กล่อง", 57],
  ["P175", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Disp. Needle 27 x 0.5\"", "", "กล่อง", 57],
  ["P176", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Disp. Syringe insulin 29G x 0.5\" (ติดเข็มถอดหัวไม่ได้)", "", "กล่อง", 230],
  ["P177", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Disp. Insulin Syringe 1ml 27G x 0.5\" (ถอดหัวเข็มได้)", "", "กล่อง", 265],
  ["P178", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Disp. Syring 3ml.", "", "กล่อง", 135],
  ["P179", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Disp. Syring 5ml.", "", "กล่อง", 160],
  ["P180", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Disp. Syring 10ml.", "", "กล่อง", 230],
  ["P181", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Elastic bandage 3\"", "", "โหล", 240],
  ["P182", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Elastic bandage 4\"", "", "โหล", 310],
  ["P183", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Elastic bandage 6\"", "", "โหล", 415],
  ["P184", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Foley catheter No.14", "", "เส้น", 19],
  ["P185", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Foley catheter No.16", "", "เส้น", 19],
  ["P186", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Foley catheter No.18", "", "เส้น", 19],
  ["P187", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Gauze Bandages 3\" x 6 หลา", "", "โหล", 63],
  ["P188", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Gauze Bandages 4\" x 6 หลา", "", "โหล", 90],
  ["P189", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Gauze Bandages 6\" x 6 หลา", "", "โหล", 123],
  ["P190", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Gauze pad 3 x 4\" 100 pcs.", "", "ห่อ", 69],
  ["P191", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Injection plug", "", "ชิ้น", 8.5],
  ["P192", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "IV Catheter No.22 (กล่องละ 50 ชิ้น)", "", "กล่อง", 460],
  ["P193", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "IV Catheter No.24 (กล่องละ 50 ชิ้น)", "", "กล่อง", 485],
  ["P194", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "IV Set", "", "ชุด", 11],
  ["P195", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "K-Y Jelly 82 g", "", "หลอด", 200],
  ["P196", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Neofix ขนาด 5 cm x 10 m", "", "กล่อง", 165],
  ["P197", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "NG tube No.14", "", "เส้น", 13],
  ["P198", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "NG tube No.16", "", "เส้น", 19],
  ["P199", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "NYLON 3/0 24 mm 3/8 Round", "", "กล่อง", 495],
  ["P200", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "NYLON 4/0 30 mm 3/8 Round", "", "กล่อง", 600],
  ["P201", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Black Braided Silk ไหมเย็บแผล 3-0", "", "ม้วน", 320],
  ["P202", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Black Braided Silk ไหมเย็บแผล 4-0", "", "ม้วน", 320],
  ["P203", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Oxygen Canular (Adult)", "", "ชุด", 13.5],
  ["P204", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Oxygen mask with bag (Adult)", "", "ชุด", 40],
  ["P205", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Oxygen neublizer (Adult)", "", "ชุด", 40],
  ["P206", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Oxygen neublizer (Child)", "", "ชุด", 40],
  ["P207", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Plaster Micropore 0.5\"", "", "กล่อง", 253],
  ["P208", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Plaster Micropore 1\"", "", "กล่อง", 253],
  ["P209", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Plaster Transpore 0.5\"", "", "กล่อง", 253],
  ["P210", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Plaster Transpore 1\"", "", "กล่อง", 253],
  ["P211", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Spore test", "", "ชิ้น", 126],
  ["P212", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Steam Chemical Integrator (แผ่นตรวจสอบการฆ่าเชื้อ)", "", "แผ่น", 6.16],
  ["P213", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Sofra tulle 10 cm x 10 cm (กล่องละ 10 ชิ้น)", "", "กล่อง", 200],
  ["P214", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Sterile glove No.6.5 (กล่องละ 50 ชิ้น)", "", "กล่อง", 840],
  ["P215", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Sterile glove No.7 (กล่องละ 50 ชิ้น)", "", "กล่อง", 800],
  ["P216", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Top dressing 4\" x 6\" Sterile (20 แพ็ค)", "", "ห่อ", 160],
  ["P217", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "Urine bag เทล่าง", "", "ถุง", 19],
  ["P218", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "ซองซิบสีใส ขนาด 15\" x23\"", "", "กิโล", 160],
  ["P219", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "ซองซิบสีใส ขนาด 20\" x30\"", "", "กิโล", 160],
  ["P220", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "ซองซิบสีใส ขนาด 9\" x13\"", "", "กิโล", 160],
  ["P221", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "ซองซิบสีชา ขนาด 15\" x23\"", "", "กิโล", 160],
  ["P222", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "ซองซิบสีชา ขนาด 9\" x13\"", "", "กิโล", 160],
  ["P223", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "ซองซิบสีใส ขนาด 10\" x 15\"", "", "กิโล", 160],
  ["P224", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "ซองซิบสีชาทึบ ขนาด 10\" x 15\"", "", "กิโล", 160],
  ["P225", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "สติ๊กเกอร์ยา", "", "ม้วน", 130],
  ["P226", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "สาย Suction No. 14", "", "เส้น", 4],
  ["P227", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "สาย Suction No. 16", "", "เส้น", 4],
  ["P228", "s", "วัสดุการแพทย์ที่ไม่ใช่ยา", "น้ำยาล้างเครื่องมือแพทย์", "", "แกลลอน", 2400],
  ["P229", "s", "วัสดุวิทยาศาสตร์การแพทย์", "STRIPS (แถบตรวจน้ำตาล) (100 ชิ้น)", "", "กล่อง", 749],
  ["P230", "s", "วัสดุวิทยาศาสตร์การแพทย์", "ถ่านเครื่องเจาะน้ำตาล", "", "ก้อน", 48.15],
  ["P231", "s", "วัสดุวิทยาศาสตร์การแพทย์", "เข็มเจาะปลายนิ้ว (200 ชิ้น)", "", "กล่อง", 399.11],
  ["P232", "s", "วัสดุวิทยาศาสตร์การแพทย์", "UPT test (100 test)", "", "กล่อง", 525],
  ["P233", "s", "วัสดุวิทยาศาสตร์การแพทย์", "ชุดตรวจอัลบูมินกลูโคส (100 test)", "", "กล่อง", 140],
  ["P234", "s", "วัสดุวิทยาศาสตร์การแพทย์", "หลอด Hematocrit สีแดง (100 หลอด)", "", "กล่อง", 71.5],
  ["P235", "s", "วัสดุวิทยาศาสตร์การแพทย์", "ดินน้ำมัน", "", "ก้อน", 71.5],
  ["P236", "s", "วัสดุวิทยาศาสตร์การแพทย์", "เครื่องตรวจน้ำตาล", "", "เครื่อง", 2247]
];
