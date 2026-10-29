/**
 * สติกเกอร์ QR จากชีต: เมนู คลังใน → พิมพ์สติกเกอร์ QR
 * - เลือกแถวในแท็บ "รายการยา" ก่อน (ลากคลุมหลายแถวได้) จะพิมพ์เฉพาะยาที่เลือก
 * - ไม่ได้เลือก หรืออยู่แท็บอื่น จะพิมพ์ยาทุกตัวที่ติ๊ก "ใช้งาน"
 * QR สร้างใหม่ทุกครั้งจากรหัสยาในชีต ยาที่เพิ่มใหม่จึงพิมพ์ได้ทันที
 * กระดาษสติกเกอร์ A4 แผ่นละ 21 ดวง (3 × 7)
 */
const LABEL_BASE = 'https://focus2541.github.io/pharmacy-scanner/stock/#';

function labelItems_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET.items);
  const all = sh.getDataRange().getValues().slice(1).filter(r => r[0])
    .map(r => ({ c: String(r[0]).trim(), n: String(r[1] || ''), t: String(r[2] || ''), s: String(r[4] || ''), on: r[6] === true || String(r[6]).toUpperCase() === 'TRUE' }));
  let picked = null;
  const act = ss.getActiveSheet();
  if (act && act.getName() === SHEET.items) {
    const rows = new Set();
    const list = act.getActiveRangeList();
    (list ? list.getRanges() : []).forEach(rg => { for (let r = rg.getRow(); r < rg.getRow() + rg.getNumRows(); r++) if (r > 1) rows.add(r); });
    if (rows.size) {
      const codes = new Set([...rows].map(r => String(act.getRange(r, 1).getValue()).trim()).filter(Boolean));
      picked = all.filter(i => codes.has(i.c));
    }
  }
  const items = picked && picked.length ? picked : all.filter(i => i.on);
  return { items, picked: !!(picked && picked.length), facility: settings_().FACILITY };
}

function labelsDialog() {
  const d = labelItems_();
  const t = HtmlService.createTemplate(LABELS_HTML_);
  t.data = JSON.stringify(d.items).replace(/</g, '\\u003c');
  t.base = LABEL_BASE;
  t.note = d.picked ? `พิมพ์เฉพาะยาที่เลือกในแท็บรายการยา ${d.items.length} ตัว` : `พิมพ์ยาทุกตัวที่ใช้งาน ${d.items.length} ตัว (ถ้าจะพิมพ์บางตัว ให้เลือกแถวในแท็บรายการยาก่อน)`;
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(460).setHeight(300), 'สติกเกอร์ QR');
}

const LABELS_HTML_ = `<!doctype html><html><head><base target="_top">
<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcode-generator/1.4.4/qrcode.min.js"></script>
<style>body{font-family:sans-serif;font-size:14px;color:#15302b;margin:0;padding:4px}
.btn{display:inline-block;padding:11px 18px;border-radius:10px;border:0;background:#0f7a68;color:#fff;font-weight:bold;font-size:15px;cursor:pointer}
.note{color:#5d7470;font-size:13px;line-height:1.5}</style></head><body>
<p id="msg"><?= note ?></p>
<p><button class="btn" id="go">เปิดหน้าพิมพ์</button></p>
<p class="note">หน้าพิมพ์จะเปิดในแท็บใหม่ แล้วกดพิมพ์ (Ctrl+P) เลือกกระดาษ A4 ขอบ "ไม่มี" หรือเลือก "บันทึกเป็น PDF" ถ้าต้องการไฟล์</p>
<script>
const items=<?!= data ?>, BASE=<?!= JSON.stringify(base) ?>;
const esc=s=>String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
function qr(t){const q=qrcode(0,'M');q.addData(t);q.make();return q.createSvgTag({cellSize:4,margin:0,scalable:true})}
function page(){
  let s='';
  for(let p=0;p<items.length;p+=21){
    s+='<div class="sheet">'+items.slice(p,p+21).map(i=>'<div class="lab"><div class="qr">'+qr(BASE+i.c)+'</div><div class="tx"><div class="th">'+esc(i.t||i.n)+'</div>'+(i.t?'<div class="en">'+esc(i.n)+'</div>':'')+'<div class="cd">'+esc(i.c)+(i.s?' · '+esc(i.s):'')+'</div><div class="scan">สแกนก่อนหยิบทุกครั้ง</div></div></div>').join('')+'</div>';
  }
  return '<!doctype html><html lang="th"><head><meta charset="utf-8"><title>สติกเกอร์ QR คลังใน</title>'+
  '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Anuphan:wght@600;700&family=Sarabun:wght@400;500&display=swap">'+
  '<style>*{box-sizing:border-box}body{margin:0;background:#eee;font-family:Sarabun,sans-serif}'+
  '.sheet{width:210mm;height:297mm;margin:10px auto;background:#fff;padding:8mm 6mm;display:grid;grid-template-columns:repeat(3,1fr);grid-auto-rows:39.5mm;gap:1.5mm;break-after:page}'+
  '.lab{border:1px dashed #b9c9c5;border-radius:2mm;padding:2mm;display:flex;gap:2mm;align-items:center;overflow:hidden}'+
  '.qr{width:28mm;height:28mm;flex:none}.qr svg{width:100%;height:100%;display:block}'+
  '.tx{min-width:0;display:flex;flex-direction:column;gap:.6mm}.th{font:700 11pt/1.15 Anuphan,sans-serif;overflow-wrap:anywhere}'+
  '.en{font-size:7.5pt;line-height:1.2;color:#333;overflow-wrap:anywhere}.cd{font-size:7pt;color:#666}.scan{font-size:6.5pt;color:#0f7a68;font-weight:500}'+
  '@media print{@page{size:A4;margin:0}body{background:#fff}.sheet{margin:0}}</style></head><body>'+s+'</body></html>';
}
document.getElementById('go').onclick=()=>{
  const w=window.open('','_blank');
  if(!w){document.getElementById('msg').textContent='เบราว์เซอร์บล็อกหน้าต่างใหม่ ให้อนุญาตป๊อปอัปของ docs.google.com แล้วกดอีกครั้ง';return}
  w.document.open();w.document.write(page());w.document.close();
  setTimeout(()=>{try{w.focus();w.print()}catch(e){}},1500);
  google.script.host.close();
};
</script></body></html>`;
