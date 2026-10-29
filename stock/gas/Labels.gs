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

/* ===================== เพิ่มยาใหม่ + สร้าง QR =====================
 * เมนู คลังใน → เพิ่มยาใหม่ + สร้าง QR : กรอกชื่อยา กดบันทึก ได้รหัสใหม่ (ต่อจากตัวสุดท้าย) พร้อม QR พิมพ์ได้ทันที
 * หรือพิมพ์ชื่อยาในแถวว่างท้ายแท็บ "รายการยา" ระบบจะใส่รหัสและติ๊ก "ใช้งาน" ให้เอง
 */
function nextCode_(sh) {
  const codes = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 1).getValues().map(r => String(r[0]).trim());
  let max = 0;
  codes.forEach(c => { const m = /^P(\d+)$/.exec(c); if (m) max = Math.max(max, Number(m[1])); });
  return 'P' + String(max + 1).padStart(3, '0');
}

function addItemDialog() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.items);
  const vals = sh.getDataRange().getValues().slice(1);
  const shelves = [...new Set(vals.map(r => String(r[4] || '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'th'));
  const units = [...new Set(vals.map(r => String(r[3] || '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'th'));
  const t = HtmlService.createTemplate(ADD_HTML_);
  t.shelves = JSON.stringify(shelves); t.units = JSON.stringify(units); t.base = LABEL_BASE;
  t.names = JSON.stringify(vals.map(r => [String(r[0]), String(r[1]), String(r[2])]));
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(480).setHeight(620), 'เพิ่มยาใหม่ + สร้าง QR');
}

/* เรียกจากหน้าต่างเพิ่มยา: บันทึกแถวใหม่ คืนรหัส */
function addItem(f) {
  const n = String(f.n || '').trim(), th = String(f.t || '').trim();
  if (!n && !th) throw new Error('ใส่ชื่อยาก่อน');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.items);
    const code = nextCode_(sh);
    const row = [code, n || th, n ? th : '', String(f.u || '').trim(), String(f.s || '').trim(), Number(f.min) || '', true, String(f.hub || '').trim()];
    sh.appendRow(row);
    sh.getRange(sh.getLastRow(), 7).insertCheckboxes().setValue(true);
    return { c: code, n: row[1], t: row[2], s: row[4] };
  } finally { lock.releaseLock(); }
}

/* พิมพ์ชื่อยาในแถวใหม่ของแท็บรายการยา → ใส่รหัสและติ๊กใช้งานให้ (simple trigger) */
function onEdit(e) {
  try {
    const sh = e && e.range && e.range.getSheet();
    if (!sh || sh.getName() !== SHEET.items || e.range.getRow() < 2) return;
    const c = e.range.getColumn();
    if (c !== 2 && c !== 3) return;
    for (let r = e.range.getRow(); r < e.range.getRow() + e.range.getNumRows(); r++) {
      const v = sh.getRange(r, 1, 1, 7).getValues()[0];
      if (!String(v[0]).trim() && String(v[1] || v[2]).trim()) {
        sh.getRange(r, 1).setValue(nextCode_(sh));
        sh.getRange(r, 7).insertCheckboxes().setValue(true);
      }
    }
  } catch (_) {}
}

const ADD_HTML_ = `<!doctype html><html><head><base target="_top">
<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcode-generator/1.4.4/qrcode.min.js"></script>
<style>body{font-family:sans-serif;font-size:14px;color:#15302b;margin:0;padding:2px 4px}
label{display:block;font-size:12.5px;color:#5d7470;margin:9px 0 3px}
input,select{width:100%;box-sizing:border-box;font:inherit;font-size:15px;padding:8px 10px;border:1px solid #c9d6d2;border-radius:8px}
.row{display:flex;gap:8px}.row>div{flex:1}
.btn{padding:11px 16px;border-radius:10px;border:1px solid #c9d6d2;background:#fff;font-weight:bold;font-size:15px;cursor:pointer}
.primary{background:#0f7a68;border-color:#0f7a68;color:#fff}
.warn{background:#fbeedd;color:#8a4a06;border-radius:8px;padding:6px 10px;font-size:12.5px;margin-top:6px}
.done{display:flex;gap:12px;align-items:center;background:#e0f1ec;border-radius:12px;padding:10px;margin-top:6px}
.done svg{width:110px;height:110px;flex:none;background:#fff;padding:4px}
.done b{font-size:16px}.muted{color:#5d7470;font-size:12.5px}
[hidden]{display:none!important}</style></head><body>
<div id="form">
  <label>ชื่อยา (อังกฤษ)</label><input id="n" placeholder="เช่น Loratadine 10 mg" autocomplete="off">
  <div class="warn" id="dup" hidden></div>
  <label>ชื่อไทย (ใส่การใช้ในวงเล็บได้)</label><input id="t" placeholder="เช่น ลอราทาดีน เม็ด 10 มก. (ยาแก้แพ้)">
  <div class="row"><div><label>หน่วยนับ</label><input id="u" list="ul" placeholder="กล่อง"><datalist id="ul"></datalist></div>
  <div><label>หมวด/ชั้น</label><select id="s"></select></div></div>
  <div class="row"><div><label>จุดสั่งซื้อ (ไม่บังคับ)</label><input id="min" type="number" min="0"></div>
  <div><label>ชื่อในคลังบวกหมื้อ (ไม่บังคับ)</label><input id="hub"></div></div>
  <p><button class="btn primary" id="save">บันทึก + สร้าง QR</button></p>
</div>
<div id="ok" hidden>
  <div class="done"><div id="qr"></div><div><b id="okName"></b><div id="okTh"></div><div class="muted" id="okCode"></div><div class="muted">เพิ่มในแท็บรายการยาแล้ว เปิดในแอปได้ทันที</div></div></div>
  <label>จำนวนดวงที่จะพิมพ์</label><input id="copies" type="number" min="1" max="21" value="1" style="width:90px">
  <p><button class="btn primary" id="print">พิมพ์สติกเกอร์</button> <button class="btn" id="more">เพิ่มอีกตัว</button></p>
</div>
<script>
const SHELVES=<?!= shelves ?>, UNITS=<?!= units ?>, NAMES=<?!= names ?>, BASE=<?!= JSON.stringify(base) ?>;
const $=id=>document.getElementById(id);
const esc=s=>String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
$('s').innerHTML=SHELVES.map(s=>'<option>'+esc(s)+'</option>').join('');
$('ul').innerHTML=UNITS.map(u=>'<option value="'+esc(u)+'">').join('');
function qr(t){const q=qrcode(0,'M');q.addData(t);q.make();return q.createSvgTag({cellSize:4,margin:0,scalable:true})}
const norm=s=>String(s).toLowerCase().replace(/[^a-z0-9ก-๙]/g,'');
function checkDup(){const a=norm($('n').value+$('t').value);if(a.length<4){$('dup').hidden=true;return}
  const k=norm($('n').value).slice(0,6);const hit=NAMES.filter(r=>k.length>=4&&norm(r[1]).startsWith(k)).slice(0,3);
  $('dup').hidden=!hit.length;$('dup').innerHTML='มีชื่อคล้ายกันอยู่แล้ว: '+hit.map(r=>esc(r[0]+' '+r[1])).join(', ')+' — ถ้าเป็นยาตัวเดียวกันไม่ต้องเพิ่มใหม่'}
$('n').oninput=checkDup;
let made=null;
$('save').onclick=()=>{const f={n:$('n').value,t:$('t').value,u:$('u').value,s:$('s').value,min:$('min').value,hub:$('hub').value};
  if(!f.n.trim()&&!f.t.trim()){$('n').focus();return}
  $('save').disabled=true;$('save').textContent='กำลังบันทึก…';
  google.script.run.withSuccessHandler(r=>{made=r;NAMES.push([r.c,r.n,r.t]);$('qr').innerHTML=qr(BASE+r.c);$('okName').textContent=r.t||r.n;$('okTh').textContent=r.t?r.n:'';$('okCode').textContent=r.c+(r.s?' · '+r.s:'');
    $('form').hidden=true;$('ok').hidden=false;$('save').disabled=false;$('save').textContent='บันทึก + สร้าง QR'})
  .withFailureHandler(e=>{alert(e.message);$('save').disabled=false;$('save').textContent='บันทึก + สร้าง QR'}).addItem(f)};
$('more').onclick=()=>{['n','t','min','hub'].forEach(id=>$(id).value='');$('dup').hidden=true;$('ok').hidden=true;$('form').hidden=false;$('n').focus()};
$('print').onclick=()=>{const k=Math.max(1,Math.min(21,Number($('copies').value)||1)),i=made;
  const lab='<div class="lab"><div class="q">'+qr(BASE+i.c)+'</div><div class="tx"><div class="th">'+esc(i.t||i.n)+'</div>'+(i.t?'<div class="en">'+esc(i.n)+'</div>':'')+'<div class="cd">'+esc(i.c)+(i.s?' · '+esc(i.s):'')+'</div><div class="scan">สแกนก่อนหยิบทุกครั้ง</div></div></div>';
  const w=window.open('','_blank');if(!w){alert('เบราว์เซอร์บล็อกหน้าต่างใหม่ ให้อนุญาตป๊อปอัปของ docs.google.com');return}
  w.document.open();w.document.write('<!doctype html><html lang="th"><head><meta charset="utf-8"><title>สติกเกอร์ '+esc(i.c)+'</title><link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Anuphan:wght@600;700&family=Sarabun:wght@400;500&display=swap"><style>*{box-sizing:border-box}body{margin:0;font-family:Sarabun,sans-serif}.sheet{width:210mm;padding:8mm 6mm;display:grid;grid-template-columns:repeat(3,1fr);grid-auto-rows:39.5mm;gap:1.5mm}.lab{border:1px dashed #b9c9c5;border-radius:2mm;padding:2mm;display:flex;gap:2mm;align-items:center;overflow:hidden}.q{width:28mm;height:28mm;flex:none}.q svg{width:100%;height:100%;display:block}.tx{min-width:0;display:flex;flex-direction:column;gap:.6mm}.th{font:700 11pt/1.15 Anuphan,sans-serif;overflow-wrap:anywhere}.en{font-size:7.5pt;line-height:1.2;color:#333;overflow-wrap:anywhere}.cd{font-size:7pt;color:#666}.scan{font-size:6.5pt;color:#0f7a68;font-weight:500}@media print{@page{size:A4;margin:0}}</style></head><body><div class="sheet">'+lab.repeat(k)+'</div></body></html>');
  w.document.close();setTimeout(()=>{try{w.focus();w.print()}catch(e){}},1500)};
</script></body></html>`;
