/**
 * ล็อกการเข้าใช้: เครื่องใหม่ต้องขออนุมัติก่อนใช้งาน
 * - เปิดแอปครั้งแรก → กรอกชื่อ → ระบบส่งอีเมลหาผู้ดูแล (ADMIN_EMAIL)
 * - ผู้ดูแลกดปุ่ม "อนุมัติ" ในอีเมล → เครื่องนั้นใช้งานได้ทันที
 * - ยกเลิกสิทธิ์: เอาติ๊กออกในแท็บ "อุปกรณ์" ช่อง "อนุมัติ" (มีผลภายใน 1 นาที)
 * ลิงก์ที่ถูกส่งต่อไปเครื่องอื่นจะเห็นแค่หน้าขอสิทธิ์ ไม่เห็นข้อมูลยา
 */
const DEV_SHEET = 'อุปกรณ์';
const DEV_HEADER = ['tok', 'ชื่อ', 'ตำแหน่ง/หน้าที่', 'อีเมล', 'ข้อมูลเครื่อง', 'ขอเมื่อ', 'อนุมัติ', 'ตัดสินเมื่อ', 'req_id'];

function devSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(DEV_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DEV_SHEET);
    sh.getRange(1, 1, 1, DEV_HEADER.length).setValues([DEV_HEADER]).setFontWeight('bold').setBackground('#e0f1ec');
    sh.setFrozenRows(1);
    sh.hideColumns(1);
    sh.hideColumns(9);
    sh.getRange('F:F').setNumberFormat('dd/mm/yyyy hh:mm');
    sh.getRange('H:H').setNumberFormat('dd/mm/yyyy hh:mm');
  }
  return sh;
}

function accessSettings_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.settings);
  const vals = sh.getDataRange().getValues();
  const get = k => { const r = vals.find(v => v[0] === k); return r ? String(r[1]) : ''; };
  let email = get('ADMIN_EMAIL'), key = get('ADMIN_KEY');
  if (!email) { email = Session.getEffectiveUser().getEmail(); sh.appendRow(['ADMIN_EMAIL', email, 'อีเมลที่รับคำขอใช้งานจากเครื่องใหม่ (เปลี่ยนได้)']); }
  if (!key) { key = Utilities.getUuid().replace(/-/g, ''); sh.appendRow(['ADMIN_KEY', key, 'รหัสลับใช้ยืนยันลิงก์อนุมัติในอีเมล ห้ามแชร์']); }
  return { email, key, facility: get('FACILITY') || DEFAULTS.FACILITY };
}

function sig_(id, key) {
  const raw = Utilities.computeHmacSha256Signature(String(id), key);
  return Utilities.base64EncodeWebSafe(raw).replace(/=+$/, '').slice(0, 22);
}

function approvedSet_() {
  const cache = CacheService.getScriptCache();
  const c = cache.get('devok');
  if (c) return JSON.parse(c);
  const sh = devSheet_();
  const n = sh.getLastRow() - 1;
  const ok = n > 0 ? sh.getRange(2, 1, n, 7).getValues().filter(r => r[0] && r[6] === true).map(r => String(r[0])) : [];
  try { cache.put('devok', JSON.stringify(ok), 60); } catch (_) {}
  return ok;
}
function isApproved_(tok) { return !!tok && approvedSet_().indexOf(String(tok)) >= 0; }

function devRow_(tok) {
  const sh = devSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return null;
  const vals = sh.getRange(2, 1, n, 9).getValues();
  const i = vals.findIndex(r => String(r[0]) === String(tok));
  return i < 0 ? null : { row: i + 2, v: vals[i] };
}

function devStatus_(tok) {
  const r = tok ? devRow_(tok) : null;
  if (!r) return 'none';
  if (r.v[6] === true) return 'approved';
  return r.v[7] ? 'rejected' : 'pending';
}

function register_(b) {
  const tok = String(b.tok || '');
  if (!/^[a-f0-9]{32}$/.test(tok)) throw new Error('ข้อมูลเครื่องไม่ถูกต้อง ลองรีเฟรชหน้า');
  const name = String(b.name || '').trim().slice(0, 60), role = String(b.role || '').trim().slice(0, 60), email = String(b.email || '').trim().slice(0, 80);
  if (!name) throw new Error('ใส่ชื่อก่อนส่งคำขอ');
  if (email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('อีเมลไม่ถูกต้อง');
  const ex = devRow_(tok);
  if (ex) return { ok: true, status: devStatus_(tok) };
  const sh = devSheet_();
  const n = sh.getLastRow() - 1;
  const pending = n > 0 ? sh.getRange(2, 7, n, 2).getValues().filter(r => r[0] !== true && !r[1]).length : 0;
  if (pending >= 30) throw new Error('มีคำขอรออนุมัติมากเกินไป ติดต่อผู้ดูแลระบบ');
  const s = accessSettings_();
  const id = Utilities.getUuid().slice(0, 12);
  sh.appendRow([tok, name, role, email, String(b.ua || '').slice(0, 120), new Date(), false, '', id]);
  sh.getRange(sh.getLastRow(), 7).insertCheckboxes();
  const url = ScriptApp.getService().getUrl();
  const link = a => `${url}?a=${a}&r=${id}&s=${sig_(id + a, s.key)}`;
  const btn = (href, text, bg) => `<a href="${href}" style="display:inline-block;padding:12px 22px;margin:4px 6px 4px 0;border-radius:10px;background:${bg};color:#fff;text-decoration:none;font-weight:bold">${text}</a>`;
  MailApp.sendEmail({
    to: s.email,
    subject: `[คลังใน ${s.facility}] ขอใช้งาน: ${name}`,
    htmlBody: `<div style="font-family:sans-serif;font-size:15px;line-height:1.6">
      <p><b>${esc_(name)}</b> ขอใช้งานแอปคลังใน ${esc_(s.facility)}</p>
      <table style="border-collapse:collapse">
        <tr><td style="color:#666;padding-right:12px">ตำแหน่ง/หน้าที่</td><td>${esc_(role) || '-'}</td></tr>
        <tr><td style="color:#666;padding-right:12px">อีเมล</td><td>${esc_(email) || '-'}</td></tr>
        <tr><td style="color:#666;padding-right:12px">เครื่อง</td><td style="font-size:12px">${esc_(String(b.ua || '').slice(0, 120))}</td></tr>
        <tr><td style="color:#666;padding-right:12px">เวลา</td><td>${Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm')}</td></tr>
      </table>
      <p>${btn(link('approve'), 'อนุมัติ', '#0f7a68')} ${btn(link('reject'), 'ไม่อนุมัติ', '#b3302b')}</p>
      <p style="color:#666;font-size:13px">ยกเลิกสิทธิ์ภายหลังได้ที่แท็บ "อุปกรณ์" ในชีต โดยเอาติ๊กช่อง "อนุมัติ" ออก</p></div>`,
  });
  return { ok: true, status: 'pending' };
}

function decide_(p) {
  const s = accessSettings_();
  const a = p.a === 'approve' ? 'approve' : 'reject';
  const page = (title, msg, color) => HtmlService.createHtmlOutput(
    `<meta name="viewport" content="width=device-width,initial-scale=1"><div style="font-family:sans-serif;max-width:420px;margin:40px auto;padding:24px;border-radius:16px;background:#f3f6f5;text-align:center">
     <div style="font-size:22px;font-weight:bold;color:${color}">${title}</div><p style="font-size:16px;color:#15302b">${msg}</p></div>`).setTitle(title);
  if (!p.r || p.s !== sig_(p.r + a, s.key)) return page('ลิงก์ไม่ถูกต้อง', 'ลิงก์นี้ใช้ไม่ได้ ให้เปิดจากอีเมลคำขอโดยตรง', '#b3302b');
  const sh = devSheet_();
  const n = sh.getLastRow() - 1;
  const vals = n > 0 ? sh.getRange(2, 1, n, 9).getValues() : [];
  const i = vals.findIndex(r => String(r[8]) === p.r);
  if (i < 0) return page('ไม่พบคำขอ', 'คำขอนี้อาจถูกลบออกจากแท็บอุปกรณ์แล้ว', '#b3302b');
  const row = i + 2, name = vals[i][1], email = vals[i][3];
  sh.getRange(row, 7, 1, 2).setValues([[a === 'approve', new Date()]]);
  CacheService.getScriptCache().remove('devok');
  if (a === 'approve' && email) {
    try { MailApp.sendEmail(String(email), `[คลังใน ${s.facility}] อนุมัติการใช้งานแล้ว`, `สวัสดี ${name}\n\nเครื่องของคุณได้รับอนุมัติให้ใช้แอปคลังใน ${s.facility} แล้ว เปิดแอปอีกครั้งเพื่อเริ่มใช้งาน\n\nhttps://focus2541.github.io/pharmacy-scanner/stock/`); } catch (_) {}
  }
  return a === 'approve'
    ? page('อนุมัติแล้ว', `${esc_(name)} ใช้งานแอปคลังในได้แล้ว`, '#0f7a68')
    : page('ไม่อนุมัติ', `${esc_(name)} จะใช้งานแอปไม่ได้`, '#b3302b');
}

function esc_(s) { return String(s || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

/* กดรันครั้งเดียวหลังวางโค้ด เพื่ออนุญาตให้สคริปต์ส่งอีเมล และสร้างแท็บอุปกรณ์ */
function setupAccess() {
  accessSettings_();
  devSheet_();
  Logger.log('ส่งอีเมลได้อีกวันนี้: ' + MailApp.getRemainingDailyQuota() + ' ฉบับ · อีเมลผู้ดูแล: ' + accessSettings_().email);
}
