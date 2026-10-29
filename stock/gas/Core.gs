/**
 * ตรรกะคำนวณยอด (ไม่แตะชีตโดยตรง ทดสอบแยกได้)
 * ยอดคงเหลือของยาแต่ละตัว = ผลรวม "คงเหลือ" ของทุกล็อต
 * หยิบออก: ตัดจากล็อตที่หมดอายุเร็วสุดก่อน
 * ถ้าหยิบเกินยอดในระบบ ยอดจะติดลบ และขึ้นสถานะ "ต้องนับใหม่"
 */
var Core = (function () {
  function uid(p) { return p + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }
  function int(v) { const n = Math.round(Number(v)); return isFinite(n) ? n : NaN; }
  function item(db, code) {
    const it = db.items.find(i => i.code === String(code));
    if (!it) throw new Error('ไม่พบรหัสยา ' + code);
    return it;
  }
  function lotsOf(db, code) { return db.lots.filter(l => l.code === code); }
  function bal(db, code) { return lotsOf(db, code).reduce((s, l) => s + l.left, 0); }
  const byExpAsc = (a, b) => (a.exp || '9999') < (b.exp || '9999') ? -1 : (a.exp || '9999') > (b.exp || '9999') ? 1 : 0;
  function touch(db, l) { if (l.row) db.dirtyLots[l.id] = 1; }
  function newLot(db, it, exp, qin, left, now) {
    const l = { id: uid('L'), code: it.code, name: it.name, exp: exp || '', qin: qin, left: left, date: iso(now) };
    db.lots.push(l); db.newLots.push(l); return l;
  }
  function iso(d) { const z = new Date(d.getTime() + 7 * 3600e3); return z.toISOString().slice(0, 10); }
  function addMove(db, m, now) {
    const mv = Object.assign({ id: uid('M'), time: now.toISOString(), exp: '', note: '', from: '', dev: '', by: '' }, m);
    db.moves.push(mv); db.newMoves.push(mv); return mv;
  }
  /* ล็อตติดลบ (จากการหยิบเกิน) หักกลบกับล็อตบวก */
  function settle(db, code) {
    const ls = lotsOf(db, code);
    ls.filter(l => l.left < 0).forEach(neg => {
      ls.filter(l => l.left > 0).sort(byExpAsc).forEach(pos => {
        if (neg.left >= 0) return;
        const m = Math.min(pos.left, -neg.left);
        pos.left -= m; neg.left += m; touch(db, pos); touch(db, neg);
      });
    });
  }

  function take(db, b, now) {
    const it = item(db, b.c), q = int(b.q);
    if (!(q >= 1)) throw new Error('จำนวนต้องเป็น 1 ขึ้นไป');
    const before = bal(db, it.code);
    let rest = q; const from = [];
    lotsOf(db, it.code).filter(l => l.left > 0).sort(byExpAsc).forEach(l => {
      if (!rest) return;
      const m = Math.min(l.left, rest);
      l.left -= m; rest -= m; touch(db, l); from.push([l.id, m]);
    });
    if (rest > 0) {
      const ls = lotsOf(db, it.code).sort(byExpAsc);
      const l = ls.length ? ls[ls.length - 1] : newLot(db, it, '', 0, 0, now);
      l.left -= rest; touch(db, l); from.push([l.id, rest]);
    }
    const after = bal(db, it.code);
    const mv = addMove(db, { type: 'หยิบ', code: it.code, name: it.name, qty: q, before, after, by: b.by || '', dev: b.dev || '',
      note: after < 0 ? 'ยอดในระบบไม่พอ ต้องนับใหม่' : '', from: JSON.stringify(from) }, now);
    return { ok: true, move: mv, code: it.code, bal: after };
  }

  /* ตัดจำหน่าย (หมดอายุ / ชำรุด / อื่นๆ): ตัดจากล็อตวันหมดอายุที่เลือกก่อน ที่เหลือตัดแบบหมดอายุก่อนออกก่อน
     แยกประเภทจาก "หยิบ" เพื่อไม่ให้ไปนับเป็นยอดใช้ */
  function dispose(db, b, now) {
    const it = item(db, b.c), q = int(b.q);
    if (!(q >= 1)) throw new Error('จำนวนต้องเป็น 1 ขึ้นไป');
    const before = bal(db, it.code);
    if (q > before) throw new Error('ยอดในระบบมี ' + before + ' ตัดได้ไม่เกินนี้ ถ้าของจริงไม่ตรงให้นับใหม่ก่อน');
    const exp = String(b.exp || '');
    let rest = q; const from = [];
    const ls = lotsOf(db, it.code).filter(l => l.left > 0).sort(byExpAsc);
    const order = exp ? ls.filter(l => l.exp === exp).concat(ls.filter(l => l.exp !== exp)) : ls;
    order.forEach(l => {
      if (!rest) return;
      const m = Math.min(l.left, rest);
      l.left -= m; rest -= m; touch(db, l); from.push([l.id, m]);
    });
    const after = bal(db, it.code);
    const reason = String(b.reason || 'หมดอายุ');
    const mv = addMove(db, { type: 'ตัดจำหน่าย', code: it.code, name: it.name, qty: q, before, after, exp, by: b.by || '', dev: b.dev || '',
      note: reason + (b.note ? ' · ' + b.note : ''), from: JSON.stringify(from) }, now);
    return { ok: true, move: mv, code: it.code, bal: after };
  }

  /* ยอดใช้ (หยิบที่ไม่ถูกยกเลิก) ต่อรหัส ในช่วง days วันล่าสุด  rows = [[time, type, code, qty, note], ...] */
  function usage(rows, days, now) {
    const since = now.getTime() - days * 864e5, out = {};
    let first = null;
    rows.forEach(r => {
      const t = new Date(r[0]).getTime();
      if (!isFinite(t)) return;
      if (first === null || t < first) first = t;
      if (t < since || r[1] !== 'หยิบ' || /ยกเลิกแล้ว/.test(String(r[4] || ''))) return;
      out[r[2]] = (out[r[2]] || 0) + (Number(r[3]) || 0);
    });
    const span = first === null ? 0 : Math.min(days, Math.max(1, (now.getTime() - Math.max(first, since)) / 864e5));
    return { days: Math.round(span), use: out };
  }

  function receive(db, b, now) {
    const it = item(db, b.c), q = int(b.q);
    if (!(q >= 1)) throw new Error('จำนวนต้องเป็น 1 ขึ้นไป');
    let exp = String(b.exp || '').slice(0, 10);
    if (!exp && db.hubExp) { try { exp = db.hubExp(it.hub || it.name) || ''; } catch (_) { exp = ''; } }
    const before = bal(db, it.code);
    const same = lotsOf(db, it.code).find(l => l.exp === exp && l.left >= 0 && l.date === iso(now));
    if (same) { same.left += q; same.qin += q; touch(db, same); } else newLot(db, it, exp, q, q, now);
    settle(db, it.code);
    const after = bal(db, it.code);
    const mv = addMove(db, { type: 'รับเข้า', code: it.code, name: it.name, qty: q, before, after, exp, by: b.by || '', dev: b.dev || '', note: b.note || '' }, now);
    return { ok: true, move: mv, code: it.code, bal: after, exp };
  }

  /* นับจริง: ยอดที่นับได้คือยอดใหม่ ของที่เหลือถือว่าเป็นล็อตที่หมดอายุช้าสุด (ของเก่าถูกใช้ไปก่อน) */
  function count(db, b, now) {
    const it = item(db, b.c), q = int(b.q);
    if (!(q >= 0)) throw new Error('จำนวนต้องเป็น 0 ขึ้นไป');
    const before = bal(db, it.code);
    const ls = lotsOf(db, it.code);
    ls.forEach(l => { if (l.left !== 0) { l.left = 0; touch(db, l); } });
    let rest = q;
    ls.filter(l => l.qin > 0).sort(byExpAsc).reverse().forEach(l => {
      if (!rest) return;
      const m = Math.min(l.qin, rest);
      l.left = m; rest -= m; touch(db, l);
    });
    if (rest > 0) {
      const latest = ls.filter(l => l.exp).sort(byExpAsc).pop();
      if (latest) { latest.left += rest; touch(db, latest); }
      else {
        let exp = String(b.exp || '');
        if (!exp && db.hubExp) { try { exp = db.hubExp(it.hub || it.name) || ''; } catch (_) {} }
        newLot(db, it, exp, rest, rest, now);
      }
    }
    const after = bal(db, it.code), diff = after - before;
    const mv = addMove(db, { type: 'นับ', code: it.code, name: it.name, qty: q, before, after, by: b.by || '', dev: b.dev || '',
      note: diff === 0 ? 'ตรงกับระบบ' : (diff > 0 ? 'เกิน ' : 'ขาด ') + Math.abs(diff) }, now);
    return { ok: true, move: mv, code: it.code, bal: after, diff };
  }

  function undo(db, b, now) {
    const m = db.moves.find(x => x.id === String(b.id));
    if (!m || m.type !== 'หยิบ') throw new Error('ไม่พบรายการที่จะยกเลิก');
    if (/ยกเลิกแล้ว/.test(m.note)) throw new Error('รายการนี้ยกเลิกไปแล้ว');
    if (m.dev && b.dev && m.dev !== b.dev) throw new Error('ยกเลิกได้เฉพาะเครื่องที่บันทึก');
    if (now - new Date(m.time) > UNDO_MINUTES_ * 60e3) throw new Error('เกิน ' + UNDO_MINUTES_ + ' นาทีแล้ว ให้ผู้ดูแลนับใหม่แทน');
    const before = bal(db, m.code);
    let from = []; try { from = JSON.parse(m.from || '[]'); } catch (_) {}
    from.forEach(([id, q]) => { const l = db.lots.find(x => x.id === id); if (l) { l.left += q; touch(db, l); } });
    settle(db, m.code);
    m.note = (m.note ? m.note + ' · ' : '') + 'ยกเลิกแล้ว'; if (m.row) db.dirtyMoves[m.id] = 1;
    const after = bal(db, m.code);
    const mv = addMove(db, { type: 'ยกเลิก', code: m.code, name: m.name, qty: m.qty, before, after, by: b.by || m.by, dev: b.dev || '', note: 'ยกเลิกการหยิบ ' + m.id }, now);
    return { ok: true, move: mv, code: m.code, bal: after };
  }
  var UNDO_MINUTES_ = 15;

  function snapshot(db, now) {
    const today = iso(now);
    const last = {};
    db.moves.forEach(m => { last[m.code] = m; });
    const items = db.items.filter(i => i.active).map(i => {
      const ls = lotsOf(db, i.code);
      const lots = ls.filter(l => l.left > 0).sort(byExpAsc).map(l => [l.exp, l.left]);
      const lm = last[i.code];
      const lc = db.moves.filter(m => m.code === i.code && m.type === 'นับ').pop();
      return { c: i.code, n: i.name, t: i.th, u: i.unit, s: i.shelf, min: i.min, b: ls.reduce((s, l) => s + l.left, 0), lots, has: ls.length > 0,
        last: lm ? [lm.type, lm.qty, lm.by, lm.time] : null, counted: lc ? lc.time : '' };
    });
    const recent = db.moves.slice(-60).reverse().map(m => [m.id, m.time, m.type, m.code, m.qty, m.after, m.by, m.note, m.dev]);
    return { facility: db.settings.FACILITY || '', today, items, recent };
  }

  return { take, receive, count, undo, dispose, usage, snapshot, bal };
})();

if (typeof module !== 'undefined') module.exports = Core;
