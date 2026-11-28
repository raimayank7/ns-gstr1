/* NSG3F: GSTR-3B from the files a business already has (no bills saved in NSBiz needed). Pure, no screen: gstr3b.js draws it,
   Node tests use it. Nothing here sends anything anywhere: every file stays in the browser.

   Outward side (Table 3.1 a/b/c/e, 3.2), one of:
   - a sales register, any Excel / CSV (Tally, Busy, Marg, the person's own sheet): register(sheets, 's'). The heading row and every
     column are found by their words; tax from IGST / CGST / SGST / cess columns, or from a GST rate (or a total tax column) with the
     place of supply. Optional columns tell exports, SEZ, nil, exempt, non-GST, credit / debit notes and reverse charge.
   - a Tally Day Book XML (also Marg, Vyapar, myBillBook "export to Tally"): tally(texts) through NSTallyIn.plan.
   - the GSTR-1 the person filed or prepared: the portal JSON (gstr1Json) or the GST Returns Offline Tool workbook (gstr1Sheets).
   Inward side (Table 4, 3.1(d), 5):
   - the portal's GSTR-2B (JSON or Excel), read by NSG2.read2b / read2bSheets of engine/g2core.js (the GSTR-2B site's reader; this file
     only uses it) and / or a purchase register (register(sheets, 'p')), whose optional columns mark reverse charge, blocked credit
     (section 17(5)), ineligible credit (section 16(4), place of supply), imports, nil / exempt / composition / non-GST purchases.
   compute(D, opts) -> the boxes of GSTR-3B in the same shape gstr3b.js uses (zero()), Table 3.2, and one message per row that was left
   out or needs a look (errors first).

   Where each figure goes (GST portal, GSTR-3B FAQ, tutorial.gst.gov.in/userguide/returns/GSTR3B.htm, read 28 Nov 2026):
   3.1(a) taxable supplies incl. B2B reverse charge supplies (value only: the buyer pays the tax), B2C, deemed exports, notes, advances
   (11A) less advances adjusted (11B); 3.1(b) exports and SEZ supplies; 3.1(c) nil and exempt; 3.1(e) non-GST; 3.1(d) is drafted from
   GSTR-2B (reverse charge inward supplies), so it is taken from 2B, or from the purchase register's reverse charge rows when given
   (they also hold purchases from unregistered suppliers, which 2B never shows). Table 4: 2B rows with ITC available = 4(A)(5), ISD = 4(A)(4),
   reverse charge = 4(A)(3), "ITC not available" = 4(D)(2); credit blocked under s.17(5) is in 4(A) and reversed in 4(B)(1), ineligible
   credit (s.16(4), place of supply) only in 4(D)(2) (Circular 170/02/2022-GST). Amendment rows (B2BA, CDNRA) are left out with a note.
   Designed & Developed by: Nivesh Sahayak (https://www.niveshsahayak.in/) */
(function (G) {
  'use strict';
  var TP = G.NSTpl || tryReq('./tplcore.js'), TI = G.NSTallyIn || tryReq('./tallyin.js');
  function tryReq(p) { try { return typeof require === 'function' ? require(p) : null; } catch (e) { return null; } }
  function G2() { return G.NSG2 || tryReq('./g2core.js'); }
  function BK() { return G.NSBooks || tryReq('./books.js'); }
  function key(s) { return String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' '); }
  function t(s) { return String(s == null ? '' : s).trim(); }
  function r2(n) { return Math.round((+n + Number.EPSILON) * 100) / 100; }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function num(v) {
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    if (TP) return TP.amt(v).v;
    var s = t(v), neg = /^\(.*\)$/.test(s) || /^-/.test(s), m = s.replace(/,/g, '').match(/\d+(\.\d+)?|\.\d+/), n = m ? parseFloat(m[0]) : 0; return neg ? -n : n;
  }
  function date(v) {
    if (typeof v === 'number' && v > 20000 && v < 80000) return new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 864e5).toISOString().slice(0, 10);
    return TP ? TP.date(v) : (/^\d{4}-\d{2}-\d{2}/.test(t(v)) ? t(v).slice(0, 10) : '');
  }
  function g15(g) { return t(g).toUpperCase().replace(/\s+/g, ''); }
  var CH = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  function gstinCheck(g14) { var s = 0; for (var i = 0; i < 14; i++) { var v = CH.indexOf(g14[i]) * (i % 2 ? 2 : 1); s += Math.floor(v / 36) + v % 36; } return CH[(36 - s % 36) % 36]; }
  function gstinOk(g) { g = g15(g); return /^\d{2}[A-Z0-9]{13}$/.test(g) && gstinCheck(g.slice(0, 14)) === g[14]; }
  function stateOf(s, gstin) {
    var x = t(s); if (!x) return /^\d{2}/.test(g15(gstin)) ? g15(gstin).slice(0, 2) : '';
    var m = x.match(/^(\d{1,2})\s*[-–]/) || x.match(/^(\d{1,2})$/); if (m) return pad(+m[1]);
    return TP ? TP.stateCode(x.replace(/^\d+\s*-\s*/, ''), gstin) : '';
  }

  /* ---------- a sales or purchase register: columns by their words ---------- */
  var FIELDS = [
    { k: 'no', rules: [[/(invoice|bill|voucher|vch|doc|document|note|inv|supplier invoice|ref|reference)\.? ?(no|number|num|#)\b/, 9], [/^(no|bill|invoice|vch no|voucher)\.?$/, 5]] },
    { k: 'date', rules: [[/(invoice|bill|voucher|vch|doc|document|note) date|^date$/, 9], [/date/, 3]] },
    { k: 'party', rules: [[/party|customer|buyer|supplier|vendor|receiver|recipient name|trade name|legal name/, 8], [/particulars|^name$|^account( name)?$|^ledger( name)?$/, 6]] },
    { k: 'gstin', rules: [[/gstin|gst no|gst number|gst in\b|gstin\/uin/, 9]] },
    { k: 'pos', rules: [[/place of supply|^pos$|supply state|state of supply/, 9], [/^state$|state name|^state code$/, 4]] },
    { k: 'txv', rules: [[/taxable|assessable|value of supply|net value|basic (value|amount)/, 9]] },
    { k: 'rate', rules: [[/gst ?%|tax ?%|gst rate|tax rate|rate of tax|igst rate|^rate ?%|^rate$/, 7]] },
    { k: 'appl', rules: [[/applicable %/, 9]] },
    { k: 'igst', rules: [[/igst|integrated/, 9]] },
    { k: 'cgst', rules: [[/cgst|central/, 9]] },
    { k: 'sgst', rules: [[/sgst|utgst|state ?\/ ?ut tax|state tax|state\/ut/, 9]] },
    { k: 'cess', rules: [[/cess/, 9]] },
    { k: 'taxtot', rules: [[/total (tax|gst)|(tax|gst) amount|^(gst|tax)$/, 7]] },
    { k: 'value', rules: [[/(invoice|bill|note|voucher|document) (value|amount|total)|grand total|gross total|net amount|total amount/, 8], [/^(total|amount)$/, 4]] },
    { k: 'dkind', rules: [[/voucher type|vch type|document type|doc type|note type|bill type|^kind$|^type$/, 7]] },
    { k: 'cat', rules: [[/supply type|nature of supply|type of supply|taxability|tax type|gst type|category|invoice type|export type|note supply type|ur type/, 8]] },
    { k: 'rcm', rules: [[/reverse charge|\brcm\b|^rc$/, 9]] },
    { k: 'itc', rules: [[/itc|eligib|input (tax )?credit|17 ?\(?5|blocked/, 8]] }
  ];
  function heading(s) { return key(String(s == null ? '' : s).replace(/[₹*:()\[\]]/g, ' ').replace(/\brs\.?\b/gi, ' ')).replace(/\.$/, ''); }
  function fieldOf(h0) {
    var h = heading(h0); if (!h) return null;
    var best = null;
    FIELDS.forEach(function (F) {
      if (/%|\brate\b/.test(h) && /^(igst|cgst|sgst|cess|txv|taxtot|value)$/.test(F.k)) return;
      if (F.k === 'gstin' && /our gstin|company gstin|e-?commerce|\beco\b|operator/.test(h)) return;
      if (F.k === 'party' && /gstin|invoice|bill no|number|date|type/.test(h)) return;
      if (F.k === 'value' && /taxable|tax amount/.test(h)) return;
      if (F.k === 'date' && /due|shipping|port|reference date|return date/.test(h)) return;
      if (F.k === 'no' && /shipping|port|gstin|phone|mobile/.test(h)) return;
      F.rules.forEach(function (r) { if (r[0].test(h) && (!best || r[1] > best.w)) best = { k: F.k, w: r[1] }; });
    });
    return best;
  }
  function sig(heads) { return TP ? TP.sig(heads) : heads.map(function (h) { return key(h).replace(/[^a-z0-9%]+/g, ''); }).filter(Boolean).join('|'); }
  function mapOf(heads) {
    var map = {}, w = {};
    (heads || []).forEach(function (c, i) { var f = fieldOf(c); if (f && (!(f.k in map) || f.w > w[f.k])) { map[f.k] = i; w[f.k] = f.w; } });
    return map;
  }
  function good(map) { return ('date' in map) && ('txv' in map || 'value' in map || 'igst' in map || 'cgst' in map || 'taxtot' in map) ? Object.keys(map).length : 0; }
  function layout(rows, saved) {
    var best = null;
    for (var r = 0; r < Math.min(rows.length, 20); r++) {
      var heads = (rows[r] || []).map(function (c) { return c == null ? '' : String(c); }), s = sig(heads);
      if (saved && s && saved[s]) return { row: r, map: saved[s].map, sig: s, heads: heads, remembered: true };
      var cands = [{ heads: heads, row: r }];
      if (rows[r + 1]) cands.push({ heads: heads.map(function (h, i) { var b = (rows[r + 1] || [])[i]; return t(h) + ' ' + t(b == null ? '' : b); }), row: r + 1 });
      cands.forEach(function (c) { var map = mapOf(c.heads), n = good(map); if (n && (!best || n > best.n)) best = { row: c.row, map: map, sig: sig(heads), heads: c.heads, n: n }; });
    }
    return best;
  }
  /* what a "type" cell says: category of supply, and note / return */
  function catOf(s) {
    var x = key(s);
    if (!x) return '';
    if (/advance adj|adjusted/.test(x)) return 'advadj';
    if (/advance/.test(x)) return 'adv';
    if (/\bsez|sewp|sewop/.test(x)) return 'sez';
    if (/export|\bexp\b|expwp|expwop|wpay|wopay|\blut\b/.test(x)) return 'exp';
    if (/deemed|^de$/.test(x)) return 'de';
    if (/non[- ]?gst|no gst|out of scope|ngst|not under gst/.test(x)) return 'ngst';
    if (/\bnil/.test(x)) return 'nil';
    if (/exempt/.test(x)) return 'exm';
    if (/composition|^comp/.test(x)) return 'comp';
    if (/import|impg|bill of entry|\bboe\b/.test(x)) return 'imp';
    if (/\bisd\b|input service distributor/.test(x)) return 'isd';
    if (/taxable|regular|^r$|b2b|b2c|b2cs|b2cl|tax invoice/.test(x)) return 'tax';
    return '';
  }
  function kindOf(s, side) {
    var x = key(s);
    if (!x) return '';
    if (side === 's') { if (/credit note|sales? return|^cn$|^c$|return/.test(x)) return 'cn'; if (/debit note|^dn$|^d$/.test(x)) return 'dn'; return /invoice|bill|sale|tax/.test(x) ? 'inv' : ''; }
    if (/supplier'?s? debit|debit note received/.test(x)) return 'dn';
    if (/debit note|purchase return|return|credit note|^cn$|^dn$|^c$|^d$/.test(x)) return 'cn';
    return /invoice|bill|purchase/.test(x) ? 'inv' : '';
  }
  function itcOf(s) {
    var x = key(s); if (!x) return 'elig';
    if (/ineligib|not eligible|16 ?\(?4|place of supply|\bpos\b/.test(x)) return 'inel';
    if (/block|17 ?\(?5|personal|^no$|^n$|no itc|not available|^none$/.test(x)) return 'blk';
    return 'elig';
  }
  function yes(s) { return /^(y|yes|true|1|rcm|reverse)/i.test(t(s)); }
  var TOTAL = /^(grand )?total\b|^sub ?total|^closing|^opening|^total of/i;

  /* sheets [{ name, rows }] -> { rows, skipped: [{ sheet, row, why, txt }], read, sheets: [info], unknown: [{ name, heads, rows }] }
     side 's' = sales, 'p' = purchases; opts { own: our GSTIN (tax from a rate: IGST or CGST + SGST), layouts, force: { sheet: { row, map } } } */
  function register(sheets, side, opts) {
    opts = opts || {};
    var own = g15(opts.own).slice(0, 2), out = { rows: [], skipped: [], read: 0, sheets: [], unknown: [] };
    (sheets || []).forEach(function (sh) {
      if (key(sh.name) === 'how to fill') return;
      var rows = sh.rows || [], L = (opts.force && opts.force[sh.name]) || layout(rows, opts.layouts);
      if (!L || !good(L.map)) { if (rows.length > 1) out.unknown.push({ name: sh.name, heads: (rows[(L && L.row) || 0] || []).slice(0, 40), rows: rows }); return; }
      var M = L.map, info = { name: sh.name, row: L.row, map: M, heads: L.heads || rows[L.row] || [], sig: L.sig || sig(rows[L.row] || []), n: 0, remembered: !!L.remembered };
      out.sheets.push(info);
      var cell = function (r, f) { var c = f in M ? r[M[f]] : ''; return c == null ? '' : c; }, last = null;
      for (var i = L.row + 1; i < rows.length; i++) {
        var r = rows[i] || [], rn = i + 1;
        if (!r.some(function (c) { return t(c) !== ''; })) continue;
        out.read++;
        var no = t(cell(r, 'no')), dt0 = cell(r, 'date'), pty = t(cell(r, 'party')), where = sh.name + ' row ' + rn;
        var skip = function (why, kept) { out.skipped.push({ sheet: sh.name, row: rn, why: why, txt: [pty, no].filter(Boolean).join(' · '), kept: kept ? 1 : 0 }); };
        if (TOTAL.test(pty) || TOTAL.test(no) || TOTAL.test(t(r[0]))) { skip('total'); continue; }
        var head = !(no || t(dt0)) && last ? last : null;   /* a line of the bill above: item registers leave the bill's own columns empty */
        var o = head ? Object.assign({}, head, { txv: 0, igst: 0, cgst: 0, sgst: 0, cess: 0, rate: 0, src: where, line: 1 }) : {
          src: where, side: side, no: no, date: date(dt0), party: pty, gstin: g15(cell(r, 'gstin')), pos: '', kind: 'inv', cat: '', rcm: false, itc: 'elig' };
        if (!head) {
          if (!o.date) { skip(t(dt0) ? 'baddate' : 'nodate'); last = null; continue; }
          if (o.gstin && !gstinOk(o.gstin)) skip(/^\d{2}[A-Z0-9]{13}$/.test(o.gstin) ? 'chk' : 'badgstin', true);   /* kept: the tax still counts; only the GSTIN needs a look */
          o.pos = stateOf(cell(r, 'pos'), side === 's' ? o.gstin : '') || (side === 'p' && o.gstin ? o.gstin.slice(0, 2) : '');
          o.kind = kindOf(cell(r, 'dkind'), side) || kindOf(cell(r, 'cat'), side) || 'inv';
          o.cat = catOf(cell(r, 'cat')) || catOf(cell(r, 'dkind'));
          o.rcm = yes(cell(r, 'rcm')); o.itc = itcOf(cell(r, 'itc'));
        }
        var txv = num(cell(r, 'txv')), ig = num(cell(r, 'igst')), cg = num(cell(r, 'cgst')), sg = num(cell(r, 'sgst')), cs = num(cell(r, 'cess')), val = num(cell(r, 'value')), rt = num(cell(r, 'rate')), tt = num(cell(r, 'taxtot'));
        var ap = 'appl' in M ? num(cell(r, 'appl')) : 0;
        if (txv < 0 || val < 0 || ig + cg + sg < 0) { if (o.kind === 'inv' && !head) o.kind = 'cn'; }
        txv = Math.abs(txv); ig = Math.abs(ig); cg = Math.abs(cg); sg = Math.abs(sg); cs = Math.abs(cs); val = Math.abs(val); tt = Math.abs(tt);
        if (!txv && !ig && !cg && !sg && !tt && !val) { skip('noamt'); last = head ? last : null; continue; }
        o.rate = rt;
        if (!(ig || cg || sg)) {
          var tax = tt || (rt && txv ? r2(txv * rt / 100 * (ap ? ap / 100 : 1)) : 0);
          if (tax) {
            var there = side === 's' ? o.pos : (o.gstin ? o.gstin.slice(0, 2) : o.pos), inter = own && there ? own !== there : false;
            if (o.cat === 'exp' || o.cat === 'sez' || o.cat === 'imp') inter = true;
            if (inter) ig = tax; else { cg = r2(tax / 2); sg = r2(tax - cg); }
            if (!own || !there) o.guess = 1;
            o.derived = 1;
          }
        }
        if (!txv && val) txv = r2(val - ig - cg - sg - cs);
        o.txv = r2(txv); o.igst = r2(ig); o.cgst = r2(cg); o.sgst = r2(sg); o.cess = r2(cs); o.value = r2(val);
        if (!o.cat) o.cat = (ig || cg || sg || rt) ? 'tax' : 'nil0';
        out.rows.push(o); info.n++; if (!head) last = o;
      }
    });
    return out;
  }

  /* ---------- GSTR-1 as filed / prepared ---------- */
  function itmSum(list) {
    var o = { txv: 0, igst: 0, cgst: 0, sgst: 0, cess: 0, rate: 0 };
    (list || []).forEach(function (i) { var d = i.itm_det || i; o.txv += num(d.txval); o.igst += num(d.iamt); o.cgst += num(d.camt); o.sgst += num(d.samt); o.cess += num(d.csamt); if (num(d.rt)) o.rate = num(d.rt); });
    ['txv', 'igst', 'cgst', 'sgst', 'cess'].forEach(function (k) { o[k] = r2(o[k]); });
    return o;
  }
  function row1(src, o) { return Object.assign({ src: src, side: 's', no: '', date: '', party: '', gstin: '', pos: '', kind: 'inv', cat: 'tax', rcm: false, itc: 'elig', gstr1: 1 }, o); }
  /* the portal's GSTR-1 JSON (the one our GSTR-1 site and the department's offline tool make, or the one downloaded after filing) */
  function gstr1Json(j) {
    var out = { gstin: '', period: '', rows: [], skipped: [], read: 0, amend: 0 };
    if (!j || typeof j !== 'object') return out;
    out.gstin = g15(j.gstin); out.period = t(j.fp || j.ret_period);
    var push = function (o) { out.read++; out.rows.push(o); };
    (j.b2b || []).forEach(function (p) { (p.inv || []).forEach(function (v) {
      var it = itmSum(v.itms), ty = t(v.inv_typ).toUpperCase();
      push(row1('GSTR-1 B2B ' + t(v.inum), Object.assign(it, { no: t(v.inum), date: date(v.idt), gstin: g15(p.ctin), pos: pad(+v.pos || 0), rcm: /^y/i.test(t(v.rchrg)), cat: /^SEW/.test(ty) ? 'sez' : ty === 'DE' ? 'de' : 'tax', value: num(v.val) })));
    }); });
    (j.b2cl || []).forEach(function (p) { (p.inv || []).forEach(function (v) { push(row1('GSTR-1 B2CL ' + t(v.inum), Object.assign(itmSum(v.itms), { no: t(v.inum), date: date(v.idt), pos: pad(+p.pos || 0), value: num(v.val) }))); }); });
    (j.b2cs || []).forEach(function (s) { push(row1('GSTR-1 B2CS ' + (s.pos || '') + ' ' + (s.rt || '') + '%', { pos: pad(+s.pos || 0), txv: num(s.txval), igst: num(s.iamt), cgst: num(s.camt), sgst: num(s.samt), cess: num(s.csamt), rate: num(s.rt) })); });
    (j.exp || []).forEach(function (e) { (e.inv || []).forEach(function (v) { push(row1('GSTR-1 EXP ' + t(v.inum), Object.assign(itmSum(v.itms), { no: t(v.inum), date: date(v.idt), pos: '96', cat: 'exp' }))); }); });
    (j.cdnr || []).forEach(function (p) { (p.nt || []).forEach(function (v) {
      var ty = t(v.inv_typ).toUpperCase();
      push(row1('GSTR-1 CDNR ' + t(v.nt_num || v.ntnum), Object.assign(itmSum(v.itms), { no: t(v.nt_num || v.ntnum), date: date(v.nt_dt || v.ntdt), gstin: g15(p.ctin), pos: pad(+v.pos || 0), kind: /^D/i.test(t(v.ntty)) ? 'dn' : 'cn', rcm: /^y/i.test(t(v.rchrg)), cat: /^SEW/.test(ty) ? 'sez' : ty === 'DE' ? 'de' : 'tax' })));
    }); });
    (j.cdnur || []).forEach(function (v) {
      var ty = t(v.typ).toUpperCase();
      push(row1('GSTR-1 CDNUR ' + t(v.nt_num || v.ntnum), Object.assign(itmSum(v.itms), { no: t(v.nt_num || v.ntnum), date: date(v.nt_dt || v.ntdt), pos: ty === 'B2CL' ? pad(+v.pos || 0) : '96', kind: /^D/i.test(t(v.ntty)) ? 'dn' : 'cn', cat: /^EXP/.test(ty) ? 'exp' : 'tax' })));
    });
    var nil = j.nil && (j.nil.inv || j.nil);
    (Array.isArray(nil) ? nil : []).forEach(function (n) {
      var a = r2(num(n.nil_amt) + num(n.expt_amt)), g = num(n.ngsup_amt);
      if (a) push(row1('GSTR-1 nil / exempt ' + t(n.sply_ty), { txv: a, cat: 'nil' }));
      if (g) push(row1('GSTR-1 non-GST ' + t(n.sply_ty), { txv: g, cat: 'ngst' }));
    });
    (j.at || []).forEach(function (a) { var o = { txv: 0, igst: 0, cgst: 0, sgst: 0, cess: 0 }; (a.itms || []).forEach(function (i) { o.txv += num(i.ad_amt); o.igst += num(i.iamt); o.cgst += num(i.camt); o.sgst += num(i.samt); o.cess += num(i.csamt); }); push(row1('GSTR-1 advances (11A) ' + (a.pos || ''), Object.assign(o, { pos: pad(+a.pos || 0), cat: 'adv' }))); });
    (j.txpd || []).forEach(function (a) { var o = { txv: 0, igst: 0, cgst: 0, sgst: 0, cess: 0 }; (a.itms || []).forEach(function (i) { o.txv += num(i.ad_amt); o.igst += num(i.iamt); o.cgst += num(i.camt); o.sgst += num(i.samt); o.cess += num(i.csamt); }); push(row1('GSTR-1 advances adjusted (11B) ' + (a.pos || ''), Object.assign(o, { pos: pad(+a.pos || 0), cat: 'advadj' }))); });
    ['b2ba', 'b2cla', 'b2csa', 'expa', 'cdnra', 'cdnura', 'ata', 'txpda'].forEach(function (k) { var x = j[k]; if (Array.isArray(x) && x.length) out.amend += x.length; });
    out.rows.forEach(function (o) { ['txv', 'igst', 'cgst', 'sgst', 'cess'].forEach(function (k) { o[k] = r2(o[k] || 0); }); });
    return out;
  }
  /* the GST Returns Offline Tool workbook (sheets b2b,sez,de / b2cl / b2cs / cdnr / cdnur / exp / exemp / at / atadj; headings in
     row 4 of the department's file, row 1 of the CSV files our GSTR-1 site makes). The workbook holds rates, not tax amounts: the tax
     is worked out from the rate, Applicable % and place of supply (IGST when it is not our state). */
  var G1H = { gstin: /gstin\/uin of recipient/i, no: /^(invoice|note) number$/i, date: /^(invoice|note) date$/i, value: /^(invoice|note) value$/i, pos: /^place of supply$/i, rcm: /^reverse charge$/i,
    appl: /^applicable %/i, ityp: /^invoice type$/i, rate: /^rate$/i, txv: /^taxable value$/i, cess: /^cess amount$/i, typ: /^type$/i, nt: /^note type$/i, nst: /^note supply type$/i, ur: /^ur type$/i,
    etyp: /^export type$/i, nil: /^nil rated/i, exm: /^exempted/i, ngst: /^non-gst/i, adv: /^gross advance (received|adjusted)$/i };
  function sheetKind(n) {
    n = key(n).replace(/\s/g, '');
    if (/^b2b/.test(n) && !/b2ba/.test(n)) return 'b2b'; if (n === 'b2cl') return 'b2cl'; if (n === 'b2cs') return 'b2cs'; if (n === 'cdnr') return 'cdnr'; if (n === 'cdnur') return 'cdnur';
    if (n === 'exp') return 'exp'; if (n === 'exemp') return 'exemp'; if (n === 'at') return 'at'; if (n === 'atadj') return 'atadj';
    if (/^(b2ba|b2cla|b2csa|cdnra|cdnura|expa|ata|atadja)$/.test(n)) return 'amend';
    return '';
  }
  function gstr1Sheets(sheets, opts) {
    opts = opts || {};
    var own = g15(opts.own).slice(0, 2), out = { gstin: '', rows: [], skipped: [], read: 0, amend: 0, sheets: [] };
    (sheets || []).forEach(function (sh) {
      var kind = sheetKind(sh.name.replace(/^.*·\s*/, '')), rows = sh.rows || [];
      rows.slice(0, 4).forEach(function (r) { (r || []).forEach(function (c) { var m = String(c || '').match(/\b\d{2}[A-Z]{5}\d{4}[A-Z][A-Z0-9]Z[A-Z0-9]\b/); if (m && !out.gstin) out.gstin = m[0]; }); });
      if (!kind) return;
      var hr = -1, M = {};
      for (var i = 0; i < Math.min(8, rows.length) && hr < 0; i++) {
        var mm = {}; (rows[i] || []).forEach(function (c, ci) { var s = t(c); Object.keys(G1H).forEach(function (k) { if (!(k in mm) && G1H[k].test(s)) mm[k] = ci; }); });
        if (Object.keys(mm).length >= 2 && ('txv' in mm || 'nil' in mm || 'adv' in mm)) { hr = i; M = mm; }
      }
      if (hr < 0) return;
      if (kind === 'amend') { out.amend += rows.slice(hr + 1).filter(function (r) { return (r || []).some(function (c) { return t(c) !== ''; }); }).length; return; }
      out.sheets.push({ name: sh.name, kind: kind });
      var c = function (r, k) { var x = k in M ? r[M[k]] : ''; return x == null ? '' : x; };
      for (var rr = hr + 1; rr < rows.length; rr++) {
        var r = rows[rr] || [], rn = rr + 1, src = sh.name + ' row ' + rn;
        if (!r.some(function (x) { return t(x) !== ''; })) continue;
        out.read++;
        if (kind === 'exemp') {
          var d0 = t(c(r, 'typ')) || t(r[0]), a = r2(num(c(r, 'nil')) + num(c(r, 'exm'))), g = num(c(r, 'ngst'));
          if (a) out.rows.push(row1(src, { txv: a, cat: 'nil', party: d0 })); if (g) out.rows.push(row1(src, { txv: r2(g), cat: 'ngst', party: d0 }));
          continue;
        }
        var txv = num(kind === 'at' || kind === 'atadj' ? c(r, 'adv') : c(r, 'txv')), rt = num(c(r, 'rate')), ap = num(c(r, 'appl')), cs = num(c(r, 'cess'));
        if (!txv && !cs) { out.skipped.push({ sheet: sh.name, row: rn, why: 'noamt', txt: t(c(r, 'no')) }); continue; }
        var o = row1(src, { no: t(c(r, 'no')), date: date(c(r, 'date')), gstin: g15(c(r, 'gstin')), pos: stateOf(c(r, 'pos')), rcm: yes(c(r, 'rcm')), rate: rt, value: num(c(r, 'value')) });
        if (kind === 'exp') { o.cat = 'exp'; o.pos = '96'; }
        if (kind === 'b2b' || kind === 'cdnr') { var ty = key(c(r, kind === 'b2b' ? 'ityp' : 'nst')); o.cat = /sez/.test(ty) ? 'sez' : /deemed/.test(ty) ? 'de' : 'tax'; }
        if (kind === 'cdnur') { var u = key(c(r, 'ur')); o.cat = /exp/.test(u) ? 'exp' : 'tax'; if (o.cat === 'exp') o.pos = '96'; }
        if (kind === 'cdnr' || kind === 'cdnur') o.kind = /^d/i.test(t(c(r, 'nt'))) ? 'dn' : 'cn';
        if (kind === 'at') o.cat = 'adv'; if (kind === 'atadj') o.cat = 'advadj';
        if (kind === 'b2cs' && /^e$/i.test(t(c(r, 'typ')))) o.eco = 1;
        var withPay = !(o.cat === 'exp' && /wop/i.test(t(c(r, 'etyp')) + t(c(r, 'ur')))) && !(o.cat === 'sez' && /without/i.test(t(c(r, kind === 'b2b' ? 'ityp' : 'nst'))));
        var tax = withPay ? r2(txv * rt / 100 * (ap ? ap / 100 : 1)) : 0, inter = o.cat === 'exp' || o.cat === 'sez' || (own && o.pos && own !== o.pos);
        o.txv = r2(txv); o.cess = r2(cs);
        if (inter) { o.igst = tax; o.cgst = 0; o.sgst = 0; } else { o.igst = 0; o.cgst = r2(tax / 2); o.sgst = r2(tax - o.cgst); }
        if (!own && !inter) o.guess = 1;
        o.derived = 1;
        out.rows.push(o);
      }
    });
    return out;
  }

  /* ---------- Tally Day Book XML: sales and notes (and purchases, for the purchase side) ---------- */
  function tally(texts, opts) {
    opts = opts || {};
    var out = { rows: [], skipped: [], read: 0, company: '', purch: [] };
    if (!TI || !TI.plan) return out;
    var own = g15(opts.own).slice(0, 2), P = TI.plan(texts, { seller: { name: 'Tally file', state: own } }), B = BK();
    out.company = P.company || '';
    P.docs.forEach(function (d) {
      var sale = d.type === 'invoice' || d.type === 'note' && d.side !== 'purchase', buy = d.type === 'pbill' || d.type === 'note' && d.side === 'purchase';
      if (!sale && !buy) return;
      out.read++;
      var tt = B && B.itemTotals ? B.itemTotals(d) : null; if (!tt) return;
      var who = d.buyer || {}, o = { src: 'Tally ' + (d.no || ''), side: sale ? 's' : 'p', no: t(d.no), date: d.date, party: t(who.name), gstin: g15(who.gstin), kind: 'inv', cat: 'tax', rcm: false, itc: 'elig',
        txv: r2(tt.taxable), igst: r2(tt.igst), cgst: r2(tt.cgst), sgst: r2(tt.sgst), cess: 0 };
      if (o.gstin && !gstinOk(o.gstin)) out.skipped.push({ sheet: 'Tally', row: d.no, why: 'chk', txt: o.party, kept: 1 });
      if (sale) { o.pos = d.pos || who.state || own; if (d.type === 'note') o.kind = d.kind === 'debit' ? 'dn' : 'cn'; if (d.kind === 'supply') o.cat = 'nil0'; if (o.pos === '96') o.cat = 'exp'; }
      else { o.pos = who.state || (o.gstin ? o.gstin.slice(0, 2) : ''); if (d.type === 'note') o.kind = 'cn'; if (d.noGst) o.cat = 'nil0'; }
      if (!(o.igst + o.cgst + o.sgst) && o.cat === 'tax') o.cat = 'nil0';
      (sale ? out.rows : out.purch).push(o);
    });
    return out;
  }

  /* ---------- the GSTR-3B boxes ---------- */
  var BOX = { a: ['txv', 'igst', 'cgst', 'sgst', 'cess'], b: ['txv', 'igst', 'cess'], c: ['txv'], d: ['txv', 'igst', 'cgst', 'sgst', 'cess'], e: ['txv'],
    i1: ['igst', 'cess'], i2: ['igst', 'cess'], i3: ['igst', 'cgst', 'sgst', 'cess'], i4: ['igst', 'cgst', 'sgst', 'cess'], i5: ['igst', 'cgst', 'sgst', 'cess'],
    r1: ['igst', 'cgst', 'sgst', 'cess'], r2: ['igst', 'cgst', 'sgst', 'cess'], d1: ['igst', 'cgst', 'sgst', 'cess'], d2: ['igst', 'cgst', 'sgst', 'cess'],
    x1: ['inter', 'intra'], x2: ['inter', 'intra'], cf: ['igst', 'cgst', 'sgst', 'cess'], 'int': ['igst', 'cgst', 'sgst', 'cess'], fee: ['cgst', 'sgst'] };
  function zero() { var o = {}; Object.keys(BOX).forEach(function (b) { o[b] = {}; BOX[b].forEach(function (h) { o[b][h] = 0; }); }); return o; }
  var H4 = ['igst', 'cgst', 'sgst', 'cess'];
  /* D = { sales: rows, purch: rows, twoB: 2B rows (NSG2), amend1: n }, opts = { own (GSTIN), from, to (the return's days), itc: '2b' | 'reg',
     credit: { igst, cgst, sgst, cess } (credit ledger balance) }. Out: { A, s32, msgs: [{ lvl: 'err'|'warn'|'info', src, why, v }], n, sums } */
  function compute(D, opts) {
    opts = opts || {};
    var A = zero(), s32 = {}, msgs = [], own = g15(opts.own).slice(0, 2), from = opts.from || '', to = opts.to || '9999', n = { s: 0, p: 0, b: 0, out: 0 };
    function add(box, h, v) { if (A[box] && h in A[box]) A[box][h] = r2(A[box][h] + v); }
    function four(box, o, sg) { H4.forEach(function (h) { add(box, h, sg * (o[h] || 0)); }); }
    function msg(lvl, o, why, v) { msgs.push({ lvl: lvl, src: o.src || '', why: why, v: v || '', txt: [o.party, o.gstin, o.no].filter(Boolean).join(' · ') }); }
    function inP(o) { return o.date && o.date >= from && o.date <= to; }
    function check(o) {
      var tax = o.igst + o.cgst + o.sgst;
      if (o.igst && (o.cgst || o.sgst)) msg('warn', o, 'both');
      if (Math.abs(o.cgst - o.sgst) > 1) msg('warn', o, 'cgsgap');
      if (o.rate && !o.derived && o.txv && Math.abs(r2(o.txv * o.rate / 100) - tax) > 1 && tax) msg('warn', o, 'ratecheck', r2(o.txv * o.rate / 100));
      if (o.guess) msg('warn', o, 'guess');
    }
    /* outward */
    (D.sales || []).forEach(function (o) {
      if (!o.gstr1 && !inP(o)) { n.out++; return; }
      if (o.gstr1 && o.date && (o.date < from || o.date > to) && o.cat !== 'adv' && o.cat !== 'advadj' && o.kind === 'inv') msg('info', o, 'g1date');
      n.s++; check(o);
      var sg = o.kind === 'cn' ? -1 : 1, cat = o.cat;
      if (cat === 'nil0') { msg('warn', o, 'nil0'); cat = 'nil'; }
      if (cat === 'advadj') { sg = -1; cat = 'adv'; }
      if (cat === 'exp' || cat === 'sez') { add('b', 'txv', sg * o.txv); add('b', 'igst', sg * o.igst); add('b', 'cess', sg * o.cess); if (o.cgst || o.sgst) msg('warn', o, 'zrcgst'); return; }
      if (cat === 'nil' || cat === 'exm') { add('c', 'txv', sg * o.txv); return; }
      if (cat === 'ngst') { add('e', 'txv', sg * o.txv); return; }
      if (cat === 'comp' || cat === 'imp' || cat === 'isd') cat = 'tax';
      add('a', 'txv', sg * o.txv);
      if (o.rcm) { if (o.igst + o.cgst + o.sgst) msg('info', o, 'rcmsale'); return; }   /* the buyer pays the tax (s.9(3)): value only */
      four('a', o, sg);
      var inter = o.igst > 0 || (own && o.pos && o.pos !== own && !o.cgst);
      if (cat !== 'adv' && !o.gstin && inter && o.pos && o.pos !== '96') { var s = s32[o.pos] || (s32[o.pos] = { pos: o.pos, txv: 0, igst: 0 }); s.txv = r2(s.txv + sg * o.txv); s.igst = r2(s.igst + sg * o.igst); }
    });
    if (D.amend1) msgs.push({ lvl: 'warn', src: 'GSTR-1', why: 'amend1', v: D.amend1, txt: '' });
    /* inward: the purchase register (reverse charge, imports, nil / exempt purchases always; credit too when ITC is taken from it) */
    var fromReg = opts.itc === 'reg' || !(D.twoB && D.twoB.length), regBy = {}, regRcm = 0;
    (D.purch || []).forEach(function (o) {
      if (!inP(o)) { n.out++; return; }
      n.p++; check(o);
      var sg = o.kind === 'cn' ? -1 : 1, cat = o.cat === 'nil0' ? 'nil' : o.cat, inter = own && o.pos ? own !== o.pos : o.igst > 0;
      if (o.gstin && o.no) regBy[o.gstin + '|' + nk(o.no)] = o;
      if (cat === 'nil' || cat === 'exm' || cat === 'comp') { add('x1', inter ? 'inter' : 'intra', sg * o.txv); return; }
      if (cat === 'ngst') { add('x2', inter ? 'inter' : 'intra', sg * o.txv); return; }
      if (o.rcm) {
        regRcm++;
        add('d', 'txv', sg * o.txv); four('d', o, sg);
        if (o.itc === 'inel') four('d2', o, sg); else { four('i3', o, sg); if (o.itc === 'blk') four('r1', o, sg); }
        return;
      }
      if (cat === 'imp') { add('i1', 'igst', sg * o.igst); add('i1', 'cess', sg * o.cess); if (!fromReg) msg('info', o, 'impreg'); return; }
      if (!fromReg) return;
      if (o.itc === 'inel') { four('d2', o, sg); return; }
      four(cat === 'isd' ? 'i4' : 'i5', o, sg);
      if (o.itc === 'blk') four('r1', o, sg);
      if (!o.gstin) msg('warn', o, 'nogstinp');
    });
    /* inward: GSTR-2B */
    var used2b = {};
    if (!fromReg) {
      (D.twoB || []).forEach(function (x) {
        var o = { src: '2B ' + (x.sec || x.kind) + ' ' + x.no, party: x.name, gstin: x.gstin, no: x.no };
        if (x.sec === 'amend') { msg('info', o, 'amend2b'); return; }
        n.b++;
        var sg = x.kind === 'cn' ? -1 : 1, rg = regBy[x.gstin + '|' + nk(x.no)], itc = rg ? rg.itc : 'elig';
        if (rg) used2b[x.gstin + '|' + nk(x.no)] = 1;
        if (x.rev) {
          if (regRcm) return;   /* the purchase register's reverse charge rows already give 3.1(d) and 4(A)(3) */
          add('d', 'txv', sg * x.txv); four('d', x, sg);
          if (!x.itc || itc === 'inel') four('d2', x, sg); else { four('i3', x, sg); if (itc === 'blk') four('r1', x, sg); }
          return;
        }
        if (!x.itc || itc === 'inel') { four('d2', x, sg); if (!x.itc) msg('info', o, 'noitc2b', x.rsn || ''); return; }
        four(x.sec === 'isd' ? 'i4' : 'i5', x, sg);
        if (itc === 'blk') { four('r1', x, sg); msg('info', o, 'blk2b'); }
      });
      (D.purch || []).forEach(function (o) {
        if (!inP(o) || o.rcm || !o.gstin || o.cat === 'nil' || o.cat === 'nil0' || o.cat === 'exm' || o.cat === 'comp' || o.cat === 'ngst' || o.cat === 'imp') return;
        if (!used2b[o.gstin + '|' + nk(o.no)] && (o.igst + o.cgst + o.sgst)) msg('info', o, 'not2b', r2(o.igst + o.cgst + o.sgst));
      });
    }
    if (opts.credit) H4.forEach(function (h) { A.cf[h] = r2(Math.max(0, num(opts.credit[h]))); });
    var order = { err: 0, warn: 1, info: 2 };
    msgs.sort(function (a, b) { return order[a.lvl] - order[b.lvl]; });
    return { A: A, s32: Object.keys(s32).map(function (k) { return s32[k]; }).filter(function (s) { return s.txv || s.igst; }).sort(function (a, b) { return a.pos < b.pos ? -1 : 1; }), msgs: msgs, n: n, fromReg: fromReg };
  }
  function nk(no) { return String(no || '').toUpperCase().replace(/[^A-Z0-9]/g, '').replace(/^0+/, ''); }
  /* rows left out by the readers as messages (errors) */
  function skips(list, side) {
    return (list || []).filter(function (s) { return s.why !== 'total'; }).map(function (s) { return { lvl: s.kept ? 'warn' : 'err', src: s.sheet + ' row ' + s.row, why: s.why, txt: s.txt || '', side: side }; });
  }
  /* the period a GSTR-1 / 2B file is for (MMYYYY) as 'YYYY-MM' */
  function month(p) { p = t(p); var m = p.match(/^(\d{2})(\d{4})$/); return m ? m[2] + '-' + m[1] : ''; }

  /* ---------- sample files (made-up; GSTINs carry real check digits; names say Sample) ---------- */
  function gst(state, pan, ent) { var g = state + pan + ent + 'Z'; return g + gstinCheck(g); }
  function sample(m) {
    m = m || '2026-10';
    var me = gst('03', 'AAAPS1234A', '1'), d = function (day) { return pad(day) + '-' + m.slice(5, 7) + '-' + m.slice(0, 4); }, mmY = m.slice(5, 7) + m.slice(0, 4);
    var c1 = gst('03', 'AABCG5678B', '1'), c2 = gst('07', 'AACCD4321C', '1'), s1 = gst('03', 'AADFS2468D', '1'), s2 = gst('06', 'AAECH1357E', '1'), s3 = gst('03', 'AAFPT9753F', '1');
    var sales = [['Sales Register: Sample Traders (sample data)'], [],
      ['Invoice No', 'Date', 'Party Name', 'GSTIN/UIN', 'Place of Supply', 'Taxable Value', 'GST %', 'IGST', 'CGST', 'SGST', 'Invoice Value', 'Type'],
      ['ST/101', d(3), 'Sample Garg Hardware', c1, '03-Punjab', '1,00,000.00', '18', '', '9,000.00', '9,000.00', '1,18,000.00', 'Tax invoice'],
      ['ST/102', d(7), 'Sample Delhi Fittings', c2, '07-Delhi', '50,000.00', '18', '9,000.00', '', '', '59,000.00', 'Tax invoice'],
      ['ST/103', d(12), 'Cash sales (sample)', '', '03-Punjab', '20,000.00', '12', '', '1,200.00', '1,200.00', '22,400.00', 'Tax invoice'],
      ['ST/104', d(15), 'Sample walk-in buyer, Haryana', '', '06-Haryana', '10,000.00', '18', '1,800.00', '', '', '11,800.00', 'Tax invoice'],
      ['ST/105', d(18), 'Sample Gulf Importers LLC', '', 'Outside India', '80,000.00', '0', '', '', '', '80,000.00', 'Export under LUT'],
      ['ST/106', d(20), 'Sample Garg Hardware', c1, '03-Punjab', '5,000.00', '0', '', '', '', '5,000.00', 'Exempt'],
      ['CN/7', d(25), 'Sample Garg Hardware', c1, '03-Punjab', '-10,000.00', '18', '', '-900.00', '-900.00', '-11,800.00', 'Credit note'],
      ['ST/107', '31-13-2026', 'Sample bad date row', '', '03-Punjab', '1,000.00', '18', '', '90', '90', '1,180', 'Tax invoice'],
      ['', '', 'Grand Total', '', '', '2,55,000.00', '', '', '', '', '', '']];
    var purch = [['Purchase Register (sample)'],
      ['Supplier Invoice No', 'Invoice Date', 'Supplier', 'GSTIN', 'Taxable Value', 'IGST', 'CGST', 'SGST', 'Reverse charge', 'ITC eligibility', 'Supply type'],
      ['SS/55', d(4), 'Sample Steel Suppliers', s1, '60,000', '', '5,400', '5,400', 'No', 'Eligible', ''],
      ['HC/9', d(9), 'Sample Haryana Chemicals', s2, '40,000', '7,200', '', '', 'No', 'Eligible', ''],
      ['TP/3', d(11), 'Sample Transport Co.', s3, '10,000', '', '250', '250', 'Yes', 'Eligible', ''],
      ['CAR/1', d(14), 'Sample Car Dealer (motor car for office use)', s1, '6,00,000', '', '54,000', '54,000', 'No', 'Blocked 17(5)', ''],
      ['MLK/1', d(16), 'Sample Dairy (unregistered)', '', '3,000', '', '', '', 'No', '', 'Exempt'],
      ['SS/56', d(21), 'Sample Steel Suppliers', s1, '8,000', '', '720', '720', 'No', 'Eligible', '']];
    var two = { data: { gstin: me, rtnprd: mmY, docdata: {
      b2b: [{ ctin: s1, trdnm: 'Sample Steel Suppliers', supfildt: d(11), supprd: mmY, inv: [{ inum: 'SS/55', dt: d(4), val: 70800, typ: 'R', pos: '03', rev: 'N', itcavl: 'Y', rsn: '', items: [{ num: 1, rt: 18, txval: 60000, igst: 0, cgst: 5400, sgst: 5400, cess: 0 }] },
                                                                  { inum: 'CAR/1', dt: d(14), val: 708000, typ: 'R', pos: '03', rev: 'N', itcavl: 'Y', rsn: '', items: [{ num: 1, rt: 18, txval: 600000, igst: 0, cgst: 54000, sgst: 54000, cess: 0 }] }] },
            { ctin: s2, trdnm: 'Sample Haryana Chemicals', supfildt: d(11), supprd: mmY, inv: [{ inum: 'HC/9', dt: d(9), val: 47200, typ: 'R', pos: '03', rev: 'N', itcavl: 'Y', rsn: '', items: [{ num: 1, rt: 18, txval: 40000, igst: 7200, cgst: 0, sgst: 0, cess: 0 }] }] },
            { ctin: s3, trdnm: 'Sample Transport Co.', supfildt: d(11), supprd: mmY, inv: [{ inum: 'TP/3', dt: d(11), val: 10000, typ: 'R', pos: '03', rev: 'Y', itcavl: 'Y', rsn: '', items: [{ num: 1, rt: 5, txval: 10000, igst: 0, cgst: 250, sgst: 250, cess: 0 }] }] }],
      cdnr: [{ ctin: s2, trdnm: 'Sample Haryana Chemicals', supfildt: d(11), supprd: mmY, nt: [{ ntnum: 'HC/CN/2', dt: d(19), val: 2360, typ: 'C', pos: '03', rev: 'N', itcavl: 'Y', rsn: '', items: [{ num: 1, rt: 18, txval: 2000, igst: 360, cgst: 0, sgst: 0, cess: 0 }] }] }] } } };
    return { me: me, month: m, sales: sales, purch: purch, twoB: two };
  }

  var API = { register: register, gstr1Json: gstr1Json, gstr1Sheets: gstr1Sheets, tally: tally, compute: compute, skips: skips, zero: zero, layout: layout, fieldOf: fieldOf, FIELDS: FIELDS, sig: sig,
    month: month, sample: sample, gstinOk: gstinOk, catOf: catOf, itcOf: itcOf, G2: G2 };
  if (typeof module !== 'undefined' && module.exports) module.exports = API; else G.NSG3F = API;
})(typeof window !== 'undefined' ? window : this);
