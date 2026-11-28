/* NSE3 (byof3-einv-gstr1, 28 Nov 2026): three-way match before filing GSTR-1. Pure (no screen); e3match.js draws it, node tests use it.
   Nothing here sends anything anywhere: every file stays in the browser.

   Three sources, each read into the same record shape
     { src: 'e'|'g'|'b', no, nk (number key), kind: 'inv'|'cn'|'dn', date (YYYY-MM-DD), gstin, party, pos, txv, igst, cgst, sgst, cess, val,
       irn, cnl (cancelled), einv (an e-invoice is expected: registered buyer, export or SEZ), at (where it came from) }
   - e-invoices: any JSON the IRP or the GST portal gives (Get IRN details reply with SignedInvoice, the INV-01 upload file, objects with
     SignedQRCode / Irn / DocNo; lists of any of these, also wrapped in data / Data), and any Excel / CSV list of IRNs (the GST portal's
     "Download details from e-invoices (Excel)" on the GSTR-1 page, the e-invoice portal's IRN reports, a sheet of your own).
     Official sources read 28 Nov 2026: tutorial.gst.gov.in/userguide/returns/Creation_of_Outward_Supplies_Return_in_GSTR-1.htm (the
     GSTR-1 page has "DOWNLOAD DETAILS FROM E-INVOICES (EXCEL)", tables 4A/4B/4C/6B/6C, 6A, 9B CDNR / CDNUR, one sheet per table; IRN and
     IRN date columns; over 500 records the file goes to "E-invoice download history"). GSTN publishes no column list of that Excel,
     so columns are found by their words and the person can say which is which (Columns screen); the INV-01 field names are the same
     ones einv.js writes (DocDtls, BuyerDtls, ValDtls, TranDtls).
   - GSTR-1: the portal / offline tool JSON or the GST Returns Offline Tool workbook, read by NSG3F.gstr1Json / gstr1Sheets (g3core.js).
   - books: a sales register (Tally, Busy, Marg, Excel, CSV) or Tally Day Book XML, read by NSG1File (g1file.js).
   compare(E, G, B, opts) -> { rows, probs, counts, scope, skipped }: one row per invoice / note number, which source has it, and what differs
   (value, tax, tax type, GSTIN, date, invoice value, cancelled e-invoice still reported). sheets(res) -> Excel sheets for NS.xlsx.write.
   Designed & Developed by: Nivesh Sahayak (https://www.niveshsahayak.in/) */
(function (G) {
  'use strict';
  function tryReq(p) { try { return typeof require === 'function' ? require(p) : null; } catch (e) { return null; } }
  var TP = G.NSTpl || tryReq('./tplcore.js');
  function G3() { return G.NSG3F || tryReq('./g3core.js'); }
  function G1() { return G.NSG1File || tryReq('./g1file.js'); }
  function t(s) { return String(s == null ? '' : s).trim(); }
  function key(s) { return t(s).toLowerCase().replace(/\s+/g, ' '); }
  function r2(n) { return Math.round((+n || 0) * 100 + (n < 0 ? -1e-9 : 1e-9)) / 100; }
  function num(v) { if (typeof v === 'number') return isFinite(v) ? v : 0; return TP ? TP.amt(v).v : (parseFloat(t(v).replace(/,/g, '')) || 0); }
  function g15(g) { return t(g).toUpperCase().replace(/\s+/g, ''); }
  var CH = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  function gstinOk(g) {
    g = g15(g); if (!/^[0-9]{2}[0-9A-Z]{13}$/.test(g)) return false;
    var s = 0; for (var i = 0; i < 14; i++) { var v = CH.indexOf(g[i]) * (i % 2 ? 2 : 1); s += Math.floor(v / 36) + v % 36; }
    return CH[(36 - s % 36) % 36] === g[14];
  }
  function date(v) {
    if (typeof v === 'number' && v > 20000 && v < 80000) return new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 864e5).toISOString().slice(0, 10);
    var s = t(v); if (/^\d{5}(\.\d+)?$/.test(s)) return date(+s);
    var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return m[0];
    return TP ? TP.date(s) : '';
  }
  /* invoice numbers: the key ignores case and spaces; the loose key also ignores signs and leading zeros (INV/001 = inv-1) */
  function nk(no) { return t(no).toUpperCase().replace(/\s+/g, ''); }
  function loose(no) { return nk(no).replace(/[^A-Z0-9]+/g, '-').replace(/(^|-|[A-Z])0+(?=\d)/g, '$1').replace(/^-|-$/g, ''); }
  function kindOf(s) { var k = key(s); if (/^(crn|cn|c)$|credit/.test(k)) return 'cn'; if (/^(dbn|dn|d)$|debit/.test(k)) return 'dn'; return 'inv'; }
  function rec(src, o) {
    var x = Object.assign({ src: src, no: '', kind: 'inv', date: '', gstin: '', party: '', pos: '', txv: 0, igst: 0, cgst: 0, sgst: 0, cess: 0, val: 0, irn: '', cnl: false, cat: '', at: '' }, o);
    x.no = t(x.no); x.nk = nk(x.no); x.gstin = g15(x.gstin);
    ['txv', 'igst', 'cgst', 'sgst', 'cess', 'val'].forEach(function (k) { x[k] = r2(Math.abs(num(x[k]))); });
    if (!x.val) x.val = r2(x.txv + x.igst + x.cgst + x.sgst + x.cess);
    x.einv = x.gstin.length === 15 || x.cat === 'exp' || x.cat === 'sez' || x.cat === 'de' || x.pos === '96';
    return x;
  }

  /* ---------- e-invoices: JSON ---------- */
  function b64(part) {
    var s = String(part).replace(/-/g, '+').replace(/_/g, '/'); s += '==='.slice((s.length + 3) % 4);
    if (typeof atob === 'function') { var b = atob(s), u = new Uint8Array(b.length); for (var i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); return new TextDecoder().decode(u); }
    return Buffer.from(s, 'base64').toString('utf8');
  }
  /* a signed JWT (SignedInvoice / SignedQRCode): its middle part; "data" holds the invoice as a JSON string */
  function jwt(s) {
    var p = t(s).split('.'); if (p.length !== 3) return null;
    try { var j = JSON.parse(b64(p[1])); if (typeof j.data === 'string') { try { return JSON.parse(j.data); } catch (e) { return j; } } return j.data && typeof j.data === 'object' ? j.data : j; } catch (e) { return null; }
  }
  function lower(o) { var k = {}; Object.keys(o).forEach(function (x) { k[x.toLowerCase().replace(/[\s_]/g, '')] = o[x]; }); return k; }
  function fromInv(v, extra, at) {
    var D = v.DocDtls || {}, B = v.BuyerDtls || {}, V = v.ValDtls || {}, T = v.TranDtls || {}, sup = t(T.SupTyp).toUpperCase();
    return rec('e', Object.assign({ no: D.No, date: date(D.Dt), kind: kindOf(D.Typ), gstin: B.Gstin === 'URP' ? '' : B.Gstin, party: B.LglNm || B.TrdNm, pos: t(B.Pos).padStart(2, '0'),
      txv: V.AssVal, igst: V.IgstVal, cgst: V.CgstVal, sgst: V.SgstVal, cess: num(V.CesVal) + num(V.StCesVal), val: V.TotInvVal,
      cat: /^EXP/.test(sup) ? 'exp' : /^SEZ/.test(sup) ? 'sez' : sup === 'DEXP' ? 'de' : '', at: at }, extra));
  }
  function readJson(j, name) {
    var out = { recs: [], skipped: [], files: 1 }, seen = 0;
    function walk(o, at) {
      if (Array.isArray(o)) { o.forEach(function (x, i) { walk(x, at + '[' + (i + 1) + ']'); }); return; }
      if (!o || typeof o !== 'object') return;
      var k = lower(o), st = t(k.status || k.irnstatus || k.einvoicestatus), cnl = /^(cnl|cancel)/i.test(st) || !!k.cncldt || !!k.canceldate;
      var extra = { irn: t(k.irn), cnl: cnl };
      if (typeof k.signedinvoice === 'string') { var inv = jwt(k.signedinvoice); if (inv && inv.DocDtls) { seen++; if (!extra.irn) extra.irn = t(inv.Irn); out.recs.push(fromInv(inv, extra, at)); return; } }
      if (o.DocDtls && (o.ValDtls || o.ItemList)) { seen++; out.recs.push(fromInv(o, extra, at)); return; }
      if (typeof k.data === 'string' && /^\s*[\[{]/.test(k.data)) { try { walk(JSON.parse(k.data), at); return; } catch (e) {} }
      if (typeof k.signedqrcode === 'string') {
        var q = jwt(k.signedqrcode);
        if (q && q.DocNo) { seen++; out.recs.push(rec('e', { no: q.DocNo, date: date(q.DocDt), kind: kindOf(q.DocTyp), gstin: q.BuyerGstin === 'URP' ? '' : q.BuyerGstin, val: q.TotInvVal, irn: t(q.Irn || k.irn), cnl: cnl, at: at, onlyVal: true })); return; }
      }
      var no = k.docno || k.invoiceno || k.documentno;
      if (no && (k.irn || k.ackno)) { seen++; out.recs.push(rec('e', { no: no, date: date(k.docdt || k.invoicedate || k.docdate), kind: kindOf(k.doctyp || k.doctype), gstin: k.buyergstin || k.ctin, val: k.totinvval || k.invoicevalue, txv: k.assval, igst: k.igstval, cgst: k.cgstval, sgst: k.sgstval, irn: t(k.irn), cnl: cnl, at: at, onlyVal: !k.assval })); return; }
      Object.keys(o).forEach(function (x) { if (o[x] && typeof o[x] === 'object') walk(o[x], at); });
    }
    walk(j, name || 'JSON');
    if (!seen) out.skipped.push({ at: name || 'JSON', why: 'nojson' });
    return out;
  }

  /* ---------- e-invoices: Excel / CSV list (one row per invoice, or one row per rate like the GSTR-1 tables) ---------- */
  var ESET = { k: 'e3einv', name: 'E-invoice list', cols: [
    { k: 'no', h: 'Document no', need: 1, a: ['invoice number', 'invoice no', 'invoice no.', 'document number', 'document no', 'doc no', 'doc number', 'docno', 'note number', 'note no', 'bill no', 'inv no', 'supplier invoice no', 'document no.', 'invoice/note number'] },
    { k: 'date', h: 'Document date', need: 1, a: ['invoice date', 'document date', 'doc date', 'docdt', 'note date', 'bill date', 'inv date', 'invoice/note date'] },
    { k: 'typ', h: 'Document type', a: ['doc type', 'document type', 'doctyp', 'note type', 'type of document', 'type'] },
    { k: 'gstin', h: 'Buyer GSTIN', a: ['gstin/uin of recipient', 'gstin of recipient', 'recipient gstin', 'buyer gstin', 'receiver gstin', 'customer gstin', 'gstin', 'gstin/uin', 'ctin', 'buyer gstin/uin'] },
    { k: 'party', h: 'Buyer name', a: ['receiver name', 'recipient name', 'buyer name', 'customer name', 'name of recipient', 'legal name', 'trade name', 'party name', 'buyer legal name'] },
    { k: 'pos', h: 'Place of supply', a: ['pos', 'place of supply', 'buyer state', 'state of supply'] },
    { k: 'txv', h: 'Taxable value', a: ['taxable value', 'assessable value', 'assval', 'taxable amount', 'total assessable value', 'taxable value (rs)', 'total taxable value', 'taxable amt', 'taxable', 'basic amount', 'assessable amount'] },
    { k: 'rate', h: 'Rate', a: ['rate', 'gst rate', 'tax rate', 'rate (%)', 'gst %'] },
    { k: 'igst', h: 'IGST', a: ['integrated tax', 'igst', 'igst amount', 'igst value', 'igstval', 'integrated tax (rs)', 'integrated tax amount'] },
    { k: 'cgst', h: 'CGST', a: ['central tax', 'cgst', 'cgst amount', 'cgst value', 'cgstval', 'central tax (rs)', 'central tax amount'] },
    { k: 'sgst', h: 'SGST / UTGST', a: ['state/ut tax', 'state tax', 'sgst', 'sgst/utgst', 'sgst amount', 'sgst value', 'sgstval', 'state/ut tax (rs)', 'state/ut tax amount', 'utgst'] },
    { k: 'cess', h: 'Cess', a: ['cess', 'cess amount', 'cess value', 'cesval', 'cess (rs)'] },
    { k: 'val', h: 'Invoice value', a: ['invoice value', 'document value', 'total invoice value', 'totinvval', 'note value', 'invoice value (rs)', 'total value', 'note value (rs)'] },
    { k: 'irn', h: 'IRN', a: ['irn', 'irn no', 'irn number', 'invoice reference number', 'invoice reference number (irn)'] },
    { k: 'irndt', h: 'IRN date', a: ['irn date', 'ack date', 'ackdt', 'acknowledgement date', 'irn generation date'] },
    { k: 'ack', h: 'Ack no', a: ['ack no', 'ack no.', 'ackno', 'acknowledgement no', 'acknowledgement number'] },
    { k: 'status', h: 'Status', a: ['status', 'irn status', 'e-invoice status', 'einvoice status', 'e invoice status', 'document status'] },
    { k: 'sup', h: 'Supply type', a: ['supply type', 'suptyp', 'invoice type', 'transaction type', 'category'] }
  ] };
  if (TP && TP.addSet) TP.addSet(ESET);
  function sheetNote(name) { var n = key(name).replace(/^.*·\s*/, '').replace(/\s/g, ''); return /^cdnr|^cdnur|credit|debit|note/.test(n); }
  function readSheets(book, opts) {
    opts = opts || {};
    var out = { recs: [], skipped: [], found: [] }, own = g15(opts.own).slice(0, 2);
    (book || []).forEach(function (sh) {
      var rows = sh.rows || [], L = (opts.force && opts.force[sh.name]) || TP.match('e3einv', sh.name, rows, opts.saved);
      if (!L || !('no' in L.map) || !('date' in L.map)) { if (rows.length > 1) out.found.push({ name: sh.name, ok: false, rows: rows.length, headRow: (L && L.row) || 0, map: L ? L.map : {} }); return; }
      var objs = TP.objects(rows, L), by = {}, list = [], noteSheet = sheetNote(sh.name);
      objs.forEach(function (o) {
        if (!t(o.no)) { if ((o._r || 0) > L.row + 1 && Object.keys(o).some(function (k) { return k[0] !== '_' && t(o[k]); })) out.skipped.push({ at: sh.name + ' row ' + o._r, why: 'nono' }); return; }
        var d = date(o.date); if (!d) { out.skipped.push({ at: sh.name + ' row ' + o._r, why: 'nodate', v: t(o.date) }); return; }
        var kind = t(o.typ) ? kindOf(o.typ) : noteSheet ? 'cn' : 'inv', g = g15(o.gstin), k = kind + '|' + nk(o.no) + '|' + g;
        var hasTax = ['igst', 'cgst', 'sgst'].some(function (f) { return t(o[f]) !== ''; });
        var x = by[k];
        if (!x) { x = by[k] = { no: t(o.no), date: d, kind: kind, gstin: /^(urp|na|-)$/i.test(g) ? '' : g, party: t(o.party), pos: posOf(o.pos, g), txv: 0, igst: 0, cgst: 0, sgst: 0, cess: 0, val: 0, irn: t(o.irn), cnl: /cancel|cnl/i.test(t(o.status)), cat: /exp/i.test(t(o.sup)) ? 'exp' : /sez/i.test(t(o.sup)) ? 'sez' : '', at: sh.name + ' row ' + o._r, _rows: 0 }; list.push(x); }
        x._rows++;
        var tv = num(o.txv), rt = num(String(o.rate || '').replace('%', ''));
        x.txv += tv; x.cess += num(o.cess); if (!x.val) x.val = num(o.val); if (!x.irn) x.irn = t(o.irn);
        if (hasTax) { x.igst += num(o.igst); x.cgst += num(o.cgst); x.sgst += num(o.sgst); }
        else if (rt && tv) {   /* only a rate given: tax worked out from it, IGST when the place of supply is not our state */
          var tax = r2(tv * rt / 100), inter = x.cat === 'exp' || x.cat === 'sez' || (own && x.pos && x.pos !== own);
          if (inter) x.igst += tax; else { x.cgst += r2(tax / 2); x.sgst += r2(tax - r2(tax / 2)); }
          x.derived = 1;
        }
      });
      list.forEach(function (x) { delete x._rows; out.recs.push(rec('e', x)); });
      out.found.push({ name: sh.name, ok: true, rows: list.length, headRow: L.row, map: L.map, sig: L.sig, heads: L.heads || rows[L.row], remembered: !!L.remembered });
    });
    return out;
  }
  function posOf(s, g) {
    var x = t(s), m = x.match(/^(\d{1,2})\b/);
    if (m) return ('0' + m[1]).slice(-2);
    if (x && TP && TP.stateCode) { var c = TP.stateCode(x.replace(/^\d+\s*-\s*/, ''), ''); if (c) return c; }
    return /^\d{2}/.test(g) ? g.slice(0, 2) : '';
  }
  /* any e-invoice input: { json: [{ name, data }], book: [sheets] } */
  function readEinv(inp, opts) {
    var out = { recs: [], skipped: [], found: [] };
    (inp.json || []).forEach(function (f) { var r = readJson(f.data, f.name); out.recs = out.recs.concat(r.recs); out.skipped = out.skipped.concat(r.skipped); out.found.push({ name: f.name, ok: r.recs.length > 0, rows: r.recs.length, json: 1 }); });
    if ((inp.book || []).length) { var s = readSheets(inp.book, opts); out.recs = out.recs.concat(s.recs); out.skipped = out.skipped.concat(s.skipped); out.found = out.found.concat(s.found); }
    /* the same IRN twice (two downloads that overlap): keep one */
    var seen = {}; out.recs = out.recs.filter(function (r) { var k = r.irn || (r.kind + '|' + r.nk + '|' + r.date + '|' + r.gstin); if (seen[k]) { if (r.cnl) seen[k].cnl = true; return false; } seen[k] = r; return true; });
    return out;
  }

  /* ---------- GSTR-1 (NSG3F readers) ---------- */
  function readG1(inp, opts) {
    var F = G3(), out = { recs: [], skipped: [], found: [], summary: [], gstin: '', amend: 0 };
    if (!F) { out.skipped.push({ at: 'GSTR-1', why: 'noreader' }); return out; }
    function take(r, name) {
      out.gstin = out.gstin || r.gstin; out.amend += r.amend || 0;
      var n = 0;
      (r.rows || []).forEach(function (o) {
        if (!t(o.no)) { if (o.txv || o.igst || o.cgst || o.sgst) out.summary.push({ at: o.src, cat: o.cat, txv: o.txv, tax: r2(o.igst + o.cgst + o.sgst + o.cess) }); return; }
        n++; out.recs.push(rec('g', { no: o.no, date: o.date, kind: o.kind, gstin: o.gstin, pos: o.pos, txv: o.txv, igst: o.igst, cgst: o.cgst, sgst: o.sgst, cess: o.cess, val: o.value, cat: o.cat, at: o.src, derived: !!o.derived, b2cl: /B2CL/.test(o.src) }));
      });
      (r.skipped || []).forEach(function (s) { out.skipped.push({ at: (s.sheet || name) + (s.row ? ' row ' + s.row : ''), why: s.why || 'skip' }); });
      out.found.push({ name: name, ok: n > 0 || out.summary.length > 0, rows: n });
    }
    (inp.json || []).forEach(function (f) { take(F.gstr1Json(f.data), f.name); });
    if ((inp.book || []).length) take(F.gstr1Sheets(inp.book, { own: opts && opts.own }), 'GSTR-1 workbook');
    return out;
  }

  /* ---------- books (NSG1File: sales register sheets or Tally XML) ---------- */
  function docTot(d, own) {
    var o = { txv: 0, igst: 0, cgst: 0, sgst: 0, cess: 0 };
    (d.items || []).forEach(function (it) {
      if (it._tx) { ['txv', 'igst', 'cgst', 'sgst', 'cess'].forEach(function (k) { o[k] += +it._tx[k] || 0; }); return; }
      var v = r2(num(it.qty || 1) * num(it.rate) * (1 - num(it.disc) / 100)), tx = r2(v * num(it.gst) / 100), intra = !d.pos || !own || d.pos === own;
      o.txv += v; if (intra) { o.cgst += r2(tx / 2); o.sgst += r2(tx - r2(tx / 2)); } else o.igst += tx;
    });
    return o;
  }
  function readBooks(inp, opts) {
    opts = opts || {};
    var F = G1(), out = { recs: [], skipped: [], found: [], lines: null };
    if (!F) { out.skipped.push({ at: 'Books', why: 'noreader' }); return out; }
    var own = g15(opts.own).slice(0, 2), docs = [];
    if ((inp.book || []).length) {
      var L = F.read(inp.book, { saved: opts.saved, force: opts.force }); out.lines = L;
      var T = F.toDocs(L.lines, { own: own }); docs = docs.concat(T.docs);
      T.probs.forEach(function (p) { out.skipped.push({ at: (p.at || []).join(' row '), why: p.k, v: p.v && p.v.d }); });
      L.found.forEach(function (f) { out.found.push({ name: f.name, ok: !!f.kind, rows: f.rows, headRow: f.headRow, map: f.map, sig: f.sig, heads: f.heads, remembered: f.remembered }); });
    }
    if ((inp.texts || []).length) {
      try { var X = F.fromTally(inp.texts, { own: own }); docs = docs.concat(X.docs); if (!own) own = X.own || ''; out.found.push({ name: 'Tally XML', ok: X.docs.length > 0, rows: X.docs.length }); }
      catch (e) { out.skipped.push({ at: 'Tally XML', why: 'xml', v: e && e.message }); }
    }
    if (!own && F.guessOwn) own = F.guessOwn(docs);
    docs.forEach(function (d) {
      var tt = docTot(d, own), sup = d.einv && d.einv.sup || '';
      out.recs.push(rec('b', { no: d.no, date: d.date, kind: d.type === 'note' ? (d.kind === 'debit' ? 'dn' : 'cn') : 'inv', gstin: d.buyer && d.buyer.gstin, party: d.buyer && d.buyer.name, pos: d.pos,
        txv: tt.txv, igst: tt.igst, cgst: tt.cgst, sgst: tt.sgst, cess: tt.cess, val: d._val, cnl: !!d.cancelled, cat: /^EXP/.test(sup) || d.pos === '96' ? 'exp' : /^SEZ/.test(sup) ? 'sez' : sup === 'DEXP' ? 'de' : '',
        at: (d._rows && d._rows[0] || []).join(' row ') }));
    });
    out.own = own;
    return out;
  }

  /* ---------- the match ---------- */
  /* B2C Large (table 5): unregistered buyer in another state, bill above Rs 1 lakh from the August 2024 period (Notification 12/2024-CT), Rs 2.5 lakh before */
  function b2cl(r, own) { own = g15(own).slice(0, 2); return !r.gstin && r.kind === 'inv' && own && r.pos && r.pos !== own && r.pos !== '96' && r.val > (r.date >= '2024-08-01' ? 100000 : 250000); }
  var TOL = 1;   /* rupees: rounding between software */
  function mon(d) { return String(d || '').slice(0, 7); }
  function months(recs) { var m = {}; recs.forEach(function (r) { if (r.date) m[mon(r.date)] = 1; }); return Object.keys(m).sort(); }
  function tax(r) { return r2(r.igst + r.cgst + r.sgst + r.cess); }
  /* E, G, B: record lists, or null when that file was not given. opts: { own } */
  function compare(E, G, B, opts) {
    opts = opts || {};
    var have = { e: !!E, g: !!G, b: !!B }, srcs = ['e', 'g', 'b'].filter(function (s) { return have[s]; }), L = { e: E || [], g: G || [], b: B || [] };
    /* compare only the months every given file covers: a GSTR-1 of October against a year of books is not "missing" anything else */
    var common = null;
    srcs.forEach(function (s) { var m = months(L[s]); common = common == null ? m : common.filter(function (x) { return m.indexOf(x) >= 0; }); });
    common = common || [];
    var inScope = function (r) { return common.indexOf(mon(r.date)) >= 0; };
    var out = { rows: [], probs: [], counts: { ok: 0, diff: 0, miss: 0, cnl: 0 }, scope: common, outside: { e: 0, g: 0, b: 0 }, have: have, dups: [] };
    /* groups: one per document. Records with the same loose number go in one bucket; inside it each record joins the group that fits it
       best (same exact number, same buyer GSTIN, same date) and has nothing yet from its own file. Books first, so they anchor the groups. */
    var buckets = {}, order = [];
    ['b', 'g', 'e'].filter(function (s) { return have[s]; }).forEach(function (s) {
      var dupAt = {};
      L[s].forEach(function (r) {
        if (!inScope(r)) { out.outside[s]++; return; }
        var dk = r.kind + '|' + r.nk + '|' + r.gstin; if (dupAt[dk] && !r.cnl) out.dups.push({ src: s, no: r.no, at: r.at, other: dupAt[dk] }); else dupAt[dk] = r.at;
        var bk = r.kind + '|' + loose(r.no), B = buckets[bk] || (buckets[bk] = []), best = null, bs = -1;
        B.forEach(function (g) {
          if (g[s]) return;
          var o = g.b || g.g || g.e, sc = (o.nk === r.nk ? 2 : 0) + (o.gstin && o.gstin === r.gstin ? 3 : 0) + (o.date === r.date ? 1 : 0) - (o.gstin && r.gstin && o.gstin !== r.gstin ? 1 : 0);
          if (sc > bs) { bs = sc; best = g; }
        });
        if (!best) { best = {}; B.push(best); order.push(best); }
        best[s] = r;
      });
    });
    order.forEach(function (g) { var ks = ['e', 'g', 'b'].filter(function (s) { return g[s]; }); g.looseNo = ks.some(function (s) { return g[s].nk !== g[ks[0]].nk; }); });
    var NAME = { e: 'e-invoice', g: 'GSTR-1', b: 'books' };
    order.forEach(function (g) {
      var any = g.e || g.g || g.b, row = { no: any.no, kind: any.kind, date: any.date, gstin: any.gstin, party: (g.b && g.b.party) || (g.e && g.e.party) || '', e: g.e || null, g: g.g || null, b: g.b || null, what: [], status: 'ok' };
      function add(code, v, sev) { row.what.push({ k: code, v: v || {}, sev: sev || 'diff' }); }
      var eCnl = g.e && g.e.cnl, needE = (g.g && g.g.einv && !g.g.cnl) || (g.b && g.b.einv && !g.b.cnl);
      if (have.e && !g.e && needE) add('miss_e', {}, 'miss');
      if (eCnl) {
        if (g.g) add('cnl_g', {}, 'diff');
        if (g.b && !g.b.cnl) add('cnl_b', {}, 'diff');
        if (!g.g && (!g.b || g.b.cnl)) { row.status = 'cnl'; out.counts.cnl++; out.rows.push(row); return; }
      }
      var listed = (g.b && (g.b.einv || b2cl(g.b, opts.own))) || (g.e && !g.e.cnl);
      if (have.g && !g.g && !eCnl && !(g.b && g.b.cnl) && listed) add('miss_g', {}, 'miss');
      if (have.g && !g.g && g.b && !listed && !g.e) add('b2c', {}, 'info');
      if (have.b && !g.b && !eCnl) add('miss_b', {}, 'miss');
      if (g.b && g.b.cnl && g.g) add('bcnl_g', {}, 'diff');
      if (g.looseNo) add('looseno', { n: ['e', 'g', 'b'].filter(function (s) { return g[s]; }).map(function (s) { return NAME[s] + ' "' + g[s].no + '"'; }).join(', ') }, 'info');
      /* pairwise differences, against the books first (they are what you keep), then e-invoice against GSTR-1 */
      var live = function (s) { return g[s] && !g[s].cnl; };
      var pairs = live('b') ? [['b', 'e'], ['b', 'g']] : [['e', 'g']];
      pairs = pairs.filter(function (p) { return live(p[0]) && live(p[1]); });
      var said = {};
      pairs.forEach(function (p) {
        var a = g[p[0]], b = g[p[1]], w = { a: NAME[p[0]], b: NAME[p[1]] };
        function once(code, v) { var s = code + '|' + JSON.stringify(v); if (said[s]) return; said[s] = 1; add(code, Object.assign({}, w, v)); }
        if (!b.onlyVal && !a.onlyVal) {
          if (Math.abs(a.txv - b.txv) > TOL) once('d_txv', { x: a.txv, y: b.txv });
          var ta = tax(a), tb = tax(b);
          if (Math.abs(ta - tb) > TOL && !(a.derived || b.derived)) once('d_tax', { x: ta, y: tb });
          var ia = a.igst > 0.5, ib = b.igst > 0.5, ca = a.cgst + a.sgst > 0.5, cb = b.cgst + b.sgst > 0.5;
          if ((ia || ca) && (ib || cb) && (ia !== ib || ca !== cb)) once('d_head', { x: ia ? 'IGST' : 'CGST + SGST', y: ib ? 'IGST' : 'CGST + SGST' });
        }
        if (a.gstin !== b.gstin && (a.gstin || b.gstin)) once('d_gstin', { x: a.gstin || '(none)', y: b.gstin || '(none)' });
        if (a.date && b.date && a.date !== b.date) once('d_date', { x: a.date, y: b.date });
        if (a.val && b.val && Math.abs(a.val - b.val) > TOL && (a.onlyVal || b.onlyVal || Math.abs(a.txv - b.txv) <= TOL)) once('d_val', { x: a.val, y: b.val });
      });
      var sev = row.what.map(function (w) { return w.sev; });
      row.status = sev.indexOf('miss') >= 0 ? 'miss' : sev.indexOf('diff') >= 0 ? 'diff' : 'ok';
      out.counts[row.status]++;
      out.rows.push(row);
    });
    out.rows.sort(function (a, b) { var o = { miss: 0, diff: 1, cnl: 2, ok: 3 }; return o[a.status] - o[b.status] || (a.date < b.date ? -1 : a.date > b.date ? 1 : 0); });
    /* totals of each file for the common months, so a gap that has no single invoice behind it still shows */
    out.totals = {};
    srcs.forEach(function (s) { var o = { n: 0, txv: 0, tax: 0 }; L[s].forEach(function (r) { if (!inScope(r) || r.cnl) return; var sg = r.kind === 'cn' ? -1 : 1; o.n++; o.txv = r2(o.txv + sg * r.txv); o.tax = r2(o.tax + sg * tax(r)); }); out.totals[s] = o; });
    return out;
  }

  /* ---------- words for each finding (English; the screen uses NS.t with these as the fallback) ---------- */
  var MSG = {
    miss_e: 'No e-invoice (IRN) found for this document, but it goes to a registered buyer, an export or SEZ.',
    miss_g: 'Not in GSTR-1.', b2c: 'Sale to an unregistered buyer: GSTR-1 shows it only in state-wise totals (B2C Small), and it needs no e-invoice.', miss_b: 'Not in your books (sales register).',
    cnl_g: 'The e-invoice was cancelled, but the document is still in GSTR-1.', cnl_b: 'The e-invoice was cancelled, but the bill is not cancelled in your books.',
    bcnl_g: 'Cancelled in your books, but still in GSTR-1.',
    looseno: 'The number is written differently: {n}. Use one way everywhere.',
    d_txv: 'Taxable value differs: {a} {x}, {b} {y}.', d_tax: 'Tax differs: {a} {x}, {b} {y}.', d_head: 'Tax type differs: {a} {x}, {b} {y}.',
    d_gstin: 'Buyer GSTIN differs: {a} {x}, {b} {y}.', d_date: 'Date differs: {a} {x}, {b} {y}.', d_val: 'Invoice value differs: {a} {x}, {b} {y}.'
  };
  function fmtV(v, M) { var o = {}; Object.keys(v || {}).forEach(function (k) { var x = v[k]; o[k] = typeof x === 'number' ? (M ? M(x) : x.toFixed(2)) : /^\d{4}-\d\d-\d\d$/.test(x) ? x.split('-').reverse().join('-') : x; }); return o; }
  function say(w, T, M) { var v = fmtV(w.v, M), s = (T && T('e3.m.' + w.k, v)) || MSG[w.k] || w.k; if (s === 'e3.m.' + w.k) s = MSG[w.k]; return s.replace(/\{(\w+)\}/g, function (m, k) { return v[k] == null ? m : v[k]; }); }
  function dmy(d) { return d ? d.split('-').reverse().join('-') : ''; }
  var KIND = { inv: 'Invoice', cn: 'Credit note', dn: 'Debit note' }, STAT = { ok: 'Matches', diff: 'Differs', miss: 'Missing', cnl: 'Cancelled e-invoice' };
  function sheets(res, T) {
    var p = [['Status', 'Document no', 'Type', 'Date', 'Buyer GSTIN', 'Buyer', 'What to check']];
    res.rows.forEach(function (r) { if (r.status === 'ok' && !r.what.length) return; (r.what.length ? r.what : [{ k: '' }]).forEach(function (w) { p.push([STAT[r.status], r.no, KIND[r.kind], dmy(r.date), r.gstin, r.party, w.k ? say(w, T) : '']); }); });
    var H = ['Status', 'Document no', 'Type', 'Date', 'Buyer GSTIN'], S = [['e', 'E-invoice'], ['g', 'GSTR-1'], ['b', 'Books']];
    S.forEach(function (s) { H.push(s[1] + ' taxable', s[1] + ' tax', s[1] + ' value'); });
    H.push('IRN');
    var a = [H].concat(res.rows.map(function (r) {
      var x = [STAT[r.status], r.no, KIND[r.kind], dmy(r.date), r.gstin];
      S.forEach(function (s) { var q = r[s[0]]; x.push(q ? q.txv : '', q ? tax(q) : '', q ? q.val : ''); });
      x.push(r.e ? r.e.irn : ''); return x;
    }));
    var tt = [['File', 'Documents', 'Taxable value (net of credit notes)', 'Tax (net)', 'Left out: outside the months compared']];
    S.forEach(function (s) { var o = res.totals[s[0]]; if (o) tt.push([s[1], o.n, o.txv, o.tax, res.outside[s[0]]]); });
    tt.push([], ['Months compared', res.scope.join(', ')], ['Made by', 'gstr1.niveshsahayak.in (free, in your browser). Check before you file; this is a working sheet.']);
    return [['To check', p, [12, 18, 12, 12, 18, 26, 80], 'wrap'], ['All documents', a], ['Totals', tt]];
  }

  /* ---------- sample files: books = the GSTR-1 site's sample register; e-invoices and GSTR-1 made from it with planted differences ---------- */
  function sample() {
    var F = G1(), rows = F.sample(), own = '03';
    var R = F.read([{ name: 'Sample sales register', rows: rows }]), docs = F.toDocs(R.lines, { own: own }).docs;
    var ownG = '03AAACN1234A1Z', s = 0; for (var i = 0; i < 14; i++) { var v = CH.indexOf(ownG[i]) * (i % 2 ? 2 : 1); s += Math.floor(v / 36) + v % 36; } ownG += CH[(36 - s % 36) % 36];
    function dd(d) { return d.split('-').reverse().join('/'); }
    var E = [], b2b = {}, cdnr = {};
    docs.forEach(function (d) {
      var tt = docTot(d, own), reg = gstinOk(d.buyer.gstin);
      if (!reg) return;
      var no = d.no, txv = tt.txv, ig = tt.igst, cg = tt.cgst, sg = tt.sgst;
      if (no === 'SMP/002' && d.type === 'invoice') return;                             /* planted: no e-invoice made */
      var eTxv = no === 'SMP/006' ? 11400 : txv, eCg = no === 'SMP/006' ? 1026 : cg, eSg = no === 'SMP/006' ? 1026 : sg;   /* planted: transport line left out of the e-invoice */
      var irn = Array(65).join('0').slice(0, 56) + ('00000000' + (E.length + 1)).slice(-8);
      E.push({ AckNo: 112610000000 + E.length, AckDt: d.date + ' 11:00:00', Irn: irn, Status: no === 'CN/001' ? 'CNL' : 'ACT',
        DocDtls: { Typ: d.type === 'note' ? 'CRN' : 'INV', No: no, Dt: dd(d.date) }, BuyerDtls: { Gstin: d.buyer.gstin, LglNm: d.buyer.name, Pos: d.pos },
        ValDtls: { AssVal: eTxv, IgstVal: ig, CgstVal: eCg, SgstVal: eSg, TotInvVal: r2(eTxv + ig + eCg + eSg) }, TranDtls: { SupTyp: 'B2B' } });
      if (no === 'SMP/008' && d.buyer.gstin.slice(0, 2) === '07') return;              /* planted: left out of GSTR-1 */
      var idt = no === 'SMP/001' ? '05-10-2026' : d.date.split('-').reverse().join('-');   /* planted: date typed wrong in GSTR-1 */
      var it = [{ num: 1, itm_det: { rt: 18, txval: txv, iamt: ig, camt: cg, samt: sg, csamt: 0 } }];
      if (d.type === 'note') { (cdnr[d.buyer.gstin] = cdnr[d.buyer.gstin] || []).push({ ntty: 'C', nt_num: no, nt_dt: idt, val: r2(txv + ig + cg + sg), pos: d.pos, rchrg: 'N', inv_typ: 'R', itms: it }); return; }
      (b2b[d.buyer.gstin] = b2b[d.buyer.gstin] || []).push({ inum: no, idt: idt, val: r2(txv + ig + cg + sg), pos: d.pos, rchrg: 'N', inv_typ: 'R', itms: it });
    });
    E.push({ AckNo: 112619999999, AckDt: '2026-11-20 10:00:00', Irn: Array(65).join('9'), Status: 'ACT', DocDtls: { Typ: 'INV', No: 'SMP/009', Dt: '20/11/2026' },   /* planted: an e-invoice with no bill in the books */
      BuyerDtls: { Gstin: docs[0].buyer.gstin, LglNm: docs[0].buyer.name, Pos: '03' }, ValDtls: { AssVal: 5000, IgstVal: 0, CgstVal: 450, SgstVal: 450, TotInvVal: 5900 }, TranDtls: { SupTyp: 'B2B' } });
    var g1 = { gstin: ownG, fp: '112026', b2b: Object.keys(b2b).map(function (g) { return { ctin: g, inv: b2b[g] }; }), cdnr: Object.keys(cdnr).map(function (g) { return { ctin: g, nt: cdnr[g] }; }),
      b2cs: [{ sply_ty: 'INTRA', pos: '03', typ: 'OE', rt: 18, txval: 1000, camt: 90, samt: 90, csamt: 0 }] };
    return { own: ownG, books: rows, einv: E, gstr1: g1 };
  }

  var API = { readJson: readJson, readSheets: readSheets, readEinv: readEinv, readG1: readG1, readBooks: readBooks, compare: compare, sheets: sheets, say: say, sample: sample,
    ESET: ESET, MSG: MSG, nk: nk, loose: loose, gstinOk: gstinOk, jwt: jwt, tax: tax };
  if (typeof module !== 'undefined' && module.exports) module.exports = API; else G.NSE3 = API;
})(typeof window !== 'undefined' ? window : this);
