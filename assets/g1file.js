/* NSG1File (byof-gstr1, 28 Nov 2026): a sales register from any software -> the bill records gstr1.js sorts into the GSTR-1 tables.
   Pure (no screen), so scripts/test_g1file.js runs it in node. The screen part is g1drop.js.

   In:
   - any Excel / CSV sheet of sales (Tally's Sales Register, Busy / Marg / Vyapar sales registers, a sheet of your own, one row per
     bill or one row per item line): the heading row and columns are found by NSTpl (field set 'g1sales', addSet / match); tax and
     sales ledger columns are summed by kind ("CGST @9%", "Output IGST 18%", "Sales GST 18%", "Exempt Sales"): one bill line per rate.
   - Tally's Day Book XML (sales vouchers, credit notes) through NSTallyIn.plan; also Marg / Vyapar "Export to Tally" files.
   Out: { docs, probs, found, own } where docs are invoice / note records shaped like the ones the document makers save (type, kind, no,
   date, buyer { name, gstin, state }, pos, items, rc, einv.sup, exp, cancelled, orig) plus _rows (where each came from in the file)
   and on every item _tx = { txv, igst, cgst, sgst, cess } taken from the file, so gstr1.js uses the file's own tax figures.
   check(docs, own) lists what to correct, row by row, before any portal file is made: [{ k, sev: 'err'|'warn'|'info', no, at, v }].
   Messages are NS.t('g1f.e.' + k, v) on the screen. Rules: GSTIN check digit (mod 36), GST rates by the bill date (same list as
   NS.gstRatesFor / docs.js RATE_SETS), B2C Large limit (Rs 1 lakh from the August 2024 period, Notification 12/2024-CT; Rs 2.5 lakh
   before), HSN digits (Notification 78/2020-CT: 4 digits up to Rs 5 crore turnover, 6 above). */
(function (G) {
  'use strict';
  var TP = G.NSTpl || (typeof require === 'function' ? require('./tplcore.js') : null);
  function key(s) { return String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); }
  function r2(n) { return Math.round((+n + (n < 0 ? -1e-9 : 1e-9)) * 100) / 100; }
  function t(s) { return String(s == null ? '' : s).trim(); }
  function num(s) { return TP.amt(s).v; }

  var RATES = [0, 0.1, 0.25, 1, 1.5, 3, 5, 6, 7.5, 12, 18, 28, 40];
  function ratesFor(date) {
    var S = [['2026-02-01', [0, 0.1, 0.25, 1.5, 3, 5, 12, 18, 40]], ['2025-09-22', [0, 0.1, 0.25, 1.5, 3, 5, 12, 18, 28, 40]], ['', [0, 0.1, 0.25, 3, 5, 12, 18, 28]]];
    return S.filter(function (r) { return String(date || '') >= r[0]; })[0][1];
  }
  function nearRate(g) { var b = null; RATES.forEach(function (x) { if (b == null || Math.abs(x - g) < Math.abs(b - g)) b = x; }); return Math.abs(b - g) <= 0.6 ? b : r2(g); }
  var CH = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  function gstinOk(g) {
    g = String(g || '').toUpperCase();
    if (!/^[0-9]{2}[0-9A-Z]{13}$/.test(g)) return false;
    var s = 0;
    for (var i = 0; i < 14; i++) { var v = CH.indexOf(g[i]) * (i % 2 ? 2 : 1); s += Math.floor(v / 36) + v % 36; }
    return CH[(36 - s % 36) % 36] === g[14];
  }
  function b2clLimit(date) { return String(date) >= '2024-08-01' ? 100000 : 250000; }

  /* ---------- the columns of a sales register ---------- */
  var SET = { k: 'g1sales', name: 'Sales register', cols: [
    { k: 'no', h: 'Invoice no', need: 1, a: ['invoice number', 'bill no', 'bill number', 'voucher no', 'vch no', 'vch/bill no', 'inv no', 'invoice no.', 'document no', 'doc no', 'note no', 'note number', 'docnum', 'bill', 'voucher number', 'vch bill no', 'invno'] },
    { k: 'date', h: 'Date', need: 1, a: ['invoice date', 'bill date', 'voucher date', 'vch date', 'doc date', 'document date', 'note date', 'docdate', 'inv date'] },
    { k: 'party', h: 'Customer', a: ['particulars', 'party', 'party name', 'customer name', 'buyer', 'buyer name', 'receiver name', 'name of customer', 'account', 'account name', 'ledger', 'billed to', 'cardname', 'name'] },
    { k: 'gstin', h: 'Customer GSTIN', a: ['gstin', 'gstin/uin', 'gstin/uin of recipient', 'gstin of recipient', 'party gstin', 'buyer gstin', 'customer gstin', 'gst no', 'gst number', 'gstin no', 'recipient gstin', 'gstin uin'] },
    { k: 'pos', h: 'Place of supply', a: ['pos', 'state', 'supply state', 'place of supply state', 'state name', 'party state', 'buyer state', 'state code', 'place of supply (state)'] },
    { k: 'type', h: 'Type', a: ['voucher type', 'vch type', 'document type', 'doc type', 'invoice type', 'bill type', 'transaction type', 'note type', 'kind'] },
    { k: 'orig', h: 'Original invoice no', a: ['original invoice number', 'original invoice no.', 'against invoice', 'against bill no', 'against', 'ref invoice no', 'reference no', 'original bill no', 'invoice no against', 'voucher ref. no.', 'voucher ref no', 'ref no', 'reference'] },
    { k: 'origdate', h: 'Original invoice date', a: ['original invoice date', 'against invoice date', 'ref date', 'reference date'] },
    { k: 'hsn', h: 'HSN or SAC', a: ['hsn', 'sac', 'hsn code', 'hsn/sac', 'hsn sac', 'hsn/sac code', 'hsn no'] },
    { k: 'desc', h: 'Item', a: ['item name', 'description', 'product', 'item description', 'stock item', 'goods', 'product name', 'item details', 'dscription'] },
    { k: 'qty', h: 'Quantity', a: ['qty', 'qnty', 'billed qty', 'total quantity'] },
    { k: 'uqc', h: 'Unit', a: ['uqc', 'uom', 'units', 'unit of measure', 'unitmsr'] },
    { k: 'price', h: 'Rate', a: ['price', 'unit price', 'rate per unit', 'item rate', 'selling price', 'pricebefdi'] },
    { k: 'txv', h: 'Taxable value', a: ['value', 'taxable amount', 'taxable amt', 'taxable', 'assessable value', 'basic amount', 'net value', 'taxable val', 'gross amt', 'gross amount', 'amount before tax', 'sub total', 'subtotal', 'item value', 'txval'] },
    { k: 'rate', h: 'GST rate %', a: ['gst rate', 'gst %', 'gst%', 'tax %', 'tax rate', 'rate %', 'rate of tax', 'rate(%)', 'gst rate %', 'applicable % of tax rate x', 'tax rate %', 'igst rate', 'rt', 'vatprcnt'] },
    { k: 'val', h: 'Invoice value', a: ['gross total', 'invoice amount', 'invoice total', 'bill amount', 'bill value', 'total', 'grand total', 'net amount', 'net amt', 'total amount', 'invoice value (rs)', 'note value', 'doctotal', 'amount'] },
    { k: 'rc', h: 'Reverse charge', a: ['reverse charge (y/n)', 'rcm', 'reverse charge y/n'] },
    { k: 'etype', h: 'Supply type', a: ['export type', 'supply type', 'invoice category', 'sez', 'category'] },
    { k: 'port', h: 'Port code', a: ['port', 'port code of export'] },
    { k: 'sbno', h: 'Shipping bill no', a: ['shipping bill number', 'shipping bill', 'sb no', 'bill of export no'] },
    { k: 'sbdate', h: 'Shipping bill date', a: ['sb date', 'shipping bill dt'] },
    { k: 'status', h: 'Status', a: ['cancelled', 'is cancelled', 'bill status', 'voucher status'] }
  ] };
  if (TP && TP.addSet) TP.addSet(SET);

  /* other columns: tax and sales ledgers, summed by kind. "CGST @9%" -> cgst at 18 %, "Sales GST 18%" -> taxable at 18 %, "Exempt Sales" -> 0 % exempt */
  function pctOf(h) { var m = String(h).match(/(\d+(?:\.\d+)?)\s*%/); return m ? +m[1] : null; }
  function ledgerOf(h) {
    var s = key(h); if (!s) return null;
    if (/round|discount|freight|cartage|packing|tcs|tds|balance|qty|quantity|price/.test(s)) return null;
    var rateCol = /\brate\b|\(%\)|^\W*(c|s|i|ut)?gst\s*%$|^(cgst|sgst|igst|utgst)\s*%$/.test(s) && !/@/.test(s);
    var p = pctOf(h);
    if (/cess/.test(s)) return { k: 'cess', rateCol: rateCol };
    if (/\bigst\b|integrated/.test(s)) return { k: 'igst', g: p, rateCol: rateCol };
    if (/\bcgst\b|central tax|central gst/.test(s)) return { k: 'cgst', g: p == null ? null : p * 2, rateCol: rateCol };
    if (/\bsgst\b|\butgst\b|state tax|state\/ut tax|state gst/.test(s)) return { k: 'sgst', g: p == null ? null : p * 2, rateCol: rateCol };
    if (/sale|sales|turnover|supply|supplies/.test(s) && !/return|total|tax\b|register|type|date|no\b/.test(s)) {
      var tk = /exempt/.test(s) ? 'exempt' : /nil/.test(s) ? 'nil' : /non[- ]?gst/.test(s) ? 'nongst' : '';
      return { k: 'sale', g: tk ? 0 : p, tk: tk, exp: /export/.test(s), sez: /\bsez\b/.test(s) };
    }
    return null;
  }

  /* ---------- reading sheets ---------- book: [{ name, rows }]; opts: { saved (remembered layouts), force { sheet: L } } */
  function read(book, opts) {
    opts = opts || {};
    var out = { lines: [], found: [], probs: [] };
    (book || []).forEach(function (sh) {
      var rows = sh.rows || [], L = (opts.force && opts.force[sh.name]) || TP.match('g1sales', sh.name, rows, opts.saved);
      if (!L || !('no' in L.map) || !('date' in L.map)) { if (rows.length > 1) out.found.push({ name: sh.name, kind: '', rows: rows.length, heads: (rows[(L && L.row) || 0] || []).slice(0, 40), headRow: (L && L.row) || 0, map: L ? L.map : {}, missing: ['Invoice no', 'Date'].filter(function (x, i) { return !L || !([ 'no', 'date' ][i] in L.map); }) }); return; }
      var heads = L.heads || rows[L.row] || [], used = {}; Object.keys(L.map).forEach(function (f) { used[L.map[f]] = 1; });
      var leds = [];
      heads.forEach(function (h, i) { if (used[i]) return; var x = ledgerOf(h); if (x) { x.i = i; x.h = h; leds.push(x); } });
      /* a column named like a rate is a rate only when its values look like rates; otherwise it holds amounts */
      leds.forEach(function (x) {
        if (!x.rateCol) return;
        var vals = []; for (var r = L.row + 1; r < Math.min(rows.length, L.row + 200); r++) { var v = num((rows[r] || [])[x.i]); if (v) vals.push(Math.abs(v)); }
        x.rateCol = vals.length > 0 && vals.every(function (v) { return v <= 40 && (RATES.indexOf(v) >= 0 || RATES.indexOf(v * 2) >= 0); });
      });
      var rateLeds = leds.filter(function (x) { return x.rateCol; }); leds = leds.filter(function (x) { return !x.rateCol; });
      var objs = TP.objects(rows, L), n0 = out.lines.length;
      objs.forEach(function (o) {
        var row = rows[o._r - 1] || [];
        o._s = sh.name; o._led = leds.map(function (x) { return { x: x, v: num(row[x.i]) }; }).filter(function (z) { return z.v; });
        if (!o.rate && rateLeds.length) { var rl = rateLeds.map(function (x) { var v = Math.abs(num(row[x.i])); return x.k === 'igst' || x.k === 'cess' ? v : v * 2; }).filter(Boolean); if (rl.length && rateLeds[0].k !== 'cess') o.rate = String(rl[0]); }
        out.lines.push(o);
      });
      out.found.push({ name: sh.name, kind: 'g1sales', rows: out.lines.length - n0, map: L.map, sig: L.sig, heads: heads, headRow: L.row, remembered: !!L.remembered, ledgers: leds.map(function (x) { return x.h; }), missing: [] });
    });
    return out;
  }

  /* ---------- rows -> bill records ---------- opts: { own: state code of the business, ownGstin, seller } */
  function kindOf(s, neg) {
    var k = key(s);
    if (/credit|return|\bcn\b|c\/n|cr note|crn/.test(k)) return 'credit';
    if (/debit|\bdn\b|d\/n|dr note|dbn/.test(k)) return 'debit';
    return neg ? 'credit' : 'invoice';
  }
  function supOf(s, exp) {
    var k = key(s).replace(/[\s_-]+/g, '');
    if (/sezwp|sezwithpay|sez.*withpayment/.test(k)) return 'SEZWP';
    if (/sezwop|sezwithoutpay|sez.*without/.test(k)) return 'SEZWOP';
    if (/deemed|dexp/.test(k)) return 'DEXP';
    if (/expwp|withpayment|wpay/.test(k)) return 'EXPWP';
    if (/expwop|withoutpayment|wopay|lut|bond/.test(k)) return 'EXPWOP';
    if (/export/.test(k) || exp) return 'EXP';
    if (/sez/.test(k)) return 'SEZ';
    return '';
  }
  /* one row -> its lines [{ g, txv, igst, cgst, sgst, cess, tk, hsn, desc, qty, uqc }] (amounts as signed in the file) */
  function linesOf(o) {
    var led = o._led || [], sales = led.filter(function (z) { return z.x.k === 'sale'; }), tax = { igst: 0, cgst: 0, sgst: 0, cess: 0 }, taxBy = {};
    led.forEach(function (z) { if (z.x.k === 'sale') return; tax[z.x.k] += z.v; if (z.x.g != null && z.x.k !== 'cess') { var b = taxBy[z.x.g] || (taxBy[z.x.g] = { igst: 0, cgst: 0, sgst: 0 }); b[z.x.k] += z.v; } });
    var base = { hsn: t(o.hsn).replace(/\s+/g, ''), desc: t(o.desc), qty: o.qty === undefined || t(o.qty) === '' ? '' : num(o.qty), uqc: t(o.uqc) };
    var out = [];
    var given = t(o.rate) !== '' ? TP.amt(String(o.rate).replace('%', '')).v : null;
    if (sales.length && (sales.length > 1 || sales[0].x.g != null || sales[0].x.tk || !t(o.txv))) {
      var sTot = sales.reduce(function (a, z) { return a + z.v; }, 0), left = Object.assign({}, tax);
      sales.forEach(function (z, i) {
        var g = z.x.g, part = Object.assign({}, base, { txv: z.v, tk: z.x.tk, exp: z.x.exp, sez: z.x.sez });
        var b = g != null ? taxBy[g] : null;
        if (b) { part.igst = b.igst; part.cgst = b.cgst; part.sgst = b.sgst; left.igst -= b.igst; left.cgst -= b.cgst; left.sgst -= b.sgst; }
        else part._share = sTot ? z.v / sTot : 0;
        part.g = g; part.cess = sTot ? r2(tax.cess * z.v / sTot) : 0;
        out.push(part);
      });
      /* tax columns without a rate in their names: shared out by value over the sales columns that have no rate-wise tax */
      var free = out.filter(function (p) { return p._share != null; }), fs = free.reduce(function (a, p) { return a + p.txv; }, 0);
      free.forEach(function (p) { var w = fs ? p.txv / fs : 0; p.igst = r2(left.igst * w); p.cgst = r2(left.cgst * w); p.sgst = r2(left.sgst * w); delete p._share; });
    } else {
      var txv = t(o.txv) !== '' ? num(o.txv) : (base.qty !== '' && t(o.price) ? r2(num(o.qty) * num(o.price)) : (sales[0] ? sales[0].v : 0));
      out.push(Object.assign({}, base, { txv: txv, g: given, igst: tax.igst, cgst: tax.cgst, sgst: tax.sgst, cess: tax.cess, sez: /\bsez\b/i.test(o.etype || '') }));
    }
    out.forEach(function (p) {
      var tx = (p.igst || 0) + (p.cgst || 0) + (p.sgst || 0);
      p.taxGiven = !!(p.igst || p.cgst || p.sgst);
      if (p.g == null) p.g = given != null ? given : p.txv && tx ? nearRate(Math.abs(tx / p.txv * 100)) : (p.tk ? 0 : null);
    });
    return out;
  }
  function stateOf(s, gstin) { var c = TP.stateCode(s, ''); if (!c && /^\d{2}/.test(t(gstin))) c = t(gstin).slice(0, 2); return c; }

  function toDocs(lines, opts) {
    opts = opts || {};
    var own = opts.own || '', seller = opts.seller || { name: 'My business', state: own, gstin: opts.ownGstin || '' };
    var groups = [], byK = {}, probs = [], cur = null;
    lines.forEach(function (o) {
      var no = t(o.no), dt = TP.date(o.date);
      if (/^(total|grand total|sub total|subtotal|opening|closing|page total)\b/i.test(t(o.party) || no || t(o.desc)) && !dt) return;   /* register totals */
      if (!no && !t(o.date) && cur && (t(o.desc) || t(o.hsn) || t(o.txv) || (o._led || []).length)) { cur.rows.push(o); return; }   /* item line under the bill row */
      if (!no) { probs.push({ k: 'nono', sev: 'err', at: [o._s, o._r] }); cur = null; return; }
      if (!dt) { probs.push({ k: 'nodate', sev: 'err', no: no, at: [o._s, o._r], v: { d: t(o.date) } }); cur = null; return; }
      var txv0 = t(o.txv) !== '' ? num(o.txv) : 0, kind = kindOf(o.type, txv0 < 0 || num(o.val) < 0);
      var k = kind + '|' + key(no) + '|' + dt + '|' + key(o.gstin || o.party);
      var g = byK[k];
      if (!g) { g = byK[k] = { k: k, kind: kind, no: no, date: dt, head: o, rows: [] }; groups.push(g); }
      g.rows.push(o); cur = g;
    });
    var docs = groups.map(function (g) {
      var o = g.head, gstin = '', party = '', pos = '', rc = '', etype = '', port = '', sbno = '', sbdate = '', val = 0, status = '', orig = '';
      g.rows.forEach(function (r) {
        gstin = gstin || t(r.gstin).toUpperCase().replace(/\s+/g, ''); party = party || t(r.party); pos = pos || t(r.pos); rc = rc || t(r.rc); etype = etype || t(r.etype);
        port = port || t(r.port); sbno = sbno || t(r.sbno); sbdate = sbdate || t(r.sbdate); status = status || t(r.status); orig = orig || t(r.orig);
        if (!val && num(r.val)) val = num(r.val);
      });
      if (/^(na|nil|-|0|urp|unregistered|n\/a|none)$/i.test(gstin)) gstin = '';
      var items = [], neg = g.kind !== 'invoice', exp = false, sez = false;
      g.rows.forEach(function (r) {
        linesOf(r).forEach(function (p) {
          if (!p.txv && !p.igst && !p.cgst && !p.sgst) return;
          var sg = neg ? -1 : 1, s = function (v) { v = +v || 0; return neg ? Math.abs(v) : v; };
          exp = exp || p.exp; sez = sez || p.sez;
          var q = p.qty === '' ? 1 : Math.abs(p.qty) || 1, txv = s(p.txv);
          items.push({ desc: p.desc || 'Sales', hsn: p.hsn, unit: p.uqc, qty: q, rate: q ? txv / q : txv, disc: '', gst: p.g == null ? '' : Math.abs(p.g), tk: p.tk || undefined,
            _tx: { txv: r2(txv), igst: r2(s(p.igst)), cgst: r2(s(p.cgst)), sgst: r2(s(p.sgst)), cess: r2(s(p.cess)) }, _given: p.taxGiven, _row: [r._s, r._r], _neg: sg < 0 && p.txv > 0 });
        });
      });
      var bst = gstinOk(gstin) ? gstin.slice(0, 2) : '';
      var posC = stateOf(pos, '') || (/^\d{2}/.test(gstin) ? gstin.slice(0, 2) : '');
      var sup = supOf(etype, exp);
      if (sup === 'EXP' || sup === 'EXPWP' || sup === 'EXPWOP') { posC = '96'; if (sup === 'EXP') sup = items.some(function (i) { return i._tx.igst; }) ? 'EXPWP' : 'EXPWOP'; }
      if (sup === 'SEZ' || (sez && !sup)) sup = items.some(function (i) { return i._tx.igst; }) ? 'SEZWP' : 'SEZWOP';
      if (!posC) {
        var anyI = items.some(function (i) { return i._tx.igst; }), anyC = items.some(function (i) { return i._tx.cgst || i._tx.sgst; });
        if (anyC || !anyI) posC = own;   /* CGST + SGST: inside the state; nothing else known: own state (said on the screen) */
      }
      var allZero = items.length && items.every(function (i) { return !+i.gst && !i._tx.igst && !i._tx.cgst && !i._tx.sgst; });
      var d = { type: g.kind === 'invoice' ? 'invoice' : 'note', no: g.no, date: g.date, seller: seller, buyer: { name: party || (gstin ? '' : 'Cash sale'), gstin: gstin, state: bst || posC }, pos: posC,
        items: items, rc: /^(y|yes|true|1)$/i.test(rc), _val: val ? Math.abs(val) : 0, _rows: g.rows.map(function (r) { return [r._s, r._r]; }), _src: 'file', _nopos: !stateOf(pos, '') && !gstin };
      if (/cancel/i.test(status) || /cancel/i.test(o.type || '')) d.cancelled = true;
      if (sup) d.einv = { sup: sup === 'EXPWP' || sup === 'EXPWOP' || sup.indexOf('SEZ') === 0 || sup === 'DEXP' ? sup : '' };
      if (sup === 'EXPWP' || sup === 'EXPWOP') d.exp = { port: port.toUpperCase(), sbNo: sbno, sbDate: TP.date(sbdate) };
      if (g.kind === 'invoice') { d.kind = allZero && !(sup || '').match(/^(EXP|SEZ)/) ? 'supply' : 'tax'; }
      else { d.kind = g.kind; d.side = 'sale'; d.withGst = true; d.orig = { no: orig }; d._odate = TP.date(g.head.origdate); }
      return d;
    });
    return { docs: docs, probs: probs };
  }

  /* ---------- Tally Day Book XML ---------- texts: strings or bytes; -> { docs, probs, own } (sales invoices and credit / debit notes to customers) */
  function fromTally(texts, opts) {
    opts = opts || {};
    var TI = G.NSTallyIn || (typeof require === 'function' ? require('./tallyin.js') : null);
    var P = TI.plan(texts, { co: 'tally', noOpening: true, seller: opts.seller });
    var own = opts.own || P.ownState || '', docs = [], probs = [];
    P.docs.forEach(function (d) {
      if (d.type === 'invoice' || (d.type === 'note' && d.side === 'sale')) {
        d.pos = d.pos || d.buyer && d.buyer.state || own;
        d.seller = opts.seller || d.seller; d._src = 'tally'; d._rows = [['Tally', d.vno || d.no]];
        if (d.type === 'note') { d.orig = { no: t(d.orig) }; d.withGst = true; }
        d.items = (d.items || []).concat((d.charges || []).filter(function (c) { return c.amt; }).map(function (c) { return { desc: c.name, hsn: '', qty: 1, rate: c.amt, gst: c.gst === 0 ? 0 : (d.items[0] && d.items[0].gst) || 0, _charge: 1 }; }));
        d.charges = [];
        docs.push(d);
      }
    });
    if (!docs.length) probs.push({ k: 'tallynone', sev: 'err', at: ['Tally', ''] });
    return { docs: docs, probs: probs, own: own, left: P.left, company: P.company };
  }

  /* own state when the user has not said: the state of bills taxed with CGST + SGST */
  function guessOwn(docs) {
    var c = {};
    docs.forEach(function (d) { if ((d.items || []).some(function (i) { return i._tx && (i._tx.cgst || i._tx.sgst); })) { var s = d.pos || d.buyer && d.buyer.state; if (s) c[s] = (c[s] || 0) + 1; } });
    return Object.keys(c).sort(function (a, b) { return c[b] - c[a]; })[0] || '';
  }
  /* bills with no place of supply that were given the own state before it was known */
  function setOwn(docs, own, seller) {
    docs.forEach(function (d) {
      if (seller) d.seller = seller;
      if (!d.pos && own && !(d.items || []).some(function (i) { return i._tx && i._tx.igst; })) d.pos = own;
      if (d.buyer && !d.buyer.state) d.buyer.state = d.pos;
    });
    return docs;
  }

  /* ---------- the checks, before anything is made ---------- */
  function check(docs, opts) {
    opts = opts || {};
    var own = opts.own || '', big = opts.aato === 'big', E = [];
    function add(k, sev, d, v, at) { E.push({ k: k, sev: sev, no: d ? d.no : '', at: at || (d && d._rows && d._rows[0]) || ['', ''], v: v || {} }); }
    if (opts.ownGstin && !gstinOk(opts.ownGstin)) add('own', 'err', null, { g: opts.ownGstin });
    if (!own) add('noown', 'warn', null, {});
    var seen = {}, fyOf = function (dt) { var y = +dt.slice(0, 4); return +dt.slice(5, 7) < 4 ? y - 1 : y; };
    docs.forEach(function (d) {
      var sk = d.type + '|' + fyOf(d.date) + '|' + key(d.no);
      if (seen[sk] && !d.cancelled) add('dup', 'err', d, { o: seen[sk].date, p: seen[sk].buyer.name || seen[sk].buyer.gstin || '' });
      else seen[sk] = d;
      if (d.cancelled) return;
      var g = d.buyer.gstin, reg = gstinOk(g), sup = d.einv && d.einv.sup || '';
      if (g && !reg) add('gstin', 'err', d, { g: g });
      if (reg && d.pos && d.pos !== g.slice(0, 2) && d.pos !== '96' && !/^SEZ/.test(sup)) add('posgst', 'warn', d, { p: d.pos, g: g });
      if (d._nopos && !reg) add(d.pos ? 'nopos' : 'noposx', d.pos ? 'warn' : 'err', d, {});
      if (String(d.no).length > 16) add('nolong', 'err', d, {});
      else if (!/^[A-Za-z0-9\/-]+$/.test(String(d.no))) add('nochar', 'err', d, {});
      var allowed = ratesFor(d.date), intra = d.pos && own ? d.pos === own : null, txv = 0, tax = 0;
      if (!(d.items || []).length) add('noline', 'err', d, {});
      (d.items || []).forEach(function (it) {
        var x = it._tx || { txv: 0, igst: 0, cgst: 0, sgst: 0 }, at = it._row || d._rows && d._rows[0];
        txv += x.txv; tax += x.igst + x.cgst + x.sgst + (x.cess || 0);
        if (it.gst === '' || it.gst == null) add('norate', 'err', d, {}, at);
        else if (allowed.indexOf(+it.gst) < 0) add('rate', 'err', d, { r: it.gst, d: d.date }, at);
        if (d.kind !== 'supply') {
          if (!it.hsn) add('nohsn', 'err', d, { i: it.desc }, at);
          else if (!/^\d+$/.test(it.hsn) || [4, 6, 8].indexOf(it.hsn.length) < 0) add('badhsn', 'err', d, { h: it.hsn }, at);
          else if (big && it.hsn.length < 6) add('hsn6', 'warn', d, { h: it.hsn }, at);
        }
        if (x.igst && (x.cgst || x.sgst)) add('mix', 'err', d, {}, at);
        else if (it._given) {
          if (Math.abs(x.cgst - x.sgst) > 1) add('half', 'warn', d, { c: r2(x.cgst), s: r2(x.sgst) }, at);
          var want = r2(x.txv * (+it.gst || 0) / 100), got = r2(x.igst + x.cgst + x.sgst);
          if (it.gst !== '' && Math.abs(want - got) > Math.max(1, x.txv * 0.0005)) add('taxrate', 'warn', d, { t: got, r: it.gst, w: want }, at);
          if (intra === true && x.igst && !/^(SEZ|EXP)/.test(sup) && d.pos !== '96') add('igstin', 'warn', d, {}, at);
          if (intra === false && (x.cgst || x.sgst)) add('cgstout', 'warn', d, { p: d.pos }, at);
        }
      });
      var calc = r2(txv + tax);
      if (d._val && Math.abs(d._val - calc) > 1.01 && d._src === 'file') add('valgap', 'warn', d, { v: d._val, c: calc });
      var val = d._val || calc;
      if (d.type === 'invoice' && !reg && !/^(EXP|SEZ)/.test(sup) && d.pos && own && d.pos !== own && d.pos !== '96' && val > b2clLimit(d.date)) add('b2cl', 'info', d, { v: r2(val), l: b2clLimit(d.date) });
      if (d.type === 'note' && !(d.orig && d.orig.no)) add('noorig', 'warn', d, {});
      if (d.type === 'note' && d._odate && d._odate > d.date) add('nbefore', 'err', d, {});
    });
    return E;
  }
  function counts(E) { var c = { err: 0, warn: 0, info: 0 }; E.forEach(function (e) { c[e.sev]++; }); return c; }
  function months(docs) { var m = {}; docs.forEach(function (d) { if (d.date) m[d.date.slice(0, 7)] = 1; }); return Object.keys(m).sort(); }

  /* ---------- sample file: a made-up item-wise sales register of two months, with a few mistakes to show the checks ---------- */
  function sample() {
    var H = ['Invoice No', 'Invoice Date', 'Customer Name', 'GSTIN/UIN', 'Place of Supply', 'Voucher Type', 'Against Invoice', 'HSN/SAC', 'Item', 'Qty', 'UQC', 'Taxable Value', 'GST Rate %', 'IGST', 'CGST', 'SGST', 'Invoice Value'];
    var R = [
      ['SMP/001', '04-10-2026', 'Garg Traders (sample)', '03AAACG1234A1Z5', 'Punjab', 'Sales', '', '7306', 'MS Pipe 1 inch', '100', 'KGS', '9500', '18', '', '855', '855', '11210'],
      ['SMP/002', '09-10-2026', 'Delhi Steels Pvt Ltd (sample)', '07AAACD9999B1Z2', 'Delhi', 'Sales', '', '7306', 'MS Pipe 2 inch', '200', 'KGS', '38000', '18', '6840', '', '', '44840'],
      ['SMP/003', '15-10-2026', 'Walk-in customer', '', 'Punjab', 'Sales', '', '7318', 'Bolts', '50', 'NOS', '1000', '18', '', '90', '90', '1180'],
      ['SMP/004', '22-10-2026', 'Rana Builders (sample)', '', 'Haryana', 'Sales', '', '7306', 'MS Pipe 2 inch', '600', 'KGS', '114000', '18', '20520', '', '', '134520'],
      ['SMP/005', '28-10-2026', 'Sharma Hardware (sample)', '03AAACS1234A1Z9', 'Punjab', 'Sales', '', '', 'Paint brush', '10', 'NOS', '500', '18', '', '45', '45', '590'],
      ['SMP/006', '03-11-2026', 'Garg Traders (sample)', '03AAACG1234A1Z5', 'Punjab', 'Sales', '', '7306', 'MS Pipe 1 inch', '120', 'KGS', '11400', '18', '', '1026', '1026', '14502'],
      ['SMP/006', '03-11-2026', 'Garg Traders (sample)', '03AAACG1234A1Z5', 'Punjab', 'Sales', '', '9965', 'Transport charges', '', '', '1000', '5', '', '25', '25', '14502'],
      ['SMP/007', '12-11-2026', 'Walk-in customer', '', 'Punjab', 'Sales', '', '7318', 'Bolts', '80', 'NOS', '1600', '12', '', '96', '96', '1792'],
      ['CN/001', '18-11-2026', 'Garg Traders (sample)', '03AAACG1234A1Z5', 'Punjab', 'Credit Note', 'SMP/006', '7306', 'MS Pipe 1 inch', '10', 'KGS', '950', '18', '', '85.5', '85.5', '1121'],
      ['SMP/008', '25-11-2026', 'Delhi Steels Pvt Ltd (sample)', '07AAACD9999B1Z2', 'Delhi', 'Sales', '', '7306', 'MS Pipe 2 inch', '50', 'KGS', '9500', '18', '1710', '', '', '11210'],
      ['SMP/008', '26-11-2026', 'Kumar & Sons (sample)', '03AAACK1234A1Z1', 'Punjab', 'Sales', '', '7318', 'Bolts', '20', 'NOS', '400', '18', '', '36', '36', '472']
    ];
    return [H].concat(R);
  }
  /* sample GSTINs: the check letter is worked out here so the planted wrong one (Sharma Hardware) is the only bad one */
  function fixSample(rows) {
    return rows.map(function (r, i) { if (!i || !r[3] || /Sharma/.test(r[2])) return r; var g = r[3].slice(0, 14), s = 0; for (var j = 0; j < 14; j++) { var v = CH.indexOf(g[j]) * (j % 2 ? 2 : 1); s += Math.floor(v / 36) + v % 36; } r = r.slice(); r[3] = g + CH[(36 - s % 36) % 36]; return r; });
  }

  var API = { SET: SET, read: read, toDocs: toDocs, fromTally: fromTally, check: check, counts: counts, months: months, guessOwn: guessOwn, setOwn: setOwn, gstinOk: gstinOk, ratesFor: ratesFor,
    ledgerOf: ledgerOf, sample: function () { return fixSample(sample()); } };
  if (typeof module !== 'undefined' && module.exports) module.exports = API; else G.NSG1File = API;
})(typeof window !== 'undefined' ? window : this);
