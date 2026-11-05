/* GSTR-1 summary (gstr1.niveshsahayak.in): reads the bills saved in the data home and sorts them into the tables of GSTR-1
   (4A B2B, 5 B2CL, 6A exports, 7 B2CS, 8 nil/exempt, 9B credit/debit notes, 11A advances, 12 HSN B2B/B2C, 13 documents issued).
   Downloads are in the layout of the GST Returns Offline Tool (V3.2.4, GSTR1_Excel_Workbook_Template_V2.2, checked 4 Nov 2026):
   one CSV per section with the tool's own headers, or one Excel workbook with the template's sheet names (header row 4, data from row 5).
   The offline tool turns them into the JSON that is uploaded on the GST portal; we never upload anything. Screen words: NS.t('g1.*'). */
(function () {
  'use strict';
  var $ = NS.$, $$ = NS.$$, esc = NS.esc, num = NS.num, r2 = NS.r2, t = NS.t, money = NS.money;
  var root = $('#g1');
  if (!root) return;
  var XLSX_JS = ['https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js', 'sha384-Pqp51FUN2/qzfxZxBCtF0stpc9ONI6MYZpVqmo8m20SoaQCzf+arZvACkLkirlPz'];
  var UT = { '04': 1, '26': 1, '31': 1, '35': 1, '38': 1 };
  /* place of supply as the offline tool's master list writes it */
  var POSN = { '01': 'Jammu & Kashmir', '04': 'Chandigarh', '26': 'Dadra & Nagar Haveli & Daman & Diu', '31': 'Lakshdweep', '35': 'Andaman & Nicobar Islands', '96': 'Foreign Country' };
  function posTxt(c) { return c ? c + '-' + (POSN[c] || NS.stateName(c)) : ''; }
  var UQC = { nos: 'NOS-NUMBERS', pcs: 'PCS-PIECES', kg: 'KGS-KILOGRAMS', gm: 'GMS-GRAMMES', ltr: 'LTR-LITRES', ml: 'MLT-MILILITRE', mtr: 'MTR-METERS', 'sq ft': 'SQF-SQUARE FEET',
    box: 'BOX-BOX', pack: 'PAC-PACKS', set: 'SET-SETS', pair: 'PRS-PAIRS', dozen: 'DOZ-DOZENS', bag: 'BAG-BAGS', ton: 'TON-TONNES', bottle: 'BTL-BOTTLES', roll: 'ROL-ROLLS', 'sq m': 'SQM-SQUARE METERS', quintal: 'QTL-QUINTAL', unit: 'UNT-UNITS' };
  var MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  function gdate(iso) { var p = String(iso || '').split('-'); return p.length === 3 ? p[2] + '-' + MON[+p[1] - 1] + '-' + p[0] : ''; }
  function key(s) { return String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); }
  function gstinOk(g) { g = String(g || '').trim().toUpperCase(); return g.length === 15 && NS.gstinCheck(g).ok; }
  /* B2C Large: inter-state bill to an unregistered buyer above Rs 1 lakh from the August 2024 period (Notification 12/2024-CT), Rs 2.5 lakh before */
  function b2clLimit(date) { return String(date) >= '2024-08-01' ? 100000 : 250000; }
  var SUPTYPE = { B2B: 'Regular B2B', SEZWP: 'SEZ supplies with payment', SEZWOP: 'SEZ supplies without payment', DEXP: 'Deemed Exp' };
  var SEC = ['b2b', 'b2cl', 'exp', 'b2cs', 'exemp', 'cdnr', 'cdnur', 'at', 'hsnb2b', 'hsnb2c', 'docs'];
  /* section: offline tool sheet name, table no. in GSTR-1, CSV headers (the tool's own files) and Excel headers (the template's order) */
  var LAYOUT = {
    b2b: ['b2b,sez,de', '4A, 4B, 6B, 6C', ['GSTIN/UIN of Recipient', 'Receiver Name', 'Invoice Number', 'Invoice date', 'Invoice Value', 'Place Of Supply', 'Reverse Charge', 'Applicable % of Tax Rate', 'Invoice Type', 'E-Commerce GSTIN', 'Rate', 'Taxable Value', 'Cess Amount']],
    b2cl: ['b2cl', '5', ['Invoice Number', 'Invoice date', 'Invoice Value', 'Place Of Supply', 'Applicable % of Tax Rate', 'Rate', 'Taxable Value', 'Cess Amount', 'E-Commerce GSTIN']],
    exp: ['exp', '6A', ['Export Type', 'Invoice Number', 'Invoice date', 'Invoice Value', 'Port Code', 'Shipping Bill Number', 'Shipping Bill Date', 'Rate', 'Taxable Value', 'Cess Amount']],
    b2cs: ['b2cs', '7', ['Type', 'Place Of Supply', 'Rate', 'Applicable % of Tax Rate', 'Taxable Value', 'Cess Amount', 'E-Commerce GSTIN'], ['Type', 'Place Of Supply', 'Applicable % of Tax Rate', 'Rate', 'Taxable Value', 'Cess Amount', 'E-Commerce GSTIN']],
    exemp: ['exemp', '8', ['Description', 'Nil Rated Supplies', 'Exempted(other than nil rated/non GST supply)', 'Non-GST Supplies']],
    cdnr: ['cdnr', '9B', ['GSTIN/UIN of Recipient', 'Receiver Name', 'Note Number', 'Note Date', 'Note Type', 'Place Of Supply', 'Reverse Charge', 'Note Supply Type', 'Note Value', 'Applicable % of Tax Rate', 'Rate', 'Taxable Value', 'Cess Amount']],
    cdnur: ['cdnur', '9B', ['UR Type', 'Note Number', 'Note Date', 'Note Type', 'Place Of Supply', 'Note Value', 'Applicable % of Tax Rate', 'Rate', 'Taxable Value', 'Cess Amount']],
    at: ['at', '11A', ['Place Of Supply', 'Applicable % of Tax Rate', 'Rate', 'Gross Advance Received', 'Cess Amount']],
    hsnb2b: ['hsn(b2b)', '12', ['HSN', 'Description', 'UQC', 'Total Quantity', 'Total Value', 'Taxable Value', 'Integrated Tax Amount', 'Central Tax Amount', 'State/UT Tax Amount', 'Cess Amount', 'Rate'],
      ['HSN', 'Description', 'UQC', 'Total Quantity', 'Total Value', 'Rate', 'Taxable Value', 'Integrated Tax Amount', 'Central Tax Amount', 'State/UT Tax Amount', 'Cess Amount']],
    hsnb2c: ['hsn(b2c)', '12', ['HSN', 'Description', 'UQC', 'Total Quantity', 'Total Value', 'Taxable Value', 'Integrated Tax Amount', 'Central Tax Amount', 'State/UT Tax Amount', 'Cess Amount', 'Rate'],
      ['HSN', 'Description', 'UQC', 'Total Quantity', 'Total Value', 'Rate', 'Taxable Value', 'Integrated Tax Amount', 'Central Tax Amount', 'State/UT Tax Amount', 'Cess Amount']],
    docs: ['docs', '13', ['Nature of Document', 'Sr. No. From', 'Sr. No. To', 'Total Number', 'Cancelled']]
  };
  var EXEMP = ['Inter-State supplies to registered persons', 'Intra-State supplies to registered persons', 'Inter-State supplies to unregistered persons', 'Intra-State supplies to unregistered persons'];

  var DOCS = [], BIZ = [], R = null;

  /* ---------- one bill: lines by rate, place of supply, value ---------- */
  function taxOn(d) {
    if (d.type === 'cashmemo') return d.withGst === true;
    if (d.type === 'note') return d.withGst !== false;
    return d.kind === 'tax';
  }
  function calc(d) {
    var origin = d.seller && d.seller.state, pos = d.pos || d.buyer && d.buyer.state || origin, tax = taxOn(d);
    var intra = !origin || !pos || origin === pos, o = { pos: pos, intra: intra, lines: [], rates: {}, taxable: 0, tax: 0 };
    /* extra charges on the bill (freight, packing...) are part of the taxable value: counted as lines at their GST rate */
    (d.items || []).concat(window.NSBooks && NSBooks.chargeLines ? NSBooks.chargeLines(d) : []).forEach(function (it) {
      var q = num(it.qty), rate = num(it.rate);
      if (!it.desc && !rate) return;
      var gross = q * rate, txv = r2(gross - gross * num(it.disc) / 100), g = tax ? num(it.gst) : 0, tx = r2(txv * g / 100);
      var l = { hsn: String(it.hsn || '').replace(/\s/g, ''), desc: it.desc || '', unit: it.unit || '', qty: q, g: g, txv: txv, igst: intra ? 0 : tx, cgst: intra ? r2(tx / 2) : 0, sgst: intra ? r2(tx / 2) : 0 };
      o.lines.push(l);
      var b = o.rates[g] || (o.rates[g] = { g: g, txv: 0, igst: 0, cgst: 0, sgst: 0 });
      b.txv += l.txv; b.igst += l.igst; b.cgst += l.cgst; b.sgst += l.sgst;
      o.taxable += txv; o.tax += tx;
    });
    var total = r2(o.taxable + o.tax);
    o.value = d.round ? Math.round(total) : total;
    o.taxable = r2(o.taxable); o.tax = r2(o.tax);
    return o;
  }
  function exportOf(d) {
    var sup = d.einv && d.einv.sup || '';
    if (/^EXP/.test(sup) || (d.buyer && d.buyer.state) === '96' || d.pos === '96') return sup === 'EXPWOP' ? 'WOPAY' : sup === 'EXPWP' ? 'WPAY' : (calc(d).tax > 0 ? 'WPAY' : 'WOPAY');
    return '';
  }
  /* what kind of outward bill it is in GSTR-1 */
  function cls(d) {
    var o = calc(d);
    if (exportOf(d)) return 'exp';
    if (gstinOk(d.buyer && d.buyer.gstin)) return 'b2b';
    return !o.intra && o.value > b2clLimit(d.date) ? 'b2cl' : 'b2cs';
  }

  /* ---------- the whole period ---------- */
  function period() {
    var m = $('#g1-month').value || NS.today().slice(0, 7), q = $('#g1-ptype').value === 'q';
    var y = +m.slice(0, 4), mo = +m.slice(5, 7), from = mo, to = mo;
    if (q) { from = Math.floor((mo - 1) / 3) * 3 + 1; to = from + 2; }
    var pad = function (n) { return (n < 10 ? '0' : '') + n; };
    var last = new Date(y, to, 0).getDate();
    return { from: y + '-' + pad(from) + '-01', to: y + '-' + pad(to) + '-' + pad(last), q: q, label: q ? MON[from - 1] + '–' + MON[to - 1] + ' ' + y : MON[mo - 1] + ' ' + y };
  }
  function addTo(map, k, base) { return map[k] || (map[k] = Object.assign({ txv: 0, igst: 0, cgst: 0, sgst: 0 }, base)); }
  function work() {
    var bz = $('#g1-biz').value, P = period(), aato = $('#g1-aato').value;
    var mine = DOCS.filter(function (d) { return key(d.seller && d.seller.name) === bz; });
    var inP = mine.filter(function (d) { return d.date >= P.from && d.date <= P.to; });
    var byNo = {}; mine.forEach(function (d) { if ((d.type === 'invoice' || d.type === 'cashmemo') && d.no) byNo[key(d.no)] = d; });
    var r = { P: P, b2b: [], b2cl: [], exp: [], b2cs: {}, exemp: [0, 1, 2, 3].map(function () { return [0, 0, 0]; }), cdnr: [], cdnur: [], at: {}, hsnb2b: {}, hsnb2c: {}, docs: {}, warn: [], n: 0,
      tot: { txv: 0, igst: 0, cgst: 0, sgst: 0 }, seller: null, skipped: 0, J: { b2b: [], b2cl: [], exp: [], cdnr: [], cdnur: [] } };
    /* one record per bill or note for the portal JSON (gstr1json.js) */
    function J(s, d, o, x) { r.J[s].push(Object.assign({ no: d.no, date: d.date, val: o.value, pos: o.pos, rc: !!d.rc, rates: Object.keys(o.rates).map(function (g) { return o.rates[g]; }) }, x)); }
    var W = function (k, v) { r.warn.push(t('g1.w.' + k, v)); };
    function tot(o, sign) { r.tot.txv += sign * o.txv; r.tot.igst += sign * o.igst; r.tot.cgst += sign * o.cgst; r.tot.sgst += sign * o.sgst; }
    function hsn(bucket, d, o, sign) {
      o.lines.forEach(function (l) {
        var svc = /^99/.test(l.hsn), u = svc ? '' : UQC[key(l.unit)] || 'OTH-OTHERS';
        var h = addTo(bucket, l.hsn + '|' + l.g + '|' + u, { hsn: l.hsn, desc: l.desc, uqc: u, qty: 0, g: l.g, svc: svc });
        h.qty += sign * (svc ? 0 : l.qty); h.txv += sign * l.txv; h.igst += sign * l.igst; h.cgst += sign * l.cgst; h.sgst += sign * l.sgst;
        if (!l.hsn) W('nohsn', { no: d.no, item: l.desc });
        else if (!/^\d+$/.test(l.hsn) || [4, 6, 8].indexOf(l.hsn.length) < 0) W('badhsn', { no: d.no, h: l.hsn });
        else if (aato === 'big' && l.hsn.length < 6) W('hsn6', { no: d.no, h: l.hsn });
        else if (bucket === r.hsnb2b && l.hsn.length < 4) W('hsn4', { no: d.no, h: l.hsn });
      });
    }
    function docRow(nature, d) { var x = r.docs[nature + '|' + series(d.no)] || (r.docs[nature + '|' + series(d.no)] = { nature: nature, list: [] }); x.list.push(d.no || ''); }
    inP.forEach(function (d) {
      var ty = d.type || 'invoice';
      if (!r.seller && d.seller) r.seller = d.seller;
      if (ty === 'invoice' || ty === 'cashmemo') {
        docRow('Invoices for outward supply', d);
        var o = calc(d), reg = gstinOk(d.buyer && d.buyer.gstin), place = (o.intra ? 1 : 0) + (reg ? 0 : 2);
        if (!taxOn(d)) {
          if (d.kind === 'supply') { r.exemp[place][1] += o.taxable; r.n++; } else r.skipped++;
          return;
        }
        r.n++;
        if (d.irn) W('irn', { no: d.no });
        if (d.buyer && d.buyer.gstin && !reg) W('badgstin', { no: d.no, g: d.buyer.gstin });
        var c = cls(d), rc = d.rc ? 'Y' : 'N', rates = Object.keys(o.rates).map(function (g) { return o.rates[g]; });
        rates.forEach(function (b) { tot(b, 1); });
        if (d.date >= '2026-02-01' && rates.some(function (b) { return +b.g === 28; })) W('r28', { no: d.no });   /* 28% left the GST rate list on 1 Feb 2026 (same rule as docs.js RATE_SETS) */
        if (c === 'b2b') {
          var st = SUPTYPE[d.einv && d.einv.sup] || 'Regular B2B';
          rates.forEach(function (b) { r.b2b.push([d.buyer.gstin.toUpperCase(), d.buyer.name || '', d.no, gdate(d.date), o.value, posTxt(o.pos), rc, '', st, '', b.g, r2(b.txv), '', b]); });
          J('b2b', d, o, { ctin: d.buyer.gstin.toUpperCase().trim(), sup: d.einv && d.einv.sup });
          hsn(r.hsnb2b, d, o, 1);
        } else if (c === 'b2cl') {
          rates.forEach(function (b) { r.b2cl.push([d.no, gdate(d.date), o.value, posTxt(o.pos), '', b.g, r2(b.txv), '', '', b]); });
          J('b2cl', d, o);
          hsn(r.hsnb2c, d, o, 1);
        } else if (c === 'exp') {
          var et = exportOf(d), ex = d.exp || {};
          rates.forEach(function (b) { r.exp.push([et, d.no, gdate(d.date), o.value, ex.port || '', ex.sbNo || '', gdate(ex.sbDate), b.g, r2(b.txv), '', b]); });
          J('exp', d, o, { etype: et, port: String(ex.port || '').toUpperCase(), sbno: /^\d{1,7}$/.test(ex.sbNo || '') ? +ex.sbNo : '', sbdate: ex.sbDate || '' });
          W('exp', { no: d.no });
          hsn(r.hsnb2c, d, o, 1);
        } else {
          rates.forEach(function (b) {
            if (!b.g) { r.exemp[place][0] += b.txv; return; }
            var s = addTo(r.b2cs, o.pos + '|' + b.g, { pos: o.pos, g: b.g }); s.txv += b.txv; s.igst += b.igst; s.cgst += b.cgst; s.sgst += b.sgst;
          });
          hsn(r.hsnb2c, d, o, 1);
        }
        if (!d.buyer || !d.pos && !d.buyer.state) W('nopos', { no: d.no });
      } else if (ty === 'note') {
        docRow(d.kind === 'debit' ? 'Debit Note' : 'Credit Note', d);
        if (!taxOn(d)) { r.skipped++; return; }
        r.n++;
        var on = calc(d), sign = d.kind === 'debit' ? 1 : -1, orig = byNo[key(d.orig && d.orig.no)], nt = d.kind === 'debit' ? 'D' : 'C';
        var pos = orig ? calc(orig).pos : on.pos, nrates = Object.keys(on.rates).map(function (g) { return on.rates[g]; });
        nrates.forEach(function (b) { tot(b, sign); });
        if (!orig) W('noorig', { no: d.no, o: d.orig && d.orig.no || '—' });
        if (gstinOk(d.buyer && d.buyer.gstin)) {
          var nst = SUPTYPE[orig && orig.einv && orig.einv.sup] || 'Regular B2B';
          nrates.forEach(function (b) { r.cdnr.push([d.buyer.gstin.toUpperCase(), d.buyer.name || '', d.no, gdate(d.date), nt, posTxt(pos), d.rc ? 'Y' : 'N', nst, on.value, '', b.g, r2(b.txv), '', b]); });
          J('cdnr', d, on, { ctin: d.buyer.gstin.toUpperCase().trim(), pos: pos, ntty: nt, sup: orig && orig.einv && orig.einv.sup });
          hsn(r.hsnb2b, d, on, sign);
        } else {
          var oc = orig ? cls(orig) : '';
          if (oc === 'b2cl' || oc === 'exp') {
            var ur = oc === 'b2cl' ? 'B2CL' : exportOf(orig) === 'WPAY' ? 'EXPWP' : 'EXPWOP';
            nrates.forEach(function (b) { r.cdnur.push([ur, d.no, gdate(d.date), nt, oc === 'b2cl' ? posTxt(pos) : '', on.value, '', b.g, r2(b.txv), '', b]); });
            J('cdnur', d, on, { ur: ur, pos: pos, ntty: nt });
          } else {   /* small B2C bills: the note changes the B2CS figures (a credit note lowers them) */
            nrates.forEach(function (b) {
              var s = addTo(r.b2cs, pos + '|' + b.g, { pos: pos, g: b.g }); s.txv += sign * b.txv; s.igst += sign * b.igst; s.cgst += sign * b.cgst; s.sgst += sign * b.sgst;
            });
          }
          hsn(r.hsnb2c, d, on, sign);
        }
      } else if (ty === 'self') docRow('Invoices for inward supply from unregistered person', d);
      else if (ty === 'receipt' && d.kind === 'adv' && d.gstAdv) {
        docRow('Receipt Voucher', d);
        var g = num(d.rate), amt = num(d.amount), apos = d.pos || d.seller && d.seller.state, base = r2(amt * 100 / (100 + g)), intra = !apos || apos === (d.seller && d.seller.state), tx = r2(amt - base);
        var a = addTo(r.at, apos + '|' + g, { pos: apos, g: g }); a.txv += base; a.igst += intra ? 0 : tx; a.cgst += intra ? tx / 2 : 0; a.sgst += intra ? tx / 2 : 0;
        W('adv', { no: d.no });
      } else if (ty === 'receipt' && d.kind === 'ref') docRow('Refund Voucher', d);
      else if (ty === 'voucher' && d.rcm) docRow('Payment Voucher', d);
      else if (ty === 'challan') {
        var nat = { job: 'Delivery Challan for job work', approval: 'Delivery Challan for supply on approval', transfer: 'Delivery Challan in case other than by way of supply (excluding at S no. 9 to 11)', other: 'Delivery Challan in case other than by way of supply (excluding at S no. 9 to 11)' }[d.kind];
        if (nat) docRow(nat, d);
      }
    });
    if (r.skipped) W('plain', { n: r.skipped });
    ['b2cs', 'at', 'hsnb2b', 'hsnb2c'].forEach(function (s) { r[s] = Object.keys(r[s]).map(function (k) { return r[s][k]; }); });
    r.b2cs.forEach(function (s) { tot0(s); }); r.at.forEach(function (s) { tot0(s); });
    function tot0(s) { s.txv = r2(s.txv); s.igst = r2(s.igst); s.cgst = r2(s.cgst); s.sgst = r2(s.sgst); }
    r.hsnb2b.forEach(tot0); r.hsnb2c.forEach(tot0);
    r.docs = Object.keys(r.docs).map(function (k) { return docSum(r.docs[k], W); });
    Object.keys(r.tot).forEach(function (k) { r.tot[k] = r2(r.tot[k]); });
    var uniq = {}; r.warn = r.warn.filter(function (w) { if (uniq[w]) return false; uniq[w] = 1; return true; });
    return r;
  }
  /* series = the number without its running count: INV/26-27/014 -> INV/26-27/ */
  function series(no) { return String(no || '').replace(/\d+\s*$/, ''); }
  function tail(no) { var m = String(no || '').match(/(\d+)\s*$/); return m ? +m[1] : NaN; }
  function docSum(x, W) {
    var l = x.list.slice().sort(function (a, b) { return (tail(a) - tail(b)) || String(a).localeCompare(String(b)); });
    var lo = tail(l[0]), hi = tail(l[l.length - 1]), seen = {}, dup = 0;
    l.forEach(function (n) { if (seen[key(n)]) dup++; seen[key(n)] = 1; });
    if (dup) W('dup', { s: series(l[0]) || l[0] });
    if (isFinite(lo) && isFinite(hi) && hi - lo + 1 > l.length) W('gap', { s: series(l[0]), n: hi - lo + 1 - l.length });
    var total = isFinite(lo) && isFinite(hi) ? Math.max(hi - lo + 1, l.length) : l.length;
    return { nature: x.nature, from: l[0], to: l[l.length - 1], total: total, cancel: total - l.length };
  }

  /* ---------- rows for the files ---------- */
  function rows(s, r, excel) {
    var cut = function (a) { return a.slice(0, a.length - 1); };
    switch (s) {
      case 'b2b': return r.b2b.map(cut);
      case 'b2cl': return r.b2cl.map(cut);
      case 'exp': return r.exp.map(cut);
      case 'b2cs': return r.b2cs.map(function (x) { return excel ? ['OE', posTxt(x.pos), '', x.g, x.txv, '', ''] : ['OE', posTxt(x.pos), x.g, '', x.txv, '', '']; });
      case 'exemp': return EXEMP.map(function (n, i) { return [n, r2(r.exemp[i][0]), r2(r.exemp[i][1]), r2(r.exemp[i][2])]; });
      case 'cdnr': return r.cdnr.map(cut);
      case 'cdnur': return r.cdnur.map(cut);
      case 'at': return r.at.map(function (x) { return [posTxt(x.pos), '', x.g, x.txv, '']; });
      case 'hsnb2b': case 'hsnb2c': return r[s].map(function (h) {
        var tv = r2(h.txv + h.igst + h.cgst + h.sgst), q = h.svc ? '' : r2(h.qty);
        return excel ? [h.hsn, h.desc, h.uqc, q, tv, h.g, h.txv, h.igst, h.cgst, h.sgst, ''] : [h.hsn, h.desc, h.uqc, q, tv, h.txv, h.igst, h.cgst, h.sgst, '', h.g];
      });
      case 'docs': return r.docs.map(function (x) { return [x.nature, x.from, x.to, x.total, x.cancel]; });
    }
    return [];
  }
  function csvCell(v) { v = v == null ? '' : String(v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }
  function save(blob, name) {
    var a = document.createElement('a'), u = URL.createObjectURL(blob);
    a.href = u; a.download = name; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(u); }, 4000);
  }
  function fname(ext) { var s = R.seller || {}; return 'GSTR-1 ' + (s.gstin || s.name || '') + ' ' + R.P.label + (ext ? ' ' + ext : ''); }
  function csv(s) {
    var L = LAYOUT[s], body = [L[2]].concat(rows(s, R, false)).map(function (r) { return r.map(csvCell).join(','); }).join('\r\n') + '\r\n';
    save(new Blob([body], { type: 'text/csv;charset=utf-8' }), fname(L[0]).replace(/[\\\/:*?"<>|]+/g, '-') + '.csv');
  }
  function excel() {
    NS.toast(t('g1.making'));
    NS.loadJs(XLSX_JS[0], XLSX_JS[1]).then(function () {
      var wb = new window.ExcelJS.Workbook();
      wb.creator = 'Nivesh Sahayak'; wb.company = 'Nivesh Sahayak';
      SEC.forEach(function (s) {
        var L = LAYOUT[s], ws = wb.addWorksheet(L[0]), hd = L[3] || L[2];
        ws.getCell('A1').value = 'Summary for ' + L[0] + ' (' + L[1] + ') ' + R.P.label;
        ws.getRow(4).values = hd; ws.getRow(4).font = { bold: true };
        rows(s, R, true).forEach(function (r, i) { ws.getRow(5 + i).values = r; });
        hd.forEach(function (h, i) { ws.getColumn(i + 1).width = Math.max(12, Math.min(40, h.length + 2)); });
      });
      return wb.xlsx.writeBuffer();
    }).then(function (buf) {
      save(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), fname('').replace(/[\\\/:*?"<>|]+/g, '-').trim() + '.xlsx');
    }).catch(function () { NS.toast(t('g1.xfail'), 'bad'); });
  }

  /* GSTR-1 JSON for the portal's Prepare Offline › Upload: checked first, problems listed in plain words */
  function ownGstin() {
    var s = R.seller || {}, bz = $('#g1-biz').value, b = BIZ.filter(function (x) { return key(x.name) === bz && x.gstin; })[0];
    return String(s.gstin || b && b.gstin || '').trim().toUpperCase();
  }
  function json() {
    var box = $('#g1-jerr'), g = ownGstin(), E = NSG1J.validate(R, g, $('#g1-aato').value);
    if (E.length) {
      box.hidden = false;
      box.innerHTML = '<b>' + esc(t('g1.j.fix', { n: E.length })) + '</b><ul>' + E.slice(0, 40).map(function (e) { return '<li>' + esc(t('g1.j.' + e[0], e[1])) + '</li>'; }).join('') + '</ul>' +
        (E.length > 40 ? '<p>' + esc(t('g1.more', { n: E.length - 40 })) + '</p>' : '');
      box.scrollIntoView({ block: 'nearest' });
      return;
    }
    box.hidden = true;
    var out = NSG1J.build(R, g);
    save(new Blob([JSON.stringify(out)], { type: 'application/json' }), 'GSTR1_' + g + '_' + out.fp + '.json');
    NS.toast(t('g1.j.done'));
  }
  function jsonButton() {
    var row = $('#g1-files .row');
    if (!row || $('[data-do="json"]', row)) return;
    var b = NS.el('button', { type: 'button', 'class': 'btn primary', 'data-do': 'json', 'data-ic': 'down' });
    b.textContent = t('g1.j.btn'); b.setAttribute('data-i18n', 'g1.j.btn');
    row.insertBefore(b, row.firstChild);
    var x = $('[data-do="xlsx"]', row); if (x) x.classList.remove('primary');
    var h = NS.el('p', { 'class': 'soft g1-jh', 'data-i18n': 'g1.j.h' }); h.textContent = t('g1.j.h');
    row.parentNode.insertBefore(h, row.nextSibling);
    var e = NS.el('div', { id: 'g1-jerr', 'class': 'callout g1-warn', role: 'alert' }); e.hidden = true;
    h.parentNode.insertBefore(e, h.nextSibling);
  }

  /* ---------- screen ---------- */
  function m(n) { return money(n, false); }
  function sumOf(list, f) { return r2(list.reduce(function (a, x) { return a + f(x); }, 0)); }
  function table(s, head, body, foot) {
    var L = LAYOUT[s];
    return '<section class="g1-sec" data-sec="' + s + '"><div class="g1-sh"><h3>' + esc(t('g1.s.' + s)) + ' <small class="soft">' + esc(t('g1.table', { n: L[1] })) + '</small></h3>' +
      '<button type="button" class="btn sm" data-csv="' + s + '">' + esc(t('g1.csv')) + '</button></div>' +
      (body.length ? '<div class="g1-tw"><table class="g1-t"><thead><tr>' + head.map(function (h, i) { return '<th' + (i >= head.length - (foot ? foot.length : 0) ? ' class="n"' : '') + '>' + esc(h) + '</th>'; }).join('') + '</tr></thead><tbody>' +
        body.map(function (r) { return '<tr>' + r.map(function (c, i) { return '<td' + (typeof c === 'number' ? ' class="n"' : '') + '>' + esc(typeof c === 'number' ? m(c) : c) + '</td>'; }).join('') + '</tr>'; }).join('') + '</tbody>' +
        (foot ? '<tfoot><tr><td colspan="' + (head.length - foot.length) + '">' + esc(t('g1.total')) + '</td>' + foot.map(function (f) { return '<td class="n">' + m(f) + '</td>'; }).join('') + '</tr></tfoot>' : '') + '</table></div>'
        : '<p class="soft">' + esc(t('g1.none')) + '</p>') + '</section>';
  }
  function taxCols(b) { return [r2(b.txv), r2(b.igst), r2(b.cgst), r2(b.sgst)]; }
  function TH() { return [t('g1.c.txv'), 'IGST', 'CGST', 'SGST/UTGST']; }
  function feet(list, f) { return [sumOf(list, function (x) { return f(x).txv; }), sumOf(list, function (x) { return f(x).igst; }), sumOf(list, function (x) { return f(x).cgst; }), sumOf(list, function (x) { return f(x).sgst; })]; }
  function last(a) { return a[a.length - 1]; }
  function draw() {
    R = work();
    var r = R, out = $('#g1-out'), s = r.seller || {};
    var tax = r2(r.tot.igst + r.tot.cgst + r.tot.sgst);
    var h = '<div class="g1-print-hd"><b>' + esc(s.name || '') + '</b>' + (s.gstin ? ' · GSTIN ' + esc(s.gstin) : '') + '<br>' + esc(t('g1.for', { p: r.P.label })) + '</div>' +
      '<div class="g1-tiles"><div><span>' + esc(t('g1.k.n')) + '</span><b>' + r.n + '</b></div><div><span>' + esc(t('g1.c.txv')) + '</span><b>' + m(r.tot.txv) + '</b></div>' +
      '<div><span>IGST</span><b>' + m(r.tot.igst) + '</b></div><div><span>CGST</span><b>' + m(r.tot.cgst) + '</b></div><div><span>SGST/UTGST</span><b>' + m(r.tot.sgst) + '</b></div><div><span>' + esc(t('g1.k.tax')) + '</span><b>' + m(tax) + '</b></div></div>' +
      (r.warn.length ? '<div class="callout g1-warn"><b>' + esc(t('g1.check')) + '</b><ul>' + r.warn.slice(0, 40).map(function (w) { return '<li>' + esc(w) + '</li>'; }).join('') + '</ul>' + (r.warn.length > 40 ? '<p>' + esc(t('g1.more', { n: r.warn.length - 40 })) + '</p>' : '') + '</div>' : '') +
      (!r.n ? '<p class="callout">' + esc(t('g1.empty')) + '</p>' : '');
    var inv = [t('g1.c.no'), t('g1.c.date'), t('g1.c.val'), t('g1.c.pos'), t('g1.c.rate')].concat(TH());
    h += table('b2b', ['GSTIN', t('g1.c.name')].concat(inv), r.b2b.map(function (x) { return [x[0], x[1], x[2], x[3], x[4], x[5], x[10] + '%'].concat(taxCols(last(x))); }), feet(r.b2b, last));
    h += table('b2cl', inv, r.b2cl.map(function (x) { return [x[0], x[1], x[2], x[3], x[5] + '%'].concat(taxCols(last(x))); }), feet(r.b2cl, last));
    h += table('exp', [t('g1.c.etype')].concat(inv.slice(0, 3), [t('g1.c.rate')], TH()), r.exp.map(function (x) { return [x[0], x[1], x[2], x[3], x[7] + '%'].concat(taxCols(last(x))); }), feet(r.exp, last));
    h += table('b2cs', [t('g1.c.pos'), t('g1.c.rate')].concat(TH()), r.b2cs.map(function (x) { return [posTxt(x.pos), x.g + '%'].concat(taxCols(x)); }), feet(r.b2cs, function (x) { return x; }));
    h += table('exemp', [t('g1.c.desc'), t('g1.c.nil'), t('g1.c.exm'), t('g1.c.ngst')], r.exemp.some(function (x) { return x[0] || x[1] || x[2]; }) ? rows('exemp', r) : [], null);
    var nh = [t('g1.c.nno'), t('g1.c.date'), t('g1.c.nt'), t('g1.c.val'), t('g1.c.rate')].concat(TH());
    h += table('cdnr', ['GSTIN', t('g1.c.name')].concat(nh), r.cdnr.map(function (x) { return [x[0], x[1], x[2], x[3], x[4] === 'C' ? t('g1.cn') : t('g1.dn'), x[8], x[10] + '%'].concat(taxCols(last(x))); }), null);
    h += table('cdnur', [t('g1.c.urt')].concat(nh), r.cdnur.map(function (x) { return [x[0], x[1], x[2], x[3] === 'C' ? t('g1.cn') : t('g1.dn'), x[5], x[7] + '%'].concat(taxCols(last(x))); }), null);
    h += table('at', [t('g1.c.pos'), t('g1.c.rate')].concat(TH()), r.at.map(function (x) { return [posTxt(x.pos), x.g + '%'].concat(taxCols(x)); }), feet(r.at, function (x) { return x; }));
    ['hsnb2b', 'hsnb2c'].forEach(function (k) {
      h += table(k, ['HSN', t('g1.c.desc'), 'UQC', t('g1.c.qty'), t('g1.c.rate')].concat(TH()), r[k].map(function (x) { return [x.hsn || '—', x.desc, x.uqc, x.svc ? '' : NS.qty(x.qty), x.g + '%'].concat(taxCols(x)); }), feet(r[k], function (x) { return x; }));
    });
    h += table('docs', [t('g1.c.nat'), t('g1.c.from'), t('g1.c.to'), t('g1.c.tot'), t('g1.c.can')], r.docs.map(function (x) { return [x.nature, x.from, x.to, String(x.total), String(x.cancel)]; }), null);
    out.innerHTML = h;
    var je = $('#g1-jerr'); if (je) je.hidden = true;
    $('#g1-files').hidden = false;
    try { localStorage.setItem('ns-g1', JSON.stringify({ biz: $('#g1-biz').value, p: $('#g1-ptype').value, m: $('#g1-month').value, a: $('#g1-aato').value })); } catch (e) {}
  }

  function bizList() {
    var seen = {}, out = [];
    BIZ.concat(DOCS.map(function (d) { return d.seller; })).forEach(function (s) { var k = key(s && s.name); if (k && !seen[k]) { seen[k] = 1; out.push([k, s.name + (s.gstin ? ' (' + s.gstin + ')' : '')]); } });
    return out;
  }
  function load() {
    var st = $('#g1-status');
    NSStore.ready().then(function () { return Promise.all([NSStore.list('doc'), NSStore.list('biz')]); }).then(function (x) {
      DOCS = x[0] || []; BIZ = x[1] || [];
      var bl = bizList(), sel = $('#g1-biz');
      sel.innerHTML = bl.length ? bl.map(function (b) { return '<option value="' + esc(b[0]) + '">' + esc(b[1]) + '</option>'; }).join('') : '<option value="">' + esc(t('g1.nobiz')) + '</option>';
      var saved = {}; try { saved = JSON.parse(localStorage.getItem('ns-g1') || '{}'); } catch (e) {}
      var prev = new Date(); prev.setDate(1); prev.setMonth(prev.getMonth() - 1);
      $('#g1-month').value = saved.m || prev.getFullYear() + '-' + ('0' + (prev.getMonth() + 1)).slice(-2);
      if (saved.biz && bl.some(function (b) { return b[0] === saved.biz; })) sel.value = saved.biz;
      if (saved.p) $('#g1-ptype').value = saved.p;
      var aa; try { aa = localStorage.getItem('ns-aato'); } catch (e) {}
      $('#g1-aato').value = saved.a || (aa && /10|big|5/.test(aa) && !/upto|small/.test(aa) ? 'big' : 'small');
      st.textContent = t(NSStore.mode() === 'vault' ? 'g1.vault' : 'g1.local', { n: DOCS.length });
      root.classList.add('on');
      draw();
    }).catch(function () { st.textContent = t('g1.fail'); });
  }
  function start() {
    jsonButton();
    load();
    document.addEventListener('ns-lang', function () { if (R) draw(); });   /* the language pack can arrive after the first drawing */
    root.addEventListener('change', function (e) { if (e.target.closest('.g1-ctl')) draw(); });
    root.addEventListener('click', function (e) {
      var b = e.target.closest('button'); if (!b || !R) return;
      if (b.dataset.csv) csv(b.dataset.csv);
      else if (b.dataset.do === 'json') json();
      else if (b.dataset.do === 'xlsx') excel();
      else if (b.dataset.do === 'allcsv') SEC.forEach(function (s, i) { if (rows(s, R).length) setTimeout(function () { csv(s); }, i * 400); });
      else if (b.dataset.do === 'print') window.print();
      else if (b.dataset.do === 'share') NS.sharePdf([$('#g1-out')], fname(''), 'a4', 'GSTR-1 ' + R.P.label, 'GSTR-1 summary ' + R.P.label);
      else if (b.dataset.do === 'refresh') load();
    });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
})();
