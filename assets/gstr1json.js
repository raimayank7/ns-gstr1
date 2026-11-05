/* GSTR-1 return JSON for the GST portal (Returns › GSTR-1 › Prepare Offline › Upload), made from the summary gstr1.js works out.
   Pure functions (no screen), so scripts/test_gstr1json.js can check them in node.
   Field names follow the JSON the GST Returns Offline Tool writes and the GSTN GSTR-1 save API: gstin, fp (MMYYYY; a quarter = its last
   month), b2b[ctin, inv[inum, idt DD-MM-YYYY, val, pos, rchrg, inv_typ, itms[num, itm_det]]], b2cl[pos, inv], b2cs[sply_ty, pos, typ, rt, txval...],
   exp[exp_typ, inv], cdnr[ctin, nt[ntty, nt_num, nt_dt...]], cdnur[typ B2CL|EXPWP|EXPWOP...], nil{inv[sply_ty INTRB2B|INTRAB2B|INTRB2C|INTRAB2C]},
   at[pos, sply_ty, itms[rt, ad_amt...]], hsn{hsn_b2b, hsn_b2c} (Table 12 split from the May 2025 period), doc_issue{doc_det[doc_num, docs]}.
   GSTN publishes no official schema file: names were checked against a schema rebuilt from portal tests (gist the-solipsist/d60d4ec9, 2026)
   and the offline tool's own output; doc_num codes and cdnur typ values are from the GSTN API list and still need one real upload to confirm.
   validate() returns problems as [key, values] for NS.t('g1.j.' + key). */
(function (G) {
  'use strict';
  var RATES = [0, 0.1, 0.25, 1, 1.5, 3, 5, 6, 7.5, 12, 18, 28, 40];
  /* Table 13: number of each nature of document */
  var DOCNUM = {
    'Invoices for outward supply': 1, 'Invoices for inward supply from unregistered person': 2, 'Revised Invoice': 3, 'Debit Note': 4, 'Credit Note': 5,
    'Receipt Voucher': 6, 'Payment Voucher': 7, 'Refund Voucher': 8, 'Delivery Challan for job work': 9, 'Delivery Challan for supply on approval': 10,
    'Delivery Challan in case of liquid gas': 11, 'Delivery Challan in case other than by way of supply (excluding at S no. 9 to 11)': 12
  };
  var NILTY = ['INTRB2B', 'INTRAB2B', 'INTRB2C', 'INTRAB2C'];   /* same order as gstr1.js exemp rows */
  var INVTYP = { SEZWP: 'SEWP', SEZWOP: 'SEWOP', DEXP: 'DE' };
  var VERSION = 'GST3.2.4';

  function r2(n) { return Math.round((+n || 0) * 100 + (n < 0 ? -1e-9 : 1e-9)) / 100; }
  function dmy(iso) { var p = String(iso || '').split('-'); return p.length === 3 ? p[2] + '-' + p[1] + '-' + p[0] : ''; }
  function code(pos) { return String(pos || '').padStart(2, '0'); }

  /* GSTIN: 15 characters, last one a check digit (mod 36) */
  var CH = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  function gstinOk(g) {
    g = String(g || '').toUpperCase();
    if (!/^[0-9]{2}[0-9A-Z]{13}$/.test(g)) return false;
    var s = 0;
    for (var i = 0; i < 14; i++) { var v = CH.indexOf(g[i]) * (i % 2 ? 2 : 1); s += Math.floor(v / 36) + v % 36; }
    return CH[(36 - s % 36) % 36] === g[14];
  }
  function stateOk(c) { c = +c; return (c >= 1 && c <= 38) || c === 97; }
  function posOk(c) { return stateOk(c) || +c === 96; }
  function noOk(n) { n = String(n || ''); return n.length > 0 && n.length <= 16 && /^[A-Za-z0-9\/-]+$/.test(n) && !/^[0\/-]+$/.test(n); }

  /* fp: month MMYYYY; a quarter is filed under its last month */
  function fp(P) { var p = String(P.to).split('-'); return p[1] + p[0]; }

  function itm(b, i, flat) {
    var d = { txval: r2(b.txv), rt: b.g, iamt: r2(b.igst), camt: r2(b.cgst), samt: r2(b.sgst), csamt: 0 };
    if (flat) { delete d.camt; delete d.samt; return d; }
    return { num: i + 1, itm_det: d };
  }
  function itmsOf(rates, flat) { return rates.map(function (b, i) { return itm(b, i, flat); }); }

  /* J = gstr1.js work().J (one record per bill/note), r = the rest of the summary */
  function build(r, gstin) {
    var J = r.J, out = { gstin: String(gstin || '').toUpperCase(), fp: fp(r.P), version: VERSION, hash: 'hash' };
    var home = out.gstin.slice(0, 2);
    function grp(list, k, make) { var m = {}, a = []; list.forEach(function (x) { var g = m[x[k]]; if (!g) { g = m[x[k]] = make(x); a.push(g); } g._push(x); }); a.forEach(function (g) { delete g._push; }); return a; }
    if (J.b2b.length) out.b2b = grp(J.b2b, 'ctin', function (x) { var g = { ctin: x.ctin, inv: [] }; g._push = function (y) {
      g.inv.push({ inum: y.no, idt: dmy(y.date), val: r2(y.val), pos: code(y.pos), rchrg: y.rc ? 'Y' : 'N', inv_typ: INVTYP[y.sup] || 'R', itms: itmsOf(y.rates) }); }; return g; });
    if (J.b2cl.length) out.b2cl = grp(J.b2cl, 'pos', function (x) { var g = { pos: code(x.pos), inv: [] }; g._push = function (y) {
      g.inv.push({ inum: y.no, idt: dmy(y.date), val: r2(y.val), itms: itmsOf(y.rates) }); }; return g; });
    if (J.exp.length) out.exp = grp(J.exp, 'etype', function (x) { var g = { exp_typ: x.etype, inv: [] }; g._push = function (y) {
      var v = { inum: y.no, idt: dmy(y.date), val: r2(y.val) };
      if (y.port) v.sbpcode = y.port; if (y.sbno) v.sbnum = y.sbno; if (y.sbdate) v.sbdt = dmy(y.sbdate);
      v.itms = itmsOf(y.rates, true); g.inv.push(v); }; return g; });
    var b2cs = r.b2cs.filter(function (x) { return x.txv; });
    if (b2cs.length) out.b2cs = b2cs.map(function (x) {
      var intra = code(x.pos) === home, o = { sply_ty: intra ? 'INTRA' : 'INTER', pos: code(x.pos), typ: 'OE', txval: r2(x.txv), rt: x.g };
      if (intra) { o.camt = r2(x.cgst); o.samt = r2(x.sgst); } else o.iamt = r2(x.igst);
      o.csamt = 0; return o;
    });
    if (J.cdnr.length) out.cdnr = grp(J.cdnr, 'ctin', function (x) { var g = { ctin: x.ctin, nt: [] }; g._push = function (y) {
      g.nt.push({ ntty: y.ntty, nt_num: y.no, nt_dt: dmy(y.date), val: r2(y.val), pos: code(y.pos), rchrg: y.rc ? 'Y' : 'N', inv_typ: INVTYP[y.sup] || 'R', itms: itmsOf(y.rates) }); }; return g; });
    if (J.cdnur.length) out.cdnur = J.cdnur.map(function (y) {
      var o = { typ: y.ur, ntty: y.ntty, nt_num: y.no, nt_dt: dmy(y.date), val: r2(y.val) };
      if (y.ur === 'B2CL') o.pos = code(y.pos);
      o.itms = itmsOf(y.rates); return o;
    });
    if (r.exemp.some(function (x) { return x[0] || x[1] || x[2]; }))
      out.nil = { inv: r.exemp.map(function (x, i) { return { sply_ty: NILTY[i], nil_amt: r2(x[0]), expt_amt: r2(x[1]), ngsup_amt: r2(x[2]) }; }) };
    if (r.at.length) out.at = grp(r.at.map(function (x) { return Object.assign({ k: code(x.pos) }, x); }), 'k', function (x) {
      var intra = code(x.pos) === home, g = { pos: code(x.pos), sply_ty: intra ? 'INTRA' : 'INTER', itms: [] };
      g._push = function (y) { g.itms.push({ rt: y.g, ad_amt: r2(y.txv), iamt: r2(y.igst), camt: r2(y.cgst), samt: r2(y.sgst), csamt: 0 }); }; return g; });
    function hsnList(list) {
      return list.map(function (h, i) {
        var u = h.svc ? 'NA' : String(h.uqc || 'OTH').split('-')[0];
        return { num: i + 1, hsn_sc: h.hsn, desc: String(h.desc || h.hsn).slice(0, 30), uqc: u, qty: h.svc ? 0 : r2(h.qty), rt: h.g,
          txval: r2(h.txv), iamt: r2(h.igst), camt: r2(h.cgst), samt: r2(h.sgst), csamt: 0 };
      });
    }
    out.hsn = { hsn_b2b: hsnList(r.hsnb2b) };
    if (r.hsnb2c.length) out.hsn.hsn_b2c = hsnList(r.hsnb2c);
    var dd = {};
    r.docs.forEach(function (x) {
      var n = DOCNUM[x.nature]; if (!n) return;
      var g = dd[n] || (dd[n] = { doc_num: n, docs: [] });
      g.docs.push({ num: g.docs.length + 1, from: String(x.from || ''), to: String(x.to || ''), totnum: x.total, cancel: x.cancel, net_issue: x.total - x.cancel });
    });
    out.doc_issue = { doc_det: Object.keys(dd).sort(function (a, b) { return a - b; }).map(function (k) { return dd[k]; }) };
    return out;
  }

  /* problems that make the portal refuse the file (or would be wrong in the return) */
  function validate(r, gstin, aato) {
    var E = [], J = r.J;
    function add(k, v) { E.push([k, v || {}]); }
    if (!gstin) add('nogstin'); else if (!gstinOk(gstin)) add('badown', { g: gstin });
    function rates(no, list) { list.forEach(function (b) { if (RATES.indexOf(+b.g) < 0) add('rate', { no: no, r: b.g }); }); }
    function one(y, needCtin) {
      if (!noOk(y.no)) add(String(y.no || '').length > 16 ? 'nolong' : 'nochar', { no: y.no || '—' });
      if (needCtin && !gstinOk(y.ctin)) add('ctin', { no: y.no, g: y.ctin || '—' });
      if ((y.pos || y.needPos) && !posOk(y.pos)) add('pos', { no: y.no });
      rates(y.no, y.rates);
    }
    J.b2b.forEach(function (y) { y.needPos = 1; one(y, true); });
    J.b2cl.forEach(function (y) { y.needPos = 1; one(y); });
    J.exp.forEach(function (y) { one(y); });
    J.cdnr.forEach(function (y) { y.needPos = 1; one(y, true); });
    J.cdnur.forEach(function (y) { if (y.ur === 'B2CL') y.needPos = 1; one(y); });
    r.b2cs.forEach(function (x) { if (!posOk(x.pos)) add('pos', { no: 'B2CS' }); if (RATES.indexOf(+x.g) < 0) add('rate', { no: 'B2CS', r: x.g }); });
    r.docs.forEach(function (x) { [x.from, x.to].forEach(function (n) { if (String(n || '').length > 16) add('nolong', { no: n }); }); });
    var min = aato === 'big' ? 6 : 4;
    r.hsnb2b.concat(r.hsnb2c).forEach(function (h) {
      if (!h.hsn) add('nohsn', { d: h.desc || '—' });
      else if (!/^\d+$/.test(h.hsn) || h.hsn.length < min || h.hsn.length > 8 || h.hsn.length % 2) add('hsn', { h: h.hsn, n: min });
    });
    var seen = {}; return E.filter(function (e) { var k = e[0] + JSON.stringify(e[1]); if (seen[k]) return false; seen[k] = 1; return true; });
  }

  var API = { build: build, validate: validate, gstinOk: gstinOk, noOk: noOk, fp: fp, DOCNUM: DOCNUM, RATES: RATES };
  if (typeof module !== 'undefined' && module.exports) module.exports = API; else G.NSG1J = API;
})(typeof window !== 'undefined' ? window : this);
