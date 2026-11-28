/* "Bring your own file" for gstr1.niveshsahayak.in (byof-gstr1, 28 Nov 2026): drop a sales register (Tally, Busy, Marg, Excel, CSV) or
   Tally's Day Book XML -> what was read, every row to correct or check -> the GSTR-1 tables, JSON and Excel of gstr1.js, made from the file.
   The file is read in this browser only: nothing is uploaded, nothing is saved to NSBiz (there is no "Save to NSBiz" here on purpose:
   a register has no item master or party details, so it would make half-filled bills). Uses NSG1File (g1file.js), NSTpl, NSTallyIn,
   SheetJS from /assets/vendor (Excel .xls too). Column choices are remembered per heading row in the data home (meta 'g1layouts'). */
(function () {
  'use strict';
  var $ = NS.$, $$ = NS.$$, esc = NS.esc, t = NS.t, F = window.NSG1File;
  var box = $('#g1-drop'); if (!box || !F || !window.NSG1) return;
  var XLJS = '/assets/vendor/xlsx-0.20.3.full.min.js';
  var ST = { files: [], book: [], texts: [], name: '', force: {}, saved: {}, E: [], lines: null, tally: null, cols: null, company: '' };
  var MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  function ml(m) { return MON[+m.slice(5, 7) - 1] + ' ' + m.slice(0, 4); }

  /* remembered column choices */
  function loadSaved() {
    var ls = {}; try { ls = JSON.parse(localStorage.getItem('ns-g1-layouts') || '{}'); } catch (e) {}
    ST.saved = ls;
    if (window.NSStore && NSStore.meta) return NSStore.ready().then(function () { return NSStore.meta('g1layouts'); }).then(function (m) { if (m && typeof m === 'object') ST.saved = Object.assign(ls, m); }, function () {});
    return Promise.resolve();
  }
  function remember(sig, map) {
    if (!sig) return; ST.saved[sig] = { kind: 'g1sales', map: map, at: Date.now() };
    var k = Object.keys(ST.saved); if (k.length > 60) delete ST.saved[k.sort(function (a, b) { return (ST.saved[a].at || 0) - (ST.saved[b].at || 0); })[0]];
    try { localStorage.setItem('ns-g1-layouts', JSON.stringify(ST.saved)); } catch (e) {}
    if (window.NSStore && NSStore.meta) NSStore.meta('g1layouts', ST.saved).catch(function () {});
  }

  /* ---------- reading the files ---------- */
  function isXml(name, b) {
    if (/\.xml$/i.test(name)) return true;
    var i = 0; if (b[0] === 0xFF || b[0] === 0xFE || b[0] === 0xEF) i = b[0] === 0xEF ? 3 : 2;
    while (i < Math.min(b.length, 400) && (b[i] === 0 || b[i] === 32 || b[i] === 10 || b[i] === 13 || b[i] === 9)) i++;
    return b[i] === 0x3C;
  }
  function sheetsOf(f, buf) {
    return NS.loadJs(XLJS, '').then(function () {
      var X = window.XLSX, csv = /\.(csv|txt)$/i.test(f.name), wb = csv ? X.read(new TextDecoder().decode(buf), { type: 'string', raw: true }) : X.read(buf, { type: 'array', cellDates: false });
      return wb.SheetNames.map(function (n) {
        var rows = X.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' });
        return { name: (ST.files.length > 1 ? f.name.replace(/\.[^.]+$/, '') + ' · ' : '') + (csv ? f.name.replace(/\.[^.]+$/, '') : n), rows: rows.map(function (r) { return r.map(function (c) { return c == null ? '' : String(c).trim(); }); }) };
      });
    });
  }
  function take(files) {
    ST.files = [].slice.call(files || []).filter(function (f) { return /\.(xlsx|xlsm|xls|csv|txt|xml)$/i.test(f.name); });
    if (!ST.files.length) { NS.toast(t('g1f.kinds'), 'bad'); return; }
    ST.book = []; ST.texts = []; ST.force = {}; ST.name = ST.files.map(function (f) { return f.name; }).join(', ');
    msg(t('g1f.reading', { f: ST.name }));
    loadSaved().then(function () {
      return ST.files.reduce(function (p, f) {
        return p.then(function () { return f.arrayBuffer(); }).then(function (buf) {
          var b = new Uint8Array(buf);
          if (isXml(f.name, b)) { ST.texts.push(b); return; }
          return sheetsOf(f, buf).then(function (sh) { ST.book = ST.book.concat(sh); });
        }).catch(function (e) { ST.bad = (ST.bad || []).concat([t('g1f.bad', { f: f.name, e: /password|encrypt/i.test(String(e && e.message)) ? t('g1f.locked') : (e && e.message || e) })]); });
      }, Promise.resolve());
    }).then(work);
  }
  function sample() {
    ST.files = [{ name: 'sample-sales-register.csv' }]; ST.name = t('g1f.samplef'); ST.texts = []; ST.force = {};
    ST.book = [{ name: 'Sample sales register', rows: F.sample() }];
    if (!$('#g1-own').value) $('#g1-own').value = sampleGstin();
    loadSaved().then(work);
  }
  /* a made-up Punjab GSTIN with a correct check letter for the sample */
  function sampleGstin() { var g = '03AAACN1234A1Z', C = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ', s = 0; for (var i = 0; i < 14; i++) { var v = C.indexOf(g[i]) * (i % 2 ? 2 : 1); s += Math.floor(v / 36) + v % 36; } return g + C[(36 - s % 36) % 36]; }

  /* ---------- from rows to bills, checks, then the tables ---------- */
  function work() {
    var g = String($('#g1-own').value || '').toUpperCase().replace(/\s+/g, ''), own = /^\d{2}/.test(g) ? g.slice(0, 2) : '';
    ST.lines = F.read(ST.book, { saved: ST.saved, force: ST.force });
    var T = F.toDocs(ST.lines.lines, { own: own }), docs = T.docs, probs = T.probs.slice(), tl = null;
    if (ST.texts.length) { try { tl = F.fromTally(ST.texts, { own: own }); docs = docs.concat(tl.docs); probs = probs.concat(tl.probs); if (tl.company) ST.company = tl.company; } catch (e) { (ST.bad = ST.bad || []).push(t('g1f.bad', { f: 'XML', e: e && e.message || e })); } }
    ST.tally = tl;
    if (!own) own = (tl && tl.own) || F.guessOwn(docs);
    var seller = { name: ST.company || (g ? t('g1f.me') + ' ' + g : t('g1f.me')), gstin: g, state: own };
    F.setOwn(docs, own, seller);
    docs.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
    ST.E = probs.concat(F.check(docs, { own: own, ownGstin: g, aato: $('#g1-aato').value }));
    ST.docs = docs; ST.own = own;
    drawRead();
    if (docs.length) {
      var ms = F.months(docs);
      NSG1.useFile({ name: ST.name, docs: docs, seller: seller, months: ms, from: docs[0].date, to: lastDay(docs[docs.length - 1].date) });
    }
  }
  function lastDay(d) { var y = +d.slice(0, 4), m = +d.slice(5, 7); return d.slice(0, 8) + ('0' + new Date(y, m, 0).getDate()).slice(-2); }

  function msg(s) { var r = $('#g1-read'); r.hidden = false; r.innerHTML = '<p class="soft" aria-live="polite">' + esc(s) + '</p>'; }
  function where(a) { return a && a[0] ? (a[0] + (a[1] !== '' && a[1] != null ? ' · ' + t('g1f.row', { n: a[1] }) : '')) : ''; }
  function emsg(e) { var v = Object.assign({ no: e.no }, e.v); if (v.p && /^\d{2}$/.test(v.p)) v.p = v.p + '-' + (NS.stateName(v.p) || ''); ['v', 'l', 'c', 's', 't', 'w'].forEach(function (k) { if (typeof v[k] === 'number') v[k] = NS.money(v[k], false); }); ['d', 'o'].forEach(function (k) { if (/^\d{4}-\d\d-\d\d$/.test(v[k] || '')) v[k] = v[k].split('-').reverse().join('-'); }); return t('g1f.e.' + e.k, v); }
  function drawRead() {
    var r = $('#g1-read'), L = ST.lines, h = '';
    (ST.bad || []).forEach(function (b) { h += '<p class="callout g1-warn">' + esc(b) + '</p>'; }); ST.bad = [];
    L.found.forEach(function (f) {
      if (f.kind) h += '<p class="g1f-sh">✔ ' + esc(t('g1f.read', { s: f.name, n: f.rows, r: f.headRow + 1 })) + (f.ledgers && f.ledgers.length ? ' <span class="soft">' + esc(t('g1f.leds', { l: f.ledgers.slice(0, 6).join(', ') + (f.ledgers.length > 6 ? '…' : '') })) + '</span>' : '') + ' <button type="button" class="btn sm" data-g1f="cols" data-s="' + esc(f.name) + '">' + esc(t('g1f.cols')) + '</button></p>';
      else h += '<p class="g1f-sh">✖ ' + esc(t('g1f.notsales', { s: f.name })) + ' <button type="button" class="btn sm" data-g1f="cols" data-s="' + esc(f.name) + '">' + esc(t('g1f.cols')) + '</button></p>';
    });
    if (ST.tally) h += '<p class="g1f-sh">✔ ' + esc(t('g1f.tally', { n: ST.tally.docs.length })) + '</p>';
    var D = ST.docs || [], inv = D.filter(function (d) { return d.type === 'invoice' && !d.cancelled; }).length, cn = D.filter(function (d) { return d.type === 'note' && d.kind === 'credit'; }).length, dn = D.filter(function (d) { return d.type === 'note' && d.kind === 'debit'; }).length, cx = D.filter(function (d) { return d.cancelled; }).length;
    var ms = F.months(D);
    if (D.length) h += '<p><b>' + esc(t('g1f.got', { b: inv, c: cn, d: dn, x: cx, m: ms.length ? (ms.length > 1 ? ml(ms[0]) + '–' + ml(ms[ms.length - 1]) : ml(ms[0])) : '' })) + '</b>' + (ST.own ? ' <span class="soft">' + esc(t('g1f.ownis', { s: ST.own + '-' + (NS.stateName(ST.own) || '') })) + '</span>' : '') + '</p>';
    if (ST.cols) h += colsHtml();
    var E = ST.E, c = F.counts(E);
    if (D.length || E.length) {
      if (!E.length) h += '<p class="callout">' + esc(t('g1f.ok')) + '</p>';
      else {
        var order = { err: 0, warn: 1, info: 2 }, list = E.slice().sort(function (a, b) { return order[a.sev] - order[b.sev]; });
        h += '<div class="g1f-probs"><h3>' + esc(t('g1f.probs')) + '</h3><p>' + esc(t('g1f.errs', { e: c.err, w: c.warn, i: c.info })) + ' <button type="button" class="btn sm" data-g1f="dl">' + esc(t('g1f.dl')) + '</button></p>' +
          '<div class="g1-tw"><table class="g1-t g1f-t"><thead><tr><th></th><th>' + esc(t('g1f.where')) + '</th><th>' + esc(t('g1f.bill')) + '</th><th>' + esc(t('g1f.what')) + '</th></tr></thead><tbody>' +
          list.slice(0, 200).map(function (e) { return '<tr class="sev-' + e.sev + '"><td><span class="g1f-b">' + esc(t('g1f.s.' + e.sev)) + '</span></td><td>' + esc(where(e.at)) + '</td><td>' + esc(e.no || '') + '</td><td class="wrap">' + esc(emsg(e)) + '</td></tr>'; }).join('') + '</tbody></table></div>' +
          (list.length > 200 ? '<p class="soft">' + esc(t('g1f.more', { n: list.length - 200 })) + '</p>' : '') +
          (c.err ? '<p><label><input type="checkbox" id="g1-ack"> ' + esc(t('g1f.ack')) + '</label></p>' : '') + '</div>';
      }
    }
    if (D.length) h += '<p><button type="button" class="btn" data-g1f="saved">' + esc(t('g1f.saved')) + '</button></p>';
    r.hidden = !h; r.innerHTML = h;
  }
  /* the portal file and the template are made only when nothing is left to correct, or the user says they checked */
  NSG1.gate = function () {
    if (!NSG1.file()) return true;
    var c = F.counts(ST.E || []); if (!c.err) return true;
    var a = $('#g1-ack'); if (a && a.checked) return true;
    NS.toast(t('g1f.fixfirst'), 'bad'); var p = $('.g1f-probs'); if (p) p.scrollIntoView({ block: 'start', behavior: 'smooth' });
    return false;
  };
  function dl() {
    var rows = [[t('g1f.where'), t('g1f.bill'), t('g1f.sev'), t('g1f.what')]].concat(ST.E.map(function (e) { return [where(e.at), e.no || '', t('g1f.s.' + e.sev), emsg(e)]; }));
    var body = '﻿' + rows.map(function (r) { return r.map(function (v) { v = String(v == null ? '' : v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }).join(','); }).join('\r\n');
    var a = document.createElement('a'), u = URL.createObjectURL(new Blob([body], { type: 'text/csv;charset=utf-8' }));
    a.href = u; a.download = 'GSTR-1 file check.csv'; document.body.appendChild(a); a.click(); a.remove(); setTimeout(function () { URL.revokeObjectURL(u); }, 4000);
  }
  /* a blank sheet with every column NSBiz reads */
  function blank() {
    NS.loadJs(XLJS, '').then(function () {
      var X = window.XLSX, wb = X.utils.book_new(), H = F.SET.cols.map(function (c) { return c.h; });
      var row = F.SET.cols.map(function (c) { return { no: 'INV/001', date: '05-04-2026', party: 'Garg Traders', pos: 'Punjab', type: 'Sales', hsn: '7306', desc: 'MS Pipe', qty: '10', uqc: 'KGS', price: '100', txv: '1000', rate: '18', val: '1180', rc: 'N' }[c.k] || ''; });
      var ws = X.utils.aoa_to_sheet([H.concat(['IGST', 'CGST', 'SGST', 'Cess']), row.concat(['', '90', '90', ''])]);
      X.utils.book_append_sheet(wb, ws, 'Sales register');
      X.writeFile(wb, 'GSTR-1 sales register (blank).xlsx');
    });
  }

  /* ---------- the columns screen ---------- */
  function sheetOf(n) { return ST.book.filter(function (b) { return b.name === n; })[0]; }
  function colsHtml() {
    var sh = sheetOf(ST.cols); if (!sh) return '';
    var f = ST.lines.found.filter(function (x) { return x.name === ST.cols; })[0] || {}, L = ST.force[ST.cols] || f, row = L.row != null ? L.row : (L.headRow || 0), map = L.map || {}, heads = sh.rows[row] || [];
    var h = '<div class="card g1f-cols"><h3>' + esc(t('g1f.colsof', { s: ST.cols })) + '</h3><p class="soft">' + esc(t('g1f.colsh')) + '</p>' +
      '<p><label>' + esc(t('g1f.hrow')) + ' <input type="number" min="1" max="15" value="' + (row + 1) + '" data-g1f-hrow style="width:5em"></label></p><div class="g1f-grid">';
    F.SET.cols.forEach(function (c) {
      h += '<label>' + esc(c.h) + (c.need ? ' *' : '') + '<select data-g1f-col="' + c.k + '"><option value="">' + esc(t('g1f.nonec')) + '</option>' +
        heads.map(function (hd, i) { return '<option value="' + i + '"' + (map[c.k] === i ? ' selected' : '') + '>' + esc(colName(i) + ': ' + (hd || '—')) + '</option>'; }).join('') + '</select></label>';
    });
    return h + '</div><p class="soft">' + esc(t('g1f.colsled')) + '</p><div class="row"><button type="button" class="btn primary" data-g1f="colsok">' + esc(t('g1f.colsok')) + '</button> <button type="button" class="btn" data-g1f="colsx">' + esc(t('g1f.close')) + '</button></div></div>';
  }
  function colName(i) { var s = ''; i++; while (i) { s = String.fromCharCode(65 + (i - 1) % 26) + s; i = Math.floor((i - 1) / 26); } return s; }
  function colsSave() {
    var sh = sheetOf(ST.cols); if (!sh) return;
    var row = Math.max(0, (+$('[data-g1f-hrow]', box).value || 1) - 1), map = {};
    $$('[data-g1f-col]', box).forEach(function (s) { if (s.value !== '') map[s.getAttribute('data-g1f-col')] = +s.value; });
    var heads = sh.rows[row] || [], sig = NSTpl.sig(heads);
    ST.force[ST.cols] = { kind: 'g1sales', row: row, map: map, sig: sig, heads: heads };
    remember(sig, map); ST.cols = null; work();
  }

  /* ---------- events ---------- */
  var fin = $('#g1-fin'), dz = $('.g1-dz', box);
  fin.addEventListener('change', function () { take(fin.files); fin.value = ''; });
  ['dragenter', 'dragover'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add('over'); }); });
  ['dragleave', 'drop'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.remove('over'); }); });
  dz.addEventListener('drop', function (e) { if (e.dataTransfer && e.dataTransfer.files) take(e.dataTransfer.files); });
  box.addEventListener('click', function (e) {
    var b = e.target.closest('[data-g1f]'); if (!b) return;
    var k = b.getAttribute('data-g1f');
    if (k === 'sample') sample();
    else if (k === 'blank') blank();
    else if (k === 'dl') dl();
    else if (k === 'cols') { ST.cols = b.getAttribute('data-s'); drawRead(); var c = $('.g1f-cols', box); if (c) c.scrollIntoView({ block: 'nearest' }); }
    else if (k === 'colsok') colsSave();
    else if (k === 'colsx') { ST.cols = null; drawRead(); }
    else if (k === 'saved') { ST.docs = []; ST.E = []; ST.book = []; ST.texts = []; $('#g1-read').hidden = true; NSG1.useSaved(); }
  });
  $('#g1-own').addEventListener('change', function () { if (ST.book.length || ST.texts.length) work(); });
  $('#g1-aato').addEventListener('change', function () { if (NSG1.file()) { ST.E = ST.E.filter(function (x) { return x.k === 'nono' || x.k === 'nodate' || x.k === 'tallynone'; }).concat(F.check(ST.docs, { own: ST.own, ownGstin: String($('#g1-own').value || '').toUpperCase().trim(), aato: $('#g1-aato').value })); drawRead(); } });
  document.addEventListener('ns-lang', function () { if (ST.docs && ST.docs.length) drawRead(); });
  try { var og = localStorage.getItem('ns-g1-own'); if (og && !$('#g1-own').value) $('#g1-own').value = og; } catch (e) {}
  $('#g1-own').addEventListener('input', function () { try { localStorage.setItem('ns-g1-own', this.value.trim()); } catch (e) {} });
})();
