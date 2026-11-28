/* Three-way match screen (byof3-einv-gstr1, 28 Nov 2026) on gstr1.niveshsahayak.in/einvoice-gstr1-books-check/: drop the e-invoice
   download, GSTR-1 and the sales register -> what was read -> documents missing in one of them or different in value, tax, GSTIN or date,
   with an Excel download. Reading and matching are NSE3 (e3core.js). Files are read in this browser only; nothing is sent or saved,
   except the column choices (localStorage ns-e3-lay). No "Save to NSBiz": the result is a list to correct in your own software. */
(function () {
  'use strict';
  var NS = window.NS, box = NS && NS.$ && NS.$('#e3'); if (!box || !window.NSE3) return;
  var $ = NS.$, $$ = NS.$$, esc = NS.esc, E3 = window.NSE3, XLJS = '/assets/vendor/xlsx-0.20.3.full.min.js';
  function t(k, v, en) { var s = NS.t(k, v); if (s === k && en) { s = en; Object.keys(v || {}).forEach(function (x) { s = s.split('{' + x + '}').join(v[x]); }); } return s; }
  var SRC = { e: 'E-invoices', g: 'GSTR-1', b: 'Books' };
  var ST = { e: null, g: null, b: null, res: null, filter: 'todo', cols: null, saved: {}, force: { e: {}, b: {} }, show: 200 };
  try { ST.saved = JSON.parse(localStorage.getItem('ns-e3-lay') || '{}'); } catch (e) {}

  /* ---------- CSS (also for the translated page copies, which carry only the box) ---------- */
  var css = document.createElement('style');
  css.textContent = '.e3{background:var(--surface);border:1px solid var(--line);border-radius:var(--r);padding:16px;margin:0 0 18px}.e3 h2{margin:0 0 6px;font-size:1.2rem}' +
    '.e3-g{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:12px;margin:10px 0}.e3-c{border:1px dashed var(--line);border-radius:var(--r);padding:12px;min-width:0}' +
    '.e3-c.over{border-color:var(--accent);background:color-mix(in srgb,var(--accent) 8%,transparent)}.e3-c h3{margin:0 0 4px;font-size:1rem}.e3-c small{display:block;color:var(--muted);margin:4px 0}' +
    '.e3-c input[type=file]{max-width:100%}.e3-st{font-size:.92rem;margin-top:6px;overflow-wrap:anywhere}.e3-row{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:8px 0}' +
    '.e3-row label{display:flex;flex-wrap:wrap;gap:6px;align-items:center}.e3-row input[type=text]{min-width:0;width:16em;max-width:100%}' +
    '.e3-chips button[aria-pressed=true]{outline:2px solid var(--accent)}.e3-tw{overflow-x:auto;border:1px solid var(--line);border-radius:var(--r)}' +
    '.e3-t{border-collapse:collapse;width:100%;font-size:.92rem}.e3-t th,.e3-t td{border-bottom:1px solid var(--line);padding:6px 8px;text-align:left;vertical-align:top}' +
    '.e3-t td.w{min-width:16em}.e3-r td{white-space:nowrap}.e3-r td.w{white-space:normal}.e3-lb{display:none}.e3-b{display:inline-block;border-radius:99px;padding:1px 8px;font-size:.8rem;border:1px solid var(--line);white-space:nowrap}' +
    '.e3-miss .e3-b{background:#fde2e2;color:#8a1c1c}.e3-diff .e3-b{background:#fff1cc;color:#6b4b00}.e3-ok .e3-b{background:#dff5e3;color:#1f5f2c}.e3-cnl .e3-b{background:#e8e8ee;color:#333}' +
    '.e3-y{color:#1f7a35;font-weight:700}.e3-n{color:#b42318;font-weight:700}.e3-cols{border:1px solid var(--line);border-radius:var(--r);padding:12px;margin:10px 0}' +
    '.e3-cg{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:8px}.e3-cg label{display:flex;flex-direction:column;gap:2px;font-size:.9rem}.e3-cg select{max-width:100%}' +
    '@media (max-width:640px){.e3-r thead{display:none}.e3-r,.e3-r tbody,.e3-r tr,.e3-r td{display:block;width:auto}.e3-r tr{padding:8px 10px;border-bottom:1px solid var(--line)}.e3-r td{border:0;padding:2px 0;white-space:normal}.e3-r td.w{min-width:0}.e3-lb{display:inline;color:var(--muted)}}';
  document.head.appendChild(css);

  /* ---------- reading files ---------- */
  function isJson(f, b) { if (/\.json$/i.test(f.name)) return true; var i = b[0] === 0xEF ? 3 : 0; while (i < 200 && /\s/.test(String.fromCharCode(b[i] || 33))) i++; return b[i] === 0x7B || b[i] === 0x5B; }
  function isXml(f, b) { if (/\.xml$/i.test(f.name)) return true; var i = 0; if (b[0] === 0xFF || b[0] === 0xFE || b[0] === 0xEF) i = b[0] === 0xEF ? 3 : 2; while (i < 400 && (b[i] === 0 || b[i] === 32 || b[i] === 10 || b[i] === 13 || b[i] === 9)) i++; return b[i] === 0x3C; }
  function sheetsOf(f, buf, many) {
    return NS.loadJs(XLJS, '').then(function () {
      var X = window.XLSX, csv = /\.(csv|txt)$/i.test(f.name), wb = csv ? X.read(new TextDecoder().decode(buf), { type: 'string', raw: true }) : X.read(buf, { type: 'array' });
      return wb.SheetNames.map(function (n) {
        var rows = X.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' });
        return { name: (many ? f.name.replace(/\.[^.]+$/, '') + ' · ' : '') + (csv ? f.name.replace(/\.[^.]+$/, '') : n), rows: rows.map(function (r) { return r.map(function (c) { return c == null ? '' : typeof c === 'number' ? c : String(c).trim(); }); }) };
      });
    });
  }
  function load(s, files) {
    files = [].slice.call(files || []);
    if (!files.length) return;
    var inp = { json: [], book: [], texts: [], names: files.map(function (f) { return f.name; }), bad: [] };
    stat(s, t('e3.reading', { f: inp.names.join(', ') }, 'Reading {f}…'));
    files.reduce(function (p, f) {
      return p.then(function () { return f.arrayBuffer(); }).then(function (buf) {
        var b = new Uint8Array(buf);
        if (isJson(f, b)) { inp.json.push({ name: f.name, data: JSON.parse(new TextDecoder().decode(b).replace(/^﻿/, '')) }); return; }
        if (isXml(f, b)) { if (s === 'b') inp.texts.push(b); else inp.bad.push(t('e3.xmlhere', { f: f.name }, '{f}: an XML file goes in the Books box (Tally Day Book).')); return; }
        return sheetsOf(f, buf, files.length > 1).then(function (sh) { inp.book = inp.book.concat(sh); });
      }).catch(function (e) { inp.bad.push(f.name + ': ' + (/password|encrypt/i.test(String(e && e.message)) ? t('e3.locked', null, 'the file is locked with a password; save a copy without it.') : t('e3.unread', null, 'could not be read') + ' (' + (e && e.message || e) + ')')); });
    }, Promise.resolve()).then(function () { ST[s] = { inp: inp }; ST.force[s] = {}; run(); });
  }
  function own() { return String($('#e3-own').value || '').toUpperCase().replace(/\s+/g, ''); }
  function readOne(s) {
    var x = ST[s]; if (!x) return;
    var o = { own: own(), saved: ST.saved, force: ST.force[s] };
    x.r = s === 'e' ? E3.readEinv(x.inp, o) : s === 'g' ? E3.readG1(x.inp, o) : E3.readBooks(x.inp, o);
  }
  function run() {
    ['e', 'g', 'b'].forEach(readOne);
    if (!own() && ST.g && ST.g.r && ST.g.r.gstin) { $('#e3-own').value = ST.g.r.gstin; ['e', 'b'].forEach(readOne); }
    var have = ['e', 'g', 'b'].filter(function (s) { return ST[s] && ST[s].r; });
    ST.res = have.length >= 2 ? E3.compare(ST.e && ST.e.r.recs, ST.g && ST.g.r.recs, ST.b && ST.b.r.recs, { own: own() }) : null;
    ST.show = 200;
    ['e', 'g', 'b'].forEach(drawStat); draw();
  }

  /* ---------- drawing ---------- */
  function stat(s, html, raw) { var el = $('#e3-st-' + s); el.innerHTML = raw ? html : esc(html); }
  var WHY = { nono: 'no document number', nodate: 'date not understood', nojson: 'no e-invoice found in this JSON', noamt: 'no amount', noreader: 'the reader did not load (reload the page)', xml: 'Tally XML could not be read', chk: 'GSTIN check letter wrong (kept)' };
  function drawStat(s) {
    var x = ST[s]; if (!x) return;
    var r = x.r || { recs: [], found: [], skipped: [] }, h = '';
    (x.inp.bad || []).forEach(function (b) { h += '<p class="callout g1-warn">' + esc(b) + '</p>'; });
    r.found.forEach(function (f) {
      h += '<div>' + (f.ok ? '✔ ' : '✖ ') + esc(f.name) + ': ' + esc(f.ok ? t('e3.got', { n: f.rows }, '{n} rows read') : t('e3.nohead', null, 'no document number and date columns found')) +
        (f.heads && s !== 'g' ? ' <button type="button" class="btn sm" data-e3="cols" data-s="' + s + '" data-n="' + esc(f.name) + '">' + esc(t('e3.cols', null, 'Columns…')) + '</button>' : '') + '</div>';
    });
    if (s === 'g' && r.summary && r.summary.length) h += '<small>' + esc(t('e3.sumrows', { n: r.summary.length }, '{n} summary rows (B2C Small, nil, advances) have no document numbers and are not compared one by one.')) + '</small>';
    if (s === 'g' && r.amend) h += '<small>' + esc(t('e3.amend', { n: r.amend }, '{n} amendment rows were left out.')) + '</small>';
    if (r.skipped.length) h += '<details><summary>' + esc(t('e3.skipped', { n: r.skipped.length }, '{n} rows left out')) + '</summary><ul>' + r.skipped.slice(0, 50).map(function (k) { return '<li>' + esc(k.at + ': ' + (t('e3.why.' + k.why, null, WHY[k.why] || k.why)) + (k.v ? ' (' + k.v + ')' : '')) + '</li>'; }).join('') + '</ul></details>';
    stat(s, h, true);
  }
  function money(v) { return NS.money ? NS.money(v, false) : (+v).toFixed(2); }
  function dmy(d) { return d ? d.split('-').reverse().join('-') : ''; }
  var KIND = { inv: 'Invoice', cn: 'Credit note', dn: 'Debit note' };
  function draw() {
    var out = $('#e3-res'), R = ST.res;
    if (!R) { out.innerHTML = '<p class="soft">' + esc(t('e3.need2', null, 'Add at least two of the three files to compare them.')) + '</p>' + (ST.cols ? colsHtml() : ''); return; }
    var c = R.counts, h = ST.cols ? colsHtml() : '';
    h += '<p>' + esc(t('e3.scope', { m: R.scope.map(ml).join(', ') || '-' }, 'Months compared (the ones every file covers): {m}.')) + '</p>';
    var left = ['e', 'g', 'b'].filter(function (s) { return R.outside[s]; }).map(function (s) { return R.outside[s] + ' ' + t('e3.s.' + s, null, SRC[s]); });
    if (left.length) h += '<p class="soft">' + esc(t('e3.outside', { l: left.join(', ') }, 'Left out because they fall in other months: {l}.')) + '</p>';
    if (!R.scope.length) h += '<p class="callout g1-warn">' + esc(t('e3.nocommon', null, 'The files have no month in common. Choose files of the same month or quarter.')) + '</p>';
    var todo = c.miss + c.diff;
    h += '<p class="callout">' + esc(todo ? t('e3.sumbad', { m: c.miss, d: c.diff, o: c.ok }, '{m} missing in a file, {d} with differences, {o} matching.') : t('e3.sumok', { o: c.ok }, 'All {o} documents match. Check the totals below once more before you file.')) + '</p>';
    h += '<div class="e3-row e3-chips" role="group">' + [['todo', 'To check (' + todo + ')'], ['miss', 'Missing (' + c.miss + ')'], ['diff', 'Different (' + c.diff + ')'], ['ok', 'Matching (' + c.ok + ')'], ['cnl', 'Cancelled e-invoices (' + c.cnl + ')'], ['all', 'All']].map(function (b) {
      return '<button type="button" class="btn sm" data-noic data-e3="f" data-f="' + b[0] + '" aria-pressed="' + (ST.filter === b[0]) + '">' + esc(t('e3.f.' + b[0], null, b[1].replace(/ \(.*/, '')) + (b[1].match(/ \(.*/) || [''])[0]) + '</button>';
    }).join('') + ' <button type="button" class="btn primary sm" data-e3="xl">' + esc(t('e3.xl', null, 'Download Excel')) + '</button></div>';
    var rows = R.rows.filter(function (r) { return ST.filter === 'all' || (ST.filter === 'todo' ? r.status === 'miss' || r.status === 'diff' : r.status === ST.filter); });
    var have = ['e', 'g', 'b'].filter(function (s) { return R.have[s]; });
    h += '<div class="e3-tw"><table class="e3-t e3-r"><thead><tr><th>' + esc(t('e3.h.st', null, 'Status')) + '</th><th>' + esc(t('e3.h.no', null, 'Document')) + '</th><th>' + esc(t('e3.h.party', null, 'Buyer')) + '</th>' +
      have.map(function (s) { return '<th>' + esc(t('e3.s.' + s, null, SRC[s])) + '</th>'; }).join('') + '<th>' + esc(t('e3.h.what', null, 'What to check')) + '</th></tr></thead><tbody>' +
      rows.slice(0, ST.show).map(function (r) {
        return '<tr class="e3-' + r.status + '"><td><span class="e3-b">' + esc(t('e3.st.' + r.status, null, { ok: 'Matches', diff: 'Differs', miss: 'Missing', cnl: 'Cancelled' }[r.status])) + '</span></td>' +
          '<td>' + esc(r.no) + '<br><small>' + esc(t('e3.k.' + r.kind, null, KIND[r.kind]) + ' · ' + dmy(r.date)) + '</small></td><td>' + esc(r.party || '') + (r.gstin ? '<br><small>' + esc(r.gstin) + '</small>' : '') + '</td>' +
          have.map(function (s) { var q = r[s]; return '<td><span class="e3-lb">' + esc(t('e3.s.' + s, null, SRC[s])) + ': </span>' + (q ? '<span class="e3-y" aria-label="yes">✔</span> ' + esc(money(q.txv)) + (q.cnl ? ' <small>(' + esc(t('e3.cnl', null, 'cancelled')) + ')</small>' : '') : '<span class="e3-n" aria-label="no">✖</span>') + '</td>'; }).join('') +
          '<td class="w">' + esc(r.what.map(function (w) { return E3.say(w, function (k, v) { return NS.t(k, v); }, money); }).join(' ')) + '</td></tr>';
      }).join('') + '</tbody></table></div>';
    if (!rows.length) h += '<p class="soft">' + esc(t('e3.none', null, 'Nothing in this list.')) + '</p>';
    if (rows.length > ST.show) h += '<p><button type="button" class="btn sm" data-e3="more">' + esc(t('e3.more', { n: rows.length - ST.show }, 'Show {n} more')) + '</button></p>';
    if (R.dups.length) h += '<details><summary>' + esc(t('e3.dups', { n: R.dups.length }, 'The same number is used twice inside one file: {n}')) + '</summary><ul>' + R.dups.slice(0, 50).map(function (d) { return '<li>' + esc(t('e3.s.' + d.src, null, SRC[d.src]) + ': ' + d.no + ' (' + d.at + '; ' + d.other + ')') + '</li>'; }).join('') + '</ul></details>';
    h += '<h3>' + esc(t('e3.tot', null, 'Totals of the months compared')) + '</h3><div class="e3-tw"><table class="e3-t"><thead><tr><th></th><th>' + esc(t('e3.h.n', null, 'Documents')) + '</th><th>' + esc(t('e3.h.txv', null, 'Taxable value')) + '</th><th>' + esc(t('e3.h.tax', null, 'Tax')) + '</th></tr></thead><tbody>' +
      have.map(function (s) { var o = R.totals[s]; return '<tr><td>' + esc(t('e3.s.' + s, null, SRC[s])) + '</td><td>' + o.n + '</td><td>' + esc(money(o.txv)) + '</td><td>' + esc(money(o.tax)) + '</td></tr>'; }).join('') + '</tbody></table></div>' +
      '<p class="soft">' + esc(t('e3.totnote', null, 'Credit notes are taken off. GSTR-1 totals here count only the documents listed one by one (B2B, B2C Large, exports, notes), not B2C Small; books count every bill. E-invoices cover only registered buyers, exports and SEZ.')) + '</p>';
    out.innerHTML = h;
  }
  function ml(m) { var M = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']; return M[+m.slice(5, 7) - 1] + ' ' + m.slice(0, 4); }

  /* ---------- the Columns screen (e-invoice list and books sheets) ---------- */
  function colName(i) { var s = ''; i++; while (i) { s = String.fromCharCode(65 + (i - 1) % 26) + s; i = Math.floor((i - 1) / 26); } return s; }
  function colsHtml() {
    var c = ST.cols, x = ST[c.s], sh = x && x.inp.book.filter(function (b) { return b.name === c.n; })[0]; if (!sh) return '';
    var f = (x.r.found || []).filter(function (y) { return y.name === c.n; })[0] || {}, L = ST.force[c.s][c.n] || f, row = L.row != null ? L.row : (L.headRow || 0), map = L.map || {}, heads = sh.rows[row] || [];
    var cols = c.s === 'e' ? E3.ESET.cols : (window.NSG1File ? NSG1File.SET.cols : []);
    return '<div class="e3-cols"><h3>' + esc(t('e3.colsof', { s: c.n }, 'Columns of {s}')) + '</h3><p><label>' + esc(t('e3.hrow', null, 'Heading row')) + ' <input type="number" min="1" max="15" value="' + (row + 1) + '" data-e3-hrow style="width:5em"></label></p><div class="e3-cg">' +
      cols.map(function (k) { return '<label>' + esc(k.h) + (k.need ? ' *' : '') + '<select data-e3-col="' + k.k + '"><option value="">-</option>' + heads.map(function (hd, i) { return '<option value="' + i + '"' + (map[k.k] === i ? ' selected' : '') + '>' + esc(colName(i) + ': ' + (hd === '' ? '-' : hd)) + '</option>'; }).join('') + '</select></label>'; }).join('') +
      '</div><div class="e3-row"><button type="button" class="btn primary" data-e3="colsok">' + esc(t('e3.colsok', null, 'Use these columns')) + '</button><button type="button" class="btn" data-e3="colsx">' + esc(t('e3.close', null, 'Close')) + '</button></div></div>';
  }
  function colsSave() {
    var c = ST.cols, sh = ST[c.s].inp.book.filter(function (b) { return b.name === c.n; })[0];
    var row = Math.max(0, (+$('[data-e3-hrow]', box).value || 1) - 1), map = {};
    $$('[data-e3-col]', box).forEach(function (s) { if (s.value !== '') map[s.getAttribute('data-e3-col')] = +s.value; });
    var heads = sh.rows[row] || [], sig = NSTpl.sig(heads), kind = c.s === 'e' ? 'e3einv' : 'g1sales';
    ST.force[c.s][c.n] = { kind: kind, row: row, map: map, sig: sig, heads: heads };
    ST.saved[sig] = { kind: kind, map: map, at: Date.now() };
    try { localStorage.setItem('ns-e3-lay', JSON.stringify(ST.saved)); } catch (e) {}
    ST.cols = null; run();
  }

  /* ---------- Excel ---------- */
  /* Excel out: extras.js NS.xlsx(file, [{ name, rows }]) on this site (SheetJS); xlsx.js NS.xlsx.write where that is the one on the page */
  function xl() {
    var R = ST.res, sh = E3.sheets(R, function (k, v) { return NS.t(k, v); }), file = 'E-invoice, GSTR-1 and books check' + (R.scope.length ? ' ' + R.scope[0] + (R.scope.length > 1 ? ' to ' + R.scope[R.scope.length - 1] : '') : '') + '.xlsx';
    if (typeof NS.xlsx === 'function') return NS.xlsx(file, sh.map(function (x) { return { name: x[0], rows: x[1] }; }));
    var blob = NS.xlsx.write(sh), a = document.createElement('a'), u = URL.createObjectURL(blob);
    a.href = u; a.download = file; document.body.appendChild(a); a.click(); a.remove(); setTimeout(function () { URL.revokeObjectURL(u); }, 4000);
  }
  function sample() {
    var S = E3.sample();
    $('#e3-own').value = S.own;
    ST.e = { inp: { json: [{ name: 'sample-e-invoices.json', data: S.einv }], book: [], texts: [], names: ['sample-e-invoices.json'], bad: [] } };
    ST.g = { inp: { json: [{ name: 'sample-GSTR-1.json', data: S.gstr1 }], book: [], texts: [], names: ['sample-GSTR-1.json'], bad: [] } };
    ST.b = { inp: { json: [], book: [{ name: 'Sample sales register', rows: S.books }], texts: [], names: ['sample-sales-register.csv'], bad: [] } };
    ST.force = { e: {}, b: {} }; run();
    NS.toast && NS.toast(t('e3.sampled', null, 'Sample files loaded (made-up data with planted differences).'));
  }

  /* ---------- events ---------- */
  ['e', 'g', 'b'].forEach(function (s) {
    var c = $('#e3-c-' + s), fin = $('#e3-f-' + s);
    fin.addEventListener('change', function () { load(s, fin.files); fin.value = ''; });
    ['dragenter', 'dragover'].forEach(function (ev) { c.addEventListener(ev, function (e) { e.preventDefault(); c.classList.add('over'); }); });
    ['dragleave', 'drop'].forEach(function (ev) { c.addEventListener(ev, function (e) { e.preventDefault(); c.classList.remove('over'); }); });
    c.addEventListener('drop', function (e) { if (e.dataTransfer && e.dataTransfer.files) load(s, e.dataTransfer.files); });
  });
  box.addEventListener('click', function (e) {
    var b = e.target.closest('[data-e3]'); if (!b) return;
    var k = b.getAttribute('data-e3');
    if (k === 'sample') sample();
    else if (k === 'clear') { ST.e = ST.g = ST.b = null; ST.res = null; ['e', 'g', 'b'].forEach(function (s) { stat(s, ''); }); draw(); }
    else if (k === 'f') { ST.filter = b.getAttribute('data-f'); ST.show = 200; draw(); }
    else if (k === 'more') { ST.show += 500; draw(); }
    else if (k === 'xl') xl();
    else if (k === 'cols') { ST.cols = { s: b.getAttribute('data-s'), n: b.getAttribute('data-n') }; draw(); var cc = $('.e3-cols', box); if (cc) cc.scrollIntoView({ block: 'nearest' }); }
    else if (k === 'colsok') colsSave();
    else if (k === 'colsx') { ST.cols = null; draw(); }
  });
  $('#e3-own').addEventListener('change', function () { try { localStorage.setItem('ns-g1-own', this.value.trim()); } catch (e) {} if (ST.e || ST.g || ST.b) run(); });
  try { var og = localStorage.getItem('ns-g1-own'); if (og && !$('#e3-own').value) $('#e3-own').value = og; } catch (e) {}
  document.addEventListener('ns-lang', function () { if (ST.res) draw(); });
  draw();
})();
