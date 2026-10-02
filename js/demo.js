/*
 * "Explain the change" sales demo. Classic script so it runs from file://.
 * Contract (consumed by the video pipeline, do not change):
 *   ?step=0-6  ?prospect=meridian|harborline  ?capture=1
 *   window.demo = { goTo(n), steps: 7, prospect }
 *   document.body.dataset.step mirrors the current step; Left/Right arrows navigate.
 * Additions: window.demo.step, next(), prev(), setProspect(id), settled() and
 *   document.body.dataset.settled ("true" once the step transition has finished).
 */
(function () {
  'use strict';

  var DATA = window.BRAINCO_DEMO_DATA;
  var STEP_COUNT = 7;
  var MINUS = '\u2212';
  var params = new URLSearchParams(window.location.search);
  var capture = params.get('capture') === '1';
  var prospectId = DATA.prospects[params.get('prospect')] ? params.get('prospect') : DATA.defaultProspect;
  var current = clampStep(params.get('step'));
  var P = DATA.prospects[prospectId];
  var settlePromise = Promise.resolve();
  var settleToken = 0;

  var LEGEND = {
    today: { label: 'Exists today', note: 'Built in a current Brain Co. engagement' },
    proposed: { label: 'Proposed', note: 'Part of the recommended product' },
    assumption: { label: 'Assumption', note: 'Needs validation in the pilot' }
  };

  var STEP_LABELS = ['Overview', 'Diff', 'Explain', 'Draft', 'Review', 'Approve', 'Audit'];

  /* ---------- formatting ---------- */

  function clampStep(v) {
    var n = parseInt(v, 10);
    if (isNaN(n)) return 0;
    return Math.max(0, Math.min(STEP_COUNT - 1, n));
  }
  function esc(s) {
    return String(s).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  }
  function pct(v) { return v.toFixed(1) + '%'; }
  function money(v) { return '$' + v.toFixed(1) + 'M'; }
  function moneyLong(v) { return '$' + v.toFixed(1) + ' million'; }
  function num(v) { return v.toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 }); }
  function mult(v) { return v.toFixed(2) + 'x'; }
  function signed(v, digits) {
    var d = digits === undefined ? 1 : digits;
    var s = Math.abs(v).toFixed(d);
    return (v < 0 ? MINUS : v > 0 ? '+' : '') + s;
  }
  function bps(d) { return (d < 0 ? MINUS : d > 0 ? '+' : '') + Math.abs(d) + ' bps'; }
  function words(text) { return text.split(/\s+/).filter(Boolean).length; }

  /* ---------- derived data ---------- */

  function derive(v) {
    return {
      tvpi: (v.distributions + v.nav) / v.paidIn,
      tvpiPrior: (v.distributions + v.navPrior) / v.paidIn,
      dpi: v.distributions / v.paidIn,
      navPct: (v.nav / v.navPrior - 1) * 100,
      dBps: Math.round((v.irr - v.irrPrior) * 100)
    };
  }
  function focusVehicle() {
    return P.vehicles.filter(function (v) { return v.id === P.focus.vehicleId; })[0];
  }
  function otherVehicles() {
    return P.vehicles.filter(function (v) { return v.id !== P.focus.vehicleId; });
  }
  function excelRow() { return 5 + P.vehicles.indexOf(focusVehicle()); }
  function excelCell() { return 'J' + excelRow(); }
  function claimText(c, version) {
    if (!c.edit) return c.text;
    return c.text.replace('{edit}', version === 'v1' ? c.edit.before : c.edit.after);
  }
  function commentary(version) {
    return P.claims.map(function (c) { return claimText(c, version); }).join(' ');
  }
  function editedClaim() {
    return P.claims.filter(function (c) { return c.id === P.edit.claimId; })[0];
  }
  function workbookFile() { return P.shortFirm + '_' + P.period.label.replace(' ', '-') + '_LP_Schedules.xlsx'; }
  function scheduleSheet() { return P.package.workbookTabs[0]; }
  function priorVersion() { return P.focus.prior.version.replace('approved ', ''); }

  function sources() {
    var v = focusVehicle();
    var m = P.focus.model;
    var s = {
      M1: { kind: 'model', short: m.cell, title: m.file + ' › ' + m.sheet + '!' + m.cell, detail: P.focus.metric + ', ' + P.period.label + ': ' + pct(v.irr) },
      P1: { kind: 'prior', short: P.period.prior.split(' ')[0] + ' ' + priorVersion(), title: P.focus.prior.file + ', ' + P.focus.prior.version, detail: v.short + ' net IRR at ' + P.period.priorAsOf + ': ' + pct(v.irrPrior) }
    };
    P.holdings.forEach(function (h) {
      s[h.id] = { kind: 'model', short: h.cell.split('!')[1], title: m.file + ' › ' + h.cell, detail: h.name + ' fair value ' + money(h.current) + ' (' + P.period.prior.split(' ')[0] + ' ' + money(h.prior) + ', ' + signed((h.current / h.prior - 1) * 100) + '%)' };
    });
    P.recording.lines.forEach(function (l) {
      s[l.id] = { kind: 'voice', short: l.id, title: 'Voice explanation, ' + P.people.partner.name + ', ' + l.t, detail: l.text };
    });
    return s;
  }

  /* ---------- small components ---------- */

  var ICON = {
    model: '<svg viewBox="0 0 12 12" aria-hidden="true"><rect x="1" y="1.5" width="10" height="9" rx="1.4" fill="none" stroke="currentColor" stroke-width="1.1"/><path d="M1 4.6h10M4.6 1.5v9" stroke="currentColor" stroke-width="1.1"/></svg>',
    voice: '<svg viewBox="0 0 12 12" aria-hidden="true"><path d="M2 5v2M4 3.5v5M6 2v8M8 3.5v5M10 5v2" stroke="currentColor" stroke-width="1.2" stroke-linecap="round"/></svg>',
    prior: '<svg viewBox="0 0 12 12" aria-hidden="true"><circle cx="6" cy="6" r="4.6" fill="none" stroke="currentColor" stroke-width="1.1"/><path d="M6 3.4V6l1.8 1.2" fill="none" stroke="currentColor" stroke-width="1.1" stroke-linecap="round"/></svg>',
    check: '<svg viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 6.3l2.3 2.3 4.7-5" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    arrow: '<svg viewBox="0 0 16 10" aria-hidden="true"><path d="M1 5h13M10 1.5L14 5l-4 3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    flag: '<svg viewBox="0 0 12 12" aria-hidden="true"><path d="M3 10.5V1.8M3 2h6l-1.4 2.2L9 6.4H3" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linejoin="round"/></svg>'
  };

  function chip(kind, why) {
    var l = LEGEND[kind];
    var title = why ? l.label + ': ' + why : l.label + ': ' + l.note;
    return '<span class="chip chip-' + kind + '" data-legend="' + kind + '" title="' + esc(title) + '" aria-label="' + esc(title) + '"><i></i>' + l.label + '</span>';
  }
  function cite(id, srcs) {
    var s = srcs[id];
    return '<span class="cite cite-' + s.kind + '" title="' + esc(s.title + ': ' + s.detail) + '">' + ICON[s.kind] + esc(s.short) + '</span>';
  }
  function avatar(person, cls) {
    return '<span class="avatar ' + (cls || '') + '" title="' + esc(person.name + ', ' + person.role) + '">' + esc(person.initials) + '</span>';
  }
  function panel(o) {
    return '<section class="panel ' + (o.cls || '') + '" data-element="' + o.id + '">' +
      '<header class="panel-head"><div class="panel-titles">' +
      (o.kicker ? '<div class="kicker">' + o.kicker + '</div>' : '') +
      '<h3>' + o.title + '</h3></div>' + chip(o.legend, o.why) + '</header>' +
      '<div class="panel-body">' + o.body + '</div></section>';
  }
  function frame(kind, state, caption) {
    var html = kind === 'deck' ? deckHTML(state) : kind === 'excel' ? excelHTML(state) : letterHTML(state);
    var scrollRegion = kind === 'excel' && state === 'approved' ? ' tabindex="0" role="region" aria-label="Excel schedule; scroll horizontally for more columns"' : '';
    return '<div class="doc-frame doc-frame-' + kind + '" data-doc="' + kind + '"' + scrollRegion + '>' + html + '</div>';
  }

  /* ---------- documents ---------- */

  function bridgeSVG() {
    var v = focusVehicle();
    var W = 500, H = 236, top = 26, bottom = 182;
    var bars = [{ label: P.period.prior.split(' ')[0] + ' NAV', value: v.navPrior, total: true }]
      .concat(P.bridge.map(function (b) { return { label: b.label, value: b.value }; }))
      .concat([{ label: P.period.label.split(' ')[0] + ' NAV', value: v.nav, total: true }]);
    var cum = v.navPrior, levels = [v.navPrior];
    P.bridge.forEach(function (b) { cum += b.value; levels.push(cum); });
    var mn = Math.min.apply(null, levels.concat([v.nav])), mx = Math.max.apply(null, levels.concat([v.nav]));
    var span = mx - mn;
    var lo = Math.floor((mn - span * 0.7) / 5) * 5, hi = Math.ceil((mx + span * 0.2) / 5) * 5;
    function y(val) { return top + (hi - val) / (hi - lo) * (bottom - top); }
    var slot = (W - 20) / bars.length, bw = slot * 0.5;
    var out = '<svg class="bridge" viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="NAV bridge">';
    out += '<line x1="10" x2="' + (W - 10) + '" y1="' + bottom + '" y2="' + bottom + '" class="axis"/>';
    var run = v.navPrior;
    bars.forEach(function (b, i) {
      var x = 10 + slot * i + (slot - bw) / 2;
      var y1, y2, cls, lbl;
      if (b.total) { y1 = y(b.value); y2 = bottom; cls = 'bar-total'; lbl = num(b.value); }
      else {
        var from = run, to = run + b.value;
        y1 = y(Math.max(from, to)); y2 = y(Math.min(from, to));
        cls = b.value < 0 ? 'bar-neg' : 'bar-pos'; lbl = signed(b.value);
        run = to;
      }
      out += '<rect x="' + x.toFixed(1) + '" y="' + y1.toFixed(1) + '" width="' + bw.toFixed(1) + '" height="' + Math.max(1.5, y2 - y1).toFixed(1) + '" class="' + cls + '"/>';
      out += '<text x="' + (x + bw / 2).toFixed(1) + '" y="' + (y1 - 6).toFixed(1) + '" class="bar-val">' + lbl + '</text>';
      if (i < bars.length - 1) {
        var level = b.total && i === 0 ? b.value : run;
        out += '<line x1="' + (x + bw).toFixed(1) + '" x2="' + (x + slot).toFixed(1) + '" y1="' + y(level).toFixed(1) + '" y2="' + y(level).toFixed(1) + '" class="connector"/>';
      }
      var lines = splitLabel(b.label);
      lines.forEach(function (ln, j) {
        out += '<text x="' + (x + bw / 2).toFixed(1) + '" y="' + (bottom + 16 + j * 13) + '" class="bar-lbl">' + esc(ln) + '</text>';
      });
    });
    out += '<path d="M14 ' + (bottom - 9) + 'l6 -4M14 ' + (bottom - 4) + 'l6 -4" class="break"/>';
    out += '</svg>';
    return out;
  }
  function splitLabel(label) {
    var w = label.split(' ');
    if (w.length < 2 || label.length <= 11) return [label];
    var best = null;
    for (var i = 1; i < w.length; i++) {
      var a = w.slice(0, i).join(' '), b = w.slice(i).join(' ');
      var score = Math.max(a.length, b.length);
      if (!best || score < best.score) best = { score: score, lines: [a, b] };
    }
    return best.lines;
  }

  function pendingCommentary() {
    var v = focusVehicle(), d = derive(v);
    return v.short + ' commentary pending. Net IRR changed ' + bps(d.dBps) + '; explanation requested from ' + P.people.partner.name + '.';
  }

  function deckHTML(state) {
    var v = focusVehicle(), d = derive(v), ok = state === 'approved';
    var q2 = P.period.prior.split(' ')[0];
    function kpi(label, value, sub, extra, focus) {
      return '<div class="deck-kpi' + (focus ? ' is-focus' : '') + (focus && ok ? ' is-updated' : '') + '">' +
        '<div class="deck-kpi-label">' + label + '</div>' +
        '<div class="deck-kpi-value">' + value + (focus && !ok ? '<span class="deck-unapproved">Unapproved</span>' : '') + '</div>' +
        '<div class="deck-kpi-sub">' + sub + (extra ? ' <b>' + extra + '</b>' : '') + '</div></div>';
    }
    return '<div class="doc deck deck-' + P.brand.deckFont + '" style="--brand:' + P.brand.color + ';--brand-accent:' + P.brand.accent + '">' +
      '<div class="deck-rule"></div>' +
      '<div class="deck-top"><div><div class="deck-kicker">' + esc(v.short) + ' · ' + P.period.label + '</div>' +
      '<div class="deck-title">Performance and value bridge</div></div>' +
      '<div class="deck-logo"><b>' + P.brand.mark + '</b><span>' + P.brand.sub + '</span></div></div>' +
      '<div class="deck-main"><div class="deck-kpis">' +
      kpi('Net IRR', pct(v.irr), q2 + ' ' + pct(v.irrPrior), bps(d.dBps), true) +
      kpi('Net TVPI', mult(d.tvpi), q2 + ' ' + mult(d.tvpiPrior)) +
      kpi('DPI', mult(d.dpi), q2 + ' ' + mult(d.dpi)) +
      kpi('Net asset value', money(v.nav), q2 + ' ' + money(v.navPrior)) +
      '</div><div class="deck-chart"><div class="deck-chart-title">NAV bridge, ' + P.period.prior + ' to ' + P.period.label + ' (USD millions)</div>' + bridgeSVG() + '</div></div>' +
      '<div class="deck-commentary ' + (ok ? 'is-updated' : 'is-pending') + '"><div class="deck-commentary-label">Commentary</div>' +
      '<p>' + esc(ok ? commentary('v2') : pendingCommentary()) + '</p></div>' +
      '<div class="deck-foot"><span>Confidential. Prepared for the limited partners of ' + esc(v.name) + ' As of ' + P.period.asOf + '.</span><span>' + P.focus.deckSlide + '</span></div>' +
      '</div>';
  }

  function excelHTML(state) {
    var ok = state === 'approved';
    var focus = focusVehicle();
    var cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K'];
    var head = ['Vehicle', 'Vintage', 'Commit.', 'Paid-in', 'Distrib.', 'NAV', 'Net TVPI', 'DPI', 'IRR ' + P.period.prior.split(' ')[0], 'IRR ' + P.period.label.split(' ')[0], 'Chg (bps)'];
    var rows = '';
    function cell(content, cls) { return '<td class="' + (cls || '') + '">' + content + '</td>'; }
    function rowHead(n) { return '<th class="rn">' + n + '</th>'; }
    function blank(n) { var s = ''; for (var i = 0; i < n; i++) s += '<td></td>'; return s; }
    rows += '<tr>' + rowHead(1) + cell('<b>Fund performance summary</b>', 'span-text') + blank(10) + '</tr>';
    rows += '<tr>' + rowHead(2) + cell('As of ' + P.period.asOf + '. USD millions, net of fees and carry.', 'muted span-text') + blank(10) + '</tr>';
    rows += '<tr>' + rowHead(3) + blank(11) + '</tr>';
    rows += '<tr class="hdr">' + rowHead(4) + head.map(function (h, i) { return cell(h, i === 0 ? '' : 'r'); }).join('') + '</tr>';
    var t = { c: 0, p: 0, d: 0, n: 0 };
    P.vehicles.forEach(function (v, i) {
      var d = derive(v), isF = v.id === focus.id;
      t.c += v.commitments; t.p += v.paidIn; t.d += v.distributions; t.n += v.nav;
      var irrCls = 'r' + (isF ? (ok ? ' sel is-updated' : ' sel pending') : '');
      rows += '<tr>' + rowHead(5 + i) +
        cell(esc(v.short) + (isF ? ' <sup>(1)</sup>' : '')) + cell(v.vintage, 'r') + cell(num(v.commitments), 'r') + cell(num(v.paidIn), 'r') +
        cell(num(v.distributions), 'r') + cell(num(v.nav), 'r') + cell(mult(d.tvpi), 'r') + cell(mult(d.dpi), 'r') +
        cell(pct(v.irrPrior), 'r') + cell(pct(v.irr), irrCls) + cell(signed(d.dBps, 0), 'r' + (Math.abs(d.dBps) >= P.focus.rule.irrBps ? ' neg' : '')) + '</tr>';
    });
    var tr = 5 + P.vehicles.length;
    rows += '<tr class="tot">' + rowHead(tr) + cell('Total') + cell('') + cell(num(t.c), 'r') + cell(num(t.p), 'r') + cell(num(t.d), 'r') + cell(num(t.n), 'r') + blank(5) + '</tr>';
    rows += '<tr>' + rowHead(tr + 1) + blank(11) + '</tr>';
    rows += '<tr class="note">' + rowHead(tr + 2) + '<td colspan="11" class="' + (ok ? 'is-updated' : 'pending-note') + '"><span class="note-k">(1) ' + esc(focus.short) + ':</span> ' + esc(ok ? commentary('v2') : pendingCommentary()) + '</td></tr>';
    for (var extra = tr + 3; extra <= 26; extra++) rows += '<tr>' + rowHead(extra) + blank(11) + '</tr>';
    var formula = '=REGISTRY("' + P.focus.registryId + '", "' + P.period.label.replace(/(Q\d) (\d{4})/, '$2$1') + '")';
    return '<div class="doc excel">' +
      '<div class="xl-title"><span class="xl-icon">X</span><span>' + esc(workbookFile()) + '</span><span class="xl-state ' + (ok ? 'ok' : 'draft') + '">' + (ok ? 'Approved v3' : 'Draft') + '</span></div>' +
      '<div class="xl-formula"><span class="xl-ref">' + excelCell() + '</span><span class="xl-fx">fx</span><span class="xl-f">' + esc(formula) + '</span></div>' +
      '<div class="xl-sheet"><table class="xl-grid"><thead><tr><th class="corner rn"></th>' + cols.map(function (c) { return '<th>' + c + '</th>'; }).join('') + '</tr></thead><tbody>' + rows + '</tbody></table></div>' +
      '<div class="xl-tabs">' + P.package.workbookTabs.map(function (t, i) { return '<span' + (i === 0 ? ' class="on"' : '') + '>' + esc(t) + '</span>'; }).join('') + '</div>' +
      '</div>';
  }

  function letterHTML(state) {
    var ok = state === 'approved';
    var v = focusVehicle(), d = derive(v);
    var body = '<h4>' + esc(v.short) + '</h4>' +
      '<p class="' + (ok ? 'is-updated' : 'pending-para') + '">' + esc(ok ? commentary('v2') : pendingCommentary()) + '</p>' +
      '<p>During the quarter, ' + esc(v.short) + ' made no new investments and called no capital. Distributions since inception total ' + moneyLong(v.distributions) + ' (' + mult(d.dpi) + ' paid-in capital), and net asset value was ' + moneyLong(v.nav) + ' at ' + P.period.asOf + ', for a net TVPI of ' + mult(d.tvpi) + '.</p>';
    otherVehicles().forEach(function (o) {
      var od = derive(o);
      body += '<h4>' + esc(o.short) + '</h4><p>' + esc(o.short) + ' net IRR was ' + pct(o.irr) + ' at ' + P.period.asOfShort + ', compared with ' + pct(o.irrPrior) + ' at ' + P.period.priorAsOf + '. ' +
        (o.distributions > 0
          ? 'Distributions since inception total ' + moneyLong(o.distributions) + ', or ' + mult(od.dpi) + ' paid-in capital, and net TVPI was ' + mult(od.tvpi) + '.'
          : 'Net TVPI was ' + mult(od.tvpi) + '; the vehicle has not yet made distributions.') + '</p>';
    });
    return '<div class="doc letter" style="--brand:' + P.brand.color + '">' +
      '<div class="letter-head"><span class="letter-mark">' + esc(P.firm) + '</span><span>Letter to Limited Partners · ' + P.period.quarterWord + '</span></div>' +
      '<div class="letter-body">' + body + '</div>' +
      '<div class="letter-foot"><span>Confidential</span><span>Page ' + P.focus.letterPage + '</span></div>' +
      '</div>';
  }

  /* ---------- steps ---------- */

  function story(n) {
    var v = focusVehicle(), d = derive(v), pp = P.people;
    var wc = words(commentary('v1'));
    return [
      { title: P.package.name, sub: 'Three outputs built from one set of figures. One change needs an explanation before anything goes to limited partners.' },
      { title: 'The prior-period diff finds what changed', sub: v.short + ' net IRR fell from ' + pct(v.irrPrior) + ' to ' + pct(v.irr) + ', past the ' + P.focus.rule.irrBps + ' bps rule. The chip links the figure to the exact model cell.' },
      { title: 'The partner explains it out loud', sub: pp.partner.name + ' answers in ' + P.recording.duration + ' of voice instead of writing prose. The transcript is split into lines that can be cited.' },
      { title: 'A cited draft in two minutes', sub: 'A ' + wc + '-word commentary block drafted from the model and the transcript. Every claim links to its sources and is checked on its own.' },
      { title: 'The reviewer’s edit becomes a typed signal', sub: pp.ir.name + ' corrects one sentence. The edit is logged as a factual correction, not a tone change, so it counts as a real error.' },
      { title: 'Approve once, update everywhere', sub: pp.cfo.name + ' approves. The figure and the commentary update the LP deck, the Excel schedule and the letter together.' },
      { title: 'Audit trail, and what stays the same for every customer', sub: 'Who changed what, when and why, by version. Swap in a different prospect and only the configuration changes.' }
    ][n];
  }

  function step0() {
    var v = focusVehicle(), d = derive(v), pp = P.people, s = sources();
    var ms = P.package.milestones.map(function (m, i) {
      var st = i === 0 ? 'done' : i === 1 ? 'now' : 'next';
      return '<li class="ms ms-' + st + '"><span class="ms-dot">' + (st === 'done' ? ICON.check : '') + '</span><span class="ms-label">' + m.label + '</span><span class="ms-date">' + m.date + '</span></li>';
    }).join('');
    var timeline = panel({ id: 'package', legend: 'proposed', why: 'Report-cycle workspace', kicker: P.firm, title: P.package.name + ' · in preparation',
      body: '<ol class="milestones">' + ms + '</ol>' });
    var o = P.package.outputs;
    function tile(kind, name, meta) {
      return '<figure class="tile">' + frame(kind, 'draft') + '<figcaption><b>' + name + '</b><span>' + meta + '</span><span class="pill pill-draft">Draft</span></figcaption></figure>';
    }
    var outputs = panel({ id: 'outputs', legend: 'today', why: 'Brand-compliant output generation across Word, PPT and Excel', kicker: 'Rendered from the package', title: 'Outputs for limited partners', cls: 'panel-outputs',
      body: '<div class="tiles">' + tile('deck', 'LP deck', o.deckSlides + ' slides') + tile('excel', 'Excel schedules', P.package.workbookTabs.length + ' tabs') + tile('letter', 'LP letter', o.letterPages + ' pages') + '</div>' });
    var vrows = P.vehicles.map(function (x) {
      var xd = derive(x), isF = x.id === v.id;
      return '<tr class="' + (isF ? 'is-focus' : '') + '"><td>' + esc(x.name) + '</td><td class="r">' + x.vintage + '</td><td class="r">' + money(x.commitments) + '</td><td class="r">' + money(x.nav) + '</td><td class="r">' + mult(xd.tvpi) + '</td><td class="r">' + pct(x.irr) + (isF ? ' <span class="pill pill-flag">' + ICON.flag + bps(xd.dBps) + '</span>' : '') + '</td></tr>';
    }).join('');
    var vehicles = panel({ id: 'vehicles', legend: 'proposed', why: 'Report object model: vehicle, asset, period, metric, commentary block, source', kicker: P.strategy + ' · as of ' + P.period.asOf, title: 'Vehicles in this package',
      body: '<table class="mini"><thead><tr><th>Vehicle</th><th class="r">Vintage</th><th class="r">Commitments</th><th class="r">NAV</th><th class="r">Net TVPI</th><th class="r">Net IRR</th></tr></thead><tbody>' + vrows + '</tbody></table>' });
    var flag = panel({ id: 'flag', legend: 'proposed', why: 'Prior-period diff with an explanation rule', kicker: 'Needs explanation', title: esc(v.short) + ' · ' + P.focus.metric, cls: 'panel-flag',
      body: '<div class="flag-figs"><span class="was">' + pct(v.irrPrior) + '</span>' + ICON.arrow + '<span class="now">' + pct(v.irr) + '</span><span class="delta">' + bps(d.dBps) + '</span></div>' +
        '<p>Over the ' + P.focus.rule.irrBps + ' bps rule. Explanation requested from ' + esc(pp.partner.name) + '.</p>' +
        '<div class="cite-row">' + cite('M1', s) + cite('P1', s) + '</div>' });
    var registry = panel({ id: 'registry', legend: 'proposed', why: 'Figure registry: one approved number feeds every output', kicker: 'Figure registry', title: P.package.figures + ' figures, one source of truth',
      body: '<div class="reg-bar"><span style="flex:' + (P.package.figures - 1) + '"></span><i></i></div>' +
        '<dl class="reg-stats"><div><dt>Within rule</dt><dd>' + (P.package.figures - 1) + '</dd></div><div><dt>Needs explanation</dt><dd class="amber">1</dd></div><div><dt>Outputs fed</dt><dd>3</dd></div></dl>' });
    var people = panel({ id: 'people', legend: 'proposed', why: 'Approval chain is customer configuration', kicker: 'Approval chain · configured', title: 'People on this package',
      body: '<ul class="people-list">' + [
        [pp.ir, 'Assembles the package'],
        [pp.partner, 'Explains changes, signs off commentary'],
        [pp.cfo, 'Approves for distribution']
      ].map(function (r) { return '<li>' + avatar(r[0]) + '<div><b>' + esc(r[0].name) + '</b><span>' + esc(r[0].role) + '</span></div><em>' + r[1] + '</em></li>'; }).join('') + '</ul>' });
    var fit = panel({ id: 'fit', legend: 'assumption', why: 'Signals sales can spot before deployment; to audit against current engagements', kicker: 'Why this prospect fits', title: 'Signals visible before deployment',
      body: '<ul class="tags"><li>Quarterly LP cycle</li><li>' + P.vehicles.length + ' vehicles</li><li>CFO-led close</li><li>Signed, reconciled letter</li></ul>' });
    return '<div class="grid g0"><div class="col-main">' + timeline + outputs + vehicles + '</div><div class="col-side">' + flag + registry + people + fit + '</div></div>';
  }

  function step1() {
    var v = focusVehicle(), d = derive(v), s = sources(), m = P.focus.model, rule = P.focus.rule;
    function periodPane(prior) {
      return '<article class="period-pane' + (prior ? '' : ' current') + '"' + (prior ? ' data-element="prior"' : '') + '>' +
        '<div class="label">' + (prior ? P.period.prior + ' · ' + P.focus.prior.version : P.period.label + ' · Draft') + ' ' + chip('proposed') + '</div>' +
        '<h2>' + esc(v.short) + ' performance</h2><div class="pane-metric">' + pct(prior ? v.irrPrior : v.irr) + (prior ? '' : cite('M1', s)) + '</div>' +
        '<div class="pane-sub' + (prior ? '' : ' focus-line') + '">' + (prior ? 'Net IRR · approved by ' + esc(P.focus.prior.approvedBy) + ' on ' + P.focus.prior.approvedOn : bps(d.dBps) + ' · over the ' + rule.irrBps + ' bps explanation rule') + '</div>' +
        '<div class="mini-statement"><div><span>Net asset value</span><b>' + money(prior ? v.navPrior : v.nav) + '</b></div>' +
        '<div><span>Net TVPI</span><b>' + mult(prior ? d.tvpiPrior : d.tvpi) + '</b></div><div><span>DPI</span><b>' + mult(d.dpi) + '</b></div></div></article>';
    }
    var compare = '<section class="compare" data-element="diff">' + periodPane(true) + periodPane(false) + '</section>';
    var source = '<aside class="inspector" data-element="source"><div class="label">Evidence inspector</div><h3>' + esc(m.sheet + '!' + m.cell) + '</h3>' +
      '<div class="inspect-block"><b>' + esc(m.file) + '</b><span>' + esc(m.version) + '</span></div>' +
      '<div class="inspect-block"><b>Formula</b><p>' + esc(m.formula) + '</p>' + chip('today', 'Source-linked citation chips; Excel-based agents read the model') + '</div>' +
      '<div class="inspect-block" data-element="rule"><b>Configured rule</b><p>Net IRR ≥ ' + rule.irrBps + ' bps or NAV ≥ ' + rule.navPct + '% quarter over quarter.</p>' + chip('proposed', 'Explanation rule; thresholds are customer configuration') + '</div>' +
      '<div class="explain-call">Explanation requested from ' + esc(P.people.partner.name) + '</div></aside>';
    var rows = P.holdings.map(function (h) {
      return '<tr><td>' + esc(h.name) + '</td><td>' + cite(h.id, s) + '</td><td class="r">' + money(h.prior) + '</td><td class="r">' + money(h.current) + '</td><td class="r neg">' + signed(h.current - h.prior).replace(MINUS, MINUS + '$').replace('+', '+$') + 'M</td></tr>';
    }).join('');
    var other = P.bridge[P.bridge.length - 1];
    rows += '<tr><td>' + esc(other.label) + '</td><td>—</td><td></td><td></td><td class="r">' + signed(other.value).replace(MINUS, MINUS + '$').replace('+', '+$') + 'M</td></tr>';
    var holdings = '<section class="desk-table" data-element="holdings"><header class="table-head"><div><div class="label">What moved inside ' + esc(v.short) + '</div><h3>Valuation model bridge</h3></div>' + chip('today', 'Excel-based agents read the model') + '</header>' +
      '<div class="table-scroll" tabindex="0" role="region" aria-label="Valuation model bridge"><table><thead><tr><th>Holding</th><th>Source cell</th><th class="r">' + P.period.prior + '</th><th class="r">' + P.period.label + '</th><th class="r">Change</th></tr></thead><tbody>' + rows + '</tbody></table></div></section>';
    return '<div class="desk">' + compare + source + holdings + '</div>';
  }

  function waveform() {
    var seed = 0;
    for (var i = 0; i < P.id.length; i++) seed = (seed * 31 + P.id.charCodeAt(i)) >>> 0;
    function rnd() { seed = (seed * 1664525 + 1013904223) >>> 0; return seed / 4294967296; }
    var n = 180, W = 1600, H = 96, bw = W / n;
    var gapsAt = P.recording.lines.map(function (l) { var p = l.t.split(':'); return (parseInt(p[0], 10) * 60 + parseInt(p[1], 10)) / P.recording.seconds; });
    var out = '<svg class="wave" viewBox="0 0 ' + W + ' ' + (H + 24) + '" preserveAspectRatio="none" role="img" aria-label="Recorded waveform">';
    for (var k = 0; k < n; k++) {
      var pos = k / n, quiet = gapsAt.some(function (g) { return Math.abs(pos - g) < 0.009; });
      var amp = quiet ? 0.05 : 0.16 + 0.84 * Math.pow(rnd(), 1.35) * (0.62 + 0.38 * Math.sin(k / 6));
      var h = Math.max(3, amp * H);
      out += '<rect x="' + (k * bw + bw * 0.22).toFixed(1) + '" y="' + ((H - h) / 2 + 20).toFixed(1) + '" width="' + (bw * 0.56).toFixed(1) + '" height="' + h.toFixed(1) + '" rx="2"/>';
    }
    P.recording.lines.forEach(function (l, j) {
      var x = gapsAt[j] * W;
      out += '<line x1="' + x.toFixed(1) + '" x2="' + x.toFixed(1) + '" y1="14" y2="' + (H + 22) + '" class="mark"/>';
    });
    out += '</svg>';
    var labels = P.recording.lines.map(function (l, j) {
      return '<span class="wave-label" style="left:' + (gapsAt[j] * 100).toFixed(2) + '%">' + l.id + ' · ' + l.t + '</span>';
    }).join('');
    return '<div class="wave-box">' + out + '<div class="wave-labels">' + labels + '</div></div>';
  }

  function step2() {
    var v = focusVehicle(), d = derive(v), pp = P.people;
    var lines = P.recording.lines.map(function (l) {
      return '<li><span class="cite cite-voice">' + ICON.voice + l.id + '</span><span class="t">' + l.t + '</span><p>' + esc(l.text) + '</p></li>';
    }).join('');
    var rec = panel({ id: 'recording', legend: 'today', why: 'Voice interview capture', kicker: 'Voice explanation', title: 'Why did ' + esc(v.short) + ' net IRR fall?', cls: 'panel-rec',
      body: '<div class="speaker">' + avatar(pp.partner, 'lg') + '<div><b>' + esc(pp.partner.name) + '</b><span>' + esc(pp.partner.role) + ' · Recorded ' + P.recording.recordedOn + '</span></div><span class="dur">' + ICON.voice + P.recording.duration + '</span></div>' +
        '<div class="wave-wrap">' + waveform() + '</div>' });
    var trans = panel({ id: 'transcript', legend: 'today', why: 'Voice interview capture produces the transcript; line IDs are citation targets', kicker: 'Transcript', title: P.recording.lines.length + ' citable lines', cls: 'panel-transcript',
      body: '<ol class="transcript">' + lines + '</ol>' });
    var req = panel({ id: 'request', legend: 'proposed', why: 'Explanation request routed from the flagged figure', kicker: 'Explanation request', title: 'Sent to ' + esc(pp.partner.name),
      body: '<div class="req-from">' + avatar(pp.ir) + '<span><b>' + esc(pp.ir.name) + '</b> · ' + esc(pp.ir.role) + ' · ' + P.request.sentOn + '</span></div>' +
        '<div class="req-fig"><span>' + esc(v.short) + ' · Net IRR</span><b>' + pct(v.irrPrior) + ' → ' + pct(v.irr) + '</b><span class="neg">' + bps(d.dBps) + '</span></div>' +
        '<div class="guide"><div class="kicker">Interview guide · ' + esc(P.request.guide) + '</div><ol>' + P.request.prompts.map(function (q) { return '<li>' + esc(q) + '</li>'; }).join('') + '</ol></div>' });
    var asm = panel({ id: 'voice-assumption', legend: 'assumption', why: 'Behavior to validate in the pilot', kicker: 'To validate', title: 'Partners will talk, not type',
      body: '<p>Assumes a partner will answer a short voice prompt every quarter. If not, the draft starts from prior-period text and the reviewer writes more.</p>' });
    return '<div class="grid g2 hero-grid"><div class="row-hero">' + rec + '</div><div class="col-main">' + trans + '</div><div class="col-side">' + req + asm + '</div></div>';
  }

  function commentaryHTML(version, s, mode) {
    return P.claims.map(function (c) {
      var t;
      if (mode === 'tracked' && c.edit) {
        var parts = c.text.split('{edit}');
        t = esc(parts[0]) + '<del>' + esc(c.edit.before) + '</del> <ins>' + esc(c.edit.after) + '</ins>' + esc(parts[1]);
      } else {
        t = esc(claimText(c, version));
      }
      var cls = 'claim';
      return '<span class="' + cls + '" data-claim="' + c.id + '">' + t + '</span>' + c.cites.map(function (id) { return cite(id, s); }).join('');
    }).join(' ');
  }

  function autoChecks(c) {
    var out = [];
    if (c.id === 'C1') out.push('Figure matches registry');
    out.push('Citations resolve');
    return out.map(function (x) { return '<span class="ac">' + ICON.check + x + '</span>'; }).join('');
  }

  function step3() {
    var v = focusVehicle(), s = sources();
    var text1 = commentary('v1');
    var draft = panel({ id: 'draft', legend: 'today', why: 'Drafting through the DAG and citation chips exist today; the registry input is proposed', kicker: 'Commentary block · v1 draft', title: esc(v.short) + ' performance commentary', cls: 'panel-draft',
      body: '<div class="prose">' + commentaryHTML('v1', s, 'plain') + '</div>' +
        '<div class="meta-row"><span>' + words(text1) + ' words</span><span>' + P.claims.length + ' claims</span><span>' + P.claims.length + ' of ' + P.claims.length + ' cited</span><span>Drafted ' + P.timeline.drafted + ', two minutes after the recording</span><span>Goes to LP letter, deck slide ' + P.focus.deckSlide + ' and schedule note (1)</span></div>' });
    var claimRows = P.claims.map(function (c) {
      var ev = c.cites.map(function (id) { return '<span><b>' + esc(s[id].short) + '</b> ' + esc(s[id].detail) + '</span>'; }).join('');
      return '<tr><td class="cid">' + c.id + '</td><td class="ctext">' + esc(claimText(c, 'v1')) + '<div class="evidence">' + ev + '</div></td><td class="srcs">' + c.cites.map(function (id) { return cite(id, s); }).join('') + '</td><td class="acs">' + autoChecks(c) + '</td><td><span class="pill pill-wait">Awaiting review</span></td></tr>';
    }).join('');
    var claims = panel({ id: 'claims', legend: 'proposed', why: 'The claim is the evaluation unit: one figure or statement, its sources and the reviewer’s disposition', kicker: 'Evaluation unit', title: 'Each claim is checked on its own', cls: 'panel-claims',
      body: '<table class="claims"><thead><tr><th>ID</th><th>Claim</th><th>Sources</th><th>Automatic checks</th><th>Disposition</th></tr></thead><tbody>' + claimRows + '</tbody></table>' +
        '<p class="foot muted">Numbers and citations can be checked automatically. Whether the wording is faithful to the source still needs a person.</p>' });
    var inputs = ['Registry figure: ' + esc(v.short) + ' net IRR ' + pct(v.irr), 'Model cells ' + s.M1.short + ', ' + s.M2.short + ', ' + s.M3.short, 'Transcript T1–T' + P.recording.lines.length, P.period.prior + ' letter, for house style'];
    var stepsDag = ['Extract claims from sources', 'Draft in house style, 60 words max', 'Attach a citation to every claim', 'Brand and tone check'];
    var dag = panel({ id: 'dag', legend: 'today', why: 'Shared ML service orchestrating reasoning and generation through a DAG', kicker: 'Shared ML service', title: 'How the draft was assembled', cls: 'panel-dag',
      body: '<div class="dag"><div class="dag-inputs">' + inputs.map(function (i, k) { return '<div class="node node-in">' + i + (k === 0 ? chip('proposed', 'The figure registry is proposed') : '') + '</div>'; }).join('') + '</div>' +
        '<div class="dag-edge"></div>' +
        stepsDag.map(function (st, i) { return '<div class="node node-step"><span>' + (i + 1) + '</span>' + st + '</div><div class="dag-edge"></div>'; }).join('') +
        '<div class="node node-out">Commentary v1 · ' + P.claims.length + ' claims, all cited</div></div>' +
        '<p class="foot muted">Drafting through the DAG exists today; the registry input is proposed.</p>' });
    return '<div class="grid g3 hero-grid"><div class="row-hero">' + draft + '</div><div class="col-main">' + claims + '</div><div class="col-side">' + dag + '</div></div>';
  }

  function step4() {
    var v = focusVehicle(), s = sources(), pp = P.people, c = editedClaim();
    var tracked = panel({ id: 'tracked', legend: 'today', why: 'Content-level feedback already records accepted, edited, rejected and ignored', kicker: 'Commentary block · v2', title: 'One sentence corrected', cls: 'panel-tracked',
      body: '<div class="editor-row">' + avatar(pp.ir) + '<span><b>' + esc(pp.ir.name) + '</b> · ' + esc(pp.ir.role) + ' · ' + P.timeline.edited + '</span><span class="pill pill-edit">Edited ' + c.id + '</span></div>' +
        '<div class="prose">' + commentaryHTML('v2', s, 'tracked') + '</div>' });
    var signal = panel({ id: 'signal', legend: 'proposed', why: 'Typed edit classification feeding the evaluation harness', kicker: 'Captured signal', title: 'Logged as a factual correction', cls: 'panel-signal',
      body: '<div class="edit-types"><span class="et on">' + ICON.check + 'Factual correction</span><span class="et">Tone or style</span></div>' +
        '<dl class="sig">' +
        '<div><dt>Claim</dt><dd>' + c.id + ' · ' + esc(claimText(c, 'v2')) + '</dd></div>' +
        '<div><dt>Label</dt><dd>' + esc(P.edit.label) + '</dd></div>' +
        '<div><dt>Why</dt><dd>' + esc(P.edit.reason) + '</dd></div>' +
        '<div><dt>Sources</dt><dd class="quotes">' + ['T3', 'T4'].map(function (id) { return '<span>' + cite(id, s) + ' “' + esc(s[id].detail) + '”</span>'; }).join('') + '</dd></div>' +
        '<div><dt>Slice</dt><dd>' + esc(P.shortFirm) + ' · ' + esc(v.short) + ' commentary · Interview source · ' + P.period.label + '</dd></div>' +
        '<div><dt>Effect</dt><dd>Counts toward the unsupported-claim rate before approval, and joins this cycle’s adjudication sample.</dd></div>' +
        '</dl>' });
    var accepted = P.claims.filter(function (x) { return !x.edit; }).map(function (x) { return x.id; });
    var treat = [
      ['Accepted without edit', 'Weak positive. Sample it.', accepted[0] + '–' + accepted[accepted.length - 1], ''],
      ['Edited', 'Classify the edit first: factual or tone.', c.id + ', factual', 'on'],
      ['Rejected or thumbs down', 'Strong but rare. Always adjudicate.', 'None', ''],
      ['Silence', 'Unknown. Never counted as approval.', 'None', ''],
      ['Corrected after approval', 'Most severe. Tracked by slice.', 'Watched after distribution', '']
    ].map(function (r) { return '<tr class="' + r[3] + '"><td>' + r[0] + '</td><td>' + r[1] + '</td><td class="here">' + r[2] + '</td></tr>'; }).join('');
    var table = panel({ id: 'signals', legend: 'proposed', why: 'Evaluation harness: signals are imperfect, so each is treated differently', kicker: 'Evaluation harness', title: 'How each signal is treated', cls: 'panel-treat',
      body: '<table class="treat"><thead><tr><th>Signal</th><th>How it counts</th><th>This block</th></tr></thead><tbody>' + treat + '</tbody></table>' +
        '<p class="foot muted">A person adjudicates about 50 claims per customer per cycle, stratified by section and source type.</p>' });
    var asm = panel({ id: 'edit-assumption', legend: 'assumption', why: 'Reviewer behavior to validate', kicker: 'To validate', title: 'One click to confirm the edit type',
      body: '<p>Assumes reviewers will confirm a pre-selected edit type. Without it, an edit cannot be counted as an error or as style.</p>' });
    return '<div class="grid g4 hero-grid"><div class="row-hero">' + tracked + '</div><div class="col-main">' + signal + asm + '</div><div class="col-side">' + table + '</div></div>';
  }

  function step5() {
    var pp = P.people, t = P.timeline;
    var chain = [[pp.ir, 'Prepared', t.edited], [pp.partner, 'Signed off', t.signedOff], [pp.cfo, 'Approved', t.approved]].map(function (r) {
      return '<div><b>' + esc(r[0].name) + ' · ' + esc(r[0].role) + '</b><span>' + r[1] + ' · ' + r[2] + '</span></div>';
    }).join('');
    var approval = '<section class="approval-rail" data-element="approval"><div class="rail-title">Approval · Commentary v3 ' + chip('proposed', 'Approval chain with version history') + '</div>' + chain + '</section>';
    function docPanel(kind, name, where) {
      return '<article class="doc-panel doc-panel-' + kind + '"><header class="doc-panel-head"><div><b>' + name + '</b><span>' + esc(where) + '</span></div>' + chip('today', 'Brand-compliant output generation across PPT, Excel and Word') + '</header>' + frame(kind, 'approved') + '</article>';
    }
    var docs = '<section class="document-desk" data-element="outputs-updated">' + docPanel('deck', 'LP deck', 'Slide ' + P.focus.deckSlide) + docPanel('excel', 'Excel schedule', scheduleSheet() + '!' + excelCell()) + docPanel('letter', 'LP letter', 'Page ' + P.focus.letterPage) + '</section>';
    var consistency = '<section class="doc-consistency" data-element="consistency"><div><div class="label">Consistency check ' + chip('proposed', 'One approved number and one commentary version feed every output') + '</div><span class="muted">LP deck · Excel schedule · LP letter</span></div><b>3 of 3 outputs match the registry</b></section>';
    return '<div class="approval-desk">' + approval + docs + consistency + '</div>';
  }

  function step6() {
    var v = focusVehicle(), d = derive(v), pp = P.people, t = P.timeline, c = editedClaim();
    var sys = { name: 'System', role: 'Registry', initials: '·' };
    var ml = { name: 'Shared ML service', role: 'Drafting', initials: 'AI' };
    var events = [
      [t.dataClose, sys, 'Data close. ' + esc(P.focus.model.file) + ' locked.', ''],
      [t.flagged, sys, esc(v.short) + ' net IRR flagged: ' + pct(v.irrPrior) + ' → ' + pct(v.irr) + ', ' + bps(d.dBps) + '.', 'Rule ' + P.focus.rule.irrBps + ' bps'],
      [P.request.sentOn, pp.ir, 'Explanation requested from ' + esc(pp.partner.name) + '.', ''],
      [P.recording.recordedOn, pp.partner, 'Voice explanation recorded, ' + P.recording.duration + '.', 'T1–T' + P.recording.lines.length],
      [t.drafted, ml, 'Commentary drafted: ' + P.claims.length + ' claims, all cited.', 'v1'],
      [t.edited, pp.ir, c.id + ' corrected: “' + esc(c.edit.before) + '” → “' + esc(c.edit.after) + '”.', 'v2 · factual'],
      [t.signedOff, pp.partner, 'Commentary signed off.', ''],
      [t.approved, pp.cfo, 'Approved and rendered to deck, schedule and letter.', 'v3 · 3 of 3']
    ].map(function (e) {
      return '<li><span class="when">' + e[0] + '</span>' + avatar(e[1], e[1] === sys || e[1] === ml ? 'sys' : '') + '<span class="what"><b>' + esc(e[1].name) + '</b> ' + e[2] + '</span>' + (e[3] ? '<span class="tag">' + e[3] + '</span>' : '') + '</li>';
    }).join('');
    var audit = panel({ id: 'audit', legend: 'proposed', why: 'Audit trail across the report cycle', kicker: esc(v.short) + ' net IRR · ' + P.period.label, title: 'Audit trail', cls: 'panel-audit',
      body: '<ol class="audit">' + events + '</ol>' });
    var parts = c.text.split('{edit}');
    var versions = panel({ id: 'versions', legend: 'proposed', why: 'Version history with diffs', kicker: 'Commentary block', title: 'Version history', cls: 'panel-versions',
      body: '<ol class="versions">' +
        '<li><span class="vtag">v1</span><div><b>Drafted</b> · Shared ML service · ' + t.drafted + '</div></li>' +
        '<li><span class="vtag">v2</span><div><b>Edited</b> · ' + esc(pp.ir.name) + ' · ' + t.edited + '<p class="vdiff">…' + esc(parts[0].trim().split(' ').slice(-5).join(' ')) + ' <del>' + esc(c.edit.before) + '</del> <ins>' + esc(c.edit.after) + '</ins>' + esc(parts[1]) + '</p></div></li>' +
        '<li><span class="vtag on">v3</span><div><b>Approved</b> · ' + esc(pp.cfo.name) + ' · ' + t.approved + '<p class="muted">No text change. Approval locked the figure and the commentary, then rendered 3 outputs.</p></div></li>' +
        '</ol>' });

    function li(items) { return '<ul>' + items.map(function (i) { return '<li>' + i + '</li>'; }).join('') + '</ul>'; }
    function item(text, kind) { return '<span class="b-item">' + text + '</span>' + chip(kind); }
    var shared = '<div class="b-col" data-element="b-shared"><div class="b-head"><b>Shared product</b><span>Same for every customer</span>' + chip('proposed', 'Built once by Brain Co.') + '</div>' +
      '<ul class="b-list">' +
      '<li>' + item('Report object model: vehicle, asset, period, metric, commentary, source', 'proposed') + '</li>' +
      '<li>' + item('Figure registry: one approved number feeds every output', 'proposed') + '</li>' +
      '<li>' + item('Source-linked citations', 'today') + '</li>' +
      '<li>' + item('Voice interview capture', 'today') + '</li>' +
      '<li>' + item('Commentary drafted from prior-period text', 'proposed') + '</li>' +
      '<li>' + item('Review and approval with version history and prior-period diff', 'proposed') + '</li>' +
      '<li>' + item('Feedback capture', 'today') + '</li>' +
      '<li>' + item('Evaluation harness', 'proposed') + '</li>' +
      '</ul></div>';
    var configured = '<div class="b-col b-config" data-element="b-config"><div class="b-head"><b>Configured for ' + esc(P.shortFirm) + '</b><span>By the deployment team, no code</span>' + chip('proposed', 'Customer configuration') + '</div>' +
      '<dl class="b-kv">' + P.config.map(function (k) { return '<div><dt>' + esc(k.k) + '</dt><dd>' + esc(k.v) + '</dd></div>'; }).join('') + '</dl></div>';
    var bespoke = '<div class="b-col" data-element="b-bespoke"><div class="b-head"><b>Bespoke</b><span>Priced as services</span>' + chip('proposed', 'Outside the object model') + '</div>' +
      li(P.bespoke.map(esc)) + '<p class="muted small">One-off analyses, unusual connectors, non-recurring report types.</p></div>';
    var declined = '<div class="b-declined" data-element="b-declined"><div class="b-head"><b>Declined in year one</b>' + chip('proposed', 'A proposed decision, not a product part: work that would fragment the product') + '<span>A proposed decision, not a product part</span></div>' +
      '<ul class="tags quiet"><li>Hospitality expansion</li><li>All enterprise reporting</li><li>Per-customer forks</li><li>Real-time dashboards</li><li>Self-serve template builder</li><li>External LP portal</li></ul></div>';
    var test = '<div class="b-test" data-element="b-test">' + chip('assumption', 'Provisional thresholds, reset after the retrospective') + '<p>Configuration-only onboarding is the bet. A retrospective on two past quarters must reach about 75% of report content from configuration with zero code changes, and the paid pilot runs on a fixed configuration budget with no forks. Stop scaling under about 60%. All thresholds are provisional.</p></div>';
    var overlay = '<aside class="overlay" data-element="boundary"><header class="overlay-head"><div><div class="kicker">Product boundary</div><h3>What Brain Co. standardizes, configures and prices as services</h3></div>' + chip('proposed', 'Shared product, customer configuration, bespoke') + '</header>' +
      '<div class="b-grid">' + shared + configured + bespoke + '</div>' + declined + test + '</aside>';
    return '<div class="grid g6"><div class="col-main dimmed">' + audit + versions + '</div>' + overlay + '</div>';
  }

  var BUILDERS = [step0, step1, step2, step3, step4, step5, step6];

  /* ---------- shell ---------- */

  function topbar() {
    return '<div class="brand"><i class="brand-mark" aria-hidden="true"></i><span>Brain Co.</span></div>' +
      '<div class="crumb">FinanceOS / ' + esc(P.firm) + ' / <b>' + esc(P.package.name) + '</b></div>' +
      '<div class="top-actions"><span class="period" id="top-step"></span><span class="sample" data-marker="sample">Sample data - fictional</span></div>';
  }

  function legendKey() {
    return '<div class="legend-key" aria-label="Legend">' + ['today', 'proposed', 'assumption'].map(function (k) { return chip(k); }).join('') + '</div>';
  }

  function nav() {
    var pros = Object.keys(DATA.prospects).map(function (id) {
      var p = DATA.prospects[id];
      return '<button type="button" class="seg' + (id === prospectId ? ' on' : '') + '" data-prospect="' + id + '">' + esc(p.shortFirm) + '<span>' + esc(p.strategy) + '</span></button>';
    }).join('');
    var dots = STEP_LABELS.map(function (l, i) {
      return '<button type="button" class="dot" data-goto="' + i + '" aria-label="Step ' + i + ': ' + l + '"><span>' + i + '</span><em>' + l + '</em></button>';
    }).join('');
    return '<div class="nav-prospect" role="group" aria-label="Prospect">' + pros + '</div>' +
      '<div class="nav-steps"><button type="button" class="arrow" data-nav="prev" aria-label="Previous step">' + ICON.arrow + '</button><div class="dots">' + dots + '</div><button type="button" class="arrow" data-nav="next" aria-label="Next step">' + ICON.arrow + '</button></div>' +
      '<div class="nav-hint">← → keys</div>';
  }

  function render() {
    P = DATA.prospects[prospectId];
    document.title = 'Explain the change · ' + P.firm + ' (sample data)';
    document.getElementById('topbar').innerHTML = topbar();
    document.getElementById('legend').innerHTML = legendKey();
    var stage = document.getElementById('stage');
    stage.innerHTML = BUILDERS.map(function (b, i) {
      return '<section class="step" data-step-panel="' + i + '" hidden>' + b() + '</section>';
    }).join('');
    document.getElementById('demo-nav').innerHTML = capture ? '' : nav();
    document.body.dataset.prospect = prospectId;
  }

  function fitDocs(root) {
    var frames = (root || document).querySelectorAll('.doc-frame');
    for (var i = 0; i < frames.length; i++) {
      var f = frames[i], doc = f.firstElementChild;
      if (!doc || !f.clientWidth) continue;
      // Keep the approved capture composition when the figures fit. Otherwise
      // use the readable sheet width, including in narrow desktop panes.
      if (f.closest('.document-desk')) {
        doc.style.zoom = '';
        if (f.dataset.doc === 'excel') {
          f.classList.remove('is-scrollable');
          var numbers = doc.querySelectorAll('tbody tr:not(.hdr) td.r');
          var needsScroll = window.innerWidth <= 900;
          for (var j = 0; j < numbers.length; j++) {
            if (numbers[j].scrollWidth > numbers[j].clientWidth) needsScroll = true;
          }
          f.classList.toggle('is-scrollable', needsScroll);
          f.scrollLeft = 0;
          var selected = doc.querySelector('td.sel');
          if (needsScroll && selected) {
            var cell = selected.getBoundingClientRect();
            f.scrollLeft = cell.left - f.getBoundingClientRect().left - (f.clientWidth - cell.width) / 2;
          }
        }
        continue;
      }
      if (!doc.dataset.natW) {
        doc.style.zoom = '';
        doc.dataset.natW = String(doc.offsetWidth);
      }
      doc.style.zoom = (f.clientWidth / +doc.dataset.natW).toFixed(4);
    }
  }

  function updateStory() {
    var s = story(current);
    document.getElementById('story-step').textContent = current === 0 ? 'Overview' : 'Step ' + current + ' of ' + (STEP_COUNT - 1) + ' · ' + STEP_LABELS[current];
    document.getElementById('top-step').textContent = current === 0 ? 'Overview' : 'Step ' + current + ' of ' + (STEP_COUNT - 1);
    document.getElementById('story-title').textContent = s.title;
    document.getElementById('story-sub').textContent = s.sub;
    var dots = document.querySelectorAll('[data-goto]');
    for (var i = 0; i < dots.length; i++) {
      var on = +dots[i].dataset.goto === current;
      dots[i].classList.toggle('on', on);
      dots[i].classList.toggle('done', +dots[i].dataset.goto < current);
      if (on) dots[i].setAttribute('aria-current', 'step'); else dots[i].removeAttribute('aria-current');
    }
    var prev = document.querySelector('[data-nav="prev"]'), next = document.querySelector('[data-nav="next"]');
    if (prev) prev.disabled = current === 0;
    if (next) next.disabled = current === STEP_COUNT - 1;
  }

  function syncUrl() {
    try {
      var u = new URL(window.location.href);
      u.searchParams.set('step', String(current));
      u.searchParams.set('prospect', prospectId);
      window.history.replaceState(null, '', u.toString());
    } catch (e) { /* file:// in some browsers refuses replaceState; the URL is only a convenience */ }
  }

  function goTo(n, opts) {
    var target = clampStep(n);
    var animate = !(opts && opts.instant);
    current = target;
    document.body.dataset.step = String(current);
    document.body.dataset.settled = 'false';
    var token = ++settleToken;
    var panels = document.querySelectorAll('[data-step-panel]');
    var active = null;
    for (var i = 0; i < panels.length; i++) {
      var on = +panels[i].dataset.stepPanel === current;
      panels[i].hidden = !on;
      panels[i].classList.remove('is-entering');
      if (on) active = panels[i];
    }
    fitDocs(active);
    if (animate && active) {
      void active.offsetWidth;
      active.classList.add('is-entering');
    }
    updateStory();
    syncUrl();
    settlePromise = new Promise(function (resolve) {
      requestAnimationFrame(function () {
        var anims = active && active.getAnimations ? active.getAnimations({ subtree: true }) : [];
        Promise.all(anims.map(function (a) { return a.finished.catch(function () {}); })).then(function () {
          if (token === settleToken) document.body.dataset.settled = 'true';
          resolve();
        });
      });
    });
    return settlePromise;
  }

  function setProspect(id) {
    if (!DATA.prospects[id] || id === prospectId) return goTo(current);
    prospectId = id;
    render();
    return goTo(current, { instant: true });
  }

  function bind() {
    document.addEventListener('keydown', function (e) {
      if (e.altKey || e.ctrlKey || e.metaKey || e.shiftKey) return;
      if (e.key === 'ArrowRight') { e.preventDefault(); goTo(current + 1); }
      else if (e.key === 'ArrowLeft') { e.preventDefault(); goTo(current - 1); }
    });
    document.getElementById('demo-nav').addEventListener('click', function (e) {
      var b = e.target.closest('button');
      if (!b) return;
      if (b.dataset.goto !== undefined) goTo(+b.dataset.goto);
      else if (b.dataset.nav === 'prev') goTo(current - 1);
      else if (b.dataset.nav === 'next') goTo(current + 1);
      else if (b.dataset.prospect) setProspect(b.dataset.prospect);
    });
    var pending = 0;
    window.addEventListener('resize', function () {
      cancelAnimationFrame(pending);
      pending = requestAnimationFrame(function () { fitDocs(document.querySelector('[data-step-panel]:not([hidden])')); });
    });
  }

  if (capture) document.documentElement.classList.add('capture');
  document.body.classList.toggle('capture', capture);
  render();
  bind();
  goTo(current, { instant: true });
  if (document.fonts && document.fonts.ready) {
    document.fonts.ready.then(function () { fitDocs(document.querySelector('[data-step-panel]:not([hidden])')); });
  }

  window.demo = {
    goTo: function (n) { return goTo(n); },
    steps: STEP_COUNT,
    get prospect() { return prospectId; },
    get step() { return current; },
    next: function () { return goTo(current + 1); },
    prev: function () { return goTo(current - 1); },
    setProspect: setProspect,
    settled: function () { return settlePromise; }
  };
})();
