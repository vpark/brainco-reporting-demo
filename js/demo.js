/*
 * FinanceOS prototype: connections, vault, templates, graph and outputs with cited sources.
 * Classic script so it runs from file://. All data is fictional (data/demo-data.js).
 * Guided-tour contract:
 *   ?step=0-6 opens a state of the demo track; ?capture=1 hides the track navigation for recording.
 *   window.demo = { goTo(n), steps: 7, step, next(), prev(), settled() }; goTo resolves once settled.
 *   body[data-step] mirrors the tour state; body[data-settled] is "true" once transitions finish.
 *   Left/Right arrow keys move through the tour.
 */
(function () {
  'use strict';

  var P = window.BRAINCO_DEMO_DATA;
  var S = P.sources;
  var F = P.fund;
  var STEP_COUNT = 7;
  var MINUS = '\u2212';
  var AUDIO = 'assets/audio/partner-interview-sample.m4a';
  var AUDIO_SECONDS = 72;
  var params = new URLSearchParams(window.location.search);
  var capture = params.get('capture') === '1';
  var current = clampStep(params.get('step'));
  var settlePromise = Promise.resolve();
  var settleToken = 0;

  var TOUR = [
    { view: 'connections', label: 'Connections', title: 'Connected sources feed the ingestion engine', sub: 'Google Drive and the Northwind data room stay in sync; each file type is parsed its own way.' },
    { view: 'vault', label: 'Vault', title: 'The vault keeps one folder per quarter', sub: 'Q2 2026 is generated and approved. Q3 2026 is new and in progress.' },
    { view: 'ingest', label: 'Q3 ingestion', title: 'Every Q3 file, where it came from and what was extracted', sub: 'Pages, cells and formulas, transcript words and the figures each file registered.' },
    { view: 'templates', label: 'Template', title: 'One configured template produces the same four files every quarter', sub: 'Sections, data tags and the source each tag is bound to, plus the firm’s rules.' },
    { view: 'graph', label: 'How it’s built', title: 'Each section has a recipe: ingredients, rules, steps, output', sub: 'Fund II performance shown; the graph view keeps the evidence flow and both resolved flags.' },
    { view: 'outputs', generated: true, label: 'Generate', title: 'Generate the package from cited evidence', sub: 'Four files, every claim verified, the same figures everywhere.' },
    { view: 'detail', drawer: 'C1', tab: 'analysis', label: 'Output detail', title: 'Click any sentence to see where it came from', sub: 'The 11.8% opens the model cell, its precedents and the calculation.' }
  ];

  var state = { view: 'connections', drawer: null, tab: null, approved: false, generated: false, running: -1, history: false, recipe: 'perf', gtab: 'recipe' };

  var LEGEND = {
    today: { label: 'Exists today', note: 'Built in a current Brain Co. engagement' },
    proposed: { label: 'Proposed', note: 'Part of the recommended product' },
    assumption: { label: 'Assumption', note: 'Needs validation in the pilot' }
  };

  var ICON = {
    excel: '<svg viewBox="0 0 12 12" aria-hidden="true"><rect x="1" y="1.5" width="10" height="9" rx="1.4" fill="none" stroke="currentColor" stroke-width="1.1"/><path d="M1 4.6h10M4.6 1.5v9" stroke="currentColor" stroke-width="1.1"/></svg>',
    voice: '<svg viewBox="0 0 12 12" aria-hidden="true"><path d="M2 5v2M4 3.5v5M6 2v8M8 3.5v5M10 5v2" stroke="currentColor" stroke-width="1.2" stroke-linecap="round"/></svg>',
    pdf: '<svg viewBox="0 0 12 12" aria-hidden="true"><path d="M3 1.2h4.2L9.5 3.5v7.3H3z" fill="none" stroke="currentColor" stroke-width="1.1" stroke-linejoin="round"/><path d="M4.5 6h3M4.5 8h3" stroke="currentColor" stroke-width="1"/></svg>',
    web: '<svg viewBox="0 0 12 12" aria-hidden="true"><circle cx="6" cy="6" r="4.6" fill="none" stroke="currentColor" stroke-width="1.1"/><path d="M1.6 6h8.8M6 1.4c-2 2.6-2 6.6 0 9.2M6 1.4c2 2.6 2 6.6 0 9.2" fill="none" stroke="currentColor" stroke-width="1"/></svg>',
    letter: '<svg viewBox="0 0 12 12" aria-hidden="true"><circle cx="6" cy="6" r="4.6" fill="none" stroke="currentColor" stroke-width="1.1"/><path d="M6 3.4V6l1.8 1.2" fill="none" stroke="currentColor" stroke-width="1.1" stroke-linecap="round"/></svg>',
    check: '<svg viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 6.3l2.3 2.3 4.7-5" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    arrow: '<svg viewBox="0 0 16 10" aria-hidden="true"><path d="M1 5h13M10 1.5L14 5l-4 3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    lock: '<svg viewBox="0 0 12 12" aria-hidden="true"><rect x="2.5" y="5.2" width="7" height="5.3" rx="1" fill="none" stroke="currentColor" stroke-width="1.1"/><path d="M4 5.2V3.8a2 2 0 0 1 4 0v1.4" fill="none" stroke="currentColor" stroke-width="1.1"/></svg>',
    play: '<svg viewBox="0 0 12 12" aria-hidden="true"><path d="M3.5 2.2v7.6L9.6 6z" fill="currentColor"/></svg>'
  };

  /* ---------- helpers ---------- */

  function clampStep(v) {
    var n = parseInt(v, 10);
    if (isNaN(n)) return 0;
    return Math.max(0, Math.min(STEP_COUNT - 1, n));
  }
  function esc(s) {
    return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
  }
  function mmss(t) { return Math.floor(t / 60) + ':' + ('0' + (t % 60)).slice(-2); }
  function fileById(id) {
    for (var i = 0; i < P.vault.length; i++) for (var j = 0; j < P.vault[i].files.length; j++) if (P.vault[i].files[j].id === id) return P.vault[i].files[j];
    return null;
  }
  function sourcesOfFile(fid) { return Object.keys(S).filter(function (k) { return S[k].file === fid; }); }
  function claimsCiting(ids) { return P.letter.filter(function (c) { return c.cites.some(function (x) { return ids.indexOf(x) >= 0; }); }); }
  function plain(c) { return c.text.replace(/\{tag:[A-Z0-9]+\|([^}]+)\}/g, '$1'); }

  // XIRR over the cash-flow rows plus NAV as the terminal value (actual/365, bisection).
  function cashRows() { return P.cashFlows.concat([['2026-09-30', F.nav, 'NAV, September 30 (terminal value)']]); }
  function xirr(rows) {
    var t0 = Date.parse(rows[0][0]);
    function xnpv(r) { return rows.reduce(function (s, f) { return s + f[1] / Math.pow(1 + r, (Date.parse(f[0]) - t0) / 864e5 / 365); }, 0); }
    var lo = -0.9, hi = 1;
    for (var i = 0; i < 100; i++) { var m = (lo + hi) / 2; if (xnpv(m) > 0) lo = m; else hi = m; }
    return (lo + hi) / 2;
  }

  function chip(kind, why) {
    var l = LEGEND[kind];
    var title = l.label + ': ' + (why || l.note);
    return '<span class="chip chip-' + kind + '" data-legend="' + kind + '" title="' + esc(title) + '"><i></i>' + l.label + '</span>';
  }
  function shortLabel(id) {
    var s = S[id];
    if (s.kind === 'excel') return s.cell;
    if (s.kind === 'pdf') return 'p.' + s.page;
    if (s.kind === 'voice') return id + ' ' + mmss(s.t);
    if (s.kind === 'web') return 'Web';
    return 'Q2 letter';
  }
  function sourceTitle(id) {
    var s = S[id], f = fileById(s.file);
    if (s.kind === 'excel') return f.name + ' › ' + s.sheet + '!' + s.cell;
    if (s.kind === 'pdf') return f.name + ', page ' + s.page;
    if (s.kind === 'voice') return f.name + ', ' + mmss(s.t);
    return f.name;
  }
  function cite(id) {
    var s = S[id];
    return '<button type="button" class="cite cite-' + s.kind + (state.sel === id ? ' is-active' : '') + '" data-cite="' + id + '" title="' + esc(sourceTitle(id)) + '" aria-pressed="' + (state.sel === id) + '">' + ICON[s.kind] + esc(shortLabel(id)) + '</button>';
  }
  function tag(id, text) {
    return '<button type="button" class="tag' + (state.sel === id ? ' is-active' : '') + '" data-cite="' + id + '" title="' + esc('Data tag: ' + sourceTitle(id)) + '">' + esc(text) + '</button>';
  }
  function sentenceHTML(c) {
    var body = esc(c.text).replace(/\{tag:([A-Z0-9]+)\|([^}]+)\}/g, function (m, id, t) { return tag(id, t); });
    var on = state.drawer === c.id;
    return '<span class="sent' + (on ? ' is-traced' : '') + '" data-claim="' + c.id + '">' + body + '<span class="cites">' + c.cites.map(cite).join('') + '</span></span> ';
  }

  function status(id) {
    if (id === 'W1' && !state.approved) return { k: 'review', t: 'Needs review', d: 'Market context from one trade-press source. A reviewer confirms it before approval.' };
    if (id === 'W1') return { k: 'ok', t: 'Verified', d: 'Confirmed by ' + P.people.ir.name + ' as market context.' };
    if (id === 'D1' || id === 'T3') return { k: 'ok', t: 'Verified', d: 'Conflict resolved ' + P.edit.when + ': the draft said “' + P.letter[3].draft + '”, which this source does not support. ' + P.edit.reviewer + ' corrected it (' + P.edit.type.toLowerCase() + ').' };
    if (id === 'M1') return { k: 'ok', t: 'Verified', d: 'Recomputed from the cited range: ' + (xirr(cashRows()) * 100).toFixed(2) + '%, matching registry value ' + F.irr.toFixed(1) + '%.' };
    if (id === 'M3') return { k: 'ok', t: 'Verified', d: '118.0 × (1 − 8.0%) = ' + (118 * 0.92).toFixed(2) + ', reported as $108.6M.' };
    return { k: 'ok', t: 'Verified', d: 'Each citing sentence was checked against this source.' };
  }

  /* ---------- views ---------- */

  var NAV = [['connections', 'Connections'], ['vault', 'Vault'], ['templates', 'Templates'], ['graph', 'How it’s built'], ['outputs', 'Outputs']];
  function navKey() { return state.view === 'ingest' ? 'vault' : state.view === 'detail' ? 'outputs' : state.view; }
  function sidebar() {
    return '<div class="firm"><span class="avatar">MC</span><div><b>' + esc(P.firm) + '</b><span>Signed in · Sam Ortiz</span></div></div>' +
      '<nav class="side-nav" aria-label="FinanceOS">' + NAV.map(function (it) {
        var on = navKey() === it[0];
        return '<button type="button" class="side-item' + (on ? ' on' : '') + '" data-view="' + it[0] + '"' + (on ? ' aria-current="page"' : '') + '>' + it[1] + '</button>';
      }).join('') + '</nav>' +
      '<div class="side-status"><div class="label">' + P.package.name + '</div><p><b>' + (state.approved ? 'Approved v3' : state.generated ? 'Ready for review' : 'In progress') + '</b></p></div>';
  }
  function head(kicker, title, chips) {
    return '<header class="view-head"><div><div class="label">' + kicker + '</div><h2>' + title + '</h2></div><div class="head-chips">' + chips + '</div></header>';
  }
  function allFiles() { return P.vault.reduce(function (a, g) { return a.concat(g.files.map(function (f) { f.type = g.type; return f; })); }, []); }

  function connectionsView() {
    var conns = [
      ['Google Drive', 'Meridian › Fund II, Portfolio', 'Connected Sep 2, 2026 · last sync Oct 16, 6:00 PM', '214 files watched'],
      ['SharePoint data room', 'Northwind sale process', 'Connected Sep 14, 2026 · last sync Oct 16, 5:45 PM', '38 files watched'],
      ['Voice capture', 'Partner interviews', 'Requests sent from FinanceOS · last recording Oct 21', '1 recording this quarter']
    ];
    var engine = [
      ['PDF', 'Parse page layout, tables and reading order; keep page coordinates for every passage', '46 pages', chip('proposed', 'PDF layout parsing')],
      ['Excel', 'Read every cell, formula and named range; build the dependency graph', '13,686 cells', chip('today', 'Excel-based agents read the workbooks')],
      ['Audio', 'Transcribe with word timestamps and speaker', '1:12, 186 words', chip('today', 'Voice interview capture')],
      ['Web', 'Snapshot the page with its capture date and URL', '1 snapshot', chip('proposed', 'Dated web snapshots')]
    ];
    return '<section class="view" data-element="connections">' + head('Connections', 'Sources stay where the firm keeps them', chip('proposed', 'Connectors kept in sync')) +
      '<div class="conn-grid">' + conns.map(function (c) { return '<article class="conn"><div class="conn-top"><b>' + c[0] + '</b><span class="st st-ok">' + ICON.check + 'Syncing</span></div><p>' + c[1] + '</p><p class="muted small">' + c[2] + '</p><p class="conn-count">' + c[3] + '</p></article>'; }).join('') + '</div>' +
      '<section class="engine" data-element="engine"><header class="view-head small"><div><div class="label">Ingestion engine</div><h3>What happens to each file type</h3></div>' + chip('proposed', 'Ingestion engine') + '</header>' +
      '<table class="tbl"><thead><tr><th>Type</th><th>What the engine does</th><th>This quarter</th><th>Status</th></tr></thead><tbody>' + engine.map(function (e) { return '<tr><td><b>' + e[0] + '</b></td><td>' + e[1] + '</td><td class="num">' + e[2] + '</td><td>' + e[3] + '</td></tr>'; }).join('') + '</tbody></table></section></section>';
  }

  function vaultView() {
    var q = [
      ['Q3 2026', 'New · in progress', 'q3', '8 files synced · package not yet generated', 'st-review'],
      ['Q2 2026', 'Generated and approved', 'q2', '4 outputs · approved by Priya Raman, Aug 7', 'st-ok'],
      ['Q1 2026', 'Generated and approved', 'q1', '4 outputs · approved May 6', 'st-ok']
    ];
    return '<section class="view" data-element="quarters">' + head('Vault · Fund II', 'Quarters', chip('proposed', 'Source vault organised by reporting period')) +
      '<div class="quarters">' + q.map(function (x) {
        var open = x[2] === 'q3';
        return '<button type="button" class="quarter' + (open ? ' is-new' : '') + '"' + (open ? ' data-view="ingest"' : ' disabled') + '><span class="folder" aria-hidden="true"></span><b>' + x[0] + '</b><span class="st ' + x[4] + '">' + x[1] + '</span><span class="muted">' + x[3] + '</span>' + (open ? '<span class="open">Open ' + ICON.arrow + '</span>' : '') + '</button>';
      }).join('') + '</div></section>';
  }

  function ingestView() {
    var files = allFiles(), figs = files.reduce(function (s, f) { return s + f.figures; }, 0);
    return '<section class="view" data-element="ingest">' + head('Vault › Q3 2026', files.length + ' files ingested, ' + figs + ' figures registered', chip('today', 'Excel-based agents; voice interview capture') + chip('proposed', 'Ingestion detail and figure registry')) +
      '<table class="tbl ingest"><thead><tr><th>File</th><th>Source</th><th>Synced</th><th>Extracted</th><th class="num">Figures</th></tr></thead><tbody>' + files.map(function (f) {
        var k = S[sourcesOfFile(f.id)[0]].kind;
        return '<tr><td><button type="button" class="file-link" data-file="' + f.id + '"><span class="ficon">' + ICON[k] + '</span><span><b>' + esc(f.name) + '</b><em>' + esc(f.type) + '</em></span></button></td><td class="muted">' + esc(f.path) + '</td><td>' + esc(f.synced) + '</td><td>' + esc(f.extracted) + '</td><td class="num">' + f.figures + '</td></tr>';
      }).join('') + '</tbody></table></section>';
  }

  function templatesView() {
    var outs = [['LP letter', '.docx', '6 pages'], ['LP deck', '.pptx', '24 slides'], ['Performance schedules', '.xlsx', '4 tabs'], ['Capital account statements', '.pdf', 'one per LP']];
    var vars = [['{Fund II Net IRR}', 'Valuation model › Summary!F18', 'M1'], ['{Fund II NAV}', 'Valuation model › Summary!F12', 'M4'], ['{Fund II Distributions}', 'Capital accounts › LP_Capital!D14', 'A1'], ['{Prior Net IRR}', 'Q2 2026 LP letter, approved v4', 'P1'], ['{Holding fair values}', 'Valuation model › Valuation!H8:H40', 'M3'], ['{Value change explanation}', 'Partner interview, requested when the rule fires', 'T1']];
    var sections = ['Fund overview', 'Fund II performance', 'Portfolio updates', 'Valuation changes', 'Capital activity', 'Outlook'];
    return '<section class="view" data-element="template">' + head('Templates · configured for ' + esc(P.shortFirm), 'Meridian LP package 2026', chip('proposed', 'Configured per customer, no code') + chip('today', 'Brand-compliant output across Word, PPT and Excel')) +
      '<div class="tpl-outs">' + outs.map(function (o, i) { return '<div class="tpl-out' + (i === 0 ? ' on' : '') + '"><b>' + o[0] + '</b><span>' + o[1] + ' · ' + o[2] + '</span></div>'; }).join('') + '</div>' +
      '<div class="tpl-grid"><div><div class="label">LP letter sections</div><ol class="tpl-sections">' + SECTION_KEYS.map(function (s, i) { return '<li' + (i === 1 ? ' class="on"' : '') + '><button type="button" class="tpl-sec" data-recipe="' + s[0] + '">' + s[1] + '</button></li>'; }).join('') + '</ol><p class="muted small">Click a section to see its recipe.</p></div>' +
      '<div data-element="variables"><div class="label">Data tags and source bindings ' + chip('proposed', 'Figure registry bindings') + '</div><table class="tbl"><thead><tr><th>Tag</th><th>Bound to</th><th>Value</th></tr></thead><tbody>' + vars.map(function (v) { return '<tr><td><code>' + esc(v[0]) + '</code></td><td>' + esc(v[1]) + '</td><td>' + cite(v[2]) + '</td></tr>'; }).join('') + '</tbody></table></div>' +
      '<div><div class="label">Rules</div><dl class="rules">' + P.config.map(function (k) { return '<div><dt>' + esc(k.k) + '</dt><dd>' + esc(k.v) + '</dd></div>'; }).join('') + '</dl></div></div></section>';
  }

  function graphView() {
    var docs = [['model', 'Valuation model'], ['capital', 'Capital accounts'], ['northwind', 'Northwind board deck'], ['cobalt', 'Cobalt board deck'], ['covenant', 'Covenant amendment'], ['web', 'Web snapshot'], ['prior', 'Q2 LP letter'], ['voice', 'Partner interview']];
    var figs = [['irr', 'Net IRR 11.8%'], ['chg', 'IRR change −240 bps'], ['nav', 'NAV $821.1M'], ['nw', 'Northwind $249.4M'], ['co', 'Cobalt $108.6M'], ['dist', 'Distributions $224.0M']];
    var secs = [['perf', 'Fund II performance'], ['val', 'Valuation changes'], ['port', 'Portfolio updates'], ['cap', 'Capital activity']];
    var edges = [['model', 'irr'], ['prior', 'chg'], ['model', 'chg'], ['voice', 'chg', 'warn'], ['model', 'nav'], ['capital', 'nav', 'bad'], ['model', 'nw'], ['northwind', 'nw'], ['model', 'co'], ['covenant', 'co'], ['cobalt', 'co'], ['capital', 'dist'], ['web', 'nw'],
      ['irr', 'perf'], ['chg', 'perf'], ['nav', 'perf'], ['nav', 'cap'], ['nw', 'val'], ['co', 'val'], ['nw', 'port'], ['co', 'port'], ['dist', 'cap']];
    var pos = {}, W = 1000, rowH = 52;
    function col(list, x, top) { list.forEach(function (n, i) { pos[n[0]] = { x: x, y: top + i * rowH, label: n[1] }; }); }
    col(docs, 10, 20); col(figs, 380, 72); col(secs, 750, 124);
    var svg = '<svg class="graph" viewBox="0 0 ' + W + ' ' + (20 + docs.length * rowH) + '" role="img" aria-label="Document, figure and section graph">';
    edges.forEach(function (e) {
      var a = pos[e[0]], b = pos[e[1]], x1 = a.x + 230, y1 = a.y + 18, x2 = b.x, y2 = b.y + 18, mx = (x1 + x2) / 2;
      svg += '<path class="edge ' + (e[2] || '') + '" d="M' + x1 + ' ' + y1 + ' C' + mx + ' ' + y1 + ' ' + mx + ' ' + y2 + ' ' + x2 + ' ' + y2 + '"/>';
    });
    Object.keys(pos).forEach(function (k) {
      var n = pos[k], cls = k === 'nav' ? ' bad' : k === 'chg' ? ' warn' : docs.some(function (d) { return d[0] === k; }) ? ' doc' : secs.some(function (s) { return s[0] === k; }) ? ' sec' : ' fig';
      svg += '<g class="node' + cls + '"><rect x="' + n.x + '" y="' + n.y + '" width="230" height="36" rx="2"/><text x="' + (n.x + 12) + '" y="' + (n.y + 23) + '">' + esc(n.label) + '</text></g>';
    });
    svg += '<text class="col-label" x="10" y="12">Vault documents</text><text class="col-label" x="380" y="64">Figures (registry)</text><text class="col-label" x="750" y="116">Template sections</text></svg>';
    var issues = '<aside class="issues" data-element="discrepancy"><header class="view-head small"><div><div class="label">Flags</div><h3>2 flags, both resolved</h3></div>' + chip('proposed', 'Verification against the registry') + '</header>' +
      '<article class="issue bad"><span class="st st-conflict">Seeded discrepancy</span><b>Fund II NAV disagrees</b><p>Valuation model ' + cite('M4') + ' $821.1M; fund-admin statement ' + cite('A2') + ' $830.5M. The administrator had not booked the $9.4M Cobalt markdown ' + cite('M3') + '.</p><p class="res"><b>Resolved.</b> The valuation model is the source of truth for NAV; note added for the administrator. Reviewer: Sam Ortiz, Oct 20.</p><p class="muted small">Planted in the sample data to show the check working.</p></article>' +
      '<article class="issue warn"><span class="st st-review">Needs explanation</span><b>Net IRR 14.2% → 11.8%, −240 bps</b><p>Over the 100 bps rule ' + cite('M1') + ' ' + cite('P1') + '.</p><p class="res"><b>Resolved.</b> Dana Whitfield’s voice note explains it ' + cite('T1') + ' ' + cite('T3') + ', Oct 21.</p></article></aside>';
    return '<div class="graph-wrap" data-element="graph"><div>' + chip('proposed', 'Evidence graph across documents, figures and sections') + svg + '</div>' + issues + '</div>';
  }

  /* ---------- section recipes ---------- */

  var SECTION_KEYS = [['overview', 'Fund overview'], ['perf', 'Fund II performance'], ['port', 'Portfolio updates'], ['val', 'Valuation changes'], ['cap', 'Capital activity'], ['outlook', 'Outlook']];
  function sectionName(k) { return SECTION_KEYS.filter(function (s) { return s[0] === k; })[0][1]; }
  function sectionStats(k) { var n = sectionName(k); return P.sections.filter(function (s) { return s.name === n; })[0]; }
  function recipe(k) {
    var bps = Math.round((F.irr - F.irrPrior) * 100), tvpi = (F.distributions + F.nav) / F.paidIn, dpi = F.distributions / F.paidIn;
    var cobalt = ((108.6 / 118.0 - 1) * 100).toFixed(1).replace('-', MINUS);
    var house = 'House style from the Q2 2026 LP letter, approved v4';
    var chain = 'Approval chain: ' + P.config.filter(function (c) { return c.k === 'Approval chain'; })[0].v;
    var R = {
      perf: { sources: ['M1', 'P1', 'T1', 'T2', 'T3', 'T4', 'D1', 'D2', 'D3', 'W1'], figures: [['M1', 'Net IRR, Q3', '11.8%'], ['P1', 'Net IRR, Q2', '14.2%'], ['M2', 'Northwind fair value', '$249.4M'], ['M3', 'Cobalt fair value', '$108.6M']], prior: S.P1.quote,
        rules: ['Explain any net IRR move of 100 bps or more, or a NAV move of 5% or more', '60 words max for the change explanation', 'Name the holdings that drive the change', house, chain],
        derived: [MINUS + Math.abs(bps) + ' bps = ' + F.irr.toFixed(1) + '% ' + MINUS + ' ' + F.irrPrior.toFixed(1) + '%, over the 100 bps rule', 'Cobalt ' + cobalt + '% = 108.6 ÷ 118.0 ' + MINUS + ' 1', 'Northwind ' + MINUS + '12.6 = 249.4 ' + MINUS + ' 262.0'],
        retrieve: '8 passages: partner interview T1–T4, Northwind deck p.6, Cobalt deck p.4, covenant amendment p.2, trade press', draft: '6 cited sentences (C1–C6)', verify: 'Numbers match the registry; each claim checked against its passage; 1 conflict caught in C4 (“covenant breach”)', review: P.edit.reviewer + ' corrected C4 as a factual correction, ' + P.edit.when,
        output: ['LP letter, page 2, Fund II paragraph', 'LP deck, slide 7 commentary', 'Performance schedules, note (1)'] },
      overview: { sources: ['A1', 'M4', 'P1'], figures: [['M4', 'Fund II NAV', '$821.1M'], ['A1', 'Distributions since inception', '$224.0M']], prior: 'Q2 2026 letter, fund overview: one line per fund with NAV and net IRR.',
        rules: ['One line per fund, in vintage order', house, chain], derived: ['Net TVPI ' + tvpi.toFixed(2) + 'x = (224.0 + 821.1) ÷ 722.0'], retrieve: 'Fund summaries from the valuation model and capital accounts', draft: 'One cited line per fund', verify: 'Every figure matches the registry', review: P.people.ir.name + ' accepted without edit', output: ['LP letter, page 1', 'LP deck, slide 3'] },
      port: { sources: ['D2', 'D3', 'W1'], figures: [['M2', 'Northwind fair value', '$249.4M'], ['M3', 'Cobalt fair value', '$108.6M']], prior: 'Q2 2026 letter, portfolio updates for each holding.',
        rules: ['One paragraph per holding with an operating update', 'Cite the board deck for every operating claim', house, chain], derived: [], retrieve: 'Board-deck passages for each holding', draft: 'One cited paragraph per holding', verify: 'Each claim checked against its board-deck passage', review: P.people.ir.name + ' accepted without edit', output: ['LP letter, pages 3–4', 'LP deck, slides 9–16'] },
      val: { sources: ['M2', 'M3', 'M4', 'A2', 'D1'], figures: [['M2', 'Northwind fair value', '$249.4M'], ['M3', 'Cobalt fair value', '$108.6M'], ['M4', 'Fund II NAV', '$821.1M']], prior: 'Q2 2026 letter, valuation changes table.',
        rules: ['List every holding whose value moved 4% or more', 'Use the valuation model as the source of truth for NAV', house, chain], derived: ['Cobalt ' + cobalt + '%', 'NAV ' + MINUS + (F.navPrior - F.nav).toFixed(1) + ' = 821.1 ' + MINUS + ' 844.6'], retrieve: 'Valuation sheet rows, the administrator statement and the covenant amendment', draft: 'Valuation-change table with cited notes', verify: 'Seeded NAV discrepancy against the administrator caught and resolved', review: P.people.ir.name + ' added a note for the administrator', output: ['LP letter, page 4', 'Performance schedules, NAV Bridge tab'] },
      cap: { sources: ['A1'], figures: [['A1', 'Paid-in capital', '$722.0M'], ['A1', 'Distributions', '$224.0M']], prior: 'Q2 2026 letter, capital activity.',
        rules: ['State calls and distributions for the quarter', house, chain], derived: ['DPI ' + dpi.toFixed(2) + 'x = 224.0 ÷ 722.0', 'Net TVPI ' + tvpi.toFixed(2) + 'x'], retrieve: 'Capital account ledger', draft: '2 cited sentences (C7–C8)', verify: 'Every figure matches the registry', review: P.people.ir.name + ' accepted without edit', output: ['LP letter, page 2', 'Capital account statements, one per LP'] },
      outlook: { sources: ['W1', 'T4'], figures: [], prior: 'Q2 2026 letter, outlook.',
        rules: ['No forecasts of returns', 'Market context must cite a dated source', house, chain], derived: [], retrieve: 'Trade press and partner commentary', draft: 'Short cited outlook paragraph', verify: '1 market claim needs reviewer confirmation', review: 'Open: ' + P.people.ir.name + ' to confirm the market claim', output: ['LP letter, page 6'] }
    };
    return R[k];
  }
  function recipeHTML(k) {
    var r = recipe(k), st = sectionStats(k), ok = st.verified === st.total;
    var steps = [
      ['Retrieve passages', r.retrieve, 'proposed'],
      ['Pull registered figures', r.figures.length ? r.figures.map(function (x) { return x[1] + ' ' + x[2]; }).join(' · ') : 'No figures in this section', 'proposed'],
      ['Compute derived values', r.derived.length ? r.derived.join(' · ') : 'None needed', 'proposed'],
      ['Draft with citations', r.draft, 'today'],
      ['Verify', r.verify, 'proposed'],
      ['Human review', r.review, 'today'],
      ['Approve', k === 'outlook' ? 'Waiting on review' : P.people.partner.name + ' signed off ' + P.timeline.signedOff + (state.approved ? '; CFO approved' : '; CFO approval pending'), 'proposed']
    ];
    var why = { today: 'Shared ML service drafting through the DAG; content feedback capture', proposed: 'Recipe engine and figure registry' };
    return '<section class="recipe" data-element="recipe"><header class="view-head small"><div><div class="label">Recipe</div><h3>' + esc(sectionName(k)) + '</h3></div>' + chip('proposed', 'Recipe engine and figure registry') + '</header>' +
      '<div class="rc-grid"><div class="rc-block"><div class="label">Ingredients</div><p class="rc-k">Vault sources</p><div class="rc-chips">' + r.sources.map(cite).join('') + chip('today', 'Source-linked citation chips') + '</div>' +
      (r.figures.length ? '<p class="rc-k">Registry figures</p><ul class="rc-figs">' + r.figures.map(function (x) { return '<li><span>' + esc(x[1]) + '</span>' + tag(x[0], x[2]) + '</li>'; }).join('') + '</ul>' : '') +
      '<p class="rc-k">Prior-period text it continues</p><p class="rc-prior">' + esc(r.prior) + '</p></div>' +
      '<div class="rc-block"><div class="label">Rules · configured for ' + esc(P.shortFirm) + '</div><ul class="rc-rules">' + r.rules.map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') + '</ul>' +
      '<div class="label rc-out-label">Output</div><ul class="rc-rules">' + r.output.map(function (o) { return '<li><b>' + esc(o) + '</b></li>'; }).join('') + '</ul><p><span class="st ' + (ok ? 'st-ok' : 'st-review') + '">' + (ok ? 'Verified' : 'Needs review') + '</span> ' + st.verified + ' of ' + st.total + ' claims verified</p></div></div>' +
      '<div class="rc-block"><div class="label">Steps</div><ol class="rc-steps">' + steps.map(function (s, i) { return '<li><span class="stage-n">' + (i + 1) + '</span><b>' + s[0] + '</b><span>' + esc(s[1]) + '</span>' + chip(s[2], why[s[2]]) + '</li>'; }).join('') + '</ol></div></section>';
  }
  function builtView() {
    var picker = '<nav class="rc-picker" data-element="sections" aria-label="Sections">' + chip('proposed', 'Configured template sections') + SECTION_KEYS.map(function (s) {
      var on = state.recipe === s[0];
      return '<button type="button" class="rc-pick' + (on ? ' on' : '') + '" data-recipe="' + s[0] + '"' + (on ? ' aria-current="true"' : '') + '>' + s[1] + '</button>';
    }).join('') + '</nav>';
    var g = state.gtab === 'graph';
    var tabs = '<div class="drawer-tabs gtabs" role="tablist"><button type="button" role="tab" class="dtab' + (g ? '' : ' on') + '" aria-selected="' + !g + '" data-gtab="recipe">Recipe</button><button type="button" role="tab" class="dtab' + (g ? ' on' : '') + '" aria-selected="' + g + '" data-gtab="graph">Graph view</button></div>';
    return '<section class="view view-built" data-element="built">' + head('How it’s built · ' + P.package.name, g ? 'How evidence flows into the report' : 'The recipe for each section', chip('proposed', 'Recipe engine and figure registry')) + tabs + (g ? graphView() : '<div class="rc-wrap">' + picker + recipeHTML(state.recipe) + '</div>') + '</section>';
  }

  var STAGES = [['Ingest', '8 files parsed'], ['Index', '1,912 passages'], ['Register figures', P.package.figures + ' figures'], ['Draft from cited evidence', '214 claims'], ['Trace sentences', 'every sentence to its span'], ['Verify', '214 checked · 2 flags resolved'], ['Ready for review', '4 files']];
  function outputsView() {
    var done = state.generated ? STAGES.length : state.running;
    var stages = '<ol class="progress">' + STAGES.map(function (s, i) {
      var st = i < done ? 'done' : i === done ? 'run' : 'wait';
      return '<li class="' + st + '"><span class="dotp">' + (i < done ? ICON.check : i + 1) + '</span><b>' + s[0] + '</b><em>' + (i < done ? s[1] : i === done ? 'Running' : 'Waiting') + '</em></li>';
    }).join('') + '</ol>';
    var files = [['LP letter', '.docx', '6 pages', '96 of 96 claims verified', 'letter'], ['LP deck', '.pptx', '24 slides', '61 of 61 claims verified', ''], ['Performance schedules', '.xlsx', '4 tabs · 186 figures from the registry', '33 of 33 note claims verified', ''], ['Capital account statements', '.pdf', '1 per LP · 38 figures from the registry', '24 of 24 claims verified', '']];
    var out = '<section class="view" data-element="generate">' + head('Outputs · ' + P.package.name, state.generated ? 'Generated, ready for review' : 'Generate the Q3 package', chip('today', 'Shared ML service orchestrating through a DAG') + chip('proposed', 'Trace and verify stages')) +
      '<div class="gen-row"><button type="button" class="btn primary" data-generate="1"' + (state.running >= 0 && !state.generated ? ' disabled' : '') + '>' + (state.generated ? 'Regenerate' : 'Generate') + '</button>' + stages + '</div>';
    if (state.generated) {
      out += '<div class="files" data-element="files">' + chip('today', 'Brand-compliant output across Word, PPT and Excel') + '<div class="file-cards">' + files.map(function (f) {
        return '<button type="button" class="file-card"' + (f[4] ? ' data-view="detail"' : ' disabled') + '><b>' + f[0] + ' <em>' + f[1] + '</em></b><span>' + f[2] + '</span><span class="st st-ok">' + f[3] + '</span>' + (f[4] ? '<span class="open">Open ' + ICON.arrow + '</span>' : '') + '</button>';
      }).join('') + '</div></div>' +
        '<div class="consistency" data-element="consistency"><header class="view-head small"><div><div class="label">Consistency check</div><h3>The same figures everywhere</h3></div>' + chip('proposed', 'Render once from the registry') + '</header><table class="tbl"><thead><tr><th>Figure</th><th>Letter</th><th>Deck</th><th>Schedules</th><th>Statements</th></tr></thead><tbody>' +
        [['Net IRR 11.8%', 'M1', 1, 1, 1, 0], ['NAV $821.1M', 'M4', 1, 1, 1, 1], ['Distributions $224.0M', 'A1', 1, 1, 1, 1], ['Net TVPI 1.45x', 'A1', 0, 1, 1, 0]].map(function (r) {
          return '<tr><td>' + r[0] + ' ' + cite(r[1]) + '</td>' + r.slice(2).map(function (v) { return '<td>' + (v ? '<span class="ok">' + ICON.check + 'Match</span>' : '<span class="muted">Not shown</span>') + '</td>'; }).join('') + '</tr>';
        }).join('') + '</tbody></table></div>';
    }
    return out + '</section>';
  }

  function letterDoc() {
    var fundII = P.letter.slice(0, 6), capital = P.letter.slice(6);
    return '<article class="paper letter-paper" data-doc="letter"><header class="paper-head"><span class="paper-mark">' + esc(P.firm) + '</span><span>Letter to Limited Partners · ' + P.period.quarterWord + '</span></header>' +
      '<h4><button type="button" class="sec-head" data-recipe-drawer="perf" title="See how this section is built">' + esc(F.short) + ' performance</button></h4><p>' + fundII.map(sentenceHTML).join('') + '</p>' +
      '<h4><button type="button" class="sec-head" data-recipe-drawer="cap" title="See how this section is built">Capital activity</button></h4><p>' + capital.map(sentenceHTML).join('') + '</p>' +
      '<footer class="paper-foot"><span>Confidential</span><span>Page 2</span></footer></article>';
  }

  function detailView() {
    var approval = '<div class="approval" data-element="approval"><span class="st ' + (state.approved ? 'st-ok' : 'st-review') + '">' + (state.approved ? 'Approved v3' : 'In review v2') + '</span>' +
      '<span>' + esc(P.people.ir.name) + ' prepared · ' + esc(P.people.partner.name) + ' signed off · ' + (state.approved ? esc(P.people.cfo.name) + ' approved ' + P.timeline.approved : esc(P.people.cfo.name) + ' to approve') + '</span>' +
      (state.approved ? '' : '<button type="button" class="btn" data-approve="1">Approve as CFO</button>') +
      '<button type="button" class="btn ghost" data-history="1">Version history</button><button type="button" class="btn ghost" data-drawer="BOUNDARY">Product boundary</button>' + chip('proposed', 'Approval chain with version history') + '</div>';
    var hist = state.history ? '<ol class="versions"><li><b>v1</b> Drafted by the shared ML service · ' + P.timeline.drafted + '</li><li><b>v2</b> ' + esc(P.edit.reviewer) + ' corrected C4 (factual) · ' + P.timeline.edited + '</li><li><b>v3</b> ' + (state.approved ? 'Approved by ' + esc(P.people.cfo.name) + ' · ' + P.timeline.approved : 'Awaiting CFO approval') + '</li></ol>' : '';
    return '<section class="view view-detail" data-element="letter">' + head('Outputs › LP letter.docx', esc(F.short) + ' performance', chip('today', 'Source-linked citation chips')) + approval + hist +
      '<p class="hint">Click any sentence or figure to see its sources.</p><div class="doc-wrap">' + letterDoc() + '</div></section>';
  }

  function boundaryHTML() {
    function item(text, kind) { return '<li><span>' + text + '</span>' + chip(kind) + '</li>'; }
    return '<section class="boundary" data-element="boundary"><header class="view-head small"><div><div class="label">Product boundary</div><h3>Shared building blocks, one product on top, the rest configured</h3></div>' + chip('proposed', 'Proposed product boundary') + '</header>' +
      '<div class="b-block"><b>Shared building blocks</b><span class="muted">Every segment, built once</span><ul>' +
      item('Source vault kept in sync', 'proposed') + item('Figure registry', 'proposed') + item('Source-linked citations', 'today') + item('Voice interview capture', 'today') +
      item('Brand-compliant output across Word, PPT and Excel', 'today') + item('Review and approval with version history', 'proposed') + item('Feedback capture', 'today') + item('Evaluation harness', 'proposed') + '</ul></div>' +
      '<div class="b-block"><b>Investor-reporting product</b><span class="muted">Private equity and real estate</span><ul>' +
      item('Report object model: fund, asset, period, metric, commentary, source', 'proposed') + item('Prior-period diff with the explain-the-change rule', 'proposed') + item('LP package templates rendered once to letter, deck, schedules and statements', 'proposed') + '</ul></div>' +
      '<div class="b-block"><b>Configured for ' + esc(P.shortFirm) + '</b><span class="muted">Per customer, no code</span><dl>' + P.config.map(function (k) { return '<div><dt>' + esc(k.k) + '</dt><dd>' + esc(k.v) + '</dd></div>'; }).join('') + '</dl></div>' +
      '<div class="b-block"><b>Bespoke services</b><span class="muted">Priced separately</span><ul class="plain">' + P.bespoke.map(function (b) { return '<li>' + esc(b) + '</li>'; }).join('') + '</ul></div>' +
      '<div class="b-block"><b>Not this year</b><p class="muted">A hospitality product · all enterprise reporting · a general workflow platform · per-customer forks · real-time dashboards · self-serve template builder · external LP portal</p></div>' +
      '<div class="b-banner">' + chip('assumption', 'Pipeline assumption; thresholds are provisional') + '<p>Assumes a pipeline across all three segments; with only these three customers, build nothing shared and keep working out go-to-market. A retrospective on two past quarters must reach about 75% of content from configuration before one paid pilot; under about 60%, stop the product. Thresholds are provisional.</p></div></section>';
  }

  /* ---------- source viewer ---------- */

  function citedBy(id) {
    var cl = claimsCiting([id]);
    return '<div class="in-block"><div class="label">Cited by</div>' + (cl.length ? '<ol class="in-claims">' + cl.map(function (c) { return '<li><span class="in-cid">' + c.id + '</span><span>' + esc(plain(c)) + '</span></li>'; }).join('') + '</ol>' : '<p class="muted">No sentence in this section cites it.</p>') + '</div>';
  }
  function statusBlock(id) {
    var st = status(id);
    return '<div class="in-block in-status"><span class="st st-' + st.k + '">' + st.t + '</span><p>' + esc(st.d) + '</p></div>';
  }
  function srcHead(id, kindLabel) {
    var f = fileById(S[id].file);
    return '<header class="in-head"><div class="label">Source viewer · ' + kindLabel + '</div><h3>' + esc(f.name) + '</h3><button type="button" class="in-link" data-view="ingest">Show in Q3 vault</button></header>';
  }

  function waveform(at) {
    var bars = '';
    for (var i = 0; i < 120; i++) {
      var tSec = i / 120 * AUDIO_SECONDS;
      var speaking = [[4, 13.4], [19, 25.4], [41, 48.3], [58, 62.2]].some(function (r) { return tSec >= r[0] && tSec <= r[1]; });
      var h = speaking ? 6 + Math.abs(Math.sin(i * 1.7) * 14 + Math.sin(i * 0.43) * 8) : 2;
      bars += '<rect x="' + (i * 3) + '" y="' + (24 - h / 2).toFixed(1) + '" width="2" height="' + h.toFixed(1) + '"/>';
    }
    var x = (at / AUDIO_SECONDS * 360).toFixed(1);
    return '<svg class="wave" viewBox="0 0 360 48" role="img" aria-label="Waveform, marker at ' + mmss(at) + '">' + bars + '<line x1="' + x + '" x2="' + x + '" y1="0" y2="48" class="wave-mark"/></svg>';
  }
  function voiceViewer(id) {
    var s = S[id];
    var lines = Object.keys(S).filter(function (k) { return S[k].kind === 'voice'; });
    return '<section class="viewer viewer-voice" data-element="viewer-voice">' + srcHead(id, 'Voice') + chip('today', 'Voice interview capture') +
      '<div class="in-block"><div class="player"><audio controls preload="none" src="' + AUDIO + '" aria-label="Sample audio of the partner interview"></audio><button type="button" class="btn" data-seek="' + s.t + '">' + ICON.play + 'Play from ' + mmss(s.t) + '</button></div>' + waveform(s.t) +
      '<p class="muted small">Sample audio: synthetic voice made for this demo. Pauses stand in for unquoted parts of the 1:12 interview.</p></div>' +
      '<div class="in-block"><div class="label">Transcript · word timestamps</div><ol class="transcript">' + lines.map(function (k) {
        return '<li class="' + (k === id ? 'on' : '') + '"><button type="button" data-cite="' + k + '"><b>' + k + '</b>' + mmss(S[k].t) + '</button><p>' + esc(S[k].text) + '</p></li>';
      }).join('') + '</ol></div>' + statusBlock(id) + citedBy(id) + '</section>';
  }
  function pdfViewer(id) {
    var s = S[id];
    return '<section class="viewer viewer-pdf" data-element="viewer-pdf">' + srcHead(id, 'PDF') + chip('proposed', 'PDF layout parsing with passage highlight') +
      '<div class="pdf-page"><div class="pdf-meta"><span>' + esc(fileById(s.file).name) + '</span><span>Page ' + s.page + ' of ' + s.of + '</span></div><h4>' + esc(s.heading) + '</h4><p>' + esc(s.before) + '<mark>' + esc(s.quote) + '</mark>' + esc(s.after) + '</p></div>' +
      statusBlock(id) + citedBy(id) + '</section>';
  }
  function excelViewer(id) {
    var s = S[id];
    var out = '<section class="viewer viewer-excel" data-element="viewer-excel">' + srcHead(id, 'Excel') + chip('today', 'Excel-based agents read the model') +
      '<div class="xl-formula"><span class="xl-ref">' + s.sheet + '!' + s.cell + '</span><span class="xl-fx">fx</span><span class="xl-f">' + esc(s.formula) + '</span></div>';
    if (id === 'M1') {
      var rows = cashRows();
      out += '<div class="xl-scroll" tabindex="0" role="region" aria-label="CF_Net precedent cells"><table class="xl-grid prec"><thead><tr><th class="rn"></th><th>A</th><th>B</th><th>C</th></tr></thead><tbody>' +
        '<tr><th class="rn">4</th><td><b>Flow</b></td><td><b>Date</b></td><td class="r"><b>USD M</b></td></tr>' +
        rows.map(function (r, i) { return '<tr class="is-prec"><th class="rn">' + (5 + i) + '</th><td>' + esc(r[2]) + '</td><td>' + r[0] + '</td><td class="r">' + (r[1] < 0 ? MINUS + Math.abs(r[1]).toFixed(1) : r[1].toFixed(1)) + '</td></tr>'; }).join('') +
        '<tr class="is-cell"><th class="rn">F18</th><td colspan="2">Summary!F18, Net IRR</td><td class="r"><b>' + F.irr.toFixed(1) + '%</b></td></tr></tbody></table></div>' +
        '<p class="muted small">Precedents: CF_Net!B5:C23, 18 LP cash flows plus the September 30 NAV ' + cite('M4') + ' as terminal value.</p>' +
        '<div class="in-block calc"><div class="label">How this number was calculated</div><pre><code>' + esc(
          'flows = read_range("CF_Net!B5:C23")   # 19 rows\n' +
          '# (2021-06-30, -45.1) ... (2026-09-30, 821.1)\n\n' +
          'def xnpv(rate, flows):\n' +
          '    t0 = flows[0][0]\n' +
          '    return sum(v / (1 + rate) ** ((d - t0).days / 365)\n' +
          '               for d, v in flows)\n\n' +
          'lo, hi = -0.9, 1.0          # bisection on XNPV = 0\n' +
          'for _ in range(100):\n' +
          '    mid = (lo + hi) / 2\n' +
          '    lo, hi = (mid, hi) if xnpv(mid, flows) > 0 else (lo, mid)\n\n' +
          'print(f"{mid:.2%}")          # ' + (xirr(rows) * 100).toFixed(2) + '%') + '</code></pre>' +
        '<p class="calc-result">' + ICON.check + 'Recomputed in this page: <b>' + (xirr(rows) * 100).toFixed(2) + '%</b>, reported as ' + tag('M1', F.irr.toFixed(1) + '%') + '. Q2 reported ' + F.irrPrior.toFixed(1) + '% ' + cite('P1') + '; change ' + MINUS + Math.round((F.irrPrior - F.irr) * 100) + ' bps.</p></div>';
    } else {
      out += '<table class="xl-grid single"><tbody><tr class="is-cell"><th class="rn">' + s.cell + '</th><td>' + esc(s.label) + '</td><td class="r"><b>' + esc(s.value) + '</b></td></tr>' + (s.prior ? '<tr><th class="rn"></th><td>Prior quarter</td><td class="r">' + esc(s.prior) + '</td></tr>' : '') + '</tbody></table>' +
        (s.note ? '<p class="small">' + esc(s.note) + '</p>' : '');
    }
    return out + statusBlock(id) + citedBy(id) + '</section>';
  }
  function excerptViewer(id) {
    var s = S[id], web = s.kind === 'web';
    return '<section class="viewer viewer-' + s.kind + '" data-element="viewer-' + s.kind + '">' + srcHead(id, web ? 'Web snapshot' : 'Prior report') + chip('proposed', web ? 'Dated web snapshots in the vault' : 'Prior-period reports in the vault') +
      '<div class="pdf-page"><div class="pdf-meta"><span>' + (web ? esc(s.url) : 'Page ' + s.page) + '</span><span>' + (web ? 'Captured ' + esc(s.captured) : esc(s.approved)) + '</span></div><p><mark>' + esc(s.quote) + '</mark></p></div>' +
      statusBlock(id) + citedBy(id) + '</section>';
  }
  /* ---------- drawer ---------- */

  function passage(id) {
    var s = S[id];
    if (s.kind === 'voice') return s.text;
    if (s.kind === 'excel') return s.label + ': ' + s.value;
    return s.quote;
  }
  function claimById(id) { return P.letter.filter(function (c) { return c.id === id; })[0]; }
  function tabsFor(c) {
    var kinds = c.cites.map(function (id) { return S[id].kind; }), t = [];
    if (kinds.indexOf('voice') >= 0) t.push(['voice', 'Voice clip']);
    if (kinds.indexOf('excel') >= 0) t.push(['analysis', 'Analysis']);
    t.push(['text', 'Text build']);
    if (kinds.indexOf('pdf') >= 0) t.push(['pdf', 'PDF']);
    if (kinds.indexOf('web') >= 0 || kinds.indexOf('letter') >= 0) t.push(['source', 'Source']);
    return t;
  }
  var TAB_KIND = { voice: ['voice'], analysis: ['excel'], pdf: ['pdf'], source: ['web', 'letter'] };
  function tabOfKind(k) { return k === 'voice' ? 'voice' : k === 'excel' ? 'analysis' : k === 'pdf' ? 'pdf' : 'source'; }
  function pick(c, tab) {
    if (state.sel && c.cites.indexOf(state.sel) >= 0 && TAB_KIND[tab].indexOf(S[state.sel].kind) >= 0) return state.sel;
    return c.cites.filter(function (id) { return TAB_KIND[tab].indexOf(S[id].kind) >= 0; })[0];
  }
  function sourceView(id) {
    var k = S[id].kind;
    if (k === 'voice') return voiceViewer(id);
    if (k === 'pdf') return pdfViewer(id);
    if (k === 'excel') return excelViewer(id);
    return excerptViewer(id);
  }
  function textBuild(c) {
    var edited = c.id === P.edit.claimId;
    var draft = edited ? plain(c).replace(c.final, c.draft) : plain(c);
    return '<section class="viewer" data-element="text-build">' + chip('today', 'Content-level feedback capture') + chip('proposed', 'Typed edits and approval with version history') +
      '<ol class="build"><li><div class="label">1 · Retrieved passages</div><ul>' + c.cites.map(function (id) { return '<li>' + cite(id) + ' ' + esc(passage(id)) + '</li>'; }).join('') + '</ul></li>' +
      '<li><div class="label">2 · Draft v1 · shared ML service · ' + P.timeline.drafted + '</div><p>' + esc(draft) + '</p></li>' +
      (edited ? '<li><div class="label">3 · Reviewer edit v2 · ' + esc(P.edit.reviewer) + ' · ' + P.edit.when + '</div><p><del>' + esc(c.draft) + '</del> <ins>' + esc(c.final) + '</ins></p><p><span class="st st-ok">' + esc(P.edit.type) + '</span> <span class="st">Tone or style</span></p><p class="small">' + esc(P.edit.reason) + '</p></li>'
        : '<li><div class="label">3 · Review · ' + esc(P.people.ir.name) + '</div><p>Accepted without edit.</p></li>') +
      '<li><div class="label">4 · Approval</div><p>' + (state.approved ? 'Approved v3 by ' + esc(P.people.cfo.name) + ', ' + P.timeline.approved : 'Signed off by ' + esc(P.people.partner.name) + '; awaiting CFO approval.') + '</p></li></ol></section>';
  }
  function drawerHTML() {
    var close = '<button type="button" class="drawer-close" data-close="1" aria-label="Close">×</button>';
    if (state.drawer === 'BOUNDARY') return close + boundaryHTML();
    if (state.drawer.indexOf('R:') === 0) return close + recipeHTML(state.drawer.slice(2));
    if (state.drawer.indexOf('S:') === 0) return close + sourceView(state.drawer.slice(2));
    var c = claimById(state.drawer), tabs = tabsFor(c);
    var tab = tabs.some(function (t) { return t[0] === state.tab; }) ? state.tab : tabs[0][0];
    var body = tab === 'text' ? textBuild(c) : sourceView(pick(c, tab));
    return close + '<header class="drawer-head" data-element="drawer"><div class="label">Sentence ' + c.id + ' · ' + c.cites.length + ' citations</div><p>' + esc(plain(c)) + '</p>' + chip('today', 'Source-linked citation chips') + '</header>' +
      '<div class="drawer-tabs" role="tablist">' + tabs.map(function (t) { return '<button type="button" role="tab" class="dtab' + (t[0] === tab ? ' on' : '') + '" aria-selected="' + (t[0] === tab) + '" data-dtab="' + t[0] + '">' + t[1] + '</button>'; }).join('') + '</div>' + body;
  }

  /* ---------- shell ---------- */

  function mainView() {
    return ({ connections: connectionsView, vault: vaultView, ingest: ingestView, templates: templatesView, graph: builtView, outputs: outputsView, detail: detailView }[state.view] || connectionsView)();
  }

  function render() {
    document.getElementById('sidebar').innerHTML = sidebar();
    document.getElementById('main').innerHTML = mainView();
    var dr = document.getElementById('drawer');
    dr.hidden = !state.drawer;
    dr.innerHTML = state.drawer ? drawerHTML() : '';
    document.getElementById('shell').classList.toggle('has-drawer', !!state.drawer);
    var x = dr.querySelector('.xl-scroll');
    if (x) { var c = x.querySelector('.is-cell'); if (c) x.scrollTop = c.offsetTop - x.clientHeight + c.offsetHeight + 8; }
  }

  function tourNav() {
    var dots = TOUR.map(function (t, i) { return '<button type="button" class="dot" data-goto="' + i + '" aria-label="Step ' + i + ': ' + t.label + '"><span>' + i + '</span><em>' + t.label + '</em></button>'; }).join('');
    return '<div class="nav-steps"><button type="button" class="arrow" data-nav="prev" aria-label="Previous">' + ICON.arrow + '</button><div class="dots">' + dots + '</div><button type="button" class="arrow" data-nav="next" aria-label="Next">' + ICON.arrow + '</button></div><div class="nav-hint">Demo track · ← → keys · or click anything</div>';
  }

  function updateCaption() {
    var t = TOUR[current];
    document.getElementById('story-step').textContent = current + ' of ' + (STEP_COUNT - 1) + ' · ' + t.label;
    document.getElementById('story-title').textContent = t.title;
    document.getElementById('story-sub').textContent = t.sub;
    var dots = document.querySelectorAll('[data-goto]');
    for (var i = 0; i < dots.length; i++) {
      var on = +dots[i].dataset.goto === current;
      dots[i].classList.toggle('on', on);
      if (on) dots[i].setAttribute('aria-current', 'step'); else dots[i].removeAttribute('aria-current');
    }
    var prev = document.querySelector('[data-nav="prev"]'), next = document.querySelector('[data-nav="next"]');
    if (prev) prev.disabled = current === 0;
    if (next) next.disabled = current === STEP_COUNT - 1;
  }

  function syncUrl() {
    try { var u = new URL(window.location.href); u.searchParams.set('step', String(current)); window.history.replaceState(null, '', u.toString()); } catch (e) { /* file:// may refuse */ }
  }

  function settle(animate) {
    var token = ++settleToken;
    document.body.dataset.settled = 'false';
    var shell = document.getElementById('shell');
    shell.classList.remove('is-entering');
    if (animate) { void shell.offsetWidth; shell.classList.add('is-entering'); }
    settlePromise = new Promise(function (resolve) {
      requestAnimationFrame(function () {
        var anims = shell.getAnimations ? shell.getAnimations({ subtree: true }) : [];
        Promise.all(anims.map(function (a) { return a.finished.catch(function () {}); })).then(function () {
          if (token === settleToken) document.body.dataset.settled = 'true';
          resolve();
        });
      });
    });
    return settlePromise;
  }

  var genTimer = null;
  function goTo(n, opts) {
    clearInterval(genTimer);
    current = clampStep(n);
    var t = TOUR[current];
    state.view = t.view;
    state.drawer = t.drawer || null;
    state.tab = t.tab || null;
    state.sel = t.drawer === 'C1' ? 'M1' : null;
    state.generated = current >= 5;
    state.running = -1;
    state.approved = false;
    state.history = false;
    state.recipe = 'perf';
    state.gtab = 'recipe';
    document.body.dataset.step = String(current);
    render();
    updateCaption();
    syncUrl();
    return settle(!(opts && opts.instant));
  }

  function generate() {
    clearInterval(genTimer);
    state.generated = false;
    state.running = 0;
    render();
    genTimer = setInterval(function () {
      state.running += 1;
      if (state.running >= STAGES.length) { clearInterval(genTimer); state.generated = true; state.running = -1; }
      render();
    }, 280);
  }

  function openSentence(id, tab) {
    state.drawer = id;
    state.tab = tab || null;
    render();
  }

  function bind() {
    document.addEventListener('click', function (e) {
      var b = e.target.closest('button');
      var sent = e.target.closest('.sent');
      if (!b && sent) { state.sel = null; openSentence(sent.dataset.claim); return; }
      if (!b) return;
      if (b.dataset.cite) {
        state.sel = b.dataset.cite;
        var inSent = b.closest('.sent');
        if (inSent) openSentence(inSent.dataset.claim, tabOfKind(S[b.dataset.cite].kind));
        else if (b.closest('#drawer') && state.drawer && state.drawer.indexOf('S:') !== 0 && state.drawer !== 'BOUNDARY' && claimById(state.drawer) && claimById(state.drawer).cites.indexOf(b.dataset.cite) >= 0) openSentence(state.drawer, tabOfKind(S[b.dataset.cite].kind));
        else { state.drawer = 'S:' + b.dataset.cite; render(); }
      }
      else if (b.dataset.file) { state.drawer = 'S:' + sourcesOfFile(b.dataset.file)[0]; render(); }
      else if (b.dataset.view) { state.view = b.dataset.view; if (state.view !== 'detail') state.drawer = null; render(); }
      else if (b.dataset.recipe) { state.view = 'graph'; state.gtab = 'recipe'; state.recipe = b.dataset.recipe; state.drawer = null; render(); }
      else if (b.dataset.recipeDrawer) { state.drawer = 'R:' + b.dataset.recipeDrawer; render(); }
      else if (b.dataset.gtab) { state.gtab = b.dataset.gtab; render(); }
      else if (b.dataset.dtab) { state.tab = b.dataset.dtab; render(); }
      else if (b.dataset.drawer) { state.drawer = b.dataset.drawer; render(); }
      else if (b.dataset.close) { state.drawer = null; render(); }
      else if (b.dataset.generate) generate();
      else if (b.dataset.history) { state.history = !state.history; render(); }
      else if (b.dataset.approve) { state.approved = true; render(); }
      else if (b.dataset.seek !== undefined) {
        var a = document.querySelector('#drawer audio');
        if (a) { a.preload = 'auto'; a.currentTime = +b.dataset.seek; a.play().catch(function () {}); }
      }
      else if (b.dataset.goto !== undefined) goTo(+b.dataset.goto);
      else if (b.dataset.nav === 'prev') goTo(current - 1);
      else if (b.dataset.nav === 'next') goTo(current + 1);
    });
    document.addEventListener('keydown', function (e) {
      if (e.altKey || e.ctrlKey || e.metaKey || e.shiftKey) return;
      if (e.target && /^(AUDIO|INPUT|TEXTAREA)$/.test(e.target.tagName)) return;
      if (e.key === 'ArrowRight') { e.preventDefault(); goTo(current + 1); }
      else if (e.key === 'ArrowLeft') { e.preventDefault(); goTo(current - 1); }
      else if (e.key === 'Escape' && state.drawer) { state.drawer = null; render(); }
    });
  }

  document.title = 'FinanceOS · ' + P.firm + ' (sample data)';
  document.getElementById('topbar').innerHTML = '<div class="brand"><i class="brand-mark" aria-hidden="true"></i><span>Brain Co.</span><span class="product">FinanceOS</span></div>' +
    '<div class="crumb">' + esc(P.firm) + ' / <b>' + esc(P.package.name) + '</b></div>' +
    '<div class="top-actions"></div>';
  document.getElementById('legend').innerHTML = '<div class="legend-key" aria-label="Legend">' + ['today', 'proposed', 'assumption'].map(function (k) { return chip(k); }).join('') + '</div>';
  document.getElementById('demo-nav').innerHTML = capture ? '' : tourNav();
  if (capture) document.documentElement.classList.add('capture');
  document.body.classList.toggle('capture', capture);
  bind();
  goTo(current, { instant: true });

  window.demo = {
    goTo: function (n) { return goTo(n); },
    steps: STEP_COUNT,
    get step() { return current; },
    next: function () { return goTo(current + 1); },
    prev: function () { return goTo(current - 1); },
    settled: function () { return settlePromise; },
    xirr: function () { return xirr(cashRows()); }
  };
})();
