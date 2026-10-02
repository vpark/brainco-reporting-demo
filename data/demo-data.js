/*
 * Sample data for the Brain Co. LP reporting demo. Every name, number, passage,
 * transcript line and cash flow shown on screen comes from this file.
 * All of it is fictional. Money is USD millions; IRRs are percentages.
 * Fund II net IRR (11.8%) is the XIRR of cashFlows plus the September 30 NAV;
 * the tests recompute it, so the figure, the code snippet and the grid cannot drift.
 */
window.BRAINCO_DEMO_DATA = {
  firm: 'Meridian Capital Partners',
  shortFirm: 'Meridian',
  brand: { mark: 'MERIDIAN', sub: 'CAPITAL PARTNERS', color: '#1E4636', accent: '#A8843F' },
  period: { label: 'Q3 2026', prior: 'Q2 2026', asOf: 'September 30, 2026', asOfShort: 'September 30', priorAsOf: 'June 30', quarterWord: 'Third Quarter 2026' },
  package: { name: 'Q3 2026 LP package', figures: 186, outputs: { deckSlides: 24, letterPages: 6, workbookTabs: 4 } },
  people: {
    partner: { name: 'Dana Whitfield', role: 'Partner', initials: 'DW' },
    ir: { name: 'Sam Ortiz', role: 'IR Associate', initials: 'SO' },
    cfo: { name: 'Priya Raman', role: 'CFO', initials: 'PR' }
  },
  fund: { name: 'Meridian Capital Partners Fund II, L.P.', short: 'Fund II', vintage: 2021, commitments: 850.0, paidIn: 722.0, distributions: 224.0, nav: 821.1, navPrior: 844.6, irrPrior: 14.2, irr: 11.8, rule: 100 },
  registryId: 'MCP-F2.NET_IRR',
  // LP net cash flows: calls negative, distributions positive. Sheet CF_Net, rows 5 onward.
  cashFlows: [
    ['2021-06-30', -45.1, 'Capital call 1'], ['2021-09-30', -45.1, 'Capital call 2'], ['2021-12-31', -45.1, 'Capital call 3'],
    ['2022-03-31', -45.1, 'Capital call 4'], ['2022-06-30', -45.1, 'Capital call 5'], ['2022-09-30', -45.1, 'Capital call 6'],
    ['2022-12-31', -45.1, 'Capital call 7'], ['2023-03-31', -45.1, 'Capital call 8'], ['2023-06-30', -45.1, 'Capital call 9'],
    ['2023-09-30', -45.1, 'Capital call 10'], ['2023-12-31', -45.1, 'Capital call 11'], ['2024-03-31', -45.1, 'Capital call 12'],
    ['2024-06-30', -45.1, 'Capital call 13'], ['2024-09-30', -45.1, 'Capital call 14'], ['2024-12-31', -45.1, 'Capital call 15'],
    ['2024-12-31', 50.5, 'Distribution 1'], ['2025-03-31', -45.5, 'Capital call 16'], ['2026-06-30', 173.5, 'Distribution 2']
  ],
  // Every citable source. kind: excel | pdf | voice | web | letter. file links it to a vault file.
  sources: {
    M1: { kind: 'excel', file: 'model', sheet: 'Summary', cell: 'F18', label: 'Net IRR, Q3 2026', value: '11.8%', formula: '=XIRR(CF_Net!C5:C23, CF_Net!B5:B23)' },
    M4: { kind: 'excel', file: 'model', sheet: 'Summary', cell: 'F12', label: 'Net asset value, September 30', value: '$821.1M', formula: '=SUM(Valuation!H8:H40)' },
    M2: { kind: 'excel', file: 'model', sheet: 'Valuation', cell: 'H22', label: 'Northwind Cold Chain fair value', value: '$249.4M', formula: '=Northwind_DCF!F44', prior: '$262.0M', note: 'Exit date input moved from Dec 31, 2026 to Jun 30, 2027 (Northwind_DCF!C6).' },
    M3: { kind: 'excel', file: 'model', sheet: 'Valuation', cell: 'H31', label: 'Cobalt Dental Partners fair value', value: '$108.6M', formula: '=H30*(1-Adj!C9)', prior: '$118.0M', note: 'Adj!C9 is an 8.0% markdown entered Sep 30, 2026.' },
    A1: { kind: 'excel', file: 'capital', sheet: 'LP_Capital', cell: 'D14', label: 'Paid-in capital and distributions since inception', value: '$722.0M paid in, $224.0M distributed', formula: '=SUMIFS(Ledger!E:E, Ledger!B:B, "Fund II")' },
    A2: { kind: 'excel', file: 'capital', sheet: 'LP_Capital', cell: 'D16', label: 'Fund II NAV per fund administrator, September 30', value: '$830.5M', formula: '=SUM(Positions!F4:F31)', note: 'Seeded discrepancy for this demo: the administrator had not yet booked the $9.4M Cobalt markdown (Valuation!H31). 830.5 − 9.4 = 821.1.' },
    D1: { kind: 'pdf', file: 'covenant', page: 2, of: 6, heading: 'Section 2. Amendments to the Credit Agreement', before: 'The definition of “Maximum Total Leverage Ratio” in Section 7.11(a) is hereby amended as follows: ', quote: 'the maximum Total Leverage Ratio is reset from 5.25:1.00 to 6.00:1.00 for each Test Period ending on or before June 30, 2027,', after: ' and the Borrower shall pay an amendment fee equal to 0.25% of the outstanding Term Loans. No Default or Event of Default has occurred and is continuing as of the Amendment Effective Date.' },
    D2: { kind: 'pdf', file: 'northwind', page: 6, of: 22, heading: 'Strategic process update', before: 'Second-round bids were received from two strategic buyers. ', quote: 'Given financing conditions, the board now expects a signed sale agreement in Q2 2027 and closing by mid-2027, versus Q4 2026 previously.', after: ' Management’s view of enterprise value is unchanged.' },
    D3: { kind: 'pdf', file: 'cobalt', page: 4, of: 18, heading: 'Q3 operating results', before: 'Same-practice revenue grew 6.1% year over year. ', quote: 'Revenue and EBITDA for the quarter were in line with the 2026 operating plan.', after: ' Two de novo practices opened in September.' },
    W1: { kind: 'web', file: 'web', captured: 'Oct 12, 2026, 8:00 AM', url: 'coldchainweekly.example/q3-deal-volume', quote: 'Third-quarter cold-storage deal volume fell 31% from a year earlier, and several sale processes have moved into 2027.' },
    P1: { kind: 'letter', file: 'prior', page: 2, quote: 'Fund II net IRR was 14.2% at June 30, 2026, compared with 14.6% at March 31.', approved: 'Approved by Priya Raman, Aug 7, 2026 (v4)' },
    T1: { kind: 'voice', file: 'voice', t: 4, text: 'Most of the move is Northwind Cold Chain. The sale process slipped, so we now expect to exit in mid-2027 instead of Q4 2026.' },
    T2: { kind: 'voice', file: 'voice', t: 19, text: 'Pushing the expected proceeds out that far lowers the current mark. Our view of the exit value hasn’t really changed.' },
    T3: { kind: 'voice', file: 'voice', t: 41, text: 'Second, Cobalt Dental Partners. In August the lender reset the leverage covenant, and we marked the position down 8%.' },
    T4: { kind: 'voice', file: 'voice', t: 58, text: 'That’s a financing reset, not a default. Operations are on plan.' }
  },
  vault: [
    { type: 'Excel', files: [
      { id: 'model', name: 'Fund II Valuation Model Q3 2026.xlsx', synced: 'Synced Oct 16, 6:00 PM', detail: 'v7, locked at data close · 14 sheets', path: 'Drive › Meridian › Fund II › Q3 2026 › Models', extracted: '14 sheets · 12,480 cells · 3,912 formulas', figures: 109 },
      { id: 'capital', name: 'Fund admin capital account statements Q3 2026.xlsx', synced: 'Synced Oct 15, 4:30 PM', detail: '3 sheets', path: 'Drive › Meridian › Fund II › Q3 2026 › Administrator', extracted: '3 sheets · 1,206 cells · 214 formulas', figures: 12 }
    ] },
    { type: 'PDF', files: [
      { id: 'northwind', name: 'Northwind Cold Chain Q3 board deck.pdf', synced: 'Synced Oct 9, 11:20 AM', detail: '22 pages', path: 'SharePoint › Northwind sale data room › Board', extracted: '22 pages · 9 tables · 6,840 words', figures: 6 },
      { id: 'cobalt', name: 'Cobalt Dental Partners Q3 board deck.pdf', synced: 'Synced Oct 10, 3:05 PM', detail: '18 pages', path: 'Drive › Meridian › Portfolio › Cobalt Dental › Board', extracted: '18 pages · 7 tables · 5,210 words', figures: 5 },
      { id: 'covenant', name: 'Cobalt Dental lender covenant amendment.pdf', synced: 'Synced Aug 26, 9:40 AM', detail: '6 pages', path: 'Drive › Meridian › Portfolio › Cobalt Dental › Legal', extracted: '6 pages · 1 table · 2,380 words', figures: 2 }
    ] },
    { type: 'Web snapshot', files: [
      { id: 'web', name: 'Cold Chain Weekly: Q3 cold-storage deal volume', synced: 'Captured Oct 12, 8:00 AM', detail: 'Trade press', path: 'Web › coldchainweekly.example', extracted: '1 page · 412 words · dated snapshot', figures: 0 }
    ] },
    { type: 'Prior reports', files: [
      { id: 'prior', name: 'Q2 2026 LP letter, approved v4', synced: 'Approved Aug 7', detail: '6 pages', path: 'Vault › Q2 2026 › Outputs', extracted: '6 pages · 2,140 words', figures: 52 }
    ] },
    { type: 'Voice', files: [
      { id: 'voice', name: 'Dana Whitfield, Fund II value change interview', synced: 'Recorded Oct 21, 10:05 AM', detail: '1:12 · transcribed with word timestamps', path: 'Brain Co. voice capture', extracted: '1:12 audio · 186 words with timestamps', figures: 0 }
    ] }
  ],
  // The Fund II section of the LP letter. Each sentence lists its citations;
  // {tag:ID|text} renders a figure as a data tag linked to that source.
  letter: [
    { id: 'C1', text: 'Fund II net IRR declined to {tag:M1|11.8%} from {tag:P1|14.2%} at June 30.', cites: ['M1', 'P1'] },
    { id: 'C2', text: 'The main driver is timing: we now expect to exit Northwind Cold Chain in mid-2027 rather than Q4 2026,', cites: ['T1', 'D2'] },
    { id: 'C3', text: 'which lowers its current mark to {tag:M2|$249.4 million}.', cites: ['T2', 'M2'] },
    { id: 'C4', text: 'We also marked Cobalt Dental Partners down {tag:M3|8%} after its lender reset a leverage covenant.', cites: ['T3', 'D1', 'M3'], draft: 'following a covenant breach', final: 'after its lender reset a leverage covenant' },
    { id: 'C5', text: 'The reset is a financing change, not a default, and the company’s operations are on plan.', cites: ['T4', 'D3'] },
    { id: 'C6', text: 'Cold-storage deal activity slowed across the market during the quarter.', cites: ['W1'] },
    { id: 'C7', text: 'Fund II made no new investments and called no capital during the quarter.', cites: ['A1'] },
    { id: 'C8', text: 'Distributions since inception total {tag:A1|$224.0 million}, and net asset value was {tag:M4|$821.1 million} at September 30, 2026.', cites: ['A1', 'M4'] }
  ],
  edit: { claimId: 'C4', type: 'Factual correction', reviewer: 'Sam Ortiz', when: 'Oct 22, 9:14 AM', reason: '“Covenant breach” is not in the sources. The amendment resets the ratio and states no default has occurred; T4 says it is not a default.' },
  timeline: { dataClose: 'Oct 16, 6:00 PM', drafted: 'Oct 21, 10:07 AM', edited: 'Oct 22, 9:14 AM', signedOff: 'Oct 23, 11:40 AM', approved: 'Nov 4, 4:12 PM' },
  sections: [
    { name: 'Fund overview', status: 'Verified', verified: 18, total: 18 },
    { name: 'Fund II performance', status: 'Verified', verified: 42, total: 42 },
    { name: 'Portfolio updates', status: 'Verified', verified: 66, total: 66 },
    { name: 'Valuation changes', status: 'Verified', verified: 38, total: 38 },
    { name: 'Capital activity', status: 'Verified', verified: 29, total: 29 },
    { name: 'Outlook', status: 'Needs review', verified: 20, total: 21 }
  ],
  config: [
    { k: 'Brand template', v: 'Meridian LP 2026: letter, deck, schedules' },
    { k: 'Metric definition', v: 'Net IRR: XIRR of LP net cash flows, NAV as terminal value' },
    { k: 'Model mapping', v: 'Fund II Valuation Model › Summary!F18' },
    { k: 'Explanation rule', v: 'Net IRR ≥ 100 bps or NAV ≥ 5% quarter over quarter' },
    { k: 'Approval chain', v: 'IR Associate → Partner → CFO' }
  ],
  bespoke: ['One-time continuation-vehicle analysis for Fund I', 'Connector to a single LP’s reporting portal']
};
