/*
 * Prospect data for the "Explain the change" demo.
 * This is the one swappable file: every name, number, transcript line and
 * configuration value shown on screen comes from here. All data is fictional.
 *
 * Money is in USD millions. IRRs are percentages. TVPI and DPI are derived in
 * js/demo.js from paidIn, distributions and nav, so they cannot drift.
 * Bridge steps must sum from the focus vehicle's navPrior to its nav.
 */
window.BRAINCO_DEMO_DATA = {
  version: 1,
  defaultProspect: 'meridian',
  prospects: {
    meridian: {
      id: 'meridian',
      firm: 'Meridian Capital Partners',
      shortFirm: 'Meridian',
      strategy: 'Private equity',
      brand: { mark: 'MERIDIAN', sub: 'CAPITAL PARTNERS', color: '#1E4636', accent: '#A8843F', deckFont: 'serif' },
      period: { label: 'Q3 2026', prior: 'Q2 2026', asOf: 'September 30, 2026', asOfShort: 'September 30', priorAsOf: 'June 30', quarterWord: 'Third Quarter 2026' },
      package: {
        name: 'Q3 2026 LP package',
        figures: 186,
        outputs: { deckSlides: 24, letterPages: 6 },
        workbookTabs: ['Fund Performance', 'Cash Flows', 'Holdings', 'NAV Bridge'],
        milestones: [
          { label: 'Data close', date: 'Oct 16' },
          { label: 'Explanations', date: 'Oct 23' },
          { label: 'Review', date: 'Oct 30' },
          { label: 'CFO approval', date: 'Nov 6' },
          { label: 'Distribution', date: 'Nov 13' }
        ]
      },
      people: {
        partner: { name: 'Dana Whitfield', role: 'Partner', initials: 'DW' },
        ir: { name: 'Sam Ortiz', role: 'IR Associate', initials: 'SO' },
        cfo: { name: 'Priya Raman', role: 'CFO', initials: 'PR' }
      },
      vehicles: [
        { id: 'fund1', name: 'Meridian Capital Partners Fund I, L.P.', short: 'Fund I', vintage: 2017, commitments: 500.0, paidIn: 478.0, distributions: 612.0, nav: 286.8, navPrior: 289.4, irrPrior: 16.1, irr: 16.0 },
        { id: 'fund2', name: 'Meridian Capital Partners Fund II, L.P.', short: 'Fund II', vintage: 2021, commitments: 850.0, paidIn: 722.0, distributions: 224.0, nav: 821.1, navPrior: 844.6, irrPrior: 14.2, irr: 11.8 },
        { id: 'coinvest', name: 'Meridian Co-Invest I, L.P.', short: 'Co-Invest I', vintage: 2023, commitments: 120.0, paidIn: 96.0, distributions: 0.0, nav: 118.1, navPrior: 118.3, irrPrior: 12.4, irr: 12.1 }
      ],
      focus: {
        vehicleId: 'fund2',
        metric: 'Net IRR',
        rule: { irrBps: 100, navPct: 5 },
        model: { file: 'Fund II Valuation Model Q3 2026.xlsx', version: 'v7, locked at data close', sheet: 'Summary', cell: 'F18', formula: '=XIRR(CF_Net!C5:C48, CF_Net!B5:B48)' },
        prior: { file: 'Q2 2026 LP package', version: 'approved v4', approvedBy: 'Priya Raman', approvedOn: 'Aug 7' },
        registryId: 'MCP-F2.NET_IRR',
        deckSlide: 7,
        letterPage: 2
      },
      bridge: [
        { label: 'Northwind Cold Chain', note: 'exit timing', value: -12.6 },
        { label: 'Cobalt Dental Partners', note: '8% markdown', value: -9.4 },
        { label: 'Other holdings, net', note: '', value: -1.5 }
      ],
      holdings: [
        { id: 'M2', name: 'Northwind Cold Chain', cell: 'Valuation!H22', prior: 262.0, current: 249.4 },
        { id: 'M3', name: 'Cobalt Dental Partners', cell: 'Valuation!H31', prior: 118.0, current: 108.6 }
      ],
      request: {
        sentOn: 'Oct 17, 8:15 AM',
        guide: 'Quarterly value change',
        prompts: [
          'What drove the change: timing, value, or both?',
          'Which holdings should LPs hear about by name?',
          'Is anything here a credit or operating problem?'
        ]
      },
      recording: {
        duration: '1:12',
        seconds: 72,
        recordedOn: 'Oct 21, 10:05 AM',
        lines: [
          { id: 'T1', t: '0:04', text: 'Most of the move is Northwind Cold Chain. The sale process slipped, so we now expect to exit in mid-2027 instead of Q4 2026.' },
          { id: 'T2', t: '0:19', text: 'Pushing the expected proceeds out that far lowers the current mark. Our view of the exit value hasn’t really changed.' },
          { id: 'T3', t: '0:41', text: 'Second, Cobalt Dental Partners. In August the lender reset the leverage covenant, and we marked the position down 8%.' },
          { id: 'T4', t: '0:58', text: 'That’s a financing reset, not a default. Operations are on plan.' }
        ]
      },
      claims: [
        { id: 'C1', text: 'Fund II net IRR declined to 11.8% from 14.2% at June 30.', cites: ['M1', 'P1'] },
        { id: 'C2', text: 'The main driver is timing: we now expect to exit Northwind Cold Chain in mid-2027 rather than Q4 2026,', cites: ['T1'] },
        { id: 'C3', text: 'which lowers its current mark.', cites: ['T2', 'M2'] },
        { id: 'C4', text: 'We also marked Cobalt Dental Partners down 8% {edit}.', cites: ['T3', 'M3'],
          edit: { before: 'following a covenant breach', after: 'after its lender reset a leverage covenant' } }
      ],
      edit: {
        claimId: 'C4',
        type: 'factual',
        label: 'Commentary faithfulness',
        reason: '“Covenant breach” is not in the source. T3 says the lender reset the covenant, and T4 says it is not a default.'
      },
      timeline: {
        dataClose: 'Oct 16, 6:00 PM',
        flagged: 'Oct 17, 8:02 AM',
        drafted: 'Oct 21, 10:07 AM',
        edited: 'Oct 22, 9:14 AM',
        signedOff: 'Oct 23, 11:40 AM',
        approved: 'Nov 4, 4:12 PM'
      },
      config: [
        { k: 'Brand template', v: 'Meridian LP 2026: deck, schedules, letter' },
        { k: 'Metric definition', v: 'Net IRR: XIRR of LP net cash flows, NAV as terminal value' },
        { k: 'Model mapping', v: 'Fund II Valuation Model › Summary!F18' },
        { k: 'Explanation rule', v: 'Net IRR ≥ 100 bps or NAV ≥ 5% quarter over quarter' },
        { k: 'Interview guide', v: 'Quarterly value change, 3 prompts' },
        { k: 'Approval chain', v: 'IR Associate → Partner → CFO' },
        { k: 'Audience variants', v: 'LP package; internal portfolio review' }
      ],
      bespoke: [
        'One-time continuation-vehicle analysis for Fund I',
        'Connector to a single LP’s reporting portal'
      ]
    },

    harborline: {
      id: 'harborline',
      firm: 'Harborline Real Estate Partners',
      shortFirm: 'Harborline',
      strategy: 'Institutional real estate',
      brand: { mark: 'HARBORLINE', sub: 'REAL ESTATE PARTNERS', color: '#173A5E', accent: '#3F8A8C', deckFont: 'sans' },
      period: { label: 'Q3 2026', prior: 'Q2 2026', asOf: 'September 30, 2026', asOfShort: 'September 30', priorAsOf: 'June 30', quarterWord: 'Third Quarter 2026' },
      package: {
        name: 'Q3 2026 LP package',
        figures: 214,
        outputs: { deckSlides: 31, letterPages: 8 },
        workbookTabs: ['Fund Performance', 'Cash Flows', 'Properties', 'Leasing', 'NAV Bridge'],
        milestones: [
          { label: 'Data close', date: 'Oct 20' },
          { label: 'Explanations', date: 'Oct 27' },
          { label: 'Review', date: 'Nov 3' },
          { label: 'CFO approval', date: 'Nov 12' },
          { label: 'Distribution', date: 'Nov 17' }
        ]
      },
      people: {
        partner: { name: 'Marcus Hale', role: 'Partner', initials: 'MH' },
        ir: { name: 'Nina Okafor', role: 'IR Associate', initials: 'NO' },
        cfo: { name: 'David Chen', role: 'CFO', initials: 'DC' }
      },
      vehicles: [
        { id: 'income2', name: 'Harborline Income Fund II, L.P.', short: 'Income Fund II', vintage: 2019, commitments: 420.0, paidIn: 405.0, distributions: 168.0, nav: 331.2, navPrior: 332.9, irrPrior: 7.4, irr: 7.2 },
        { id: 'va3', name: 'Harborline Value-Add Fund III, L.P.', short: 'Value-Add Fund III', vintage: 2022, commitments: 640.0, paidIn: 512.0, distributions: 46.0, nav: 584.8, navPrior: 598.3, irrPrior: 12.6, irr: 10.3 },
        { id: 'pacifica', name: 'Harborline–Pacifica Logistics JV', short: 'Pacifica JV', vintage: 2024, commitments: 150.0, paidIn: 88.0, distributions: 0.0, nav: 91.5, navPrior: 91.8, irrPrior: 3.9, irr: 3.6 }
      ],
      focus: {
        vehicleId: 'va3',
        metric: 'Net IRR',
        rule: { irrBps: 100, navPct: 5 },
        model: { file: 'Fund III Valuation Model Q3 2026.xlsx', version: 'v5, locked at data close', sheet: 'Summary', cell: 'E12', formula: '=XIRR(LP_Net!D6:D41, LP_Net!C6:C41)' },
        prior: { file: 'Q2 2026 LP package', version: 'approved v3', approvedBy: 'David Chen', approvedOn: 'Aug 12' },
        registryId: 'HRE-F3.NET_IRR',
        deckSlide: 9,
        letterPage: 3
      },
      bridge: [
        { label: 'Lakeshore Commons', note: 'stabilization timing', value: -8.1 },
        { label: 'Bayview Logistics Center', note: '6% markdown', value: -5.8 },
        { label: 'Other holdings, net', note: '', value: 0.4 }
      ],
      holdings: [
        { id: 'M2', name: 'Lakeshore Commons', cell: 'Valuation!H14', prior: 142.6, current: 134.5 },
        { id: 'M3', name: 'Bayview Logistics Center', cell: 'Valuation!H19', prior: 96.5, current: 90.7 }
      ],
      request: {
        sentOn: 'Oct 21, 9:30 AM',
        guide: 'Quarterly value change',
        prompts: [
          'What drove the change: timing, value, or both?',
          'Which properties should LPs hear about by name?',
          'Is anything here a tenant or leasing problem?'
        ]
      },
      recording: {
        duration: '1:05',
        seconds: 65,
        recordedOn: 'Oct 24, 2:40 PM',
        lines: [
          { id: 'T1', t: '0:03', text: 'Most of the change is Lakeshore Commons in Denver. Lease-up ran slower than we underwrote, so we now expect stabilization in Q3 2027 instead of Q1 2027.' },
          { id: 'T2', t: '0:21', text: 'That pushes the expected sale out about two quarters and lowers the current value. The rents we’re signing are in line with plan.' },
          { id: 'T3', t: '0:39', text: 'Second is Bayview Logistics Center. The new appraisal moved the cap rate out 25 basis points, so we marked it down 6%.' },
          { id: 'T4', t: '0:54', text: 'Bayview is fully leased. This is a market move, not a tenant issue.' }
        ]
      },
      claims: [
        { id: 'C1', text: 'Value-Add Fund III net IRR declined to 10.3% from 12.6% at June 30.', cites: ['M1', 'P1'] },
        { id: 'C2', text: 'Most of the change is timing: we now expect Lakeshore Commons to stabilize in Q3 2027 rather than Q1 2027,', cites: ['T1'] },
        { id: 'C3', text: 'which lowers its current value.', cites: ['T2', 'M2'] },
        { id: 'C4', text: 'We also marked Bayview Logistics Center down 6% {edit}.', cites: ['T3', 'M3'],
          edit: { before: 'after a tenant default', after: 'after its appraisal cap rate widened 25 basis points' } }
      ],
      edit: {
        claimId: 'C4',
        type: 'factual',
        label: 'Commentary faithfulness',
        reason: '“Tenant default” is not in the source. T3 attributes the markdown to the appraisal cap rate, and T4 says it is not a tenant issue.'
      },
      timeline: {
        dataClose: 'Oct 20, 5:30 PM',
        flagged: 'Oct 21, 8:05 AM',
        drafted: 'Oct 24, 2:42 PM',
        edited: 'Oct 27, 9:30 AM',
        signedOff: 'Oct 27, 4:05 PM',
        approved: 'Nov 10, 11:20 AM'
      },
      config: [
        { k: 'Brand template', v: 'Harborline 2026: deck, schedules, letter' },
        { k: 'Metric definition', v: 'Net IRR: XIRR of LP net cash flows, NAV as terminal value' },
        { k: 'Model mapping', v: 'Fund III Valuation Model › Summary!E12' },
        { k: 'Explanation rule', v: 'Net IRR ≥ 100 bps or NAV ≥ 5% quarter over quarter' },
        { k: 'Interview guide', v: 'Quarterly value change, 3 prompts' },
        { k: 'Approval chain', v: 'IR Associate → Partner → CFO' },
        { k: 'Audience variants', v: 'LP package; JV partner summary' }
      ],
      bespoke: [
        'One-time appraisal-firm data feed for Bayview',
        'Pacifica JV partner waterfall reconciliation'
      ]
    }
  }
};

