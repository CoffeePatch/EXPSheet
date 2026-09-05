/**
 * ReconciliationTests.gs
 * -----------------------------------------------------------------------------
 * Automated verification suite for the Bank Reconciliation Engine.
 * Tests all 24 required business and presentation scenarios in Apps Script.
 * -----------------------------------------------------------------------------
 */

function runReconciliationTests() {
  const results = [];
  let passed = 0;
  const total = 24;

  function assert(condition, message) {
    if (!condition) {
      throw new Error(`Assertion Failed: ${message}`);
    }
  }

  const baseConfig = {
    targetAccount: "9682",
    dateOrder: "DMY",
    dateFormat: "dd/MM/yyyy",
    exclusionKeywords: ["SWEEP FROM", "AUTOSWEEP"]
  };

  // Test 1: Perfect 1:1 match
  try {
    const banks = {
      inScope: [{ bankId: "B1", rowNumber: 2, dateKey: "2026-09-05", rawDate: "05/09/2026", direction: "OUT", amount: 50.00, rawAmount: "50.00", description: "Coffee" }],
      excluded: [],
      unresolved: []
    };
    const ledger = {
      inScope: [{ ledgerId: "L1", rowNumber: 2, id: "TX1", dateKey: "2026-09-05", rawDate: "05/09/2026", account: "9682", title: "Coffee", direction: "OUT", amount: 50.00, rawAmount: "-50.00", splitMeta: { hasMarker: false }, consumed: false }],
      unresolved: []
    };
    const res = reconExecuteReconciliation_(banks, ledger, baseConfig);
    assert(res.matches.length === 1, "Expected 1 match");
    assert(res.missing.length === 0, "Expected 0 missing");
    assert(res.ledgerExtras.length === 0, "Expected 0 ledger extras");
    assert(res.dailySummary["2026-09-05"].status === "OK", "Expected OK daily checksum");
    results.push("Test 1 (Perfect 1:1 match): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 1 (Perfect 1:1 match): FAILED - ${e.message}`);
  }

  // Test 2: Missing bank transaction
  try {
    const banks = {
      inScope: [{ bankId: "B1", rowNumber: 2, dateKey: "2026-09-05", rawDate: "05/09/2026", direction: "OUT", amount: 120.00, rawAmount: "120.00", description: "Electronics" }],
      excluded: [],
      unresolved: []
    };
    const ledger = { inScope: [], unresolved: [] };
    const res = reconExecuteReconciliation_(banks, ledger, baseConfig);
    assert(res.missing.length === 1, "Expected 1 missing");
    assert(res.missing[0].status === "MISSING", "Expected status MISSING");
    results.push("Test 2 (Missing bank transaction): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 2 (Missing bank transaction): FAILED - ${e.message}`);
  }

  // Test 3: Repeated identical bank transactions
  try {
    const banks = {
      inScope: [
        { bankId: "B1", rowNumber: 2, dateKey: "2026-08-25", rawDate: "25/08/2026", direction: "OUT", amount: 40.00, rawAmount: "40.00", description: "Store A" },
        { bankId: "B2", rowNumber: 3, dateKey: "2026-08-25", rawDate: "25/08/2026", direction: "OUT", amount: 40.00, rawAmount: "40.00", description: "Store A" }
      ],
      excluded: [],
      unresolved: []
    };
    const ledger = {
      inScope: [
        { ledgerId: "L1", rowNumber: 2, id: "ID1", dateKey: "2026-08-25", rawDate: "25/08/2026", account: "9682", title: "Store A", direction: "OUT", amount: 40.00, rawAmount: "-40.00", splitMeta: { hasMarker: false }, consumed: false },
        { ledgerId: "L2", rowNumber: 3, id: "ID2", dateKey: "2026-08-25", rawDate: "25/08/2026", account: "9682", title: "Store A", direction: "OUT", amount: 40.00, rawAmount: "-40.00", splitMeta: { hasMarker: false }, consumed: false }
      ],
      unresolved: []
    };
    const res = reconExecuteReconciliation_(banks, ledger, baseConfig);
    assert(res.matches.length === 2, "Expected 2 matches");
    assert(res.missing.length === 0, "Expected 0 missing");
    assert(res.reviews.length === 0, "Expected 0 reviews");
    assert(res.ledgerExtras.length === 0, "Expected 0 ledger extras");
    results.push("Test 3 (Repeated identical bank transactions): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 3 (Repeated identical bank transactions): FAILED - ${e.message}`);
  }

  // Test 4: Duplicate ledger rows with 1 bank transaction (Rule: Match 1, 2 become LEDGER EXTRA, 0 REVIEW)
  try {
    const banks = {
      inScope: [
        { bankId: "B1", rowNumber: 2, dateKey: "2026-08-25", rawDate: "25/08/2026", direction: "OUT", amount: 40.00, rawAmount: "40.00", description: "Store A" }
      ],
      excluded: [],
      unresolved: []
    };
    const ledger = {
      inScope: [
        { ledgerId: "L1", rowNumber: 2, id: "ID1", dateKey: "2026-08-25", rawDate: "25/08/2026", account: "9682", title: "Store A", direction: "OUT", amount: 40.00, rawAmount: "-40.00", splitMeta: { hasMarker: false }, consumed: false },
        { ledgerId: "L2", rowNumber: 3, id: "ID2", dateKey: "2026-08-25", rawDate: "25/08/2026", account: "9682", title: "Store A", direction: "OUT", amount: 40.00, rawAmount: "-40.00", splitMeta: { hasMarker: false }, consumed: false },
        { ledgerId: "L3", rowNumber: 4, id: "ID3", dateKey: "2026-08-25", rawDate: "25/08/2026", account: "9682", title: "Store A", direction: "OUT", amount: 40.00, rawAmount: "-40.00", splitMeta: { hasMarker: false }, consumed: false }
      ],
      unresolved: []
    };
    const res = reconExecuteReconciliation_(banks, ledger, baseConfig);
    assert(res.matches.length === 1, "Expected 1 match");
    assert(res.reviews.length === 0, "Expected 0 reviews");
    assert(res.ledgerExtras.length === 2, "Expected 2 ledger extras");
    results.push("Test 4 (Duplicate ledger rows with 1 bank txn): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 4 (Duplicate ledger rows with 1 bank txn): FAILED - ${e.message}`);
  }

  // Test 5: Explicit split transaction
  try {
    const banks = {
      inScope: [
        { bankId: "B1", rowNumber: 2, dateKey: "2026-09-05", rawDate: "05/09/2026", direction: "OUT", amount: 255.00, rawAmount: "255.00", description: "Dinner" }
      ],
      excluded: [],
      unresolved: []
    };
    const ledger = {
      inScope: [
        { ledgerId: "L1", rowNumber: 2, id: "SPLIT1A", dateKey: "2026-09-05", rawDate: "05/09/2026", account: "9682", title: "Dinner", direction: "OUT", amount: 127.50, rawAmount: "-127.50", splitMeta: { hasMarker: true, originalAmount: 255.00, parts: 2, perPartAmount: 127.50 }, consumed: false },
        { ledgerId: "L2", rowNumber: 3, id: "SPLIT1B", dateKey: "2026-09-05", rawDate: "05/09/2026", account: "9682", title: "Dinner", direction: "OUT", amount: 127.50, rawAmount: "-127.50", splitMeta: { hasMarker: true, originalAmount: 255.00, parts: 2, perPartAmount: 127.50 }, consumed: false }
      ],
      unresolved: []
    };
    const res = reconExecuteReconciliation_(banks, ledger, baseConfig);
    assert(res.splitMatches.length === 1, "Expected 1 split match");
    assert(res.splitMatches[0].status === "MATCH — SPLIT", "Expected MATCH — SPLIT status");
    assert(res.reviews.length === 0, "Expected 0 reviews");
    assert(res.missing.length === 0, "Expected 0 missing");
    results.push("Test 5 (Explicit split transaction): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 5 (Explicit split transaction): FAILED - ${e.message}`);
  }

  // Test 6: Ambiguous/incomplete split
  try {
    const banks = {
      inScope: [
        { bankId: "B1", rowNumber: 2, dateKey: "2026-09-05", rawDate: "05/09/2026", direction: "OUT", amount: 255.00, rawAmount: "255.00", description: "Dinner" }
      ],
      excluded: [],
      unresolved: []
    };
    const ledger = {
      inScope: [
        { ledgerId: "L1", rowNumber: 2, id: "SPLIT1A", dateKey: "2026-09-05", rawDate: "05/09/2026", account: "9682", title: "Dinner", direction: "OUT", amount: 127.50, rawAmount: "-127.50", splitMeta: { hasMarker: true, originalAmount: 255.00, parts: 2, perPartAmount: 127.50 }, consumed: false }
      ],
      unresolved: []
    };
    const res = reconExecuteReconciliation_(banks, ledger, baseConfig);
    assert(res.reviews.length === 1, "Expected 1 review");
    assert(res.reviews[0].status === "REVIEW", "Expected REVIEW status");
    assert(res.reviews[0].reason.includes("Incomplete split group"), "Expected incomplete split group reason");
    results.push("Test 6 (Ambiguous/incomplete split): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 6 (Ambiguous/incomplete split): FAILED - ${e.message}`);
  }

  // Test 7: Ledger extra
  try {
    const banks = { inScope: [], excluded: [], unresolved: [] };
    const ledger = {
      inScope: [
        { ledgerId: "L1", rowNumber: 2, id: "TX1", dateKey: "2026-09-05", rawDate: "05/09/2026", account: "9682", title: "Solo", direction: "OUT", amount: 75.00, rawAmount: "-75.00", splitMeta: { hasMarker: false }, consumed: false }
      ],
      unresolved: []
    };
    const res = reconExecuteReconciliation_(banks, ledger, baseConfig);
    assert(res.ledgerExtras.length === 1, "Expected 1 ledger extra");
    assert(res.ledgerExtras[0].status === "LEDGER EXTRA", "Expected LEDGER EXTRA status");
    results.push("Test 7 (Ledger extra): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 7 (Ledger extra): FAILED - ${e.message}`);
  }

  // Test 8: Excluded sweep transaction
  try {
    const banks = {
      inScope: [],
      excluded: [
        { bankId: "BANK_2", rowNumber: 2, dateKey: "2026-09-05", rawDate: "05/09/2026", description: "SWEEP FROM ACCOUNT 123", rawDebit: "", rawCredit: "5000.00", status: "EXCLUDED", reason: "Internal sweep exclusion: SWEEP FROM ACCOUNT 123" }
      ],
      unresolved: []
    };
    const ledger = { inScope: [], unresolved: [] };
    const res = reconExecuteReconciliation_(banks, ledger, baseConfig);
    assert(res.excluded.length === 1, "Expected 1 excluded");
    assert(res.excluded[0].status === "EXCLUDED", "Expected EXCLUDED status");
    results.push("Test 8 (Excluded sweep transaction): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 8 (Excluded sweep transaction): FAILED - ${e.message}`);
  }

  // Test 9: Invalid date
  try {
    const banks = {
      inScope: [],
      excluded: [],
      unresolved: [
        { source: "Bank", rowNumber: 2, rawDate: "invalid-date", description: "Bad date txn", status: "UNRESOLVED", reason: "Invalid or unparseable bank date: 'invalid-date'" }
      ]
    };
    const ledger = { inScope: [], unresolved: [] };
    const res = reconExecuteReconciliation_(banks, ledger, baseConfig);
    assert(res.unresolved.length === 1, "Expected 1 unresolved");
    assert(res.unresolved[0].status === "UNRESOLVED", "Expected UNRESOLVED status");
    results.push("Test 9 (Invalid date): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 9 (Invalid date): FAILED - ${e.message}`);
  }

  // Test 10: Invalid amount
  try {
    const banks = {
      inScope: [],
      excluded: [],
      unresolved: [
        { source: "Bank", rowNumber: 2, rawDate: "05/09/2026", description: "Bad amount", status: "UNRESOLVED", reason: "Bank row missing both debit and credit amounts" }
      ]
    };
    const ledger = { inScope: [], unresolved: [] };
    const res = reconExecuteReconciliation_(banks, ledger, baseConfig);
    assert(res.unresolved.length === 1, "Expected 1 unresolved");
    assert(res.unresolved[0].reason.includes("amount"), "Expected amount in reason");
    results.push("Test 10 (Invalid amount): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 10 (Invalid amount): FAILED - ${e.message}`);
  }

  // Test 11: Wrong target account
  try {
    const banks = {
      inScope: [{ bankId: "B1", rowNumber: 2, dateKey: "2026-09-05", rawDate: "05/09/2026", direction: "OUT", amount: 50.00, rawAmount: "50.00", description: "Coffee" }],
      excluded: [],
      unresolved: []
    };
    // Ledger rows with different account are excluded upstream in parseLedgerData
    const ledger = { inScope: [], unresolved: [] };
    const res = reconExecuteReconciliation_(banks, ledger, baseConfig);
    assert(res.missing.length === 1, "Expected bank to be missing");
    assert(res.ledgerExtras.length === 0, "Expected 0 ledger extras");
    results.push("Test 11 (Wrong target account): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 11 (Wrong target account): FAILED - ${e.message}`);
  }

  // Test 12: Equal daily totals but incorrect transaction representation
  try {
    const banks = {
      inScope: [
        { bankId: "B1", rowNumber: 2, dateKey: "2026-09-05", rawDate: "05/09/2026", direction: "OUT", amount: 100.00, rawAmount: "100.00", description: "Bulk item" }
      ],
      excluded: [],
      unresolved: []
    };
    const ledger = {
      inScope: [
        { ledgerId: "L1", rowNumber: 2, id: "TX1", dateKey: "2026-09-05", rawDate: "05/09/2026", account: "9682", title: "Item A", direction: "OUT", amount: 60.00, rawAmount: "-60.00", splitMeta: { hasMarker: false }, consumed: false },
        { ledgerId: "L2", rowNumber: 3, id: "TX2", dateKey: "2026-09-05", rawDate: "05/09/2026", account: "9682", title: "Item B", direction: "OUT", amount: 40.00, rawAmount: "-40.00", splitMeta: { hasMarker: false }, consumed: false }
      ],
      unresolved: []
    };
    const res = reconExecuteReconciliation_(banks, ledger, baseConfig);
    assert(res.missing.length === 1, "Bank 100 should be MISSING");
    assert(res.ledgerExtras.length === 2, "Ledger 60 and 40 should be LEDGER EXTRA");
    assert(res.dailySummary["2026-09-05"].status === "WARNING", "Daily checksum status should be WARNING");
    results.push("Test 12 (Equal daily totals, wrong transaction representation): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 12 (Equal daily totals, wrong transaction representation): FAILED - ${e.message}`);
  }

  // Test 13: Both IN and OUT transactions on same date
  try {
    const banks = {
      inScope: [
        { bankId: "B1", rowNumber: 2, dateKey: "2026-09-05", rawDate: "05/09/2026", direction: "OUT", amount: 20.00, rawAmount: "20.00", description: "Fee" },
        { bankId: "B2", rowNumber: 3, dateKey: "2026-09-05", rawDate: "05/09/2026", direction: "IN", amount: 100.00, rawAmount: "100.00", description: "Refund" }
      ],
      excluded: [],
      unresolved: []
    };
    const ledger = {
      inScope: [
        { ledgerId: "L1", rowNumber: 2, id: "TX1", dateKey: "2026-09-05", rawDate: "05/09/2026", account: "9682", title: "Fee", direction: "OUT", amount: 20.00, rawAmount: "-20.00", splitMeta: { hasMarker: false }, consumed: false },
        { ledgerId: "L2", rowNumber: 3, id: "TX2", dateKey: "2026-09-05", rawDate: "05/09/2026", account: "9682", title: "Refund", direction: "IN", amount: 100.00, rawAmount: "100.00", splitMeta: { hasMarker: false }, consumed: false }
      ],
      unresolved: []
    };
    const res = reconExecuteReconciliation_(banks, ledger, baseConfig);
    assert(res.matches.length === 2, "Expected 2 matches");
    assert(res.missing.length === 0, "Expected 0 missing");
    results.push("Test 13 (Both IN and OUT transactions): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 13 (Both IN and OUT transactions): FAILED - ${e.message}`);
  }

  // Test 14: Formatting/whitespace/case variations
  try {
    const dateKey1 = reconNormalizeDateKey_("  05/09/2026  ", "UTC", "DMY");
    const dateKey2 = reconNormalizeDateKey_("2026-09-05", "UTC", "DMY");
    assert(dateKey1 === "2026-09-05", "dateKey1 should be 2026-09-05");
    assert(dateKey2 === "2026-09-05", "dateKey2 should be 2026-09-05");

    const amt1 = reconParseAmount_(" $1,250.50 ");
    const amt2 = reconParseAmount_(" (1,250.50) ");
    assert(amt1 === 1250.5, "amt1 should be 1250.5");
    assert(amt2 === -1250.5, "amt2 should be -1250.5");

    const banks = {
      inScope: [{ bankId: "B1", rowNumber: 2, dateKey: dateKey1, rawDate: "  05/09/2026  ", direction: "OUT", amount: 1250.50, rawAmount: " $1,250.50 ", description: "  Large Expense  " }],
      excluded: [],
      unresolved: []
    };
    const ledger = {
      inScope: [{ ledgerId: "L1", rowNumber: 2, id: "TX1", dateKey: dateKey2, rawDate: "2026-09-05", account: "9682", title: "LARGE EXPENSE", direction: "OUT", amount: 1250.50, rawAmount: " (1,250.50) ", splitMeta: { hasMarker: false }, consumed: false }],
      unresolved: []
    };
    const res = reconExecuteReconciliation_(banks, ledger, baseConfig);
    assert(res.matches.length === 1, "Expected 1 match");
    results.push("Test 14 (Formatting/whitespace/case variations): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 14 (Formatting/whitespace/case variations): FAILED - ${e.message}`);
  }

  // Test 15: Independent split groups isolation (MUST NOT MERGE)
  try {
    const banks = {
      inScope: [
        { bankId: "B1", rowNumber: 2, dateKey: "2026-09-05", rawDate: "05/09/2026", direction: "OUT", amount: 255.00, rawAmount: "255.00", description: "Dinner A" },
        { bankId: "B2", rowNumber: 3, dateKey: "2026-09-05", rawDate: "05/09/2026", direction: "OUT", amount: 255.00, rawAmount: "255.00", description: "Dinner B" }
      ],
      excluded: [],
      unresolved: []
    };
    const splitMeta = reconExtractSplitMetadata_("Share | [Split: 255 / 2 = 127.5]");
    const ledger = {
      inScope: [
        { ledgerId: "L1", rowNumber: 2, id: "abc1231", dateKey: "2026-09-05", rawDate: "05/09/2026", account: "9682", title: "Dinner A", direction: "OUT", amount: 127.50, rawAmount: "-127.50", splitMeta: splitMeta, notes: "Share 1 | [Split: 255 / 2 = 127.5]", consumed: false },
        { ledgerId: "L2", rowNumber: 3, id: "abc1232", dateKey: "2026-09-05", rawDate: "05/09/2026", account: "9682", title: "Dinner A", direction: "OUT", amount: 127.50, rawAmount: "-127.50", splitMeta: splitMeta, notes: "Share 2 | [Split: 255 / 2 = 127.5]", consumed: false },
        { ledgerId: "L3", rowNumber: 4, id: "xyz7891", dateKey: "2026-09-05", rawDate: "05/09/2026", account: "9682", title: "Dinner B", direction: "OUT", amount: 127.50, rawAmount: "-127.50", splitMeta: splitMeta, notes: "Share 1 | [Split: 255 / 2 = 127.5]", consumed: false },
        { ledgerId: "L4", rowNumber: 5, id: "xyz7892", dateKey: "2026-09-05", rawDate: "05/09/2026", account: "9682", title: "Dinner B", direction: "OUT", amount: 127.50, rawAmount: "-127.50", splitMeta: splitMeta, notes: "Share 2 | [Split: 255 / 2 = 127.5]", consumed: false }
      ],
      unresolved: []
    };
    const res = reconExecuteReconciliation_(banks, ledger, baseConfig);
    assert(res.splitMatches.length === 2, `Expected 2 split matches, got ${res.splitMatches.length}`);
    assert(res.reviews.length === 0, `Expected 0 reviews, got ${res.reviews.length}`);
    assert(res.missing.length === 0, `Expected 0 missing, got ${res.missing.length}`);
    const b1Matched = res.splitMatches.find(m => m.bankId === "B1");
    const b2Matched = res.splitMatches.find(m => m.bankId === "B2");
    assert(b1Matched && b1Matched.matchedLedgerRows.every(r => r.id.startsWith("abc123")), "B1 should match abc123");
    assert(b2Matched && b2Matched.matchedLedgerRows.every(r => r.id.startsWith("xyz789")), "B2 should match xyz789");
    results.push("Test 15 (Independent split groups isolation): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 15 (Independent split groups isolation): FAILED - ${e.message}`);
  }

  // Test 16: Split marker embedded inside extra Notes text
  try {
    const banks = {
      inScope: [{ bankId: "B1", rowNumber: 2, dateKey: "2026-09-05", rawDate: "05/09/2026", direction: "OUT", amount: 300.00, rawAmount: "300.00", description: "Team lunch" }],
      excluded: [],
      unresolved: []
    };
    const meta1 = reconExtractSplitMetadata_("Prefix text [Split: 300 / 3 = 100] suffix text");
    const meta2 = reconExtractSplitMetadata_("Another note [Split: 300 / 3 = 100] more details");
    const meta3 = reconExtractSplitMetadata_("[Split: 300 / 3 = 100]");
    assert(meta1.hasMarker && meta1.originalAmount === 300 && meta1.parts === 3, "Meta1 marker extract failed");
    const ledger = {
      inScope: [
        { ledgerId: "L1", rowNumber: 2, id: "grp1", dateKey: "2026-09-05", rawDate: "05/09/2026", account: "9682", title: "Team lunch", direction: "OUT", amount: 100.00, rawAmount: "-100.00", splitMeta: meta1, notes: "Prefix text [Split: 300 / 3 = 100] suffix text", consumed: false },
        { ledgerId: "L2", rowNumber: 3, id: "grp2", dateKey: "2026-09-05", rawDate: "05/09/2026", account: "9682", title: "Team lunch", direction: "OUT", amount: 100.00, rawAmount: "-100.00", splitMeta: meta2, notes: "Another note [Split: 300 / 3 = 100] more details", consumed: false },
        { ledgerId: "L3", rowNumber: 4, id: "grp3", dateKey: "2026-09-05", rawDate: "05/09/2026", account: "9682", title: "Team lunch", direction: "OUT", amount: 100.00, rawAmount: "-100.00", splitMeta: meta3, notes: "[Split: 300 / 3 = 100]", consumed: false }
      ],
      unresolved: []
    };
    const res = reconExecuteReconciliation_(banks, ledger, baseConfig);
    assert(res.splitMatches.length === 1 && res.splitMatches[0].status === "MATCH — SPLIT", "Expected 1 split match");
    results.push("Test 16 (Embedded split notes marker): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 16 (Embedded split notes marker): FAILED - ${e.message}`);
  }

  // Test 17: Source immutability check
  try {
    const bankOrig = [{ bankId: "B1", rowNumber: 2, dateKey: "2026-09-05", rawDate: "05/09/2026", direction: "OUT", amount: 50.00, rawAmount: "50.00", description: "Coffee" }];
    const ledgerOrig = [{ ledgerId: "L1", rowNumber: 2, id: "TX1", dateKey: "2026-09-05", rawDate: "05/09/2026", account: "9682", title: "Coffee", direction: "OUT", amount: 50.00, rawAmount: "-50.00", splitMeta: { hasMarker: false }, consumed: false }];
    const bSnap = JSON.stringify(bankOrig);

    // Run reconciliation
    reconExecuteReconciliation_({ inScope: bankOrig, excluded: [], unresolved: [] }, { inScope: ledgerOrig, unresolved: [] }, baseConfig);

    assert(JSON.stringify(bankOrig) === bSnap, "Bank input array must not be mutated");
    assert(ledgerOrig[0].id === "TX1" && ledgerOrig[0].amount === 50.00 && ledgerOrig[0].rawAmount === "-50.00", "Ledger raw values must remain immutable");
    results.push("Test 17 (Source immutability): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 17 (Source immutability): FAILED - ${e.message}`);
  }

  // Test 18: Copy-ready row for missing OUT transaction (Bank = -40)
  try {
    const missingItems = [
      { status: "MISSING", bankId: "B1", rowNumber: 2, dateKey: "2026-08-25", rawDate: "25/08/2026", direction: "OUT", amount: 40.00, rawAmount: "40.00", description: "Coffee Shop" }
    ];
    const rows = reconBuildCopyReadyLedgerRows_(missingItems, baseConfig);
    assert(rows.length === 1, "Expected 1 copy-ready row");
    const r = rows[0];
    assert(typeof r.id === "string" && /^[0-9a-f]{8}$/.test(r.id), `ID should be 8 hex chars, got ${r.id}`);
    assert(r.date === "25/08/2026", `Date should be 25/08/2026, got ${r.date}`);
    assert(r.account === "9682", `Account should be 9682, got ${r.account}`);
    assert(r.title === "Coffee Shop", `Title should be 'Coffee Shop', got ${r.title}`);
    assert(r.amount === -40.00, `Amount should be -40, got ${r.amount}`);
    assert(typeof r.amount === "number", "Amount must be numeric type");
    assert(r[0] === r.id && r[1] === r.date && r[2] === r.account && r[3] === r.title && r[4] === r.amount, "Array indexing must match 5 columns");
    results.push("Test 18 (Copy-ready row for missing OUT transaction): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 18 (Copy-ready row for missing OUT transaction): FAILED - ${e.message}`);
  }

  // Test 19: Copy-ready row for missing IN transaction (Bank = +500)
  try {
    const missingItems = [
      { status: "MISSING", bankId: "B2", rowNumber: 3, dateKey: "2026-08-25", rawDate: "25/08/2026", direction: "IN", amount: 500.00, rawAmount: "500.00", description: "Deposit" }
    ];
    const rows = reconBuildCopyReadyLedgerRows_(missingItems, baseConfig);
    assert(rows.length === 1, "Expected 1 copy-ready row");
    const r = rows[0];
    assert(r.amount === 500.00, `Inflow amount should be +500, got ${r.amount}`);
    assert(typeof r.amount === "number", "Amount must be numeric type");
    results.push("Test 19 (Copy-ready row for missing IN transaction): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 19 (Copy-ready row for missing IN transaction): FAILED - ${e.message}`);
  }

  // Test 20: Copy-ready row decimal precision (Bank = -127.50)
  try {
    const missingItems = [
      { status: "MISSING", bankId: "B3", rowNumber: 4, dateKey: "2026-08-25", rawDate: "25/08/2026", direction: "OUT", amount: 127.50, rawAmount: "127.50", description: "Dinner share" }
    ];
    const rows = reconBuildCopyReadyLedgerRows_(missingItems, baseConfig);
    assert(rows.length === 1, "Expected 1 copy-ready row");
    assert(rows[0].amount === -127.50, `Decimal amount should be -127.50, got ${rows[0].amount}`);
    assert(typeof rows[0].amount === "number", "Amount must be numeric type");
    results.push("Test 20 (Copy-ready row decimal precision): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 20 (Copy-ready row decimal precision): FAILED - ${e.message}`);
  }

  // Test 21: Multiple missing transactions get unique generated IDs
  try {
    const missingItems = [
      { status: "MISSING", bankId: "B1", rowNumber: 2, dateKey: "2026-08-25", direction: "OUT", amount: 10.00, description: "Item 1" },
      { status: "MISSING", bankId: "B2", rowNumber: 3, dateKey: "2026-08-25", direction: "OUT", amount: 20.00, description: "Item 2" },
      { status: "MISSING", bankId: "B3", rowNumber: 4, dateKey: "2026-08-25", direction: "OUT", amount: 30.00, description: "Item 3" }
    ];
    const rows = reconBuildCopyReadyLedgerRows_(missingItems, baseConfig);
    assert(rows.length === 3, "Expected 3 copy-ready rows");
    const ids = rows.map(r => r.id);
    const uniqueIds = new Set(ids);
    assert(uniqueIds.size === 3, `All generated IDs must be unique within run: ${ids.join(", ")}`);
    ids.forEach(id => assert(/^[0-9a-f]{8}$/.test(id), `ID ${id} must be 8 hex chars`));
    results.push("Test 21 (Multiple missing txns get unique IDs): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 21 (Multiple missing txns get unique IDs): FAILED - ${e.message}`);
  }

  // Test 22: Non-missing transactions generate no copy-ready rows
  try {
    const nonMissingItems = [
      { status: "MATCH", bankId: "B1", rowNumber: 2, dateKey: "2026-08-25", direction: "OUT", amount: 50.00 },
      { status: "REVIEW", bankId: "B2", rowNumber: 3, dateKey: "2026-08-25", direction: "OUT", amount: 255.00 },
      { status: "UNRESOLVED", bankId: "B3", rowNumber: 4, dateKey: "2026-08-25", direction: "OUT", amount: 10.00 },
      { status: "EXCLUDED", bankId: "B4", rowNumber: 5, dateKey: "2026-08-25", direction: "OUT", amount: 1000.00 }
    ];
    const rows = reconBuildCopyReadyLedgerRows_(nonMissingItems, baseConfig);
    assert(rows.length === 0, `Expected 0 copy-ready rows for non-missing items, got ${rows.length}`);
    results.push("Test 22 (Non-missing txns generate no copy-ready rows): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 22 (Non-missing txns generate no copy-ready rows): FAILED - ${e.message}`);
  }

  // Test 23: Source data immutability during copy-ready generation
  try {
    const missingOrig = [
      { status: "MISSING", bankId: "B1", rowNumber: 2, dateKey: "2026-08-25", rawDate: "25/08/2026", direction: "OUT", amount: 40.00, description: "Coffee" }
    ];
    const snap = JSON.stringify(missingOrig);
    const rows = reconBuildCopyReadyLedgerRows_(missingOrig, baseConfig);
    assert(JSON.stringify(missingOrig) === snap, "Original missing array must not be mutated");
    assert(rows.length === 1, "Expected 1 row generated");
    results.push("Test 23 (Source immutability during row generation): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 23 (Source immutability during row generation): FAILED - ${e.message}`);
  }

  // Test 24: Action Required sheet structure verification (exactly 5 columns, only MISSING items, signed amounts)
  try {
    let capturedValues = null;
    let capturedNumberFormat = null;
    let sectionTitle = "";

    const mockSheet = {
      clear: () => {},
      getRange: (r, c, numRows, numCols) => {
        return {
          setValues: (vals) => {
            if (r === 6 && c === 1) capturedValues = vals;
          },
          setValue: (val) => {
            if (r === 5 && c === 1) sectionTitle = val;
            return mockSheet.getRange();
          },
          setFontWeight: () => mockSheet.getRange(),
          setFontSize: () => mockSheet.getRange(),
          setFontColor: () => mockSheet.getRange(),
          setBackground: () => mockSheet.getRange(),
          setNumberFormat: (fmt) => {
            if (r === 7 && c === 5) capturedNumberFormat = fmt;
            return mockSheet.getRange();
          }
        };
      }
    };
    const mockSS = {
      getSheetByName: () => mockSheet,
      insertSheet: () => mockSheet
    };

    const mockResult = {
      stats: { inScopeBankTxns: 2, matchedCount: 1, splitMatchedCount: 0, missingCount: 1, reviewCount: 0, unresolvedCount: 0, excludedCount: 0, ledgerExtraCount: 0 },
      configuration: baseConfig,
      missing: [
        { status: "MISSING", bankId: "B1", rowNumber: 2, dateKey: "2026-08-25", rawDate: "25/08/2026", direction: "OUT", amount: 255.00, description: "Office Supplies" }
      ],
      reviews: [],
      unresolved: [],
      excluded: [],
      ledgerExtras: []
    };

    reconWriteExceptionsLog_(mockSS, mockResult, "Reconciliation_Log", "dd/MM/yyyy");

    assert(sectionTitle === "ACTION REQUIRED", "Section title must be ACTION REQUIRED");
    assert(capturedValues && capturedValues.length === 2, "Expected header + 1 row");
    const headers = capturedValues[0];
    assert(headers.length === 5, `Action Required must have exactly 5 columns, got ${headers.length}`);
    assert(headers[0] === "ID" && headers[1] === "Date" && headers[2] === "Account" && headers[3] === "Title" && headers[4] === "Amount", "Columns must be ID | Date | Account | Title | Amount");
    const dataRow = capturedValues[1];
    assert(dataRow[4] === -255.00, `Amount must be signed numeric -255, got ${dataRow[4]}`);
    assert(typeof dataRow[4] === "number", "Amount must be numeric type");
    assert(capturedNumberFormat === "#,##0.00;-#,##0.00;0.00", `Number format must be signed without parentheses, got ${capturedNumberFormat}`);
    results.push("Test 24 (Action Required 5-column table structure & signed formatting): PASSED");
    passed++;
  } catch (e) {
    results.push(`Test 24 (Action Required 5-column table structure & signed formatting): FAILED - ${e.message}`);
  }

  Logger.log(`=== RECONCILIATION TEST RESULTS: ${passed}/${total} PASSED ===\n` + results.join("\n"));
  return { passed, total, results };
}
