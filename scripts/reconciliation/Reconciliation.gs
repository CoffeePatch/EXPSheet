/**
 * Reconciliation.gs
 * -----------------------------------------------------------------------------
 * Manual reconciliation between the List sheet and a pasted bank statement.
 * -----------------------------------------------------------------------------
 */

const RECON_CONFIG = Object.freeze({
  SHEET_LIST: "List",
  SHEET_BANK_RAW: "SheetName",
  SHEET_RECON_LOG: "Reconciliation_Log",
  SHEET_RECON_SUMMARY: "Reconciliation_Summary",
  SHEET_RECON_DETAILS: "Reconciliation_Details",

  // External Bank Workbook (Optional)
  BANK_SPREADSHEET_ID: "Sheet_id",
  BANK_SHEET_NAME: "SheetName",

  // Target Ledger Account
  RECON_TARGET_ACCOUNT: "AccountName",

  // Date Parsing and Formats
  RECON_DATE_ORDER: "DMY", // "DMY" or "MDY"
  RECON_START_DATE: "01/05/2026",
  RECON_END_DATE: null,
  RECON_USE_OVERLAP_START_DATE: true,
  RECON_LOG_DATE_FORMAT: "dd/MM/yyyy",

  // Output Modes: "EXCEPTIONS_ONLY" (default), "ALL_TRANSACTIONS", "TWO_SHEETS"
  OUTPUT_MODE: "EXCEPTIONS_ONLY",

  // Dry-Run Mode: When true, runs calculation and logs result model without modifying any output sheets
  DRY_RUN: false,

  // Controlled Amount Tolerance: allowable difference for 1:1 match on same date & direction
  AMOUNT_TOLERANCE: 1.00,

  // Bank column header candidates (searched case-insensitively with normalized spaces)
  BANK_DATE_HEADERS: ["Date", "Txn Date", "Value Date", "Transaction Date", "Posting Date"],
  BANK_DESC_HEADERS: ["Details", "Description", "Particulars", "Narrative", "Narration", "Remarks"],
  BANK_DEBIT_HEADERS: ["Debit", "Withdrawal", "Out", "Dr", "Debit Amount"],
  BANK_CREDIT_HEADERS: ["Credit", "Deposit", "In", "Cr", "Credit Amount"],

  // Ledger column header candidates
  LEDGER_ID_HEADERS: ["ID", "Txn ID", "Transaction ID"],
  LEDGER_DATE_HEADERS: ["Date", "Txn Date", "Transaction Date"],
  LEDGER_ACCOUNT_HEADERS: ["Account", "Account Name", "Wallet"],
  LEDGER_AMOUNT_HEADERS: ["Amount", "Amt"],
  LEDGER_TITLE_HEADERS: ["Title", "Description", "Particulars"],
  LEDGER_NOTES_HEADERS: ["Notes", "Note", "Memo", "Transaction Notes"],
  LEDGER_TYPE_HEADERS: ["Type", "Direction", "Transaction Type"],
  LEDGER_CATEGORY_HEADERS: ["Category"],

  // Bank exclusion rules (internal sweeps, autosweeps, etc.)
  BANK_EXCLUSION_PATTERNS: ["SWEEP"],
  BANK_EXCLUSION_KEYWORDS: ["SWEEP"]
});

function reconcileBankStatement() {
  const ss = SpreadsheetApp.getActive();
  const listSheetName = reconNormalizeSheetName_(RECON_CONFIG.SHEET_LIST);
  const bankSheetName = reconNormalizeSheetName_(RECON_CONFIG.SHEET_BANK_RAW);
  const logSheetName = reconNormalizeSheetName_(RECON_CONFIG.SHEET_RECON_LOG);
  const summarySheetName = reconNormalizeSheetName_(RECON_CONFIG.SHEET_RECON_SUMMARY) || "Reconciliation_Summary";
  const detailsSheetName = reconNormalizeSheetName_(RECON_CONFIG.SHEET_RECON_DETAILS) || "Reconciliation_Details";
  const outputMode = reconNormalizeOutputMode_(RECON_CONFIG.OUTPUT_MODE);

  if (!listSheetName || !bankSheetName || !logSheetName) {
    throw new Error("Missing required sheet names in RECON_CONFIG.");
  }

  const listSheet = ss.getSheetByName(listSheetName);
  const bankSpreadsheetId = String(RECON_CONFIG.BANK_SPREADSHEET_ID || "").trim();
  const bankWorkbook = bankSpreadsheetId ? SpreadsheetApp.openById(bankSpreadsheetId) : ss;
  const bankSheetEffectiveName = reconNormalizeSheetName_(RECON_CONFIG.BANK_SHEET_NAME) || bankSheetName;
  const bankSheet = bankWorkbook.getSheetByName(bankSheetEffectiveName);

  if (!listSheet || !bankSheet) {
    const bankLocation = bankSpreadsheetId
      ? `spreadsheetId=${bankSpreadsheetId}, sheet="${bankSheetEffectiveName}"`
      : `sheet="${bankSheetEffectiveName}" in active spreadsheet`;
    throw new Error(`Missing sheet(s). Required: "${listSheetName}" and bank ${bankLocation}.`);
  }

  if (!RECON_CONFIG.RECON_TARGET_ACCOUNT) {
    throw new Error("Set RECON_CONFIG.RECON_TARGET_ACCOUNT to a target ledger account before reconciling.");
  }

  const tz = ss.getSpreadsheetTimeZone();
  const listValues = listSheet.getDataRange().getValues();
  const bankValues = bankSheet.getDataRange().getValues();

  // Layer 2: Resolve Schemas
  const bankMeta = reconResolveBankColumns_(bankValues, bankSheetEffectiveName);
  const ledgerMeta = reconResolveLedgerColumns_(listValues, listSheetName);

  // Layer 2: Resolve Date Bounds
  const rangeBounds = reconResolveDateRange_(bankValues, bankMeta, listValues, ledgerMeta, tz, RECON_CONFIG);

  // Layer 3 & 4: Parse & Classify Source Data
  const bankExclusionPatterns = RECON_CONFIG.BANK_EXCLUSION_PATTERNS || RECON_CONFIG.BANK_EXCLUSION_KEYWORDS || ["SWEEP"];
  const bankParsed = reconParseBankData_(
    bankValues,
    bankMeta,
    tz,
    RECON_CONFIG.RECON_DATE_ORDER,
    rangeBounds.rangeStartKey,
    rangeBounds.rangeEndKey,
    bankExclusionPatterns
  );

  const ledgerParsed = reconParseLedgerData_(
    listValues,
    ledgerMeta,
    tz,
    RECON_CONFIG.RECON_DATE_ORDER,
    RECON_CONFIG.RECON_TARGET_ACCOUNT,
    rangeBounds.rangeStartKey,
    rangeBounds.rangeEndKey
  );

  // Layer 5 & 6: Execute Pure Reconciliation Engine
  const engineConfig = {
    targetAccount: RECON_CONFIG.RECON_TARGET_ACCOUNT,
    startKey: rangeBounds.rangeStartKey,
    endKey: rangeBounds.rangeEndKey,
    dateFormat: RECON_CONFIG.RECON_LOG_DATE_FORMAT,
    amountTolerance: typeof RECON_CONFIG.AMOUNT_TOLERANCE === "number" ? RECON_CONFIG.AMOUNT_TOLERANCE : 1.00
  };

  const result = reconExecuteReconciliation_(bankParsed, ledgerParsed, engineConfig);

  // Dry-Run Path (Read-Only validation)
  if (RECON_CONFIG.DRY_RUN) {
    Logger.log(`[DRY RUN] Reconciliation completed without modifying any sheets: ` +
      `${result.stats.matchedCount} matched, ${result.stats.splitMatchedCount} split matched, ` +
      `${result.stats.missingCount} missing, ${result.stats.reviewCount} review, ` +
      `${result.stats.unresolvedCount} unresolved, ${result.stats.excludedCount} excluded, ` +
      `${result.stats.ledgerExtraCount} ledger extras.`
    );
    return result;
  }

  // Layer 7: Output & Presentation
  if (outputMode === "EXCEPTIONS_ONLY") {
    reconWriteExceptionsLog_(ss, result, logSheetName, RECON_CONFIG.RECON_LOG_DATE_FORMAT);
  } else if (outputMode === "ALL_TRANSACTIONS") {
    reconWriteExceptionsLog_(ss, result, logSheetName, RECON_CONFIG.RECON_LOG_DATE_FORMAT);
    reconWriteForensicDetails_(ss, result, detailsSheetName, RECON_CONFIG.RECON_LOG_DATE_FORMAT);
  } else if (outputMode === "TWO_SHEETS") {
    reconWriteDailySummary_(ss, result, summarySheetName, RECON_CONFIG.RECON_LOG_DATE_FORMAT);
    reconWriteForensicDetails_(ss, result, detailsSheetName, RECON_CONFIG.RECON_LOG_DATE_FORMAT);
  }

  Logger.log(
    `Reconciliation completed: ${result.stats.matchedCount} matched, ` +
    `${result.stats.splitMatchedCount} split matched, ${result.stats.missingCount} missing, ` +
    `${result.stats.reviewCount} review, ${result.stats.unresolvedCount} unresolved, ` +
    `${result.stats.excludedCount} excluded, ${result.stats.ledgerExtraCount} ledger extras.`
  );
}

function reconNormalizeOutputMode_(value) {
  const raw = String(value || "EXCEPTIONS_ONLY").trim().toUpperCase();
  if (raw === "EXCEPTIONS_ONLY" || raw === "ALL_TRANSACTIONS" || raw === "TWO_SHEETS") return raw;
  return "EXCEPTIONS_ONLY";
}

function reconNormalizeHeader_(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function reconNormalizeSheetName_(value) {
  return String(value || "").trim();
}

function reconRoundCurrency_(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
}

function reconPad2_(value) {
  return String(value).padStart(2, "0");
}

function reconIsValidDate_(year, month, day) {
  if (!year || !month || !day) return false;
  if (year < 1900 || year > 2100) return false;
  if (month < 1 || month > 12) return false;
  if (day < 1 || day > 31) return false;

  const daysInMonth = [
    31,
    (year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0)) ? 29 : 28,
    31, 30, 31, 30, 31, 31, 30, 31, 30, 31
  ];
  return day <= daysInMonth[month - 1];
}

/** ============================================================================
 *  LAYER 2: SCHEMA & HEADER RESOLUTION
 *  ============================================================================ */

/**
 * Resolves column indexes for the bank statement dynamically from header candidates.
 */
function reconResolveBankColumns_(values, bankSheetName) {
  if (!values || values.length === 0) {
    throw new Error(`${bankSheetName} is empty; paste your bank statement before reconciling.`);
  }

  const dateHeaders = (RECON_CONFIG.BANK_DATE_HEADERS || []).map(reconNormalizeHeader_);
  const descHeaders = (RECON_CONFIG.BANK_DESC_HEADERS || []).map(reconNormalizeHeader_);
  const debitHeaders = (RECON_CONFIG.BANK_DEBIT_HEADERS || []).map(reconNormalizeHeader_);
  const creditHeaders = (RECON_CONFIG.BANK_CREDIT_HEADERS || []).map(reconNormalizeHeader_);

  for (let r = 0; r < Math.min(values.length, 10); r++) {
    const row = values[r] || [];
    const headers = row.map(reconNormalizeHeader_);

    const dateIdx = headers.findIndex((h) => dateHeaders.includes(h));
    if (dateIdx < 0) continue;

    const debitIdx = headers.findIndex((h) => debitHeaders.includes(h));
    const creditIdx = headers.findIndex((h) => creditHeaders.includes(h));
    const descIdx = headers.findIndex((h) => descHeaders.includes(h));

    if (debitIdx < 0 || creditIdx < 0) {
      throw new Error(
        `${bankSheetName} header row found at row ${r + 1}, but missing Debit or Credit column. Found: [${headers.join(", ")}].`
      );
    }

    return { startRow: r + 1, dateIdx, descIdx, debitIdx, creditIdx };
  }

  throw new Error(`${bankSheetName} headers not found within first 10 rows. Expected Date, Debit, and Credit columns.`);
}

/**
 * Resolves column indexes for the List ledger dynamically from header candidates.
 */
function reconResolveLedgerColumns_(values, listSheetName) {
  if (!values || values.length === 0) {
    throw new Error(`${listSheetName} is empty; header row is required.`);
  }

  const headerRow = values[0] || [];
  const headers = headerRow.map(reconNormalizeHeader_);

  const idHeaders = (RECON_CONFIG.LEDGER_ID_HEADERS || []).map(reconNormalizeHeader_);
  const dateHeaders = (RECON_CONFIG.LEDGER_DATE_HEADERS || []).map(reconNormalizeHeader_);
  const accountHeaders = (RECON_CONFIG.LEDGER_ACCOUNT_HEADERS || []).map(reconNormalizeHeader_);
  const amountHeaders = (RECON_CONFIG.LEDGER_AMOUNT_HEADERS || []).map(reconNormalizeHeader_);
  const titleHeaders = (RECON_CONFIG.LEDGER_TITLE_HEADERS || []).map(reconNormalizeHeader_);
  const notesHeaders = (RECON_CONFIG.LEDGER_NOTES_HEADERS || []).map(reconNormalizeHeader_);
  const typeHeaders = (RECON_CONFIG.LEDGER_TYPE_HEADERS || []).map(reconNormalizeHeader_);
  const catHeaders = (RECON_CONFIG.LEDGER_CATEGORY_HEADERS || []).map(reconNormalizeHeader_);

  const dateIdx = headers.findIndex((h) => dateHeaders.includes(h));
  const accountIdx = headers.findIndex((h) => accountHeaders.includes(h));
  const amountIdx = headers.findIndex((h) => amountHeaders.includes(h));

  const missing = [];
  if (dateIdx < 0) missing.push("Date");
  if (accountIdx < 0) missing.push("Account");
  if (amountIdx < 0) missing.push("Amount");

  if (missing.length > 0) {
    throw new Error(`${listSheetName} missing required column header(s): ${missing.join(", ")}. Found: [${headers.join(", ")}].`);
  }

  return {
    startRow: 1,
    dateIdx,
    accountIdx,
    amountIdx,
    idIdx: headers.findIndex((h) => idHeaders.includes(h)),
    titleIdx: headers.findIndex((h) => titleHeaders.includes(h)),
    notesIdx: headers.findIndex((h) => notesHeaders.includes(h)),
    typeIdx: headers.findIndex((h) => typeHeaders.includes(h)),
    catIdx: headers.findIndex((h) => catHeaders.includes(h))
  };
}

/**
 * Resolves effective start and end date keys, handling overlap constraints cleanly.
 */
function reconResolveDateRange_(bankValues, bankMeta, listValues, ledgerMeta, tz, config) {
  let firstBankDateKey = "";
  let lastBankDateKey = "";

  for (let r = bankMeta.startRow; r < bankValues.length; r++) {
    const row = bankValues[r] || [];
    const dk = reconNormalizeDateKey_(row[bankMeta.dateIdx], tz, config.RECON_DATE_ORDER);
    if (!dk) continue;
    if (!firstBankDateKey || dk < firstBankDateKey) firstBankDateKey = dk;
    if (!lastBankDateKey || dk > lastBankDateKey) lastBankDateKey = dk;
  }

  if (!firstBankDateKey || !lastBankDateKey) {
    throw new Error("Bank statement contains no valid dates; cannot determine reconciliation range.");
  }

  const cfgStartKey = reconNormalizeDateKey_(config.RECON_START_DATE, tz, config.RECON_DATE_ORDER) || "";
  const cfgEndKey = reconNormalizeDateKey_(config.RECON_END_DATE, tz, config.RECON_DATE_ORDER) || "";

  let rangeStartKey = cfgStartKey || firstBankDateKey;
  let rangeEndKey = cfgEndKey || lastBankDateKey;

  if (rangeStartKey > rangeEndKey) {
    throw new Error(`Configured start date "${rangeStartKey}" is after end date "${rangeEndKey}".`);
  }

  if (config.RECON_USE_OVERLAP_START_DATE) {
    const firstManualKey = reconFindFirstManualDateKey_(
      listValues,
      ledgerMeta,
      tz,
      config.RECON_DATE_ORDER,
      config.RECON_TARGET_ACCOUNT,
      rangeStartKey,
      rangeEndKey
    );
    if (firstManualKey && firstManualKey > rangeStartKey) {
      rangeStartKey = firstManualKey;
      if (rangeStartKey > rangeEndKey) {
        throw new Error(`Overlap start date "${rangeStartKey}" is after end date "${rangeEndKey}".`);
      }
    }
  }

  return { rangeStartKey, rangeEndKey };
}

/**
 * Finds earliest manual date key for target account.
 */
function reconFindFirstManualDateKey_(values, meta, tz, dateOrder, targetAccount, rangeStartKey, rangeEndKey) {
  if (!values || values.length <= meta.startRow) return "";
  let firstKey = "";

  for (let r = meta.startRow; r < values.length; r++) {
    const row = values[r];
    const acc = String(row[meta.accountIdx] || "").trim();
    if (acc !== targetAccount) continue;

    const dk = reconNormalizeDateKey_(row[meta.dateIdx], tz, dateOrder);
    if (!dk) continue;
    if (rangeStartKey && dk < rangeStartKey) continue;
    if (rangeEndKey && dk > rangeEndKey) continue;

    if (!firstKey || dk < firstKey) firstKey = dk;
  }

  return firstKey;
}

/** ============================================================================
 *  LAYER 4: BANK CLASSIFICATION & PARSING
 *  ============================================================================ */

/**
 * Checks whether a bank transaction description matches any exclusion pattern.
 * Standalone keyword/pattern "SWEEP" (case-insensitive, whitespace/newline normalized, substring/token aware).
 *
 * @param {string} desc Bank transaction description
 * @param {Array<string>} patterns Array of exclusion keywords/patterns (e.g. ["SWEEP"])
 * @return {boolean} True if description matches an exclusion pattern
 */
function reconIsExcludedBankDesc_(desc, patterns) {
  if (!desc) return false;
  const rawPatterns = (patterns && patterns.length) ? patterns : ["SWEEP"];
  // Normalize whitespace, tabs, and newlines to single space
  const normalizedDesc = String(desc).replace(/[\r\n\t]+/g, " ").replace(/\s+/g, " ").trim().toUpperCase();
  if (!normalizedDesc) return false;

  for (let i = 0; i < rawPatterns.length; i++) {
    const rawPat = rawPatterns[i];
    const pat = String(rawPat || "").replace(/[\r\n\t]+/g, " ").replace(/\s+/g, " ").trim().toUpperCase();
    if (!pat) continue;

    // Check standalone word / token boundary (e.g. \bSWEEP\b) or substring match
    const escaped = pat.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const tokenRegex = new RegExp(`(^|[^a-zA-Z0-9_])${escaped}([^a-zA-Z0-9_]|$)`, "i");
    if (tokenRegex.test(normalizedDesc) || normalizedDesc.includes(pat)) {
      return true;
    }
  }
  return false;
}

/**
 * Parses and classifies raw bank statement rows into in-scope, excluded, and unresolved.
 */
function reconParseBankData_(values, meta, tz, dateOrder, startKey, endKey, exclusionKeywords) {
  const inScope = [];
  const excluded = [];
  const unresolved = [];
  const patterns = exclusionKeywords || (typeof RECON_CONFIG !== "undefined" ? (RECON_CONFIG.BANK_EXCLUSION_PATTERNS || RECON_CONFIG.BANK_EXCLUSION_KEYWORDS) : ["SWEEP"]);

  for (let r = meta.startRow; r < values.length; r++) {
    const row = values[r] || [];
    const sheetRow = r + 1;

    const rawDate = row[meta.dateIdx];
    const rawDebit = row[meta.debitIdx];
    const rawCredit = row[meta.creditIdx];
    const desc = meta.descIdx >= 0 ? String(row[meta.descIdx] || "").trim() : "";

    // Skip totally empty rows
    if ((rawDate === "" || rawDate === null || rawDate === undefined) &&
        (rawDebit === "" || rawDebit === null || rawDebit === undefined) &&
        (rawCredit === "" || rawCredit === null || rawCredit === undefined) &&
        !desc) {
      continue;
    }

    // Exclusion check (e.g. SWEEP)
    const isExcluded = reconIsExcludedBankDesc_(desc, patterns);
    if (isExcluded) {
      excluded.push({
        bankId: `BANK_${sheetRow}`,
        rowNumber: sheetRow,
        rawDate: rawDate,
        dateKey: reconNormalizeDateKey_(rawDate, tz, dateOrder),
        description: desc,
        rawDebit: rawDebit,
        rawCredit: rawCredit,
        status: "EXCLUDED",
        reason: `Internal sweep exclusion: ${desc}`
      });
      continue;
    }

    // Validate Date
    const dateKey = reconNormalizeDateKey_(rawDate, tz, dateOrder);
    if (!dateKey) {
      unresolved.push({
        source: "Bank",
        rowNumber: sheetRow,
        rawDate: rawDate,
        dateKey: "",
        description: desc,
        status: "UNRESOLVED",
        reason: `Invalid or unparseable bank date: "${rawDate}"`
      });
      continue;
    }

    // Filter by Date Range
    if (startKey && dateKey < startKey) continue;
    if (endKey && dateKey > endKey) continue;

    // Validate Amounts
    const debit = reconParseAmount_(rawDebit);
    const credit = reconParseAmount_(rawCredit);
    const hasDebit = isFinite(debit) && debit !== 0;
    const hasCredit = isFinite(credit) && credit !== 0;

    if (!hasDebit && !hasCredit) {
      unresolved.push({
        source: "Bank",
        rowNumber: sheetRow,
        rawDate: rawDate,
        dateKey: dateKey,
        description: desc,
        status: "UNRESOLVED",
        reason: `Bank row missing both debit and credit amounts`
      });
      continue;
    }

    // Debit is OUT
    if (hasDebit) {
      inScope.push({
        bankId: `BANK_${sheetRow}_OUT`,
        rowNumber: sheetRow,
        dateKey: dateKey,
        rawDate: rawDate,
        direction: "OUT",
        amount: reconRoundCurrency_(Math.abs(debit)),
        rawAmount: rawDebit,
        description: desc
      });
    }

    // Credit is IN
    if (hasCredit) {
      inScope.push({
        bankId: `BANK_${sheetRow}_IN`,
        rowNumber: sheetRow,
        dateKey: dateKey,
        rawDate: rawDate,
        direction: "IN",
        amount: reconRoundCurrency_(Math.abs(credit)),
        rawAmount: rawCredit,
        description: desc
      });
    }
  }

  return { inScope, excluded, unresolved };
}

/**
 * Parses and filters ledger rows for target account and date range.
 */
function reconParseLedgerData_(values, meta, tz, dateOrder, targetAccount, startKey, endKey) {
  const inScope = [];
  const unresolved = [];

  for (let r = meta.startRow; r < values.length; r++) {
    const row = values[r] || [];
    const sheetRow = r + 1;

    const acc = String(row[meta.accountIdx] || "").trim();
    if (acc !== targetAccount) {
      continue; // Filter strictly to target account
    }

    const rawDate = row[meta.dateIdx];
    const rawAmt = row[meta.amountIdx];
    const id = meta.idIdx >= 0 ? String(row[meta.idIdx] || "").trim() : "";
    const title = meta.titleIdx >= 0 ? String(row[meta.titleIdx] || "").trim() : "";
    const notes = meta.notesIdx >= 0 ? String(row[meta.notesIdx] || "").trim() : "";
    const type = meta.typeIdx >= 0 ? String(row[meta.typeIdx] || "").trim().toUpperCase() : "";

    // Validate Date
    const dateKey = reconNormalizeDateKey_(rawDate, tz, dateOrder);
    if (!dateKey) {
      unresolved.push({
        source: "Ledger",
        rowNumber: sheetRow,
        id: id,
        rawDate: rawDate,
        dateKey: "",
        title: title,
        status: "UNRESOLVED",
        reason: `Invalid or unparseable ledger date: "${rawDate}"`
      });
      continue;
    }

    // Date Range Filter
    if (startKey && dateKey < startKey) continue;
    if (endKey && dateKey > endKey) continue;

    // Validate Amount
    const parsedAmt = reconParseAmount_(rawAmt);
    if (!isFinite(parsedAmt) || parsedAmt === 0) {
      unresolved.push({
        source: "Ledger",
        rowNumber: sheetRow,
        id: id,
        rawDate: rawDate,
        dateKey: dateKey,
        title: title,
        status: "UNRESOLVED",
        reason: `Invalid or zero ledger amount: "${rawAmt}"`
      });
      continue;
    }

    // Determine Direction
    let direction;
    if (type === "IN") {
      direction = "IN";
    } else if (type === "OUT") {
      direction = "OUT";
    } else {
      direction = parsedAmt > 0 ? "IN" : "OUT";
    }

    // Extract Split Metadata from Notes or Title
    let splitMeta = reconExtractSplitMetadata_(notes);
    if (!splitMeta.hasMarker) {
      splitMeta = reconExtractSplitMetadata_(title);
    }

    inScope.push({
      ledgerId: `LEDGER_${sheetRow}`,
      rowNumber: sheetRow,
      id: id,
      dateKey: dateKey,
      rawDate: rawDate,
      account: acc,
      title: title,
      notes: notes,
      type: type,
      direction: direction,
      amount: reconRoundCurrency_(Math.abs(parsedAmt)),
      rawAmount: rawAmt,
      splitMeta: splitMeta,
      consumed: false
    });
  }

  return { inScope, unresolved };
}

/** ============================================================================
 *  LAYER 5 & 6: RECONCILIATION ENGINE & RESULT MODEL (PURE LOGIC)
 *  ============================================================================ */

/**
 * Pure reconciliation engine. Reconciles Bank -> Ledger with stateful consumption.
 * @returns {Object} Full result model.
 */
function reconExecuteReconciliation_(bankParsed, ledgerParsed, config) {
  const inScopeBank = bankParsed.inScope || [];
  const inScopeLedger = ledgerParsed.inScope || [];
  const excludedBank = bankParsed.excluded || [];
  const unresolvedAll = (bankParsed.unresolved || []).concat(ledgerParsed.unresolved || []);

  const tolerance = (config && typeof config.amountTolerance === "number")
    ? config.amountTolerance
    : (typeof RECON_CONFIG !== "undefined" && typeof RECON_CONFIG.AMOUNT_TOLERANCE === "number"
        ? RECON_CONFIG.AMOUNT_TOLERANCE
        : 1.00);

  const consumedRowNumbers = new Set();
  const matches = [];
  const splitMatches = [];
  const reviews = [];
  const missing = [];

  // Match Bank -> Ledger
  for (let i = 0; i < inScopeBank.length; i++) {
    const b = inScopeBank[i];

    // Step 5A: 1:1 Match Search with Controlled Amount Tolerance
    const candidates = inScopeLedger.filter(
      (l) =>
        !consumedRowNumbers.has(l.rowNumber) &&
        l.dateKey === b.dateKey &&
        l.direction === b.direction &&
        reconRoundCurrency_(Math.abs(b.amount - l.amount)) <= reconRoundCurrency_(tolerance)
    );

    if (candidates.length >= 1) {
      // Sort candidates: exact match first (diff = 0), then smallest diff, then lowest row number
      candidates.sort((c1, c2) => {
        const diff1 = reconRoundCurrency_(Math.abs(b.amount - c1.amount));
        const diff2 = reconRoundCurrency_(Math.abs(b.amount - c2.amount));
        if (diff1 !== diff2) return diff1 - diff2;
        return c1.rowNumber - c2.rowNumber;
      });

      const cand = candidates[0];
      cand.consumed = true;
      consumedRowNumbers.add(cand.rowNumber);

      const diff = reconRoundCurrency_(Math.abs(b.amount - cand.amount));
      const reason = diff > 0
        ? `Matched within amount tolerance (difference ${diff.toFixed(2)})`
        : "Exact 1:1 match";

      const m = Object.assign({}, b, {
        status: "MATCH",
        matchedLedgerRows: [cand],
        difference: diff,
        reason: reason
      });
      matches.push(m);
      continue;
    }

    // Step 5B: Explicit Split Group Search (1 Bank -> Many Ledger)
    const dateDirCandidates = inScopeLedger.filter(
      (l) =>
        !consumedRowNumbers.has(l.rowNumber) &&
        l.dateKey === b.dateKey &&
        l.direction === b.direction
    );

    let splitHandled = false;
    const potentialGroups = {};

    for (let c = 0; c < dateDirCandidates.length; c++) {
      const l = dateDirCandidates[c];
      const meta = l.splitMeta;

      // A recognized split marker is strictly required
      if (!meta || !meta.hasMarker) {
        continue;
      }

      // The split marker must reference this bank amount
      if (reconRoundCurrency_(meta.originalAmount) !== b.amount) {
        continue;
      }

      // Upstream split identity: base ID if ID exists (length > 1), else title context
      let upstreamId;
      if (l.id && l.id.length > 1) {
        upstreamId = `id_${l.id.slice(0, -1)}`;
      } else if (l.title) {
        upstreamId = `ctx_${reconNormalizeHeader_(l.title)}`;
      } else {
        upstreamId = `row_${l.rowNumber}`;
      }

      const key = `${b.dateKey}_${b.direction}_${upstreamId}_${meta.originalAmount}_${meta.parts}`;
      if (!potentialGroups[key]) potentialGroups[key] = [];
      potentialGroups[key].push(l);
    }

    for (const gKey in potentialGroups) {
      const gRows = potentialGroups[gKey];
      const markerRow = gRows[0];
      const meta = markerRow.splitMeta;
      const expectedParts = meta.parts;
      const totalSharesSum = reconRoundCurrency_(
        gRows.reduce((acc, r) => acc + r.amount, 0)
      );

      // Strict split identity requirements:
      // 1. Same exact date & direction (already filtered)
      // 2. Recognized split marker in notes/title (already verified)
      // 3. Compatible upstream split identity / base ID (enforced via gKey)
      // 4. Expected number of parts (gRows.length === expectedParts)
      // 5. Exact financial sum === bank amount
      if (totalSharesSum === b.amount && (gRows.length === expectedParts || expectedParts <= 0)) {
        for (let k = 0; k < gRows.length; k++) {
          gRows[k].consumed = true;
          consumedRowNumbers.add(gRows[k].rowNumber);
        }

        const sm = Object.assign({}, b, {
          status: "MATCH — SPLIT",
          matchedLedgerRows: gRows
        });
        splitMatches.push(sm);
        splitHandled = true;
        break;
      } else {
        // Incomplete / broken split group
        const rev = Object.assign({}, b, {
          status: "REVIEW",
          reason: `Incomplete split group: found ${gRows.length} of ${expectedParts} share(s) totaling ${totalSharesSum} for bank amount ${b.amount}`,
          matchedLedgerRows: gRows
        });
        reviews.push(rev);
        splitHandled = true;
        break;
      }
    }

    if (splitHandled) continue;

    // Step 5C: Missing Bank Event
    const miss = Object.assign({}, b, {
      status: "MISSING",
      reason: "No unconsumed ledger row found matching date and amount"
    });
    missing.push(miss);
  }

  // Step 5D: Identify Ledger Extras
  const ledgerExtras = [];
  for (let j = 0; j < inScopeLedger.length; j++) {
    const l = inScopeLedger[j];
    if (!consumedRowNumbers.has(l.rowNumber)) {
      const extra = Object.assign({}, l, {
        status: "LEDGER EXTRA",
        reason: "Unconsumed ledger row on target account"
      });
      ledgerExtras.push(extra);
    }
  }

  // Step 5E: Secondary Sanity Check (Daily Checksum)
  const allDatesSet = {};
  inScopeBank.forEach((b) => { allDatesSet[b.dateKey] = true; });
  inScopeLedger.forEach((l) => { allDatesSet[l.dateKey] = true; });
  const sortedDates = Object.keys(allDatesSet).sort();

  const dailySummary = {};
  for (let d = 0; d < sortedDates.length; d++) {
    const dk = sortedDates[d];

    const bankIn = reconRoundCurrency_(
      inScopeBank
        .filter((b) => b.dateKey === dk && b.direction === "IN")
        .reduce((sum, b) => sum + b.amount, 0)
    );
    const bankOut = reconRoundCurrency_(
      inScopeBank
        .filter((b) => b.dateKey === dk && b.direction === "OUT")
        .reduce((sum, b) => sum + b.amount, 0)
    );
    const ledgerIn = reconRoundCurrency_(
      inScopeLedger
        .filter((l) => l.dateKey === dk && l.direction === "IN")
        .reduce((sum, l) => sum + l.amount, 0)
    );
    const ledgerOut = reconRoundCurrency_(
      inScopeLedger
        .filter((l) => l.dateKey === dk && l.direction === "OUT")
        .reduce((sum, l) => sum + l.amount, 0)
    );

    const inflowDiff = reconRoundCurrency_(ledgerIn - bankIn);
    const outflowDiff = reconRoundCurrency_(ledgerOut - bankOut);
    const overallDiff = reconRoundCurrency_(inflowDiff + outflowDiff);

    // Count actionable exceptions on this date (MISSING, REVIEW, UNRESOLVED)
    const dayExceptions = missing
      .filter((b) => b.dateKey === dk)
      .concat(reviews.filter((b) => b.dateKey === dk))
      .concat(unresolvedAll.filter((u) => u.dateKey === dk));

    // Calculate tolerance variances on this date from accepted 1:1 matches
    let tolInDiff = 0;
    let tolOutDiff = 0;
    const dayMatches = matches.filter((m) => m.dateKey === dk);
    for (let mIdx = 0; mIdx < dayMatches.length; mIdx++) {
      const dm = dayMatches[mIdx];
      const matchedRow = dm.matchedLedgerRows && dm.matchedLedgerRows[0];
      if (matchedRow) {
        const diff = reconRoundCurrency_(matchedRow.amount - dm.amount);
        if (dm.direction === "IN") tolInDiff = reconRoundCurrency_(tolInDiff + diff);
        else if (dm.direction === "OUT") tolOutDiff = reconRoundCurrency_(tolOutDiff + diff);
      }
    }

    const unexplainedInflowDiff = reconRoundCurrency_(inflowDiff - tolInDiff);
    const unexplainedOutflowDiff = reconRoundCurrency_(outflowDiff - tolOutDiff);

    let status, reason;
    if (unexplainedInflowDiff === 0 && unexplainedOutflowDiff === 0) {
      if (dayExceptions.length === 0) {
        status = "OK";
        reason = (tolInDiff !== 0 || tolOutDiff !== 0)
          ? "Daily totals match within tolerance and zero transaction exceptions"
          : "Daily totals match and zero transaction exceptions";
      } else {
        status = "WARNING";
        reason = `Daily totals balance, but ${dayExceptions.length} transaction exception(s) exist on this date`;
      }
    } else {
      status = "DISCREPANCY";
      reason = `Daily variance: Inflow diff ${inflowDiff}, Outflow diff ${outflowDiff}`;
    }

    dailySummary[dk] = {
      dateKey: dk,
      bankIn: bankIn,
      bankOut: bankOut,
      ledgerIn: ledgerIn,
      ledgerOut: ledgerOut,
      inflowDiff: inflowDiff,
      outflowDiff: outflowDiff,
      overallDiff: overallDiff,
      status: status,
      reason: reason,
      exceptionCount: dayExceptions.length
    };
  }

  const stats = {
    totalBankRows: inScopeBank.length + excludedBank.length,
    inScopeBankTxns: inScopeBank.length,
    matchedCount: matches.length,
    splitMatchedCount: splitMatches.length,
    reviewCount: reviews.length,
    missingCount: missing.length,
    unresolvedCount: unresolvedAll.length,
    excludedCount: excludedBank.length,
    ledgerExtraCount: ledgerExtras.length
  };

  return {
    configuration: config,
    stats: stats,
    matches: matches,
    splitMatches: splitMatches,
    reviews: reviews,
    missing: missing,
    excluded: excludedBank,
    unresolved: unresolvedAll,
    ledgerExtras: ledgerExtras,
    dailySummary: dailySummary
  };
}

function reconNormalizeDateKey_(value, tz, dateOrder) {
  if (value === null || value === undefined || value === "") return "";

  if (Object.prototype.toString.call(value) === "[object Date]") {
    if (!isNaN(value.getTime())) {
      if (typeof Utilities !== "undefined" && Utilities.formatDate && tz) {
        return Utilities.formatDate(value, tz, "yyyy-MM-dd");
      }
      return `${value.getFullYear()}-${reconPad2_(value.getMonth() + 1)}-${reconPad2_(value.getDate())}`;
    }
    return "";
  }

  if (typeof value === "number" && isFinite(value)) {
    // Support numeric date inputs (Sheets serial dates: 10000..90000)
    if (value >= 10000 && value <= 90000) {
      const serialEpochMs = Date.UTC(1899, 11, 30) + value * 86400000;
      const serialDate = new Date(serialEpochMs);
      if (!isNaN(serialDate.getTime())) {
        if (typeof Utilities !== "undefined" && Utilities.formatDate && tz) {
          return Utilities.formatDate(serialDate, tz, "yyyy-MM-dd");
        }
        return `${serialDate.getUTCFullYear()}-${reconPad2_(serialDate.getUTCMonth() + 1)}-${reconPad2_(serialDate.getUTCDate())}`;
      }
    }

    // Epoch milliseconds
    if (value >= 1e12 && value <= 1e15) {
      const epochDate = new Date(value);
      if (!isNaN(epochDate.getTime())) {
        if (typeof Utilities !== "undefined" && Utilities.formatDate && tz) {
          return Utilities.formatDate(epochDate, tz, "yyyy-MM-dd");
        }
        return `${epochDate.getUTCFullYear()}-${reconPad2_(epochDate.getUTCMonth() + 1)}-${reconPad2_(epochDate.getUTCDate())}`;
      }
    }
    return "";
  }

  const raw = String(value || "").trim();
  if (!raw) return "";

  const datePart = raw.split(/[T ]/)[0];

  // Pattern 1: ISO YYYY-MM-DD or YYYY/MM/DD
  const isoMatch = datePart.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (isoMatch) {
    const y = Number(isoMatch[1]);
    const m = Number(isoMatch[2]);
    const d = Number(isoMatch[3]);
    if (reconIsValidDate_(y, m, d)) {
      return `${y}-${reconPad2_(m)}-${reconPad2_(d)}`;
    }
    return "";
  }

  // Pattern 2: DD/MM/YYYY or MM/DD/YYYY
  const dmyMatch = datePart.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (dmyMatch) {
    const first = Number(dmyMatch[1]);
    const second = Number(dmyMatch[2]);
    const year = Number(dmyMatch[3]);

    if (!isFinite(first) || !isFinite(second) || !isFinite(year)) return "";

    let day = first;
    let month = second;
    const preference = String(dateOrder || "DMY").toUpperCase();
    const normalizedPreference = preference === "MDY" ? "MDY" : "DMY";

    if (first <= 12 && second > 12) {
      day = second;
      month = first;
    } else if (first > 12 && second <= 12) {
      day = first;
      month = second;
    } else if (normalizedPreference === "MDY") {
      day = second;
      month = first;
    } else {
      day = first;
      month = second;
    }

    if (!reconIsValidDate_(year, month, day)) return "";
    return `${String(year)}-${reconPad2_(month)}-${reconPad2_(day)}`;
  }

  // Strictly avoid ambiguous new Date(raw) fallback. Return empty string if not strictly matched.
  return "";
}

function reconExtractSplitMetadata_(text) {
  if (!text) return { hasMarker: false };
  const str = String(text);

  // Pattern 1: [Split: 255 / 2 = 127.5] or [Split: 255/2]
  let m = str.match(/\[\s*Split:\s*([\d,.]+)\s*\/\s*(\d+)\s*(?:=\s*([\d,.]+))?\s*\]/i);
  if (!m) {
    // Pattern 2: (Split: 255/2) or (Split: 255 / 2 = 127.5)
    m = str.match(/\(\s*Split:\s*([\d,.]+)\s*\/\s*(\d+)\s*(?:=\s*([\d,.]+))?\s*\)/i);
  }

  if (m) {
    const orig = reconParseAmount_(m[1]);
    const parts = parseInt(m[2], 10);
    const perPart = m[3] ? reconParseAmount_(m[3]) : reconRoundCurrency_(orig / parts);
    return {
      hasMarker: true,
      originalAmount: orig,
      parts: parts,
      perPartAmount: perPart
    };
  }

  return { hasMarker: false };
}

function reconDateKeyToDate_(dateKey) {
  const parts = String(dateKey || "").split("-");
  if (parts.length !== 3) return null;

  const year = Number(parts[0]);
  const month = Number(parts[1]);
  const day = Number(parts[2]);

  if (!isFinite(year) || !isFinite(month) || !isFinite(day)) return null;
  if (!reconIsValidDate_(year, month, day)) return null;
  return new Date(year, month - 1, day);
}

function reconIsValidDate_(year, month, day) {
  if (!year || !month || !day) return false;
  const parsed = new Date(year, month - 1, day);
  return (
    parsed.getFullYear() === year &&
    parsed.getMonth() === month - 1 &&
    parsed.getDate() === day
  );
}

function reconPad2_(value) {
  return String(value).padStart(2, "0");
}

function reconParseAmount_(value) {
  if (value === null || value === undefined || value === "") return NaN;
  if (typeof value === "number") return value;

  let text = String(value || "").trim();
  if (!text) return NaN;

  let negative = false;
  if (/^\(.*\)$/.test(text)) {
    negative = true;
    text = text.slice(1, -1);
  }

  text = text.replace(/,/g, "").replace(/[^\d.-]/g, "");
  if (!text) return NaN;

  const parsed = Number(text);
  if (!isFinite(parsed)) return NaN;
  return negative ? -Math.abs(parsed) : parsed;
}

/** ============================================================================
 *  LAYER 7: OUTPUT & PRESENTATION WRITERS
 *  ============================================================================ */

/**
 * Generates an 8-character lowercase hexadecimal unique ID (e.g., "79e0f922").
 * Uses Utilities.getUuid() when available, with a random hex fallback.
 * Prevents collisions against already generated IDs within the run.
 *
 * @param {Set<string>} [existingIds] Set of IDs to avoid duplicates
 * @param {Function} [customGenerator] Optional custom generator function (for testing)
 * @return {string} 8-character hex ID
 */
function reconGenerateHexId_(existingIds, customGenerator) {
  if (typeof customGenerator === "function") {
    const customId = String(customGenerator());
    if (existingIds) existingIds.add(customId);
    return customId;
  }

  let id = "";
  let attempts = 0;
  do {
    if (typeof Utilities !== "undefined" && Utilities.getUuid) {
      id = Utilities.getUuid().replace(/-/g, "").slice(0, 8).toLowerCase();
    } else {
      id = "";
      for (let i = 0; i < 8; i++) {
        id += Math.floor(Math.random() * 16).toString(16);
      }
    }
    attempts++;
  } while (existingIds && existingIds.has(id) && attempts < 100);

  if (existingIds) {
    existingIds.add(id);
  }
  return id;
}

/**
 * Formats a normalized dateKey (YYYY-MM-DD) into display string matching configured date format.
 *
 * @param {string} dateKey Normalized YYYY-MM-DD
 * @param {string} [dateFormat] "dd/MM/yyyy", "MM/dd/yyyy", or "yyyy-MM-dd"
 * @return {string}
 */
function reconFormatDateKeyToDisplay_(dateKey, dateFormat) {
  if (!dateKey) return "";
  const parts = String(dateKey).split("-");
  if (parts.length !== 3) return String(dateKey);
  const y = parts[0], m = parts[1], d = parts[2];
  const fmt = String(dateFormat || "dd/MM/yyyy").toLowerCase();
  if (fmt.includes("yyyy") && fmt.startsWith("dd")) {
    return `${d}/${m}/${y}`;
  }
  if (fmt.includes("yyyy") && fmt.startsWith("mm")) {
    return `${m}/${d}/${y}`;
  }
  if (fmt.startsWith("yyyy")) {
    return `${y}-${m}-${d}`;
  }
  return `${d}/${m}/${y}`;
}

/**
 * Builds Action Required ledger entry rows for MISSING bank transactions.
 * Columns: ID | Date | Account | Title | Amount
 * Amount uses signed numeric values: outflows negative (e.g. -255), inflows positive (e.g. 255).
 * Never uses parentheses or accounting formatting.
 *
 * @param {Array<Object>} missingItems Array of MISSING bank transactions
 * @param {Object} [config] Configuration object (targetAccount, dateFormat, idGenerator)
 * @return {Array<Array>} Array of 5-element rows [ID, Date, Account, Title, Amount] with attached named properties
 */
function reconBuildMissingActionRows_(missingItems, config) {
  if (!Array.isArray(missingItems) || missingItems.length === 0) {
    return [];
  }

  const targetAccount = String(
    (config && (config.targetAccount || config.RECON_TARGET_ACCOUNT)) ||
    (typeof RECON_CONFIG !== "undefined" && (RECON_CONFIG.targetAccount || RECON_CONFIG.RECON_TARGET_ACCOUNT)) ||
    "9682"
  ).trim();

  const dateFormat = (config && (config.dateFormat || config.DATE_FORMAT || config.RECON_LOG_DATE_FORMAT)) ||
    (typeof RECON_CONFIG !== "undefined" && (RECON_CONFIG.dateFormat || RECON_CONFIG.DATE_FORMAT || RECON_CONFIG.RECON_LOG_DATE_FORMAT)) ||
    "dd/MM/yyyy";

  const customGenerator = config && config.idGenerator;
  const generatedIds = new Set();
  const actionRows = [];

  for (let i = 0; i < missingItems.length; i++) {
    const item = missingItems[i];
    if (!item) continue;
    // Strict filter: only MISSING bank transactions
    if (item.status && item.status !== "MISSING") continue;

    const id = reconGenerateHexId_(generatedIds, customGenerator);
    const displayDate = item.dateKey ? reconFormatDateKeyToDisplay_(item.dateKey, dateFormat) : String(item.rawDate || "").trim();
    const title = String(item.description || item.title || "").trim();
    const amt = reconRoundCurrency_(item.amount);
    // Outflows negative (e.g. -255), inflows positive (e.g. 255)
    const signedAmount = amt === 0 ? 0 : (item.direction === "OUT" ? -Math.abs(amt) : Math.abs(amt));

    const row = [id, displayDate, targetAccount, title, signedAmount];
    row.id = id;
    row.date = displayDate;
    row.account = targetAccount;
    row.title = title;
    row.amount = signedAmount;
    actionRows.push(row);
  }

  return actionRows;
}

// Alias for compatibility
var reconBuildCopyReadyLedgerRows_ = reconBuildMissingActionRows_;

/**
 * Builds unified, flat, filterable/sortable exception rows for Reconciliation_Log.
 * Exactly 11 columns in order:
 * ID | Date | Account | Title | Amount | P | D | C | Notes | Source | Category of Problem
 *
 * @param {Object} result Reconciliation result model
 * @param {Object} [config] Configuration object (targetAccount, dateFormat, idGenerator)
 * @return {Array<Array>} Array of 11-element rows with attached named properties
 */
function reconBuildUnifiedExceptionRows_(result, config) {
  if (!result) return [];

  const targetAccount = String(
    (config && (config.targetAccount || config.RECON_TARGET_ACCOUNT)) ||
    (typeof RECON_CONFIG !== "undefined" && (RECON_CONFIG.targetAccount || RECON_CONFIG.RECON_TARGET_ACCOUNT)) ||
    "9682"
  ).trim();

  const dateFormat = (config && (config.dateFormat || config.DATE_FORMAT || config.RECON_LOG_DATE_FORMAT)) ||
    (typeof RECON_CONFIG !== "undefined" && (RECON_CONFIG.dateFormat || RECON_CONFIG.DATE_FORMAT || RECON_CONFIG.RECON_LOG_DATE_FORMAT)) ||
    "dd/MM/yyyy";

  const customGenerator = config && config.idGenerator;
  const generatedIds = new Set();
  const rows = [];

  const addRow = (id, dateStr, account, title, signedAmt, notes, source, category, origItem) => {
    const row = [id, dateStr, account, title, signedAmt, "", "", "", notes, source, category];
    row.id = id;
    row.date = dateStr;
    row.account = account;
    row.title = title;
    row.amount = signedAmt;
    row.p = "";
    row.d = "";
    row.c = "";
    row.notes = notes;
    row.source = source;
    row.categoryOfProblem = category;
    row.category = category;
    row.originalItem = origItem;
    rows.push(row);
  };

  // 1. Bank Missing (Missing in Ledger)
  (result.missing || []).forEach((item) => {
    const id = reconGenerateHexId_(generatedIds, customGenerator);
    const dateStr = item.dateKey ? reconFormatDateKeyToDisplay_(item.dateKey, dateFormat) : String(item.rawDate || "").trim();
    const title = String(item.description || item.title || "").trim();
    const amt = reconRoundCurrency_(item.amount);
    const signedAmt = amt === 0 ? 0 : (item.direction === "OUT" ? -Math.abs(amt) : Math.abs(amt));
    const notes = item.reason || "No matching ledger transaction found";
    addRow(id, dateStr, targetAccount, title, signedAmt, notes, "Bank", "Missing in Ledger", item);
  });

  // 2. Reviews (Split review / investigation)
  (result.reviews || []).forEach((item) => {
    const id = reconGenerateHexId_(generatedIds, customGenerator);
    const dateStr = item.dateKey ? reconFormatDateKeyToDisplay_(item.dateKey, dateFormat) : String(item.rawDate || "").trim();
    const title = String(item.description || item.title || "").trim();
    let signedAmt = "";
    if (typeof item.amount === "number" && !isNaN(item.amount)) {
      const a = reconRoundCurrency_(item.amount);
      signedAmt = a === 0 ? 0 : (item.direction === "OUT" ? -Math.abs(a) : Math.abs(a));
    }
    const notes = item.reason || "Split review required";
    addRow(id, dateStr, targetAccount, title, signedAmt, notes, item.source || "Bank", "Review", item);
  });

  // 3. Unresolved Data Errors
  (result.unresolved || []).forEach((item) => {
    const isLedger = item.source === "Ledger";
    const id = isLedger ? (String(item.id || "").trim()) : reconGenerateHexId_(generatedIds, customGenerator);
    const dateStr = item.dateKey ? reconFormatDateKeyToDisplay_(item.dateKey, dateFormat) : String(item.rawDate || "").trim();
    const title = String(item.title || item.description || "").trim();
    let signedAmt = "";
    if (typeof item.amount === "number" && !isNaN(item.amount)) {
      const a = reconRoundCurrency_(item.amount);
      signedAmt = a === 0 ? 0 : (item.direction === "OUT" ? -Math.abs(a) : Math.abs(a));
    } else if (item.rawDebit) {
      const d = reconParseAmount_(item.rawDebit);
      if (isFinite(d) && d !== 0) signedAmt = -Math.abs(reconRoundCurrency_(d));
    } else if (item.rawCredit) {
      const c = reconParseAmount_(item.rawCredit);
      if (isFinite(c) && c !== 0) signedAmt = Math.abs(reconRoundCurrency_(c));
    } else if (item.rawAmount) {
      const r = reconParseAmount_(item.rawAmount);
      if (isFinite(r)) signedAmt = reconRoundCurrency_(r);
    }
    const notes = item.reason || "Unresolved transaction data";
    addRow(id, dateStr, item.account || targetAccount, title, signedAmt, notes, item.source || (isLedger ? "Ledger" : "Bank"), "Unresolved Data", item);
  });

  // 4. Ledger Extras (Missing in Bank)
  (result.ledgerExtras || []).forEach((item) => {
    // Preserve EXISTING ledger ID (not row number)
    const id = String(item.id || "").trim();
    const dateStr = item.dateKey ? reconFormatDateKeyToDisplay_(item.dateKey, dateFormat) : String(item.rawDate || "").trim();
    const title = String(item.title || item.description || "").trim();
    const amt = reconRoundCurrency_(item.amount);
    const signedAmt = amt === 0 ? 0 : (item.direction === "OUT" ? -Math.abs(amt) : Math.abs(amt));
    const notes = item.notes || item.reason || "Unconsumed ledger row on target account";
    addRow(id, dateStr, item.account || targetAccount, title, signedAmt, notes, "Ledger", "Missing in Bank", item);
  });

  return rows;
}

/**
 * Writes the primary exception-oriented reconciliation log.
 * Replaces multiple tables with ONE unified 11-column filterable table:
 * ID | Date | Account | Title | Amount | P | D | C | Notes | Source | Category of Problem
 */
function reconWriteExceptionsLog_(ss, result, sheetName, dateFormat) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  sheet.clear();

  const stats = result.stats;
  const cfg = result.configuration || {};

  // Header Dashboard Cards (Rows 1-4)
  const kpiData = [
    ["BANK RECONCILIATION REPORT", "", "", "", "", "", "", "", "", "", ""],
    [
      `Target Account: ${cfg.targetAccount || ""}`,
      `Date Range: ${cfg.startKey || "Start"} to ${cfg.endKey || "End"}`,
      `Executed: ${new Date().toISOString().replace("T", " ").slice(0, 19)}`,
      "", "", "", "", "", "", "", ""
    ],
    [
      `Bank In-Scope: ${stats.inScopeBankTxns || 0}`,
      `1:1 Matched: ${stats.matchedCount || 0}`,
      `Split Matched: ${stats.splitMatchedCount || 0}`,
      `MISSING: ${stats.missingCount || 0}`,
      `REVIEW: ${stats.reviewCount || 0}`,
      `UNRESOLVED: ${stats.unresolvedCount || 0}`,
      `EXCLUDED (Audit): ${stats.excludedCount || 0}`,
      `LEDGER EXTRA (Info): ${stats.ledgerExtraCount || 0}`,
      "", "", ""
    ],
    ["", "", "", "", "", "", "", "", "", "", ""]
  ];

  sheet.getRange(1, 1, kpiData.length, kpiData[0].length).setValues(kpiData);
  sheet.getRange("A1").setFontWeight("bold").setFontSize(14);
  sheet.getRange("A2:C2").setFontColor("#555555");
  sheet.getRange("A3:H3").setFontWeight("bold");

  // Flat Unified Table (Starts at Row 5)
  // Exactly 11 columns in this order:
  // ID | Date | Account | Title | Amount | P | D | C | Notes | Source | Category of Problem
  const headers = ["ID", "Date", "Account", "Title", "Amount", "P", "D", "C", "Notes", "Source", "Category of Problem"];
  const unifiedRows = reconBuildUnifiedExceptionRows_(result, Object.assign({}, cfg, { dateFormat: dateFormat }));

  const tableRows = [headers];
  if (unifiedRows.length > 0) {
    unifiedRows.forEach((r) => {
      tableRows.push([r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10]]);
    });
  } else {
    tableRows.push([
      "-",
      "-",
      "-",
      "Zero exceptions found",
      "-",
      "", "", "", "", "", ""
    ]);
  }

  const startRow = 5;
  sheet.getRange(startRow, 1, tableRows.length, headers.length).setValues(tableRows);

  // Format headers
  sheet.getRange(startRow, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#f3f3f3");

  // Format Date (Column 2) and Amount (Column 5) as signed numeric currency, NEVER parentheses: #,##0.00;-#,##0.00;0.00
  if (unifiedRows.length > 0) {
    if (dateFormat) {
      sheet.getRange(startRow + 1, 2, unifiedRows.length, 1).setNumberFormat(dateFormat);
    }
    sheet.getRange(startRow + 1, 5, unifiedRows.length, 1).setNumberFormat("#,##0.00;-#,##0.00;0.00");
  }
}

/**
 * Writes the forensic details table containing all transactions (including MATCH and MATCH — SPLIT).
 */
function reconWriteForensicDetails_(ss, result, sheetName, dateFormat) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  sheet.clear();

  const headers = [
    "Date",
    "Status",
    "Bank Row",
    "Bank Dir",
    "Bank Amount",
    "Bank Description",
    "Matched Ledger Rows",
    "Ledger Amount(s)",
    "Ledger Titles / IDs",
    "Reason / Details"
  ];
  const rows = [headers];

  const allTxns = [].concat(
    result.matches,
    result.splitMatches,
    result.reviews,
    result.missing,
    result.excluded
  );

  allTxns.sort((a, b) => String(a.dateKey || "").localeCompare(String(b.dateKey || "")));

  allTxns.forEach((item) => {
    const dateVal = item.dateKey ? reconDateKeyToDate_(item.dateKey) : (item.rawDate || "");
    const status = item.status || "";
    const bankRow = item.rowNumber || "";
    const bankDir = item.direction || "";
    const bankAmt = item.amount ? (item.direction === "OUT" ? -item.amount : item.amount) : "";
    const bankDesc = item.description || "";

    let lRows = "";
    let lAmts = "";
    let lTitles = "";

    if (item.matchedLedgerRows && item.matchedLedgerRows.length > 0) {
      lRows = item.matchedLedgerRows.map((l) => l.rowNumber).join(", ");
      lAmts = item.matchedLedgerRows.map((l) => (l.direction === "OUT" ? -l.amount : l.amount)).join(", ");
      lTitles = item.matchedLedgerRows.map((l) => l.title || l.id).join(", ");
    }

    rows.push([
      dateVal,
      status,
      bankRow,
      bankDir,
      bankAmt,
      bankDesc,
      lRows,
      lAmts,
      lTitles,
      item.reason || ""
    ]);
  });

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
  if (dateFormat && rows.length > 1) {
    sheet.getRange(2, 1, rows.length - 1, 1).setNumberFormat(dateFormat);
    sheet.getRange(2, 5, rows.length - 1, 1).setNumberFormat("#,##0.00;(#,##0.00);-");
  }
}

/**
 * Writes the secondary daily checksum table.
 */
function reconWriteDailySummary_(ss, result, sheetName, dateFormat) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  sheet.clear();

  const headers = [
    "Date",
    "Checksum Status",
    "Ledger In",
    "Bank In",
    "Inflow Diff",
    "Ledger Out",
    "Bank Out",
    "Outflow Diff",
    "Overall Diff",
    "Notes / Status Reason"
  ];
  const rows = [headers];

  const daily = result.dailySummary || {};
  const dateKeys = Object.keys(daily).sort();

  dateKeys.forEach((dk) => {
    const d = daily[dk];
    const dateVal = reconDateKeyToDate_(dk) || dk;

    rows.push([
      dateVal,
      d.status,
      d.ledgerIn,
      d.bankIn,
      d.inflowDiff,
      -d.ledgerOut,
      -d.bankOut,
      d.outflowDiff,
      d.overallDiff,
      d.reason
    ]);
  });

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
  if (dateFormat && rows.length > 1) {
    sheet.getRange(2, 1, rows.length - 1, 1).setNumberFormat(dateFormat);
    sheet.getRange(2, 3, rows.length - 1, 7).setNumberFormat("#,##0.00;(#,##0.00);-");
  }
}