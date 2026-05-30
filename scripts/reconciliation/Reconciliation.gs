/**
 * Reconciliation.gs
 * -----------------------------------------------------------------------------
 * Manual reconciliation between the List sheet and a pasted bank statement.
 * -----------------------------------------------------------------------------
 */

const RECON_CONFIG = Object.freeze({
  SHEET_LIST: "List", // Change to match your ledger tab name.
  SHEET_BANK_RAW: "Bank_Raw", // Change to match the pasted bank statement tab.
  SHEET_RECON_LOG: "Reconciliation_Log", // Change to customize the output tab name.
  RECON_LOG_DATE_FORMAT: "dd/MM/yyyy", // Output format for the Date column in Reconciliation_Log.

  // Optional: if your bank statement lives in a different Google Sheets file,
  // set BANK_SPREADSHEET_ID and (optionally) BANK_SHEET_NAME.
  // - If BANK_SPREADSHEET_ID is null/blank, the script reads Bank_Raw from the active spreadsheet.
  // - If BANK_SPREADSHEET_ID is set, the script opens that spreadsheet and reads BANK_SHEET_NAME
  //   (or falls back to SHEET_BANK_RAW when BANK_SHEET_NAME is not set).
  BANK_SPREADSHEET_ID: null,
  BANK_SHEET_NAME: null,

  // Optional: when true, ignore any dates before BOTH datasets have started (overlap window).
  // This helps avoid early-bank-date mismatches when your manual ledger starts later.
  RECON_USE_OVERLAP_START_DATE: false,

  // Output control:
  // - "SUMMARY_ONLY": current behavior (1 row per mismatched date).
  // - "SINGLE_SHEET": one flat, sortable table with SUMMARY + TXN rows in SHEET_RECON_LOG.
  // - "TWO_SHEETS": writes a date summary to SHEET_RECON_SUMMARY and transactions to SHEET_RECON_DETAILS.
  OUTPUT_MODE: "SUMMARY_ONLY",
  SHEET_RECON_SUMMARY: "Reconciliation_Summary",
  SHEET_RECON_DETAILS: "Reconciliation_Details",

  RECON_TARGET_ACCOUNT: null, // TODO: set this to the exact List sheet account name in column C (case-sensitive, e.g., "9682").
  RECON_DATE_ORDER: "DMY",
  // Optional: supply a start and/or end date to restrict the reconciliation range.
  // Accepts a Date object or any value parseable by `reconNormalizeDateKey_`.
  // If only start is provided, the bank statement's last row date is used as end.
  // If only end is provided, the bank statement's first row date is used as start.
  // If neither provided, the full bank statement range is used.
  RECON_START_DATE: null,
  RECON_END_DATE: null,

  BANK_DATE_HEADERS: ["Txn Date", "Value Date"],
  BANK_DESC_HEADERS: ["Details", "Description", "Particulars", "Narration"],
  BANK_DEBIT_HEADER: "Debit",
  BANK_CREDIT_HEADER: "Credit",

  // Optional: used only for detailed outputs (SINGLE_SHEET / TWO_SHEETS)
  // If none match, the Title/Desc column will be blank.
  MANUAL_DESC_HEADERS: ["Title", "Notes", "Description", "Category"]
});

function reconcileBankStatement() {
  const ss = SpreadsheetApp.getActive();
  const listSheetName = reconNormalizeSheetName_(RECON_CONFIG.SHEET_LIST);
  const bankSheetName = reconNormalizeSheetName_(RECON_CONFIG.SHEET_BANK_RAW);
  const logSheetName = reconNormalizeSheetName_(RECON_CONFIG.SHEET_RECON_LOG);
  const outputMode = reconNormalizeOutputMode_(RECON_CONFIG.OUTPUT_MODE);
  const summarySheetName = reconNormalizeSheetName_(RECON_CONFIG.SHEET_RECON_SUMMARY) || "Reconciliation_Summary";
  const detailsSheetName = reconNormalizeSheetName_(RECON_CONFIG.SHEET_RECON_DETAILS) || "Reconciliation_Details";

  if (!listSheetName || !bankSheetName || !logSheetName) {
    const missingSheets = [];
    if (!listSheetName) missingSheets.push("SHEET_LIST");
    if (!bankSheetName) missingSheets.push("SHEET_BANK_RAW");
    if (!logSheetName) missingSheets.push("SHEET_RECON_LOG");
    throw new Error(
      `Set RECON_CONFIG.${missingSheets.join(
        ", "
      )} to your sheet names before reconciling.`
    );
  }

  const listSheet = ss.getSheetByName(listSheetName);

  const bankSpreadsheetId = String(RECON_CONFIG.BANK_SPREADSHEET_ID || "").trim();
  const bankWorkbook = bankSpreadsheetId ? SpreadsheetApp.openById(bankSpreadsheetId) : ss;
  const bankSheetEffectiveName =
    reconNormalizeSheetName_(RECON_CONFIG.BANK_SHEET_NAME) || bankSheetName;
  const bankSheet = bankWorkbook.getSheetByName(bankSheetEffectiveName);

  if (!listSheet || !bankSheet) {
    const bankLocation = bankSpreadsheetId
      ? `spreadsheetId=${bankSpreadsheetId}, sheet="${bankSheetEffectiveName}"`
      : `sheet="${bankSheetEffectiveName}" in active spreadsheet`;
    throw new Error(
      `Missing sheet(s). Required: "${listSheetName}" and bank ${bankLocation}.`
    );
  }

  if (!RECON_CONFIG.RECON_TARGET_ACCOUNT) {
    throw new Error(
      "Set RECON_CONFIG.RECON_TARGET_ACCOUNT to a List sheet account name before reconciling."
    );
  }

  const tz = ss.getSpreadsheetTimeZone();
  const manualValues = listSheet.getDataRange().getValues();
  const bankValues = bankSheet.getDataRange().getValues();

  // Determine totals and bank column metadata
  const bankMeta = reconResolveBankColumns_(bankValues, bankSheetEffectiveName);

  // Determine bank statement first/last date keys for range resolution.
  let firstBankDateKey = "";
  let lastBankDateKey = "";
  for (let r = bankMeta.startRow; r < bankValues.length; r++) {
    const row = bankValues[r] || [];
    const dk = reconNormalizeDateKey_(row[bankMeta.dateIdx], tz, RECON_CONFIG.RECON_DATE_ORDER);
    if (!dk) continue;
    if (!firstBankDateKey) firstBankDateKey = dk;
    lastBankDateKey = dk;
  }

  if (!firstBankDateKey || !lastBankDateKey) {
    throw new Error(`${bankSheetName} contains no valid dates; cannot determine reconciliation range.`);
  }

  // Resolve configured start/end date keys.
  const cfgStartKey = reconNormalizeDateKey_(RECON_CONFIG.RECON_START_DATE, tz, RECON_CONFIG.RECON_DATE_ORDER) || "";
  const cfgEndKey = reconNormalizeDateKey_(RECON_CONFIG.RECON_END_DATE, tz, RECON_CONFIG.RECON_DATE_ORDER) || "";

  let rangeStartKey = cfgStartKey || "";
  let rangeEndKey = cfgEndKey || "";

  if (rangeStartKey && !rangeEndKey) {
    rangeEndKey = lastBankDateKey;
  } else if (!rangeStartKey && rangeEndKey) {
    rangeStartKey = firstBankDateKey;
  } else if (!rangeStartKey && !rangeEndKey) {
    rangeStartKey = firstBankDateKey;
    rangeEndKey = lastBankDateKey;
  }

  if (rangeStartKey > rangeEndKey) {
    throw new Error(`Configured start date "${rangeStartKey}" is after end date "${rangeEndKey}".`);
  }

  if (RECON_CONFIG.RECON_USE_OVERLAP_START_DATE) {
    const firstManualKey = reconFindFirstManualDateKey_(
      manualValues,
      RECON_CONFIG.RECON_TARGET_ACCOUNT,
      tz,
      rangeStartKey,
      rangeEndKey
    );
    if (firstManualKey) {
      rangeStartKey = firstManualKey > rangeStartKey ? firstManualKey : rangeStartKey;
      if (rangeStartKey > rangeEndKey) {
        throw new Error(
          `Overlap start date "${rangeStartKey}" is after end date "${rangeEndKey}".`
        );
      }
    }
  }

  const bankTotals = reconAggregateBankTotals_(bankValues, bankMeta, tz, rangeStartKey, rangeEndKey);
  const manualTotalsFiltered = reconAggregateManualTotals_(manualValues, RECON_CONFIG.RECON_TARGET_ACCOUNT, tz, rangeStartKey, rangeEndKey);

  if (outputMode === "SUMMARY_ONLY") {
    const logRows = reconBuildLogRows_(manualTotalsFiltered, bankTotals);
    reconWriteLog_(ss, logRows, logSheetName);
    return;
  }

  const manualDescIdx = reconResolveManualDescIndex_(manualValues);
  const manualByDate = reconAggregateManualByDate_(manualValues, RECON_CONFIG.RECON_TARGET_ACCOUNT, tz, rangeStartKey, rangeEndKey, manualDescIdx);
  const bankByDate = reconAggregateBankByDate_(bankValues, bankMeta, tz, rangeStartKey, rangeEndKey);

  if (outputMode === "SINGLE_SHEET") {
    const combinedRows = reconBuildCombinedRows_(manualByDate, bankByDate);
    reconWriteTable_(ss, combinedRows, logSheetName, RECON_CONFIG.RECON_LOG_DATE_FORMAT, 1);
    return;
  }

  // TWO_SHEETS
  const summaryRows = reconBuildSummaryRowsFromDaily_(manualByDate, bankByDate);
  const detailsRows = reconBuildDetailsRowsFromDaily_(manualByDate, bankByDate);
  reconWriteTable_(ss, summaryRows, summarySheetName, RECON_CONFIG.RECON_LOG_DATE_FORMAT, 1);
  reconWriteTable_(ss, detailsRows, detailsSheetName, RECON_CONFIG.RECON_LOG_DATE_FORMAT, 1);
}

function reconNormalizeOutputMode_(value) {
  const raw = String(value || "SUMMARY_ONLY").trim().toUpperCase();
  if (raw === "SINGLE_SHEET" || raw === "TWO_SHEETS" || raw === "SUMMARY_ONLY") return raw;
  return "SUMMARY_ONLY";
}

function reconFindFirstManualDateKey_(values, accountFilter, tz, rangeStartKey, rangeEndKey) {
  if (!values || values.length < 2) return "";
  let firstKey = "";

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const account = String(row[2] || "").trim();
    if (account !== accountFilter) continue;

    const dateKey = reconNormalizeDateKey_(row[0], tz, RECON_CONFIG.RECON_DATE_ORDER);
    if (!dateKey) continue;
    if (rangeStartKey && dateKey < rangeStartKey) continue;
    if (rangeEndKey && dateKey > rangeEndKey) continue;

    if (!firstKey || dateKey < firstKey) firstKey = dateKey;
  }

  return firstKey;
}

function reconAggregateManualTotals_(values, accountFilter, tz, rangeStartKey, rangeEndKey) {
  const totals = {};
  if (!values || values.length < 2) return totals;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const account = String(row[2] || "").trim();
    if (account !== accountFilter) continue;
    const dateKey = reconNormalizeDateKey_(row[0], tz, RECON_CONFIG.RECON_DATE_ORDER);
    if (!dateKey) continue;

    if (rangeStartKey && dateKey < rangeStartKey) continue;
    if (rangeEndKey && dateKey > rangeEndKey) continue;

    const amount = reconParseAmount_(row[4]);
    if (!isFinite(amount) || amount === 0) continue;

    const bucket = reconEnsureTotals_(totals, dateKey);
    if (amount > 0) {
      bucket.inflow = reconRoundCurrency_(bucket.inflow + amount);
    } else {
      bucket.outflow = reconRoundCurrency_(bucket.outflow + amount);
    }
  }

  return totals;
}

function reconResolveManualDescIndex_(values) {
  if (!values || values.length === 0) return -1;
  const headerRow = values[0] || [];
  const headers = headerRow.map(reconNormalizeHeader_);
  const candidates = (RECON_CONFIG.MANUAL_DESC_HEADERS || []).map(reconNormalizeHeader_);
  for (let i = 0; i < candidates.length; i++) {
    const idx = headers.indexOf(candidates[i]);
    if (idx >= 0) return idx;
  }
  return -1;
}

function reconAggregateManualByDate_(values, accountFilter, tz, rangeStartKey, rangeEndKey, descIdx) {
  const totals = {};
  if (!values || values.length < 2) return totals;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const account = String(row[2] || "").trim();
    if (account !== accountFilter) continue;

    const dateKey = reconNormalizeDateKey_(row[0], tz, RECON_CONFIG.RECON_DATE_ORDER);
    if (!dateKey) continue;
    if (rangeStartKey && dateKey < rangeStartKey) continue;
    if (rangeEndKey && dateKey > rangeEndKey) continue;

    const amount = reconParseAmount_(row[4]);
    if (!isFinite(amount) || amount === 0) continue;

    const title = descIdx >= 0 ? String(row[descIdx] || "").trim() : "";
    const bucket = reconEnsureTotalsDetailed_(totals, dateKey);
    bucket.manualTxns.push({
      dateKey,
      rowNumber: r + 1,
      title,
      amount
    });

    if (amount > 0) {
      bucket.inflow = reconRoundCurrency_(bucket.inflow + amount);
    } else {
      bucket.outflow = reconRoundCurrency_(bucket.outflow + amount);
    }
  }

  return totals;
}

function reconResolveBankColumns_(values, bankSheetName) {
  if (!values || values.length === 0) {
    throw new Error(
      `${bankSheetName} is empty; paste your bank statement before reconciling.`
    );
  }

  const dateHeaders = (RECON_CONFIG.BANK_DATE_HEADERS || []).map(reconNormalizeHeader_);
  const descHeaders = (RECON_CONFIG.BANK_DESC_HEADERS || []).map(reconNormalizeHeader_);
  const debitHeader = reconNormalizeHeader_(RECON_CONFIG.BANK_DEBIT_HEADER);
  const creditHeader = reconNormalizeHeader_(RECON_CONFIG.BANK_CREDIT_HEADER);

  for (let r = 0; r < values.length; r++) {
    const row = values[r] || [];
    const headers = row.map(reconNormalizeHeader_);
    const dateIdx = headers.findIndex((header) => dateHeaders.includes(header));

    if (dateIdx < 0) continue;

    const debitIdx = headers.indexOf(debitHeader);
    const creditIdx = headers.indexOf(creditHeader);
    const descIdx = descHeaders.length ? headers.findIndex((header) => descHeaders.includes(header)) : -1;

    if (debitIdx < 0 || creditIdx < 0) {
      throw new Error(
        `${bankSheetName} header row found at row ${r + 1} but missing "${RECON_CONFIG.BANK_DEBIT_HEADER}" or "${RECON_CONFIG.BANK_CREDIT_HEADER}".`
      );
    }

    return { startRow: r + 1, dateIdx, descIdx, debitIdx, creditIdx };
  }

  const dateHeadersDisplay = RECON_CONFIG.BANK_DATE_HEADERS.join('" or "');
  throw new Error(
    `${bankSheetName} headers not found. Expected "${dateHeadersDisplay}" plus "${RECON_CONFIG.BANK_DEBIT_HEADER}" and "${RECON_CONFIG.BANK_CREDIT_HEADER}".`
  );
}

function reconAggregateBankTotals_(values, meta, tz, rangeStartKey, rangeEndKey) {
  const totals = {};
  if (!values || values.length === 0) return totals;

  for (let r = meta.startRow; r < values.length; r++) {
    const row = values[r];
    const dateKey = reconNormalizeDateKey_(row[meta.dateIdx], tz, RECON_CONFIG.RECON_DATE_ORDER);
    if (!dateKey) continue;

    if (rangeStartKey && dateKey < rangeStartKey) continue;
    if (rangeEndKey && dateKey > rangeEndKey) continue;

    const debit = reconParseAmount_(row[meta.debitIdx]);
    const credit = reconParseAmount_(row[meta.creditIdx]);

    if (
      (!isFinite(debit) || debit === 0) &&
      (!isFinite(credit) || credit === 0)
    ) {
      continue;
    }

    const bucket = reconEnsureTotals_(totals, dateKey);

    if (isFinite(credit) && credit !== 0) {
      bucket.inflow = reconRoundCurrency_(bucket.inflow + Math.abs(credit));
    }

    if (isFinite(debit) && debit !== 0) {
      // Treat debits as outflows regardless of the sign in the statement export.
      const debitOutflow = -Math.abs(debit);
      bucket.outflow = reconRoundCurrency_(bucket.outflow + debitOutflow);
    }
  }

  return totals;
}

function reconNormalizeHeader_(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function reconBuildLogRows_(manualTotals, bankTotals) {
  const headers = [
    "Date",
    "Manual Inflow",
    "Bank Inflow",
    "Inflow Diff",
    "Manual Outflow",
    "Bank Outflow",
    "Outflow Diff",
    "Overall Diff"
  ];

  const rows = [headers];
  const dateKeys = reconCollectDateKeys_(manualTotals, bankTotals);

  dateKeys.forEach((dateKey) => {
    const dateValue = reconDateKeyToDate_(dateKey);
    if (!dateValue) {
      throw new Error(`Invalid normalized date key "${dateKey}".`);
    }

    const manual = manualTotals[dateKey] || { inflow: 0, outflow: 0 };
    const bank = bankTotals[dateKey] || { inflow: 0, outflow: 0 };

    const inflowDiff = reconRoundCurrency_(manual.inflow - bank.inflow);
    const outflowDiff = reconRoundCurrency_(manual.outflow - bank.outflow);

    const overallDiff = reconRoundCurrency_(inflowDiff + outflowDiff);

    if (inflowDiff !== 0 || outflowDiff !== 0) {
      rows.push([
        dateValue,
        manual.inflow,
        bank.inflow,
        inflowDiff,
        manual.outflow,
        bank.outflow,
        outflowDiff,
        overallDiff
      ]);
    }
  });

  return rows;
}

function reconCollectDateKeys_(manualTotals, bankTotals) {
  const keys = Object.create(null);
  Object.keys(manualTotals || {}).forEach((key) => {
    keys[key] = true;
  });
  Object.keys(bankTotals || {}).forEach((key) => {
    keys[key] = true;
  });
  return Object.keys(keys).sort();
}

function reconWriteLog_(ss, rows, logSheetName) {
  let logSheet = ss.getSheetByName(logSheetName);
  if (!logSheet) {
    logSheet = ss.insertSheet(logSheetName);
  }

  logSheet.clearContents();
  if (rows.length === 0) return;
  logSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  const dateFormat = String(RECON_CONFIG.RECON_LOG_DATE_FORMAT || "").trim();
  if (dateFormat && rows.length > 1) {
    logSheet.getRange(2, 1, rows.length - 1, 1).setNumberFormat(dateFormat);
  }
}

function reconNormalizeDateKey_(value, tz, dateOrder) {
  if (!value) return "";

  if (Object.prototype.toString.call(value) === "[object Date]") {
    if (!isNaN(value.getTime())) {
      return Utilities.formatDate(value, tz, "yyyy-MM-dd");
    }
  }

  const raw = String(value || "").trim();
  if (!raw) return "";

  const datePart = raw.split(/[T ]/)[0];
  if (/^\d{4}-\d{2}-\d{2}$/.test(datePart)) {
    return datePart;
  }

  if (/^\d{1,2}[/-]\d{1,2}[/-]\d{4}$/.test(datePart)) {
    const parts = datePart.split(/[/-]/);
    const first = Number(parts[0]);
    const second = Number(parts[1]);
    const year = Number(parts[2]);

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
    } else if (first <= 12 && second <= 12 && normalizedPreference === "MDY") {
      day = second;
      month = first;
    } else if (first <= 12 && second <= 12) {
      day = first;
      month = second;
    }

    if (!reconIsValidDate_(year, month, day)) return "";
    return `${String(year)}-${reconPad2_(month)}-${reconPad2_(day)}`;
  }

  const parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, tz, "yyyy-MM-dd");
  }

  return "";
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

function reconRoundCurrency_(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
}

function reconNormalizeSheetName_(value) {
  return String(value || "").trim();
}

function reconEnsureTotals_(map, key) {
  if (!map[key]) {
    map[key] = { inflow: 0, outflow: 0 };
  }
  return map[key];
}

function reconAggregateBankByDate_(values, meta, tz, rangeStartKey, rangeEndKey) {
  const totals = {};
  if (!values || values.length === 0) return totals;

  for (let r = meta.startRow; r < values.length; r++) {
    const row = values[r];
    const dateKey = reconNormalizeDateKey_(row[meta.dateIdx], tz, RECON_CONFIG.RECON_DATE_ORDER);
    if (!dateKey) continue;
    if (rangeStartKey && dateKey < rangeStartKey) continue;
    if (rangeEndKey && dateKey > rangeEndKey) continue;

    const debit = reconParseAmount_(row[meta.debitIdx]);
    const credit = reconParseAmount_(row[meta.creditIdx]);

    if ((!isFinite(debit) || debit === 0) && (!isFinite(credit) || credit === 0)) {
      continue;
    }

    const title = meta.descIdx >= 0 ? String(row[meta.descIdx] || "").trim() : "";
    const bucket = reconEnsureTotalsDetailed_(totals, dateKey);

    if (isFinite(credit) && credit !== 0) {
      const amt = Math.abs(credit);
      bucket.inflow = reconRoundCurrency_(bucket.inflow + amt);
      bucket.bankTxns.push({
        dateKey,
        rowNumber: r + 1,
        title,
        amount: amt
      });
    }

    if (isFinite(debit) && debit !== 0) {
      const amt = -Math.abs(debit);
      bucket.outflow = reconRoundCurrency_(bucket.outflow + amt);
      bucket.bankTxns.push({
        dateKey,
        rowNumber: r + 1,
        title,
        amount: amt
      });
    }
  }

  return totals;
}

function reconWriteTable_(ss, rows, sheetName, dateFormat, dateColIndex1Based) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clearContents();
  if (!rows || rows.length === 0) return;
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  const fmt = String(dateFormat || "").trim();
  const dateCol = Number(dateColIndex1Based) || 1;
  if (fmt && rows.length > 1) {
    sheet.getRange(2, dateCol, rows.length - 1, 1).setNumberFormat(fmt);
  }
}

function reconEnsureTotalsDetailed_(map, key) {
  if (!map[key]) {
    map[key] = { inflow: 0, outflow: 0, manualTxns: [], bankTxns: [] };
  }
  return map[key];
}

function reconBuildCombinedRows_(manualDaily, bankDaily) {
  const headers = [
    "Date",
    "RowType",
    "Source",
    "Sheet Row",
    "Title/Desc",
    "Amount",
    "Day Manual In",
    "Day Bank In",
    "Inflow Diff",
    "Day Manual Out",
    "Day Bank Out",
    "Outflow Diff",
    "Overall Diff"
  ];

  const rows = [headers];
  const dateKeys = reconCollectDateKeys_(manualDaily, bankDaily);

  dateKeys.forEach((dateKey) => {
    const manual = manualDaily[dateKey] || { inflow: 0, outflow: 0, manualTxns: [] };
    const bank = bankDaily[dateKey] || { inflow: 0, outflow: 0, bankTxns: [] };

    const inflowDiff = reconRoundCurrency_(manual.inflow - bank.inflow);
    const outflowDiff = reconRoundCurrency_(manual.outflow - bank.outflow);
    const overallDiff = reconRoundCurrency_(inflowDiff + outflowDiff);

    if (inflowDiff === 0 && outflowDiff === 0) return;

    const dateValue = reconDateKeyToDate_(dateKey);
    if (!dateValue) throw new Error(`Invalid normalized date key "${dateKey}".`);

    rows.push([
      dateValue,
      "SUMMARY",
      "",
      "",
      "Totals",
      "",
      manual.inflow,
      bank.inflow,
      inflowDiff,
      manual.outflow,
      bank.outflow,
      outflowDiff,
      overallDiff
    ]);

    (manual.manualTxns || []).forEach((txn) => {
      rows.push([
        dateValue,
        "TXN",
        "Manual",
        txn.rowNumber,
        txn.title || "",
        txn.amount,
        manual.inflow,
        bank.inflow,
        inflowDiff,
        manual.outflow,
        bank.outflow,
        outflowDiff,
        overallDiff
      ]);
    });

    (bank.bankTxns || []).forEach((txn) => {
      rows.push([
        dateValue,
        "TXN",
        "Bank",
        txn.rowNumber,
        txn.title || "",
        txn.amount,
        manual.inflow,
        bank.inflow,
        inflowDiff,
        manual.outflow,
        bank.outflow,
        outflowDiff,
        overallDiff
      ]);
    });
  });

  return rows;
}

function reconBuildSummaryRowsFromDaily_(manualDaily, bankDaily) {
  const headers = [
    "Date",
    "Manual Inflow",
    "Bank Inflow",
    "Inflow Diff",
    "Manual Outflow",
    "Bank Outflow",
    "Outflow Diff",
    "Overall Diff"
  ];
  const rows = [headers];
  const dateKeys = reconCollectDateKeys_(manualDaily, bankDaily);

  dateKeys.forEach((dateKey) => {
    const dateValue = reconDateKeyToDate_(dateKey);
    if (!dateValue) throw new Error(`Invalid normalized date key "${dateKey}".`);

    const manual = manualDaily[dateKey] || { inflow: 0, outflow: 0 };
    const bank = bankDaily[dateKey] || { inflow: 0, outflow: 0 };

    const inflowDiff = reconRoundCurrency_(manual.inflow - bank.inflow);
    const outflowDiff = reconRoundCurrency_(manual.outflow - bank.outflow);
    const overallDiff = reconRoundCurrency_(inflowDiff + outflowDiff);

    if (inflowDiff === 0 && outflowDiff === 0) return;

    rows.push([
      dateValue,
      manual.inflow,
      bank.inflow,
      inflowDiff,
      manual.outflow,
      bank.outflow,
      outflowDiff,
      overallDiff
    ]);
  });

  return rows;
}

function reconBuildDetailsRowsFromDaily_(manualDaily, bankDaily) {
  const headers = [
    "Date",
    "Source",
    "Sheet Row",
    "Title/Desc",
    "Amount",
    "Inflow Diff",
    "Outflow Diff",
    "Overall Diff"
  ];
  const rows = [headers];
  const dateKeys = reconCollectDateKeys_(manualDaily, bankDaily);

  dateKeys.forEach((dateKey) => {
    const manual = manualDaily[dateKey] || { inflow: 0, outflow: 0, manualTxns: [] };
    const bank = bankDaily[dateKey] || { inflow: 0, outflow: 0, bankTxns: [] };

    const inflowDiff = reconRoundCurrency_(manual.inflow - bank.inflow);
    const outflowDiff = reconRoundCurrency_(manual.outflow - bank.outflow);
    const overallDiff = reconRoundCurrency_(inflowDiff + outflowDiff);

    if (inflowDiff === 0 && outflowDiff === 0) return;

    const dateValue = reconDateKeyToDate_(dateKey);
    if (!dateValue) throw new Error(`Invalid normalized date key "${dateKey}".`);

    (manual.manualTxns || []).forEach((txn) => {
      rows.push([
        dateValue,
        "Manual",
        txn.rowNumber,
        txn.title || "",
        txn.amount,
        inflowDiff,
        outflowDiff,
        overallDiff
      ]);
    });

    (bank.bankTxns || []).forEach((txn) => {
      rows.push([
        dateValue,
        "Bank",
        txn.rowNumber,
        txn.title || "",
        txn.amount,
        inflowDiff,
        outflowDiff,
        overallDiff
      ]);
    });
  });

  return rows;
}
