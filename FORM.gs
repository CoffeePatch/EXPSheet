/**
 * FORM.gs
 * -----------------------------------------------------------------------------
 * Production-grade processor for Google Form expense/transfer rows.
 *
 * Goals:
 * - Idempotent-ish processing (skip already Processed=YES rows)
 * - Header-safe dynamic mapping
 * - Batch write to List sheet
 * - Batch mark Processed in Form sheet (single write)
 * - Locking to avoid concurrent trigger overlap
 * - Clear logging and strict validation
 * -----------------------------------------------------------------------------
 */

/** =========================
 *  CONFIGURATION
 *  ========================= */
const CONFIG = Object.freeze({
  SHEET_FORM: "Form",
  SHEET_LIST: "List",

  // Update if your form headers change.
  COL_SPLIT_PERSON: "Split Person",
  COL_TRANSFER_PERSON: "Transfer Person",

  // Required baseline headers in Form sheet
  REQUIRED_HEADERS: [
    "Timestamp",
    "Date",
    "Time",
    "Title",
    "Amount",
    "Transaction Type",
    "Processed"
  ],

  // Additional headers by transaction mode
  REQUIRED_TRANSFER_HEADERS: ["Out Account", "In Account"],
  REQUIRED_TRANSACTION_HEADERS: ["Account", "Category", "Type", "Transaction Notes"],

  DATE_FORMAT: "dd/MM/yyyy",
  TIME_FORMAT: "HH:mm",
  PROCESSED_YES: "YES",
  DEFAULT_PERSON: "me",
  LOCK_WAIT_MS: 30000
});

/**
 * Main entrypoint for trigger/manual run.
 */
function processFormResponses() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(CONFIG.LOCK_WAIT_MS)) {
    console.warn("processFormResponses: Could not obtain lock; exiting to avoid overlap.");
    return;
  }

  try {
    const ctx = buildContext_();
    const result = processUnprocessedRows_(ctx);

    if (result.outRows.length > 0) {
      appendToList_(ctx.listSheet, result.outRows);
    }

    if (result.processedSheetRows.length > 0) {
      markRowsProcessedBatch_(ctx.formSheet, ctx.indexes.processed + 1, result.processedSheetRows);
    }

    console.log(
      JSON.stringify({
        event: "processFormResponses.summary",
        scannedRows: result.scanned,
        generatedRows: result.outRows.length,
        processedRows: result.processedSheetRows.length,
        failedRows: result.failedRows
      })
    );
  } catch (err) {
    console.error(`processFormResponses.fatal: ${err && err.stack ? err.stack : err}`);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/** =========================
 *  Core processing
 *  ========================= */

/**
 * Build validated context with sheets, data, and header indexes.
 * @returns {{
 *   ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
 *   formSheet: GoogleAppsScript.Spreadsheet.Sheet,
 *   listSheet: GoogleAppsScript.Spreadsheet.Sheet,
 *   values: any[][],
 *   headers: string[],
 *   indexes: Object<string, number>
 * }}
 */
function buildContext_() {
  const ss = SpreadsheetApp.getActive();
  const formSheet = ss.getSheetByName(CONFIG.SHEET_FORM);
  const listSheet = ss.getSheetByName(CONFIG.SHEET_LIST);

  if (!formSheet || !listSheet) {
    throw new Error(`Missing sheet(s). Required: "${CONFIG.SHEET_FORM}" and "${CONFIG.SHEET_LIST}".`);
  }

  const values = formSheet.getDataRange().getValues();
  if (values.length === 0) {
    throw new Error(`"${CONFIG.SHEET_FORM}" is empty; header row is required.`);
  }

  const headers = values[0].map((h) => String(h || "").trim());
  const indexes = buildHeaderIndex_(headers);

  validateHeaders_(indexes, headers);

  return { ss, formSheet, listSheet, values, headers, indexes };
}

/**
 * Process all Form rows where Processed != YES.
 */
function processUnprocessedRows_(ctx) {
  const { values, indexes } = ctx;

  const outRows = [];
  const processedSheetRows = [];
  const failedRows = [];
  let scanned = 0;

  if (values.length < 2) {
    return { scanned, outRows, processedSheetRows, failedRows };
  }

  for (let r = 1; r < values.length; r++) {
    scanned++;
    const row = values[r];
    const sheetRow = r + 1;

    const processedVal = normalizeText_(row[indexes.processed]);
    if (processedVal === CONFIG.PROCESSED_YES) continue;

    try {
      const mapped = mapBaseRow_(row, indexes);
      const generated = generateOutputRows_(row, indexes, mapped);
      if (generated.length > 0) {
        outRows.push.apply(outRows, generated);
      }
      processedSheetRows.push(sheetRow);
    } catch (err) {
      failedRows.push({ row: sheetRow, error: String(err.message || err) });
      console.error(`Row ${sheetRow} processing error: ${err && err.stack ? err.stack : err}`);
      // Continue with next row; do not mark as processed.
    }
  }

  return { scanned, outRows, processedSheetRows, failedRows };
}

/**
 * Map row fields used in both transfer and transaction paths.
 */
function mapBaseRow_(row, idx) {
  const timestamp = row[idx.timestamp];
  const dateInput = row[idx.date] || timestamp;
  const timeInput = String(row[idx.time] || "").trim();

  const title = String(row[idx.title] || "").trim();
  const rawAmount = Math.abs(Number(row[idx.amount]));
  const trxTypeRaw = normalizeText_(row[idx.transactionType]);

  if (!isFinite(rawAmount) || rawAmount <= 0) {
    throw new Error(`Invalid amount "${row[idx.amount]}".`);
  }

  const finalDate = formatDate_(dateInput);
  if (!finalDate) throw new Error("Could not parse Date/Timestamp.");

  const finalTime = resolveTime_(timeInput, timestamp);

  return {
    finalDate,
    finalTime,
    title,
    rawAmount,
    isTransfer: trxTypeRaw.includes("TRANSFER"),
    isTransaction: trxTypeRaw.includes("TRANSAC")
  };
}

/**
 * Generate output rows for List sheet from one Form row.
 */
function generateOutputRows_(row, idx, mapped) {
  if (mapped.isTransfer) {
    return buildTransferRows_(row, idx, mapped);
  }

  if (mapped.isTransaction) {
    return buildTransactionRows_(row, idx, mapped);
  }

  // Unknown transaction type => do not process silently.
  throw new Error(`Unsupported Transaction Type "${row[idx.transactionType]}".`);
}

/** =========================
 *  Row builders
 *  ========================= */

function buildTransferRows_(row, idx, mapped) {
  const outAcc = String(row[idx.outAccount] || "").trim();
  const inAcc = String(row[idx.inAccount] || "").trim();

  if (!outAcc || !inAcc) {
    throw new Error("Transfer requires both Out Account and In Account.");
  }

  const personRaw = idx.transferPerson >= 0 ? String(row[idx.transferPerson] || "").trim() : "";
  const person = personRaw || CONFIG.DEFAULT_PERSON;
  const manualNotes = idx.transactionNotes >= 0 ? String(row[idx.transactionNotes] || "").trim() : "";
  const outNote = manualNotes ? `Transfer Out - ${manualNotes}` : "Transfer Out";
  const inNote = manualNotes ? `Transfer In - ${manualNotes}` : "Transfer In";

  return [
    [
      mapped.finalDate, mapped.finalTime, outAcc, mapped.title,
      -mapped.rawAmount, person, "Transfer", outNote, "Completed", "OUT", mapped.finalDate
    ],
    [
      mapped.finalDate, mapped.finalTime, inAcc, mapped.title,
      mapped.rawAmount, person, "Transfer", inNote, "Completed", "IN", mapped.finalDate
    ]
  ];
}

function buildTransactionRows_(row, idx, mapped) {
  const account = String(row[idx.account] || "").trim();
  const category = String(row[idx.category] || "").trim();
  const direction = normalizeText_(row[idx.type] || "OUT");
  const notes = String(row[idx.transactionNotes] || "").trim();

  if (!account) throw new Error("Transaction requires Account.");
  if (!category) throw new Error("Transaction requires Category.");
  if (direction !== "OUT" && direction !== "IN") {
    throw new Error(`Type must be IN or OUT, received "${row[idx.type]}".`);
  }

  const splitRaw =
    idx.splitPerson >= 0 && String(row[idx.splitPerson] || "").trim()
      ? String(row[idx.splitPerson])
      : CONFIG.DEFAULT_PERSON;

  const personList = parsePersons_(splitRaw);
  const count = personList.length || 1;
  const sign = direction === "OUT" ? -1 : 1;
  const perHeadAmount = (mapped.rawAmount / count) * sign;

  const isCreditorWallet = /credit|pending/i.test(account);

  return personList.map((person) => {
    const isMe = person.toLowerCase() === "me";

    const rowStatus = isMe && !isCreditorWallet ? "Settled" : "Pending";
    const settlementDate = rowStatus === "Settled" ? mapped.finalDate : "";

    const splitSuffix = count > 1 ? ` (Split: ${mapped.rawAmount}/${count})` : "";
    const note = `${notes}${splitSuffix}`.trim();

    return [
      mapped.finalDate, mapped.finalTime, account, mapped.title,
      perHeadAmount, person, category, note,
      rowStatus, direction, settlementDate
    ];
  });
}

/** =========================
 *  Batch write helpers
 *  ========================= */

function appendToList_(listSheet, rows) {
  const startRow = listSheet.getLastRow() + 1;
  listSheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  console.log(`Appended ${rows.length} row(s) to "${CONFIG.SHEET_LIST}".`);
}

/**
 * Mark non-contiguous row numbers as processed with one batch write strategy:
 * 1) read whole Processed column values once
 * 2) mutate target indexes in memory
 * 3) write full column back once
 */
function markRowsProcessedBatch_(formSheet, processedCol1Based, sheetRows) {
  const lastRow = formSheet.getLastRow();
  if (lastRow < 2) return;

  const height = lastRow - 1; // data rows only
  const rng = formSheet.getRange(2, processedCol1Based, height, 1);
  const vals = rng.getValues();

  for (let i = 0; i < sheetRows.length; i++) {
    const sheetRow = sheetRows[i];
    const dataIdx = sheetRow - 2; // row2 => idx0
    if (dataIdx >= 0 && dataIdx < vals.length) {
      vals[dataIdx][0] = CONFIG.PROCESSED_YES;
    }
  }

  rng.setValues(vals);
}

/** =========================
 *  Header and parsing helpers
 *  ========================= */

function buildHeaderIndex_(headers) {
  const raw = {};
  headers.forEach((h, i) => {
    if (!h) return;
    raw[h] = i;
  });

  // Canonical indexes used in code; -1 means optional/missing.
  return {
    timestamp: indexOrThrow_(raw, "Timestamp"),
    date: indexOrThrow_(raw, "Date"),
    time: indexOrThrow_(raw, "Time"),
    title: indexOrThrow_(raw, "Title"),
    amount: indexOrThrow_(raw, "Amount"),
    transactionType: indexOrThrow_(raw, "Transaction Type"),
    processed: indexOrThrow_(raw, "Processed"),

    outAccount: indexOrMinusOne_(raw, "Out Account"),
    inAccount: indexOrMinusOne_(raw, "In Account"),
    account: indexOrMinusOne_(raw, "Account"),
    category: indexOrMinusOne_(raw, "Category"),
    type: indexOrMinusOne_(raw, "Type"),
    transactionNotes: indexOrMinusOne_(raw, "Transaction Notes"),

    splitPerson: indexOrMinusOne_(raw, CONFIG.COL_SPLIT_PERSON),
    transferPerson: indexOrMinusOne_(raw, CONFIG.COL_TRANSFER_PERSON)
  };
}

function validateHeaders_(idx, headers) {
  const missingBase = CONFIG.REQUIRED_HEADERS.filter((h) => headers.indexOf(h) === -1);
  if (missingBase.length) {
    throw new Error(`Missing required headers: ${missingBase.join(", ")}`);
  }

  // At least one person column must exist.
  if (idx.splitPerson < 0 && idx.transferPerson < 0) {
    throw new Error(
      `Missing both person columns: "${CONFIG.COL_SPLIT_PERSON}" and "${CONFIG.COL_TRANSFER_PERSON}".`
    );
  }

  // Validate mode-specific minimums exist in header row (not necessarily used every row).
  const missingTransfer = CONFIG.REQUIRED_TRANSFER_HEADERS.filter((h) => headers.indexOf(h) === -1);
  const missingTxn = CONFIG.REQUIRED_TRANSACTION_HEADERS.filter((h) => headers.indexOf(h) === -1);

  if (missingTransfer.length) {
    console.warn(`Transfer headers missing: ${missingTransfer.join(", ")}`);
  }
  if (missingTxn.length) {
    console.warn(`Transaction headers missing: ${missingTxn.join(", ")}`);
  }
}

function indexOrThrow_(map, key) {
  if (map[key] === undefined) throw new Error(`Required header not found: "${key}"`);
  return map[key];
}

function indexOrMinusOne_(map, key) {
  return map[key] === undefined ? -1 : map[key];
}

function parsePersons_(raw) {
  const arr = String(raw || "")
    .split(/[.,;\n]+/)
    .map((p) => p.trim())
    .filter(Boolean);

  return arr.length ? arr : [CONFIG.DEFAULT_PERSON];
}

/** =========================
 *  Date/time/text helpers
 *  ========================= */

function formatDate_(value) {
  const d = new Date(value);
  if (isNaN(d.getTime())) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), CONFIG.DATE_FORMAT);
}

function resolveTime_(input, fallbackTimestamp) {
  const clean = String(input || "").trim();
  if (clean) {
    const parsed = parseTimeInput_(clean);
    return parsed || clean;
  }

  const d = new Date(fallbackTimestamp);
  if (isNaN(d.getTime())) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), CONFIG.TIME_FORMAT);
}

function parseTimeInput_(input) {
  const raw = String(input || "").trim().toLowerCase().replace(/\s+/g, "");
  if (!raw) return "";

  const namedTimeMap = {
    morning: "09:00",
    afternoon: "12:00",
    evening: "17:00",
    night: "21:00",
    knight: "21:00"
  };
  if (namedTimeMap[raw]) return namedTimeMap[raw];

  let m = raw.match(/^(\d{1,2}):(\d{2})(a|p|am|pm)?$/);
  if (m) {
    const hour = Number(m[1]);
    const minute = Number(m[2]);
    const suffix = m[3] || "";
    return normalizeHourMinute_(hour, minute, suffix);
  }

  m = raw.match(/^(\d{1,2})(\d{2})(a|p|am|pm)$/);
  if (m) {
    const hour = Number(m[1]);
    const minute = Number(m[2]);
    const suffix = m[3] || "";
    return normalizeHourMinute_(hour, minute, suffix);
  }

  m = raw.match(/^(\d{1,2})(a|p|am|pm)$/);
  if (m) {
    const hour = Number(m[1]);
    const suffix = m[2] || "";
    return normalizeHourMinute_(hour, 0, suffix);
  }

  return "";
}

function normalizeHourMinute_(hour, minute, suffix) {
  if (!isFinite(hour) || !isFinite(minute) || minute < 0 || minute > 59) return "";

  let h = hour;
  if (!suffix && (h < 0 || h > 23)) return "";

  if (suffix) {
    if (h < 1 || h > 12) return "";
    const isPm = suffix === "p" || suffix === "pm";
    if (isPm && h !== 12) h += 12;
    if (!isPm && h === 12) h = 0;
  }

  return `${String(h).padStart(2, "0")}:${String(minute).padStart(2, "0")}`;
}

function normalizeText_(value) {
  return String(value || "").trim().toUpperCase();
}
