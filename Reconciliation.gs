/**
 * Reconciliation.gs
 * -----------------------------------------------------------------------------
 * Manual reconciliation between the List sheet and a pasted bank statement.
 * -----------------------------------------------------------------------------
 */

const RECON_CONFIG = Object.freeze({
  SHEET_LIST: "List",
  SHEET_BANK_RAW: "Bank_Raw",
  SHEET_RECON_LOG: "Reconciliation_Log",

  RECON_TARGET_ACCOUNT: "", // TODO: set this to your account string (e.g., "9682") before running.
  RECON_DATE_ORDER: "DMY",

  BANK_DATE_HEADERS: ["Txn Date", "Value Date"],
  BANK_DEBIT_HEADER: "Debit",
  BANK_CREDIT_HEADER: "Credit"
});

function reconcileBankStatement() {
  const ss = SpreadsheetApp.getActive();
  const listSheet = ss.getSheetByName(RECON_CONFIG.SHEET_LIST);
  const bankSheet = ss.getSheetByName(RECON_CONFIG.SHEET_BANK_RAW);

  if (!listSheet || !bankSheet) {
    throw new Error(
      `Missing sheet(s). Required: "${RECON_CONFIG.SHEET_LIST}" and "${RECON_CONFIG.SHEET_BANK_RAW}".`
    );
  }

  if (!RECON_CONFIG.RECON_TARGET_ACCOUNT) {
    throw new Error("RECON_CONFIG.RECON_TARGET_ACCOUNT must be set to a List sheet account name.");
  }

  const tz = ss.getSpreadsheetTimeZone();
  const manualValues = listSheet.getDataRange().getValues();
  const bankValues = bankSheet.getDataRange().getValues();

  const manualTotals = reconAggregateManualTotals_(manualValues, RECON_CONFIG.RECON_TARGET_ACCOUNT, tz);
  const bankMeta = reconResolveBankColumns_(bankValues);
  const bankTotals = reconAggregateBankTotals_(bankValues, bankMeta, tz);

  const logRows = reconBuildLogRows_(manualTotals, bankTotals);
  reconWriteLog_(ss, logRows);
}

function reconAggregateManualTotals_(values, accountFilter, tz) {
  const totals = {};
  if (!values || values.length < 2) return totals;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const account = String(row[2] || "").trim();
    if (account !== accountFilter) continue;

    const dateKey = reconNormalizeDateKey_(row[0], tz, RECON_CONFIG.RECON_DATE_ORDER);
    if (!dateKey) continue;

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

function reconResolveBankColumns_(values) {
  if (!values || values.length === 0) {
    throw new Error("Bank_Raw is empty; paste your bank statement before reconciling.");
  }

  const dateHeaders = (RECON_CONFIG.BANK_DATE_HEADERS || []).map(reconNormalizeHeader_);
  const debitHeader = reconNormalizeHeader_(RECON_CONFIG.BANK_DEBIT_HEADER);
  const creditHeader = reconNormalizeHeader_(RECON_CONFIG.BANK_CREDIT_HEADER);

  for (let r = 0; r < values.length; r++) {
    const row = values[r] || [];
    const headers = row.map(reconNormalizeHeader_);
    const dateIdx = headers.findIndex((header) => dateHeaders.includes(header));

    if (dateIdx < 0) continue;

    const debitIdx = headers.indexOf(debitHeader);
    const creditIdx = headers.indexOf(creditHeader);

    if (debitIdx < 0 || creditIdx < 0) {
      throw new Error(
        `Bank_Raw header row found at row ${r + 1} but missing "${RECON_CONFIG.BANK_DEBIT_HEADER}" or "${RECON_CONFIG.BANK_CREDIT_HEADER}".`
      );
    }

    return { startRow: r + 1, dateIdx, debitIdx, creditIdx };
  }

  const dateHeaderLabel = RECON_CONFIG.BANK_DATE_HEADERS.join('" or "');
  throw new Error(
    `Bank_Raw headers not found. Expected "${dateHeaderLabel}" plus "${RECON_CONFIG.BANK_DEBIT_HEADER}" and "${RECON_CONFIG.BANK_CREDIT_HEADER}".`
  );
}

function reconAggregateBankTotals_(values, meta, tz) {
  const totals = {};
  if (!values || values.length === 0) return totals;

  for (let r = meta.startRow; r < values.length; r++) {
    const row = values[r];
    const dateKey = reconNormalizeDateKey_(row[meta.dateIdx], tz, RECON_CONFIG.RECON_DATE_ORDER);
    if (!dateKey) continue;

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
    "Outflow Diff"
  ];

  const rows = [headers];
  const dateKeys = reconCollectDateKeys_(manualTotals, bankTotals);

  dateKeys.forEach((dateKey) => {
    const manual = manualTotals[dateKey] || { inflow: 0, outflow: 0 };
    const bank = bankTotals[dateKey] || { inflow: 0, outflow: 0 };

    const inflowDiff = reconRoundCurrency_(manual.inflow - bank.inflow);
    const outflowDiff = reconRoundCurrency_(manual.outflow - bank.outflow);

    if (inflowDiff !== 0 || outflowDiff !== 0) {
      rows.push([
        dateKey,
        manual.inflow,
        bank.inflow,
        inflowDiff,
        manual.outflow,
        bank.outflow,
        outflowDiff
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

function reconWriteLog_(ss, rows) {
  let logSheet = ss.getSheetByName(RECON_CONFIG.SHEET_RECON_LOG);
  if (!logSheet) {
    logSheet = ss.insertSheet(RECON_CONFIG.SHEET_RECON_LOG);
  }

  logSheet.clearContents();
  if (rows.length === 0) return;
  logSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
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

    if (first <= 12 && second > 12) {
      day = second;
      month = first;
    } else if (first <= 12 && second <= 12 && preference === "MDY") {
      day = second;
      month = first;
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

function reconIsValidDate_(year, month, day) {
  if (!year || !month || !day) return false;
  if (month < 1 || month > 12) return false;
  if (day < 1 || day > 31) return false;
  return true;
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

function reconEnsureTotals_(map, key) {
  if (!map[key]) {
    map[key] = { inflow: 0, outflow: 0 };
  }
  return map[key];
}
