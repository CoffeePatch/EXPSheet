# Bank Reconciliation (Reconciliation.gs)

## The Core Problem

Manual expense tracking can miss transactions, double-count entries, or record the wrong day. `scripts/reconciliation/Reconciliation.gs` exists to **catch those gaps** by comparing your **daily inflow/outflow totals** in the `List` sheet against totals computed from a raw **bank statement export** (for example SBI). Only dates with mismatches are surfaced so you can quickly fix the manual ledger.

---

## Prerequisites & Setup

### 1) Required sheets

Your spreadsheet must include:

- **`List`**: your manual ledger.
- **`Bank_Raw`**: where you paste the bank statement export.
- **`Reconciliation_Log`**: created/overwritten by the script.

### 2) `List` sheet format

The reconciliation logic expects these columns by position:

| Column | Required content |
|---|---|
| A | Date |
| C | Account (string identifier) |
| E | Amount (positive = inflow, negative = outflow) |

### 3) Bank export headers

The script scans `Bank_Raw` until it finds a header row containing:

- **`Txn Date`** or **`Value Date`**
- **`Debit`**
- **`Credit`**

These must appear **on the same row**. If your export uses different labels, update `BANK_DATE_HEADERS`, `BANK_DEBIT_HEADER`, and `BANK_CREDIT_HEADER` in `RECON_CONFIG`.

### 4) Required configuration

In `scripts/reconciliation/Reconciliation.gs`, set the account you want to reconcile and confirm your tab names:

```javascript
const RECON_CONFIG = Object.freeze({
  SHEET_LIST: "List",
  SHEET_BANK_RAW: "Bank_Raw",
  SHEET_RECON_LOG: "Reconciliation_Log",
  RECON_LOG_DATE_FORMAT: "dd/MM/yyyy",
  RECON_START_DATE: "", // optional
  RECON_END_DATE: "", // optional
  RECON_TARGET_ACCOUNT: "9682", // exact List sheet account name in column C (case-sensitive)
});
```

The script will stop if `RECON_TARGET_ACCOUNT` or the sheet names are not set.

---

## Step-by-Step Execution (Bank Statement Workflow)

1. **Download the bank statement**
   - Export the statement as **Excel** from your bank (for example SBI internet banking).
2. **Paste into `Bank_Raw`**
   - Open `Bank_Raw` and paste the **entire** export (including any header blocks above the column headers).
3. **Confirm configuration**
   - Ensure `RECON_TARGET_ACCOUNT` matches the exact account string used in `List` column C.
   - Set `RECON_DATE_ORDER` if your dates are ambiguous (e.g., `03/04/2026`).
   - Use `RECON_START_DATE`/`RECON_END_DATE` to limit the reconciliation window. If you only set one date, the script will use the first/last bank statement row to complete the range.
4. **Run the script**
   - In Apps Script, run `reconcileBankStatement()`.
5. **Review `Reconciliation_Log`**
   - The log lists only dates where manual and bank totals differ.

---

## Understanding the Output (`Reconciliation_Log`)

The script writes these columns:

| Column | Meaning |
|---|---|
| Date | Displayed using `RECON_LOG_DATE_FORMAT` (default `dd/MM/yyyy`). |
| Manual Inflow | Sum of positive amounts in `List` for that date. |
| Bank Inflow | Sum of bank `Credit` values for that date. |
| Inflow Diff | `Manual Inflow - Bank Inflow`. |
| Manual Outflow | Sum of negative amounts in `List` for that date. |
| Bank Outflow | Sum of bank `Debit` values as **negative totals**. |
| Outflow Diff | `Manual Outflow - Bank Outflow`. |
| Overall Diff | `Inflow Diff + Outflow Diff` (net difference for the date). |

### What the diffs mean

- **Inflow Diff**
  - `> 0`: Manual ledger shows more inflow than the bank (possible extra entry).
  - `< 0`: Bank shows more inflow than the ledger (missing income entry).

- **Outflow Diff** (outflows are negative totals)
  - `> 0`: Manual outflow is **less negative** than the bank (missing expense entry).
  - `< 0`: Manual outflow is **more negative** than the bank (duplicate or over-recorded expense).

Only rows with **non-zero diffs** appear in the log.

---

## Notes

- Reconciliation is **single-account**: only `List` rows where column C equals `RECON_TARGET_ACCOUNT` are included.
- `Reconciliation_Log` is **overwritten** on each run.
- The script **does not modify** `List`.

### Date range behavior

- If `RECON_START_DATE` is set and `RECON_END_DATE` is blank, the end date defaults to the **last dated row** in `Bank_Raw`.
- If `RECON_END_DATE` is set (for December-only workflows) and `RECON_START_DATE` is blank, the start date defaults to the **first dated row** in `Bank_Raw`.
- If both dates are set, reconciliation runs only within that explicit range.
- If neither date is set, the entire statement range is used (first to last dated row).
