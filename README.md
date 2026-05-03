# EXPSheet – Expense Tracker & Bank Reconciliation

EXPSheet is a Google Apps Script–based expense tracker that maintains a clean `List` ledger and provides a **bank reconciliation helper** to catch missed or mismatched manual entries. This README focuses on using the reconciliation workflow with **State Bank of India (SBI)** exports.

---

## The Core Problem (Why This Script Exists)

Manual expense tracking is prone to missed entries, duplicate entries, or incorrect dates. `Reconciliation.gs` solves this by **cross-referencing daily inflow/outflow totals** from your `List` sheet against totals computed from your raw SBI bank export. The reconciliation log highlights **only the dates where the manual ledger and the bank statement disagree**, so you can quickly spot what needs to be fixed.

---

## Prerequisites & Setup

### 1) Google Sheets structure

Your spreadsheet must include these tabs:

- **`List`**: Your normalized ledger (manual tracking output).
- **`Bank_Raw`**: Where you paste the SBI statement export.
- **`Reconciliation_Log`**: Created/overwritten by the script.

Required `List` columns (by position):

| Column | Required content |
|---|---|
| A | Date |
| C | Account (string identifier) |
| E | Amount (positive = inflow, negative = outflow) |

### 2) SBI export headers

Your SBI export **does not have headers on row 1**. The script scans down until it finds a header row containing **`Txn Date`** or **`Value Date`**, plus **`Debit`** and **`Credit`**. These headers must appear **in the same row**.

### 3) Set `RECON_TARGET_ACCOUNT`

Open `Reconciliation.gs` and set the account you want to reconcile in `RECON_CONFIG` near the top of the file:

```javascript
const RECON_CONFIG = Object.freeze({
  // ...
  RECON_TARGET_ACCOUNT: "9682", // exact List sheet account name in column C (case-sensitive)
});
```

This is required; the script will stop if it is empty.

---

## Configuration Reference (RECON_CONFIG)

| Field | Purpose |
|---|---|
| `SHEET_LIST` | Name of your ledger tab (`List` by default). |
| `SHEET_BANK_RAW` | Name of the tab containing pasted bank exports. |
| `SHEET_RECON_LOG` | Output tab for reconciliation results. |
| `RECON_TARGET_ACCOUNT` | **Required.** Exact account string to match in `List` column C. |
| `RECON_DATE_ORDER` | `DMY` or `MDY` when dates are ambiguous (e.g., `03/04/2026`). |
| `BANK_DATE_HEADERS` | Acceptable date headers (default: `Txn Date` or `Value Date`). |
| `BANK_DEBIT_HEADER` | SBI debit column header (default: `Debit`). |
| `BANK_CREDIT_HEADER` | SBI credit column header (default: `Credit`). |

---

## Step-by-Step Execution (SBI Workflow)

1. **Download the SBI statement**
   - From SBI internet banking, export the statement as **Excel**.
2. **Paste into `Bank_Raw`**
   - Open the `Bank_Raw` tab.
   - Clear existing content (optional).
   - Paste the **entire** export (including the messy header block above the actual column headers).
3. **Confirm configuration**
   - Ensure `RECON_TARGET_ACCOUNT` matches the exact account string in your `List` sheet column C.
   - Adjust `RECON_DATE_ORDER` if needed.
4. **Run the reconciliation**
   - In Apps Script, run `reconcileBankStatement()`.
5. **Review `Reconciliation_Log`**
   - The log shows only dates with mismatches.

---

## Understanding the Output (`Reconciliation_Log`)

The script writes these columns:

| Column | Meaning |
|---|---|
| Date | Normalized as `YYYY-MM-DD`. |
| Manual Inflow | Sum of positive amounts in `List` for that date. |
| Bank Inflow | Sum of SBI `Credit` values for that date. |
| Inflow Diff | `Manual Inflow - Bank Inflow`. |
| Manual Outflow | Sum of negative amounts in `List` for that date. |
| Bank Outflow | Sum of SBI `Debit` values as **negative totals**. |
| Outflow Diff | `Manual Outflow - Bank Outflow`. |

### How to interpret the diffs

- **Inflow Diff**
  - `> 0`: Manual ledger shows more inflow than the bank (possible extra entry).
  - `< 0`: Bank shows more inflow than the ledger (missing income entry).

- **Outflow Diff** (remember: outflows are negative totals)
  - `> 0`: Manual outflow is **less negative** than the bank (missing expense entry).
  - `< 0`: Manual outflow is **more negative** than the bank (possible duplicate or over-recorded expense).

Only rows with **non-zero diffs** are listed, so the log is focused on issues that need attention.

---

## Notes & Limitations

- The reconciliation is **single-account**: it only compares rows in `List` where column C matches `RECON_TARGET_ACCOUNT` exactly.
- The script **does not modify** your `List` sheet.
- The log is **overwritten** on every run.
- If your SBI export uses different header names, update `BANK_DATE_HEADERS`, `BANK_DEBIT_HEADER`, and `BANK_CREDIT_HEADER` accordingly.

---

## Troubleshooting

- **Error: `Bank_Raw is empty`** → Paste your SBI export into `Bank_Raw`.
- **Error: headers not found** → Ensure the header row includes `Txn Date` or `Value Date` plus `Debit` and `Credit` in the same row.
- **Empty log** → No discrepancies detected, or `RECON_TARGET_ACCOUNT` does not match any account in `List` column C.

---

## License

This project is MIT-licensed. Use at your own risk and always keep backups of your spreadsheet data.
