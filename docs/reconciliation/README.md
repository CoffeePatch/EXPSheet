# Bank Reconciliation

This folder contains the reconciliation helper and the manual setup notes for comparing `List` against a bank export pasted into `Bank_Raw`.

`Reconciliation.gs` now supports custom date ranges:

- start date only set: the script uses the last bank-row date as the end bound
- end date only set: the script uses the first bank-row date as the start bound
- both dates set: the script uses that exact range
- neither date set: the script uses the full statement range from the first bank row to the last bank row

---

## Required Sheets

| Sheet | Purpose |
|---|---|
| `List` | Manual ledger to reconcile. |
| `Bank_Raw` | Paste the bank export here. |
| `Reconciliation_Log` | Output sheet created or overwritten by the script. |

The script reads `List` by position:

| Column | Required content |
|---|---|
| A | Date |
| C | Account identifier |
| E | Amount, where positive values are inflow and negative values are outflow |

---

## Configuration

Update the `RECON_CONFIG` block in [Reconciliation.gs](Reconciliation.gs):

| Setting | Meaning |
|---|---|
| `SHEET_LIST` | Name of the ledger tab. |
| `SHEET_BANK_RAW` | Name of the bank export tab. |
| `SHEET_RECON_LOG` | Name of the output tab. |
| `RECON_TARGET_ACCOUNT` | Exact account name to reconcile from `List` column C. |
| `RECON_DATE_ORDER` | `DMY` or `MDY` for ambiguous dates. |
| `RECON_START_DATE` | Optional start date. |
| `RECON_END_DATE` | Optional end date. |
| `BANK_DATE_HEADERS` | Acceptable date header names in the bank export. |
| `BANK_DEBIT_HEADER` | Debit header name in the export. |
| `BANK_CREDIT_HEADER` | Credit header name in the export. |

Use either a real date value or a parsable date string for `RECON_START_DATE` and `RECON_END_DATE`.

---

## Step-by-Step Setup

1. Export your bank statement and paste the full sheet into `Bank_Raw`, including any header rows above the transaction table.
2. Confirm the bank export includes a header row with a date field plus `Debit` and `Credit` columns.
3. Set `RECON_TARGET_ACCOUNT` to the exact value used in `List` column C.
4. Optionally set `RECON_START_DATE` and/or `RECON_END_DATE`.
5. Save the script.
6. Run `reconcileBankStatement()` from the Apps Script editor.
7. Review `Reconciliation_Log` for the dates with non-zero differences.

---

## Output Columns

The script writes these columns to `Reconciliation_Log`:

| Column | Meaning |
|---|---|
| Date | Displayed using `RECON_LOG_DATE_FORMAT`. |
| Manual Inflow | Positive amounts from `List` for the date. |
| Bank Inflow | Bank `Credit` totals for the date. |
| Inflow Diff | `Manual Inflow - Bank Inflow`. |
| Manual Outflow | Negative amounts from `List` for the date. |
| Bank Outflow | Bank `Debit` totals as negative values. |
| Outflow Diff | `Manual Outflow - Bank Outflow`. |
| Overall Diff | `Inflow Diff + Outflow Diff`. |

Only dates with non-zero diffs are written.

---

## Tips

- Reconciliation is single-account. Only `List` rows matching `RECON_TARGET_ACCOUNT` are included.
- `Reconciliation_Log` is overwritten on every run.
- The script does not modify `List`.
- If the bank export uses different labels, update the header constants in `RECON_CONFIG` instead of changing the raw data.
