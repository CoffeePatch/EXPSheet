# Bank Reconciliation

This folder contains the manual reconciliation workflow for comparing the `List` ledger against a pasted bank export.

## Main Script

- [Reconciliation.gs](Reconciliation.gs) contains `reconcileBankStatement()`.

## What The Script Does

`reconcileBankStatement()` totals inflows and outflows from `List`, reads the export in `Bank_Raw`, and writes only the dates with differences into `Reconciliation_Log`.

- The ledger is not modified.
- The log is cleared and rebuilt on each run.
- Dates are normalized before the comparison happens.

## Output Format Options

The current log is optimized for spotting date-level mismatches, but it is not ideal when you want to sort or filter individual transactions. For that reason, there are three practical output styles:

1. Date summary only.
	- One row per day.
	- Shows the total inflow and outflow difference for that date.
	- Good for quick variance checks, but weaker for filtering specific transactions.

2. Transaction-level reconciliation log.
	- One row per transaction.
	- Includes the transaction date, description, account, bank amount, manual amount, and difference.
	- Best for sorting, filtering, and tracing a mismatch back to the exact transaction.

3. Day grouped transaction log with repeated date rows.
	- Copies all transactions for each day into the output sheet.
	- Repeats the date on every transaction row and adds a difference column.
	- Easier to sort and filter than a summary-only sheet, while keeping the output simple.

## Required Sheets

| Sheet | Purpose |
|---|---|
| `List` | Ledger to reconcile. The script reads the Date, Account, and Amount columns. |
| `Bank_Raw` | Paste the bank statement export here. |
| `Reconciliation_Log` | Output sheet created or overwritten by the script. |

## Configuration

Update the `RECON_CONFIG` block in [Reconciliation.gs](Reconciliation.gs).

Key values:

- `SHEET_LIST`
- `SHEET_BANK_RAW`
- `SHEET_RECON_LOG`
- `RECON_TARGET_ACCOUNT`
- `RECON_DATE_ORDER`
- `RECON_LOG_DATE_FORMAT`

### External Bank Workbook (Optional)

If your bank statement is stored in a separate Google Sheets file:

- Set `BANK_SPREADSHEET_ID` to the external spreadsheet ID.
- Set `BANK_SHEET_NAME` to the sheet/tab name inside that spreadsheet.

If `BANK_SHEET_NAME` is not set, the script falls back to `SHEET_BANK_RAW` as the tab name.
If `BANK_SPREADSHEET_ID` is not set, the script reads `SHEET_BANK_RAW` from the active spreadsheet (current behavior).

### Overlap Window (Optional)

If `RECON_USE_OVERLAP_START_DATE` is `true`, the script ignores any dates that occur before the manual ledger has started (within the selected range). This helps avoid early-bank-date mismatches when the bank export includes dates earlier than the first manual transaction.

### OUTPUT_MODE

`OUTPUT_MODE` controls what gets written to output sheets:

- `SUMMARY_ONLY` (default): current behavior. Writes one row per mismatched date into `SHEET_RECON_LOG`.
- `SINGLE_SHEET`: writes a single flat, sortable table into `SHEET_RECON_LOG`.
- `TWO_SHEETS`: writes two sheets (`SHEET_RECON_SUMMARY` and `SHEET_RECON_DETAILS`).

## Bank Export Expectations

The script searches `Bank_Raw` for a header row that contains one of the configured date headers plus the debit and credit headers.

| Setting | Default |
|---|---|
| Date headers | `Txn Date`, `Value Date` |
| Debit header | `Debit` |
| Credit header | `Credit` |

If your export uses different labels, update the constants in `RECON_CONFIG` instead of renaming the sheet manually.

## Run Steps

1. Set `RECON_TARGET_ACCOUNT` to the exact account name you want to reconcile.
2. Confirm the sheet names in `RECON_CONFIG` match your workbook.
3. Paste the bank export into `Bank_Raw`.
4. Run `reconcileBankStatement()` from the Apps Script editor.

## Notes

- The script dynamically searches for columns named "Date", "Account", and "Amount" on `List` (case-insensitive).
- It does not change the `List` sheet.
- If the comparison produces no differences, the log still gets rebuilt with the header row.