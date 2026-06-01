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

Recommended workflow:

- Use option 2 if you want the cleanest reconciliation experience.
- Use option 3 if you want a simpler change that still makes filtering by date much easier.
- Keep option 1 only if the goal is just to compare daily totals.

If you want, the next step can be to convert the script to one of these formats before changing the actual output logic.

## Required Sheets

| Sheet | Purpose |
|---|---|
| `List` | Ledger to reconcile. The script reads the date, account, and amount columns. |
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
- `BANK_SPREADSHEET_ID` (optional, when the bank statement is in a different workbook)
- `BANK_SHEET_NAME` (optional, bank sheet name override)
- `RECON_USE_OVERLAP_START_DATE` (optional)
- `OUTPUT_MODE`
- `SHEET_RECON_SUMMARY` (when `OUTPUT_MODE` is `TWO_SHEETS`)
- `SHEET_RECON_DETAILS` (when `OUTPUT_MODE` is `TWO_SHEETS`)
- `BANK_DATE_HEADERS`
- `BANK_DESC_HEADERS` (optional, for detailed output)
- `BANK_DEBIT_HEADER`
- `BANK_CREDIT_HEADER`
 - `MANUAL_DESC_HEADERS` (optional, for detailed output)
 - `RECON_START_DATE` (optional)
 - `RECON_END_DATE` (optional)

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
	- Includes a `SUMMARY` row per mismatched date plus all manual and bank `TXN` rows for that date.
	- Every `TXN` row repeats the day totals and diffs so you can sort/filter by `Overall Diff`.
- `TWO_SHEETS`: writes two sheets.
	- `SHEET_RECON_SUMMARY`: one row per mismatched date.
	- `SHEET_RECON_DETAILS`: all manual + bank transactions for mismatched dates (in a single sheet).

If `RECON_START_DATE` and/or `RECON_END_DATE` are set in `RECON_CONFIG`, the script restricts reconciliation to that inclusive date range. Behavior:

- If only `RECON_START_DATE` is provided, the script uses the bank statement's last date as the end date.
- If only `RECON_END_DATE` is provided, the script uses the bank statement's first date as the start date.
- If both are omitted, the script uses the full bank statement range (first to last date).

Date value format notes:

- In Apps Script / JavaScript, **do not** write dates like `01/05/2026` without quotes. That gets evaluated as math (`1/5/2026`) and will not be treated as a date.
- Use one of these instead:
	- A `Date` object: `new Date(2026, 4, 1)` (months are 0-based: 4 = May)
	- A string: `"01/05/2026"` (respects `RECON_DATE_ORDER`) or `"2026-05-01"`

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
3. (Optional) Set `RECON_START_DATE` and/or `RECON_END_DATE` in `RECON_CONFIG` to restrict the range.
3. Paste the bank export into `Bank_Raw`.
4. Run `reconcileBankStatement()` from the Apps Script editor.

## Notes

- The script uses column A for date, column C for account, and column E for amount on `List`.
- It does not change the `List` sheet.
- If the comparison produces no differences, the log still gets rebuilt with the header row.