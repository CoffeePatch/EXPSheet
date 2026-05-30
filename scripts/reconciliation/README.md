# Bank Reconciliation

This folder contains the manual reconciliation workflow for comparing the `List` ledger against a pasted bank export.

## Main Script

- [Reconciliation.gs](Reconciliation.gs) contains `reconcileBankStatement()`.

## What The Script Does

`reconcileBankStatement()` totals inflows and outflows from `List`, reads the export in `Bank_Raw`, and writes only the dates with differences into `Reconciliation_Log`.

- The ledger is not modified.
- The log is cleared and rebuilt on each run.
- Dates are normalized before the comparison happens.

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
- `BANK_DATE_HEADERS`
- `BANK_DEBIT_HEADER`
- `BANK_CREDIT_HEADER`

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

- The script uses column A for date, column C for account, and column E for amount on `List`.
- It does not change the `List` sheet.
- If the comparison produces no differences, the log still gets rebuilt with the header row.