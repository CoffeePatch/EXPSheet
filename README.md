# EXPSheet

A lightweight companion for AppSheet-based expense tracking in Google Sheets, with an optional Apps Script bank reconciliation helper.

---

## Table of Contents

1. [Features](#features)
2. [Sheet Structure](#sheet-structure)
3. [Setup Requirements](#setup-requirements)
4. [Bank Reconciliation (Manual vs Bank)](#bank-reconciliation-manual-vs-bank)
5. [Contributing](#contributing)
6. [License](#license)
7. [Disclaimer](#disclaimer)

---

## Features

| # | Feature | Details |
|---|---------|---------|
| 1 | **AppSheet-driven entry (no Apps Script required)** | Use AppSheet to capture expenses and transfers directly into Google Sheets. |
| 2 | **Optional bank reconciliation helper** | `scripts/reconciliation/Reconciliation.gs` compares your `List` tab against a pasted bank statement. |
| 3 | **Flexible sheet add-ons** | Optional tabs like `Data`, `Balance`, `LUX`, or `transactions` can be added for your own analysis. |

---

## Sheet Structure

The spreadsheet is expected to contain the following tabs:

| Tab name | Role |
|----------|------|
| `List` | **Primary transaction log.** Required if you use the reconciliation script. |
| `Bank_Raw` | *(Reconciliation only)* Paste your bank export here. |
| `Reconciliation_Log` | *(Reconciliation only)* Output from the reconciliation script. |
| `Data` | *(Optional)* Static reference data — categories, accounts, and people used in AppSheet dropdowns. |
| `Balance` | *(Optional)* Summary formulas and per-person or per-account balance calculations. |
| `LUX` | *(Custom)* User-specific tracking or dashboard view. |
| `transactions` | *(Custom)* Advanced or archival transaction log. |

AppSheet can write directly to the `List` sheet using your own schema. If you plan to use the reconciliation helper, ensure `List` has Date in column A, Account in column C, and Amount in column E.

---

## Bank Reconciliation (Manual vs Bank)

The script includes a manual reconciliation helper in `scripts/reconciliation/Reconciliation.gs` to compare your `List` tab against a pasted bank statement.

**Expected tabs**
- `List` (existing): Date in column A, Account in column C, Amount in column E.
- `Bank_Raw`: Paste your bank export here. The script scans down to find a header row containing `Txn Date` or `Value Date`, plus `Debit` and `Credit` (or update the header names in `RECON_CONFIG`).
- `Reconciliation_Log`: Created/overwritten by the script.
  - If your tab names differ, update `SHEET_LIST`, `SHEET_BANK_RAW`, and `SHEET_RECON_LOG` in `RECON_CONFIG`.

**How to run**
1. Update `RECON_CONFIG` values in `scripts/reconciliation/Reconciliation.gs`:
    - `RECON_TARGET_ACCOUNT` (account name to match in `List`)
    - `RECON_DATE_ORDER` (`DMY` or `MDY`, used when dates are ambiguous like `03/04/2026`)
    - `SHEET_LIST`, `SHEET_BANK_RAW`, `SHEET_RECON_LOG` (if you use different tab names)
    - `RECON_LOG_DATE_FORMAT` (format for the Date column in `Reconciliation_Log`)
    - `RECON_START_DATE`, `RECON_END_DATE` (optional date bounds; blank defaults to statement start/end)
    - `BANK_DATE_HEADERS` (acceptable date header names in `Bank_Raw`)
    - `BANK_DEBIT_HEADER`, `BANK_CREDIT_HEADER` (debit/credit header names in `Bank_Raw`)
2. Run `reconcileBankStatement()` from the Apps Script editor.

The log outputs: Date, Manual Inflow, Bank Inflow, Inflow Diff, Manual Outflow, Bank Outflow, Outflow Diff, Overall Diff. Dates are formatted using `RECON_LOG_DATE_FORMAT`, and the reconciliation process does not modify the `List` sheet.

---

## Setup Requirements

- A **Google Account** with access to Google Sheets and AppSheet.
- Your own **Google Sheet** with the tabs and columns described above, *or* a copy of the EXPSheet template.
- **Apps Script permissions** are only required if you run the reconciliation helper in `scripts/reconciliation/Reconciliation.gs`.

---

## Contributing

Contributions, bug reports, and feature requests are welcome!

1. **Open an issue** at [github.com/CoffeePatch/EXPSheet/issues](https://github.com/CoffeePatch/EXPSheet/issues) to discuss what you would like to change.
2. **Fork** the repository and create a branch for your changes.
3. Submit a **pull request** with a clear description of what you changed and why.

---

## License

This project is open-source under the MIT License. A `LICENSE` file will be added to the repository root in a future update.

---

## Disclaimer

This project is community-maintained and is **not** an official Google product.  
Always keep a backup of your spreadsheet data before deploying script updates or making structural changes to your sheets.  
The authors accept no liability for data loss or inaccuracies resulting from the use of this software.
