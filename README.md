# EXPSheet

A professional, extensible **Google Sheets + Apps Script** toolkit for expense tracking and reconciliation, with AppSheet-based data entry guidance. AppSheet is Google’s no-code application development platform for building apps on top of data sources like Google Sheets.

---

## Table of Contents

1. [Features](#features)
2. [Sheet Structure](#sheet-structure)
3. [Setup Requirements](#setup-requirements)
4. [AppSheet Input Types & Fill Methods](#appsheet-input-types--fill-methods)
5. [Bank Reconciliation (Manual vs Bank)](#bank-reconciliation-manual-vs-bank)
6. [Contributing](#contributing)
7. [License](#license)
8. [Disclaimer](#disclaimer)

---

## Features

| # | Feature | Details |
|---|---------|---------|
| 1 | **AppSheet-friendly intake** | Define an AppSheet form that writes submissions into the `Form` sheet. |
| 2 | **Structured sheet layout** | A consistent tab and column layout to support expense tracking workflows. |
| 3 | **Bank reconciliation helper** | Includes a reconciliation script to compare manual entries against bank statements. |

---

## Sheet Structure

The spreadsheet is expected to contain the following tabs:

| Tab name | Role |
|----------|------|
| `Form` | **Input sheet.** Receives raw AppSheet submissions for downstream processing or manual workflows. |
| `List` | **Output sheet.** Normalised, itemised transactions populated by your workflow or automation. |
| `Data` | *(Optional)* Static reference data — categories, accounts, and people used in the form dropdowns. |
| `Balance` | *(Optional)* Summary formulas and per-person or per-account balance calculations. |
| `LUX` | *(Custom)* User-specific tracking or dashboard view. |
| `transactions` | *(Custom)* Advanced or archival transaction log. |

### AppSheet input sheet (`Form`) — recommended columns

Downstream processing typically resolves columns **by header name**, so order does not matter. The following are recommended column headers in row 1:

| Column header | Format / notes |
|---------------|----------------|
| `Timestamp` | Set automatically by AppSheet (`YYYY-MM-DD HH:mm:ss`). Fallback for Date and Time fields. |
| `Date` | Date of the expense (`dd/MM/yyyy`). Leave blank to fall back to Timestamp date. |
| `Time` | Time of the expense (`HH:mm`). Also accepts shorthand like `530p`, `5:30pm`, or `5p`, plus named periods: `morning`→`09:00`, `afternoon`→`12:00`, `evening`→`17:00`, `night` (or `knight`)→`21:00`. Leave blank to fall back to Timestamp time. |
| `Title` | Short description of the expense or transfer. |
| `Amount` | Positive number. Apply sign in your processing workflow if needed. |
| `Transaction Type` | Must contain the word `Transaction` or `Transfer` (case-insensitive). |
| `Processed` | Optional status flag for your workflow (for example, `YES` when processed). |
| `Account` | Account used for the expense *(Transaction rows only)*. |
| `Category` | Expense category *(Transaction rows only)*. |
| `Type` | `IN` or `OUT` *(Transaction rows only)*. |
| `Transaction Notes` | Free-text notes *(Transaction rows only, optional)*. |
| `Split Person` | Comma-, semicolon-, newline-, or dot-separated list of people sharing the expense. Use `me` for yourself. Defaults to `me` if blank. |
| `Out Account` | Source account *(Transfer rows only)*. |
| `In Account` | Destination account *(Transfer rows only)*. |
| `Transfer Person` | Person associated with the transfer *(Transfer rows only, optional)*. Defaults to `me`. |

> **Tip:** At least one of `Split Person` or `Transfer Person` must be present as a column; both can coexist.

## AppSheet Input Types & Fill Methods

This section explains what data each field accepts, and the different ways you can fill the AppSheet form so values correctly replicate into the `Form` sheet.

### Field-by-field accepted data/input

| Field | Data type | Accepted input | Required | Applies to |
|---|---|---|---|---|
| `Timestamp` | DateTime | Auto-filled by AppSheet on submit | Yes (sheet header required) | All rows |
| `Date` | Date | Any valid date value parseable by Google Sheets/Apps Script | Recommended | Transaction + Transfer |
| `Time` | Time/Text | `HH:mm` preferred; also supports `530p`, `5:30pm`, `5p`, and named periods (`morning`, `afternoon`, `evening`, `night`/`knight`) which normalize to `09:00`, `12:00`, `17:00`, `21:00` | Optional | Transaction + Transfer |
| `Title` | Text | Any short description | Recommended | Transaction + Transfer |
| `Amount` | Number | Positive numeric value (`100`, `100.50`) | Yes | Transaction + Transfer |
| `Transaction Type` | Text/Choice | Must include `Transaction` or `Transfer` (case-insensitive) | Yes | All rows |
| `Processed` | Text flag | Optional status flag you can set/update in your workflow (for example `YES`) | Optional | All rows |
| `Account` | Text/Choice | Account name | Yes for transaction rows | Transaction |
| `Category` | Text/Choice | Category name | Yes for transaction rows | Transaction |
| `Type` | Text/Choice | `IN` or `OUT` | Yes for transaction rows | Transaction |
| `Transaction Notes` | Text | Free text notes | Optional | Transaction |
| `Split Person` | Text list | One or many people separated by comma, semicolon, new line, or dot (`person1.person2`). **Note:** dot is always treated as a separator, so names containing `.` are split into multiple people; use names without dots (for example `Dr Smith` or `JK Rowling`). | Optional (defaults to `me`) | Transaction |
| `Out Account` | Text/Choice | Source account name | Yes for transfer rows | Transfer |
| `In Account` | Text/Choice | Destination account name | Yes for transfer rows | Transfer |
| `Transfer Person` | Text/Choice | Person name; defaults to `me` if blank | Optional | Transfer |

### How many ways can you fill the AppSheet form?

You can fill the AppSheet form in **3 practical ways**:

1. **Single-person transaction**  
   Use `Transaction Type = Transaction`, set `Amount`, `Account`, `Category`, `Type`, and one person (or leave blank to default to `me`).

2. **Split transaction (multi-person)**  
   Use `Transaction Type = Transaction`, and in `Split Person` provide multiple names using any supported separator:
   - comma: `me, Alice, Bob`
   - semicolon: `me; Alice; Bob`
   - new lines: one person per line  
   - dot: `me.Alice.Bob`  
   Split the total amount equally across the listed people in your processing workflow.

3. **Account transfer**  
   Use `Transaction Type = Transfer`, and provide `Out Account` + `In Account` (+ optional `Transfer Person`).

### How this replicates in the `Form` sheet

- Every AppSheet submission creates one raw row in the `Form` tab.
- Use your own workflow (Apps Script, formulas, or external tools) to transform `Form` rows into `List`.
- If you track processing status, update the `Processed` column as part of that workflow.

### `List` sheet — suggested columns

If you normalize `Form` entries into `List`, the following column order is commonly used. The reconciliation helper only requires Date (column A), Account (column C), and Amount (column E).

| # | Column | Description |
|---|--------|-------------|
| 1 | Date | Transaction date (`dd/MM/yyyy`) |
| 2 | Time | Transaction time (`HH:mm`) |
| 3 | Account | Account involved |
| 4 | Title | Description carried over from the form |
| 5 | Amount | Signed amount (negative = OUT, positive = IN). For splits, this is the per-person share. |
| 6 | Person | Name of the person for this row (`me` = yourself) |
| 7 | Category | Expense category (or `Transfer` for transfers) |
| 8 | Notes | Transaction notes, with an auto-appended split suffix when applicable |
| 9 | Status | `Settled`, `Pending`, or `Completed` |
| 10 | Direction | `IN` or `OUT` |
| 11 | Settlement Date | Populated when Status is `Settled` or `Completed`; blank otherwise |

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
- **Apps Script permissions** granted on first run if you run the reconciliation helper script.

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
