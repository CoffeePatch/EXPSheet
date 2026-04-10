# EXPSheet

A professional, extensible **Google Apps Script** project for automated expense tracking, transfer logging, and smart multi-person splits — all driven by a simple Google Form and analysed in Google Sheets.

---

## Table of Contents

1. [Features](#features)
2. [Sheet Structure](#sheet-structure)
3. [Setup Requirements](#setup-requirements)
4. [Installation](#installation)
5. [Setting Up Triggers](#setting-up-triggers)
6. [Form Input Types & Fill Methods](#form-input-types--fill-methods)
7. [Usage Notes & Best Practices](#usage-notes--best-practices)
8. [Contributing](#contributing)
9. [License](#license)
10. [Disclaimer](#disclaimer)

---

## Features

| # | Feature | Details |
|---|---------|---------|
| 1 | **Form-driven expense entry** | Submit expenses and transfers via a linked Google Form — no manual spreadsheet editing needed. |
| 2 | **Automatic transaction splitting** | Group expenses are automatically divided equally across named people. |
| 3 | **Transfer support** | Logs both the debit side (Out Account) and the credit side (In Account) from a single form entry. |
| 4 | **Idempotent / duplicate-safe processing** | Each row is stamped `YES` in the `Processed` column after it is handled; re-runs are completely safe. |
| 5 | **Batch / catch-up processing** | Unprocessed rows from periods of downtime are automatically backfilled on the next successful run. |
| 6 | **Concurrent-run protection** | `LockService` prevents two trigger executions from racing and writing duplicate rows. |
| 7 | **Header-safe dynamic mapping** | Column positions are resolved by name at runtime — reordering columns does not break the script. |
| 8 | **Robust error logging** | Row-level failures are logged and skipped; successful rows continue to be processed normally. |
| 9 | **Direct time-driven scheduling** | Configure one trigger directly on `processFormResponses` and choose the time cadence that fits your workflow. |

---

## Sheet Structure

The spreadsheet is expected to contain the following tabs:

| Tab name | Role |
|----------|------|
| `Form` | **Input sheet.** Receives raw Google Form submissions. The script reads from this sheet. |
| `List` | **Output sheet.** Normalised, itemised transactions (one row per person/leg) written by the script. |
| `Data` | *(Optional)* Static reference data — categories, accounts, and people used in the form dropdowns. |
| `Balance` | *(Optional)* Summary formulas and per-person or per-account balance calculations. |
| `LUX` | *(Custom)* User-specific tracking or dashboard view. |
| `transactions` | *(Custom)* Advanced or archival transaction log. |

### `Form` sheet — required columns

The script resolves columns **by header name**, so order does not matter. All of the following must exist as column headers in row 1:

| Column header | Format / notes |
|---------------|----------------|
| `Timestamp` | Set automatically by Google Forms (`YYYY-MM-DD HH:mm:ss`). Fallback for Date and Time fields. |
| `Date` | Date of the expense (`dd/MM/yyyy`). Leave blank to fall back to Timestamp date. |
| `Time` | Time of the expense (`HH:mm`). Also accepts shorthand like `530p`, `5:30pm`, or `5p`, plus named periods: `morning`→`09:00`, `afternoon`→`12:00`, `evening`→`17:00`, `night` (or `knight`)→`21:00`. Leave blank to fall back to Timestamp time. |
| `Title` | Short description of the expense or transfer. |
| `Amount` | Positive number. Negative sign is applied automatically by the script based on direction. |
| `Transaction Type` | Must contain the word `Transaction` or `Transfer` (case-insensitive). |
| `Processed` | **Managed by the script.** Leave blank in new rows; the script writes `YES` here when done. |
| `Account` | Account used for the expense *(Transaction rows only)*. |
| `Category` | Expense category *(Transaction rows only)*. |
| `Type` | `IN` or `OUT` *(Transaction rows only)*. |
| `Transaction Notes` | Free-text notes *(Transaction rows only, optional)*. |
| `Split Person` | Comma-, semicolon-, newline-, or dot-separated list of people sharing the expense. Use `me` for yourself. Defaults to `me` if blank. |
| `Out Account` | Source account *(Transfer rows only)*. |
| `In Account` | Destination account *(Transfer rows only)*. |
| `Transfer Person` | Person associated with the transfer *(Transfer rows only, optional)*. Defaults to `me`. |

> **Tip:** At least one of `Split Person` or `Transfer Person` must be present as a column; both can coexist.

## Form Input Types & Fill Methods

This section explains what data each field accepts, and the different ways you can fill the form so values correctly replicate into the `Form` sheet.

### Field-by-field accepted data/input

| Field | Data type | Accepted input | Required | Applies to |
|---|---|---|---|---|
| `Timestamp` | DateTime | Auto-filled by Google Forms on submit | Yes (sheet header required) | All rows |
| `Date` | Date | Any valid date value parseable by Google Sheets/Apps Script | Recommended | Transaction + Transfer |
| `Time` | Time/Text | `HH:mm` preferred; also supports `530p`, `5:30pm`, `5p`, and named periods (`morning`, `afternoon`, `evening`, `night`/`knight`) which normalize to `09:00`, `12:00`, `17:00`, `21:00` | Optional | Transaction + Transfer |
| `Title` | Text | Any short description | Recommended | Transaction + Transfer |
| `Amount` | Number | Positive numeric value (`100`, `100.50`) | Yes | Transaction + Transfer |
| `Transaction Type` | Text/Choice | Must include `Transaction` or `Transfer` (case-insensitive) | Yes | All rows |
| `Processed` | Text flag | Leave blank; script writes `YES` after successful processing | Script-managed | All rows |
| `Account` | Text/Choice | Account name | Yes for transaction rows | Transaction |
| `Category` | Text/Choice | Category name | Yes for transaction rows | Transaction |
| `Type` | Text/Choice | `IN` or `OUT` | Yes for transaction rows | Transaction |
| `Transaction Notes` | Text | Free text notes | Optional | Transaction |
| `Split Person` | Text list | One or many people separated by comma, semicolon, new line, or dot (`person1.person2`). | Optional (defaults to `me`) | Transaction |
| `Out Account` | Text/Choice | Source account name | Yes for transfer rows | Transfer |
| `In Account` | Text/Choice | Destination account name | Yes for transfer rows | Transfer |
| `Transfer Person` | Text/Choice | Person name; defaults to `me` if blank | Optional | Transfer |

### How many ways can you fill the form?

You can fill the form in **3 practical ways**:

1. **Single-person transaction**  
   Use `Transaction Type = Transaction`, set `Amount`, `Account`, `Category`, `Type`, and one person (or leave blank to default to `me`).

2. **Split transaction (multi-person)**  
   Use `Transaction Type = Transaction`, and in `Split Person` provide multiple names using any supported separator:
   - comma: `me, Alice, Bob`
   - semicolon: `me; Alice; Bob`
   - new lines: one person per line  
   - dot: `me.Alice.Bob`  
   The script splits the total amount equally and writes one `List` row per person.

3. **Account transfer**  
   Use `Transaction Type = Transfer`, and provide `Out Account` + `In Account` (+ optional `Transfer Person`).  
   The script writes two `List` rows: one `OUT` and one `IN`.

### How this replicates in the `Form` sheet

- Every Google Form submission creates one raw row in the `Form` tab.
- `processFormResponses` reads unprocessed rows (`Processed` not `YES`).
- Valid rows are transformed and appended to `List`.
- Successfully handled `Form` rows are marked `YES` in `Processed`.

### `List` sheet — output columns

The script appends one or more rows per processed form entry. The columns written are (in order):

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

## Setup Requirements

- A **Google Account** with access to Google Sheets and Google Forms.
- Your own **Google Sheet** with the tabs and columns described above, *or* a copy of the EXPSheet template.
- **Apps Script permissions** granted on first run (the script accesses your spreadsheet and requires script-lock access).

---

## Installation

1. Open your Google Sheet.
2. Navigate to **Extensions → Apps Script**.
3. In the editor, create a new script file named `FORM.gs` (click the **+** next to *Files*).
4. Copy the entire contents of [`FORM.gs`](./FORM.gs) from this repository and paste it into the editor.
5. Review the `CONFIG` block at the top of the file:

   ```javascript
   const CONFIG = Object.freeze({
     SHEET_FORM: "Form",          // Name of your intake tab
     SHEET_LIST: "List",          // Name of your output tab
     COL_SPLIT_PERSON: "Split Person",     // Exact header for the split-persons column
     COL_TRANSFER_PERSON: "Transfer Person", // Exact header for the transfer-person column
     DATE_FORMAT: "dd/MM/yyyy",
     TIME_FORMAT: "HH:mm",
     ...
   });
   ```

   Update any value that differs from your actual sheet or column names.

6. **Save** (`Ctrl+S` / `Cmd+S`).

> **Important:** If you ever rename a sheet tab or a column header, you must update the corresponding value in `CONFIG` or the script will throw a clear error and stop safely.

---

## Setting Up Triggers

### Recommended trigger strategy

Use **one installable time-driven trigger** for `processFormResponses` and remove old `On form submit` triggers for the same function.

Why this is recommended:
- It is stable for both real-time and backlog processing.
- It automatically catches up rows that were missed during downtime.
- It avoids duplicate/overlapping trigger patterns.

### Clean old triggers first (important)

1. Open Apps Script → **Triggers** (clock icon).
2. Delete any old trigger that runs `processFormResponses` with **Event source = From spreadsheet** and **Event type = On form submit**.
3. Keep only one active trigger pattern for this project (time-driven).

### Create the correct trigger (UI method)

1. Apps Script → **Triggers** → **Add Trigger**
2. Choose function: `processFormResponses`
3. Event source: `Time-driven`
4. Type of time based trigger: choose the cadence you want (for example every minute, every 5 minutes, every 10 minutes, every 15 minutes, or every 30 minutes)
5. Save

> Use exactly one active time-driven trigger for `processFormResponses`.

### Manual run (debugging / one-off catch-up)

Select **`processFormResponses`** from the run-function dropdown and click **Run**. This processes all pending rows immediately.

### Quick troubleshooting

- **Error rate shows 100% in Triggers:** delete old/wrong trigger entries and recreate one clean time-driven trigger for `processFormResponses`.
- **No new rows processed:** verify the trigger is `Time-driven` and function name is exactly `processFormResponses`.
- **Script still uses old settings:** save the latest `FORM.gs`, then remove and re-add the trigger.

### How the processing loop works

```text
On your selected trigger cadence (or on manual run):
  ┌─ Acquire script lock (30 s timeout) ─────────────────────────────┐
  │  1. Read all rows from the "Form" sheet.                          │
  │  2. Skip rows where Processed = "YES".                            │
  │  3. For each unprocessed row:                                     │
  │       a. Validate required fields (amount, date, type, etc.).     │
  │       b. Build output row(s) — one per person for splits,         │
  │          two per transfer (debit + credit).                        │
  │       c. On error: log & skip row (it will retry next run).       │
  │  4. Append all valid output rows to "List" in one batch write.    │
  │  5. Mark all successfully processed rows as "YES" in one write.   │
  └─ Release lock ────────────────────────────────────────────────────┘
```

### Catch-up / backfill for historical rows

If the trigger was not running for a period (e.g., script errors, quota exceeded, or project disabled), all rows where `Processed` is still blank will be picked up automatically on the next successful run — **no manual intervention is required**. Simply fix any script error and let the trigger fire, or run `processFormResponses` manually.

---

## Usage Notes & Best Practices

### Permissions
- The script must be run by a user who has **Edit access** to all relevant Google Sheets.
- Triggers run as the **owner** of the Apps Script project. Ensure the project owner's account remains active and has the necessary permissions.

### Duplicate prevention
- The `Processed` column is the source of truth. Once a row is marked `YES`, it is never re-processed.
- The script uses `LockService` to serialise concurrent trigger executions, preventing race conditions during batch writes.

### Changing sheet or column names
1. Update the corresponding constant in the `CONFIG` block in `FORM.gs`.
2. Save and verify your existing trigger still points to `processFormResponses`.

### Error handling
- If a **required header is missing**, the script throws immediately and logs the problem — no partial writes occur.
- If an **individual row** has invalid data (bad amount, missing account, unknown transaction type), that row is logged and skipped; all other rows in the same run continue normally.
- Failed rows are **not** marked as processed, so they will be retried on the next run once the data is corrected.

### Split logic
- The `Split Person` field accepts a comma-, semicolon-, newline-, or dot-separated list (e.g., `me, Alice, Bob` or `me.Alice.Bob`).
- The total amount is divided equally; each person receives their own row in `List`.
- `me` rows with a non-credit/non-pending account are automatically marked `Settled`.

### Transfer logic
- A transfer creates **two rows** in `List`: one `OUT` from the source account and one `IN` to the destination account, both marked `Completed`.
- If `Transaction Notes` is filled for a transfer row, the note is appended to both transfer entries (for example: `Transfer Out - your note` and `Transfer In - your note`).

---

## Contributing

Contributions, bug reports, and feature requests are welcome!

1. **Open an issue** at [github.com/CoffeePatch/EXPSheet/issues](https://github.com/CoffeePatch/EXPSheet/issues) to discuss what you would like to change.
2. **Fork** the repository and create a branch for your changes.
3. Submit a **pull request** with a clear description of what you changed and why.

For quick customisations (different date formats, extra output columns, alternate trigger intervals), editing the `CONFIG` block in `FORM.gs` is usually sufficient without needing to fork.

---

## License

This project is open-source under the MIT License. A `LICENSE` file will be added to the repository root in a future update.

---

## Disclaimer

This project is community-maintained and is **not** an official Google product.  
Always keep a backup of your spreadsheet data before deploying script updates or making structural changes to your sheets.  
The authors accept no liability for data loss or inaccuracies resulting from the use of this software.
