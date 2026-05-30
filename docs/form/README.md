# AppSheet Intake Setup

This folder contains the intake-side Apps Script and the setup notes for the new **AppSheet-first** workflow.

Google Forms is no longer the primary entry point. AppSheet writes rows into the `Form` sheet, and `processFormResponses()` turns those rows into normalized ledger entries in `List`.

---

## What This Workflow Does

- AppSheet provides the no-code user interface.
- Users submit expenses and transfers through an AppSheet form.
- The submission lands in the `Form` sheet.
- `FORM.gs` reads new rows, validates them, and appends one or more rows to `List`.
- The script marks each successful source row as `YES` in `Processed`.

---

## Required Sheets

Your spreadsheet should contain at least these tabs:

| Sheet | Purpose |
|---|---|
| `Form` | AppSheet intake table. New submissions are written here. |
| `List` | Normalized ledger output written by the script. |

Optional tabs from the wider project can still exist, but the intake script only requires `Form` and `List`.

---

## Required Columns in `Form`

The column names must match the script headers exactly:

| Column | Recommended AppSheet type | Notes |
|---|---|---|
| `Timestamp` | DateTime | AppSheet can set this automatically with `NOW()`. |
| `Date` | Date | Transaction date. |
| `Time` | Time or Text | Optional; script accepts time text like `5:30pm` or `morning`. |
| `Title` | Text | Short description. |
| `Amount` | Number | Positive value. |
| `Transaction Type` | Enum | Use `Transaction` or `Transfer`. |
| `Processed` | Text | Hide this from users; leave blank by default. |
| `Account` | Text or Ref | Required for transaction rows. |
| `Category` | Text or Ref | Required for transaction rows. |
| `Type` | Enum | `IN` or `OUT` for transaction rows. |
| `Transaction Notes` | LongText | Optional notes for transactions. |
| `Split Person` | Text | Optional list of people for split transactions. |
| `Out Account` | Text or Ref | Required for transfer rows. |
| `In Account` | Text or Ref | Required for transfer rows. |
| `Transfer Person` | Text or Ref | Optional person for transfer rows. |

The script defaults `Split Person` and `Transfer Person` to `me` when blank.

---

## Step-by-Step AppSheet Setup

1. Open your spreadsheet in Google Sheets.
2. Make sure the `Form` and `List` tabs already exist.
3. Go to **Extensions → AppSheet → Create an app** or open AppSheet directly and start from the existing spreadsheet.
4. Add the spreadsheet as the app data source and select the `Form` table.
5. Keep the `Form` table name unchanged if you want to use the script without editing `CONFIG.SHEET_FORM`.
6. In **Data → Columns**, set the column types listed above.
7. Hide `Processed` from the user-facing form and keep it read-only.
8. Set `Timestamp` to auto-fill with `NOW()`.
9. Optionally set `Date` to `TODAY()` and `Split Person` to `me` as default values.
10. Configure `Transaction Type` as an enum with two values: `Transaction` and `Transfer`.
11. Configure `Type` as an enum with `IN` and `OUT`.
12. If you have master lists for accounts, categories, or people, connect them as `Ref` tables or enum lists.
13. Create an AppSheet form view for the `Form` table and publish it.
14. Submit a test row from the app and confirm that a new row appears in the `Form` sheet.

---

## Apps Script Setup

1. Open **Extensions → Apps Script** from the spreadsheet.
2. Add the script from [FORM.gs](FORM.gs).
3. Review the `CONFIG` block and update sheet names or header names only if your spreadsheet differs from the default layout.
4. Save the project.
5. Create one installable time-driven trigger for `processFormResponses()`.
6. Avoid using a Google Forms submit trigger. The AppSheet app writes directly to the sheet, so the time-driven trigger is the stable path.

Recommended cadence: every 5 or 10 minutes, depending on how quickly you need rows processed.

---

## Validation Checklist

Use this list to verify the setup end to end:

1. AppSheet submits a new row into `Form`.
2. `Processed` remains blank until the script runs.
3. `processFormResponses()` appends normalized output rows to `List`.
4. The source row is marked `YES` in `Processed` after a successful run.
5. Split transactions create multiple rows when multiple people are entered.
6. Transfer rows create the paired `OUT` and `IN` ledger entries.

---

## Customization Notes

- If your tab names differ, update `CONFIG.SHEET_FORM` and `CONFIG.SHEET_LIST`.
- If your AppSheet column labels differ, update the corresponding header constants in `CONFIG`.
- The script accepts flexible date and time input, but it works best when AppSheet stores clean Date and Time values.
