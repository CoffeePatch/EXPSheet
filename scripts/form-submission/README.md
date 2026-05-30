# Form Submission Workflow

This folder is the home of the Google Forms intake workflow. If you are using the form-driven version of EXPSheet, keep `FORM.gs` here so the documentation and the script stay together.

## What This Workflow Does

The form backend reads raw form responses, normalizes them into ledger rows, and writes the result into `List`.

- Transaction rows can be split across multiple people.
- Transfer rows create both the source and destination legs.
- Each form row is marked as processed after a successful run so reprocessing stays safe.

## Expected Sheets

| Sheet | Role |
|---|---|
| `Form` | Raw Google Form responses. The backend reads from this sheet. |
| `List` | Normalized output ledger written by the script. |
| `Data` | Optional lookup data for dropdowns. |
| `Balance` | Optional summary sheet. |
| `LUX` | Optional custom dashboard or personal tracking tab. |
| `transactions` | Optional archive or advanced transaction log. |

## Required Form Headers

The backend resolves columns by header name, so the order can change, but the names must stay consistent.

| Column | Applies to | Notes |
|---|---|---|
| `Timestamp` | All rows | Auto-filled by Google Forms. Used as a fallback when `Date` or `Time` is blank. |
| `Date` | All rows | Date of the expense or transfer. |
| `Time` | All rows | Accepts `HH:mm` and shorthand values such as `530p`, `5:30pm`, or `5p`. |
| `Title` | All rows | Short description of the entry. |
| `Amount` | All rows | Positive amount. The script applies the sign based on direction. |
| `Transaction Type` | All rows | Must include `Transaction` or `Transfer`. |
| `Processed` | All rows | Leave blank. The script writes `YES` after success. |
| `Account` | Transaction rows | Ledger account for a transaction entry. |
| `Category` | Transaction rows | Category name. |
| `Type` | Transaction rows | `IN` or `OUT`. |
| `Transaction Notes` | Transaction rows | Optional free text. |
| `Split Person` | Transaction rows | One or more people separated by comma, semicolon, new line, or dot. Defaults to `me`. |
| `Out Account` | Transfer rows | Source account. |
| `In Account` | Transfer rows | Destination account. |
| `Transfer Person` | Transfer rows | Optional person field, defaults to `me`. |

## Output Columns In `List`

The backend appends one or more rows per processed form submission.

| # | Column | Purpose |
|---|---|---|
| 1 | Date | Normalized transaction date. |
| 2 | Time | Normalized transaction time. |
| 3 | Account | Account involved in the row. |
| 4 | Title | Entry title from the form. |
| 5 | Amount | Signed amount, with split rows divided equally. |
| 6 | Person | Person attached to the row. |
| 7 | Category | Category or `Transfer` for transfer rows. |
| 8 | Notes | Original notes plus any split or transfer suffix. |
| 9 | Status | `Settled`, `Pending`, or `Completed`. |
| 10 | Direction | `IN` or `OUT`. |
| 11 | Settlement Date | Filled when the row is settled or completed. |

## Practical Entry Patterns

1. Single-person transaction: set `Transaction Type` to `Transaction`, fill in the transaction fields, and leave `Split Person` blank to default to `me`.
2. Split transaction: set `Transaction Type` to `Transaction`, then enter multiple people in `Split Person` using any supported separator. The backend divides the amount equally and writes one ledger row per person.
3. Transfer: set `Transaction Type` to `Transfer`, provide `Out Account` and `In Account`, and optionally set `Transfer Person`.

## Setup And Triggers

1. Put `FORM.gs` in this folder if you are using the Google Forms workflow.
2. Update the script constants to match your sheet and header names.
3. Use one installable time-driven trigger for the form processor instead of mixing trigger types.
4. If a trigger was down for a while, the next successful run backfills any rows that are still unprocessed.

## Notes

- Dot is always treated as a separator in `Split Person`, so avoid dots in names.
- Rows that fail validation stay unprocessed and will retry on the next run.
- If you are not using the form-driven workflow anymore, the AppSheet flow is documented in [../appsheet/README.md](../appsheet/README.md).
