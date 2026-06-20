# Form Submission Workflow (Alternative / Legacy)

This folder is the home of the Google Forms intake workflow. **Note: AppSheet is now the primary and recommended intake workflow for EXPSheet.** If you are still using the form-driven version of EXPSheet, keep `FORM.gs` here so the documentation and the script stay together.

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

## Required Form Headers

The backend resolves columns by header name, so the order can change, but the names must stay consistent.

| Column | Applies to | Notes |
|---|---|---|
| `Timestamp` | All rows | Auto-filled by Google Forms. |
| `Date` | All rows | Date of the expense or transfer. |
| `Time` | All rows | Accepts `HH:mm` and shorthand values. |
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

The backend appends one or more rows per processed form submission, matching the updated core schema:

| Column | Purpose |
|---|---|
| Date | Normalized transaction date. |
| Time | Normalized transaction time. |
| Account | Account involved in the row. |
| Title | Entry title from the form. |
| Amount | Signed amount, with split rows divided equally. |
| Debt Entity | Person attached to the row. |
| Category | Category or `Transfer` for transfer rows. |
| Notes | Original notes plus any split or transfer suffix. |
| Status | `Settled`, `Pending`, or `Completed`. |
| Type | `IN` or `OUT`. |
| Settlement Date | Filled when the row is settled or completed. |

## Setup And Triggers

1. Put `FORM.gs` in this folder if you are using the Google Forms workflow.
2. Update the script constants to match your sheet and header names.
3. Use one installable time-driven trigger for the form processor instead of mixing trigger types.

## Notes

- If you are transitioning to the modern, automated workflow, the AppSheet flow is documented in [../appsheet/README.md](../appsheet/README.md).
