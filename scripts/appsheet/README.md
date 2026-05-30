# AppSheet Workflow

This folder documents the AppSheet-based intake path and the backend sweeper that normalizes rows after users submit them.

## Main Script

- [IndTran.gs](IndTran.gs) contains `processTransaction()`.

## What The Script Does

`processTransaction()` reads the `List` sheet directly, normalizes rows, and appends any extra rows needed for transfers or split transactions.

- It scans the last portion of the sheet to avoid missed rows near the bottom.
- It normalizes headers before looking them up, so spacing and case are tolerated.
- It rewrites internal transfers and investment rows so the ledger stays consistent.
- It splits comma-separated people into multiple rows and divides the amount equally.

## Expected List Columns

The script reads these header names after normalization:

| Column | Purpose |
|---|---|
| `Account` | Ledger account or transfer target. |
| `Amount` | Transaction amount. |
| `Category` | Used to detect transfer and investment handling. |
| `Person` | Person name or comma-separated split list. |
| `Notes` | Optional notes carried through to cloned rows. |
| `Status` | Updated for split rows that should remain pending. |
| `Settlement Date` | Cleared for rows that should not inherit a settlement date. |
| `Type` | Set to `IN` for appended transfer legs. |

## Processing Rules

1. If the row is empty, it is ignored.
2. If `Category` is `internal transfers` or `investments` and `Person` is not `me`, the script converts the original row to `Me` and appends an `IN` leg.
3. If `Person` contains commas, the script treats it as a split transaction, divides the amount equally, updates the original row, and appends one row per additional person.
4. Appended split rows are marked `Pending`, and their settlement date is cleared.

## Notes

- Keep the `List` headers aligned with what the script expects, even though lookup is case-insensitive and trims spaces.
- This workflow is intentionally separate from the Google Forms flow documented in [../form-submission/README.md](../form-submission/README.md).
- If you change the ledger layout, update the script and this README together.
