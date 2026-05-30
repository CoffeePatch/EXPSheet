# EXPSheet

EXPSheet is a Google Sheets and Apps Script workflow for expense intake, split transactions, transfer logging, and manual bank reconciliation.

The repository includes the script files at the root and matching setup notes under `docs/`:

- [FORM.gs](FORM.gs)
- [Reconciliation.gs](Reconciliation.gs)
- [Form intake setup](docs/form/README.md)
- [Bank reconciliation setup](docs/reconciliation/README.md)

## How It Fits Together

1. AppSheet or Google Forms writes new submissions into the `Form` sheet.
2. `processFormResponses()` reads unprocessed rows and appends normalized entries to `List`.
3. `reconcileBankStatement()` compares `List` against a pasted bank export in `Bank_Raw`.

## Quick Start

1. Create or confirm the `Form`, `List`, `Bank_Raw`, and `Reconciliation_Log` tabs in your spreadsheet.
2. Import the Apps Script files from this repository into your Apps Script project.
3. Follow the setup guide in [docs/form/README.md](docs/form/README.md) for intake.
4. Follow [docs/reconciliation/README.md](docs/reconciliation/README.md) for bank reconciliation.
5. Set one installable time-driven trigger for `processFormResponses()`.
6. Run `reconcileBankStatement()` manually whenever you need a bank comparison.

## Notes

- The intake workflow can be driven by AppSheet or by a Google Form, depending on how you configure the `Form` sheet.
- The reconciliation helper supports explicit start and end dates, or it can fall back to the first and last bank rows when only one bound is provided.
- If you rename a sheet or column header, update the matching constant in the relevant script.

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
- Dot (`.`) is always treated as a person separator in `Split Person`.
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
