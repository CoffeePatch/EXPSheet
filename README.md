# EXPSheet

EXPSheet is a Google Sheets and Apps Script workflow for AppSheet-based expense intake, split transactions, transfer logging, and manual bank reconciliation.

The project is now organized by workflow:

- [Form intake setup](docs/form/README.md)
- [Bank reconciliation setup](docs/reconciliation/README.md)

---

## How The Pieces Fit Together

1. AppSheet writes new submissions into the `Form` sheet.
2. `processFormResponses()` reads those rows and appends normalized entries to `List`.
3. `reconcileBankStatement()` compares the `List` sheet against a pasted bank export in `Bank_Raw`.

---

## Repository Layout

| Path | Purpose |
|---|---|
| [docs/form/FORM.gs](docs/form/FORM.gs) | Intake processor used by Apps Script. |
| [docs/form/README.md](docs/form/README.md) | AppSheet setup and trigger instructions. |
| [docs/reconciliation/Reconciliation.gs](docs/reconciliation/Reconciliation.gs) | Bank reconciliation script with optional date range support. |
| [docs/reconciliation/README.md](docs/reconciliation/README.md) | Manual reconciliation setup and usage notes. |

---

## Quick Start

1. Create or confirm the `Form`, `List`, `Bank_Raw`, and `Reconciliation_Log` tabs in your spreadsheet.
2. Import the Apps Script files from the `docs/` folders into your Apps Script project.
3. Follow the detailed AppSheet setup in [docs/form/README.md](docs/form/README.md).
4. Set one installable time-driven trigger for `processFormResponses()`.
5. Run `reconcileBankStatement()` manually whenever you need a bank comparison.

---

## Notes

- The intake workflow is configured manually in AppSheet; Google Forms is no longer the primary entry point.
- The reconciliation helper now supports explicit start and end dates, or can fall back to the first and last bank rows when only one bound is provided.
- If you rename a sheet or a column header, update the matching constant in the relevant script.
