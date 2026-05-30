# EXPSheet

EXPSheet is split into three workflow-specific areas under `scripts`:

| Path | Purpose | Main file |
|---|---|---|
| [scripts/form-submission/](scripts/form-submission/) | Google Forms intake workflow, split handling, and form-response processing. | `FORM.gs` belongs here if you use the form-driven flow. |
| [scripts/appsheet/](scripts/appsheet/) | AppSheet-captured transaction sweeper. | [IndTran.gs](scripts/appsheet/IndTran.gs) |
| [scripts/reconciliation/](scripts/reconciliation/) | Manual bank reconciliation against a pasted export. | [Reconciliation.gs](scripts/reconciliation/Reconciliation.gs) |

## Start Here

1. If you are using the Google Forms flow, open [scripts/form-submission/README.md](scripts/form-submission/README.md). That folder is the intended home for the form-backend script.
2. If you are using AppSheet, open [scripts/appsheet/README.md](scripts/appsheet/README.md).
3. If you are reconciling your ledger against a bank export, open [scripts/reconciliation/README.md](scripts/reconciliation/README.md).

## Shared Ledger

The three workflows all revolve around the same spreadsheet ledger:

- `List` is the shared output / ledger sheet.
- If you rename a sheet tab or header, update the matching constants in the relevant script.
- Keep the workflow-specific docs inside their own folder so the setup stays focused and easier to maintain.

## Repository Map

| Folder | What to expect |
|---|---|
| `scripts/form-submission/` | Detailed notes for the Google Forms intake workflow and the expected `FORM.gs` placement. |
| `scripts/appsheet/` | The AppSheet transaction sweeper plus its setup notes. |
| `scripts/reconciliation/` | The reconciliation script, required config values, and bank-export assumptions. |
