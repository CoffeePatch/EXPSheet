# Bank Reconciliation Engine

This folder contains the automated Bank → Ledger reconciliation system for comparing the manual ledger (`List`) against bank statement exports.

## Core Purpose

The bank statement is the **SOURCE OF TRUTH**. The reconciliation engine answers:

> *"For every relevant bank transaction, is its value represented in the target ledger account?"*

- **Direction**: `Bank → Ledger`
- **Reconciliation Scope**: Scoped to the configured target account (e.g. `9682`).
- **Daily Totals**: Used strictly as a **secondary checksum / sanity check**, not as the primary matching algorithm.

## Architecture & Features

The engine is built on a decoupled 7-layer architecture:
1. **Configuration**: Centralized `RECON_CONFIG` with flexible header candidates and exclusion rules.
2. **Schema & Header Resolution**: Dynamic header resolution for both `List` and bank exports. Eliminates hardcoded column positions.
3. **Data Normalization**: Deterministic calendar date parsing (`YYYY-MM-DD`) without ambiguous native Date guessing; currency cents rounding. Preserves raw source values for diagnostics.
4. **Bank Classification**: Categorizes bank rows into `IN_SCOPE`, `EXCLUDED` (e.g. internal sweeps containing `"SWEEP"`), or `UNRESOLVED` (malformed dates/amounts).
5. **Reconciliation Engine**:
   - **Controlled Amount Tolerance**: Normal 1:1 matching matches on exact date, direction, and amount difference $\le$ `AMOUNT_TOLERANCE` (default `1.00`). Exact matches are prioritized. Opposite directions and fuzzy dates are strictly rejected.
   - **Repeated Identical / Near-Equal Transactions**: Stateful consumption tracking ensures repeated transactions are cleanly paired 1:1 without false duplicate warnings.
   - **Explicit Split Matching**: Recognizes upstream split markers produced by the sweeper (`[Split: X / N = Y]` or `(Split: X/N)`) and shared base IDs. Enforces exact group identity and financial sum equality.
   - **Zero Combinatorial Guessing**: Never guesses arbitrary subsets of ledger rows.
   - **Status Model**: `MATCH`, `MATCH — SPLIT`, `MISSING`, `REVIEW`, `UNRESOLVED`, `EXCLUDED`, `LEDGER EXTRA`.
6. **Diagnostics & Result Model**: Pure JavaScript data model completely decoupled from Google Sheets I/O.
7. **Output & Presentation**:
   - **Unified Filterable Working Table (`Reconciliation_Log`)**: Replaces fragmented sections with ONE flat, filterable/sortable table with exactly 11 columns in this order:
     `ID | Date | Account | Title | Amount | P | D | C | Notes | Source | Category of Problem`
     - Columns `P`, `D`, `C` are blank (`""`).
     - Amount is strictly signed numeric (outflows negative e.g. `-255.00`, inflows positive e.g. `255.00`, never parentheses).
     - Bank missing rows have Category `"Missing in Ledger"`, Source `"Bank"`, and generated 8-hex IDs (`79e0f922`).
     - Ledger extra rows have Category `"Missing in Bank"`, Source `"Ledger"`, and preserve the original ledger `ID` (never row numbers).
     - Reviews have Category `"Review"`, Source `"Bank"`.
     - Data errors have Category `"Unresolved Data"`.
     - Successful `MATCH` and `EXCLUDED` transactions do not clutter this table.
   - **Forensic View (`Reconciliation_Details`)**: Full transaction audit log including successful `MATCH` and `MATCH — SPLIT` rows.
   - **Daily Checksum View (`Reconciliation_Summary`)**: Daily totals, variances, and checksum state (`OK`, `WARNING`, `DISCREPANCY`). Legitimate tolerance matches do not produce false checksum discrepancies.

## Required Sheets

| Sheet | Purpose |
|---|---|
| `List` | The manual ledger. Requires `Date`, `Account`, and `Amount` headers (supports `ID`, `Title`, `Notes`, `Type`, `Category`). |
| `Bank_Raw` | Raw bank statement export tab. Requires `Date`, `Debit`, and `Credit` headers (`Description`/`Details` optional). |
| `Reconciliation_Log` | Primary exception-oriented report created/rebuilt by the script (ONE flat 11-column filterable table). |
| `Reconciliation_Details` | (Optional/Forensic) Contains all transactions and matches. |
| `Reconciliation_Summary` | (Optional/Daily Checksum) Contains daily inflow/outflow checksums and status. |

## Configuration (`RECON_CONFIG`)

Update constants in [Reconciliation.gs](Reconciliation.gs):

- `RECON_TARGET_ACCOUNT`: Target ledger account name (e.g. `"9682"`).
- `AMOUNT_TOLERANCE`: Allowable difference for 1:1 match on same date & direction (default `1.00`).
- `BANK_EXCLUSION_PATTERNS`: Array of patterns for internal sweep exclusion (e.g. `["SWEEP"]`). Normalized across case, whitespace, and newlines.
- `SHEET_LIST`, `SHEET_BANK_RAW`, `SHEET_RECON_LOG`: Sheet/tab names.
- `BANK_SPREADSHEET_ID`: (Optional) External bank workbook ID. Leave empty if bank sheet is in the active workbook.
- `BANK_SHEET_NAME`: (Optional) Tab name in external workbook.
- `RECON_DATE_ORDER`: `"DMY"` (default) or `"MDY"`.
- `RECON_START_DATE` / `RECON_END_DATE`: Date bounds.
- `RECON_USE_OVERLAP_START_DATE`: When `true`, ignores dates before the earliest manual transaction in range.
- `DRY_RUN`: When `true`, runs reconciliation calculations and logs diagnostics without mutating any spreadsheet sheets.
- `OUTPUT_MODE`:
  - `"EXCEPTIONS_ONLY"` (default): Rebuilds `Reconciliation_Log` with the unified 11-column table and summary cards.
  - `"ALL_TRANSACTIONS"`: Writes `Reconciliation_Log` plus full forensic details to `Reconciliation_Details`.
  - `"TWO_SHEETS"`: Writes `Reconciliation_Summary` (daily checksum) and `Reconciliation_Details` (forensic log).

## Running Reconciliation

1. Paste your bank statement into `Bank_Raw` (or configure `BANK_SPREADSHEET_ID`).
2. Verify `RECON_TARGET_ACCOUNT` in `RECON_CONFIG`.
3. Run `reconcileBankStatement()` in Apps Script.

## Verification & Automated Tests

A comprehensive test suite covering all 35 business and presentation scenarios is available in [ReconciliationTests.gs](ReconciliationTests.gs).

Run `runReconciliationTests()` in the Apps Script editor to verify:
1. Perfect 1:1 match
2. Missing bank transaction
3. Repeated identical bank transactions
4. Duplicate ledger rows with 1 bank transaction (Rule: Match 1, 2 become LEDGER EXTRA, 0 REVIEW)
5. Explicit split transaction
6. Ambiguous / incomplete split
7. Ledger extra
8. Excluded sweep transaction
9. Invalid date (UNRESOLVED)
10. Invalid amount (UNRESOLVED)
11. Wrong target account
12. Equal daily totals but incorrect transaction representation (Daily Checksum WARNING)
13. Both IN and OUT transactions
14. Formatting, whitespace, and case variations
15. Independent split groups isolation (MUST NOT MERGE)
16. Split marker embedded inside extra notes text
17. Source immutability check
18. Action row for missing OUT transaction (signed negative amount e.g. -40)
19. Action row for missing IN transaction (signed positive amount e.g. +500)
20. Action row decimal precision (-127.50 preserved as numeric value)
21. Multiple missing transactions get unique generated hex IDs
22. Non-missing transactions generate no rows in missing action helper
23. Source data immutability during row generation
24. Unified 11-column table structure & signed formatting verification (`#,##0.00;-#,##0.00;0.00`, never parentheses)
25. Sweep exclusion real-world narrative (`"SWEEP TRF CREDT 0045279680696 OF Mr. TEJESWARA RAO ..."`)
26. Sweep exclusion case-insensitivity and multiline normalization
27. Non-sweep transactions are not excluded (`"UBER TRIP MUMBAI"`)
28. Excluded sweep creates zero checksum warnings or missing items
29. Match within amount tolerance (`diff <= 1.00`, records diagnostic difference note, checksum `OK`)
30. Amount difference exceeding tolerance does not match (`diff > 1.00`)
31. Opposite direction does not match under tolerance
32. Exact match preferred over tolerance match on same date
33. Repeated near-equal transactions consume 1:1 without double counting
34. Unified exception rows produce exactly 11 columns for all exception categories
35. Ledger extras in unified table preserve original ledger ID (never row numbers)