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
4. **Bank Classification**: Categorizes bank rows into `IN_SCOPE`, `EXCLUDED` (e.g. internal sweeps such as `"SWEEP FROM"`), or `UNRESOLVED` (malformed dates/amounts).
5. **Reconciliation Engine**:
   - **Exact 1:1 Matching**: Matches unconsumed ledger rows on exact date, direction, and amount.
   - **Repeated Identical Transactions**: Stateful consumption tracking ensures repeated legitimate transactions are cleanly paired (Bank #1 → Ledger #1, Bank #2 → Ledger #2) without false duplicate warnings.
   - **Explicit Split Matching**: Recognizes upstream split markers produced by the sweeper (`[Split: X / N = Y]` or `(Split: X/N)`) and shared base IDs. Enforces exact group identity and financial sum equality.
   - **Zero Combinatorial Guessing**: Never guesses arbitrary subsets of ledger rows.
   - **Status Model**: `MATCH`, `MATCH — SPLIT`, `MISSING`, `REVIEW`, `UNRESOLVED`, `EXCLUDED`, `LEDGER EXTRA`.
6. **Diagnostics & Result Model**: Pure JavaScript data model completely decoupled from Google Sheets I/O.
7. **Output & Presentation**:
30:    - **Exception-Oriented Primary View (`Reconciliation_Log`)**: Focuses exclusively on actionable items (`MISSING`, `REVIEW`, `UNRESOLVED`) with side-by-side comparison. Displays KPI cards and an informational section for `EXCLUDED` and `LEDGER EXTRA` rows.
31:    - **Copy-Ready Ledger Entries**: Generates a clean 5-column table (`ID | Date | Account | Title | Amount`) directly below the actionable exceptions on `Reconciliation_Log` for all `MISSING` bank transactions. Allows users to directly copy rows into the `List` sheet with minimal manual editing. Preserves signed numeric precision (debits negative, credits positive) and generates 8-character unique hex IDs (`79e0f922`).
32:    - **Forensic View (`Reconciliation_Details`)**: Full transaction audit log including successful `MATCH` and `MATCH — SPLIT` rows.
33:    - **Daily Checksum View (`Reconciliation_Summary`)**: Daily totals, variances, and checksum state (`OK`, `WARNING`, `DISCREPANCY`).
34: 
35: ## Required Sheets
36: 
37: | Sheet | Purpose |
38: |---|---|
39: | `List` | The manual ledger. Requires `Date`, `Account`, and `Amount` headers (supports `ID`, `Title`, `Notes`, `Type`, `Category`). |
40: | `Bank_Raw` | Raw bank statement export tab. Requires `Date`, `Debit`, and `Credit` headers (`Description`/`Details` optional). |
41: | `Reconciliation_Log` | Primary exception-oriented report created/rebuilt by the script (includes Action Required and Copy-Ready Ledger Entries). |
42: | `Reconciliation_Details` | (Optional/Forensic) Contains all transactions and matches. |
43: | `Reconciliation_Summary` | (Optional/Daily Checksum) Contains daily inflow/outflow checksums and status. |
44: 
45: ## Configuration (`RECON_CONFIG`)
46: 
47: Update constants in [Reconciliation.gs](Reconciliation.gs):
48: 
49: - `RECON_TARGET_ACCOUNT`: Target ledger account name (e.g. `"9682"`).
50: - `SHEET_LIST`, `SHEET_BANK_RAW`, `SHEET_RECON_LOG`: Sheet/tab names.
51: - `BANK_SPREADSHEET_ID`: (Optional) External bank workbook ID. Leave empty if bank sheet is in the active workbook.
52: - `BANK_SHEET_NAME`: (Optional) Tab name in external workbook.
53: - `RECON_DATE_ORDER`: `"DMY"` (default) or `"MDY"`.
54: - `RECON_START_DATE` / `RECON_END_DATE`: Date bounds.
55: - `RECON_USE_OVERLAP_START_DATE`: When `true`, ignores dates before the earliest manual transaction in range.
56: - `DRY_RUN`: When `true`, runs reconciliation calculations and logs diagnostics without mutating any spreadsheet sheets.
57: - `OUTPUT_MODE`:
58:   - `"EXCEPTIONS_ONLY"` (default): Rebuilds `Reconciliation_Log` with actionable exceptions, copy-ready ledger rows, and summary cards.
59:   - `"ALL_TRANSACTIONS"`: Writes `Reconciliation_Log` plus full forensic details to `Reconciliation_Details`.
60:   - `"TWO_SHEETS"`: Writes `Reconciliation_Summary` (daily checksum) and `Reconciliation_Details` (forensic log).
61: - `BANK_EXCLUSION_KEYWORDS`: Keywords for audit exclusion (e.g. `["SWEEP FROM", "AUTOSWEEP"]`).
62: 
63: ## Running Reconciliation
64: 
65: 1. Paste your bank statement into `Bank_Raw` (or configure `BANK_SPREADSHEET_ID`).
66: 2. Verify `RECON_TARGET_ACCOUNT` in `RECON_CONFIG`.
67: 3. Run `reconcileBankStatement()` in Apps Script.
68: 
69: ## Verification & Automated Tests
70: 
71: A comprehensive test suite covering all 23 business and presentation scenarios is available in [ReconciliationTests.gs](ReconciliationTests.gs).
72: 
73: Run `runReconciliationTests()` in the Apps Script editor to verify:
74: 1. Perfect 1:1 match
75: 2. Missing bank transaction
76: 3. Repeated identical bank transactions
77: 4. Duplicate ledger rows with 1 bank transaction (Rule: Match 1, 2 become LEDGER EXTRA, 0 REVIEW)
78: 5. Explicit split transaction
79: 6. Ambiguous / incomplete split
80: 7. Ledger extra
81: 8. Excluded sweep transaction
82: 9. Invalid date (UNRESOLVED)
83: 10. Invalid amount (UNRESOLVED)
84: 11. Wrong target account
85: 12. Equal daily totals but incorrect transaction representation (Daily Checksum WARNING)
86: 13. Both IN and OUT transactions
87: 14. Formatting, whitespace, and case variations
88: 15. Independent split groups isolation (MUST NOT MERGE)
89: 16. Split marker embedded inside extra notes text
90: 17. Source immutability check
91: 18. Copy-ready row for missing OUT transaction (signed negative amount)
92: 19. Copy-ready row for missing IN transaction (signed positive amount)
93: 20. Copy-ready row decimal precision (-127.50 preserved as numeric value)
94: 21. Multiple missing transactions get unique generated hex IDs
95: 22. Non-missing transactions (MATCH, REVIEW, UNRESOLVED, EXCLUDED) generate no copy-ready rows
96: 23. Source data immutability during copy-ready generation
   - **Action Required Primary Table (`Reconciliation_Log`)**: Contains exactly 5 columns: `ID | Date | Account | Title | Amount`. Populated strictly for `MISSING` bank transactions. Allows users to edit Title directly in the sheet and manually copy rows into `List`. Signed numeric values are enforced (negative for outflows e.g. `-255`, positive for inflows e.g. `255`, never parentheses).
   - **Informational & Diagnostics Table (`Reconciliation_Log`)**: Displays `REVIEW`, `UNRESOLVED`, `EXCLUDED`, and `LEDGER EXTRA` rows in a dedicated section below Action Required.
   - **Forensic View (`Reconciliation_Details`)**: Full transaction audit log including successful `MATCH` and `MATCH — SPLIT` rows.
   - **Daily Checksum View (`Reconciliation_Summary`)**: Daily totals, variances, and checksum state (`OK`, `WARNING`, `DISCREPANCY`).

## Required Sheets

| Sheet | Purpose |
|---|---|
| `List` | The manual ledger. Requires `Date`, `Account`, and `Amount` headers (supports `ID`, `Title`, `Notes`, `Type`, `Category`). |
| `Bank_Raw` | Raw bank statement export tab. Requires `Date`, `Debit`, and `Credit` headers (`Description`/`Details` optional). |
| `Reconciliation_Log` | Primary exception-oriented report created/rebuilt by the script (Action Required & Informational items). |
| `Reconciliation_Details` | (Optional/Forensic) Contains all transactions and matches. |
| `Reconciliation_Summary` | (Optional/Daily Checksum) Contains daily inflow/outflow checksums and status. |

## Configuration (`RECON_CONFIG`)

Update constants in [Reconciliation.gs](Reconciliation.gs):

- `RECON_TARGET_ACCOUNT`: Target ledger account name (e.g. `"9682"`).
- `SHEET_LIST`, `SHEET_BANK_RAW`, `SHEET_RECON_LOG`: Sheet/tab names.
- `BANK_SPREADSHEET_ID`: (Optional) External bank workbook ID. Leave empty if bank sheet is in the active workbook.
- `BANK_SHEET_NAME`: (Optional) Tab name in external workbook.
- `RECON_DATE_ORDER`: `"DMY"` (default) or `"MDY"`.
- `RECON_START_DATE` / `RECON_END_DATE`: Date bounds.
- `RECON_USE_OVERLAP_START_DATE`: When `true`, ignores dates before the earliest manual transaction in range.
- `DRY_RUN`: When `true`, runs reconciliation calculations and logs diagnostics without mutating any spreadsheet sheets.
- `OUTPUT_MODE`:
  - `"EXCEPTIONS_ONLY"` (default): Rebuilds `Reconciliation_Log` with Action Required and Informational tables.
  - `"ALL_TRANSACTIONS"`: Writes `Reconciliation_Log` plus full forensic details to `Reconciliation_Details`.
  - `"TWO_SHEETS"`: Writes `Reconciliation_Summary` (daily checksum) and `Reconciliation_Details` (forensic log).
- `BANK_EXCLUSION_KEYWORDS`: Keywords for audit exclusion (e.g. `["SWEEP FROM", "AUTOSWEEP"]`).

## Running Reconciliation

1. Paste your bank statement into `Bank_Raw` (or configure `BANK_SPREADSHEET_ID`).
2. Verify `RECON_TARGET_ACCOUNT` in `RECON_CONFIG`.
3. Run `reconcileBankStatement()` in Apps Script.

## Verification & Automated Tests

A comprehensive test suite covering all 24 business and presentation scenarios is available in [ReconciliationTests.gs](ReconciliationTests.gs).

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
18. Action Required row for missing OUT transaction (signed negative amount e.g. -40)
19. Action Required row for missing IN transaction (signed positive amount e.g. +500)
20. Action Required row decimal precision (-127.50 preserved as numeric value)
21. Multiple missing transactions get unique generated hex IDs
22. Non-missing transactions (MATCH, REVIEW, UNRESOLVED, EXCLUDED) generate no rows in Action Required
23. Source data immutability during Action Required row generation
24. Action Required 5-column table structure & signed formatting verification (never parentheses)