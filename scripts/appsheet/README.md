# AppSheet Budget Workflow

This folder documents the primary AppSheet-based intake path and its backend **PORTFOLIO NEXUS Sweeper Engine** that normalizes and restructures rows after users submit them.

## Main Script

- [IndTran.gs](IndTran.gs) contains the `processTransaction()` function.

## The PORTFOLIO NEXUS Sweeper Engine

`processTransaction()` is an "Armor-Plated" sweeper designed to be triggered on every spreadsheet change. It reads the `List` sheet directly, normalizes rows, and performs complex financial routing based on the content of the rows. 

**Core Mechanism:** To prevent missing any newly inserted rows from AppSheet (which often inserts rows above blank formatted "ghost" rows at the bottom of the sheet), the engine scans the *entire* spreadsheet on every run. Because the engine is idempotent, it safely ignores already-processed rows and only acts on new ones.

The engine relies on a robust schema. It looks for columns dynamically (case-insensitive and trimmed), specifically relying on:
- `Account` and `Amount`: To ensure the row has data.
- `Category`: Used to trigger Transfer logic.
- `Debt Entity` (with fallback to `DebtEntity` or `Person`): Used to trigger Split Bill logic.

### Module 1: Transfer & Investment Engine

**Trigger:** `Category` equals `Internal Transfer`, `Investments`, `Credit Card Payment`, or `Credit Card`.
**Logic:**
- The engine enforces a negative amount on the original row (the "OUT" leg) and clears the `Debt Entity` column.
- It automatically inserts a new row directly underneath the original row. This is the "IN" leg.
- The "IN" leg has a positive amount, assigns the destination account, and is marked as `[Transfer: In]`.
- Row IDs are updated cleanly (e.g., suffixed with `A` and `B`) to maintain unique keys for AppSheet.

### Module 2: Split Bill / Lending Engine

**Trigger:** The `Debt Entity` column contains a comma (e.g., `Me , Laxman`).
**Logic:**
- The engine splits the text by the comma and divides the original amount equally among all individuals.
- It updates the original row in-place for the first person (usually `Me`), setting the amount to the divided negative share and the `Status` to `Settled`.
- For each additional person, it inserts a new row directly underneath, appending the divided negative share and setting their `Status` to `Pending`. 
- An audit string (e.g., `[Split: 452 / 2 = 226]`) is appended to the `Notes` of all resulting rows.

### Module 3: Standard Expense Engine (Fallback)

**Trigger:** `Category` is standard (not a transfer) AND `Debt Entity` is either entirely blank or equals `"Me"`.
**Logic:**
- This is the catch-all for normal, everyday expenses.
- The engine guarantees the amount is properly negative.
- It normalizes `Debt Entity` to be blank (removing `"Me"`).

## Setup And Triggers

1. Put `IndTran.gs` into your Google Apps Script project.
2. In the Google Apps Script Triggers dashboard, create a new trigger:
   - **Choose which function to run:** `processTransaction`
   - **Select event source:** `From spreadsheet`
   - **Select event type:** `On change`
3. This guarantees that every time AppSheet syncs a new row, the engine immediately processes, splits, and stacks the data correctly.
