# Bank Reconciliation (Reconciliation.gs)

## The Core Problem

Manual expense tracking can miss transactions, double-count entries, or record the wrong day. `Reconciliation.gs` exists to **catch those gaps** by comparing your **daily inflow/outflow totals** in the `List` sheet against totals computed from the raw **SBI bank export**. Only dates with mismatches are surfaced so you can quickly fix the manual ledger.

---

## Prerequisites & Setup

### 1) Required sheets

Your spreadsheet must include:

- **`List`**: your manual ledger.
- **`Bank_Raw`**: where you paste the SBI statement export.
- **`Reconciliation_Log`**: created/overwritten by the script.

### 2) `List` sheet format

The reconciliation logic expects these columns by position:

| Column | Required content |
|---|---|
| A | Date |
| C | Account (string identifier) |
| E | Amount (positive = inflow, negative = outflow) |

### 3) SBI export headers

The script scans `Bank_Raw` until it finds a header row containing:

- **`Txn Date`** or **`Value Date`**
- **`Debit`**
- **`Credit`**

These must appear **on the same row**. If your export uses different labels, update `BANK_DATE_HEADERS`, `BANK_DEBIT_HEADER`, and `BANK_CREDIT_HEADER` in `RECON_CONFIG`.

### 4) Required configuration

In `Reconciliation.gs`, set the account you want to reconcile:

```javascript
const RECON_CONFIG = Object.freeze({
  // ...
  RECON_TARGET_ACCOUNT: "9682", // exact List sheet account name in column C (case-sensitive)
});
```

The script will stop if `RECON_TARGET_ACCOUNT` is not set.

---

## Step-by-Step Execution (SBI Workflow)

1. **Download the SBI statement**
   - Export the statement as **Excel** from SBI internet banking.
2. **Paste into `Bank_Raw`**
   - Open `Bank_Raw` and paste the **entire** export (including any header blocks above the column headers).
3. **Confirm configuration**
   - Ensure `RECON_TARGET_ACCOUNT` matches the exact account string used in `List` column C.
   - Set `RECON_DATE_ORDER` if your dates are ambiguous (e.g., `03/04/2026`).
4. **Run the script**
   - In Apps Script, run `reconcileBankStatement()`.
5. **Review `Reconciliation_Log`**
   - The log lists only dates where manual and bank totals differ.

---

## Understanding the Output (`Reconciliation_Log`)

The script writes these columns:

| Column | Meaning |
|---|---|
| Date | Normalized as `YYYY-MM-DD`. |
| Manual Inflow | Sum of positive amounts in `List` for that date. |
| Bank Inflow | Sum of SBI `Credit` values for that date. |
| Inflow Diff | `Manual Inflow - Bank Inflow`. |
| Manual Outflow | Sum of negative amounts in `List` for that date. |
| Bank Outflow | Sum of SBI `Debit` values as **negative totals**. |
| Outflow Diff | `Manual Outflow - Bank Outflow`. |

### What the diffs mean

- **Inflow Diff**
  - `> 0`: Manual ledger shows more inflow than the bank (possible extra entry).
  - `< 0`: Bank shows more inflow than the ledger (missing income entry).

- **Outflow Diff** (outflows are negative totals)
  - `> 0`: Manual outflow is **less negative** than the bank (missing expense entry).
  - `< 0`: Manual outflow is **more negative** than the bank (duplicate or over-recorded expense).

Only rows with **non-zero diffs** appear in the log.

---

## Notes

- Reconciliation is **single-account**: only `List` rows where column C equals `RECON_TARGET_ACCOUNT` are included.
- `Reconciliation_Log` is **overwritten** on each run.
- The script **does not modify** `List`.
