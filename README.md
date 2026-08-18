# ERP vs Tally — Supplier Ledger Reconciliation

Automates the manual date + Dr/Cr check between an ERP supplier ledger and a
Tally ledger export.

## Run it

Double-click **`run.bat`**, or:

```
python app.py
```

Then open <http://127.0.0.1:5000> and upload the workbook.

The workbook needs one Tally ledger sheet and one ERP ledger sheet. Both are
detected automatically; if it picks wrong, override them with the dropdowns on
the result page and press Re-run.

Command line (same engine, no browser):

```
python reco.py Book1.xlsx
```

Exits with code 1 if the reconciliation does not tie to zero, so it can be
wired into a scheduled job.

## What you get

The summary is a ledger control sheet — the same four lines an accountant would
check by hand, side by side:

| | Tally | ERP |
|---|---|---|
| **1. Opening balance** | from the export header, 0.00 if not stated | same |
| **2. Total DEBIT** | payments, TDS, discounting, bank settlements | same |
| **3. Total CREDIT** | purchases and other supplier charges | same |
| **4. Closing balance** | Opening + Credit − Debit | same |

A control check proves `Opening + Credit − Debit − Closing = 0.00` on both sides,
so the totals cannot silently drift from the source export.

The **Debit side** and **Credit side** are then each broken into three lines,
with the full row-level detail underneath:

- **Matched** — counterpart found and values agree
- **Tally internal contra (net nil)** — self-reversing bank pairs; ERP is not
  expected to carry them, so they are shown separately rather than counted as a
  break
- **Mismatched** — no counterpart, or values disagree

Every detail row shows the reference, document key, description, amount, **what
it matched against in the other system**, the document difference and the reason.
Filter pills isolate matched / contra / mismatched, or just the Tally rows or
just the ERP rows.

Excel report sheets: `Summary`, `Reconciliation`, `Exceptions`, `Debit detail`,
`Credit detail`, `Tally (parsed)`, `ERP (parsed)`, `Data quality`, `Method`.

## Files

| File | Purpose |
|---|---|
| `reco_engine.py` | All parsing, matching and report logic. The single source of truth. |
| `app.py` | Web upload UI on top of the engine. |
| `reco.py` | Command-line runner on top of the engine. |
| `run.bat` | Windows launcher. |

Uploads and generated reports land in `uploads/` and `reports/` and are purged
automatically after 24 hours.

## Matching rules

| Entry type | Rule |
|---|---|
| **Purchase** | Must match on the **invoice / document reference**. A purchase with no invoice number is reported as unmatchable rather than guessed at. |
| **Bank** | **Date + amount, exact.** No tolerance is applied to bank entries. Date window 15 days, and the receipt/payment wording is cross-checked between systems. |
| **TDS** | Every TDS booked in Tally is searched for on the **ERP Debit side only**. Found on the credit side instead, it is reported as a direction error; not found at all, it is Mismatched. |
| **All others** | Amount within **± Rs 1.00** (Tally 10.46 vs ERP 11.00 counts as matched), date within 45 days, then the **reference number is confirmed**. |
| **Unmatched** | Anything without a counterpart goes straight to **Mismatched**. No special buckets. |

Tolerances live at the top of `reco_engine.py`: `TOL` (1.00), `TOL_BANK`
(0.005 — effectively exact), `DATE_WINDOW` (45), `BANK_DATE_WINDOW` (15).

### What is checked on every match

Direction (Dr/Cr), amount within the applicable tolerance, date, and the
reference number. Each matched row records the **criteria actually checked** and
the **counterpart it matched against**, so any pairing can be re-traced.

### When several rows could match

Candidates are ranked by: reference agreement → nearest date → closest amount →
earliest row in the file. If more than one candidate ties on all of those, the
match is still made deterministically but flagged **AMBIGUOUS** for a human to
confirm. See "Same vendor, several transactions in one day" below.

## Same vendor, several transactions in one day

- **References present** — the reference decides. Order in the file is
  irrelevant; three same-day, same-amount invoices pair correctly.
- **No references, distinct amounts** — the amount decides cleanly.
- **No references, identical amounts** — genuinely undecidable. Rows are paired
  deterministically and flagged AMBIGUOUS. Because the amounts are identical the
  totals are unaffected whichever way they pair, so the reconciliation still
  ties; only the row-level attribution needs a human eye.

The real fix is to carry the reference number into ERP for these entries.

## The proof

Every classified difference accumulates into a **balance bridge** walking the
Tally closing balance to the ERP closing balance. The residual **must be zero**.
A non-zero residual means the reconciliation is incomplete and must not be
signed off — the web page shows a red banner and the CLI exits non-zero.

## Scope note

Validated against a single supplier ledger. Before rolling out across all
suppliers, re-test the assumption in step 2 (dropping the series prefix). It is
safe where document numbers do not collide across series; the Data quality sheet
flags every conflict it depends on, so a bad case shows up rather than silently
mis-matching.

---

## Author

Built by **Manohar Kumar Sah** ([@mmpbrother94](https://github.com/mmpbrother94)).
