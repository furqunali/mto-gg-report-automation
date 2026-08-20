# 🍽️ Foodservice Report Automation — MTO & G&G / FMO

**Turn a raw point-of-sale export into a fully-built monthly Excel report *and* a standalone HTML dashboard — in one command.**

> 🔴 **Live demo dashboard:** **https://mto-gg-dashboard.vercel.app**
>
> This repository is a **sanitized public demo**. The automation logic is the real production code; the sample data is fictional (a made-up "Demo Market" chain with generic store and category names). See [Security](#-security).

---

## 🎯 Problem

Every month, a foodservice/POS report was assembled **by hand in Excel**, and it took hours of error-prone, repetitive work:

- unmerge the raw "Product Group Ranking" export and split it per store,
- VLOOKUP this month's sales back into a Summary sheet keyed on UPC,
- add any newly-sold items into the right category (without breaking the subtotals),
- roll the 12-month trend window forward one month,
- refresh the No-Movement tab and reconcile every total to the grand total.

One wrong row insert silently corrupts a reconciliation block of ~146 absolute-row formulas — so the manual process was not just slow, it was fragile.

## 💡 Solution

Drop the new export into the report folder, point one config line at last month's workbook, and run **one command**. The tool produces, in ~30–60 seconds:

- `01 - MTO & G&G - FMO <Month> 2026.xlsx` — the finished, already-recalculated Excel report, and
- `Dashboard <Month> 2026.html` — a self-contained, offline, emailable dashboard.

The new month's label, output filename and dashboard title are all **auto-derived** from the template — the only thing maintained by hand is which workbook is "last month's".

## 🏗️ Architecture

The pipeline lives in [`automation/update_report.py`](automation/update_report.py) and runs in two phases: `openpyxl` does the data work, then Excel is driven headless via COM to do the structural work only Excel can do safely.

```
   Raw "Product Group Ranking" export (.xlsx, 4 stores stacked)
                        │
                        ▼
   ┌─────────────────────────────────────────────────────────┐
   │  parse_raw()            join everything on UPC            │
   │  write_store_sheets()   rebuild the 4 store tabs + totals │
   │  widen_vlookup_ranges() so a long month never drops items │   openpyxl
   │  update_trend()         shift-and-freeze the 12-mo window │   (data layer)
   │  fix_grand_total_formula() reconcile to the grand total   │
   │  build_dashboard()      Excel Dashboard tab (KPIs+charts) │
   │  restyle_original_charts() restore chart look openpyxl drops
   │  write_change_log()     what was added / missing / unmapped
   └─────────────────────────────────────────────────────────┘
                        │
                        ▼
   ┌─────────────────────────────────────────────────────────┐
   │  com_finalize()  (pywin32 / Excel COM, headless)          │
   │   • insert NEW items inside their category so Excel        │   Excel COM
   │     auto-expands subtotals + repairs Table1 & reconcile    │   (structure layer)
   │   • full recalc, then save                                 │
   └─────────────────────────────────────────────────────────┘
                        │
          ┌─────────────┴─────────────┐
          ▼                           ▼
   Finished .xlsx report      build_html_dashboard()
   (opens pre-computed)       standalone .html (inline SVG)
```

**Why the split.** The Summary sheet is really two tables sharing row numbers — an item block (`A:P`) and a side-by-side 12-month trend table (`R:AE`) — plus a reconciliation block and an Excel structured Table (`Table1`) whose ~146 formulas reference item subtotals *by absolute row number*. Inserting a row with a Python library would corrupt those references; Excel repairs them automatically on an in-app insert. So `openpyxl` handles everything that doesn't move rows, and Excel COM handles the one thing that does.

## ✨ Key Features

- **Near-zero-touch run** — auto-detects the newest `*Product Group Ranking*.xlsx`, derives the month/filename/title from the template. One line to maintain; one command (or one double-click via `run_report.bat`) to run.
- **UPC-keyed rebuild** — re-splits all four store tabs from the stacked raw export and adds a bold TOTAL row (Qty + $) to each.
- **Shift-and-freeze 12-month trend** — slides the window left, freezes last month's *computed* value, keeps the live formula for the new month. **Idempotent** — re-running a month can't double-advance it.
- **Hands-off new-item insertion** — new items are placed into the correct category via Excel COM so subtotals, `Table1` and the reconciliation formulas all self-repair.
- **Self-healing totals** — rebuilds the grand-total formula from the units-total row so dollars and units can never diverge (it fixed a latent ~$123/mo template bug).
- **Original-look charts preserved** — `openpyxl` drops cell-linked chart titles and plot-area fills on save, so the script *restores just those*, rather than rebuilding and losing the original styling.
- **Change Log tab** — lists new items (with exact target category + per-store qty/$), no-movement items, unmapped categories and unknown stores for a quick review.
- **Standalone HTML dashboard** — self-contained inline-SVG (no build step, no external requests): KPI tiles, sales by store, 12-month volume trend, MTO vs G&G split, and category volume sorted high→low.

## 🧰 Tech Stack

| Layer | Technology |
|---|---|
| Language | Python 3.11 |
| Excel read/write | [`openpyxl`](https://openpyxl.readthedocs.io/) |
| Headless insert / recalc | [`pywin32`](https://pypi.org/project/pywin32/) — Excel COM (Windows + Excel only) |
| Dashboard | Vanilla HTML + inline SVG / JavaScript (zero dependencies) |
| Launcher | Windows `.bat` one-click |

## 🧠 Engineering Decisions

- **`openpyxl` for data, Excel COM for structure.** `openpyxl` is fast and CI-friendly for reading/writing values, but it cannot safely insert a row into a sheet whose reconciliation and structured-table formulas are keyed by absolute row number. Rather than reimplement Excel's formula-repair logic in Python, the tool hands that single operation to Excel itself over COM — the right tool for each half of the job.
- **Restore, don't rebuild, charts.** `openpyxl` preserves a chart's series and area fill but drops cell-linked titles and plot-area fills on save. Rebuilding charts from scratch would lose the report's established look, so the script surgically re-adds only the two things Excel drops.
- **A self-contained HTML dashboard.** The audience is non-technical and often reads on email. A single `.html` file with everything inlined (no CDN, no build, no server) opens in any browser, works offline, and can be attached to an email as-is.
- **Idempotency as a safety contract.** The trend shift and the total-fix are no-ops when already applied, so a re-run is safe. The one non-idempotent step (item insert) is documented, and every run writes a **new** output file — inputs are never modified.

## 📊 Results / Demo

**▶ Live dashboard: https://mto-gg-dashboard.vercel.app**

The dashboard ([`demo/index.html`](demo/index.html)) renders KPI tiles and four inline-SVG charts entirely client-side — no build step and no external requests.

![Dashboard preview](demo/preview.png)

What used to be a multi-hour manual Excel session is now a single ~30–60s run that outputs a finished, pre-computed workbook and a shareable dashboard.

## ⚙️ Setup / Installation

> **Platform note:** the data/dashboard layer is cross-platform, but the headless insert-and-recalc step uses **Excel COM and requires Windows with Microsoft Excel installed** (`pywin32`). Developed against Excel 16.0 / Python 3.11.

```bash
pip install openpyxl pywin32

# 1) Put last month's finished workbook and the new "Product Group Ranking"
#    export in the report folder.
# 2) Point TEMPLATE_FILE (top of automation/update_report.py) at last month's workbook.
python automation/update_report.py
```

Or on Windows, double-click [`automation/run_report.bat`](automation/run_report.bat).

Each pipeline stage is a `CONFIG` feature switch (all default ON) — e.g. `AUTO_INSERT_ITEMS`, `UPDATE_TREND`, `BUILD_HTML_DASHBOARD`, `FIX_VLOOKUP_RANGES` — so any step can be toggled off for a one-off run. See [`automation/HANDOFF.md`](automation/HANDOFF.md) for the full operator notes.

## 🔒 Security

This is a **sanitized public demo**:

- All store names, categories, vendors and figures are **fictional** — a made-up "Demo Market" chain (Northgate / Riverside / Lakeview / Summit).
- `.gitignore` **blocks all Excel files** (`*.xlsx`, `*.xlsm`, `*.xlsb`, `~$*`), so **no real workbook or customer data is ever committed** to the repo.
- The dashboard is fully self-contained and makes **no external network requests**.

The automation logic in this repo is the genuine article; only the data is synthetic.

## 🗺️ Roadmap

- [x] UPC-keyed store rebuild + per-store totals
- [x] 12-month trend auto-roll (shift-and-freeze, idempotent)
- [x] Hands-off new-item insert via Excel COM
- [x] Excel + standalone HTML dashboards
- [ ] Eyeball auto-inserted rows on the first live month with genuinely new items (logic validated with synthetic items)
- [ ] Extend the pipeline to the companion **Target Margin / Weighted-Average** report

## 📄 License

[MIT](LICENSE)
