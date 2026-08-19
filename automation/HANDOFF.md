# MTO & G&G – FMO automation — handoff / resume notes

_Last updated: 2026-07-31. Status: **v1.3 — everything below + store-tab totals,
rebuilt clean Summary charts, a standalone HTML dashboard, and a near-zero-touch
run (auto-detects the ranking file, auto-names the month, one-click .bat).
Tested + validated. Safe to use.**_

## HOW TO RUN IT EACH MONTH (the short version)
1. Download the "Product Group Ranking" from Focal Point, unmerge it, and
   **save it into the report folder** (the folder that holds the workbooks —
   `…\MTO & G&G-FMO Reports\`). Keep the words "Product Group Ranking" in the
   name; the number in parentheses doesn't matter — the script grabs the NEWEST
   one automatically.
2. Open `Automation\update_report.py` and set the ONE line `TEMPLATE_FILE` to
   last month's finished workbook (e.g. this month point it at July's file).
3. Double-click `Automation\run_report.bat` (or run `python update_report.py`).
   Excel works headless for ~30-60 s. When it's done you'll have, in the folder:
   `01 - MTO & G&G - FMO <Month> 2026.xlsx` and `Dashboard <Month> 2026.html`.
   The new month's name, output filename and dashboard title are all derived
   automatically from the template — nothing else to type.
4. Review the **Dashboard** tab, the **Change Log** tab, and open the .html in
   any browser. Everything is already calculated (no F9, no pasting).

## v1.3 additions
- **Store-tab totals**: each store sheet (0008/0028/0062/0025) now ends in a
  bold TOTAL row (sum of Qty + $). Doesn't affect the Summary's UPC lookups.
- **Summary charts kept in the ORIGINAL look**: the per-category line charts
  came out degraded because the Python library, on save, drops each chart's
  (cell-linked) TITLE and its plot-area fill. The library actually KEEPS the
  teal chart-area fill, the red line and the positions - so instead of
  rebuilding, we just restore the two dropped things in openpyxl:
  `restyle_original_charts()` re-adds each white bold category title and the
  teal plot-area fill, and drops the one broken (#REF!) placeholder chart.
  The big multi-series "Catogory Chart" is handled separately: openpyxl keeps
  its title/gradient/legend/colours, so it's left untouched except for
  re-adding its DATA TABLE (the value grid under the lines) which openpyxl
  drops. Result: all Summary charts look exactly like the June original, window
  advanced to the new month. (An alternate clean-sparkline rebuild is available behind
  `REBUILD_SUMMARY_CHARTS=True`, off by default. A COM chart-copy path also
  exists but is unused — the Excel clipboard is blocked in headless sessions.)
- **Standalone HTML dashboard** (`Dashboard <Month>.html`): self-contained,
  offline, opens in any browser, safe to email. KPI tiles + 4 SVG charts;
  the category chart is sorted high→low. Same blue/gold/green palette.
- **Near-zero-touch run**: RAW_FILE / NEW_MONTH_LABEL / REPORT_TITLE_PERIOD /
  OUT_FILE all default to None = auto. Only TEMPLATE_FILE is maintained by hand.
  `run_report.bat` gives a double-click launch.
- **COM insert hardened**: new items are inserted by shifting ONLY the A:P item
  block (not whole rows), so the side-by-side Trend table and the far-right
  charts are never disturbed. Verified: after a test insert the Trend categories
  and totals row stay intact and the report still reconciles.

## v1.2 — what a single run now does (fully hands-off)
The script now finishes the whole month in one command. openpyxl does the data
work, then **Excel is driven via COM (win32com)** to insert new items, recalc,
and save — so the delivered file opens already computed (no manual F9).
1. Rewrites the 4 store sheets from the raw ranking. ✅
2. Widens the capped VLOOKUP ranges. ✅
3. Advances the 12-month **Trend** table (shift-and-freeze — see v1.1 below). ✅
4. **Fixes a latent template bug**: the company Total-DMK $ cell (P392) was
   missing one category term (+P387) that every other column includes, so the
   grand-total dollars were understated ~$123/mo. Rebuilt from the units-total
   row set so they can never diverge. Now reconciles exactly
   (sites $117,752.72 == grand total). Logged in the Change Log.
5. **Removes junk sheets**: `Sheet1` (held only an AI concept image) and
   `No Movement old` (16,378 phantom cols). File dropped 10.9 MB → ~0.5 MB.
6. Builds a **Dashboard** summary sheet (first tab): 5 KPI tiles + 4 charts
   (Sales by Store, 12-month volume trend, MTO vs G&G by store, category
   volume). All cells are live formula refs into the reconciliation block
   (rows 390-394) + Trend totals, so Excel repairs them on recalc/insert.
   Colors: Total DMK = blue, Uber & GrubHub = gold, Compound = green
   (CVD-validated). Print area excludes the off-screen staging tables.
7. **Auto-inserts NEW items** into the Summary via Excel COM: finds each item's
   category, inserts a row inside the block (Excel expands the subtotal SUM and
   repairs Table1 / reconciliation / trend refs automatically), sets category /
   name / UPC / unit, copies the VLOOKUP formulas. Change Log lists what was
   inserted — and only shows that note when items actually appeared.
8. Excel full-recalcs and saves. If COM fails for any reason, the openpyxl file
   is still valid — the console tells you to open + F9 + paste from Change Log.

## How the monthly report works (understood & verified)
- **Raw input** = "Product Group Ranking" file from PDI Focal Point. One sheet, all 4
  stores stacked. Store name in col C, category in col D, item rows = name(F) / UPC(G) /
  unit(H) / qty(I) / $(J). Subtotal rows say "<cat> Total" in col B.
- **Store → sheet map:** Demo Market Northgate→`0008`, Riverside Station→`0028`,
  Lakeview Station→`0062`, Summit Station→`0025`.
- **Everything is joined on UPC.** Summary col E = UPC; each site sheet is VLOOKUP'd by
  UPC (site col F) → qty (col H) and $ (col I).
- **Summary "left block"** (rows 7–~389): category header → items → `Total` subtotal.
- **DANGER ZONE — do NOT let a script move rows here:** rows ~390–394 reconciliation
  (`=G11+G21+…` absolute refs) and rows ~400–431 an Excel **structured Table `Table1`**
  (`=G274`, `Table1[[#This Row]…]`, `SUBTOTAL(109,Table1[…])`) — ~146 formulas keyed to
  the item subtotals **by absolute row number**. Excel repairs these on an in-app row
  insert; openpyxl does NOT → it would corrupt them. That's why new-item insertion is
  left to Excel (see below).
- **Trend table** = cols R–AE, rows 5–34 (12-month category qty). Shares rows with the
  left block, so it's another reason not to insert rows via script.

## Category renames (raw → summary), the only non-1:1 mappings
`Breakfast Bagel (Biscuits)`→`Breakfast Bagel`, `Fried Chicken`→`Chicken Tenders`,
`DMK - Catering`→`Catering`. (Ambiguous: `World Cup Slushies` — verify manually.)

## What v1 automates (script: `update_report.py`)
1. Rewrites the 4 store sheets from the raw file (the big copy/paste chore). ✅
2. Widens the 3 capped VLOOKUP ranges (`0008`=$I$758, `0028`=$I$821, `0062`=$I$862)
   to full column so a longer month never silently drops items. Same cells, safe. ✅
3. Adds a **Change Log** sheet: NEW items (+ exact target category & per-store qty/$ to
   paste), NO-MOVEMENT / missing items, unmapped categories, unknown stores. ✅
4. Leaves the Summary formulas / Table1 / reconciliation untouched. VLOOKUPs
   recalc in Excel (press F9). Writes to a NEW file; never edits your inputs.
5. **Auto-updates the 12-month Trend table** (Summary cols T–AE, rows 6–33) —
   see below.

## What v1.1 adds: Trend table auto-update (`UPDATE_TREND`)
The newest Trend column (AE) was already a **live formula** summing the four
site subtotals for each category (`=$G$11+$I$11+$K$11+$M$11`). So the monthly
job is a **shift-and-freeze**, exactly what you did by hand:
- Slide the 12-month window left one: `T..AC ← U..AD`, dropping the oldest month.
- **Freeze** last month's *computed* AE value into AD (read from the template's
  cached values, so it's the real number, not a stale formula).
- Leave AE as its live formula → it recomputes to the new month from this
  month's store sheets when Excel recalculates.
- Shift the month headers the same way; set AE's header to `NEW_MONTH_LABEL`.
- Rewrite the row-34 column totals as `SUM(col6:col33)` (was `:32`, which
  silently dropped a category — e.g. World Cup Slushies — once it left AE).
- **Idempotent:** if AE's header already equals `NEW_MONTH_LABEL`, the shift is
  skipped, so re-running the month can't double-advance the window.
- No raw→trend category map is needed (the AE formulas already encode the
  summary→trend mapping by row). A brand-new product group with no Trend row is
  the one case still needing a manual row-add — rare; flagged via Change Log's
  unmapped-category list if the category is also new to the Summary.

**No Movement tab:** nothing to build — every cell is a live formula mirroring
the Summary row-by-row (`='Summary Comparision Report'!B7`…), so it recalcs
itself. The Change Log's NO-MOVEMENT list gives the explicit missing-item set.

## Validation done (June data)
- Site totals reproduce ground truth exactly: Northgate 4,948/$17,205.66 · Riverside
  6,284/$21,435.21 · Lakeview 11,713/$41,046.08 · Summit 6,302/$26,998.46 · **Total
  29,247 / $106,685.41 and it reconciles (sum-of-sites == Total).**
- **Bug found in the current June file:** its Summary *Northgate* column held stale cached
  values (17,552 / $75,700) that don't reconcile; the automation's output is correct.
- New-item detection + category mapping unit-tested (3 simulated new items placed right).
- Table1 / reconciliation / trend confirmed intact in the output.

## Requirements & switches
Requirements: Windows + **Excel installed** and **pywin32** (`pip install pywin32`).
Both confirmed present on this machine (Excel 16.0, py 3.11). Run steps are in the
"HOW TO RUN IT EACH MONTH" section at the top.

CONFIG feature switches (all default ON): `ADD_STORE_TOTALS`,
`REBUILD_SUMMARY_CHARTS`, `BUILD_DASHBOARD`, `BUILD_HTML_DASHBOARD`,
`AUTO_INSERT_ITEMS`, `RECALC_WITH_EXCEL`, `CLEAN_JUNK_SHEETS`, `UPDATE_TREND`,
`FIX_VLOOKUP_RANGES`. Set any to False to skip that step. The auto-fields
(`RAW_FILE`, `NEW_MONTH_LABEL`, `REPORT_TITLE_PERIOD`, `OUT_FILE`) can be set to a
literal value to override the auto-detection for a one-off run.

## Idempotency / safety notes
- Re-running the same month is safe: the Trend shift skips if AE already shows
  `NEW_MONTH_LABEL`; the grand-total fix is a no-op once consistent. But new-item
  **insert is NOT idempotent** — running twice with the same template would insert
  twice. Always build from LAST month's file, not from an already-processed output.
- The build always writes a NEW `OUT_FILE`; your inputs are never modified.

## NEXT / TODO
- [x] ~~12-month Trend auto-update~~ — DONE (v1.1).
- [x] ~~No Movement tab~~ — self-updating mirror, no work needed.
- [x] ~~Hands-off new-item insert (Excel COM)~~ — DONE (v1.2).
- [x] ~~Drop Sheet1 + No Movement old~~ — DONE (v1.2).
- [x] ~~Dashboard + professional charts~~ — DONE (v1.2).
- [ ] Once a real month with genuinely NEW items runs, eyeball the auto-inserted
      rows once to confirm placement (logic validated with a synthetic item).
- [ ] THEN move on to Folder 2 "Target Margin Report" (per user, after this is
      confirmed working in production).
