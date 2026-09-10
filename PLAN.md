# Plan — EIMS Inventory + RUPS lookup web app

> Historical design record. For the current, verified system documentation see [WORKFLOW.md](WORKFLOW.md).

## Goal
A locally hosted web app that:
1. Displays EIMS Inventory data (cached .txt file + a Refresh button that pulls live via `curl --negotiate`).
2. Filters EIMS (Operation / Mgr_Name / Department / Prodgroup3), editable from the web UI.
3. Selects EIMS lots → searches RUPS by LotNumber → RUPS API.
4. Accepts manual Visual ID / Lot entry for direct RUPS lookups.

## Decisions (from the original user requirements)
- EIMS source: cached `EIMS_Inventory_Report.txt` plus a Refresh button (live download).
- Apply the preset FILTERS but allow editing them in the UI.
- Select EIMS lots to search RUPS.
- Map EIMS `LotNumber` → RUPS lot.

## Design
- Main file: `streamlit_app.py` (keep `app.py` and `export_eims_inventory.py`, reuse their logic).
- Reuse: `query_units()`, `export_excel()`, RUPS HEADER/URL; `download()` via curl, `parse_tab()`, FILTERS.
- Two tabs:
  - A. Manual RUPS search: enter Lot(s) + Visual ID(s) → Search.
  - B. EIMS Inventory: pick product + filters → table → Search RUPS.

## Steps
1. Scaffold: constants, BASE_DIR, paths, HEADER/URL, FILTERS.
2. EIMS layer: `download_eims()` (curl), `load_eims()` (parse cache), `apply_filters()`.
3. RUPS layer: `query_units(values)` + `export_excel()`.
4. UI rendering: filter form + table + manual form + results area.
5. Handlers / routes: page, filter, refresh, EIMS→RUPS, manual, download.
6. `main()`: bind the port, open the browser.

## Verification
1. Run `miniforge python streamlit_app.py` → the browser opens.
2. The EIMS tab shows filtered rows; changing a filter changes the row count.
3. Refresh EIMS → re-downloads the live file and updates the timestamp.
4. Tick a lot → Search RUPS → RUPS table + Excel download.
5. Manual Visual ID/Lot → RUPS results.

## Scope
- Included: local UI, EIMS display/filter/refresh, RUPS lookup (selection + manual), Excel export.
- Excluded: changing authentication, server deployment, editing EIMS data, other RIMS tabs.

## Open considerations
- Rendering a large EIMS table as HTML may be slow → limit the preview size.
- Many lots → combine into a single API call (comma-joined).

---

## Extensions delivered after the original plan
- Migrated from the stdlib `http.server` to **Streamlit** (`streamlit_app.py`).
- **PPV / Class** classification based on the SSPEC portion of the EIMS Product column.
- Filtering across the **full mapped product list** (Prodgroup3 multiselect).
- **Critical Dispose** section: `Days_At_Operation > 60`.
- Modern UI: Intel blue theme (`.streamlit/config.toml`), gradient banner, metric cards, colored badges, progress columns.
- RUPS search: **EIMS ↔ RUPS reconciliation per lot** (EIMS_Qty vs RUPS_units), warns on mismatch and reports lots that were not found.
- The by-lot table gained `Prodgroup3`, `LOSE_OPERATION`, and `PI_Dispose` columns.
- `LOSE_OPERATION` → `PI Dispose` mapping via `LOSE_OPERATION MAPPING.xlsx`.
- Display structure: **outer block = Product**, inner block = **PI_Dispose**.
- Added the `tenx-pe-knowledge-mcp-server` MCP entry to `.vscode/mcp.json`.

---

## Later phases (superseded this plan)
- Replaced Streamlit with a **Flask + HTML/JS app** in `web/` (port 8600) because Streamlit reruns
  made click-to-jump interactions laggy. Streamlit now lives in `streamlit_backup/`.
- Added **ATMf (InTMS) ticket submission** using signal 221 plus an action-flow PATCH.
- Added a **daily 07:00 VN scheduler** with 24-hour caching and pre-warming.
- Automated the **MARS loss-operation query** through the SQLPathFinder CLI.

See [WORKFLOW.md](WORKFLOW.md) for the full, current details.
