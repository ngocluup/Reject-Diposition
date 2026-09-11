---
name: reject-management
description: "Streamlit web app (Intel VNAT) that reconciles reject/units between EIMS Inventory and RUPS, classifies lots PPV/Class/Eng_Assessment, and creates ATMf (InTMS) reject tickets. Use when: user asks about the Reject Management app, EIMS <-> RUPS reconciliation, EIMS inventory filtering, RUPS unit lookup, PPV vs Class classification, PI Dispose / LOSE_OPERATION mapping, quantity mismatch, or submitting/creating ATMf reject tickets (signal 221 Reject Management - POR), scrap/transfer-to-HVE-Lab requests."
argument-hint: "Say what you need: 'run the app', 'filter product X', 'reconcile EIMS vs RUPS', 'submit ATMf ticket for lots', 'why Qty mismatch', 'add product to mapping'"
---

# Reject Management — EIMS ↔ RUPS reconciliation + ATMf ticketing

Internal Intel VNAT Streamlit app. Loads EIMS Inventory, cross-references lots against
RUPS, reconciles quantities, classifies lots (PPV / Class / Eng_Assessment), emails a
report, and creates ATMf reject tickets in InTMS. Single file: `streamlit_app.py`.

## Run
```
C:\Users\ngocluup\AppData\Local\miniforge3\envs\ngocluup\python.exe -m streamlit run streamlit_app.py
```
Share on LAN via `run_share.bat` (port 8501). Deps via conda-forge only (PyPI/GitHub blocked):
`conda install -n ngocluup -y -c conda-forge streamlit pandas requests openpyxl`.

## What the app supports

### 1. Manual RUPS search
Paste lots / visual IDs → query RUPS → view + download Excel.

### 2. EIMS Inventory (main flow)
- **Data source**: download/refresh EIMS `.txt` via `curl --negotiate` (Kerberos SSO, no password).
- **Filters**: Product (Prodgroup3), Operation, Mgr_Name, Department. Default products = ADLN, ARLS816B.
- **Reconcile EIMS ↔ RUPS**: auto-query RUPS for filtered lots; compare EIMS_Qty vs RUPS unit
  count; flag Match / N/A / Not found / MISMATCH; sort by Status then Last-used days.
- **Classify** each lot PPV / Class / Eng_Assessment (SSPEC + LOSE_OPERATION → PI Dispose map).
- **Display by product**: one card per product, PPV/Class/Eng metrics, one tab per group.
- **Export**: download any lot table as CSV/Excel or copy it to the clipboard.

### 3. Submit ATMf ticket (InTMS)
- Tick lots directly in each table; every table has its own **Select all / Clear**.
- Auto-map Prodgroup3 → ATMf Product Name (override list + fuzzy); "Force product" fallback.
- **One ticket per product** (grouped); pick a per-product **Request** (Scrap / Transfer to HVE Lab).
- Prefills the "Fill in required information*" step (request + lot list) and prints the ticket URL.

## Key facts (verified — do not re-derive)
- **RUPS** POST `https://rups.intel.com/RUPS_api`, headers SITE/TOKEN/WWID; returns ONE row per
  VID; `QUANTITY` is lot-level repeated per row — never sum. RUPS_Qty = `groupby(lot)["VID"].nunique()`.
- **EIMS** cached `data/EIMS_Inventory_Report.txt` (tab-delimited); join key `LotNumber`.
- **PPV vs Class** from RUPS `SSPEC` (non-empty = PPV). `LOSE_OPERATION` → PI Dispose via
  `data/LOSE_OPERATION MAPPING.xlsx`; op containing 7226 → Eng_Assessment.
- **ATMf/InTMS** host `atmf.intel.com` (NOT intms.intel.com). Create: POST
  `/api/custom/ims/ticketing/` body `{"signal":221,"form_json":{...}}` (signal 221 =
  "Reject Management - POR"). form_json string fields: Customer, Payer BU, Product Stage,
  Product Name, Test Area, Factories, Shipping/LYA/FA Only, NFO shipment.
  The "Fill in required information*" step (action 4) must be PATCHed AFTER create:
  PATCH `/api/custom/ims/actionflow/<id>/` body
  `{"action_flow_json":{"id":"4","rich_text":"<p>...</p>"}}`. Product list in
  `data/intms_products.json` (432 items, read utf-8-sig). Ticket URL:
  `https://atmf.intel.com/intms/panel/detail/<id>`.
- All auth is Kerberos SSO via `curl.exe --negotiate -u :` (no password); always pass a timeout.

## Gotchas
- `str(NaN) == "nan"` is truthy; guard with `pd.notna(v) and str(v).strip().lower() != "nan"`.
- API error responses can be a LIST (e.g. `["signal can not be found!"]`) — guard `isinstance(dict)`.
- `data_editor` can't color cells → a Flag column (⚠ mismatch) is used instead.
- After code edits, remind the user to press **Refresh EIMS** (caches hold data).
