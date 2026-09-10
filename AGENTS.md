# AGENTS.md — Unit management (RUPS + EIMS reject app)

Internal Intel tool for reject/unit management. Shows EIMS Inventory data, filters it,
cross-references lots against the RUPS API, classifies PPV/Class, and submits ATMf tickets.
Runs on the developer's machine, shared with the team over the Intel network / VPN.

**Primary app is now the Flask + HTML/JS web app in [web/](web/). Streamlit is kept as a
backup in [streamlit_backup/](streamlit_backup/).**

## Big picture
- **[web/](web/)** — PRIMARY app (Flask). `core.py` (backend logic), `app.py` (JSON API +
  server on port 8600), `templates/index.html`, `static/{style.css,app.js}` (single page).
- **[streamlit_backup/](streamlit_backup/)** — old Streamlit app (`streamlit_app.py`,
  `.streamlit/`, `run_share.bat`, port 8501). Backup / reference; shares root data/ & output/.
- **[scripts/](scripts/)** — standalone data pullers (`export_eims_inventory.py`, `rups_data.py`).
- **[data/](data/)** — cached `EIMS_Inventory_Report.txt`, `LOSE_OPERATION MAPPING.xlsx`,
  `Lot_loss_operation.csv`, `intms_products.json`.
- **[output/](output/)** — generated Excel/email/ATMf request scratch files.
- **[WORKFLOW.md](WORKFLOW.md)** — full verified reference: every workflow, step, and script I/O.
- **[PLAN.md](PLAN.md)** — original design decisions and history.

**All project files, docs, code comments and UI text are written in English.**

## Run / share
- **Production (users):** `run_prod.bat` — served from the `prod/` git worktree (branch `main`) on
  port 8600. LAN: `http://<host>:8600`.
- **Development (you):** `run_dev.bat` — working copy (branch `dev`) on port 8601, scheduler
  disabled, orange DEV banner. Edit freely; users are unaffected.
- **Publish:** `deploy.bat` (fast-forwards `main` inside the worktree), then restart `run_prod.bat`.
- **Fresh clone:** `setup_prod.bat` creates the worktree.
- **Streamlit (backup):** `streamlit_backup\run_share.bat` (port 8501).

## Secrets
- RUPS credentials live in **`config.local.json`** (git-ignored) or env vars
  `RUPS_SITE`/`RUPS_TOKEN`/`RUPS_WWID`. Copy `config.example.json` to start.
- **Never hard-code tokens/WWIDs, and never commit factory data** (EIMS/MARS exports are ignored).

## Environment (critical — see [python-scripts instructions](.github/instructions/python-scripts.instructions.md))
- No `python`/`python3` on PATH. Use the full **miniforge** interpreter path above for the app.
- **PyPI/pip and github.com are blocked/timeout.** Install deps ONLY via conda-forge:
  `conda install -n ngocluup -y -c conda-forge streamlit pandas requests openpyxl`
  (conda prints to stderr → PowerShell may report exit 1 on success; verify with an import check).
- Standalone `scripts/` follow the stdlib-only, Python 2.7/3.3-compatible rules in the instructions
  file (urllib fallback, `ssl.SSLContext(PROTOCOL_SSLv23)` for verify=False, CSV+BOM for Excel).

## Data-source conventions (verified — do not re-derive)
- **EIMS**: cached `data/EIMS_Inventory_Report.txt`, refreshed live via `curl --negotiate` (Kerberos,
  no password). Join key `LotNumber`. Default filters live in `DEFAULT_FILTERS` in the app.
- **RUPS** (`CUSTOMIZED_API_SEARCH_UNIT_INFO`, POST `https://rups.intel.com/RUPS_api`):
  returns **one row per unit (VID)**. `QUANTITY` is lot-level, **repeated on every row — never sum it**.
  RUPS unit count = `df.groupby(lot)["VID"].nunique()`. Response shape: `return.data = [[{...}], []]`.
- EIMS_Qty (lot-level) vs RUPS unit count can legitimately differ — that mismatch is expected and surfaced.
- PPV vs Class derived from RUPS `SSPEC` (non-empty = PPV, empty = Class).
- `LOSE_OPERATION` → PI Dispose via `data/LOSE_OPERATION MAPPING.xlsx` (Operation | Domain | PI Dispose).

## Gotchas
- `str(NaN) == "nan"` is truthy; guard filters with
  `pd.notna(v) and str(v).strip() and str(v).strip().lower() != "nan"`.
- RUPS `TOKEN`/`WWID` headers are sensitive credentials — do not commit publicly or log.
- Network calls must always pass a `timeout`.
