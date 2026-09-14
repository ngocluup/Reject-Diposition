# Reject Management

Internal Intel tool that reconciles reject/unit data between **EIMS Inventory** and the
**RUPS** API, enriches it with loss-operation history from **Oracle MARS**, classifies lots as
PPV / Class / Eng_Assessment, and submits **ATMf (InTMS)** reject tickets.

Runs as a small Flask app on the developer's machine and is shared with the team over the
Intel network / VPN.

> **Full documentation:** [WORKFLOW.md](WORKFLOW.md) — every workflow, step, and script I/O.

---

## Features

- **EIMS ↔ RUPS reconciliation** — compares lot-level EIMS quantity against the distinct unit
  (VID) count in RUPS and highlights mismatches.
- **Automatic classification** — PPV / Class from RUPS `SSPEC`; `Eng_Assessment` from the
  loss operation; PI Dispose mapped from an Excel lookup.
- **MARS loss-operation enrichment** — driven automatically through the SQLPathFinder CLI.
- **ATMf ticket submission** — creates signal-221 tickets and fills the required action-flow step.
  Remembers the product mapping you picked, flags lots that already have a ticket, and blocks
  duplicate submissions.
- **Search by lot** — look up any lot regardless of the current filters, reusing data already on screen.
- **Excel-like tables** — quick search, sortable headers and per-column filter dropdowns everywhere.
- **Export** — download any lot table as CSV/Excel, or copy it to the clipboard.
- **Fully automated daily refresh** at 07:00 Vietnam time, with 24-hour caching and pre-warming.

---

## Requirements

- Windows, joined to the Intel domain, VPN connected (Kerberos SSO is used for EIMS and ATMf).
- Python 3.14 via miniforge: `C:\Users\<you>\AppData\Local\miniforge3\envs\ngocluup\python.exe`
- Packages (conda-forge only — PyPI is blocked on this network):
  ```
  conda install -n ngocluup -y -c conda-forge pandas requests openpyxl flask waitress
  ```
- SQLPathFinder 3 installed (for the MARS loss-operation query).

---

## Setup

```powershell
git clone <repo-url> "Unit management"
cd "Unit management"
copy config.example.json config.local.json
# then edit config.local.json and fill in RUPS_TOKEN and RUPS_WWID
```

`config.local.json` is git-ignored and must **never** be committed.

---

## Running

| Environment | Command | Port | Scheduler |
|---|---|---|---|
| **Production** (what users open) | `run_prod.bat` | 8600 | enabled (07:00 VN) |
| **Development** (your sandbox) | `run_dev.bat` | 8601 | disabled |

Both are selected with the `RM_ENV` environment variable (`prod` / `dev`). The dev instance
shows an orange DEV banner so the two can never be confused.

Share URLs (Intel network / VPN): `http://<hostname>:8600`

---

## Development workflow

Production is served from a **separate git worktree** pinned to `main`, so editing files in your
working copy cannot affect users.

```
Unit management/        <- your working copy, branch "dev",  port 8601
Unit management/prod/   <- git worktree,     branch "main", port 8600  (git-ignored)
```

1. Work on the `dev` branch and test at <http://localhost:8601>.
2. When you are happy, run `deploy.bat` — it merges `dev` into `main`, updates the prod
   worktree, and restarts the production server.

---

## Configuration

| Variable | Default | Purpose |
|---|---|---|
| `RM_ENV` | `prod` | `prod` or `dev` |
| `RM_PORT` | 8600 / 8601 | Listening port |
| `RM_DATA_DIR` | `./data` | Shared data directory |
| `RM_OUTPUT_DIR` | `./output` | Scratch directory |
| `RM_CONFIG` | `./config.local.json` | Path to the secrets file |
| `RUPS_SITE` / `RUPS_TOKEN` / `RUPS_WWID` | from `config.local.json` | RUPS API credentials |

---

## Project layout

```
web/            Flask app (core.py = logic, app.py = API + server)
scripts/        Standalone data pullers and the SQLPathFinder query grid
data/           Cached EIMS / MARS data and mapping files
output/         Generated scratch files (git-ignored)
streamlit_backup/  Previous Streamlit implementation, kept for reference
```

---

## Security

- Credentials live only in `config.local.json` or environment variables.
- `config.local.json`, `output/`, and cached EIMS data are git-ignored.
- Never commit RUPS tokens, WWIDs, or exported factory data.
