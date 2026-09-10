---
applyTo: "**/*.py"
description: "Conventions for writing & running Python scripts on this machine (no modern Python 3; PyPI/GitHub are blocked)."
---

# Python conventions on this machine

## Environments
- There is **no** `python` / `python3` on PATH → always use a full path.
- **Preferred (web app):** `C:\Users\ngocluup\AppData\Local\miniforge3\envs\ngocluup\python.exe`
  (Python 3.14 — has pandas, requests, openpyxl, flask, waitress, oracledb).
- `C:\Python27\python.exe` (2.7.9) — working `ssl`; pip 1.5.6 is too old to install anything.
- `C:\Python33\python.exe` (3.3.0) — no pip; `ssl` lacks `_create_unverified_context`.
- The `py` launcher points to Python 2.7.9.
- **Standalone scripts in `scripts/`** default to `C:\Python27\python.exe` and must stay stdlib-only.

## Network
- github.com and PyPI installs are **blocked / time out**.
- Install packages **only via conda-forge**:
  `conda install -n ngocluup -y -c conda-forge <packages>`
  (conda writes warnings to stderr, so PowerShell reports exit code 1 even on success — verify
  with an import check instead of trusting the exit code).
- Direct HTTPS to internal Intel sites works fine.

## Coding rules
- **Stdlib-only for `scripts/`.** The `web/` app may use the miniforge dependencies.
- **HTTP:** keep `requests` optional behind `try/except ImportError` with a `urllib` fallback.
- **Always set a `timeout`** on network calls.
- **SSL `verify=False`** in a way that works on both 2.7 and 3.3:
  ```python
  import ssl
  ctx = ssl.SSLContext(ssl.PROTOCOL_SSLv23)
  ctx.check_hostname = False
  ctx.verify_mode = ssl.CERT_NONE
  ```
  Do not use `ssl._create_unverified_context()` (missing on Python 3.3).
- **Python 2 & 3 compatibility:** branch on `sys.version_info[0]`.

## Excel export
- Without `openpyxl`/`pandas`, emit **CSV with a UTF-8 BOM** (Excel opens it directly).
- Python 3: `open(path, "w", newline="", encoding="utf-8-sig")`.
- Python 2: `open(path, "wb")`, write `b"\xef\xbb\xbf"`, then encode each cell to utf-8.

## Security
- Files containing `TOKEN` / `WWID` hold sensitive credentials — **never commit them publicly**.

## RUPS API notes
- POST `https://rups.intel.com/RUPS_api`, headers `SITE`/`TOKEN`/`WWID`, JSON body `{api, unit_list}`.
- Response: `return.data = [[{...unit fields...}], []]` → walk the inner lists and collect dicts as CSV rows.
- RUPS returns **one row per unit (VID)**; `QUANTITY` is lot-level and repeated on every row —
  never sum it. Count units with `groupby(lot)["VID"].nunique()`.

## Language
- **All project files, docs, code comments and UI text must be written in English.**
