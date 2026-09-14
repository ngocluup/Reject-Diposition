"""Backend logic for the Reject Management web app (Flask).

Pure Python (no Streamlit). Reuses the verified RUPS / EIMS / ATMf logic:
- RUPS: POST rups.intel.com, one row per VID, QUANTITY is lot-level (never sum).
- EIMS: cached tab file, refreshed via curl --negotiate (Kerberos SSO).
- ATMf/InTMS: create ticket (signal 221) then PATCH action-flow step 4.

Run everything with the miniforge interpreter (conda-forge deps only).
"""
import io
import json
import os
import subprocess
import time
import threading
import difflib

import pandas as pd
import requests

try:
    requests.packages.urllib3.disable_warnings()  # type: ignore[attr-defined]
except Exception:
    pass

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# data/ and output/ can be redirected so a prod worktree and the dev checkout
# share one copy of the (large, slow to rebuild) EIMS/MARS data.
DATA_DIR = os.environ.get("RM_DATA_DIR") or os.path.join(BASE_DIR, "data")
OUTPUT_DIR = os.environ.get("RM_OUTPUT_DIR") or os.path.join(BASE_DIR, "output")
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---------------------------------------------------------------- secrets
# Credentials never live in source control. Resolution order:
#   1. environment variables (RUPS_SITE / RUPS_TOKEN / RUPS_WWID)
#   2. config.local.json in the project root (git-ignored)
# Copy config.example.json to config.local.json and fill it in.
CONFIG_PATH = os.environ.get("RM_CONFIG") or os.path.join(BASE_DIR, "config.local.json")


def _load_local_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8-sig") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
        except (ValueError, OSError):
            pass
    return {}


_LOCAL_CONFIG = _load_local_config()


def _secret(name, default=""):
    return os.environ.get(name) or _LOCAL_CONFIG.get(name) or default


# ---------------------------------------------------------------- environment
# "prod" serves real users; "dev" is a sandbox on another port with the
# scheduler disabled so it can never fight prod over EIMS/MARS refreshes.
APP_ENV = (os.environ.get("RM_ENV") or "prod").strip().lower()
IS_DEV = APP_ENV == "dev"

# ---------------------------------------------------------------- RUPS config
RUPS_URL = "https://rups.intel.com/RUPS_api"
RUPS_HEADER = {
    "SITE": _secret("RUPS_SITE", "VNAT"),
    "TOKEN": _secret("RUPS_TOKEN"),
    "WWID": _secret("RUPS_WWID"),
}

# ---------------------------------------------------------------- EIMS config
EIMS_SITE = "vnat"
EIMS_URL = "https://vnatmfg.intel.com/vf-rims-portal/home/%s/EIMS_Inventory_Report.txt" % EIMS_SITE
EIMS_TXT_PATH = os.path.join(DATA_DIR, "EIMS_Inventory_Report.txt")
EIMS_LOT_COLUMN = "LotNumber"
EIMS_FILTER_COLUMNS = ["Operation", "Mgr_Name", "Department", "Prodgroup3"]

LOSE_MAP_PATH = os.path.join(DATA_DIR, "LOSE_OPERATION MAPPING.xlsx")
LOT_LOSS_PATH = os.path.join(DATA_DIR, "Lot_loss_operation.csv")

# ---------------------------------------------------------------- SPF / MARS config
# SQLPathFinder CLI runs the "EIMS lot.VG2" query grid against MARS (it handles
# its own MARS auth). Input CSV must have LotNumber as column 2; output is
# Lot_loss_operation.csv. See /memories/repo/mars-oracle-access.md.
SPF_EXE = os.path.expandvars(
    r"%USERPROFILE%\My Programs\SQLPathFinder3\sqlpathfinder3.exe")
SPF_VG2 = os.path.join(BASE_DIR, "scripts", "EIMS lot.VG2")
SPF_INPUT_CSV = os.path.join(DATA_DIR, "EIMS_Inventory_filtered.csv")
SPF_RUN_DIR = os.path.join(OUTPUT_DIR, "spf_run")

# ---------------------------------------------------------------- InTMS config
INTMS_URL = "https://atmf.intel.com/api/custom/ims/ticketing/"
INTMS_ACTIONFLOW_URL = "https://atmf.intel.com/api/custom/ims/actionflow/"
INTMS_TICKET_URL = "https://atmf.intel.com/intms/panel/detail/%s"
INTMS_SIGNAL_ID = 221
INTMS_PRODUCTS_PATH = os.path.join(DATA_DIR, "intms_products.json")
INTMS_CUSTOMERS = ["CCG", "PSG", "DCG"]
INTMS_PAYER_BUS = ["ProdCo", "TMGf", "I'm not sure!"]
INTMS_PRODUCT_STAGES = ["HVM", "NPI", "Not product specific", "N/A"]
INTMS_TEST_AREAS = ["PPV", "Class", "Burn In", "OLB", "Others"]
INTMS_FACTORIES = ["VNAT", "CDAT", "CRAT", "KMAT8", "KuAT", "PGAT", "PG18"]
INTMS_DEFAULTS = {
    "Customer": "CCG",
    "Payer BU": "ProdCo",
    "Product Stage": "HVM",
    "Factories": "VNAT",
    "Shipping/LYA/FA Only": "No",
    "NFO shipment": "No",
}
INTMS_REQUESTS = ["Scrap these lot", "Transfer these lot to HVE Lab"]

CRITICAL_DAYS = 60

DEFAULT_FILTERS = {
    "Operation": ["7000"],
    "Mgr_Name": [
        "Bui Anh Huy", "Dong Dang Phu", "Le Thanh Loc",
        "Nguyen Dang Hai", "Nguyen Duc Loc", "Sky Tran",
    ],
    "Department": ["A/T MFG", "VNAT TEG-T"],
    "Prodgroup3": [
        "ADLN", "ARLR816B", "ARLS816B", "ARLU281", "ASL", "BMG21",
        "BTL12P", "GRR", "MTLU281", "NVLH", "NVLHX",
        "RPLP282", "RPLP282IOTG", "RPRP282", "TWL",
    ],
}

INTMS_PRODUCT_MAP = {
    "ADLN": "ADL N 0+8+1",
}


# ======================================================================= memory
# Everything the app learns from real submissions, shared by the whole team:
#   product_map : Prodgroup3 -> the ATMf Product Name actually used
#   submitted   : lot -> the ticket it went out on
# One small JSON file next to the data, rewritten atomically.
MEMORY_PATH = os.path.join(DATA_DIR, "atmf_memory.json")
_memory = None
_memory_lock = threading.Lock()


def _blank_memory():
    return {"product_map": {}, "submitted": {}}


def load_memory():
    global _memory
    with _memory_lock:
        if _memory is None:
            _memory = _blank_memory()
            if os.path.exists(MEMORY_PATH):
                try:
                    with open(MEMORY_PATH, "r", encoding="utf-8-sig") as f:
                        data = json.load(f)
                    if isinstance(data, dict):
                        _memory["product_map"] = dict(data.get("product_map") or {})
                        _memory["submitted"] = dict(data.get("submitted") or {})
                except (ValueError, OSError):
                    pass
        return _memory


def _save_memory_locked():
    """Write via a temp file so a crash can never leave a half-written JSON."""
    tmp = MEMORY_PATH + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(_memory, f, indent=2, ensure_ascii=False)
        os.replace(tmp, MEMORY_PATH)
    except OSError:
        pass


def remember_product_choice(prodgroup3, atmf_name):
    """Learn the ATMf Product Name a user picked for a Prodgroup3."""
    code = str(prodgroup3 or "").strip().upper()
    name = str(atmf_name or "").strip()
    if not code or not name:
        return
    load_memory()
    with _memory_lock:
        if _memory["product_map"].get(code) == name:
            return
        _memory["product_map"][code] = name
        _save_memory_locked()


def remember_submission(lots, info):
    """Record that these lots went out on a ticket."""
    load_memory()
    with _memory_lock:
        for lot in lots:
            _memory["submitted"][str(lot)] = info
        _save_memory_locked()


def submitted_map():
    return dict(load_memory()["submitted"])


def learned_product_map():
    return dict(load_memory()["product_map"])


def forget_submissions(lots):
    """Drop the submitted marker for these lots (e.g. a ticket was cancelled)."""
    load_memory()
    removed = 0
    with _memory_lock:
        for lot in lots:
            if _memory["submitted"].pop(str(lot), None) is not None:
                removed += 1
        if removed:
            _save_memory_locked()
    return removed


# ======================================================================= RUPS
_rups_cache = {}
_rups_lock = threading.Lock()
# Data is refreshed once a day (07:00 VN), so cache RUPS results for a full day.
# Everything served afterwards reuses the cache - no repeated RUPS queries.
_RUPS_TTL = 24 * 3600


def query_units(values):
    """Call RUPS API. Returns (status, records). In-process TTL cache (30 min)."""
    key = tuple(sorted(values))
    now = time.time()
    with _rups_lock:
        hit = _rups_cache.get(key)
        if hit and now - hit[0] < _RUPS_TTL:
            return hit[1]
    body = json.dumps({
        "api": "CUSTOMIZED_API_SEARCH_UNIT_INFO",
        "unit_list": ",".join(values),
    })
    res = requests.post(RUPS_URL, data=body, headers=RUPS_HEADER,
                        verify=False, timeout=240)
    records = []
    try:
        data = res.json().get("return", {}).get("data", [])
        for group in data:
            if isinstance(group, list):
                records.extend(i for i in group if isinstance(i, dict))
            elif isinstance(group, dict):
                records.append(group)
    except ValueError:
        pass
    out = (res.status_code, records)
    with _rups_lock:
        _rups_cache[key] = (now, out)
    return out


def clear_rups_cache():
    with _rups_lock:
        _rups_cache.clear()


# Cache full reconcile results per (scope, filtered-lot-set). Reused until the
# next daily refresh, so opening / re-filtering does not re-run RUPS.
_recon_cache = {}
_recon_lock = threading.Lock()

# Flat lot -> reconciled row index, filled by every reconciliation that runs.
# "Search by lot" reads from here first so it reuses the numbers already on
# screen instead of issuing a second RUPS query for the same lots.
_lot_rows = {}


def remember_lot_rows(rows):
    with _recon_lock:
        for r in rows:
            _lot_rows[r["lot"]] = r


def cached_lot_rows(lots):
    """Return (found_rows, missing_lots) for the requested lots."""
    found, missing = [], []
    with _recon_lock:
        for lot in lots:
            row = _lot_rows.get(lot)
            if row is None:
                missing.append(lot)
            else:
                found.append(dict(row))
    return found, missing


def clear_recon_cache():
    with _recon_lock:
        _recon_cache.clear()
        _lot_rows.clear()


def records_to_df(records):
    rows = []
    for rec in records:
        row = {}
        for k, v in rec.items():
            row[k] = json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else v
        rows.append(row)
    return pd.DataFrame(rows)


def df_to_excel_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="RUPS")
    return buf.getvalue()


# ======================================================================= EIMS
def download_eims():
    result = subprocess.run(
        ["curl.exe", "-s", "-S", "--negotiate", "-u", ":",
         "--fail", "-o", EIMS_TXT_PATH, EIMS_URL],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        raise RuntimeError("curl failed (exit %d): %s"
                           % (result.returncode, result.stderr.strip()))
    return os.path.getsize(EIMS_TXT_PATH)


def refresh_loss_operation(timeout=1800):
    """Run SQLPathFinder (CLI) to refresh data/Lot_loss_operation.csv from MARS.

    Writes the current EIMS report as the VG2 input CSV (LotNumber col 2), then
    runs SPF which connects to MARS itself (its own auth), queries
    F_LOT_HISTORY_V3 for those lots, and writes Lot_loss_operation.csv.
    Returns (ok, message). Safe to skip if SPF isn't installed.

    A normal run takes 2-4 minutes, but MARS can be much slower under load, so
    the default timeout is generous - this only ever runs in the background.
    """
    if not os.path.exists(SPF_EXE) or not os.path.exists(SPF_VG2):
        return False, "SQLPathFinder or VG2 not found - skipped."
    if not os.path.exists(EIMS_TXT_PATH):
        return False, "No EIMS file to feed SPF."
    if _spf_running():
        return False, "SQLPathFinder is already running - skipped."
    # 1) Materialise the SPF input CSV (comma-delimited, LotNumber as col 2).
    try:
        df = pd.read_csv(EIMS_TXT_PATH, sep="\t", dtype=str, keep_default_na=False)
        df.columns = [c.strip() for c in df.columns]
        df.to_csv(SPF_INPUT_CSV, index=False)
    except Exception as e:  # noqa: BLE001
        return False, "Failed to build SPF input CSV: %s" % e
    # 2) Run SPF CLI (headless-ish, minimized). It detaches; poll for completion.
    os.makedirs(SPF_RUN_DIR, exist_ok=True)
    out_before = os.path.getmtime(LOT_LOSS_PATH) if os.path.exists(LOT_LOSS_PATH) else 0
    try:
        subprocess.Popen([SPF_EXE, "/run", "/minimize", "/log",
                          "/startdir=%s" % SPF_RUN_DIR, SPF_VG2])
    except Exception as e:  # noqa: BLE001
        return False, "Failed to launch SPF: %s" % e
    # 3) Wait until the output CSV is (re)written and SPF exits.
    waited = 0
    while waited < timeout:
        time.sleep(5)
        waited += 5
        running = _spf_running()
        newer = (os.path.exists(LOT_LOSS_PATH)
                 and os.path.getmtime(LOT_LOSS_PATH) > out_before)
        if newer and not running:
            return True, "Loss-operation refreshed (%d bytes)." % os.path.getsize(LOT_LOSS_PATH)
    return False, "SPF timed out after %ds." % timeout


def _spf_running():
    try:
        r = subprocess.run(["tasklist", "/FI", "IMAGENAME eq sqlpathfinder3.exe"],
                           capture_output=True, text=True)
        return "sqlpathfinder3.exe" in (r.stdout or "")
    except Exception:  # noqa: BLE001
        return False



_eims_cache = {"mtime": None, "df": None}


def load_eims_df():
    if not os.path.exists(EIMS_TXT_PATH):
        return None
    mtime = os.path.getmtime(EIMS_TXT_PATH)
    if _eims_cache["mtime"] == mtime and _eims_cache["df"] is not None:
        return _eims_cache["df"]
    df = pd.read_csv(EIMS_TXT_PATH, sep="\t", dtype=str, keep_default_na=False)
    df.columns = [c.strip() for c in df.columns]
    _eims_cache["mtime"] = mtime
    _eims_cache["df"] = df
    return df


def eims_last_update():
    if os.path.exists(EIMS_TXT_PATH):
        return time.strftime("%Y-%m-%d %H:%M:%S",
                             time.localtime(os.path.getmtime(EIMS_TXT_PATH)))
    return "not loaded"


_map_cache = {}


def load_lose_map():
    if "lose" in _map_cache:
        return _map_cache["lose"]
    out = {}
    if os.path.exists(LOSE_MAP_PATH):
        m = pd.read_excel(LOSE_MAP_PATH, dtype=str).fillna("")
        m.columns = [c.strip() for c in m.columns]
        if "Operation" in m.columns and "PI Dispose" in m.columns:
            out = {str(op).strip(): str(pi).strip()
                   for op, pi in zip(m["Operation"], m["PI Dispose"])
                   if str(op).strip()}
    _map_cache["lose"] = out
    return out


def load_inqty_map():
    if "inqty" in _map_cache:
        return _map_cache["inqty"]
    inqty, cd3 = {}, {}
    if os.path.exists(LOT_LOSS_PATH):
        m = pd.read_csv(LOT_LOSS_PATH, dtype=str, keep_default_na=False)
        m.columns = [c.strip().upper() for c in m.columns]
        if "LOT" in m.columns:
            # MARS returns the same row once per source node, so drop exact
            # duplicates before aggregating or INQTY gets multiplied.
            m = m.drop_duplicates()
            lot = m["LOT"].astype(str).str.strip()
            if "INQTY" in m.columns:
                qty = pd.to_numeric(m["INQTY"], errors="coerce").fillna(0)
                inqty = qty.groupby(lot).sum().astype(int).to_dict()
            if "CREATE_DATA3" in m.columns:
                # A lot has several rows and only some carry the loss
                # operation; keep the first NON-BLANK value per lot.
                vals = m["CREATE_DATA3"].astype(str).str.strip()
                for lt, v in zip(lot, vals):
                    if v and not cd3.get(lt):
                        cd3[lt] = v
    _map_cache["inqty"] = (inqty, cd3)
    return inqty, cd3


def load_intms_products():
    if "products" in _map_cache:
        return _map_cache["products"]
    out = []
    if os.path.exists(INTMS_PRODUCTS_PATH):
        try:
            with open(INTMS_PRODUCTS_PATH, "r", encoding="utf-8-sig") as f:
                data = json.load(f)
            out = [str(x) for x in data] if isinstance(data, list) else []
        except (ValueError, OSError):
            out = []
    _map_cache["products"] = out
    return out


def clear_caches():
    _eims_cache["mtime"] = None
    _eims_cache["df"] = None
    _map_cache.clear()
    clear_rups_cache()
    clear_recon_cache()


# ======================================================================= helpers
def strip_via(value):
    s = str(value).strip()
    if not s:
        return ""
    out = []
    for part in s.split(","):
        p = part.strip()
        if not p:
            continue
        up = p.upper()
        idx = up.find(" VIA ")
        if idx != -1:
            p = p[:idx].strip()
        elif up.endswith(" VIA"):
            p = p[:-4].strip()
        if p and p not in out:
            out.append(p)
    return ", ".join(out)


def cd3_map_for(lot):
    """MARS CREATE_DATA3 (the loss operation) for one lot, or ""."""
    _, cd3 = load_inqty_map()
    return cd3.get(lot, "")


def map_pi_dispose(lose_value, lose_map):
    if not lose_value or not lose_map:
        if lose_value and "7226" in str(lose_value):
            return "Eng_Assessment"
        return ""
    parts = [p.strip() for p in str(lose_value).split(",") if p.strip()]
    labels = []
    for p in parts:
        lab = "Eng_Assessment" if "7226" in p else lose_map.get(p, "")
        if lab and lab not in labels:
            labels.append(lab)
    return ", ".join(labels)


def product_has_sspec(product):
    if not isinstance(product, str):
        return False
    tokens = product.split()
    return bool(tokens) and len(tokens[-1]) > 1


def guess_atmf_product(prodgroup3, products, rups_product=""):
    if not products:
        return ""
    code = str(prodgroup3 or "").strip().upper()

    # What a human actually picked last time always wins over any guessing.
    learned = learned_product_map().get(code)
    if learned and learned in products:
        return learned

    if code in INTMS_PRODUCT_MAP and INTMS_PRODUCT_MAP[code] in products:
        return INTMS_PRODUCT_MAP[code]

    def _norm(s):
        return "".join(ch for ch in str(s).upper() if ch.isalnum())

    targets = {p: _norm(p) for p in products}
    for src in (code, rups_product):
        n = _norm(src)
        if not n:
            continue
        hits = [p for p, t in targets.items() if t.startswith(n) or n in t]
        if hits:
            return min(hits, key=len)
        close = difflib.get_close_matches(n, list(targets.values()), n=1, cutoff=0.6)
        if close:
            for p, t in targets.items():
                if t == close[0]:
                    return p
    return ""


def slugify(name):
    return "".join(ch if ch.isalnum() else "-" for ch in str(name)).strip("-") or "none"


# ======================================================================= filter
def filter_options(df):
    """Return {col: sorted unique values} for the filter widgets."""
    opts = {}
    cols = ["Prodgroup3"] + [c for c in EIMS_FILTER_COLUMNS if c != "Prodgroup3"]
    for c in cols:
        if c in df.columns:
            opts[c] = sorted(v for v in df[c].unique() if str(v).strip())
    return opts


def apply_filters(df, products, selected):
    """products: list of Prodgroup3; selected: {col: [values]}."""
    mask = df["Prodgroup3"].isin(products) if products else pd.Series(True, index=df.index)
    for col, vals in (selected or {}).items():
        if vals and col in df.columns:
            mask = mask & df[col].isin(vals)
    return df[mask]


def days_numeric(df):
    if "Days_At_Operation" in df.columns:
        return pd.to_numeric(df["Days_At_Operation"], errors="coerce")
    return pd.Series(index=df.index, dtype="float64")


def product_summary(filtered):
    """Return list of per-product summary dicts."""
    rows = []
    if "Prodgroup3" not in filtered.columns or not len(filtered):
        return rows
    lotcol = EIMS_LOT_COLUMN if EIMS_LOT_COLUMN in filtered.columns else None
    for prod, grp in filtered.groupby("Prodgroup3"):
        if not str(prod).strip():
            continue
        n_lots = grp[lotcol].nunique() if lotcol else len(grp)
        crit = grp[days_numeric(grp) > CRITICAL_DAYS]
        n_crit = crit[lotcol].nunique() if lotcol else len(crit)
        qty = int(pd.to_numeric(grp.get("Quantity"), errors="coerce").fillna(0).sum()) \
            if "Quantity" in grp.columns else 0
        cat = grp["Product"].map(
            lambda p: "PPV" if product_has_sspec(p) else "Class") \
            if "Product" in grp.columns else pd.Series("Class", index=grp.index)
        ppv = grp[cat.values == "PPV"]
        cls = grp[cat.values == "Class"]
        n_ppv = ppv[lotcol].nunique() if lotcol else len(ppv)
        n_cls = cls[lotcol].nunique() if lotcol else len(cls)
        rows.append({
            "product": prod,
            "slug": slugify(prod),
            "lots": int(n_lots),
            "unit": qty,
            "ppv_lots": int(n_ppv),
            "class_lots": int(n_cls),
            "critical": int(n_crit),
        })
    rows.sort(key=lambda r: r["critical"], reverse=True)
    return rows


# ======================================================================= reconcile
def reconcile(filtered, scope="all"):
    """Cached reconcile. Result is reused (per scope + lot set) until the next
    daily refresh, so re-opening or re-filtering does NOT re-query RUPS."""
    lotcol = EIMS_LOT_COLUMN
    scope_df = filtered
    if scope == "critical":
        scope_df = filtered[days_numeric(filtered) > CRITICAL_DAYS]
    lots = tuple(sorted(v for v in scope_df[lotcol].unique() if str(v).strip())) \
        if lotcol in scope_df.columns else ()
    key = (scope, lots)
    with _recon_lock:
        hit = _recon_cache.get(key)
        if hit is not None:
            return hit
    result = _reconcile_impl(filtered, scope=scope)
    with _recon_lock:
        _recon_cache[key] = result
    return result


def _reconcile_impl(filtered, scope="all"):
    """Run RUPS for the filtered lots and return a structured reconciliation.

    scope: 'all' or 'critical'. Returns dict with metrics + per-product groups.
    """
    lotcol = EIMS_LOT_COLUMN
    scope_df = filtered
    if scope == "critical":
        scope_df = filtered[days_numeric(filtered) > CRITICAL_DAYS]

    lots = sorted(v for v in scope_df[lotcol].unique() if str(v).strip()) \
        if lotcol in scope_df.columns else []
    if not lots:
        return {"ok": True, "lots": 0, "records": 0, "metrics": {},
                "products": [], "status": 200}

    # EIMS qty + product per lot.
    eims_qty, eims_prod, eims_pid = {}, {}, {}
    if "Quantity" in scope_df.columns:
        q = pd.to_numeric(scope_df["Quantity"], errors="coerce").fillna(0)
        eims_qty = q.groupby(scope_df[lotcol]).sum().astype(int).to_dict()
    if "Prodgroup3" in scope_df.columns:
        for lot, grp in scope_df.groupby(lotcol):
            prods = sorted({str(v).strip() for v in grp["Prodgroup3"]
                            if str(v).strip() and str(v).strip().lower() != "nan"})
            eims_prod[lot] = ", ".join(prods)
    # The full EIMS "Product" string (shown as Product ID in the UI). EIMS pads
    # it with runs of spaces, so collapse them to keep the column readable.
    if "Product" in scope_df.columns:
        for lot, grp in scope_df.groupby(lotcol):
            pids = sorted({" ".join(str(v).split()) for v in grp["Product"]
                           if str(v).strip() and str(v).strip().lower() != "nan"})
            eims_pid[lot] = ", ".join(pids)

    status, records = query_units(lots)
    if not records:
        # Still index these lots (as EIMS-only rows) so "search by lot" does not
        # re-run the same empty RUPS query every single time.
        blank = [{
            "lot": lot,
            "product": eims_prod.get(lot, ""),
            "product_id": eims_pid.get(lot, ""),
            "eims_qty": int(eims_qty[lot]) if lot in eims_qty else None,
            "rups_qty": 0,
            "last_used": None,
            "operation": strip_via(cd3_map_for(lot)),
            "op_conflict": False,
            "pi_dispose": map_pi_dispose(strip_via(cd3_map_for(lot)), load_lose_map()),
            "status": "Not found",
        } for lot in lots]
        remember_lot_rows(blank)
        return {"ok": True, "lots": len(lots), "records": 0,
                "status": status, "metrics": {}, "products": group_by_product(pd.DataFrame(blank)),
                "summary": blank, "not_found": lots,
                "message": "No RUPS records for the %d searched lot(s)." % len(lots)}

    df = records_to_df(records)
    lot_series = df["EIMS_LOT"].astype(str) if "EIMS_LOT" in df.columns \
        else pd.Series("", index=df.index)
    found = set(lot_series)
    not_found = [v for v in lots if v not in found]

    if "VID" in df.columns:
        rups_qty = df.groupby(lot_series)["VID"].nunique().to_dict()
    else:
        rups_qty = lot_series.value_counts().to_dict()

    lose_by_lot = {}
    if "LOSE_OPERATION" in df.columns:
        for lot, grp in df.groupby(lot_series):
            vals = sorted({str(v).strip() for v in grp["LOSE_OPERATION"]
                           if str(v).strip() and str(v).strip().lower() != "nan"})
            lose_by_lot[lot] = ", ".join(vals)

    pi_lookup = load_lose_map()
    last_edit = {}
    if "LAST_EDIT_DATE" in df.columns:
        led = pd.to_datetime(df["LAST_EDIT_DATE"], errors="coerce")
        last_edit = led.groupby(lot_series).max().to_dict()
    now = pd.Timestamp.now()
    inqty_map, cd3_map = load_inqty_map()

    rows = []
    for lot in lots:
        rq = int(rups_qty.get(lot, 0))
        eq = eims_qty.get(lot)
        if eq is None:
            match = "N/A" if rq else "Not found"
        else:
            match = "Match" if int(eq) == rq else "MISMATCH"
        lose_s = strip_via(lose_by_lot.get(lot, ""))
        cd3_s = strip_via(cd3_map.get(lot, ""))
        if lose_s and cd3_s:
            op_val, op_conflict = lose_s, (lose_s != cd3_s)
        elif lose_s:
            op_val, op_conflict = lose_s, False
        else:
            op_val, op_conflict = cd3_s, False
        pi_val = map_pi_dispose(op_val, pi_lookup)
        last_used = None
        if match == "Match" and pi_val in ("PPV", "Class"):
            ts = last_edit.get(lot)
            if ts is not None and pd.notna(ts):
                last_used = round((now - ts).total_seconds() / 86400.0, 1)
        rows.append({
            "lot": lot,
            "product": eims_prod.get(lot, ""),
            "product_id": eims_pid.get(lot, ""),
            "eims_qty": int(eq) if eq is not None else None,
            "rups_qty": rq,
            "last_used": last_used,
            "operation": op_val,
            "op_conflict": bool(op_conflict),
            "pi_dispose": pi_val,
            "status": match,
        })

    summary = pd.DataFrame(rows)
    n_mis = int((summary["status"] == "MISMATCH").sum())
    sspec = df["SSPEC"].astype(str).str.strip() if "SSPEC" in df.columns \
        else pd.Series("", index=df.index)

    # Feed the lot index so "search by lot" can reuse these numbers.
    remember_lot_rows(rows)

    products = group_by_product(summary)

    return {
        "ok": True,
        "lots": len(lots),
        "records": len(df),
        "status": status,
        "not_found": not_found,
        "metrics": {
            "ppv_units": int((sspec != "").sum()),
            "class_units": int((sspec == "").sum()),
            "not_found": len(not_found),
            "mismatch": n_mis,
        },
        "products": products,
        "summary": summary.to_dict("records"),
    }


def group_by_product(summary):
    """Shape a summary frame into products[] -> groups[] (by PI) -> lots[].

    Shared by the reconciliation view and the lot search so both render with
    exactly the same component and selection behaviour.
    """
    if summary is None or not len(summary):
        return []
    status_order = {"Match": 0, "N/A": 1, "Not found": 2, "MISMATCH": 3}
    pi_order = {"PPV": 0, "Class": 1, "Eng_Assessment": 2}
    products = []
    for prod in sorted(summary["product"].unique(), key=lambda x: (x == "", x)):
        sub = summary[summary["product"] == prod]
        groups = []
        for pi in sorted(sub["pi_dispose"].unique(),
                         key=lambda x: (pi_order.get(x, 8), x == "", x)):
            psub = sub[sub["pi_dispose"] == pi].copy()
            psub["_ord"] = psub["status"].map(lambda s: status_order.get(s, 9))
            psub["_lu"] = pd.to_numeric(psub["last_used"], errors="coerce")
            psub = psub.sort_values(["_ord", "_lu"], ascending=[True, False],
                                    na_position="last")
            groups.append({
                "pi": pi or "Other",
                "lots": psub.drop(columns=["_ord", "_lu"]).to_dict("records"),
                "mismatch": int((psub["status"] == "MISMATCH").sum()),
            })
        products.append({
            "product": prod or "(no product)",
            "slug": slugify(prod or "no-product"),
            "n_lots": len(sub),
            "n_ppv": int((sub["pi_dispose"] == "PPV").sum()),
            "n_class": int((sub["pi_dispose"] == "Class").sum()),
            "n_eng": int((sub["pi_dispose"] == "Eng_Assessment").sum()),
            "mismatch": int((sub["status"] == "MISMATCH").sum()),
            "groups": groups,
        })
    return products


# ======================================================================= lot search
# Free-form lookup for any lot, independent of the current filters. The result
# is shaped exactly like the reconciliation (products -> PI groups -> lots) so
# the same table component renders it and lots stay selectable for ATMf.

# A single search is interactive, so keep it small and fast.
MAX_SEARCH_LOTS = 200
MAX_SUGGESTIONS = 8


def parse_lot_input(text):
    """Split a free-form lot list into clean tokens.

    Accepts anything a user is likely to paste: commas, semicolons, tabs,
    newlines or plain spaces (so an Excel column pastes straight in).
    Order is preserved and duplicates are dropped.
    """
    if not text:
        return []
    out, seen = [], set()
    for raw in str(text).replace(",", " ").replace(";", " ").split():
        tok = raw.strip().strip("\"'")
        if not tok:
            continue
        key = tok.upper()
        if key not in seen:
            seen.add(key)
            out.append(tok)
    return out


def _eims_lot_index():
    """Map UPPER(lot) -> real lot string, for case-insensitive lookups."""
    df = load_eims_df()
    if df is None or EIMS_LOT_COLUMN not in df.columns:
        return {}
    lots = df[EIMS_LOT_COLUMN].astype(str).str.strip()
    return {v.upper(): v for v in lots.unique() if v}


def resolve_lots(tokens):
    """Turn user tokens into real lot numbers.

    An exact (case-insensitive) hit wins. Otherwise the token is treated as a
    prefix/substring so "ADLN" expands to every matching EIMS lot. Tokens that
    match nothing are kept anyway (the lot may have left inventory already).
    """
    index = _eims_lot_index()
    resolved, unknown, expanded = [], [], {}
    seen = set()

    def _add(lot):
        if lot and lot not in seen:
            seen.add(lot)
            resolved.append(lot)

    for tok in tokens:
        up = tok.upper()
        if up in index:
            _add(index[up])
            continue
        hits = [v for k, v in index.items() if k.startswith(up)]
        if not hits:
            hits = [v for k, v in index.items() if up in k]
        if hits:
            hits = sorted(hits)[:MAX_SEARCH_LOTS]
            expanded[tok] = len(hits)
            for h in hits:
                _add(h)
        else:
            _add(tok)
            unknown.append(tok)
    return resolved[:MAX_SEARCH_LOTS], unknown, expanded


def suggest_lots(token, limit=MAX_SUGGESTIONS):
    """Close EIMS lot numbers for a token that matched nothing."""
    index = _eims_lot_index()
    if not index:
        return []
    near = difflib.get_close_matches(str(token).upper(), list(index),
                                     n=limit, cutoff=0.6)
    return [index[k] for k in near]


def search_lots(text):
    """Look up lots and return them grouped by product -> PI dispose.

    Reads from the reconciliation the app has already computed (the same rows
    shown under "RUPS reconciliation"), so a search costs nothing and always
    agrees with what is on screen. Only lots that have never been reconciled
    fall through to a small, targeted reconciliation of their own.
    """
    tokens = parse_lot_input(text)
    if not tokens:
        return {"ok": True, "products": [], "summary": [], "query": [],
                "message": "Enter one or more lot numbers."}

    wanted, unknown, expanded = resolve_lots(tokens)
    if not wanted:
        return {"ok": True, "products": [], "summary": [], "query": tokens,
                "message": "No lots matched."}

    rows, missing = cached_lot_rows(wanted)
    reused = len(rows)

    # Anything not reconciled yet: reconcile just those lots. That reuses the
    # exact same code path as the main view and fills the index for next time.
    if missing:
        df = load_eims_df()
        if df is not None and EIMS_LOT_COLUMN in df.columns:
            lot_col = df[EIMS_LOT_COLUMN].astype(str).str.strip()
            slice_df = df[lot_col.isin(missing)]
            if len(slice_df):
                reconcile(slice_df, scope="all")
        found, still_missing = cached_lot_rows(missing)
        rows.extend(found)
        # Lots that exist nowhere - keep them visible so the user sees the gap.
        for lot in still_missing:
            rows.append({
                "lot": lot, "product": "", "product_id": "",
                "eims_qty": None, "rups_qty": 0, "last_used": None,
                "operation": "", "op_conflict": False, "pi_dispose": "",
                "status": "Not found",
            })

    # Preserve the order the user asked for.
    order = {lot: i for i, lot in enumerate(wanted)}
    rows.sort(key=lambda r: order.get(r["lot"], 10 ** 6))

    summary = pd.DataFrame(rows)
    products = group_by_product(summary)
    not_found = [r["lot"] for r in rows if r["status"] == "Not found"]

    # Hints for terms that matched nothing anywhere.
    suggestions = {}
    for tok in unknown:
        row = next((r for r in rows if r["lot"] == tok), None)
        if row and row["status"] == "Not found":
            near = suggest_lots(tok)
            if near:
                suggestions[tok] = near

    return {
        "ok": True,
        "query": tokens,
        "expanded": expanded,
        "suggestions": suggestions,
        "lots": len(rows),
        "reused": reused,
        "fetched": len(rows) - reused,
        "not_found": not_found,
        "truncated": len(wanted) >= MAX_SEARCH_LOTS,
        "products": products,
        "summary": summary.to_dict("records"),
    }
# ======================================================================= ATMf
def submit_intms_ticket(signal_id, form_json, timeout=60):
    payload = {"signal": int(signal_id), "form_json": form_json}
    body_path = os.path.join(OUTPUT_DIR, "_intms_body.json")
    with open(body_path, "w", encoding="utf-8") as f:
        json.dump(payload, f)
    cmd = [
        "curl.exe", "-s", "-k", "--negotiate", "-u", ":",
        "-H", "Content-Type: application/json",
        "-X", "POST", INTMS_URL,
        "--data-binary", "@%s" % body_path,
        "--max-time", str(timeout),
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    raw = (proc.stdout or "").strip()
    if not raw:
        return False, None, (proc.stderr or "").strip() or "empty response"
    try:
        data = json.loads(raw)
    except ValueError:
        return False, None, raw[:500]
    if isinstance(data, dict):
        tid = data.get("id")
        if tid:
            return True, tid, data
        return False, None, json.dumps(data)[:500]
    return False, None, json.dumps(data)[:500]


def fill_intms_action(ticket_id, step_id, rich_text, timeout=60):
    payload = {"action_flow_json": {"id": str(step_id), "rich_text": rich_text}}
    body_path = os.path.join(OUTPUT_DIR, "_intms_patch.json")
    with open(body_path, "w", encoding="utf-8") as f:
        json.dump(payload, f)
    cmd = [
        "curl.exe", "-s", "-k", "--negotiate", "-u", ":",
        "-H", "Content-Type: application/json",
        "-X", "PATCH", "%s%s/" % (INTMS_ACTIONFLOW_URL, ticket_id),
        "--data-binary", "@%s" % body_path,
        "--max-time", str(timeout),
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    raw = (proc.stdout or "").strip()
    if not raw:
        return False, (proc.stderr or "").strip() or "empty response"
    try:
        data = json.loads(raw)
    except ValueError:
        return False, raw[:300]
    return (isinstance(data, dict) and "id" in data), json.dumps(data)[:300]


def submit_tickets(tickets):
    """tickets: list of {product, test_area, request, lots:[...], form overrides}.

    Creates one ATMf ticket per entry and PATCHes the fill-in step. Returns
    a list of result dicts.
    """
    results = []
    for t in tickets:
        form = {
            "Customer": t.get("customer", INTMS_DEFAULTS["Customer"]),
            "Payer BU": t.get("payer", INTMS_DEFAULTS["Payer BU"]),
            "Product Stage": t.get("stage", INTMS_DEFAULTS["Product Stage"]),
            "Product Name": t["product_name"],
            "Test Area": t.get("test_area", "Class"),
            "Factories": t.get("factory", INTMS_DEFAULTS["Factories"]),
            "Shipping/LYA/FA Only": t.get("shipping", INTMS_DEFAULTS["Shipping/LYA/FA Only"]),
            "NFO shipment": t.get("nfo", INTMS_DEFAULTS["NFO shipment"]),
        }
        ok, tid, detail = submit_intms_ticket(INTMS_SIGNAL_ID, form)
        filled = None
        if ok and tid:
            lots_html = "".join("<p>%s</p>" % l for l in t["lots"])
            rich = "<p>%s</p>%s" % (t.get("request", INTMS_REQUESTS[0]), lots_html)
            filled, _ = fill_intms_action(tid, 4, rich)
            # Learn from a real submission: the product mapping that was used,
            # and which lots are now spoken for.
            remember_product_choice(t.get("product"), t["product_name"])
            remember_submission(t["lots"], {
                "ticket_id": tid,
                "url": INTMS_TICKET_URL % tid,
                "product_name": t["product_name"],
                "test_area": t.get("test_area", "Class"),
                "request": t.get("request", INTMS_REQUESTS[0]),
                "at": time.strftime("%Y-%m-%d %H:%M"),
            })
        results.append({
            "lots": t["lots"],
            "product_name": t["product_name"],
            "test_area": t.get("test_area", "Class"),
            "ok": bool(ok and tid),
            "ticket_id": tid,
            "url": INTMS_TICKET_URL % tid if tid else "",
            "filled": bool(filled),
            "detail": "" if (ok and tid) else str(detail),
        })
    return results

