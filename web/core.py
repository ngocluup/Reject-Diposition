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


# ======================================================================= RUPS
_rups_cache = {}
_rups_lock = threading.Lock()
# Data is refreshed once a day (07:00 VN), so cache RUPS results for a full day.
# Everything served afterwards reuses the cache — no repeated RUPS queries.
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


def clear_recon_cache():
    with _recon_lock:
        _recon_cache.clear()


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


def refresh_loss_operation(timeout=600):
    """Run SQLPathFinder (CLI) to refresh data/Lot_loss_operation.csv from MARS.

    Writes the current EIMS report as the VG2 input CSV (LotNumber col 2), then
    runs SPF which connects to MARS itself (its own auth), queries
    F_LOT_HISTORY_V3 for those lots, and writes Lot_loss_operation.csv.
    Returns (ok, message). Safe to skip if SPF isn't installed.
    """
    if not os.path.exists(SPF_EXE) or not os.path.exists(SPF_VG2):
        return False, "SQLPathFinder or VG2 not found — skipped."
    if not os.path.exists(EIMS_TXT_PATH):
        return False, "No EIMS file to feed SPF."
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
    eims_qty, eims_prod = {}, {}
    if "Quantity" in scope_df.columns:
        q = pd.to_numeric(scope_df["Quantity"], errors="coerce").fillna(0)
        eims_qty = q.groupby(scope_df[lotcol]).sum().astype(int).to_dict()
    if "Prodgroup3" in scope_df.columns:
        for lot, grp in scope_df.groupby(lotcol):
            prods = sorted({str(v).strip() for v in grp["Prodgroup3"]
                            if str(v).strip() and str(v).strip().lower() != "nan"})
            eims_prod[lot] = ", ".join(prods)

    status, records = query_units(lots)
    if not records:
        return {"ok": True, "lots": len(lots), "records": 0,
                "status": status, "metrics": {}, "products": [],
                "summary": [], "not_found": lots,
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

    # Group by product -> PI group.
    status_order = {"Match": 0, "N/A": 1, "Not found": 2, "MISMATCH": 3}
    products = []
    for prod in sorted(summary["product"].unique(), key=lambda x: (x == "", x)):
        sub = summary[summary["product"] == prod]
        groups = []
        pi_order = {"PPV": 0, "Class": 1, "Eng_Assessment": 2}
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


# ======================================================================= email
def build_report(summary_records):
    """Light-theme HTML email from reconciliation summary records."""
    summary = pd.DataFrame(summary_records)
    total = len(summary)
    n_match = int((summary["status"] == "Match").sum()) if total else 0
    n_mis = int((summary["status"] == "MISMATCH").sum()) if total else 0
    prods = sorted({p for p in summary.get("product", []) if str(p).strip()}) \
        if total else []
    subject = "EIMS <-> RUPS Lot Reconciliation - %s (%d lots)" % (
        time.strftime("%Y-%m-%d"), total)

    BG, CARD, BORDER = "#ffffff", "#f6f8fa", "#d0d7de"
    TEXT, MUTED, HEAD_BG = "#1f2328", "#656d76", "#eaeef2"
    RED = "background:#ffebe9;color:#cf222e;font-weight:700"
    order = {"Match": 0, "N/A": 1, "Not found": 2, "MISMATCH": 3}

    def _table(frame):
        frame = frame.sort_values(
            "status", key=lambda s: s.map(lambda v: order.get(v, 9)), kind="stable")
        cols = [("lot", "Lot"), ("eims_qty", "EIMS_Qty"), ("rups_qty", "RUPS_Qty"),
                ("last_used", "Last_Used_Days"), ("operation", "Operation"),
                ("pi_dispose", "PI_Dispose")]
        cols = [(k, h) for k, h in cols if k in frame.columns]
        head = "".join(
            "<th style='border:1px solid %s;padding:6px 10px;background:%s;color:%s;"
            "text-align:left'>%s</th>" % (BORDER, HEAD_BG, TEXT, h) for _, h in cols)
        body = ""
        for _, r in frame.iterrows():
            mis = str(r.get("status", "")) == "MISMATCH"
            cells = ""
            for k, _h in cols:
                val = r[k]
                val = "" if val is None or (isinstance(val, float) and pd.isna(val)) else val
                extra = RED if (mis and k in ("eims_qty", "rups_qty")) else ("color:%s" % TEXT)
                cells += "<td style='border:1px solid %s;padding:6px 10px;%s'>%s</td>" \
                    % (BORDER, extra, val)
            body += "<tr>%s</tr>" % cells
        return ("<table style='border-collapse:collapse;font-family:Segoe UI,Arial;"
                "font-size:12px;margin:6px 0;width:100%%'><thead><tr>%s</tr></thead>"
                "<tbody>%s</tbody></table>" % (head, body))

    parts = ["<div style='background:%s;padding:20px;font-family:Segoe UI,Arial;"
             "font-size:14px;color:%s'>" % (BG, TEXT)]
    parts.append(
        "<div style='background:linear-gradient(120deg,#0969da,#0a66c2 55%%,#1a7f37);"
        "padding:20px 24px;border-radius:16px;color:#fff;margin-bottom:14px'>"
        "<div style='font-size:22px;font-weight:800'>RUPS + EIMS inventory reconciliation</div>"
        "<div style='font-size:13px;color:#d6e4ff;margin-top:4px'>Report generated %s</div></div>"
        % time.strftime("%Y-%m-%d %H:%M"))
    parts.append("<p>Hi all,</p>")
    parts.append("<div style='background:%s;border:1px solid %s;border-radius:10px;"
                 "padding:12px 16px;margin:8px 0'><b>Summary</b><ul style='margin:6px 0'>"
                 % (CARD, BORDER))
    parts.append("<li>Products: <b>%s</b></li>" % (", ".join(prods) if prods else "-"))
    parts.append("<li>Total lots: <b>%d</b></li>" % total)
    parts.append("<li style='color:#1a7f37'>Matched: <b>%d</b></li>" % n_match)
    parts.append("<li style='color:#cf222e'>Quantity mismatches: <b>%d</b></li>" % n_mis)
    parts.append("</ul></div>")

    if total and "product" in summary.columns:
        for grp in sorted(summary["product"].unique(), key=lambda x: (x == "", x)):
            sub = summary[summary["product"] == grp]
            label = grp if str(grp).strip() else "(no product)"
            g_mis = int((sub["status"] == "MISMATCH").sum())
            accent = "#cf222e" if g_mis else "#0969da"
            head = ("<span>&#128230; %s</span> <span style='color:%s;font-weight:400'>"
                    "&mdash; %d lot(s)</span>" % (label, MUTED, len(sub)))
            if g_mis:
                head += " <span style='color:#cf222e'>&middot; %d mismatch</span>" % g_mis
            parts.append(
                "<div style='border:1px solid %s;border-left:4px solid %s;"
                "border-radius:12px;padding:12px 16px;margin:14px 0;background:%s'>"
                "<div style='font-size:16px;font-weight:800;margin-bottom:4px'>%s</div>%s</div>"
                % (BORDER, accent, CARD, head, _table(sub)))
    parts.append("<p style='color:%s'>Thanks,</p></div>" % MUTED)
    return subject, "".join(parts)


def send_via_outlook(subject, html_body, to="", cc="", attachment_path=None,
                     display_only=True):
    ps = r"""
$ol = New-Object -ComObject Outlook.Application
$mail = $ol.CreateItem(0)
$mail.Subject = $env:MAIL_SUBJECT
$mail.To = $env:MAIL_TO
$mail.CC = $env:MAIL_CC
$mail.HTMLBody = [System.IO.File]::ReadAllText($env:MAIL_BODY_FILE)
if ($env:MAIL_ATTACH -and (Test-Path $env:MAIL_ATTACH)) { $mail.Attachments.Add($env:MAIL_ATTACH) | Out-Null }
"""
    ps += "$mail.Display()\n" if display_only else "$mail.Send()\n"
    body_file = os.path.join(OUTPUT_DIR, "_email_body.html")
    with open(body_file, "w", encoding="utf-8") as f:
        f.write(html_body)
    ps_file = os.path.join(OUTPUT_DIR, "_send_outlook.ps1")
    with open(ps_file, "w", encoding="utf-8") as f:
        f.write(ps)
    env = dict(os.environ)
    env.update({
        "MAIL_SUBJECT": subject, "MAIL_TO": to, "MAIL_CC": cc,
        "MAIL_BODY_FILE": body_file, "MAIL_ATTACH": attachment_path or "",
    })
    result = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps_file],
        capture_output=True, text=True, env=env)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "Outlook COM failed")
