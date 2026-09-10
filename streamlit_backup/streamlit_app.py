"""
Streamlit app: RUPS lookup + EIMS Inventory search.

Two parts (tabs):
  1. Manual RUPS search  -> paste a Lot list (and/or Visual IDs), search RUPS, view + download.
  2. EIMS Inventory      -> load/refresh EIMS data, filter it, pick lots, auto-search RUPS.

Run:
    C:\\Users\\ngocluup\\AppData\\Local\\miniforge3\\envs\\ngocluup\\python.exe -m streamlit run streamlit_app.py

Requires: streamlit, pandas, requests, openpyxl.
Install (conda-forge works on this machine; PyPI/pip times out):
    conda install -n ngocluup -y -c conda-forge streamlit pandas requests openpyxl
"""
import io
import json
import os
import subprocess
import time

import difflib

import pandas as pd
import requests
import streamlit as st


try:
    requests.packages.urllib3.disable_warnings()  # type: ignore[attr-defined]
except Exception:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Backup copy lives in streamlit_backup/; share data/output with the project root.
_ROOT = os.path.dirname(BASE_DIR)
DATA_DIR = os.path.join(_ROOT, "data")
OUTPUT_DIR = os.path.join(_ROOT, "output")
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---------------------------------------------------------------- secrets
# Credentials come from env vars or the git-ignored config.local.json in the
# project root. Never hard-code them here.
_CONFIG_PATH = os.path.join(_ROOT, "config.local.json")
_LOCAL_CONFIG = {}
if os.path.exists(_CONFIG_PATH):
    try:
        with open(_CONFIG_PATH, "r", encoding="utf-8-sig") as _f:
            _LOCAL_CONFIG = json.load(_f) or {}
    except (ValueError, OSError):
        _LOCAL_CONFIG = {}


def _secret(name, default=""):
    return os.environ.get(name) or _LOCAL_CONFIG.get(name) or default


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

# LOSE_OPERATION -> PI Dispose mapping file (Operation | Domain | PI Dispose).
LOSE_MAP_PATH = os.path.join(DATA_DIR, "LOSE_OPERATION MAPPING.xlsx")

# Lot loss-operation data (from AQUA/SQLPathFinder): LOT, CREATE_DATA1..4, INQTY.
LOT_LOSS_PATH = os.path.join(DATA_DIR, "Lot_loss_operation.csv")

# ---------------------------------------------------------------- InTMS config
# ATMf ticketing API (InTMS). Auth: Kerberos SSO via curl --negotiate (like EIMS).
# NOTE: the live ATMf instance is atmf.intel.com (intms.intel.com is a different DB).
INTMS_URL = "https://atmf.intel.com/api/custom/ims/ticketing/"
INTMS_ACTIONFLOW_URL = "https://atmf.intel.com/api/custom/ims/actionflow/"
INTMS_TICKET_URL = "https://atmf.intel.com/intms/panel/detail/%s"
# Signal 221 = "Reject Management - POR" (ticketing_type 3). Verified from an
# existing ATMf ticket. This is the ticket template new tickets are created under.
INTMS_SIGNAL_ID = 221
INTMS_PRODUCTS_PATH = os.path.join(DATA_DIR, "intms_products.json")
# Valid option lists for the signal-221 form (from its form_fields_json).
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


# ======================================================================= RUPS
@st.cache_data(show_spinner=False, ttl=1800)
def _query_units_cached(values_tuple):
    """Cached RUPS call keyed on a hashable tuple of values."""
    body = json.dumps({
        "api": "CUSTOMIZED_API_SEARCH_UNIT_INFO",
        "unit_list": ",".join(values_tuple),
    })
    res = requests.post(RUPS_URL, data=body, headers=RUPS_HEADER, verify=False, timeout=120)
    records = []
    try:
        data = res.json().get("return", {}).get("data", [])
        for group in data:
            if isinstance(group, list):
                records.extend(item for item in group if isinstance(item, dict))
            elif isinstance(group, dict):
                records.append(group)
    except ValueError:
        pass
    return res.status_code, records


def query_units(values):
    """Call RUPS API with a list of lots / visual IDs, return (status, records).

    Cached (30 min) on the value set so re-runs / filter tweaks don't re-hit the
    API unless the lot set changes. Use the Refresh button to clear.
    """
    return _query_units_cached(tuple(values))


def records_to_df(records):
    """Flatten records into a DataFrame (nested values become JSON strings)."""
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
    """Download the EIMS tab file using Windows curl.exe with integrated auth."""
    result = subprocess.run(
        ["curl.exe", "-s", "-S", "--negotiate", "-u", ":",
         "--fail", "-o", EIMS_TXT_PATH, EIMS_URL],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        raise RuntimeError("curl failed (exit %d): %s" % (result.returncode, result.stderr.strip()))
    return os.path.getsize(EIMS_TXT_PATH)


@st.cache_data(show_spinner=False)
def load_eims_df(mtime):
    """Read the cached EIMS tab file into a DataFrame. Cached on file mtime."""
    df = pd.read_csv(EIMS_TXT_PATH, sep="\t", dtype=str, keep_default_na=False)
    df.columns = [c.strip() for c in df.columns]
    return df


def eims_last_update():
    if os.path.exists(EIMS_TXT_PATH):
        return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(os.path.getmtime(EIMS_TXT_PATH)))
    return "not loaded"


@st.cache_data(show_spinner=False)
def load_lose_map():
    """Return {Operation: PI Dispose} from the LOSE_OPERATION mapping file."""
    if not os.path.exists(LOSE_MAP_PATH):
        return {}
    m = pd.read_excel(LOSE_MAP_PATH, dtype=str).fillna("")
    m.columns = [c.strip() for c in m.columns]
    if "Operation" not in m.columns or "PI Dispose" not in m.columns:
        return {}
    return {str(op).strip(): str(pi).strip()
            for op, pi in zip(m["Operation"], m["PI Dispose"]) if str(op).strip()}


def strip_via(value):
    """Drop 'VIA ...' suffixes and dedupe operations.

    '7226 VIA 1438' -> '7226';  '7571, 7571' -> '7571';
    '7571 VIA 2438, 7571' -> '7571'.
    """
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
    """Map a (possibly comma-joined) operation value to PI Dispose label(s).

    Special rule: any operation containing 7226 -> Eng_Assessment.
    """
    if not lose_value or not lose_map:
        # still honour the 7226 rule even without a mapping table
        if lose_value and "7226" in str(lose_value):
            return "Eng_Assessment"
        return ""
    parts = [p.strip() for p in str(lose_value).split(",") if p.strip()]
    labels = []
    for p in parts:
        if "7226" in p:
            lab = "Eng_Assessment"
        else:
            lab = lose_map.get(p, "")
        if lab and lab not in labels:
            labels.append(lab)
    return ", ".join(labels)


@st.cache_data(show_spinner=False)
def load_inqty_map():
    """Return (inqty_by_lot, create_data3_by_lot) from Lot_loss_operation.csv."""
    if not os.path.exists(LOT_LOSS_PATH):
        return {}, {}
    m = pd.read_csv(LOT_LOSS_PATH, dtype=str, keep_default_na=False)
    m.columns = [c.strip().upper() for c in m.columns]
    if "LOT" not in m.columns:
        return {}, {}
    lot = m["LOT"].astype(str).str.strip()
    inqty = {}
    if "INQTY" in m.columns:
        qty = pd.to_numeric(m["INQTY"], errors="coerce").fillna(0)
        inqty = qty.groupby(lot).sum().astype(int).to_dict()
    cd3 = {}
    if "CREATE_DATA3" in m.columns:
        cd3 = dict(zip(lot, m["CREATE_DATA3"].astype(str).str.strip()))
    return inqty, cd3


@st.cache_data(show_spinner=False)
def load_intms_products():
    """Return the list of valid ATMf 'Product Name' options (signal 221)."""
    if not os.path.exists(INTMS_PRODUCTS_PATH):
        return []
    try:
        with open(INTMS_PRODUCTS_PATH, "r", encoding="utf-8-sig") as f:
            data = json.load(f)
        return [str(x) for x in data] if isinstance(data, list) else []
    except (ValueError, OSError):
        return []


# Known Prodgroup3 -> ATMf Product Name overrides (verified). Add more as needed.
INTMS_PRODUCT_MAP = {
    "ADLN": "ADL N 0+8+1",
}


def guess_atmf_product(prodgroup3, products, rups_product=""):
    """Best-effort auto-map an EIMS Prodgroup3 code to an ATMf Product Name.

    1) exact override from INTMS_PRODUCT_MAP,
    2) fuzzy match of a normalised code against the ATMf option list.
    Returns "" when nothing is confidently found.
    """
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
        # Prefix/substring hit first (e.g. "ADLN" -> "ADLN0+8+1").
        hits = [p for p, t in targets.items() if t.startswith(n) or n in t]
        if hits:
            return min(hits, key=len)
        close = difflib.get_close_matches(n, list(targets.values()), n=1, cutoff=0.6)
        if close:
            for p, t in targets.items():
                if t == close[0]:
                    return p
    return ""


def product_has_sspec(product):
    """A Product string has an SSPEC/QDF when its last token is >1 char.

    Product is fixed-width, e.g.:
      '378TDT5V E B 5M ADSUA' -> last token 'ADSUA' (SSPEC present)  -> PPV
      '498AABAV A A CE     A' -> last token 'A'     (SSPEC blank)    -> Class
    """
    if not isinstance(product, str):
        return False
    tokens = product.split()
    return bool(tokens) and len(tokens[-1]) > 1


def add_category(df):
    """Add a 'Category' column: PPV (has SSPEC) or Class (no SSPEC)."""
    out = df.copy()
    if "Product" in out.columns:
        out.insert(0, "Category",
                   out["Product"].map(lambda p: "PPV" if product_has_sspec(p) else "Class"))
    else:
        out.insert(0, "Category", "Class")
    return out


CRITICAL_DAYS = 60


def days_at_operation_numeric(df):
    """Return Days_At_Operation as a numeric Series (NaN when not parseable)."""
    if "Days_At_Operation" in df.columns:
        return pd.to_numeric(df["Days_At_Operation"], errors="coerce")
    return pd.Series(index=df.index, dtype="float64")


def split_text(text):
    return [v.strip() for v in text.replace(",", "\n").splitlines() if v.strip()]


def run_rups_and_show(values, split=False, eims_qty=None, eims_prod=None):
    """Query RUPS for values, show table + download. Optionally split PPV/Class
    and report lots that returned no RUPS record.

    split=True: classify each unit by its RUPS SSPEC column
    (SSPEC present -> PPV, SSPEC empty -> Class), render two tables,
    and list any searched lots not found in RUPS.
    eims_qty: optional {lot: EIMS quantity} to cross-check EIMS vs RUPS per lot.
    """
    if not values:
        st.warning("Please provide at least one Lot or Visual ID.")
        return
    with st.spinner("Querying RUPS for %d value(s)..." % len(values)):
        try:
            status, records = query_units(values)
        except Exception as e:  # noqa: BLE001
            st.error("Error calling RUPS API: %s" % e)
            return
    if not records:
        st.error("HTTP %s - No data found for any of the %d searched value(s)."
                 % (status, len(values)))
        st.badge("Not found (%d)" % len(values), icon=":material/error:", color="red")
        st.dataframe(pd.DataFrame({"Searched (not found)": list(values)}),
                     width="stretch", hide_index=True)
        return
    df = records_to_df(records)

    if not split:
        st.success("Found %d record(s)." % len(df), icon=":material/check_circle:")
        st.dataframe(df, width="stretch", height=400)
        _rups_download(df)
        return

    # --- Split PPV / Class by SSPEC, and find lots with no RUPS record ---
    lot_series = df["EIMS_LOT"].astype(str) if "EIMS_LOT" in df.columns else pd.Series("", index=df.index)
    found = set(lot_series)
    not_found = [v for v in values if v not in found]

    # --- Per-lot cross-check: EIMS quantity vs RUPS unit count ---
    # RUPS returns one row per unit (VID); the QUANTITY column is the lot-level
    # quantity repeated on every row, so summing it is wrong. Count units instead.
    if "VID" in df.columns:
        rups_qty = df.groupby(lot_series)["VID"].nunique().to_dict()
    else:
        rups_qty = lot_series.value_counts().to_dict()
    # Per-lot LOSE_OPERATION values from RUPS (blank when none).
    lose_map = {}
    if "LOSE_OPERATION" in df.columns:
        for lot, grp in df.groupby(lot_series):
            vals = sorted({str(v).strip() for v in grp["LOSE_OPERATION"]
                           if pd.notna(v) and str(v).strip() and str(v).strip().lower() != "nan"})
            lose_map[lot] = ", ".join(vals)
    pi_lookup = load_lose_map()
    # Per-lot latest edit datetime from RUPS (max LAST_EDIT_DATE).
    last_edit = {}
    if "LAST_EDIT_DATE" in df.columns:
        led = pd.to_datetime(df["LAST_EDIT_DATE"], errors="coerce")
        last_edit = led.groupby(lot_series).max().to_dict()
    now = pd.Timestamp.now()
    inqty_map, cd3_map = load_inqty_map()
    summary_rows = []
    for lot in values:
        rq = int(rups_qty.get(lot, 0))
        eq = eims_qty.get(lot) if eims_qty else None
        if eq is None:
            match = "N/A" if rq else "Not found"
        else:
            match = "Match" if int(eq) == rq else "MISMATCH"
        lose_val = lose_map.get(lot, "")
        cd3_val = cd3_map.get(lot, "")
        lose_s = strip_via(lose_val)
        cd3_s = strip_via(cd3_val)
        # Merge LOSE_OPERATION (RUPS) and Create_Data3 (loss file):
        # take whichever is present; conflict if both present and differ.
        if lose_s and cd3_s:
            op_val = lose_s
            op_conflict = (lose_s != cd3_s)
        elif lose_s:
            op_val = lose_s
            op_conflict = False
        else:
            op_val = cd3_s
            op_conflict = False
        pi_val = map_pi_dispose(op_val, pi_lookup)
        # "Last used" only for matched lots categorised as PPV or Class.
        last_used = None
        if match == "Match" and pi_val in ("PPV", "Class"):
            ts = last_edit.get(lot)
            if ts is not None and pd.notna(ts):
                last_used = round((now - ts).total_seconds() / 86400.0, 1)
        in_qty = inqty_map.get(lot)
        summary_rows.append({
            "Lot": lot,
            "Prodgroup3": eims_prod.get(lot, "") if eims_prod else "",
            "EIMS_Qty": int(eq) if eq is not None else None,
            "RUPS_Qty": rq,
            "_in_qty": int(in_qty) if in_qty is not None else None,
            "Last_Used_Days": last_used,
            "Operation": op_val,
            "_op_conflict": op_conflict,
            "PI_Dispose": pi_val,
            "Status": match,
        })
    summary = pd.DataFrame(summary_rows)
    mismatches = summary[summary["Status"] == "MISMATCH"]

    # Stash the latest result so the email section can build a report.
    st.session_state["last_summary"] = summary
    st.session_state["last_rups_df"] = df

    if len(mismatches):
        st.toast("%d lot(s) MISMATCH between EIMS and RUPS!" % len(mismatches),
                 icon=":material/warning:")

    match_map = dict(zip(summary["Lot"], summary["Status"]))

    sspec = df["SSPEC"].astype(str).str.strip() if "SSPEC" in df.columns else pd.Series("", index=df.index)
    n_ppv = int((sspec != "").sum())
    n_cls = int((sspec == "").sum())

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("PPV units", n_ppv, border=True)
    c2.metric("Class units", n_cls, border=True)
    c3.metric("Lots not found", len(not_found), border=True,
              delta="%d missing" % len(not_found) if not_found else None,
              delta_color="inverse")
    c4.metric("Qty mismatches", len(mismatches), border=True,
              delta="%d off" % len(mismatches) if len(mismatches) else None,
              delta_color="inverse")

    # --- By-lot summary: product = big block, PI Dispose = inner block ---
    st.badge("By lot — EIMS vs RUPS (%d)" % len(summary),
             icon=":material/rule:", color="violet")
    st.caption(":material/confirmation_number: Each table has its own **Select all / Clear** "
               "buttons — tick lots, then use the **Submit ATMf ticket** panel below.")

    _atmf_picks = []  # lots ticked across all tables below

    def _hl_status(val):
        try:
            dark = st.context.theme.type == "dark"
        except Exception:
            dark = True
        if dark:
            colors = {
                "MISMATCH": "background-color:#3a1d1f;color:#ff8a8a;font-weight:600",
                "Not found": "background-color:#33290f;color:#f0c674;font-weight:600",
                "Match": "background-color:#173026;color:#5ed99a;font-weight:600",
            }
        else:
            colors = {
                "MISMATCH": "background-color:#fdecec;color:#cf222e;font-weight:600",
                "Not found": "background-color:#fbf3e0;color:#bf8700;font-weight:600",
                "Match": "background-color:#e6f6ec;color:#1a7f37;font-weight:600",
            }
        return colors.get(val, "")

    _status_order = {"Match": 0, "N/A": 1, "Not found": 2, "MISMATCH": 3}

    _col_cfg = {
        "Last_Used_Days": st.column_config.NumberColumn(
            "Last used (days)",
            help="Days since the lot was last edited on RUPS (PPV/Class matched lots only)",
            format="%.1f",
        ),
    }

    def _sort_status(frame):
        """Auto-sort by Status (Match -> N/A -> Not found -> MISMATCH),
        then by Last_Used_Days descending (most stale first)."""
        f = frame.copy()
        f["_status_ord"] = f["Status"].map(lambda v: _status_order.get(v, 9))
        sort_cols, ascending = ["_status_ord"], [True]
        if "Last_Used_Days" in f.columns:
            f["_lu_ord"] = pd.to_numeric(f["Last_Used_Days"], errors="coerce")
            sort_cols.append("_lu_ord")
            ascending.append(False)
        f = f.sort_values(sort_cols, ascending=ascending, kind="stable",
                          na_position="last")
        return f.drop(columns=[c for c in ("_status_ord", "_lu_ord")
                               if c in f.columns])

    _HIDDEN_COLS = ["_in_qty", "_op_conflict", "Status"]

    def _cell_style(frame):
        """Red Operation on conflict; red EIMS_Qty & RUPS_Qty when they differ."""
        try:
            dark = st.context.theme.type == "dark"
        except Exception:
            dark = True
        red = ("background-color:#3a1d1f;color:#ff8a8a;font-weight:600" if dark
               else "background-color:#fdecec;color:#cf222e;font-weight:600")
        styles = pd.DataFrame("", index=frame.index, columns=frame.columns)
        if "Operation" in frame.columns and "_op_conflict" in frame.columns:
            conflict = frame["_op_conflict"].fillna(False).astype(bool)
            styles.loc[conflict, "Operation"] = red
        if "Status" in frame.columns and "EIMS_Qty" in frame.columns and "RUPS_Qty" in frame.columns:
            mis = frame["Status"].astype(str) == "MISMATCH"
            styles.loc[mis, "EIMS_Qty"] = red
            styles.loc[mis, "RUPS_Qty"] = red
        return styles

    def _styled(frame):
        show = frame.drop(columns=[c for c in _HIDDEN_COLS if c in frame.columns])
        cell_styles = _cell_style(frame).drop(
            columns=[c for c in _HIDDEN_COLS if c in frame.columns])
        return show.style.apply(lambda _: cell_styles, axis=None)

    _FLAG_MAP = {"MISMATCH": "⚠", "Not found": "❓", "N/A": "–", "Match": "✓"}

    def _editable_table(frame, key):
        """Self-contained per-table picker: its own Select all / Clear buttons and
        a Submit checkbox column. Ticked lots are collected into _atmf_picks.
        Cell coloring isn't available in data_editor, so a Flag column
        (⚠ = mismatch) keeps mismatches visible."""
        disp = _sort_status(frame).copy()
        flag = disp["Status"].map(_FLAG_MAP).fillna("") if "Status" in disp.columns \
            else pd.Series("", index=disp.index)
        keep = [c for c in ["Lot", "EIMS_Qty", "RUPS_Qty",
                            "Last_Used_Days", "Operation", "PI_Dispose"]
                if c in disp.columns]
        out = disp[keep].copy()
        out.insert(0, "Flag", flag.values)

        # Per-table bulk selection state.
        def_key = "%s_def" % key
        nonce_key = "%s_nonce" % key
        b1, b2, b3 = st.columns([1, 1, 5])
        if b1.button("Select all", icon=":material/select_all:",
                     key="%s_selall" % key, width="stretch"):
            st.session_state[def_key] = True
            st.session_state[nonce_key] = st.session_state.get(nonce_key, 0) + 1
        if b2.button("Clear", icon=":material/deselect:",
                     key="%s_clear" % key, width="stretch"):
            st.session_state[def_key] = False
            st.session_state[nonce_key] = st.session_state.get(nonce_key, 0) + 1
        tbl_default = bool(st.session_state.get(def_key, False))
        tbl_nonce = st.session_state.get(nonce_key, 0)

        out.insert(0, "Submit", tbl_default)
        edited = st.data_editor(
            out, width="stretch", hide_index=True, key="%s_%d" % (key, tbl_nonce),
            column_config={
                "Submit": st.column_config.CheckboxColumn(
                    "✓", help="Tick to submit an ATMf ticket for this lot"),
                "Flag": st.column_config.TextColumn("!", help="⚠ = EIMS/RUPS mismatch"),
                "Last_Used_Days": _col_cfg["Last_Used_Days"],
            },
            disabled=[c for c in out.columns if c != "Submit"],
        )
        _n = int(edited["Submit"].sum())
        b3.caption("%d / %d lot selected" % (_n, len(out)))
        for lot in edited[edited["Submit"]]["Lot"].astype(str).tolist():
            _atmf_picks.append(lot)

    _PI_META = {
        "PPV": (":material/verified:", "green", "🟢"),
        "Class": (":material/category:", "blue", "🔵"),
        "Eng_Assessment": (":material/science:", "orange", "🟠"),
    }

    def _show_by_pi(frame, key_prefix=""):
        has_pi = "PI_Dispose" in frame.columns and frame["PI_Dispose"].str.strip().ne("").any()
        if not has_pi:
            _editable_table(frame, key="ed_%s_all" % key_prefix)
            return

        pi_order = {"PPV": 0, "Class": 1, "Eng_Assessment": 2}
        pis = sorted(frame["PI_Dispose"].unique(),
                     key=lambda x: (pi_order.get(x, 8), x == "", x))

        # Per-group metric row (counts + mismatch).
        mcols = st.columns(len(pis))
        for col, pi in zip(mcols, pis):
            psub = frame[frame["PI_Dispose"] == pi]
            n_mis = int((psub["Status"] == "MISMATCH").sum())
            _, _, emoji = _PI_META.get(pi, (":material/help:", "gray", "⚪"))
            label = pi if pi else "Other"
            col.metric("%s %s" % (emoji, label), len(psub),
                       delta=("%d mismatch" % n_mis) if n_mis else None,
                       delta_color="inverse", border=True)

        # One tab per PI group.
        tab_labels = []
        for pi in pis:
            _, _, emoji = _PI_META.get(pi, (":material/help:", "gray", "⚪"))
            psub = frame[frame["PI_Dispose"] == pi]
            tab_labels.append("%s %s (%d)" % (emoji, pi if pi else "Other", len(psub)))
        tabs = st.tabs(tab_labels)
        for tab, pi in zip(tabs, pis):
            psub = frame[frame["PI_Dispose"] == pi]
            with tab:
                _editable_table(psub, key="ed_%s_%s" % (key_prefix, pi or "other"))

    if "Prodgroup3" in summary.columns and summary["Prodgroup3"].str.strip().ne("").any():
        prods = sorted(summary["Prodgroup3"].unique(), key=lambda x: (x == "", x))
        for grp in prods:
            sub = summary[summary["Prodgroup3"] == grp]
            label = grp if grp else "(no product)"
            key_pref = "".join(ch if ch.isalnum() else "_" for ch in str(label)) or "none"
            n_mis = int((sub["Status"] == "MISMATCH").sum())
            n_ppv = int((sub["PI_Dispose"] == "PPV").sum())
            n_cls = int((sub["PI_Dispose"] == "Class").sum())
            n_eng = int((sub["PI_Dispose"] == "Eng_Assessment").sum())
            # Compact collapsible per product; auto-expand ones with mismatches.
            title = "%s  ·  %d lot(s)  ·  🟢 PPV %d  ·  🔵 Class %d" % (
                label, len(sub), n_ppv, n_cls)
            if n_eng:
                title += "  ·  🟠 Eng %d" % n_eng
            if n_mis:
                title += "  ·  ⚠ %d mismatch" % n_mis
            with st.expander(title, expanded=bool(n_mis)):
                _show_by_pi(sub, key_prefix=key_pref)
    else:
        _editable_table(summary, key="ed_all")

    # --- ATMf ticket panel (below tables so ticked lots are already collected) ---
    with st.container(border=True):
        st.markdown("**:material/confirmation_number: Submit ATMf ticket** "
                    "— signal **221** *(Reject Management - POR)*")
        _products = load_intms_products()

        # Per-lot product auto-mapping from Prodgroup3 (+ RUPS product fallback).
        _uniq = summary.drop_duplicates("Lot")
        _info = _uniq.set_index(_uniq["Lot"].astype(str))
        _rups_prod = {}
        if "EIMS_LOT" in df.columns and "PRODUCT_NAME" in df.columns:
            for _lot, _grp in df.groupby(df["EIMS_LOT"].astype(str)):
                _vals = [str(v).strip() for v in _grp["PRODUCT_NAME"]
                         if pd.notna(v) and str(v).strip()]
                _rups_prod[_lot] = _vals[0] if _vals else ""

        c_cust, c_pay, c_stage = st.columns(3)
        atmf_customer = c_cust.selectbox(
            "Customer", INTMS_CUSTOMERS,
            index=INTMS_CUSTOMERS.index(INTMS_DEFAULTS["Customer"]), key="atmf_customer")
        atmf_payer = c_pay.selectbox(
            "Payer BU", INTMS_PAYER_BUS,
            index=INTMS_PAYER_BUS.index(INTMS_DEFAULTS["Payer BU"]), key="atmf_payer")
        atmf_stage = c_stage.selectbox(
            "Product Stage", INTMS_PRODUCT_STAGES,
            index=INTMS_PRODUCT_STAGES.index(INTMS_DEFAULTS["Product Stage"]),
            key="atmf_stage")

        c_fac, c_ship, c_nfo = st.columns(3)
        atmf_factory = c_fac.selectbox(
            "Factories", INTMS_FACTORIES,
            index=INTMS_FACTORIES.index(INTMS_DEFAULTS["Factories"]), key="atmf_factory")
        atmf_shipping = c_ship.selectbox(
            "Shipping/LYA/FA Only", ["No", "Yes"], key="atmf_shipping")
        atmf_nfo = c_nfo.selectbox("NFO shipment", ["No", "Yes"], key="atmf_nfo")

        atmf_test_mode = st.radio(
            "Test Area", ["Auto (per lot PPV/Class)"] + INTMS_TEST_AREAS,
            horizontal=True, key="atmf_test",
            help="Auto = each lot uses its own PPV/Class classification.")
        _force_prod = st.selectbox(
            "Force product (optional — override auto-map for all lots)",
            ["(auto per lot)"] + _products, key="atmf_force_prod") \
            if _products else "(auto per lot)"
        # Build a per-lot mapping, then group into one ticket per (product, area).
        per_lot = []
        for lot in _atmf_picks:
            row = _info.loc[lot] if lot in _info.index else {}
            pg3 = str(row.get("Prodgroup3", "")) if hasattr(row, "get") else ""
            pi = str(row.get("PI_Dispose", "")).strip() if hasattr(row, "get") else ""
            if _force_prod and _force_prod != "(auto per lot)":
                prod = _force_prod
            else:
                prod = guess_atmf_product(pg3, _products, _rups_prod.get(lot, ""))
            area = ("PPV" if pi == "PPV" else "Class") \
                if atmf_test_mode.startswith("Auto") else atmf_test_mode
            per_lot.append({"Lot": lot, "Prodgroup3": pg3,
                            "Product Name": prod, "Test Area": area})

        # One ticket per (Product Name, Test Area); all its lots listed inside.
        _REQUESTS = ["Scrap these lot", "Transfer these lot to HVE Lab"]
        groups = {}
        for pl in per_lot:
            key = (pl["Product Name"], pl["Test Area"])
            groups.setdefault(key, {"Product Name": pl["Product Name"],
                                    "Test Area": pl["Test Area"],
                                    "Prodgroup3": pl["Prodgroup3"],
                                    "Lots": []})
            groups[key]["Lots"].append(pl["Lot"])
        preview = list(groups.values())
        for g in preview:
            g["OK"] = "✓" if g["Product Name"] else "⚠ no map"

        if _atmf_picks:
            st.caption("These %d ticket(s) will be created — pick a **Request** "
                       "per product:" % len(preview))
            _pv = pd.DataFrame([
                {"Request": _REQUESTS[0], "Product Name": g["Product Name"],
                 "Test Area": g["Test Area"], "Lots": ", ".join(g["Lots"]),
                 "# Lots": len(g["Lots"]), "OK": g["OK"]} for g in preview])
            _edited_pv = st.data_editor(
                _pv, width="stretch", hide_index=True, key="atmf_req_editor",
                column_config={
                    "Request": st.column_config.SelectboxColumn(
                        "Request", options=_REQUESTS, required=True,
                        help="A separate request per product"),
                    "Product Name": st.column_config.TextColumn(disabled=True),
                    "Test Area": st.column_config.TextColumn(disabled=True),
                    "Lots": st.column_config.TextColumn(disabled=True),
                    "# Lots": st.column_config.NumberColumn(disabled=True),
                    "OK": st.column_config.TextColumn(disabled=True),
                })
            for i, g in enumerate(preview):
                g["Request"] = str(_edited_pv.iloc[i]["Request"])
            _unmapped = [", ".join(g["Lots"]) for g in preview if not g["Product Name"]]
            if _unmapped:
                st.warning("No product auto-mapped for: %s — use "
                           "**Force product** above." % "; ".join(_unmapped),
                           icon=":material/warning:")
        else:
            st.caption(":orange[No lot ticked — tick the ✓ column in a table above.]")

        _all_mapped = bool(_atmf_picks) and all(g["Product Name"] for g in preview)
        atmf_submit = st.button(
            "Submit ticket(s)", type="primary", icon=":material/send:",
            disabled=not _all_mapped, key="atmf_submit")

        if atmf_submit:
            results = []
            with st.spinner("Submitting %d ticket(s)…" % len(preview)):
                for g in preview:
                    form = {
                        "Customer": atmf_customer,
                        "Payer BU": atmf_payer,
                        "Product Stage": atmf_stage,
                        "Product Name": g["Product Name"],
                        "Test Area": g["Test Area"],
                        "Factories": atmf_factory,
                        "Shipping/LYA/FA Only": atmf_shipping,
                        "NFO shipment": atmf_nfo,
                    }
                    ok, tid, detail = submit_intms_ticket(INTMS_SIGNAL_ID, form)
                    filled = None
                    if ok and tid:
                        # Prefill "Fill in required information*" (action 4):
                        # this product's request first line, then one lot per line.
                        lots_html = "".join("<p>%s</p>" % l for l in g["Lots"])
                        rich = "<p>%s</p>%s" % (g.get("Request", _REQUESTS[0]), lots_html)
                        filled, _ = fill_intms_action(tid, 4, rich)
                    results.append((", ".join(g["Lots"]), g["Test Area"],
                                    ok, tid, detail, filled))
            for lots, area, ok, tid, detail, filled in results:
                if ok and tid:
                    note = "" if filled else " (fill-in step not prefilled)"
                    url = INTMS_TICKET_URL % tid
                    st.success("%s (%s) → ticket #%s%s" % (lots, area, tid, note),
                               icon=":material/check_circle:")
                    st.markdown("&nbsp;&nbsp;:material/link: [%s](%s)" % (url, url))
                else:
                    st.error("%s (%s) failed: %s" % (lots, area, detail),
                             icon=":material/error:")


    # --- Merged per-unit table with Category + Match columns ---
    merged = df.copy()
    merged.insert(0, "Category", (sspec != "").map({True: "PPV", False: "Class"}).values)
    if "EIMS_LOT" in df.columns:
        merged.insert(0, "Match", df["EIMS_LOT"].astype(str).map(match_map).values)

    st.badge("RUPS units (%d)" % len(merged), icon=":material/table_rows:", color="blue")
    st.dataframe(merged, width="stretch", height=340)

    _rups_download(merged)


def _rups_download(df):
    st.download_button(
        "Download Excel",
        data=df_to_excel_bytes(df),
        file_name="RUPS_data_%s.xlsx" % time.strftime("%Y%m%d_%H%M%S"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        icon=":material/download:",
    )


# ======================================================================= EMAIL
def build_report(summary):
    """Return (subject, html_body) mirroring the web layout in a clean LIGHT theme:
    each product = a bordered card, inside PPV/Class/Eng_Assessment blocks."""
    total = len(summary)
    n_match = int((summary["Status"] == "Match").sum())
    n_mis = int((summary["Status"] == "MISMATCH").sum())
    prods = sorted({p for p in summary["Prodgroup3"] if str(p).strip()})
    subject = "EIMS <-> RUPS Lot Reconciliation - %s (%d lots)" % (
        time.strftime("%Y-%m-%d"), total)

    # Light theme palette
    BG = "#ffffff"
    CARD = "#f6f8fa"
    BORDER = "#d0d7de"
    TEXT = "#1f2328"
    MUTED = "#656d76"
    HEAD_BG = "#eaeef2"
    RED = "background:#ffebe9;color:#cf222e;font-weight:700"
    order = {"Match": 0, "N/A": 1, "Not found": 2, "MISMATCH": 3}

    def _table(frame):
        frame = frame.sort_values(
            "Status", key=lambda s: s.map(lambda v: order.get(v, 9)), kind="stable")
        cols = [c for c in ["Lot", "EIMS_Qty", "RUPS_Qty",
                            "Last_Used_Days", "Operation", "PI_Dispose"]
                if c in frame.columns]
        head = "".join(
            "<th style='border:1px solid %s;padding:6px 10px;background:%s;color:%s;"
            "text-align:left'>%s</th>" % (BORDER, HEAD_BG, TEXT, c) for c in cols)
        rows = ""
        for _, r in frame.iterrows():
            mis = str(r.get("Status", "")) == "MISMATCH"
            cells = ""
            for c in cols:
                val = "" if pd.isna(r[c]) else r[c]
                extra = RED if (mis and c in ("EIMS_Qty", "RUPS_Qty")) else ("color:%s" % TEXT)
                cells += ("<td style='border:1px solid %s;padding:6px 10px;%s'>%s</td>"
                          % (BORDER, extra, val))
            rows += "<tr>%s</tr>" % cells
        return ("<table style='border-collapse:collapse;font-family:Segoe UI,Arial;"
                "font-size:12px;margin:6px 0;width:100%%'>"
                "<thead><tr>%s</tr></thead><tbody>%s</tbody></table>" % (head, rows))

    def _pi_blocks(frame):
        out = ""
        pi_order = {"PPV": 0, "Class": 1, "Eng_Assessment": 2}
        pis = sorted(frame["PI_Dispose"].unique(),
                     key=lambda x: (pi_order.get(x, 8), x == "", x))
        for pi in pis:
            psub = frame[frame["PI_Dispose"] == pi]
            label = pi if str(pi).strip() else "Other"
            hue = {"PPV": "#1a7f37", "Class": "#0969da",
                   "Eng_Assessment": "#9a6700"}.get(pi, "#656d76")
            m = int((psub["Status"] == "MISMATCH").sum())
            mtxt = (" &middot; <span style='color:#cf222e'>%d mismatch</span>" % m) if m else ""
            out += ("<div style='margin:10px 0 2px 0;font-weight:700;font-size:13px;color:%s'>"
                    "&#9632; %s <span style='color:%s;font-weight:400'>— %d lot(s)%s</span></div>"
                    % (hue, label, MUTED, len(psub), mtxt))
            out += _table(psub)
        return out

    parts = ["<div style='background:%s;padding:20px;font-family:Segoe UI,Arial;"
             "font-size:14px;color:%s'>" % (BG, TEXT)]
    # Header banner
    parts.append(
        "<div style='background:linear-gradient(120deg,#0969da,#0a66c2 55%%,#1a7f37);"
        "padding:20px 24px;border-radius:16px;color:#ffffff;margin-bottom:14px'>"
        "<div style='font-size:22px;font-weight:800'>&#9881; RUPS + EIMS inventory reconciliation</div>"
        "<div style='font-size:13px;color:#d6e4ff;margin-top:4px'>Report generated %s</div></div>"
        % time.strftime("%Y-%m-%d %H:%M"))

    parts.append("<p>Hi all,</p>")
    parts.append("<p style='color:%s'>Please find the EIMS &harr; RUPS reconciliation results below.</p>" % TEXT)
    parts.append("<div style='background:%s;border:1px solid %s;border-radius:10px;"
                 "padding:12px 16px;margin:8px 0'>" % (CARD, BORDER))
    parts.append("<b>Summary</b><ul style='margin:6px 0'>")
    parts.append("<li>Products: <b>%s</b></li>" % (", ".join(prods) if prods else "-"))
    parts.append("<li>Total lots: <b>%d</b></li>" % total)
    parts.append("<li style='color:#1a7f37'>Matched: <b>%d</b></li>" % n_match)
    parts.append("<li style='color:#cf222e'>Quantity mismatches: <b>%d</b></li>" % n_mis)
    parts.append("</ul></div>")

    if "Prodgroup3" in summary.columns and summary["Prodgroup3"].str.strip().ne("").any():
        for grp in sorted(summary["Prodgroup3"].unique(), key=lambda x: (x == "", x)):
            sub = summary[summary["Prodgroup3"] == grp]
            label = grp if str(grp).strip() else "(no product)"
            g_mis = int((sub["Status"] == "MISMATCH").sum())
            accent = "#cf222e" if g_mis else "#0969da"
            head = ("<span style='color:%s'>&#128230; %s</span> "
                    "<span style='color:%s;font-weight:400'>— %d lot(s)</span>"
                    % (TEXT, label, MUTED, len(sub)))
            if g_mis:
                head += " <span style='color:#cf222e'>&middot; %d mismatch</span>" % g_mis
            parts.append(
                "<div style='border:1px solid %s;border-left:4px solid %s;border-radius:12px;"
                "padding:12px 16px;margin:14px 0;background:%s'>"
                "<div style='font-size:16px;font-weight:800;margin-bottom:4px'>%s</div>%s</div>"
                % (BORDER, accent, CARD, head,
                   _pi_blocks(sub) if "PI_Dispose" in sub.columns else _table(sub)))
    else:
        parts.append(_table(summary))

    parts.append("<p style='color:%s'>Thanks,<br>[Your name]</p></div>" % MUTED)
    return subject, "".join(parts)


def send_via_outlook(subject, html_body, to="", cc="", attachment_path=None,
                     display_only=True):
    """Create an Outlook mail item via PowerShell COM. Displays a draft for review
    (display_only=True) instead of auto-sending."""
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


# ======================================================================= InTMS
def submit_intms_ticket(signal_id, form_json, timeout=60):
    """Create an InTMS/ATMf ticket via curl Kerberos SSO (--negotiate -u :).

    Returns (ok, ticket_id, detail). ok=True only when the API returns JSON;
    detail is the parsed JSON dict on success or an error string otherwise.
    """
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
    # Success: API returns a dict with an "id". Errors often come back as a
    # dict of field errors or a list of messages — surface those as text.
    if isinstance(data, dict):
        tid = data.get("id")
        if tid:
            return True, tid, data
        return False, None, json.dumps(data)[:500]
    return False, None, json.dumps(data)[:500]


def fill_intms_action(ticket_id, step_id, rich_text, timeout=60):
    """PATCH an action-flow step's rich_text (e.g. 'Fill in required information').

    Sending rich_text at ticket creation does NOT reach that step — it must be
    PATCHed afterwards. Returns (ok, detail).
    """
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



# ======================================================================= UI
st.set_page_config(
    page_title="Reject Management",
    page_icon=":material/memory:",
    layout="wide",
)

st.html(
    """
    <div style="
        position: relative; overflow: hidden;
        background: linear-gradient(120deg, #0a2a4a 0%, #10495c 45%, #3a2a6a 100%);
        padding: 30px 34px; border-radius: 20px; margin-bottom: 10px;
        border: 1px solid rgba(120,180,255,0.18);
        box-shadow: 0 10px 40px rgba(0,90,180,0.35), inset 0 1px 0 rgba(255,255,255,0.06);">
      <div style="
        position:absolute; top:-60px; right:-40px; width:220px; height:220px;
        background: radial-gradient(circle, rgba(88,166,255,0.35) 0%, rgba(88,166,255,0) 70%);
        filter: blur(6px);"></div>
      <div style="
        color:#eaf3ff; font-size:32px; font-weight:800; letter-spacing:.3px;
        text-shadow: 0 2px 14px rgba(88,166,255,0.35); position:relative;">
        ⚙️ Reject Management
      </div>
      <div style="color:rgba(220,235,255,0.82); font-size:15px; margin-top:8px; position:relative;">
        RUPS + EIMS inventory reconciliation — search by lot / visual ID, explore inventory for DOE.
      </div>
    </div>
    """
)

tab_manual, tab_eims = st.tabs([
    ":material/search: Manual RUPS search",
    ":material/inventory_2: EIMS inventory",
])

# ---- Part 1: manual lot list ------------------------------------------------
with tab_manual:
    st.subheader("Search RUPS by lot / visual ID")
    with st.container(border=True):
        col_lots, col_vis = st.columns(2)
        with col_lots:
            lots_text = st.text_area(
                "Lot(s)", height=160,
                placeholder="One per line or comma-separated, e.g.\nARL75630A\nARL75630B",
            )
        with col_vis:
            visual_text = st.text_area(
                "Visual ID(s)", height=160,
                placeholder="One per line or comma-separated",
            )
        search_manual = st.button(
            "Search RUPS", type="primary", icon=":material/search:", key="btn_manual")
    if search_manual:
        values = split_text(lots_text) + split_text(visual_text)
        run_rups_and_show(values)

# ---- Part 2: EIMS Inventory -------------------------------------------------
with tab_eims:
    # --- Section 1: data source ---
    with st.container(border=True):
        st.subheader(":material/database: 1. Data source")
        with st.container(horizontal=True, vertical_alignment="center"):
            if st.button("Refresh EIMS", icon=":material/refresh:", key="btn_refresh"):
                with st.spinner("Downloading EIMS data & reloading all sources..."):
                    try:
                        download_eims()
                        # Clear every cached data source so all flows re-run:
                        # EIMS data, PI-Dispose mapping, and lot loss-operation file.
                        load_eims_df.clear()
                        load_lose_map.clear()
                        load_inqty_map.clear()
                        _query_units_cached.clear()
                        st.success("Refreshed EIMS + mapping + loss + RUPS data.")
                    except Exception as e:  # noqa: BLE001
                        st.error("Refresh failed: %s" % e)
                st.rerun()
            st.badge("Last update: %s" % eims_last_update(),
                     icon=":material/schedule:", color="blue")

    if not os.path.exists(EIMS_TXT_PATH):
        st.info("No local EIMS file yet. Click **Refresh EIMS** to download it.",
                icon=":material/info:")
    else:
        df = load_eims_df(os.path.getmtime(EIMS_TXT_PATH))

        # --- Compact filter bar (small, low-key) inside an expander ---
        prod_options = sorted(v for v in df["Prodgroup3"].unique() if v != "") \
            if "Prodgroup3" in df.columns else []
        with st.expander(":material/tune: Filters", expanded=False):
            eims_view = st.segmented_control(
                "View",
                ["PPV/Class review (Op 7000)", "Operation 4000 (all products)"],
                default="PPV/Class review (Op 7000)",
                key="eims_view",
            ) or "PPV/Class review (Op 7000)"
            _is_4000 = eims_view.startswith("Operation 4000")
            if _is_4000:
                view_filters = {"Operation": ["4000"]}
                view_all_products = True
            else:
                view_filters = DEFAULT_FILTERS
                view_all_products = False

            default_prods = [] if view_all_products else \
                [p for p in view_filters.get("Prodgroup3", []) if p in prod_options]
            products = st.multiselect(
                "Products (Prodgroup3)",
                prod_options,
                default=default_prods or prod_options,
                key="eims_products_%s" % ("4000" if _is_4000 else "def"),
            )
            other_cols = [c for c in EIMS_FILTER_COLUMNS if c != "Prodgroup3"]
            fcols = st.columns(len(other_cols))
            selected = {}
            for i, col in enumerate(other_cols):
                options = sorted(v for v in df[col].unique() if v != "") \
                    if col in df.columns else []
                default = [v for v in view_filters.get(col, []) if v in options]
                with fcols[i]:
                    selected[col] = st.multiselect(
                        col, options, default=default,
                        key="eims_%s_%s" % (col, "4000" if _is_4000 else "def"))

        mask = df["Prodgroup3"].isin(products) if products else pd.Series(True, index=df.index)
        for col, vals in selected.items():
            if vals:
                mask &= df[col].isin(vals)
        filtered = add_category(df[mask])

        ppv_df = filtered[filtered["Category"] == "PPV"]
        class_df = filtered[filtered["Category"] == "Class"]
        critical_df = filtered[days_at_operation_numeric(filtered) > CRITICAL_DAYS]

        # --- Section 3: summary ---
        with st.container(border=True):
            st.subheader(":material/insights: 3. Summary")
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Products", len(products), border=True)
            m2.metric("Total rows", len(filtered), border=True)
            m3.metric("PPV", len(ppv_df), border=True)
            m4.metric("Class", len(class_df), border=True)
            m5.metric("Critical Dispose (>%dd)" % CRITICAL_DAYS, len(critical_df),
                      delta="%d over limit" % len(critical_df) if len(critical_df) else None,
                      delta_color="inverse", border=True)

            # Per-product lot summary table.
            if "Prodgroup3" in filtered.columns and len(filtered):
                _lotcol = EIMS_LOT_COLUMN if EIMS_LOT_COLUMN in filtered.columns else None
                rows = []
                for prod, grp in filtered.groupby("Prodgroup3"):
                    n_lots = grp[_lotcol].nunique() if _lotcol else len(grp)
                    crit_grp = grp[days_at_operation_numeric(grp) > CRITICAL_DAYS]
                    n_crit = crit_grp[_lotcol].nunique() if _lotcol else len(crit_grp)
                    qty = int(pd.to_numeric(grp.get("Quantity"), errors="coerce").fillna(0).sum()) \
                        if "Quantity" in grp.columns else 0
                    ppv_grp = grp[grp["Category"] == "PPV"]
                    cls_grp = grp[grp["Category"] == "Class"]
                    n_ppv = ppv_grp[_lotcol].nunique() if _lotcol else len(ppv_grp)
                    n_cls = cls_grp[_lotcol].nunique() if _lotcol else len(cls_grp)
                    rows.append({
                        "Product": prod,
                        "Lots": n_lots,
                        "Unit": qty,
                        "PPV (lots)": n_ppv,
                        "Class (lots)": n_cls,
                        "Critical Dispose (>%dd)" % CRITICAL_DAYS: n_crit,
                    })
                prod_summary = pd.DataFrame(rows).sort_values(
                    "Critical Dispose (>%dd)" % CRITICAL_DAYS,
                    ascending=False, ignore_index=True)

                def _prod_slug(name):
                    return "".join(ch if ch.isalnum() else "_"
                                   for ch in str(name)) or "none"

                prod_summary.insert(
                    0, "Detail",
                    prod_summary["Product"].map(lambda p: "#prod-%s" % _prod_slug(p)))
                st.markdown("**:material/summarize: Lots by product**")
                st.dataframe(
                    prod_summary, width="stretch", hide_index=True, height=280,
                    column_config={
                        "Detail": st.column_config.LinkColumn(
                            "Go", help="Jump to this product's detail below",
                            display_text="Open ↓"),
                    })
                focus_product = ""


        # --- Section 4: inventory tables ---
        with st.container(border=True):
            st.subheader(":material/table_view: 4. Inventory tables")
            t_all, t_crit = st.tabs([
                "Inventory (%d)" % len(filtered),
                "Critical Dispose >%dd (%d)" % (CRITICAL_DAYS, len(critical_df)),
            ])
            with t_all:
                st.dataframe(filtered, width="stretch", height=340)
            with t_crit:
                st.caption(":red[**Critical Dispose = lots sitting at the same "
                           "operation for more than %d days.**]" % CRITICAL_DAYS)
                crit_sorted = critical_df.sort_values(
                    "Days_At_Operation",
                    key=lambda s: pd.to_numeric(s, errors="coerce"),
                    ascending=False).copy()
                if "Days_At_Operation" in crit_sorted.columns:
                    crit_sorted["Days_At_Operation"] = pd.to_numeric(
                        crit_sorted["Days_At_Operation"], errors="coerce").round(1)
                crit_max = float(crit_sorted["Days_At_Operation"].max()) \
                    if len(crit_sorted) and "Days_At_Operation" in crit_sorted.columns else 0
                st.dataframe(
                    crit_sorted,
                    width="stretch", height=340,
                    column_config={
                        "Days_At_Operation": st.column_config.ProgressColumn(
                            "Days_At_Operation",
                            help="Days sitting at the current operation",
                            format="%.1f d",
                            min_value=0,
                            max_value=max(crit_max, CRITICAL_DAYS),
                        ),
                    },
                )

        # --- Section 5: RUPS reconciliation (auto) ---
        with st.container(border=True):
            st.markdown("<div id='rups-detail'></div>", unsafe_allow_html=True)
            st.subheader(":material/travel_explore: 5. RUPS reconciliation")
            rups_scope = st.segmented_control(
                "Lot group",
                ["All inventory", "Critical Dispose only"],
                default="All inventory",
                label_visibility="collapsed",
            ) or "All inventory"
            scope_df = critical_df if rups_scope == "Critical Dispose only" else filtered

            picked = sorted(v for v in scope_df[EIMS_LOT_COLUMN].unique() if v != "") \
                if EIMS_LOT_COLUMN in scope_df.columns else []
            st.caption("Auto-searching %d lot(s) in RUPS." % len(picked))
            eims_qty = {}
            if "Quantity" in scope_df.columns and EIMS_LOT_COLUMN in scope_df.columns:
                q = pd.to_numeric(scope_df["Quantity"], errors="coerce").fillna(0)
                eims_qty = q.groupby(scope_df[EIMS_LOT_COLUMN]).sum().astype(int).to_dict()
            eims_prod = {}
            if "Prodgroup3" in scope_df.columns and EIMS_LOT_COLUMN in scope_df.columns:
                for lot, grp in scope_df.groupby(EIMS_LOT_COLUMN):
                    prods = sorted({str(v).strip() for v in grp["Prodgroup3"]
                                    if pd.notna(v) and str(v).strip() and str(v).strip().lower() != "nan"})
                    eims_prod[lot] = ", ".join(prods)
            if picked:
                run_rups_and_show(picked, split=True, eims_qty=eims_qty, eims_prod=eims_prod)
            else:
                st.info("No lots in the current filter to search.", icon=":material/info:")

        # --- Section 6: report email (only builds when requested) ---
        with st.container(border=True):
            st.subheader(":material/mail: 6. Report email")
            last_summary = st.session_state.get("last_summary")
            if last_summary is None or last_summary.empty:
                st.info("Run a RUPS search above first, then compose the report.",
                        icon=":material/info:")
            elif st.button("Compose report draft", icon=":material/edit_note:",
                           key="btn_compose"):
                st.session_state["compose_report"] = True

            if (last_summary is not None and not last_summary.empty
                    and st.session_state.get("compose_report")):
                subject_def, html_body = build_report(last_summary)
                c1, c2 = st.columns(2)
                to = c1.text_input("To", placeholder="alias@intel.com; alias2@intel.com")
                cc = c2.text_input("Cc", placeholder="optional")
                subject = st.text_input("Subject", value=subject_def)
                with st.expander("Preview email body"):
                    st.html(html_body)
                attach = st.checkbox("Attach RUPS Excel", value=True)
                if st.button("Open draft in Outlook", type="primary",
                             icon=":material/mail:", key="btn_email"):
                    attach_path = None
                    if attach and st.session_state.get("last_rups_df") is not None:
                        attach_path = os.path.join(
                            OUTPUT_DIR, "RUPS_report_%s.xlsx" % time.strftime("%Y%m%d_%H%M%S"))
                        with open(attach_path, "wb") as f:
                            f.write(df_to_excel_bytes(st.session_state["last_rups_df"]))
                    try:
                        send_via_outlook(subject, html_body, to=to, cc=cc,
                                         attachment_path=attach_path, display_only=True)
                        st.success("Outlook draft opened — review and send.",
                                   icon=":material/check_circle:")
                    except Exception as e:  # noqa: BLE001
                        st.error("Could not open Outlook: %s" % e, icon=":material/error:")


