"""Flask app: Reject Management (EIMS <-> RUPS reconciliation + ATMf ticketing).

Single-page UI with a JSON API. Backend logic lives in core.py.

Run:
    C:\\Users\\ngocluup\\AppData\\Local\\miniforge3\\envs\\ngocluup\\python.exe web\\app.py
"""
import os
import math
import threading
import datetime as _dt

from flask import Flask, jsonify, render_template, request, send_file

import core

app = Flask(__name__)
# Reject NaN/Infinity so the browser can always JSON.parse the response.
app.config["JSON_ALLOW_NAN"] = False


def _clean(obj):
    """Recursively replace NaN/Infinity floats with None (valid JSON)."""
    if isinstance(obj, float):
        return None if (math.isnan(obj) or math.isinf(obj)) else obj
    if isinstance(obj, dict):
        return {k: _clean(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_clean(v) for v in obj]
    return obj


def sjson(payload, status=200):
    """jsonify with NaN-safe cleaning."""
    return jsonify(_clean(payload)), status


@app.route("/")
def index():
    # Cache-buster: newest mtime of the static assets, so the browser always
    # fetches the latest app.js / style.css after an edit.
    ver = 0
    for fn in ("app.js", "style.css"):
        fp = os.path.join(app.static_folder, fn)
        if os.path.exists(fp):
            ver = max(ver, int(os.path.getmtime(fp)))
    return render_template(
        "index.html",
        asset_ver=ver,
        app_env=core.APP_ENV,
        is_dev=core.IS_DEV,
        last_update=core.eims_last_update(),
        has_eims=os.path.exists(core.EIMS_TXT_PATH),
        default_filters=core.DEFAULT_FILTERS,
        critical_days=core.CRITICAL_DAYS,
        intms={
            "customers": core.INTMS_CUSTOMERS,
            "payers": core.INTMS_PAYER_BUS,
            "stages": core.INTMS_PRODUCT_STAGES,
            "test_areas": core.INTMS_TEST_AREAS,
            "factories": core.INTMS_FACTORIES,
            "requests": core.INTMS_REQUESTS,
            "defaults": core.INTMS_DEFAULTS,
            "signal": core.INTMS_SIGNAL_ID,
        },
    )


@app.route("/api/refresh", methods=["POST"])
def api_refresh():
    try:
        size = core.download_eims()
        # Also refresh loss-operation data from MARS via SQLPathFinder.
        loss_ok, loss_msg = core.refresh_loss_operation()
        core.clear_caches()
        return jsonify({"ok": True, "size": size,
                        "loss_ok": loss_ok, "loss_msg": loss_msg,
                        "last_update": core.eims_last_update()})
    except Exception as e:  # noqa: BLE001
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/options")
def api_options():
    """Filter option lists for a given view (7000 default vs 4000 all)."""
    df = core.load_eims_df()
    if df is None:
        return jsonify({"ok": False, "error": "No EIMS file. Refresh first."}), 400
    view = request.args.get("view", "default")
    if view == "op4000":
        preset = {"Operation": ["4000"]}
        all_products = True
    else:
        preset = core.DEFAULT_FILTERS
        all_products = False
    opts = core.filter_options(df)
    return jsonify({
        "ok": True,
        "options": opts,
        "preset": preset,
        "all_products": all_products,
    })


@app.route("/api/summary", methods=["POST"])
def api_summary():
    df = core.load_eims_df()
    if df is None:
        return jsonify({"ok": False, "error": "No EIMS file. Refresh first."}), 400
    body = request.get_json(force=True) or {}
    products = body.get("products") or []
    selected = body.get("filters") or {}
    filtered = core.apply_filters(df, products, selected)
    return sjson({
        "ok": True,
        "total_rows": int(len(filtered)),
        "products": core.product_summary(filtered),
        "n_products": int(filtered["Prodgroup3"].nunique()
                          if "Prodgroup3" in filtered.columns else 0),
    })


@app.route("/api/reconcile", methods=["POST"])
def api_reconcile():
    df = core.load_eims_df()
    if df is None:
        return jsonify({"ok": False, "error": "No EIMS file. Refresh first."}), 400
    body = request.get_json(force=True) or {}
    products = body.get("products") or []
    selected = body.get("filters") or {}
    scope = body.get("scope", "all")
    filtered = core.apply_filters(df, products, selected)
    try:
        result = core.reconcile(filtered, scope=scope)
    except Exception as e:  # noqa: BLE001
        import traceback
        traceback.print_exc()
        return jsonify({"ok": False, "error": "RUPS/reconcile error: %s" % e}), 500
    return sjson(result)


@app.route("/api/lot/search", methods=["POST"])
def api_lot_search():
    body = request.get_json(force=True) or {}
    query = body.get("query") or ""
    try:
        return sjson(core.search_lots(query))
    except Exception as e:  # noqa: BLE001
        import traceback
        traceback.print_exc()
        return jsonify({"ok": False, "error": "Lot search failed: %s" % e}), 500


@app.route("/api/atmf/products")
def api_atmf_products():
    return jsonify({"ok": True, "products": core.load_intms_products()})


@app.route("/api/atmf/guess", methods=["POST"])
def api_atmf_guess():
    body = request.get_json(force=True) or {}
    prods = core.load_intms_products()
    out = {}
    for pg3 in body.get("prodgroups", []):
        out[pg3] = core.guess_atmf_product(pg3, prods)
    return jsonify({"ok": True, "map": out})


@app.route("/api/atmf/submitted")
def api_atmf_submitted():
    """Lots that already went out on a ticket, so the UI can flag them."""
    return sjson({"ok": True, "submitted": core.submitted_map()})


@app.route("/api/atmf/forget", methods=["POST"])
def api_atmf_forget():
    """Clear the submitted marker for some lots (ticket cancelled / redo)."""
    body = request.get_json(force=True) or {}
    lots = body.get("lots") or []
    removed = core.forget_submissions(lots)
    return jsonify({"ok": True, "removed": removed})


@app.route("/api/atmf/submit", methods=["POST"])
def api_atmf_submit():
    body = request.get_json(force=True) or {}
    tickets = body.get("tickets") or []
    if not tickets:
        return jsonify({"ok": False, "error": "No tickets."}), 400
    try:
        results = core.submit_tickets(tickets)
        return jsonify({"ok": True, "results": results})
    except Exception as e:  # noqa: BLE001
        return jsonify({"ok": False, "error": str(e)}), 500


# ---------------------------------------------------------------- scheduler
# Auto-refresh EIMS once a day at 07:00 Vietnam time (UTC+7, no DST).
VN_TZ = _dt.timezone(_dt.timedelta(hours=7))
DAILY_REFRESH_HOUR = 7


def _seconds_until_next_run():
    now = _dt.datetime.now(VN_TZ)
    target = now.replace(hour=DAILY_REFRESH_HOUR, minute=0, second=0, microsecond=0)
    if target <= now:
        target += _dt.timedelta(days=1)
    return (target - now).total_seconds()


def _prewarm():
    """Pre-compute the default reconciliation so nobody waits on RUPS.

    Runs both scopes for the default (Op 7000) view. Cached for the day.
    """
    try:
        df = core.load_eims_df()
        if df is None:
            return
        products = core.DEFAULT_FILTERS.get("Prodgroup3", [])
        selected = {k: v for k, v in core.DEFAULT_FILTERS.items() if k != "Prodgroup3"}
        filtered = core.apply_filters(df, products, selected)
        for scope in ("all", "critical"):
            core.reconcile(filtered, scope=scope)
        print("[prewarm] reconciliation cached for the day.")
    except Exception as e:  # noqa: BLE001
        print("[prewarm] failed: %s" % e)


def _daily_refresh_loop():
    import time as _t
    while True:
        _t.sleep(_seconds_until_next_run())
        try:
            core.download_eims()
            loss_ok, loss_msg = core.refresh_loss_operation()
            core.clear_caches()
            print("[scheduler] EIMS auto-refreshed at %s VN (loss: %s)"
                  % (_dt.datetime.now(VN_TZ).strftime("%Y-%m-%d %H:%M:%S"), loss_msg))
            _prewarm()  # rebuild the daily cache right away
        except Exception as e:  # noqa: BLE001
            print("[scheduler] auto-refresh failed: %s" % e)
        _t.sleep(60)  # guard so an early wake doesn't double-fire


def start_daily_refresh():
    """Start the once-a-day 07:00 VN refresh thread (idempotent).

    Only the production instance schedules refreshes. A dev instance must never
    run this, or two processes would download EIMS and drive SQLPathFinder at
    the same time.
    """
    if core.IS_DEV:
        print("[scheduler] DEV mode - daily refresh disabled.")
        return
    if getattr(start_daily_refresh, "_started", False):
        return
    start_daily_refresh._started = True
    t = threading.Thread(target=_daily_refresh_loop, name="daily-eims-refresh",
                         daemon=True)
    t.start()
    print("[scheduler] daily EIMS refresh armed for 07:00 VN "
          "(next in %.0f min)" % (_seconds_until_next_run() / 60.0))


if __name__ == "__main__":
    import sys
    default_port = int(os.environ.get("RM_PORT") or (8601 if core.IS_DEV else 8600))
    port = int(sys.argv[1]) if len(sys.argv) > 1 else default_port
    start_daily_refresh()  # auto-refresh EIMS once a day at 07:00 VN time
    threading.Thread(target=_prewarm, name="prewarm", daemon=True).start()
    try:
        from waitress import serve
        print("Reject Management [%s] running at http://localhost:%d"
              % (core.APP_ENV.upper(), port))
        serve(app, host="0.0.0.0", port=port, threads=8)
    except ImportError:
        app.run(host="0.0.0.0", port=port, debug=False)
