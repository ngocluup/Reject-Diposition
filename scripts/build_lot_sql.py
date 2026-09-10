# -*- coding: utf-8 -*-
"""
Build a ready-to-run SQL file from eims_lot_history.sql by injecting a lot list
read from a CSV (LotNumber column).

Usage (C:\\Python27\\python.exe build_lot_sql.py ...):
  build_lot_sql.py                         # uses ../data/EIMS_Inventory_Report.csv
  build_lot_sql.py path\\to\\lots.csv        # custom CSV (needs a LotNumber column)
  build_lot_sql.py --col Lot lots.csv      # custom lot column name

Output: scripts/eims_lot_history_<timestamp>.sql  with {{LOT_LIST}} replaced by
        'LOT1','LOT2',...  ready to paste/run in SQLPathFinder or any Oracle client.

Stdlib only, Python 2.7 / 3 compatible.
"""
import csv
import os
import sys
import time

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_SQL = os.path.join(BASE_DIR, "eims_lot_history.sql")
DATA_DIR = os.path.join(os.path.dirname(BASE_DIR), "data")
# Prefer the cached EIMS report (.txt is tab-delimited; .csv if exported).
DEFAULT_CSV = os.path.join(DATA_DIR, "EIMS_Inventory_Report.txt")
if not os.path.exists(DEFAULT_CSV):
    DEFAULT_CSV = os.path.join(DATA_DIR, "EIMS_Inventory_Report.csv")
LOT_TOKEN = "{{LOT_LIST}}"


def parse_args(argv):
    csv_path = None
    lot_col = "LotNumber"
    args = argv[1:]
    i = 0
    while i < len(args):
        a = args[i]
        if a in ("--col", "-c"):
            lot_col = args[i + 1]
            i += 2
        elif a in ("--help", "-h"):
            print(__doc__)
            sys.exit(0)
        else:
            csv_path = a
            i += 1
    if csv_path is None:
        csv_path = DEFAULT_CSV
    return csv_path, lot_col


def read_lots(csv_path, lot_col):
    if not os.path.exists(csv_path):
        sys.exit("CSV not found: %s" % csv_path)
    lots = []
    seen = {}
    # EIMS files can be tab- or comma-delimited; sniff a little.
    f = open(csv_path, "r")
    try:
        sample = f.read(4096)
        f.seek(0)
        delim = "\t" if sample.count("\t") > sample.count(",") else ","
        reader = csv.DictReader(f, delimiter=delim)
        if reader.fieldnames is None:
            sys.exit("CSV appears empty: %s" % csv_path)
        headers = [h.strip() for h in reader.fieldnames]
        if lot_col not in headers:
            sys.exit("Column '%s' not found. Available columns: %s"
                     % (lot_col, ", ".join(headers)))
        # map back to the raw fieldname (in case of surrounding spaces)
        raw_col = reader.fieldnames[headers.index(lot_col)]
        for row in reader:
            val = (row.get(raw_col) or "").strip()
            if val and val.lower() != "nan" and val not in seen:
                seen[val] = 1
                lots.append(val)
    finally:
        f.close()
    return lots


def build_sql(lots):
    tf = open(TEMPLATE_SQL, "r")
    try:
        template = tf.read()
    finally:
        tf.close()
    # Oracle limits an IN (...) list to 1000 items, so chunk into
    # (lot IN (...) OR lot IN (...) ...) when needed.
    chunk = 1000
    groups = []
    for i in range(0, len(lots), chunk):
        vals = ",".join("'%s'" % lot.replace("'", "''") for lot in lots[i:i + chunk])
        groups.append("a0.lot IN (%s)" % vals)
    predicate = groups[0] if len(groups) == 1 else "(" + " OR ".join(groups) + ")"
    # Replace the whole "a0.lot IN ({{LOT_LIST}})" predicate line.
    return template.replace("a0.lot IN (%s)" % LOT_TOKEN, predicate)


def main():
    csv_path, lot_col = parse_args(sys.argv)
    lots = read_lots(csv_path, lot_col)
    if not lots:
        sys.exit("No lots found in %s (column '%s')." % (csv_path, lot_col))
    sql = build_sql(lots)
    out_path = os.path.join(BASE_DIR, "eims_lot_history_%s.sql" % time.strftime("%Y%m%d_%H%M%S"))
    of = open(out_path, "w")
    try:
        of.write(sql)
    finally:
        of.close()
    print("CSV          : %s (column '%s')" % (csv_path, lot_col))
    print("Lots injected: %d" % len(lots))
    print("Output SQL   : %s" % out_path)


if __name__ == "__main__":
    main()
