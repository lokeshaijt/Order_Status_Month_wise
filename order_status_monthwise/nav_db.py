"""Fetch the Order Status data straight from the NAV SQL Server database.

The SQL lives in nav_query.sql (repo root, editable). It must return ONE ROW PER ORDER LINE
(or one row per order line per planned week) with these column aliases:

    ord_id, buyer_id, buyer_name, nav_doc_no, buyer_contact_no, shipment_date,
    line_key, prod_name, pending_qty, stock_in_process, ready_stock, pending_prod,
    can_pack_cfc, planned_week, planned_qty

  * line_key      - unique per order line (e.g. doc no + line no). Rows with the same line_key are
                    merged; their planned_week/planned_qty go into the "Planned in NN" columns.
  * planned_week  - ISO week number (NULL if nothing planned), planned_qty - qty planned that week.
  * shipment_date - Buyer Requested Shipment Date (date/datetime).

The rows are written into a workbook with the SAME layout as the ERP export (sheet
"Order Status-By shipment Date", header row 6, data from row 7, a subtotal row per order),
so build() formats it exactly as it formats the ERP file.

    python -m order_status_monthwise.nav_db -c config.ini                  # fetch -> ERP-style xlsx
    python -m order_status_monthwise.nav_db -c config.ini --compare REAL.xlsx   # check against a real export
"""
import argparse
import configparser
import datetime as dt
import os
from collections import OrderedDict
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter as L

SHEET = "Order Status-By shipment Date"
HEADER_ROW, FIRST_ROW = 6, 7
REQUIRED = ["ord_id", "buyer_id", "buyer_name", "nav_doc_no", "buyer_contact_no", "shipment_date",
            "line_key", "prod_name", "pending_qty", "stock_in_process", "ready_stock", "pending_prod",
            "can_pack_cfc", "planned_week", "planned_qty"]
# ERP column layout (1-based). Columns 7-9 and 17-21 are not shown in the report; build() drops them.
FIXED_HEADERS = {
    1: "Ord Id", 2: "Buyer Id", 3: "Buyer Name", 4: "NAV Doc No", 5: "Buyer Contact No",
    6: "Buyer Requested Shipment Date", 7: "Container No", 8: "Shipment No", 9: "Ord Qty",
    10: "Prod Name", 11: "Pending Ord Qty", 12: "Stock In Process", 13: "Ready Stock",
    14: "Stock In Process + Ready Stock", 15: "Pending Prod", 16: "CAN PACK (IN CFC)",
    17: "Posted Qty", 18: "Location", 19: "Vessel", 20: "ETD", 21: "Remarks",
}
FIRST_PLANNED_COL = 22
NAV_BLANK_DATE = dt.date(1900, 1, 1)  # NAV stores "no date" as 1753-01-01


# ---------------------------------------------------------------- database
def db_password(cfg):
    sec = cfg["database"]
    user = sec.get("username", "")
    try:
        import keyring
        pw = keyring.get_password("order_status_db", user)
        if pw:
            return pw
    except Exception:
        pass
    pw = os.environ.get("NAV_DB_PASSWORD")
    if not pw:
        raise RuntimeError("No database password: run `python auto_report.py --set-db-password` "
                           "or set NAV_DB_PASSWORD (or use auth = windows).")
    return pw


def connection_string(cfg):
    sec = cfg["database"]
    parts = [
        f"DRIVER={{{sec.get('driver', 'ODBC Driver 18 for SQL Server')}}}",
        f"SERVER={sec['server']}",
        f"DATABASE={sec['database']}",
        "ApplicationIntent=ReadOnly",
        f"Encrypt={'yes' if sec.getboolean('encrypt', True) else 'no'}",
        f"TrustServerCertificate={'yes' if sec.getboolean('trust_server_certificate', True) else 'no'}",
    ]
    if sec.get("auth", "windows").lower() == "windows":
        parts.append("Trusted_Connection=yes")
    else:
        parts += [f"UID={sec['username']}", f"PWD={db_password(cfg)}"]
    return ";".join(parts) + ";"


def fetch_rows(cfg, query_path=None):
    """Run nav_query.sql and return a list of dicts (lower-case column names)."""
    import pyodbc  # imported here so the Streamlit app does not need it

    sec = cfg["database"]
    query_path = Path(query_path or sec.get("query_file", Path(__file__).resolve().parent.parent / "nav_query.sql"))
    sql = query_path.read_text(encoding="utf-8").replace("{company}", sec.get("company", ""))
    with pyodbc.connect(connection_string(cfg), timeout=30, autocommit=True) as conn:
        conn.timeout = sec.getint("query_timeout_seconds", 300)
        cur = conn.cursor()
        cur.execute(sql)
        cols = [c[0].lower() for c in cur.description]
        missing = [c for c in REQUIRED if c not in cols]
        if missing:
            raise ValueError(f"{query_path.name} does not return column(s): {', '.join(missing)}")
        return [dict(zip(cols, r)) for r in cur.fetchall()]


# ---------------------------------------------------------------- rows -> ERP-style workbook
def _num(v):
    if v is None or v == "":
        return 0
    f = float(v)
    return int(f) if f.is_integer() else round(f, 4)


def _date(v):
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        v = v.date()
    if isinstance(v, str):
        for f in ("%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d %H:%M:%S"):
            try:
                v = dt.datetime.strptime(v[:19], f).date()
                break
            except ValueError:
                continue
        else:
            return None
    return None if v < NAV_BLANK_DATE else v


def _text(v):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return s or None


def _id(v):
    """Keep numeric ids numeric (as in the ERP export), everything else as trimmed text."""
    if isinstance(v, (int, float)) and not isinstance(v, bool) and float(v).is_integer():
        return int(v)
    if hasattr(v, "is_integer") or type(v).__name__ == "Decimal":
        f = float(v)
        return int(f) if f.is_integer() else f
    return _text(v)


def write_erp_workbook(rows, day, path):
    """Write `rows` (from fetch_rows) as an ERP-layout workbook. Returns a stats dict."""
    lines = OrderedDict()
    for r in rows:
        key = _text(r["line_key"])
        ln = lines.get(key)
        if ln is None:
            ln = lines[key] = dict(r, weeks={})
        wk, q = r.get("planned_week"), _num(r.get("planned_qty"))
        if wk not in (None, "") and q:
            ln["weeks"][int(wk)] = ln["weeks"].get(int(wk), 0) + q

    orders = OrderedDict()
    skipped = []
    for ln in lines.values():
        doc = _text(ln["nav_doc_no"])
        d = _date(ln["shipment_date"])
        if d is None:
            skipped.append(doc)
            continue
        o = orders.setdefault(doc, {"head": ln, "date": d, "lines": []})
        o["date"] = min(o["date"], d)
        o["lines"].append(ln)
    skipped = sorted(set(skipped) - set(orders))

    # planned-week columns: every week that has a planned qty, plus the current week, in date order
    cur_year, cur_week, _ = day.isocalendar()
    weeks = {cur_week} | {w for ln in lines.values() for w in ln["weeks"]}
    # weeks numerically lower than the current week are taken to be next year's
    weeks = sorted(weeks, key=lambda w: (w < cur_week, w))
    wcol = {w: FIRST_PLANNED_COL + i for i, w in enumerate(weeks)}
    c_total = FIRST_PLANNED_COL + len(weeks)
    c_unpl = c_total + 1

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET
    ws["A1"] = "Order Status-By shipment Date"
    ws["A2"] = f"Generated from NAV database on {dt.datetime.now():%d-%m-%Y %H:%M} for report date {day:%d-%m-%Y}"
    for c, h in FIXED_HEADERS.items():
        ws.cell(HEADER_ROW, c, h)
    for w, c in wcol.items():
        ws.cell(HEADER_ROW, c, f"Planned in {w:02d}")
    ws.cell(HEADER_ROW, c_total, "Planned Qty Total")
    ws.cell(HEADER_ROW, c_unpl, "Un Planned")

    row = FIRST_ROW
    n_lines = 0
    first_pl, last_pl = L(FIRST_PLANNED_COL), L(c_total - 1)
    for doc, o in sorted(orders.items(), key=lambda kv: (kv[1]["date"], kv[0])):
        h = o["head"]
        first = row
        for i, ln in enumerate(o["lines"]):
            if i == 0:
                ws.cell(row, 1, _id(h["ord_id"]))
                ws.cell(row, 2, _id(h["buyer_id"]))
                ws.cell(row, 3, _text(h["buyer_name"]))
                ws.cell(row, 4, doc)
                ws.cell(row, 5, _text(h["buyer_contact_no"]))
                ws.cell(row, 6, f"{o['date']:%d-%m-%Y}")
            pending = _num(ln["pending_qty"])
            ws.cell(row, 9, pending)                       # product-row marker (hidden column)
            ws.cell(row, 10, _text(ln["prod_name"]))
            ws.cell(row, 11, pending)
            ws.cell(row, 12, _num(ln["stock_in_process"]))
            ws.cell(row, 13, _num(ln["ready_stock"]))
            ws.cell(row, 14, f'=IFERROR(L{row}+M{row},"")')
            pp = _num(ln["pending_prod"])
            ws.cell(row, 15, pp)
            ws.cell(row, 16, _num(ln["can_pack_cfc"]))
            for w, c in wcol.items():
                ws.cell(row, c, ln["weeks"].get(w, 0))
            ws.cell(row, c_total, f"=SUM({first_pl}{row}:{last_pl}{row})")
            ws.cell(row, c_unpl, f"=({pp}-{L(c_total)}{row})")
            row += 1
            n_lines += 1
        # subtotal row (same shape as the ERP export)
        ws.cell(row, 9, " ")
        for c in list(range(11, 16)) + list(range(FIRST_PLANNED_COL, c_unpl + 1)):
            ws.cell(row, c, f"=SUM({L(c)}{first}:{L(c)}{row - 1})")
        row += 1
    wb.save(path)
    stats = {"orders": len(orders), "lines": n_lines, "weeks": [f"{w:02d}" for w in weeks],
             "skipped_no_date": skipped}
    if skipped:
        print(f"WARNING: {len(skipped)} order(s) without a shipment date were left out: {', '.join(skipped[:20])}")
    return stats


def fetch_to_workbook(cfg, day, path):
    rows = fetch_rows(cfg)
    stats = write_erp_workbook(rows, day, path)
    stats["rows"] = len(rows)
    return stats


# ---------------------------------------------------------------- compare against a real ERP export
def _read_lines(path):
    """{(nav_doc_no, prod_name): {...values}} from an ERP-layout workbook, located by header names."""
    from .parse import blocks

    ws = openpyxl.load_workbook(path)[SHEET]
    hdr = {str(ws.cell(HEADER_ROW, c).value).strip(): c for c in range(1, ws.max_column + 1)}
    need = ["NAV Doc No", "Buyer Requested Shipment Date", "Prod Name", "Pending Ord Qty",
            "Stock In Process", "Ready Stock", "Pending Prod", "CAN PACK (IN CFC)"]
    miss = [n for n in need if n not in hdr]
    if miss:
        raise ValueError(f"{Path(path).name}: header(s) not found in row {HEADER_ROW}: {miss}")
    plan = [c for n, c in hdr.items() if n.startswith("Planned in")]
    out = {}
    for b in blocks(ws):
        doc = date = None
        for r in b["rows"]:
            doc = doc or _text(ws.cell(r, hdr["NAV Doc No"]).value)
            date = date or _text(ws.cell(r, hdr["Buyer Requested Shipment Date"]).value)
        for r in b["rows"]:
            v = lambda n: _num(ws.cell(r, hdr[n]).value) if not str(ws.cell(r, hdr[n]).value or "").startswith("=") else None
            key = (doc, _text(ws.cell(r, hdr["Prod Name"]).value))
            rec = {"date": date, "pending": v("Pending Ord Qty"), "sip": v("Stock In Process"),
                   "ready": v("Ready Stock"), "pending_prod": v("Pending Prod"), "can_pack": v("CAN PACK (IN CFC)"),
                   "planned": sum(_num(ws.cell(r, c).value) for c in plan
                                  if not str(ws.cell(r, c).value or "").startswith("="))}
            if key in out:  # same product twice on one order: add up
                for k in rec:
                    if k != "date" and rec[k] is not None:
                        out[key][k] = (out[key][k] or 0) + rec[k]
            else:
                out[key] = rec
    return out


def compare(real_path, db_path, limit=50):
    real, db = _read_lines(real_path), _read_lines(db_path)
    only_real = sorted(set(real) - set(db), key=str)
    only_db = sorted(set(db) - set(real), key=str)
    diffs = []
    for k in sorted(set(real) & set(db), key=str):
        for f in real[k]:
            if real[k][f] != db[k][f]:
                diffs.append((k, f, real[k][f], db[k][f]))
    print(f"ERP export : {len(real)} lines   database: {len(db)} lines")
    print(f"Only in ERP export ({len(only_real)}):")
    for k in only_real[:limit]:
        print("   ", k)
    print(f"Only in database ({len(only_db)}):")
    for k in only_db[:limit]:
        print("   ", k)
    print(f"Value differences ({len(diffs)}):")
    for k, f, a, b in diffs[:limit]:
        print(f"    {k}  {f}: ERP={a}  DB={b}")
    ok = not (only_real or only_db or diffs)
    print("MATCH - the database query reproduces the ERP export." if ok else "NOT MATCHING yet - fix nav_query.sql.")
    return ok


def discover(cfg):
    """Log in and list databases, NAV companies and the tables/fields the report most likely uses."""
    import pyodbc

    sec = cfg["database"]
    print("Connecting to", sec["server"], "as", sec.get("username") or "(Windows login)")
    with pyodbc.connect(connection_string(cfg), timeout=30, autocommit=True) as conn:
        cur = conn.cursor()
        print("\nDatabases you can see:")
        for (n,) in cur.execute("SELECT name FROM sys.databases WHERE database_id > 4 ORDER BY name"):
            print("   ", n)
        print(f"\nIn database [{sec['database']}]:")
        try:
            comps = [r[0] for r in cur.execute("SELECT [Name] FROM [Company] ORDER BY [Name]")]
            print("  NAV companies (put the right one in [database] company):")
            for c in comps:
                print("   ", c)
        except pyodbc.Error:
            print("  (no [Company] table - is this the NAV database?)")
        pat = ["%Sales Header", "%Sales Line", "%Customer", "%Item", "%Prod_ Order Line", "%Production%Plan%",
               "%Order Status%", "%Planning%"]
        where = " OR ".join("TABLE_NAME LIKE ?" for _ in pat)
        tables = [r[0] for r in cur.execute(
            f"SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE {where} ORDER BY TABLE_NAME", *pat)]
        print("\n  Tables that look relevant:")
        for t in tables:
            print("   ", t)
        comp = sec.get("company", "")
        for t in ("Sales Header", "Sales Line"):
            full = f"{comp}${t}" if comp else None
            if full and full in tables:
                cols = [r[0] for r in cur.execute(
                    "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = ? ORDER BY ORDINAL_POSITION", full)]
                print(f"\n  Columns of [{full}] ({len(cols)}):")
                print("    " + ", ".join(cols))


def main(argv=None):
    ap = argparse.ArgumentParser(description="Fetch Order Status data from the NAV database")
    ap.add_argument("--discover", action="store_true", help="list databases, companies and relevant tables")
    ap.add_argument("-c", "--config", default=str(Path(__file__).resolve().parent.parent / "config.ini"))
    ap.add_argument("-d", "--date", help="report date dd-mm-yyyy (default today)")
    ap.add_argument("-o", "--output", help="ERP-style xlsx to write (default ERP_from_DB_<date>.xlsx)")
    ap.add_argument("--compare", metavar="REAL_XLSX", help="compare with a real ERP export of the same day")
    a = ap.parse_args(argv)
    cfg = configparser.ConfigParser(interpolation=None)
    if not cfg.read(a.config, encoding="utf-8"):
        raise SystemExit(f"Config file not found: {a.config}")
    if a.discover:
        discover(cfg)
        return 0
    day = dt.datetime.strptime(a.date, "%d-%m-%Y").date() if a.date else dt.date.today()
    out = a.output or f"ERP_from_DB_{day:%d-%m-%Y}.xlsx"
    print("written", out, fetch_to_workbook(cfg, day, out))
    if a.compare:
        return 0 if compare(a.compare, out) else 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
