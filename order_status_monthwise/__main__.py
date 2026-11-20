"""CLI:  python -m order_status_monthwise OrderStatusRep-CBE_RODTEP-_CBE_EXP_20.11.2026.xlsx [-d 20-11-2026] [-o out.xlsx]"""
import argparse
import datetime as dt
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from .build import build

DEFAULT_TEMPLATE = Path(__file__).resolve().parent.parent / "templates" / "Order_Status_Monthwise_template.xlsx"


def date_from_name(name):
    m = re.search(r"(\d{2})[.\-](\d{2})[.\-](\d{4})", name)
    if not m:
        raise SystemExit("Cannot find a dd.mm.yyyy date in the file name; pass --date dd-mm-yyyy")
    return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))


def recalc(path):
    """Recalculate with LibreOffice so formula values are cached in the file."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("LibreOffice not found - skipping recalculation (Excel will calculate on open).")
        return
    with tempfile.TemporaryDirectory() as td:
        subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", td, str(path)],
                       check=True, stdout=subprocess.DEVNULL)
        shutil.move(str(Path(td) / Path(path).name), str(path))


def main(argv=None):
    ap = argparse.ArgumentParser(description="Order Status month-wise report")
    ap.add_argument("input", help="ERP file OrderStatusRep(ort)-CBE_RODTEP-_CBE_EXP_<date>.xlsx")
    ap.add_argument("-d", "--date", help="report date dd-mm-yyyy (default: taken from the file name)")
    ap.add_argument("-o", "--output", help="output path (default: Order_Status_Monthwise_<date>.xlsx)")
    ap.add_argument("-t", "--template", default=str(DEFAULT_TEMPLATE), help="format template workbook (sheet Nov-26)")
    ap.add_argument("--no-recalc", action="store_true", help="skip LibreOffice recalculation")
    a = ap.parse_args(argv)
    d = dt.datetime.strptime(a.date, "%d-%m-%Y").date() if a.date else date_from_name(Path(a.input).name)
    out = a.output or f"Order_Status_Monthwise_{d:%d-%m-%Y}.xlsx"
    stats = build(a.input, d, out, a.template)
    if not a.no_recalc:
        recalc(out)
    print("written", out, stats)


if __name__ == "__main__":
    sys.exit(main())
