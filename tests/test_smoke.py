import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from order_status_monthwise.__main__ import date_from_name


def test_date_from_name():
    assert str(date_from_name("OrderStatusRep-CBE_RODTEP-_CBE_EXP_20.11.2026.xlsx")) == "2026-11-20"
    assert str(date_from_name("OrderStatusReport-14-11-2026-CBE.xlsx")) == "2026-11-14"
