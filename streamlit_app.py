"""Streamlit front end: upload the ERP export, download the month-wise report."""
import datetime as dt
import tempfile
from pathlib import Path

import streamlit as st

from order_status_monthwise.__main__ import DEFAULT_TEMPLATE, date_from_name, recalc
from order_status_monthwise.build import build

st.set_page_config(page_title="Order Status – Month-wise", page_icon="📦")
st.title("Order Status – Month-wise report")
st.caption("Upload the ERP export (sheet *Order Status-By shipment Date*) to build the month-wise report.")

up = st.file_uploader("ERP file (OrderStatusRep…CBE_EXP_dd.mm.yyyy.xlsx)", type=["xlsx"])

default_date = dt.date.today()
if up is not None:
    try:
        default_date = date_from_name(up.name)
    except SystemExit:
        st.info("No dd.mm.yyyy date found in the file name – pick the report date below.")
report_date = st.date_input("Report date", value=default_date, format="DD-MM-YYYY")

if up is not None and st.button("Build report", type="primary"):
    with tempfile.TemporaryDirectory() as td:
        src = Path(td) / up.name
        src.write_bytes(up.getvalue())
        out_name = f"Order_Status_Monthwise_{report_date:%d-%m-%Y}.xlsx"
        out = Path(td) / out_name
        try:
            with st.spinner("Building…"):
                stats = build(str(src), report_date, str(out), str(DEFAULT_TEMPLATE))
                recalc(out)  # no-op if LibreOffice is not installed
        except Exception as e:  # show the build's own error message to the user
            st.error(f"Build failed: {e}")
            st.stop()
        data = out.read_bytes()

    st.success("Report built.")
    st.table({sheet: {"orders": v["orders"], "lines": v["lines"]} for sheet, v in stats.items()})
    st.download_button(
        "Download " + out_name,
        data=data,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
