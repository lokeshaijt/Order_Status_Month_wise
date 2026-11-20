# Order Status – Month-wise report

Input : ERP export `OrderStatusReport-CBE_RODTEP-_CBE_EXP_<dd.mm.yyyy>.xlsx` (sheet *Order Status-By shipment Date*)
Output: `Order_Status_Monthwise_<dd-mm-yyyy>.xlsx`

## Run
```bash
pip install -r requirements.txt
python -m order_status_monthwise "OrderStatusRep-CBE_RODTEP-_CBE_EXP_20.11.2026.xlsx"
# optional: -d 20-11-2026  -o out/report.xlsx  --no-recalc
```
LibreOffice (`soffice`) is used, if installed, to cache formula values in the output.

## Rules
- Split by *Buyer Requested Shipment Date*: each past month with pending orders, running month, +1, +2. Later orders are left out (count printed).
- Format (widths, borders, fonts, heights) is copied from sheet `Nov-26` of `templates/Order_Status_Monthwise_template.xlsx`. To change the format, edit that sheet and commit it.
- Buyer Contact No shown; CAN PACK (IN CFC) subtotal on every order block and on TOTAL; planned-week columns detected from the ERP headers (`Planned in NN`, `Planned Qty Total`, `Un Planned`).
- Row 1 shows `No. of Sales Order` = distinct NAV Doc No on the sheet.
- ERP formulas are re-pointed to the report's columns; if the ERP references a column that is not in the report, the build fails loudly.
