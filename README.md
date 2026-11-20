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

## Daily automatic e-mail (Windows server, 10:30 AM)

Flow: NAV Job Queue saves the Order Status report to a shared folder → Task Scheduler runs
`run_report.bat` at 10:30 → `auto_report.py` picks up today's file, builds the report and mails it
via Office 365. If anything fails (e.g. file missing), an error e-mail goes to `[alerts] error_to`.

### 1. NAV (ask the NAV admin)
Create a Job Queue entry that runs the *Order Status – By shipment Date* report daily at **10:00**
and saves it as Excel into the shared folder, file name containing the date as `dd.mm.yyyy`
(e.g. `OrderStatusRep-CBE_RODTEP-_CBE_EXP_20.11.2026.xlsx`). If the file is late, the script keeps
checking every 5 min for up to 30 min.

### 2. Office 365 (ask IT)
A mailbox to send from (e.g. `reports@…`) with **Authenticated SMTP** enabled
(Microsoft 365 admin → Users → mailbox → Mail → Manage email apps → *Authenticated SMTP*).
If MFA is enforced on that account, IT must allow SMTP AUTH for it or provide an app password.

### 3. Server setup (once)
```bat
cd D:\OrderStatusReport
git clone https://github.com/lokeshaijt/Order_Status_Month_wise.git .
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
copy config.example.ini config.ini
notepad config.ini
.venv\Scripts\python auto_report.py --set-password
.venv\Scripts\python auto_report.py --dry-run
```
`--set-password` stores the mailbox password in Windows Credential Manager (run it as the same
user the scheduled task runs as). `--dry-run` sends a `[TEST]` mail only to `error_to`.

### 4. Schedule
```bat
schtasks /Create /TN "Order Status Monthwise Report" /TR "D:\OrderStatusReport\run_report.bat" /SC DAILY /ST 10:30 /RU DOMAIN\svc_user /RP *
```
Or in Task Scheduler: *Run whether user is logged on or not*, trigger Daily 10:30.
Add `/D MON,TUE,WED,THU,FRI,SAT` with `/SC WEEKLY` instead to skip Sundays.

### Re-running
- `run_report.bat --date 20-11-2026 --force` – rebuild and resend a day
- `run_report.bat --no-email` – build only
- A report is mailed at most once per day (marker file `.sent_<date>` in the output folder).
- Log: `logs\auto_report.log`.
