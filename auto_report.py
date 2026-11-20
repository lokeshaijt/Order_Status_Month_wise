"""Daily automation: pick up the ERP export NAV saved, build the month-wise report, e-mail it.

    python auto_report.py                 # normal daily run (Task Scheduler)
    python auto_report.py --dry-run       # build + send only to [alerts] error_to
    python auto_report.py --no-email      # build only
    python auto_report.py --date 20-11-2026 --force   # re-run / re-send a given day

Settings are read from config.ini (copy config.example.ini). The SMTP password is read from
Windows Credential Manager (via `keyring`) or the O365_SMTP_PASSWORD environment variable.
"""
import argparse
import configparser
import datetime as dt
import glob
import logging
import os
import shutil
import smtplib
import sys
import time
import traceback
from email.message import EmailMessage
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from order_status_monthwise.__main__ import DEFAULT_TEMPLATE, date_from_name, recalc  # noqa: E402
from order_status_monthwise.build import build  # noqa: E402

XLSX_MIME = ("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
log = logging.getLogger("auto_report")


# ---------------------------------------------------------------- config / helpers
def load_config(path):
    cfg = configparser.ConfigParser(interpolation=None)
    if not cfg.read(path, encoding="utf-8"):
        raise SystemExit(f"Config file not found: {path} (copy config.example.ini to config.ini)")
    return cfg


def addr_list(value):
    return [a.strip() for a in (value or "").replace(";", ",").split(",") if a.strip()]


def smtp_password(cfg):
    user = cfg["email"]["sender"]
    try:
        import keyring  # Windows Credential Manager
        pw = keyring.get_password("order_status_report", user)
        if pw:
            return pw
    except Exception:  # keyring not installed or no backend
        pass
    pw = os.environ.get("O365_SMTP_PASSWORD")
    if not pw:
        raise RuntimeError("No SMTP password: store it with `python auto_report.py --set-password` "
                           "or set the O365_SMTP_PASSWORD environment variable.")
    return pw


def setup_logging(log_dir):
    if log.handlers:
        return
    log_dir.mkdir(parents=True, exist_ok=True)
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    fh = logging.FileHandler(log_dir / "auto_report.log", encoding="utf-8")
    fh.setFormatter(fmt)
    sh = logging.StreamHandler()
    sh.setFormatter(fmt)
    log.setLevel(logging.INFO)
    log.addHandler(fh)
    log.addHandler(sh)


# ---------------------------------------------------------------- 1. find the ERP file
def find_input(cfg, day):
    """Newest file in the input folder matching the pattern whose name carries `day`."""
    sec = cfg["input"]
    pattern = str(Path(sec["folder"]) / sec.get("pattern", "*.xlsx"))
    wait_min = sec.getint("wait_minutes", 30)
    poll_min = sec.getint("poll_minutes", 5)
    deadline = time.monotonic() + wait_min * 60
    while True:
        hits = []
        for f in glob.glob(pattern):
            if Path(f).name.startswith("~$"):  # Excel lock file
                continue
            try:
                if date_from_name(Path(f).name) == day:
                    hits.append(Path(f))
            except SystemExit:
                continue
        if hits:
            newest = max(hits, key=lambda p: p.stat().st_mtime)
            # make sure NAV has finished writing it (size stable for 10 s)
            s1 = newest.stat().st_size
            time.sleep(10)
            if newest.stat().st_size == s1 and s1 > 0:
                return newest
        if time.monotonic() >= deadline:
            raise FileNotFoundError(f"No ERP export dated {day:%d.%m.%Y} found matching {pattern} "
                                    f"after waiting {wait_min} min. Check the NAV Job Queue entry.")
        log.info("ERP file for %s not there yet - checking again in %s min", day, poll_min)
        time.sleep(poll_min * 60)


# ---------------------------------------------------------------- 3. e-mail
def send_mail(cfg, to, cc, subject, body, attachment=None):
    e = cfg["email"]
    msg = EmailMessage()
    msg["From"] = e["sender"]
    msg["To"] = ", ".join(to)
    if cc:
        msg["Cc"] = ", ".join(cc)
    msg["Subject"] = subject
    msg.set_content(body)
    if attachment:
        msg.add_attachment(Path(attachment).read_bytes(), maintype=XLSX_MIME[0], subtype=XLSX_MIME[1],
                           filename=Path(attachment).name)
    with smtplib.SMTP(e.get("smtp_host", "smtp.office365.com"), e.getint("smtp_port", 587), timeout=60) as s:
        s.starttls()
        s.login(e.get("username", e["sender"]), smtp_password(cfg))
        s.send_message(msg, to_addrs=to + cc)


def fmt(template, day, **extra):
    return template.replace("\\n", "\n").format(date=f"{day:%d-%m-%Y}", **extra)


# ---------------------------------------------------------------- main run
def run(cfg, day, dry_run=False, no_email=False, force=False):
    out_dir = Path(cfg["output"]["folder"])
    out_dir.mkdir(parents=True, exist_ok=True)
    marker = out_dir / f".sent_{day:%Y-%m-%d}"
    if marker.exists() and not (force or dry_run or no_email):
        log.info("Report for %s already sent (%s) - nothing to do. Use --force to resend.", day, marker)
        return

    src = find_input(cfg, day)
    log.info("Input: %s", src)

    out = out_dir / f"Order_Status_Monthwise_{day:%d-%m-%Y}.xlsx"
    stats = build(str(src), day, str(out), str(DEFAULT_TEMPLATE))
    if cfg["output"].getboolean("recalc", True):
        recalc(out)
    log.info("Built %s %s", out, stats)

    archive = cfg["input"].get("archive_folder", "").strip()
    if archive:
        Path(archive).mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, Path(archive) / src.name)

    if no_email:
        return
    e = cfg["email"]
    if dry_run:
        to, cc = addr_list(cfg["alerts"]["error_to"]), []
    else:
        to, cc = addr_list(e["to"]), addr_list(e.get("cc", ""))
    if not to:
        raise RuntimeError("No recipients configured ([email] to).")
    summary = "\n".join(f"  {sheet}: {v['orders']} orders, {v['lines']} lines" for sheet, v in stats.items())
    subject = fmt(e.get("subject", "Order Status - Month-wise {date}"), day)
    if dry_run:
        subject = "[TEST] " + subject
    body = fmt(e.get("body", "Dear All,\n\nPlease find attached the Order Status month-wise report "
                             "as on {date}.\n\n{summary}\n\nRegards"), day, summary=summary)
    send_mail(cfg, to, cc, subject, body, out)
    log.info("Mailed to %s cc %s", to, cc)
    if not dry_run:
        marker.write_text(dt.datetime.now().isoformat())


def main(argv=None):
    ap = argparse.ArgumentParser(description="Daily Order Status month-wise report: fetch, build, e-mail")
    ap.add_argument("-c", "--config", default=str(HERE / "config.ini"))
    ap.add_argument("-d", "--date", help="report date dd-mm-yyyy (default: today)")
    ap.add_argument("--dry-run", action="store_true", help="send only to [alerts] error_to, subject prefixed [TEST]")
    ap.add_argument("--no-email", action="store_true", help="build the report, do not send")
    ap.add_argument("--force", action="store_true", help="send even if already sent for this date")
    ap.add_argument("--set-password", action="store_true", help="store the SMTP password in Windows Credential Manager")
    a = ap.parse_args(argv)

    cfg = load_config(a.config)
    if a.set_password:
        import getpass
        import keyring
        keyring.set_password("order_status_report", cfg["email"]["sender"],
                             getpass.getpass(f"Password for {cfg['email']['sender']}: "))
        print("Saved in Windows Credential Manager.")
        return 0

    setup_logging(Path(cfg["output"].get("log_folder", str(HERE / "logs"))))
    day = dt.datetime.strptime(a.date, "%d-%m-%Y").date() if a.date else dt.date.today()
    try:
        run(cfg, day, a.dry_run, a.no_email, a.force)
        return 0
    except Exception as exc:
        log.error("FAILED: %s\n%s", exc, traceback.format_exc())
        try:
            send_mail(cfg, addr_list(cfg["alerts"]["error_to"]), [],
                      f"FAILED: Order Status month-wise report {day:%d-%m-%Y}",
                      f"The automatic report did not go out.\n\nError: {exc}\n\n{traceback.format_exc()}")
        except Exception as mail_exc:
            log.error("Could not send the failure e-mail either: %s", mail_exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
