#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Auto-run variant: measure internet speed + submit to YOUR (experimental) Google Form
Runs twice daily at 07:00 and 13:30 (local machine time, Asia/Muscat has no DST).

Requirements (install once):
    pip install speedtest-cli requests

Usage:
    python submit_speed_and_send_autorun.py
Place a shortcut to this script (via python.exe) in the Windows Startup folder to start on login.
"""

import os
import csv
import time
from datetime import datetime, date, timedelta
import requests

# -------- User-configurable metadata (EDIT IF NEEDED) --------
SCHOOL_CODE = "1561"                          # 1- رمز المدرسة
SCHOOL_SECTOR = "السيب"                       # 2- قطاع المدرسة: مسقط / قريات / السيب / العامرات / بوشر / مطرح
SCHOOL_NAME = "مدرسة ابو القاسم الزهراوي"    # 3- اسم المدرسة
SERVICE_PROVIDER = "عمانتل"                   # 4- موفر الخدمة: عمانتل / أوريدو / أواصر
LINE_NUMBER = "24424428"                      # 5- رقم الخط
SERVICE_TYPE = "فايبر"                        # 6- نوع الخدمة: فايبر / الجيل الخامس G 5
DEVICE_NAME = os.environ.get("COMPUTERNAME") or "Device"  # 7- اسم الجهاز (تلقائي)

# Scheduling (24h format, local time)
SCHEDULES = [("07:00", 7, 0), ("13:30", 13, 30)]

LOG_DIR = os.path.join(os.getcwd(), "logs")
LOG_FILE = os.path.join(LOG_DIR, "speed_log.csv")

# -------- Google Form wiring (your EXPERIMENTAL form) --------
FORM_ACTION_URL = "https://docs.google.com/forms/u/2/d/e/1FAIpQLSdZgyPaDsPtm-9B9dkKEwYhpEmedTC1QtC0BvpLH9pP3Saf2g/formResponse"

ENTRY_IDS = {
    "school_code": "entry.413551861",      # 1- رمز المدرسة
    "sector": "entry.2132382965",          # 2- قطاع المدرسة
    "school_name": "entry.256332732",      # 3- اسم المدرسة
    "provider": "entry.1818381540",        # 4- موفر الخدمة
    "line_number": "entry.1211366125",     # 5- رقم الخط
    "service_type": "entry.230321102",     # 6- نوع الخدمة
    "internet_speed": "entry.1693223006",  # 7- سرعت الأنترنت
    "notes": "entry.1576525453",           # 8- ملاحظات
}

HIDDEN_BASE = {
    "fvv": "1",
    "pageHistory": "0",
    "fbzx": "-7716787293459453204",
}

ALLOWED_SECTORS = ["مسقط", "قريات", "السيب", "العامرات", "بوشر", "مطرح"]
ALLOWED_PROVIDERS = ["عمانتل", "أوريدو", "أواصر"]
ALLOWED_SERVICE_TYPES = ["فايبر", "الجيل الخامس G 5"]

# -------- Core functions --------
def measure_speed():
    try:
        import speedtest  # from speedtest-cli
    except ImportError:
        raise SystemExit("Missing dependency: speedtest-cli. Install via: pip install speedtest-cli")

    s = speedtest.Speedtest()
    s.get_best_server()
    download_mbps = round(s.download() / 1_000_000, 2)
    time.sleep(0.5)
    upload_mbps = round(s.upload(pre_allocate=False) / 1_000_000, 2)
    ping_ms = round(s.results.ping, 2)

    best = s.get_best_server()
    server_host = best.get("host", "unknown")
    ip_addr = s.results.client.get("ip", "unknown")

    return {
        "download": download_mbps,
        "upload": upload_mbps,
        "ping": ping_ms,
        "server": server_host,
        "ip": ip_addr,
    }

def ensure_log_header(path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    if not os.path.exists(path):
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([
                "timestamp", "download_mbps", "upload_mbps", "ping_ms",
                "server", "ip", "device",
                "school_code", "sector", "school_name",
                "provider", "line_number", "service_type", "schedule_label", "submit_status"
            ])

def append_log(path, row):
    with open(path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(row)

def build_payload(results, timestamp):
    if SCHOOL_SECTOR not in ALLOWED_SECTORS:
        raise ValueError(f"Invalid sector '{SCHOOL_SECTOR}'. Allowed: {ALLOWED_SECTORS}")
    if SERVICE_PROVIDER not in ALLOWED_PROVIDERS:
        raise ValueError(f"Invalid provider '{SERVICE_PROVIDER}'. Allowed: {ALLOWED_PROVIDERS}")
    if SERVICE_TYPE not in ALLOWED_SERVICE_TYPES:
        raise ValueError(f"Invalid service type '{SERVICE_TYPE}'. Allowed: {ALLOWED_SERVICE_TYPES}")

    notes_text = (
        f"تنزيل: {results['download']} Mbps | "
        f"رفع: {results['upload']} Mbps | "
        f"Ping: {results['ping']} ms | "
        f"السيرفر: {results['server']} | "
        f"IP: {results['ip']} | "
        f"الجهاز: {DEVICE_NAME} | "
        f"التاريخ/الوقت: {timestamp}"
    )

    payload = {
        ENTRY_IDS["school_code"]: str(SCHOOL_CODE),
        ENTRY_IDS["sector"]: SCHOOL_SECTOR,
        ENTRY_IDS["school_name"]: SCHOOL_NAME,
        ENTRY_IDS["provider"]: SERVICE_PROVIDER,
        ENTRY_IDS["line_number"]: str(LINE_NUMBER),
        ENTRY_IDS["service_type"]: SERVICE_TYPE,
        ENTRY_IDS["internet_speed"]: f"{results['download']} Mbps",
        ENTRY_IDS["notes"]: notes_text,
    }
    payload.update(HIDDEN_BASE)
    return payload

def submit_form(payload):
    headers = {
        "Content-Type": "application/x-www-form-urlencoded;charset=utf-8",
        "User-Agent": "Mozilla/5.0",
        "Referer": "https://docs.google.com/forms/d/e/1FAIpQLSdZgyPaDsPtm-9B9dkKEwYhpEmedTC1QtC0BvpLH9pP3Saf2g/viewform",
    }

    variants = [payload, {k: v for k, v in payload.items() if k != "fbzx"}]
    last_code, last_preview = None, ""

    for variant in variants:
        try:
            resp = requests.post(FORM_ACTION_URL, data=variant, headers=headers, timeout=30)
            last_code = resp.status_code
            last_preview = resp.text[:500]
            if resp.status_code in (200, 302):
                return True, last_code, last_preview
        except Exception as e:
            last_preview = str(e)
    return False, last_code, last_preview

def run_once(schedule_label):
    """Measure, log, submit once."""
    print(f"\n[{schedule_label}] بدء القياس والإرسال...")
    results = measure_speed()
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print("نتائج السرعة: "
          f"تنزيل {results['download']} Mbps | "
          f"رفع {results['upload']} Mbps | "
          f"Ping {results['ping']} ms | "
          f"سيرفر {results['server']} | IP {results['ip']}")

    ensure_log_header(LOG_FILE)

    payload = build_payload(results, ts)
    ok, code, preview = submit_form(payload)
    status_txt = "SUCCESS" if ok else f"FAIL({code})"
    append_log(LOG_FILE, [
        ts, results["download"], results["upload"], results["ping"],
        results["server"], results["ip"], DEVICE_NAME,
        SCHOOL_CODE, SCHOOL_SECTOR, SCHOOL_NAME,
        SERVICE_PROVIDER, LINE_NUMBER, SERVICE_TYPE, schedule_label, status_txt
    ])

    if ok:
        print(f"[{schedule_label}] تم الإرسال بنجاح ✅ (HTTP {code})")
    else:
        print(f"[{schedule_label}] فشل الإرسال ❌ (HTTP {code})")
        print("Preview:", preview)

def next_dt_for(hour, minute, today=None):
    if today is None:
        today = date.today()
    candidate = datetime.combine(today, datetime.min.time()).replace(hour=hour, minute=minute)
    now = datetime.now()
    if candidate <= now:
        candidate += timedelta(days=1)
    return candidate

def loop_scheduler():
    print("سيعمل السكربت تلقائيًا مرتين يوميًا: 07:00 و 13:30.")
    print("اترك النافذة مفتوحة أو شغل السكربت ضمن الخلفية.")

    # Track last run date for each label to avoid duplicates after sleep/wake
    last_run = {label: None for (label, _, _) in SCHEDULES}

    while True:
        now = datetime.now()
        ran_any = False

        # Catch-up: if a scheduled time already passed today and not run yet, run it now
        for (label, h, m) in SCHEDULES:
            sched_today = datetime.combine(date.today(), datetime.min.time()).replace(hour=h, minute=m)
            if now >= sched_today and last_run[label] != date.today():
                run_once(label)
                last_run[label] = date.today()
                ran_any = True

        if ran_any:
            # After catch-up, short nap then continue
            time.sleep(30)
            continue

        # Otherwise, sleep until the next scheduled event
        upcoming = []
        for (label, h, m) in SCHEDULES:
            sched_today = datetime.combine(date.today(), datetime.min.time()).replace(hour=h, minute=m)
            if last_run[label] == date.today():
                # already done today; consider tomorrow's occurrence
                sched = next_dt_for(h, m, date.today())
            else:
                # not yet run today; if time passed, next_dt_for will pick tomorrow
                sched = next_dt_for(h, m, date.today())
                # If we're still before today's time and not run, choose today's
                if datetime.now() < sched_today:
                    sched = sched_today
            upcoming.append((sched, label))

        next_time, next_label = min(upcoming, key=lambda x: x[0])
        wait_seconds = max(5, int((next_time - datetime.now()).total_seconds()))
        # sleep in chunks to be responsive to system clock changes
        print(f"ينام حتى {next_time.strftime('%Y-%m-%d %H:%M')} للحدث [{next_label}] (~{wait_seconds//60} دقيقة).")
        while wait_seconds > 0:
            nap = min(wait_seconds, 300)  # sleep up to 5 minutes per chunk
            time.sleep(nap)
            wait_seconds = int((next_time - datetime.now()).total_seconds())

        # Time reached; run the job
        run_once(next_label)
        last_run[next_label] = date.today()

def main():
    loop_scheduler()

if __name__ == "__main__":
    main()
