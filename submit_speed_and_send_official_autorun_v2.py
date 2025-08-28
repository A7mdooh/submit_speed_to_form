#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Autorun v2 (OFFICIAL form): measure internet speed + submit to Muscat Directorate Google Form
- Robust against speedtest 403:
  * speedtest.Speedtest(secure=True)
  * Retries with backoff
  * Fallback to CLI: `speedtest --json`
- Runs twice daily at 07:00 and 13:30 (local time).

Requirements:
    pip install speedtest-cli requests
Optional:
    pip install -U speedtest-cli

Usage:
    python submit_speed_and_send_official_autorun_v2.py
"""

import os
import csv
import time
import json
import subprocess
from datetime import datetime, date, timedelta
import requests

# -------- User-configurable metadata (PRE-FILLED) --------
SCHOOL_CODE = "1561"                          # 1- رمز المدرسة
SCHOOL_SECTOR = "السيب"                       # 2- قطاع المدرسة
SCHOOL_NAME = "مدرسة ابو القاسم الزهراوي"    # 3- اسم المدرسة
SERVICE_PROVIDER = "عمانتل"                   # 4- موفر الخدمة
LINE_NUMBER = "24424428"                      # 5- رقم الخط
SERVICE_TYPE = "فايبر"                        # 6- نوع الخدمة (أو 'الجيل الخامس 5 G')
DEVICE_NAME = os.environ.get("COMPUTERNAME") or "Device"  # 7- اسم الجهاز (تلقائي)

# Scheduling (24h, local time)
SCHEDULES = [("07:00", 7, 0), ("13:30", 13, 30)]

LOG_DIR = os.path.join(os.getcwd(), "logs")
LOG_FILE = os.path.join(LOG_DIR, "speed_log.csv")

# -------- OFFICIAL Google Form wiring --------
FORM_ACTION_URL = "https://docs.google.com/forms/u/2/d/e/1FAIpQLSfoYtl3gmt9FYa7g39v4az1OOtrkYHDcfAX6M-vhI6J-hX50A/formResponse"

# Select questions (IDs from *_sentinel fields)
ENTRY_SECTOR_ID = "entry.1313908626"       # 2- قطاع المدرسة
ENTRY_PROVIDER_ID = "entry.927675658"      # 4- موفر الخدمة
ENTRY_SERVICE_TYPE_ID = "entry.66731299"   # 6- نوع الخدمة

ALLOWED_SECTORS = ["مسقط", "قريات", "السيب", "العامرات", "بوشر", "مطرح"]
ALLOWED_PROVIDERS = ["عمانتل", "أوريدو", "أواصر"]
ALLOWED_SERVICE_TYPES = ["فايبر", "الجيل الخامس 5 G"]  # لاحظ المسافة بين 5 و G

# Text/textarea questions (IDs inferred from form HTML)
# Confirmed: Q1 (school code) = entry.899161738
ENTRY_TEXT_IDS = {
    "Q1_school_code": "entry.899161738",
    # Four text fields (Q3 name, Q5 line, Q7 speed text, Q8 notes):
    "X_A": "entry.560537791",
    "X_B": "entry.1862560773",
    "X_C": "entry.181224386",
    "X_D": "entry.556952249",
}

# Try a few plausible mappings for X_A..X_D to (Q3,Q5,Q7,Q8)
TEXT_MAPPING_TRIES = [
    {"Q3_school_name": "X_A", "Q5_line_number": "X_B", "Q7_internet_speed": "X_C", "Q8_notes": "X_D"},
    {"Q3_school_name": "X_A", "Q5_line_number": "X_B", "Q7_internet_speed": "X_D", "Q8_notes": "X_C"},
    {"Q3_school_name": "X_B", "Q5_line_number": "X_A", "Q7_internet_speed": "X_C", "Q8_notes": "X_D"},
]

# Hidden params (token may vary; we try with/without)
HIDDEN_CANDIDATES = [
    {},  # without fbzx
    {"fbzx": "8122308104194036559"},  # token captured from saved HTML example
]
HIDDEN_ALWAYS = {"fvv": "1", "pageHistory": "0"}

# -------- Speed measurement (robust) --------
def measure_speed_python(max_attempts=3, backoff=20):
    """Try Python API with secure=True; retry on 403 ConfigRetrievalError"""
    try:
        import speedtest  # from speedtest-cli
    except ImportError:
        raise SystemExit("Missing dependency: speedtest-cli. Install via: pip install speedtest-cli")

    last_err = None
    for attempt in range(1, max_attempts + 1):
        try:
            s = speedtest.Speedtest(secure=True)
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
        except Exception as e:
            last_err = e
            msg = str(e)
            if "403" in msg or "ConfigRetrievalError" in msg:
                wait = backoff * attempt
                print(f"[speedtest-python] محاولة {attempt}/{max_attempts} فشلت (403 محتمل). انتظر {wait}s...")
                time.sleep(wait)
            else:
                print(f"[speedtest-python] محاولة {attempt}/{max_attempts} فشلت ({e}). إعادة المحاولة خلال 5s...")
                time.sleep(5)
    raise last_err

def measure_speed_cli():
    """Fallback to CLI: speedtest --json (supports both speedtest-cli and Ookla formats)"""
    candidates = ["speedtest", "speedtest.exe"]
    for exe in candidates:
        try:
            cp = subprocess.run([exe, "--json"], capture_output=True, text=True, timeout=180)
            if cp.returncode == 0 and cp.stdout:
                data = json.loads(cp.stdout)
                # speedtest-cli JSON
                if all(k in data for k in ("download", "upload", "ping", "server", "client")):
                    download_mbps = round(float(data["download"]) / 1_000_000, 2)
                    upload_mbps = round(float(data["upload"]) / 1_000_000, 2)
                    ping_ms = round(float(data["ping"]), 2)
                    server_host = data["server"].get("host") or f"{data['server'].get('name','')}".strip()
                    ip_addr = data["client"].get("ip", "unknown")
                    return {"download": download_mbps, "upload": upload_mbps, "ping": ping_ms, "server": server_host or "unknown", "ip": ip_addr}
                # Ookla JSON
                if data.get("type") == "result":
                    dl = data.get("download", {}).get("bandwidth")
                    ul = data.get("upload", {}).get("bandwidth")
                    ping_ms = data.get("ping", {}).get("latency")
                    if dl and ul and ping_ms is not None:
                        download_mbps = round((dl * 8) / 1_000_000, 2)
                        upload_mbps = round((ul * 8) * 8 / 1_000_000, 2) if isinstance(ul, (int, float)) else None
                        server_host = (data.get("server", {}) or {}).get("host", "unknown")
                        ip_addr = (data.get("interface", {}) or {}).get("externalIp", "unknown")
                        return {"download": download_mbps, "upload": upload_mbps, "ping": round(float(ping_ms), 2), "server": server_host, "ip": ip_addr}
            else:
                print(f"[speedtest-cli] فشل التشغيل ({exe}), rc={cp.returncode}, err={cp.stderr[:200]}")
        except FileNotFoundError:
            continue
        except Exception as e:
            print(f"[speedtest-cli] استثناء: {e}")
    raise RuntimeError("تعذر تشغيل speedtest CLI. تأكد أن 'speedtest' في PATH أو ثبّت speedtest-cli.")

def measure_speed():
    try:
        return measure_speed_python()
    except Exception as e:
        print(f"[تحذير] قياس السرعة عبر بايثون فشل ({e}). المحاولة عبر CLI...")
        return measure_speed_cli()

# -------- Logging --------
def ensure_log_header(path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    if not os.path.exists(path):
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([
                "timestamp", "download_mbps", "upload_mbps", "ping_ms",
                "server", "ip", "device",
                "school_code", "sector", "school_name",
                "provider", "line_number", "service_type", "schedule_label", "submit_status", "used_mapping", "used_hidden"
            ])

def append_log(path, row):
    with open(path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(row)

# -------- Submit to Google Form --------
def build_payload_base(results, ts):
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
        f"التاريخ/الوقت: {ts}"
    )

    base = {
        ENTRY_TEXT_IDS["Q1_school_code"]: str(SCHOOL_CODE),
        ENTRY_SECTOR_ID: SCHOOL_SECTOR,
        ENTRY_PROVIDER_ID: SERVICE_PROVIDER,
        ENTRY_SERVICE_TYPE_ID: SERVICE_TYPE,
    }

    return base, notes_text

def try_submit_with_mapping(mapping, hidden_extra, results, ts):
    base, notes_text = build_payload_base(results, ts)
    q3_id = ENTRY_TEXT_IDS[mapping["Q3_school_name"]]
    q5_id = ENTRY_TEXT_IDS[mapping["Q5_line_number"]]
    q7_id = ENTRY_TEXT_IDS[mapping["Q7_internet_speed"]]
    q8_id = ENTRY_TEXT_IDS[mapping["Q8_notes"]]

    payload = dict(base)
    payload[q3_id] = SCHOOL_NAME
    payload[q5_id] = str(LINE_NUMBER)
    payload[q7_id] = f"{results['download']} Mbps"
    payload[q8_id] = notes_text
    payload.update(HIDDEN_ALWAYS)
    payload.update(hidden_extra)

    headers = {
        "Content-Type": "application/x-www-form-urlencoded;charset=utf-8",
        "User-Agent": "Mozilla/5.0",
        "Referer": "https://docs.google.com/forms/d/e/1FAIpQLSfoYtl3gmt9FYa7g39v4az1OOtrkYHDcfAX6M-vhI6J-hX50A/viewform",
    }

    r = requests.post(FORM_ACTION_URL, data=payload, headers=headers, timeout=30)
    return (r.status_code in (200, 302)), r.status_code, r.text[:500]

def submit_official(results, ts):
    # Try all mapping x hidden combinations
    last_status = (False, None, "")
    used_mapping = None
    used_hidden = None
    for mapping in TEXT_MAPPING_TRIES:
        for hidden in HIDDEN_CANDIDATES:
            ok, code, preview = try_submit_with_mapping(mapping, hidden, results, ts)
            print(f"- تجربة mapping={mapping} hidden={hidden} => Status {code}")
            last_status = (ok, code, preview)
            if ok:
                used_mapping = mapping
                used_hidden = hidden
                return True, code, preview, used_mapping, used_hidden
    return False, last_status[1], last_status[2], used_mapping, used_hidden

# -------- Scheduler helpers --------
def run_once(schedule_label):
    print(f"\n[{schedule_label}] بدء القياس والإرسال...")
    results = measure_speed()
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print("نتائج السرعة: "
          f"تنزيل {results['download']} Mbps | "
          f"رفع {results['upload']} Mbps | "
          f"Ping {results['ping']} ms | "
          f"سيرفر {results['server']} | IP {results['ip']}")

    ensure_log_header(LOG_FILE)

    ok, code, preview, used_mapping, used_hidden = submit_official(results, ts)
    status_txt = "SUCCESS" if ok else f"FAIL({code})"
    append_log(LOG_FILE, [
        ts, results["download"], results["upload"], results["ping"],
        results["server"], results["ip"], DEVICE_NAME,
        SCHOOL_CODE, SCHOOL_SECTOR, SCHOOL_NAME,
        SERVICE_PROVIDER, LINE_NUMBER, SERVICE_TYPE, schedule_label, status_txt,
        str(used_mapping), str(used_hidden)
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
    print("اترك النافذة مفتوحة أو شغّل من Task Scheduler/Startup للتشغيل الصامت.")

    last_run = {label: None for (label, _, _) in SCHEDULES}

    while True:
        now = datetime.now()
        ran_any = False

        for (label, h, m) in SCHEDULES:
            sched_today = datetime.combine(date.today(), datetime.min.time()).replace(hour=h, minute=m)
            if now >= sched_today and last_run[label] != date.today():
                run_once(label)
                last_run[label] = date.today()
                ran_any = True

        if ran_any:
            time.sleep(30)
            continue

        upcoming = []
        for (label, h, m) in SCHEDULES:
            sched_today = datetime.combine(date.today(), datetime.min.time()).replace(hour=h, minute=m)
            if last_run[label] == date.today():
                sched = next_dt_for(h, m, date.today())
            else:
                sched = next_dt_for(h, m, date.today())
                if datetime.now() < sched_today:
                    sched = sched_today
            upcoming.append((sched, label))

        next_time, next_label = min(upcoming, key=lambda x: x[0])
        wait_seconds = max(5, int((next_time - datetime.now()).total_seconds()))
        print(f"ينام حتى {next_time.strftime('%Y-%m-%d %H:%M')} للحدث [{next_label}] (~{wait_seconds//60} دقيقة).")
        while wait_seconds > 0:
            nap = min(wait_seconds, 300)
            time.sleep(nap)
            wait_seconds = int((next_time - datetime.now()).total_seconds())

        run_once(next_label)
        last_run[next_label] = date.today()

def main():
    loop_scheduler()

if __name__ == "__main__":
    main()
