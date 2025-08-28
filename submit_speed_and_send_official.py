#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Submit speed test to the OFFICIAL Google Form (Muscat Directorate).
- Extracted from: "قياس سرعة الانترنت في مدارس مديرية مسقط"
- Uses robust mapping tries for text fields (since Google Forms hides names in saved HTML).

Requirements:
    pip install speedtest-cli requests

Run:
    python submit_speed_and_send_official.py
"""

import os
import csv
import time
from datetime import datetime
import requests

# -------- User-configurable metadata (PRE-FILLED) --------
SCHOOL_CODE = "1561"                          # 1- رمز المدرسة
SCHOOL_SECTOR = "السيب"                       # 2- قطاع المدرسة: مسقط / قريات / السيب / العامرات / بوشر / مطرح
SCHOOL_NAME = "مدرسة ابو القاسم الزهراوي"    # 3- اسم المدرسة
SERVICE_PROVIDER = "عمانتل"                   # 4- موفر الخدمة: عمانتل / أوريدو / أواصر
LINE_NUMBER = "24424428"                      # 5- رقم الخط
SERVICE_TYPE = "فايبر"                        # 6- نوع الخدمة: فايبر / الجيل الخامس 5 G
DEVICE_NAME = os.environ.get("COMPUTERNAME") or "Device"  # 7- اسم الجهاز (تلقائي)
LOG_DIR = os.path.join(os.getcwd(), "logs")
LOG_FILE = os.path.join(LOG_DIR, "speed_log.csv")

# -------- OFFICIAL Google Form wiring (extracted from uploaded HTML) --------
FORM_ACTION_URL = "https://docs.google.com/forms/u/2/d/e/1FAIpQLSfoYtl3gmt9FYa7g39v4az1OOtrkYHDcfAX6M-vhI6J-hX50A/formResponse"

# Select questions (IDs from *_sentinel fields)
ENTRY_SECTOR_ID = "entry.1313908626"       # 2- قطاع المدرسة
ENTRY_PROVIDER_ID = "entry.927675658"      # 4- موفر الخدمة
ENTRY_SERVICE_TYPE_ID = "entry.66731299"   # 6- نوع الخدمة

ALLOWED_SECTORS = ["مسقط", "قريات", "السيب", "العامرات", "بوشر", "مطرح"]
ALLOWED_PROVIDERS = ["عمانتل", "أوريدو", "أواصر"]
ALLOWED_SERVICE_TYPES = ["فايبر", "الجيل الخامس 5 G"]

# Text/textarea questions (inferred IDs from hidden entry.* in the form HTML)
# Confirmed: Q1 (school code) = entry.899161738
# The remaining four are for: Q3 school name, Q5 line number, Q7 speed text, Q8 notes
ENTRY_TEXT_IDS = {
    "Q1_school_code": "entry.899161738",
    "X_A": "entry.560537791",
    "X_B": "entry.1862560773",
    "X_C": "entry.181224386",
    "X_D": "entry.556952249",
}

# We'll try a few plausible mappings for X_A..X_D to (Q3,Q5,Q7,Q8)
TEXT_MAPPING_TRIES = [
    # Guess 1 (most likely): in document order
    {"Q3_school_name": "X_A", "Q5_line_number": "X_B", "Q7_internet_speed": "X_C", "Q8_notes": "X_D"},
    # Guess 2: swap C/D (speed vs notes)
    {"Q3_school_name": "X_A", "Q5_line_number": "X_B", "Q7_internet_speed": "X_D", "Q8_notes": "X_C"},
    # Guess 3: swap A/B (name vs line), keep C/D order
    {"Q3_school_name": "X_B", "Q5_line_number": "X_A", "Q7_internet_speed": "X_C", "Q8_notes": "X_D"},
]

HIDDEN_BASE = {
    "fvv": "1",
    "pageHistory": "0",
    # fbzx token varies; we'll try with/without
    # When present in saved HTML: "8122308104194036559" (example)
}

def measure_speed():
    try:
        import speedtest
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
        import csv
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([
                "timestamp", "download_mbps", "upload_mbps", "ping_ms",
                "server", "ip", "device",
                "school_code", "sector", "school_name",
                "provider", "line_number", "service_type"
            ])

def append_log(path, row):
    import csv
    with open(path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(row)

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

def try_submit_with_mapping(mapping, results, ts):
    base, notes_text = build_payload_base(results, ts)

    # Resolve mapping to actual entry.* keys
    q3_id = ENTRY_TEXT_IDS[mapping["Q3_school_name"]]
    q5_id = ENTRY_TEXT_IDS[mapping["Q5_line_number"]]
    q7_id = ENTRY_TEXT_IDS[mapping["Q7_internet_speed"]]
    q8_id = ENTRY_TEXT_IDS[mapping["Q8_notes"]]

    payload = dict(base)
    payload[q3_id] = SCHOOL_NAME
    payload[q5_id] = str(LINE_NUMBER)
    payload[q7_id] = f"{results['download']} Mbps"   # field 7
    payload[q8_id] = (
        f"تنزيل: {results['download']} Mbps | "
        f"رفع: {results['upload']} Mbps | "
        f"Ping: {results['ping']} ms | "
        f"السيرفر: {results['server']} | "
        f"IP: {results['ip']} | "
        f"الجهاز: {DEVICE_NAME} | "
        f"التاريخ/الوقت: {ts}"
    )

    # Try with and without fbzx (token varies)
    headers = {
        "Content-Type": "application/x-www-form-urlencoded;charset=utf-8",
        "User-Agent": "Mozilla/5.0",
        "Referer": "https://docs.google.com/forms/d/e/1FAIpQLSfoYtl3gmt9FYa7g39v4az1OOtrkYHDcfAX6M-vhI6J-hX50A/viewform",
    }

    variants = [payload, dict(payload, **{"fbzx": "8122308104194036559"})]

    for variant in variants:
        r = requests.post(FORM_ACTION_URL, data=variant, headers=headers, timeout=30)
        if r.status_code in (200, 302):
            return True, r.status_code, r.text[:500], mapping

    return False, r.status_code, r.text[:500], mapping

def main():
    print("بدء قياس السرعة... قد يستغرق الأمر دقيقة.")
    results = measure_speed()
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print("\nنتائج السرعة:")
    print(f"التاريخ/الوقت: {ts}")
    print(f"تنزيل (Mbps): {results['download']}")
    print(f"رفع (Mbps): {results['upload']}")
    print(f"Ping (ms): {results['ping']}")
    print(f"السيرفر: {results['server']}")
    print(f"IP العميل: {results['ip']}")
    print(f"الجهاز: {DEVICE_NAME}")

    ensure_log_header(LOG_FILE)
    append_log(LOG_FILE, [
        ts, results["download"], results["upload"], results["ping"],
        results["server"], results["ip"], DEVICE_NAME,
        SCHOOL_CODE, SCHOOL_SECTOR, SCHOOL_NAME,
        SERVICE_PROVIDER, LINE_NUMBER, SERVICE_TYPE
    ])
    print(f"\nتم حفظ النتيجة محليًا في: {LOG_FILE}")

    print("\nجارٍ إرسال النتائج إلى النموذج الرسمي...")
    last_preview = ""
    for mapping in TEXT_MAPPING_TRIES:
        ok, code, preview, used = try_submit_with_mapping(mapping, results, ts)
        print(f"- تجربة بالترتيب {used} => Status {code}")
        last_preview = preview
        if ok:
            print("تم الإرسال بنجاح ✅")
            return

    print("تعذّر الإرسال بعد كل المحاولات ❌")
    print("Preview:", last_preview[:500])

if __name__ == "__main__":
    main()
