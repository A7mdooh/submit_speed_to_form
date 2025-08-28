#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
سكربت متقدم لقياس سرعة الإنترنت وإرسال النتائج إلى نموذج Google Forms.

الافتراضات المضبوطة حسب طلبك:
  --school-code     1561
  --sector          السيب
  --school-name     "أبو القاسم الزهراوي للتعليم الأساسي 5-9"
  --provider        عمانتل
  --line-number     24424428
  --service-type    فايبر
  --note            "تم القياس بسلك إيثرنت عبر speedtest-cli"

الميزات:
- وسائط سطر أوامر لتغيير القيم عند الحاجة.
- ربط ملاحظاتك مع entry.899161738 مباشرة.
- إعادة محاولات تلقائية مع Backoff عند فشل الإرسال.
- تسجيل محلي للنتائج في CSV (logs/speed_log.csv).
- التحقق من صحة القيم (مزود/قطاع/خدمة).
- ترويسة HTTP مناسبة.

المتطلبات:
    pip install pyinstaller speedtest-cli requests

أمثلة تشغيل:
    python submit_speed_to_form_advanced.py
    # أو لتغيير أي قيمة:
    python submit_speed_to_form_advanced.py --sector مسقط --note "اختبار مسائي"
"""

import os
import csv
import time
import platform
from pathlib import Path
from datetime import datetime
import argparse
import requests

try:
    import speedtest  # from speedtest-cli
except ImportError as e:
    raise SystemExit("الرجاء تثبيت speedtest-cli أولاً: pip install speedtest-cli") from e


# ---------- إعدادات ثابتة للنموذج ----------
FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSfoYtl3gmt9FYa7g39v4az1OOtrkYHDcfAX6M-vhI6J-hX50A/formResponse"

E_SCHOOL_CODE   = "entry.560537791"   # 1- رمز المدرسة
E_SECTOR        = "entry.1313908626"  # 2- قطاع المدرسة
E_SCHOOL_NAME   = "entry.1862560773"  # 3- اسم المدرسة
E_PROVIDER      = "entry.927675658"   # 4- موفر الخدمة
E_LINE_NUMBER   = "entry.181224386"   # 5- رقم الخط
E_SERVICE_TYPE  = "entry.66731299"    # 6- نوع الخدمة
E_SPEED_FIELD   = "entry.556952249"   # 7- سرعة الإنترنت (نص حر)
E_NOTES         = "entry.899161738"   # 8- ملاحظات

VALID_SECTORS   = {"مسقط","قريات","السيب","العامرات","بوشر","مطرح"}
VALID_PROVIDERS = {"عمانتل","أوريدو","أواصر"}
VALID_SERVICES  = {"فايبر","الجيل الخامس 5 G"}

LOG_DIR  = Path("logs")
LOG_FILE = LOG_DIR / "speed_log.csv"


def bps_to_mbps(bps: float) -> float:
    return round(bps / 1_000_000, 2)


def measure_speed(timeout_sec: int = 30) -> dict:
    st = speedtest.Speedtest(timeout=timeout_sec)
    st.get_servers([])
    st.get_best_server()
    download_bps = st.download()
    upload_bps   = st.upload()
    ping_ms      = st.results.ping
    return {
        "download_mbps": bps_to_mbps(download_bps),
        "upload_mbps":   bps_to_mbps(upload_bps),
        "ping_ms":       round(ping_ms, 1),
        "server":        st.results.server.get("host", ""),
        "sponsor":       st.results.server.get("sponsor", ""),
        "client_ip":     st.results.client.get("ip", ""),
    }


def build_speed_text(results: dict) -> str:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [
        f"التاريخ/الوقت: {ts}",
        f"تنزيل (Mbps): {results['download_mbps']}",
        f"رفع (Mbps): {results['upload_mbps']}",
        f"Ping (ms): {results['ping_ms']}",
        f"السيرفر: {results['server']} ({results['sponsor']})",
        f"IP العميل: {results['client_ip']}",
        f"الجهاز: {platform.node()}",
    ]
    return "\n".join(lines)


def submit_to_form(payload: dict, retries: int = 3, backoff_sec: float = 2.0) -> bool:
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
        "Referer": FORM_URL.replace("formResponse", "viewform"),
        "Origin": "https://docs.google.com",
    }
    for attempt in range(1, retries + 1):
        try:
            resp = requests.post(FORM_URL, data=payload, headers=headers, timeout=30)
            if resp.status_code == 200:
                return True
            else:
                print(f"[تحذير] حالة HTTP غير متوقعة: {resp.status_code} (محاولة {attempt}/{retries})")
        except requests.RequestException as e:
            print(f"[خطأ] فشل الطلب: {e} (محاولة {attempt}/{retries})")
        time.sleep(backoff_sec * attempt)  # backoff تزايدي
    return False


def log_to_csv(row: dict):
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    new_file = not LOG_FILE.exists()
    # ثبّت ترتيب الأعمدة
    fieldnames = [
        "timestamp","download_mbps","upload_mbps","ping_ms",
        "server","sponsor","client_ip","sector","provider",
        "service_type","line_number"
    ]
    with LOG_FILE.open("a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        if new_file:
            writer.writeheader()
        writer.writerow(row)


def parse_args():
    parser = argparse.ArgumentParser(description="قياس سرعة الإنترنت وإرسالها إلى Google Forms")

    # افتراضات حسب طلبك (قابلة للتغيير بسطر الأوامر)
    parser.add_argument("--school-code", default="1561")
    parser.add_argument("--sector", default="السيب", choices=sorted(VALID_SECTORS))
    parser.add_argument("--school-name", default="أبو القاسم الزهراوي للتعليم الأساسي 5-9")
    parser.add_argument("--provider", default="عمانتل", choices=sorted(VALID_PROVIDERS))
    parser.add_argument("--line-number", default="24424428")
    parser.add_argument("--service-type", default="فايبر", choices=sorted(VALID_SERVICES))
    parser.add_argument("--note", default="تم القياس بسلك إيثرنت عبر speedtest-cli")
    parser.add_argument("--timeout", type=int, default=30, help="مهلة speedtest بالثواني (افتراضي 30)")
    parser.add_argument("--no-submit", action="store_true", help="تنفيذ القياس فقط بدون إرسال للنموذج")
    parser.add_argument("--retries", type=int, default=3, help="عدد محاولات الإرسال")
    parser.add_argument("--backoff", type=float, default=2.0, help="زمن الـ backoff الابتدائي")

    return parser.parse_args()


def main():
    args = parse_args()

    print("بدء قياس السرعة... قد يستغرق الأمر دقيقة.")
    results = measure_speed(timeout_sec=args.timeout)
    speed_text = build_speed_text(results)

    print("\nنتائج السرعة:")
    print(speed_text)

    # سجل محليًا دائمًا
    log_row = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "download_mbps": results["download_mbps"],
        "upload_mbps": results["upload_mbps"],
        "ping_ms": results["ping_ms"],
        "server": results["server"],
        "sponsor": results["sponsor"],
        "client_ip": results["client_ip"],
        "sector": args.sector,
        "provider": args.provider,
        "service_type": args.service_type,
        "line_number": args.line_number,
    }
    log_to_csv(log_row)
    print(f"\nتم حفظ نتيجة محليًا في: {LOG_FILE.resolve()}")

    if args.no_submit:
        print("\nتم تنفيذ القياس بدون إرسال (--no-submit).")
        return

    # إعداد حمولة الإرسال للنموذج
    payload = {
        E_SCHOOL_CODE:  args.school_code,
        E_SECTOR:       args.sector,
        E_SCHOOL_NAME:  args.school_name,
        E_PROVIDER:     args.provider,
        E_LINE_NUMBER:  args.line_number,
        E_SERVICE_TYPE: args.service_type,
        E_SPEED_FIELD:  speed_text,
        E_NOTES:        args.note,   # <-- ربط الملاحظات مع entry.899161738
        # حقول لازمة لميكانيكية Google Forms
        "fvv": "1",
        "fbzx": str(int(time.time() * 1000)),
        "pageHistory": "0",
    }

    print("\nجارٍ إرسال النتائج إلى Google Form...")
    ok = submit_to_form(payload, retries=args.retries, backoff_sec=args.backoff)

    if ok:
        print("تم الإرسال بنجاح ✅")
    else:
        print("تعذّر الإرسال بعد جميع المحاولات ❌")


if __name__ == "__main__":
    main()
