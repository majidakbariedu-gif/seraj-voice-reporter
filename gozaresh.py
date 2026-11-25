import os
import json
import builtins
import requests
import pandas as pd
import jdatetime
import arabic_reshaper
from bidi.algorithm import get_display
from google import genai
from google.genai import types

# ================= اصلاح نمایش فارسی در CMD =================
def fix_fa(text):
    if isinstance(text, str):
        reshaped = arabic_reshaper.reshape(text)
        return get_display(reshaped)
    return text

_orig_print = builtins.print
def print(*args, **kwargs):
    fixed_args = [fix_fa(str(a)) for a in args]
    _orig_print(*fixed_args, **kwargs)

_orig_input = builtins.input
def input(prompt=""):
    return _orig_input(fix_fa(str(prompt)))
# ============================================================

# ================= 1. این دو خط را پر کنید =================
BALE_TOKEN = "206132192:amc4sp-7S6eXaNh3X0Frz519j6hWMX3InZA"
GEMINI_API_KEY = "AQ.Ab8RN6IBwQneF4yZ1SrQ_zDlsv1Ud6tNXn5L8WEmiGfmqUYzqg"
# ==========================================================

BALE_API_URL = f"https://tapi.bale.ai/bot{BALE_TOKEN}"
EXCEL_FILE = "گزارش_طرح_سراج.xlsx"
ARCHIVE_DIR = "archive_voices"
META_FILE = "voices_index.json"

os.makedirs(ARCHIVE_DIR, exist_ok=True)
ai_client = genai.Client(api_key=GEMINI_API_KEY)

def load_index():
    if os.path.exists(META_FILE):
        with open(META_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"last_update_id": 0, "voices": []}

def save_index(data):
    with open(META_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def jalali_to_timestamp(jalali_str, is_end=False):
    if not jalali_str.strip():
        return None
    parts = [int(x) for x in jalali_str.strip().split("/")]
    jdt = jdatetime.datetime(parts[0], parts[1], parts[2], 23 if is_end else 0, 59 if is_end else 0, 59 if is_end else 0)
    return int(jdt.togregorian().timestamp())

def sync_new_voices_from_bale(index_data):
    print("در حال دریافت وویس‌های جدید از بله و ذخیره در آرشیو کامپیوتر...")
    offset = index_data["last_update_id"] + 1 if index_data["last_update_id"] > 0 else None
    params = {"timeout": 10}
    if offset:
        params["offset"] = offset

    res = requests.get(f"{BALE_API_URL}/getUpdates", params=params).json()
    if not res.get("ok"):
        print("خطا در اتصال به بله! توکن را بررسی کنید:", res)
        return index_data

    updates = res.get("result", [])
    existing_ids = {v["msg_id"] for v in index_data["voices"]}
    new_count = 0

    for item in updates:
        if item["update_id"] > index_data["last_update_id"]:
            index_data["last_update_id"] = item["update_id"]

        msg = item.get("message") or item.get("channel_post")
        if not msg:
            continue

        msg_id = msg.get("message_id")
        msg_ts = msg.get("date", 0)
        voice = msg.get("voice") or msg.get("audio")

        if voice and msg_id not in existing_ids:
            msg_jalali = jdatetime.datetime.fromtimestamp(msg_ts).strftime("%Y/%m/%d")
            safe_date = msg_jalali.replace("/", "-")
            file_path = os.path.join(ARCHIVE_DIR, f"{safe_date}_msg_{msg_id}.ogg")

            file_info = requests.get(f"{BALE_API_URL}/getFile", params={"file_id": voice["file_id"]}).json()
            if file_info.get("ok"):
                remote_path = file_info["result"]["file_path"]
                audio_bytes = requests.get(f"https://tapi.bale.ai/file/bot{BALE_TOKEN}/{remote_path}").content
                with open(file_path, "wb") as f:
                    f.write(audio_bytes)

                index_data["voices"].append({
                    "msg_id": msg_id,
                    "timestamp": msg_ts,
                    "jalali_date": msg_jalali,
                    "file_path": file_path,
                    "analyzed": False
                })
                existing_ids.add(msg_id)
                new_count += 1

    save_index(index_data)
    if new_count > 0:
        print(f" -> {new_count} وویس جدید از بله دانلود و در آرشیو ذخیره شد.")
    return index_data

def analyze_voice(audio_path, msg_jalali_date):
    uploaded_file = ai_client.files.upload(file=audio_path)
    prompt = f"""
    این یک پیام صوتی گزارش عملکرد از کادر آموزشی/تربیتی طرح سراج است که در تاریخ {msg_jalali_date} ارسال شده است.
    لطفاً به فایل صوتی دقیق گوش بده و اطلاعات زیر را استخراج کن. اگر موردی در وویس ذکر نشده بود، عبارت "ذکر نشده" را بنویس (اگر تاریخ گفته نشد از {msg_jalali_date} استفاده کن).

    خروجی باید دقیقاً یک آبجکت JSON با کلیدهای زیر باشد:
    - "نام_گزارش_دهنده"
    - "روز"
    - "تاریخ"
    - "زنگ"
    - "نام_درس"
    - "حوزه_گزارش"
    - "غایبین"
    - "اسامی_دانش_آموزان_مطرح_شده"
    - "توضیحات"
    - "تکلیف"
    - "آزمون"
    - "اولویت_پیگیری"
    - "توضیحات_بیشتر"
    """
    response = ai_client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[uploaded_file, prompt],
        config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.1)
    )
    ai_client.files.delete(name=uploaded_file.name)
    return json.loads(response.text)

def main():
    print("=== سیستم گزارش‌گیری هوشمند طرح سراج ===")
    
    index_data = load_index()
    index_data = sync_new_voices_from_bale(index_data)

    print("\nحالا بازه تاریخی گزارش‌گیری را مشخص کنید (برای دریافت همه موارد تحلیل‌نشده، فقط Enter بزنید):")
    start_str = input("از چه تاریخی؟ (مثلاً 1405/08/25): ")
    end_str = input("تا چه تاریخی؟ (مثلاً 1405/09/03): ")

    has_date_filter = bool(start_str.strip() or end_str.strip())
    start_ts = jalali_to_timestamp(start_str, is_end=False)
    end_ts = jalali_to_timestamp(end_str, is_end=True)

    extracted_rows = []

    for v in index_data["voices"]:
        if not os.path.exists(v["file_path"]):
            continue

        if not has_date_filter and v.get("analyzed"):
            continue

        if start_ts and v["timestamp"] < start_ts:
            continue
        if end_ts and v["timestamp"] > end_ts:
            continue

        print(f"در حال تحلیل وویس تاریخ {v['jalali_date']} (شناسه {v['msg_id']})...")
        try:
            data = analyze_voice(v["file_path"], v["jalali_date"])
            extracted_rows.append(data)
            v["analyzed"] = True
            print(f" -> انجام شد: {data.get('نام_گزارش_دهنده')} | اولویت: {data.get('اولویت_پیگیری')}")
        except Exception as e:
            print(f" -> خطا در تحلیل هوش مصنوعی: {e}")

    save_index(index_data)

    if extracted_rows:
        df_new = pd.DataFrame(extracted_rows)
        if os.path.exists(EXCEL_FILE):
            df_old = pd.read_excel(EXCEL_FILE)
            df_final = pd.concat([df_old, df_new], ignore_index=True)
        else:
            df_final = df_new
        df_final.to_excel(EXCEL_FILE, index=False)
        print(f"\n✅ عالی! {len(extracted_rows)} گزارش با موفقیت استخراج و در فایل اکسل ثبت شد.")
    else:
        print("\nهیچ وویسی مطابق با بازه زمانی انتخابی شما یافت نشد.")

    input("\nبرای خروج کلید Enter را بزنید...")

if __name__ == "__main__":
    main()