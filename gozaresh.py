import os
import json
import base64
import requests
import pandas as pd
import jdatetime
from openai import OpenAI

# خواندن ایمن اطلاعات محرمانه از تنظیمات Secrets گیت‌هاب
BALE_TOKEN = os.environ.get("BALE_TOKEN")
GAPGPT_API_KEY = os.environ.get("GAPGPT_API_KEY")
MY_CHAT_ID = int(os.environ.get("MY_CHAT_ID", "0"))

BALE_API_URL = f"https://tapi.bale.ai/bot{BALE_TOKEN}"
EXCEL_FILE = "گزارش_طرح_سراج.xlsx"
ARCHIVE_DIR = "archive_voices"
META_FILE = "voices_index.json"

os.makedirs(ARCHIVE_DIR, exist_ok=True)

ai_client = OpenAI(
    api_key=GAPGPT_API_KEY,
    base_url="https://api.gapgpt.app/v1"
)

def load_index():
    if os.path.exists(META_FILE):
        with open(META_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"last_update_id": 0, "voices": []}

def save_index(data):
    with open(META_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def sync_new_voices_from_bale(index_data):
    offset = index_data["last_update_id"] + 1 if index_data["last_update_id"] > 0 else None
    params = {"timeout": 10}
    if offset:
        params["offset"] = offset

    res = requests.get(f"{BALE_API_URL}/getUpdates", params=params).json()
    if not res.get("ok"):
        print("خطا در اتصال به بله:", res)
        return index_data

    updates = res.get("result", [])
    existing_ids = {v["msg_id"] for v in index_data["voices"]}

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

    save_index(index_data)
    return index_data

def analyze_voice(audio_path, msg_jalali_date):
    with open(audio_path, "rb") as audio_file:
        encoded_audio = base64.b64encode(audio_file.read()).decode("utf-8")

    prompt = f"""
    این یک پیام صوتی گزارش عملکرد از کادر آموزشی/تربیتی طرح سراج است که در تاریخ {msg_jalali_date} ارسال شده است.
    لطفاً به فایل صوتی دقیق گوش بده و اطلاعات زیر را استخراج کن. اگر موردی در وویس ذکر نشده بود، عبارت "ذکر نشده" را بنویس (اگر تاریخ گفته نشد از {msg_jalali_date} استفاده کن).

    خروجی باید دقیقاً و صرفاً یک آبجکت JSON معتبر (بدون هیچ متن اضافی یا علامت مارک‌داون) با کلیدهای زیر باشد:
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

    response = ai_client.chat.completions.create(
        model="gemini-2.5-flash",
        temperature=0.1,
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {
                        "type": "input_audio",
                        "input_audio": {"data": encoded_audio, "format": "ogg"}
                    }
                ]
            }
        ]
    )

    raw_text = response.choices[0].message.content.strip()
    if raw_text.startswith("```"):
        raw_text = raw_text.strip("`").replace("json", "", 1).strip()
    return json.loads(raw_text)

def send_excel_to_bale(file_path, caption_text):
    with open(file_path, "rb") as doc:
        res = requests.post(
            f"{BALE_API_URL}/sendDocument",
            data={"chat_id": MY_CHAT_ID, "caption": caption_text},
            files={"document": doc}
        ).json()
        print("وضعیت ارسال اکسل به بله:", res.get("ok"))

def main():
    today_jalali = jdatetime.datetime.now().strftime("%Y/%m/%d")
    print(f"=== شروع گزارش‌گیری خودکار شبانه ({today_jalali}) ===")

    index_data = load_index()
    index_data = sync_new_voices_from_bale(index_data)

    extracted_rows = []
    for v in index_data["voices"]:
        if v.get("analyzed") or not os.path.exists(v["file_path"]):
            continue

        try:
            data = analyze_voice(v["file_path"], v["jalali_date"])
            extracted_rows.append(data)
            v["analyzed"] = True
        except Exception as e:
            print(f"خطا در تحلیل وویس {v['msg_id']}: {e}")

    save_index(index_data)

    if extracted_rows:
        df_new = pd.DataFrame(extracted_rows)
        df_new.to_excel(EXCEL_FILE, index=False)

        caption = f"📊 گزارش روزانه طرح سراج ({today_jalali})\nتعداد گزارش‌های پردازش‌شده: {len(extracted_rows)} مورد"
        send_excel_to_bale(EXCEL_FILE, caption)
        print("فایل اکسل با موفقیت ساخته و در بله ارسال شد.")
    else:
        print("امشب هیچ وویس جدیدی برای تحلیل ثبت نشده بود.")

if __name__ == "__main__":
    main()
