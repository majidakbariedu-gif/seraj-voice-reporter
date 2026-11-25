import os
import json
import base64
import requests
import pandas as pd
import jdatetime
from openai import OpenAI
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# خواندن ایمن اطلاعات محرمانه از Secrets گیت‌هاب
BALE_TOKEN = os.environ.get("BALE_TOKEN")
GAPGPT_API_KEY = os.environ.get("GAPGPT_API_KEY")
MY_CHAT_ID = int(os.environ.get("MY_CHAT_ID") or 0)

BALE_API_URL = f"https://tapi.bale.ai/bot{BALE_TOKEN}"
EXCEL_FILE = "گزارش_طرح_سراج.xlsx"
ARCHIVE_DIR = "archive_voices"
META_FILE = "voices_index.json"
MIN_VOICE_DURATION = 4  # نادیده گرفتن وویس‌های تصادفی زیر ۴ ثانیه

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
            duration = voice.get("duration", 0)
            if duration < MIN_VOICE_DURATION:
                print(f"وویس شماره {msg_id} به دلیل کوتاه بودن ({duration} ثانیه) نادیده گرفته شد.")
                continue

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
    این یک پیام صوتی در گروه طرح آموزشی/تربیتی سراج است که در تاریخ {msg_jalali_date} ارسال شده است.
    قوانین بسیار مهم:
    1. اگر فایل صوتی فقط صدای نویز، سکوت، ضربه تصادفی، سرفه، یا مکالمه کاملاً نامربوط و فاقد هرگونه گزارش آموزشی/تربیتی/انضباطی است، از حدس زدن و ساختن جمله خودداری کن و صرفاً یک JSON با کلید `"معتبر": false` برگردان.
    2. در صورتی که صوت حاوی گزارش یا اطلاع‌رسانی مرتبط است، اطلاعات زیر را دقیق استخراج کن و به هیچ وجه از خودت چیزی اضافه نکن. اگر موردی در وویس گفته نشد، عبارت "ذکر نشده" را بنویس (اگر تاریخ گفته نشد از {msg_jalali_date} استفاده کن).
    3. مقدار فیلد `"اولویت_پیگیری"` را بر اساس حساسیت موضوع (مثلاً غیبت غیرموجه، مشکل انضباطی یا افت تحصیلی شدید = زیاد، موارد عادی پیگیری = متوسط، گزارش روتین بدون مشکل = کم) دقیقاً یکی از سه کلمه `"زیاد"`، `"متوسط"` یا `"کم"` قرار بده.

    خروجی باید دقیقاً و صرفاً یک آبجکت JSON معتبر (بدون هیچ متن اضافی یا علامت مارک‌داون) با کلیدهای زیر باشد:
    - "معتبر": true
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
        temperature=0.0,
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
    
    result = json.loads(raw_text)
    if not result.get("معتبر", True):
        return None
    result.pop("معتبر", None)
    return result

def create_styled_excel(df, file_path):
    # مرتب‌سازی و خوانا کردن نام ستون‌ها
    column_rename = {
        "نام_گزارش_دهنده": "نام گزارش‌دهنده",
        "روز": "روز هفته",
        "تاریخ": "تاریخ",
        "زنگ": "زنگ",
        "نام_درس": "نام درس",
        "حوزه_گزارش": "حوزه گزارش",
        "غایبین": "غایبین",
        "اسامی_دانش_آموزان_مطرح_شده": "دانش‌آموزان مطرح‌شده",
        "توضیحات": "شرح گزارش و توضیحات",
        "تکلیف": "تکلیف",
        "آزمون": "آزمون",
        "اولویت_پیگیری": "اولویت پیگیری",
        "توضیحات_بیشتر": "توضیحات تکمیلی"
    }
    df = df.rename(columns=column_rename)

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="گزارش روزانه طرح سراج")
        ws = writer.sheets["گزارش روزانه طرح سراج"]

        # ۱. راست‌به‌چپ کردن کامل شیت اکسل
        ws.views.sheetView[0].rightToLeft = True

        # تنظیمات فونت و استایل‌های رسمی
        font_family = "B Nazanin"
        header_font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
        cell_font = Font(name=font_family, size=12, bold=False, color="1F2937")
        
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")  # سرمه‌ای رسمی
        zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")   # خاکستری بسیار ملایم
        white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

        # رنگ‌های شرطی ستون اولویت
        high_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")    # قرمز ملایم
        high_font = Font(name=font_family, size=12, bold=True, color="991B1B")

        med_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")     # سبز ملایم
        med_font = Font(name=font_family, size=12, bold=True, color="166534")

        low_fill = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")     # زرد ملایم
        low_font = Font(name=font_family, size=12, bold=True, color="854D0E")

        thin_side = Side(border_style="thin", color="CBD5E1")
        cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        # استایل‌دهی ردیف سرستون‌ها
        ws.row_dimensions[1].height = 34
        priority_col_idx = None

        for col_idx, cell in enumerate(ws[1], start=1):
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = cell_border
            if cell.value == "اولویت پیگیری":
                priority_col_idx = col_idx

        # استایل‌دهی ردیف‌های داده و رنگ‌بندی ستون اولویت
        for row_idx in range(2, ws.max_row + 1):
            row_fill = zebra_fill if row_idx % 2 == 0 else white_fill
            max_lines = 1

            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                val_str = str(cell.value or "")
                
                cell.font = cell_font
                cell.fill = row_fill
                cell.border = cell_border
                
                # تراز وسط برای ستون‌های کوتاه و راست‌چین برای متن‌های توضیحی
                if col_idx in [1, 2, 3, 4, 5, 6, 12]:
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                else:
                    cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)

                # رنگ‌بندی اختصاصی ستون اولویت پیگیری
                if col_idx == priority_col_idx:
                    if "زیاد" in val_str or "بالا" in val_str:
                        cell.fill = high_fill
                        cell.font = high_font
                    elif "متوسط" in val_str:
                        cell.fill = med_fill
                        cell.font = med_font
                    elif "کم" in val_str or "پایین" in val_str:
                        cell.fill = low_fill
                        cell.font = low_font

                # تخمین تعداد خطوط برای تنظیم ارتفاع ردیف
                estimated_lines = max(1, (len(val_str) // 35) + 1)
                if estimated_lines > max_lines:
                    max_lines = estimated_lines

            ws.row_dimensions[row_idx].height = max(28, max_lines * 22)

        # تنظیم هوشمند عرض هر ستون بر اساس محتوا
        custom_widths = {
            "نام گزارش‌دهنده": 18,
            "روز هفته": 13,
            "تاریخ": 14,
            "زنگ": 12,
            "نام درس": 16,
            "حوزه گزارش": 18,
            "غایبین": 22,
            "دانش‌آموزان مطرح‌شده": 24,
            "شرح گزارش و توضیحات": 42,
            "تکلیف": 24,
            "آزمون": 22,
            "اولویت پیگیری": 16,
            "توضیحات تکمیلی": 30
        }

        for col_idx, cell in enumerate(ws[1], start=1):
            col_letter = get_column_letter(col_idx)
            col_title = str(cell.value or "")
            ws.column_dimensions[col_letter].width = custom_widths.get(col_title, 18)

        # ثابت کردن ردیف سرستون هنگام اسکرول
        ws.freeze_panes = "A2"

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
            v["analyzed"] = True
            if data is not None:
                extracted_rows.append(data)
            else:
                print(f"وویس {v['msg_id']} نامرتبط تشخیص داده شد و در اکسل نیامد.")
        except Exception as e:
            print(f"خطا در تحلیل وویس {v['msg_id']}: {e}")

    save_index(index_data)

    if extracted_rows:
        df_new = pd.DataFrame(extracted_rows)
        create_styled_excel(df_new, EXCEL_FILE)

        caption = f"📊 گزارش روزانه طرح سراج ({today_jalali})\nتعداد گزارش‌های معتبر پردازش‌شده: {len(extracted_rows)} مورد"
        send_excel_to_bale(EXCEL_FILE, caption)
        print("فایل اکسل با طراحی جدید ساخته و در بله ارسال شد.")
    else:
        print("امشب هیچ وویس معتبر جدیدی برای تحلیل ثبت نشده بود.")

if __name__ == "__main__":
    main()
