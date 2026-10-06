import asyncio
import os
import json
import time
import tempfile
from datetime import datetime, timedelta
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.types import (
    FSInputFile,
    ReplyKeyboardMarkup,
    KeyboardButton,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
)
from dotenv import load_dotenv
import edge_tts
from docx import Document
import PyPDF2
import csv
import pandas as pd
from pptx import Presentation
import re
import aiohttp
import logging
from PIL import Image
import pytesseract

# =================== ENV ===================
load_dotenv(".env")
BOT_TOKEN = os.getenv("BOT_TOKEN")
ADMIN_IDS = [int(x) for x in os.getenv("ADMIN_IDS", "").split(",") if x.strip()]

# Tesseract yo'lini sozlash (Windows uchun, agar kerak bo'lsa)
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Windows uchun
# Linux/Mac uchun bu qatorni o'chirib tashlang yoki moslashtiring

# =================== Fayllar ===================
USERS_FILE = "users.json"

def load_users():
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"[ERROR] Users faylni o'qishda xato: {e}")
            return {}
    return {}

def save_users():
    try:
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            json.dump(users, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"[ERROR] Users faylni saqlashda xato: {e}")

users = load_users()
user_state = {}

# =================== Tariflar ===================
tariflar = {
    "2️⃣ 15 limit - 5 000 so'm": {"limit": 15, "price": 5000},
    "3️⃣ Cheksiz (1 hafta) - 7 000 so'm": {"limit": -1, "price": 7000},
    "4️⃣ Cheksiz (1 oy) - 10 000 so'm": {"limit": -1, "price": 10000},
}

# =================== Klaviaturalar ===================
def main_keyboard():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📊 Mening holatim")],
            [KeyboardButton(text="💳 limitni oshirish"), KeyboardButton(text="🎤 Ovoz tanlash")]
        ],
        resize_keyboard=True
    )

def tarif_buttons():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text=name) for name in tariflar.keys()],
            [KeyboardButton(text="🔙 Orqaga")]
        ],
        resize_keyboard=True,
        one_time_keyboard=True
    )

def payment_confirmation_keyboard():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="✅ To'lovni tasdiqlash")],
            [KeyboardButton(text="🔙 Orqaga")]
        ],
        resize_keyboard=True,
        one_time_keyboard=True
    )

def voice_selection_keyboard():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="👨 Erkak ovozi"), KeyboardButton(text="👩 Ayol ovozi")],
            [KeyboardButton(text="🔙 Orqaga")]
        ],
        resize_keyboard=True,
        one_time_keyboard=True
    )

# =================== Ovozlar (O'zbekcha) ===================
VOICES = {
    "male": "uz-UZ-SardorNeural",
    "female": "uz-UZ-MadinaNeural"
}

# =================== Rasmdagi matnni o'qish ===================
async def extract_text_from_image(file_path):
    try:
        img = Image.open(file_path)
        text = pytesseract.image_to_string(img, lang='uzb+uzb_cyrl')
        text = text.strip()
        if not text:
            return None
        return preprocess_text(text)
    except Exception as e:
        print(f"[ERROR] Rasmdagi matnni o'qishda xato: {e}")
        return None

# =================== Kiril-Lotin transliteratsiyasi ===================
def transliterate_kirill_to_latin(text):
    """Kiril yozuvidagi o'zbekcha matnni lotinga aylantiradi."""
    translit_dict = {
        'а': 'a', 'А': 'A', 'б': 'b', 'Б': 'B', 'в': 'v', 'В': 'V', 
        'г': 'g', 'Г': 'G', 'д': 'd', 'Д': 'D', 'е': 'e', 'Е': 'E', 
        'ё': 'yo', 'Ё': 'Yo', 'ж': 'j', 'Ж': 'J', 'з': 'z', 'З': 'Z', 
        'и': 'i', 'И': 'I', 'й': 'y', 'Й': 'Y', 'к': 'k', 'К': 'K', 
        'л': 'l', 'Л': 'L', 'м': 'm', 'М': 'M', 'н': 'n', 'Н': 'N', 
        'о': 'o', 'О': 'O', 'п': 'p', 'П': 'P', 'р': 'r', 'Р': 'R', 
        'с': 's', 'С': 'S', 'т': 't', 'Т': 'T', 'у': 'u', 'У': 'U', 
        'ф': 'f', 'Ф': 'F', 'х': 'x', 'Х': 'X', 'ц': 'ts', 'Ц': 'Ts', 
        'ч': 'ch', 'Ч': 'Ch', 'ш': 'sh', 'Ш': 'Sh', 'ъ': "'", 'Ъ': "'", 
        'ы': 'i', 'Ы': 'I', 'ь': '', 'Ь': '', 'э': 'e', 'Э': 'E', 
        'ю': 'yu', 'Ю': 'Yu', 'я': 'ya', 'Я': 'Ya',
        'ғ': 'gʻ', 'Ғ': 'Gʻ', 'қ': 'q', 'Қ': 'Q', 'ҳ': 'h', 'Ҳ': 'H', 
        'ў': 'oʻ', 'Ў': 'Oʻ'
    }
    
    kirill_chars = set('абвгдеёжзийклмнопрстуфхцчшъыьэюяғқҳў')
    if not any(c in kirill_chars for c in text):
        return text
    
    result = ''
    i = 0
    while i < len(text):
        if i < len(text) - 1 and text[i:i+2] in translit_dict:
            result += translit_dict[text[i:i+2]]
            i += 2
        elif text[i] in translit_dict:
            result += translit_dict[text[i]]
            i += 1
        else:
            result += text[i]
            i += 1
    return result

# =================== Rim raqamlarni oddiy raqamga aylantirish ===================
def roman_to_int(s):
    """Rim raqamlarini oddiy raqamga aylantiradi."""
    roman_values = {'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100, 'D': 500, 'M': 1000}
    result = 0
    prev_value = 0
    s = s.upper()
    
    valid_roman = r'^(M{0,3})(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$'
    if not re.match(valid_roman, s):
        return None
    
    for char in reversed(s):
        curr_value = roman_values.get(char, 0)
        if not curr_value:
            return None
        if curr_value >= prev_value:
            result += curr_value
        else:
            result -= curr_value
        prev_value = curr_value
    return result if result > 0 else None

# =================== Matnni oldindan qayta ishlash ===================
def preprocess_text(text):
    """Matnni transliteratsiya qiladi, rim raqamlarni oddiy raqamga va asrlarni so'zga aylantiradi, '-' ni 'inchi' ga aylantiradi."""
    text = transliterate_kirill_to_latin(text)
    
    number_to_ordinal = {
        1: 'birinchi', 2: 'ikkinchi', 3: 'uchinchi', 4: 'to‘rtinchi', 5: 'beshinchi',
        6: 'oltinchi', 7: 'yettinchi', 8: 'sakkizinchi', 9: 'to‘qqizinchi', 10: 'o‘ninchi',
        11: 'o‘n birinchi', 12: 'o‘n ikkinchi', 13: 'o‘n uchinchi', 14: 'o‘n to‘rtinchi',
        15: 'o‘n beshinchi', 16: 'o‘n oltinchi', 17: 'o‘n yettinchi', 18: 'o‘n sakkizinchi',
        19: 'o‘n to‘qqizinchi', 20: 'yigirmanchi', 21: 'yigirma birinchi', 22: 'yigirma ikkinchi',
        23: 'yigirma uchinchi', 24: 'yigirma to‘rtinchi', 25: 'yigirma beshinchi',
        26: 'yigirma oltinchi', 27: 'yigirma yettinchi', 28: 'yigirma sakkizinchi',
        29: 'yigirma to‘qqizinchi', 30: 'o‘ttizinchi', 31: 'o‘ttiz birinchi', 32: 'o‘ttiz ikkinchi',
        33: 'o‘ttiz uchinchi', 34: 'o‘ttiz to‘rtinchi', 35: 'o‘ttiz beshinchi',
        36: 'o‘ttiz oltinchi', 37: 'o‘ttiz yettinchi', 38: 'o‘ttiz sakkizinchi',
        39: 'o‘ttiz to‘qqizinchi', 40: 'qirqinchi', 41: 'qirq birinchi', 42: 'qirq ikkinchi',
        43: 'qirq uchinchi', 44: 'qirq to‘rtinchi', 45: 'qirq beshinchi', 46: 'qirq oltinchi',
        47: 'qirq yettinchi', 48: 'qirq sakkizinchi', 49: 'qirq to‘qqizinchi', 50: 'ellikinchi',
        51: 'ellik birinchi', 52: 'ellik ikkinchi', 53: 'ellik uchinchi', 54: 'ellik to‘rtinchi',
        55: 'ellik beshinchi', 56: 'ellik oltinchi', 57: 'ellik yettinchi', 58: 'ellik sakkizinchi',
        59: 'ellik to‘qqizinchi', 60: 'oltmishinchi', 61: 'oltmish birinchi', 62: 'oltmish ikkinchi',
        63: 'oltmish uchinchi', 64: 'oltmish to‘rtinchi', 65: 'oltmish beshinchi',
        66: 'oltmish oltinchi', 67: 'oltmish yettinchi', 68: 'oltmish sakkizinchi',
        69: 'oltmish to‘qqizinchi', 70: 'yetmishinchi', 71: 'yetmish birinchi', 72: 'yetmish ikkinchi',
        73: 'yetmish uchinchi', 74: 'yetmish to‘rtinchi', 75: 'yetmish beshinchi',
        76: 'yetmish oltinchi', 77: 'yetmish yettinchi', 78: 'yetmish sakkizinchi',
        79: 'yetmish to‘qqizinchi', 80: 'saksoninchi', 81: 'sakson birinchi', 82: 'sakson ikkinchi',
        83: 'sakson uchinchi', 84: 'sakson to‘rtinchi', 85: 'sakson beshinchi',
        86: 'sakson oltinchi', 87: 'sakson yettinchi', 88: 'sakson sakkizinchi',
        89: 'sakson to‘qqizinchi', 90: 'to‘qsoninchi', 91: 'to‘qson birinchi', 92: 'to‘qson ikkinchi',
        93: 'to‘qson uchinchi', 94: 'to‘qson to‘rtinchi', 95: 'to‘qson beshinchi',
        96: 'to‘qson oltinchi', 97: 'to‘qson yettinchi', 98: 'to‘qson sakkizinchi',
        99: 'to‘qson to‘qqizinchi', 100: 'yuzinchi'
    }

    def replace_roman_asr(match):
        roman = match.group(1)
        num = roman_to_int(roman)
        if num is not None and num in number_to_ordinal:
            return f"{number_to_ordinal[num]} asr"
        return match.group(0)

    def replace_roman(match):
        roman = match.group(0)
        num = roman_to_int(roman)
        if num is not None:
            return str(num)
        return roman

    def replace_hyphen(match):
        num_str = match.group(1)
        try:
            num = int(num_str)
            if num in number_to_ordinal:
                return number_to_ordinal[num]
            return num_str
        except ValueError:
            return num_str

    text = re.sub(r'\b([IVXLCDM]+)\s*(?:asr|ASR)\b', replace_roman_asr, text, flags=re.IGNORECASE)
    text = re.sub(r'\b[IVXLCDM]+\b', replace_roman, text, flags=re.IGNORECASE)
    text = re.sub(r'\b(\d+)\s*-\b', replace_hyphen, text)

    # Takomillashtirish: tinish belgilari atrofida bo'shliq qo'shish va chunarsiz o'qishni oldini olish uchun
    text = re.sub(r'([.,!?;])', r'\1 ', text)  # Tinish belgilardan keyin bo'shliq qo'shish
    text = re.sub(r'\s+', ' ', text)  # Ortiqcha bo'shliqlarni olib tashlash
    text = re.sub(r'(\d+)', r' \1 ', text)  # Raqamlarni ajratish
    text = re.sub(r'\s+', ' ', text)  # Yana tozalash

    return text

# =================== Fayl o'qish funksiyalari ===================
def read_txt(path):
    """Oddiy matn faylini o'qiydi."""
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
            if not content.strip():
                print(f"[WARNING] TXT fayl bo'sh: {path}")
                return ""
            return content
    except FileNotFoundError:
        print(f"[ERROR] TXT fayl topilmadi: {path}")
        return ""
    except Exception as e:
        print(f"[ERROR] TXT faylni o'qishda xato: {path} - {e}")
        return ""

def read_docx(path):
    """DOCX faylini o'qiydi."""
    try:
        doc = Document(path)
        text = []
        for p in doc.paragraphs:
            if p.text.strip():
                text.append(p.text)
        for table in doc.tables:
            for row in table.rows:
                row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                if row_text:
                    text.append(" | ".join(row_text))
        result = "\n".join(text)
        if not result.strip():
            print(f"[WARNING] DOCX fayl bo'sh: {path}")
        return result
    except FileNotFoundError:
        print(f"[ERROR] DOCX fayl topilmadi: {path}")
        return ""
    except Exception as e:
        print(f"[ERROR] DOCX faylni o'qishda xato: {path} - {e}")
        return ""

def read_pdf(path):
    """PDF faylini o'qiydi."""
    try:
        text = ""
        with open(path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            if not reader.pages:
                print(f"[WARNING] PDF faylda sahifalar yo'q: {path}")
                return ""
            for page_num, page in enumerate(reader.pages):
                page_text = page.extract_text() or ""
                if page_text.strip():
                    text += page_text + "\n"
        if not text.strip():
            print(f"[WARNING] PDF fayldan matn o'qilmadi: {path}")
        return text
    except FileNotFoundError:
        print(f"[ERROR] PDF fayl topilmadi: {path}")
        return ""
    except Exception as e:
        print(f"[ERROR] PDF faylni o'qishda xato: {path} - {e}")
        return ""

def read_csv_file(path):
    """CSV faylini o'qiydi."""
    try:
        text = []
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            reader = csv.reader(f)
            for row in reader:
                row_text = [cell.strip() for cell in row if cell.strip()]
                if row_text:
                    text.append(" | ".join(row_text))
        result = "\n".join(text)
        if not result.strip():
            print(f"[WARNING] CSV fayl bo'sh: {path}")
        return result
    except FileNotFoundError:
        print(f"[ERROR] CSV fayl topilmadi: {path}")
        return ""
    except Exception as e:
        print(f"[ERROR] CSV faylni o'qishda xato: {path} - {e}")
        return ""

def read_json_file(path):
    """JSON faylini o'qiydi."""
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            data = json.load(f)
            result = json.dumps(data, ensure_ascii=False, indent=2)
            if not result.strip():
                print(f"[WARNING] JSON fayl bo'sh: {path}")
            return result
    except FileNotFoundError:
        print(f"[ERROR] JSON fayl topilmadi: {path}")
        return ""
    except json.JSONDecodeError as e:
        print(f"[ERROR] JSON faylni parse qilishda xato: {path} - {e}")
        return ""
    except Exception as e:
        print(f"[ERROR] JSON faylni o'qishda xato: {path} - {e}")
        return ""

def read_excel(path):
    """Excel faylini o'qiydi."""
    try:
        df = pd.read_excel(path)
        if df.empty:
            print(f"[WARNING] Excel fayl bo'sh: {path}")
            return ""
        return df.to_string(index=False)
    except FileNotFoundError:
        print(f"[ERROR] Excel fayl topilmadi: {path}")
        return ""
    except Exception as e:
        print(f"[ERROR] Excel faylni o'qishda xato: {path} - {e}")
        return ""

def read_pptx(path):
    """PowerPoint faylini o'qiydi."""
    try:
        prs = Presentation(path)
        texts = []
        for slide_num, slide in enumerate(prs.slides):
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    texts.append(shape.text)
                if shape.has_table:
                    for row in shape.table.rows:
                        row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                        if row_text:
                            texts.append(" | ".join(row_text))
        result = "\n".join(texts)
        if not result.strip():
            print(f"[WARNING] PPTX fayldan matn o'qilmadi: {path}")
        return result
    except FileNotFoundError:
        print(f"[ERROR] PPTX fayl topilmadi: {path}")
        return ""
    except Exception as e:
        print(f"[ERROR] PPTX faylni o'qishda xato: {path} - {e}")
        return ""

def read_code_file(path):
    """Kod faylini o'qiydi."""
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
            if not content.strip():
                print(f"[WARNING] Kod fayl bo'sh: {path}")
            return content
    except FileNotFoundError:
        print(f"[ERROR] Kod fayl topilmadi: {path}")
        return ""
    except Exception as e:
        print(f"[ERROR] Kod faylni o'qishda xato: {path} - {e}")
        return ""

def read_html_xml(path):
    """HTML/XML faylini o'qiydi va teglarni olib tashlaydi."""
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
            text = re.sub(r"<[^>]+>", "", content)
            if not text.strip():
                print(f"[WARNING] HTML/XML fayldan matn o'qilmadi: {path}")
            return text
    except FileNotFoundError:
        print(f"[ERROR] HTML/XML fayl topilmadi: {path}")
        return ""
    except Exception as e:
        print(f"[ERROR] HTML/XML faylni o'qishda xato: {path} - {e}")
        return ""

def clean_text(text):
    """Matnni tozalaydi, lekin yangi qatorlarni saqlaydi."""
    if not text:
        return ""
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', ' ', text)
    text = re.sub(r' +', ' ', text)
    lines = text.split('\n')
    cleaned_lines = [line.strip() for line in lines if line.strip()]
    return '\n'.join(cleaned_lines).strip()

def read_any_file(path):
    """Istalgan fayl formatini o'qiydi."""
    ext = os.path.splitext(path)[1].lower()
    funcs = {
        ".txt": read_txt,
        ".pdf": read_pdf,
        ".docx": read_docx,
        ".csv": read_csv_file,
        ".json": read_json_file,
        ".xls": read_excel,
        ".xlsx": read_excel,
        ".pptx": read_pptx,
        ".html": read_html_xml,
        ".xml": read_html_xml,
        ".py": read_code_file,
        ".js": read_code_file,
        ".cpp": read_code_file,
        ".java": read_code_file,
        ".c": read_code_file,
        ".ts": read_code_file,
    }
    reader = funcs.get(ext, read_txt)
    try:
        content = reader(path)
        if not content or not content.strip():
            print(f"[ERROR] Fayldan matn o'qilmadi: {path}")
            return ""
        content = preprocess_text(content)
        content = clean_text(content)
        return content
    except Exception as e:
        print(f"[ERROR] Faylni o'qishda xato: {path} - {e}")
        return ""

# =================== Yagona TTS funksiyasi ===================
async def text_to_speech(user_id, text, file_path="output.mp3"):
    """Matnni ovozga aylantiradi."""
    text = (text or "").strip()
    if not text:
        raise ValueError("Matn bo'sh. Iltimos, matn yuboring.")

    text = preprocess_text(text)
    voice_key = users.get(str(user_id), {}).get("voice", "male")
    voice = VOICES.get(voice_key, "uz-UZ-SardorNeural")

    max_chunk = 4200
    chunks = []
    if len(text) <= max_chunk:
        chunks = [text]
    else:
        words = text.split()
        cur = []
        cur_len = 0
        for w in words:
            if cur_len + len(w) + 1 > max_chunk:
                chunks.append(" ".join(cur))
                cur = [w]
                cur_len = len(w)
            else:
                cur.append(w)
                cur_len += len(w) + 1
        if cur:
            chunks.append(" ".join(cur))

    tmp_files = []
    for i, chunk in enumerate(chunks):
        success = False
        for attempt in range(3):
            try:
                tmp_fd, tmp_name = tempfile.mkstemp(suffix=f"_{i}.mp3")
                os.close(tmp_fd)
                communicate = edge_tts.Communicate(chunk, voice)
                await communicate.save(tmp_name)

                if os.path.exists(tmp_name) and os.path.getsize(tmp_name) > 0:
                    tmp_files.append(tmp_name)
                    success = True
                    break
                else:
                    if os.path.exists(tmp_name):
                        os.remove(tmp_name)
                    raise RuntimeError("Audio fayl hosil bo'lmadi")
            except Exception as e:
                print(f"[ERROR] TTS urinish {attempt+1} xato: {e}")
                if os.path.exists(tmp_name):
                    try:
                        os.remove(tmp_name)
                    except:
                        pass
                await asyncio.sleep(1)
        if not success:
            for p in tmp_files:
                if os.path.exists(p):
                    try:
                        os.remove(p)
                    except:
                        pass
            raise RuntimeError("Ovoz yaratishda xatolik yuz berdi.")

    try:
        with open(file_path, "wb") as outfile:
            for p in tmp_files:
                with open(p, "rb") as infile:
                    outfile.write(infile.read())
    except Exception as e:
        print(f"[ERROR] Audio fayllarni birlashtirishda xato: {e}")
        for p in tmp_files:
            if os.path.exists(p):
                try:
                    os.remove(p)
                except:
                    pass
        raise RuntimeError("Audio fayllarni birlashtirishda xato.")

    for p in tmp_files:
        try:
            os.remove(p)
        except:
            pass

    if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
        raise RuntimeError("Yakuniy audio fayl hosil bo'lmadi.")

    return file_path

# =================== Bot va Dispatcher ===================
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

# =================== /start komandasi ===================
@dp.message(Command("start"))
async def start_cmd(message: types.Message):
    uid = str(message.from_user.id)
    if uid not in users:
        users[uid] = {
            "limit": 7, "tarif": 0, "usage_count": 0, "voice": "male",
            "pending_tarif": None, "pending_limit": None, "pending_price": None,
            "username": message.from_user.username or "",
            "expire_date": None, "start_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
    else:
        # Har qanday jarayonni tugatish: user_state va pending maydonlarni tozalash
        user_state.pop(message.from_user.id, None)
        users[uid]["pending_tarif"] = None
        users[uid]["pending_limit"] = None
        users[uid]["pending_price"] = None
        users[uid]["payment_status"] = None  # Agar mavjud bo'lsa

    save_users()

    u = users.get(uid, {})
    limit = u.get("limit", 0)
    if message.from_user.id in ADMIN_IDS:
        limit_text = "♾ Siz admin ekansiz — cheklov yo'q."
    elif limit == -1:
        limit_text = "♾ Cheksiz (1 oy) — siz premyum tarifdasiz."
    else:
        limit_text = str(limit)
    text = (
        "👋 Assalomu alaykum!\n"
        "🎧 Slushay botiga xush kelibsiz!\n\n"
        "🔹 Bu bot siz yuborgan matnni yoki faylni ovozga aylantirib beradi. Shunchaki menga matn yoki istalgan formatdagi faylni yuboring (faqat o'zbek tilida bo'lsin).\n\n"
        f"🔹 Sizning joriy limit: {limit_text}\n\n"
        "• 📋 Quyidagi tugmalar orqali kerakli bo'limni tanlang:\n"
        "• 📊 Mening holatim — profilingiz va mavjud limit haqida ma'lumot beradi.\n"
        "• 💳 Limitni oshirish — tariflar va to'lov usullari bilan tanishing.\n"
        "• 🎤 Ovoz tanlash — erkak yoki ayol ovozini tanlang.\n\n"
        "• bo'tning qisqa tugmalari \n"
        "• /start bo'tmi qayta ishga tushirish \n"
        "• /help bo'tda qandaydur muammo yoki kamchilik bo'sa buni adminga yuborish. \n\n\n"
        "✅ Boshlash uchun menga matn yoki fayl yuboring!"
    )
    await message.answer(text, reply_markup=main_keyboard())

@dp.message(Command("help"))
async def ask_problem(message: types.Message):
    await message.answer(
        "❗ Botda qanday muammo yuz bermoqda? Yoki qanday qiyinchilik bo'lyapti? \n"
        "Iltimos, batafsil yozib yuboring. Biz tez orada hal qilamiz."
    )
    user_state[message.from_user.id] = "awaiting_problem_description"


# =================== Global o'zgaruvchilar ===================
ACTIVE_TIMEOUT = 300  # 5 daqiqa (soniyalarda)
active_users = {}  # Faol foydalanuvchilarni kuzatish uchun dictionary

# =================== Foydalanuvchi faolligini yangilash ===================
def update_active_user(user_id, username):
    """Foydalanuvchining so'nggi o'zaro aloqa vaqtini yangilaydi."""
    active_users[str(user_id)] = {
        "username": username or "Noma'lum",
        "last_interaction": time.time()
    }

# =================== Faol foydalanuvchilarni olish ===================
def get_active_users():
    """So'nggi ACTIVE_TIMEOUT ichida faol bo'lgan foydalanuvchilarni qaytaradi."""
    current_time = time.time()
    active = []
    for uid, info in active_users.items():
        if current_time - info["last_interaction"] <= ACTIVE_TIMEOUT:
            user_data = users.get(uid, {})
            active.append({
                "user_id": uid,
                "username": info["username"],
                "limit": user_data.get("limit", 0),
                "tarif": user_data.get("tarif", 0),
                "usage_count": user_data.get("usage_count", 0),
                "expire_date": user_data.get("expire_date", None)
            })
    return active

# =================== Admin: /status komandasi ===================
@dp.message(Command("status"))
async def cmd_status(message: types.Message):
    update_active_user(message.from_user.id, message.from_user.username)
    if message.from_user.id not in ADMIN_IDS:
        await message.answer("❌ Siz admin emassiz.")
        return

    if not users:
        await message.answer("👥 Hali foydalanuvchilar yo'q.")
        return

    total_users = len(users)
    total_requests = sum(u.get("usage_count", 0) for u in users.values())
    active = get_active_users()
    today = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

    text = (
        f"📊 <b>Bot statistikasi ({today})</b>\n\n"
        f"👥 Umumiy foydalanuvchilar soni: <b>{total_users}</b>\n"
        f"🗣 Jami so'rovlar soni: <b>{total_requests}</b>\n"
        f"👤 Faol foydalanuvchilar soni: <b>{len(active)}</b>\n"
        f"🛠 Bot ishlamoqda 🔵\n\n"
    )

    if active:
        tarif_name = {20: "20 limit", 25: "25 limit", -1: "Cheksiz (1 oy)", 0: "Bepul"}
        text += "<b>Faol foydalanuvchilar ro'yxati:</b>\n"
        for user in active:
            limit = "♾ Cheksiz" if user["limit"] == -1 else str(user["limit"])
            tarif = tarif_name.get(user["tarif"], "Noma'lum")
            expire = user["expire_date"] or "Yo'q"
            text += (
                f"👤 Username: @{user['username']}\n"
                f"🆔 ID: {user['user_id']}\n"
                f"📦 Tarif: {tarif}\n"
                f"🔢 Limit: {limit}\n"
                f"📈 Ishlatilgan: {user['usage_count']}\n"
                f"📅 Tugash sanasi: {expire}\n"
                f"{'-'*30}\n"
            )
    else:
        text += "👥 Hozirda faol foydalanuvchilar yo'q.\n"

    await message.answer(text, parse_mode="HTML")
# =================== /give admin komandasi ===================
@dp.message(Command("give"))
async def cmd_give_limit(message: types.Message):
    admin_id = message.from_user.id

    if admin_id not in ADMIN_IDS:
        await message.answer("❌ Bu funksiya faqat adminlar uchun!")
        return

    args = message.text.split()
    if len(args) < 3:
        await message.answer("ℹ️ Foydalanish: <b>/give @username yoki chat_id limit_soni</b>\nMasalan:\n/give @testuser 3\n/give 123456789 3\n/give @testuser -1 (1 oy cheksiz)\n/give @testuser -2 (1 hafta cheksiz)", parse_mode="HTML")
        return

    identifier = args[1]
    try:
        limit_value = int(args[2])
    except ValueError:
        await message.answer("❌ Limit soni faqat butun son bo'lishi kerak.")
        return

    target_user = None
    if identifier.startswith("@"):
        username = identifier.replace("@", "")
        for uid, info in users.items():
            if info.get("username") == username:
                target_user = uid
                break
    elif identifier.isdigit():
        target_user = identifier

    if not target_user:
        target_user = identifier if identifier.isdigit() else str(abs(hash(identifier)))
        users[target_user] = {
            "limit": 0,
            "tarif": 0,
            "usage_count": 0,
            "voice": "male",
            "pending_tarif": None,
            "pending_limit": None,
            "pending_price": None,
            "username": identifier.replace("@", "") if identifier.startswith("@") else "",
            "expire_date": None,
            "start_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

    target_id = target_user
    username_display = identifier if identifier.startswith("@") else f"ID {identifier}"

    if limit_value == -1:
        expires = datetime.now() + timedelta(days=30)
        users[target_user]["limit"] = -1
        users[target_user]["expire_date"] = expires.strftime("%Y-%m-%d %H:%M:%S")
        await message.answer(f"✅ <b>{username_display}</b> foydalanuvchisiga 1 oyga cheksiz limit berildi!", parse_mode="HTML")
        try:
            await bot.send_message(int(target_id), "🎉 Admin sizga 1 oyga cheksiz limit berdi! Endi cheklovsiz foydalaning.")
        except:
            await message.answer(f"⚠️ Foydalanuvchiga xabar yuborib bo'lmadi, lekin limit qo'shildi.")
    elif limit_value == -2:
        expires = datetime.now() + timedelta(days=7)
        users[target_user]["limit"] = -1
        users[target_user]["expire_date"] = expires.strftime("%Y-%m-%d %H:%M:%S")
        await message.answer(f"✅ <b>{username_display}</b> foydalanuvchisiga 1 hafta cheksiz limit berildi!", parse_mode="HTML")
        try:
            await bot.send_message(int(target_id), "🎉 Admin sizga 1 hafta cheksiz limit berdi! Endi cheklovsiz foydalaning.")
        except:
            await message.answer(f"⚠️ Foydalanuvchiga xabar yuborib bo'lmadi, lekin limit qo'shildi.")
    else:
        current_limit = users[target_user].get("limit", 0)
        if current_limit == -1:
            await message.answer("❌ Foydalanuvchi allaqachon cheksiz limitga ega, yangi limit qo'shish mumkin emas.")
            return
        else:
            new_limit = current_limit + limit_value
            users[target_user]["limit"] = new_limit
            await message.answer(f"✅ <b>{username_display}</b> foydalanuvchisiga {limit_value} ta limit qo'shildi. Yangi limit: {new_limit}", parse_mode="HTML")
            try:
                await bot.send_message(int(target_id), f"🎉 Admin sizga {limit_value} ta qo'shimcha limit berdi! Yangi limitingiz: {new_limit}")
            except:
                await message.answer(f"⚠️ Foydalanuvchiga xabar yuborib bo'lmadi, lekin limit qo'shildi.")
    save_users()

# =================== Mening holatim ===================
@dp.message(F.text == "📊 Mening holatim")
async def my_status(message: types.Message):
    uid = str(message.from_user.id)
    u = users.get(uid, {})
    tarif_name = {20: "20 limit", 25: "25 limit", -1: "Cheksiz (1 oy)", 0: "Bepul"}

    if message.from_user.id in ADMIN_IDS:
        text = (
            "📊 Siz admin ekansiz.\n"
            "🔹 Cheklovlar yo'q — sizning limitingiz cheksiz.\n"
        )
        await message.answer(text, reply_markup=main_keyboard())
        return

    limit = u.get("limit", 0)
    usage = u.get("usage_count", 0)
    expire = u.get("expire_date")
    text = "📊 Sizning holatingiz:\n"
    text += f"🔹 Tarif: {tarif_name.get(u.get('tarif', 0), 'Nomaʼlum')}\n"
    text += f"🔹 Limit: {'♾ Cheksiz' if limit == -1 else limit}\n"
    text += f"🔹 Ishlatilgan: {usage}\n"
    if expire:
        text += f"🔹 Amal qilish muddati: {expire[:10]}"
    await message.answer(text, reply_markup=main_keyboard())

# =================== Tariflar ===================
@dp.message(F.text == "💳 limitni oshirish")
async def show_tarifs(message: types.Message):
    await message.answer("📦 Quyidagi tariflardan birini tanlang:", reply_markup=tarif_buttons())
    user_state[message.from_user.id] = "awaiting_tarif"

# =================== Ovoz tanlash ===================
@dp.message(F.text == "🎤 Ovoz tanlash")
async def ask_voice(message: types.Message):
    kb = ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="👨 Erkak ovozi"), KeyboardButton(text="👩 Ayol ovozi")]
        ],
        resize_keyboard=True
    )
    await message.answer("🎤 Qaysi ovozda gaplashishni xohlaysiz?", reply_markup=kb)

@dp.message(F.text.in_(["👨 Erkak ovozi", "👩 Ayol ovozi"]))
async def set_voice(message: types.Message):
    uid = str(message.from_user.id)
    users[uid]["voice"] = "male" if message.text == "👨 Erkak ovozi" else "female"
    save_users()
    await message.answer("✅ Ovoz tanlandi.", reply_markup=main_keyboard())

# =================== Fayl yuborish bo'limi ===================
@dp.message(F.text == "📄 Faylni audioga aylantirish")
async def ask_file(message: types.Message):
    await message.answer(
        "📂 Iltimos, menga istalgan fayl yuboring (.txt, .pdf, .docx, .csv, .json, .html, .xlsx, .pptx, .py va hokazo).\n"
        "❗ 50 MBgacha fayl yuboring.\n\n"
        "Men fayl ichidagi MATNni o'qib, ovozga aylantiraman."
    )
    user_state[message.from_user.id] = None

# =================== Admin: /send komandasi ===================
@dp.message(Command("send"))
async def cmd_send(message: types.Message):
    if message.from_user.id not in ADMIN_IDS:
        await message.answer("❌ Siz admin emassiz.")
        return

    await message.answer("Barcha foydalanuvchilarga nima haqida ma'lumot bermoqchisiz? \nXabarni yuboring (matn, rasm, video, fayl va hokazo).")
    user_state[message.from_user.id] = "awaiting_broadcast_message"

# =================== Faylni qabul qilish handleri ===================
import logging
import aiohttp
logging.basicConfig(level=logging.INFO)


@dp.message(F.document)
async def handle_file(message: types.Message, bot: Bot):
    state = user_state.get(message.from_user.id)
    if state == "awaiting_broadcast_message" and message.from_user.id in ADMIN_IDS:
        sent_count = 0
        for uid in list(users.keys()):
            try:
                await bot.copy_message(chat_id=int(uid), from_chat_id=message.chat.id, message_id=message.message_id)
                sent_count += 1
            except Exception as e:
                print(f"[ERROR] Xabar yuborishda xato foydalanuvchi {uid} ga: {e}")
        await message.answer(f"✅ Xabar {sent_count} ta foydalanuvchiga yuborildi.")
        user_state.pop(message.from_user.id, None)
        return

    if state == "awaiting_payment":
        await message.answer("❗ Faqat pastdagi tugmadan foydalaning.")
        return
    if state == "awaiting_check":
        await message.answer("❗ To'lov tasdiqlash uchun faqat rasm yuboring.")
        return
    uid = str(message.from_user.id)
    file_id = message.document.file_id
    file_name = message.document.file_name or f"file_{int(time.time())}"
    temp_dir = tempfile.gettempdir()
    temp_path = os.path.join(temp_dir, file_name)

    u = users.get(uid, {})
    if message.from_user.id not in ADMIN_IDS and u.get('limit', 0) <= 0 and u.get('limit') != -1:
        await message.answer("❌ Sizning limitingiz tugagan. Botdan yana foydalanish uchun limitni oshirish pastdagi tugma orqali yangi limit sotib oling.", reply_markup=main_keyboard())
        return

    try:
        file = await bot.get_file(file_id)
        file_path = file.file_path
        file_url = f"https://api.telegram.org/file/bot{BOT_TOKEN}/{file_path}"
        await message.answer(f"📥 Fayl yuklab olinmoqda... ")

        # Faylni to‘g‘ridan-to‘g‘ri yuklab olish
        async with aiohttp.ClientSession() as session:
            async with session.get(file_url, timeout=aiohttp.ClientTimeout(total=300)) as resp:
                if resp.status == 200:
                    with open(temp_path, "wb") as f:
                        f.write(await resp.read())
                    await message.answer(f"📥 Fayl yuklab olindi... ")
                else:
                    await message.answer(f"❌ Yuklab olishda xato: HTTP {resp.status}")
                    return
    except Exception as e:
        await message.answer(f"❗ Fayl yuklab olishda xato: {e}")
        return

    try:
        content = read_any_file(temp_path)
        print("content")
        if not content or not content.strip():
            print(f"[ERROR] Fayldan matn o'qilmadi: {temp_path}")
            await message.answer("❌ Faylda matn topilmadi yoki format o'qilmaydi.")
            if os.path.exists(temp_path):
                os.remove(temp_path)
            return

        if len(content) > 20000:
            content = content[:15000] + "..."

        await message.answer("🎙 Fayl ichidagi matn ovozga aylantirilmoqda...")

        out_path = os.path.join(temp_dir, f"{message.from_user.id}_tts.mp3")
        await text_to_speech(message.from_user.id, content, out_path)

        await message.answer_voice(FSInputFile(out_path), caption="🎧 Mana faylingizning ovozli shakli:")

        if message.from_user.id not in ADMIN_IDS and u.get('limit') != -1:
            users[uid]['usage_count'] = u.get('usage_count', 0) + 1
            users[uid]['limit'] = max(0, u.get('limit', 0) - 1)
            save_users()
            remaining_limit = users[uid]['limit']
            await message.answer(f"🔹 Sizda {remaining_limit} ta limit qoldi.", reply_markup=main_keyboard())

    except Exception as e:
        await message.answer(f"❗ Xatolik yuz berdi: {e}")
    finally:
        for path in [temp_path, locals().get('out_path', None)]:
            if path and os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

# =================== Video handleri ===================
@dp.message(F.video)
async def handle_video(message: types.Message):
    state = user_state.get(message.from_user.id)
    if state == "awaiting_broadcast_message" and message.from_user.id in ADMIN_IDS:
        sent_count = 0
        for uid in list(users.keys()):
            try:
                await bot.copy_message(chat_id=int(uid), from_chat_id=message.chat.id, message_id=message.message_id)
                sent_count += 1
            except Exception as e:
                print(f"[ERROR] Xabar yuborishda xato foydalanuvchi {uid} ga: {e}")
        await message.answer(f"✅ Xabar {sent_count} ta foydalanuvchiga yuborildi.")
        user_state.pop(message.from_user.id, None)
        return

    if state == "awaiting_payment":
        await message.answer("❗ Faqat pastdagi tugmadan foydalaning.")
        return
    if state == "awaiting_check":
        await message.answer("❗ To'lov tasdiqlash uchun faqat rasm yuboring.")
        return

    uid = str(message.from_user.id)
    u = users.get(uid, {})
    if message.from_user.id not in ADMIN_IDS and u.get('limit', 0) <= 0 and u.get('limit') != -1:
        await message.answer("❌ Sizning limitingiz tugagan. Botdan yana foydalanish uchun limitni oshirish pastdagi tugma orqali yangi limit sotib oling.", reply_markup=main_keyboard())
        return

    if message.caption:
        text = message.caption
        await message.answer("🗣 Video ostidagi matn ovozga aylantirilmoqda...")
        out_path = f"{message.from_user.id}_video_caption.mp3"
        try:
            await text_to_speech(message.from_user.id, text, out_path)
            await message.answer_voice(FSInputFile(out_path), caption="🎧 Mana video ostidagi matningizning ovozli shakli:")

            if message.from_user.id not in ADMIN_IDS and u.get('limit') != -1:
                users[uid]['usage_count'] = u.get('usage_count', 0) + 1
                users[uid]['limit'] = max(0, u.get('limit', 0) - 1)
                save_users()
                remaining_limit = users[uid]['limit']
                await message.answer(f"🔹 Sizda {remaining_limit} ta limit qoldi.", reply_markup=main_keyboard())
        except Exception as e:
            await message.answer(f"❗ Ovoz yaratishda xato: {e}", reply_markup=main_keyboard())
        finally:
            if os.path.exists(out_path):
                try:
                    os.remove(out_path)
                except:
                    pass
    else:
        await message.answer("Iltimos menga oddiy matn, fayl tashlang yoki rasm, video ostida matni bor turdagi narsalarni tashlang men ularni sizga o'qib beraman", reply_markup=main_keyboard())

# =================== Rasm (caption va chek) handleri ===================
@dp.message(F.photo)
async def handle_photo(message: types.Message):
    state = user_state.get(message.from_user.id)
    if state == "awaiting_broadcast_message" and message.from_user.id in ADMIN_IDS:
        sent_count = 0
        for uid in list(users.keys()):
            try:
                await bot.copy_message(chat_id=int(uid), from_chat_id=message.chat.id, message_id=message.message_id)
                sent_count += 1
            except Exception as e:
                print(f"[ERROR] Xabar yuborishda xato foydalanuvchi {uid} ga: {e}")
        await message.answer(f"✅ Xabar {sent_count} ta foydalanuvchiga yuborildi.")
        user_state.pop(message.from_user.id, None)
        return

    if state == "awaiting_payment":
        await message.answer("❗ Faqat pastdagi tugmadan foydalaning.")
        return
    if state == "awaiting_check":
        uid = str(message.from_user.id)
        u = users.get(uid)
        if not u or not u.get("pending_tarif"):
            await message.answer("❌ Avval tarif tanlang va to'lov tasdiqlang.")
            return

        tarif = u["pending_tarif"]
        price = u["pending_price"]
        username = message.from_user.username or "Noma'lum"
        full_name = message.from_user.full_name or "Noma'lum"
        caption = (
            f"📥 Yangi to'lov cheki:\n"
            f"Foydalanuvchi: @{username} ({full_name})\n"
            f"ID: {message.from_user.id}\n"
            f"Tarif: {tarif}\n"
            f"Summa: {price} so'm"
        )

        inline_kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✅ Tasdiqlash", callback_data=f"approve|{uid}")],
            [InlineKeyboardButton(text="❌ Rad etish", callback_data=f"reject|{uid}")]
        ])

        for admin_id in ADMIN_IDS:
            try:
                await bot.send_photo(admin_id, message.photo[-1].file_id, caption=caption, reply_markup=inline_kb)
            except:
                pass

        await message.answer("✅ Chek adminga yuborildi. Tasdiqlanishini kuting.", reply_markup=main_keyboard())
        user_state.pop(message.from_user.id, None)
        u["payment_status"] = "pending"
        save_users()
    elif message.caption:
        text = message.caption
        uid = str(message.from_user.id)
        u = users.get(uid, {})
        if message.from_user.id not in ADMIN_IDS and u.get('limit', 0) <= 0 and u.get('limit') != -1:
            await message.answer("❌ Sizning limitingiz tugagan. Botdan yana foydalanish uchun limitni oshirish pastdagi tugma orqali yangi limit sotib oling.", reply_markup=main_keyboard())
            return

        await message.answer("🗣 Rasm ostidagi matn ovozga aylantirilmoqda...")
        out_path = f"{message.from_user.id}_caption.mp3"
        try:
            await text_to_speech(message.from_user.id, text, out_path)
            await message.answer_voice(
                voice=FSInputFile(out_path),
                caption=(
                    "🎧 <b>Mana matningizning ovozli shakli:</b><br><br>"
                    "• <b>/start</b> — bo‘tni qayta ishga tushirish<br>"
                    "• <b>/help</b> — bo‘t ishlamayotgan bo‘lsa yoki kamchilik bo‘lsa, adminga yuborish"
                ),
                parse_mode="HTML"
            )
            if message.from_user.id not in ADMIN_IDS and u.get('limit') != -1:
                users[uid]['usage_count'] = u.get('usage_count', 0) + 1
                users[uid]['limit'] = max(0, u.get('limit', 0) - 1)
                save_users()
                remaining_limit = users[uid]['limit']
                await message.answer(f"🔹 Sizda {remaining_limit} ta limit qoldi.", reply_markup=main_keyboard())
        except Exception as e:
            await message.answer(f"❗ Ovoz yaratishda xato: {e}", reply_markup=main_keyboard())
        finally:
            if os.path.exists(out_path):
                try:
                    os.remove(out_path)
                except:
                    pass
    else:
        await message.answer("🖼 Bu rasmda matn (rasm ostida matn ) yo'q.", reply_markup=main_keyboard())

# =================== Matnli xabarlar ===================
@dp.message()
async def handle_message(message: types.Message):
    uid = str(message.from_user.id)
    text = (message.text or "").strip()

    state = user_state.get(message.from_user.id)

    if state == "awaiting_broadcast_message" and message.from_user.id in ADMIN_IDS:
        sent_count = 0
        for target_uid in list(users.keys()):
            try:
                await bot.copy_message(chat_id=int(target_uid), from_chat_id=message.chat.id, message_id=message.message_id)
                sent_count += 1
            except Exception as e:
                print(f"[ERROR] Xabar yuborishda xato foydalanuvchi {target_uid} ga: {e}")
        await message.answer(f"✅ Xabar {sent_count} ta foydalanuvchiga yuborildi.")
        user_state.pop(message.from_user.id, None)
        return

    if message.text and message.text.startswith("/status"):
        return 
    if message.from_user.id in ADMIN_IDS and text.startswith("/give"):
        return

    state = user_state.get(message.from_user.id)
    
    if state and state.startswith("awaiting_reject_reason"):
        _, rejected_uid = state.split("|")
        rejected_uid = str(rejected_uid)
        u = users.get(rejected_uid)
        if not u:
            await message.answer("❌ Foydalanuvchi topilmadi.")
            user_state.pop(message.from_user.id, None)
            return

        reason = text.strip()
        if not reason:
            await message.answer("❌ Sabab bo'sh bo'lishi mumkin emas. Iltimos, sabab kiriting.")
            return

        u["pending_tarif"] = None
        u["pending_limit"] = None
        u["pending_price"] = None
        u["payment_status"] = "rejected"
        save_users()

        try:
            await bot.send_message(int(rejected_uid), f"❌ To'lovingiz rad etildi.\n\n📝 Sababi: {reason}")
        except Exception as e:
            print(f"[ERROR] Userga xabar yuborishda xato: {e}")
            await message.answer(f"⚠️ Xabar yuborib bo'lmadi, lekin to'lov rad etildi.")
        
        await message.answer("✅ Sabab foydalanuvchiga yuborildi.", reply_markup=main_keyboard())
        user_state.pop(message.from_user.id, None)
        return

    if state == "awaiting_problem_description":
        if not text:
            await message.answer("❗ Iltimos, muammo haqida matn yuboring.")
            return

        username = message.from_user.username or "Noma'lum"
        full_name = message.from_user.full_name or "Noma'lum"
        caption = (
            f"📥 Yangi bot muammosi haqida xabar:\n"
            f"Foydalanuvchi: @{username} ({full_name})\n"
            f"ID: {message.from_user.id}\n"
            f"Muammo tavsifi: {text}"
        )

        for admin_id in ADMIN_IDS:
            try:
                await bot.send_message(admin_id, caption)
            except:
                pass

        await message.answer("✅ Rahmat! Muammo yoki kamchilik uchun uzr so'raymiz. Adminlar bu muammoni tez orada o'rganib chiqib, sizga izoh qoldirishadi.", reply_markup=main_keyboard())
        user_state.pop(message.from_user.id, None)
        return

    if state == "awaiting_check":
        await message.answer("❗ To'lov tasdiqlash uchun faqat rasm yuboring.")
        return

    if state == "awaiting_tarif":
        if text not in tariflar:
            await message.answer("❗ Faqat pastdagi tug.border.")
            return
        t = tariflar[text]
        users[uid].update({"pending_tarif": text, "pending_limit": t["limit"], "pending_price": t["price"]})
        save_users()
        kb = ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text="✅ To'lovni tasdiqlash")]], resize_keyboard=True)
        await message.answer(
            f"💳 Siz {text} tanladingiz.\n"
            f"To'lov summasi: {t['price']} so'm\n"
            f"Karta raqami: <code>9860 1606 1980 9199</code>\n\n"
            f"To'lovni amalga oshirgandan so'ng, chekni yuboring.",
            parse_mode="HTML",
            reply_markup=kb
        )
        user_state[message.from_user.id] = "awaiting_check"
        return

    if text == "✅ To'lovni tasdiqlash":
        await message.answer("📸 Iltimos, to'lov chekining rasmini yuboring.")
        user_state[message.from_user.id] = "awaiting_check"
        return

    u = users.get(uid, {})
    if message.from_user.id not in ADMIN_IDS and u.get('limit', 0) <= 0 and u.get('limit') != -1:
        await message.answer("❌ Sizning limitingiz tugagan. Botdan yana foydalanish uchun limitni oshirish pastdagi tugma orqali yangi limit sotib oling.", reply_markup=main_keyboard())
        return

    if not text:
        await message.answer("❗ Iltimos, matn yuboring.", reply_markup=main_keyboard())
        return

    if len(text) > 20000:
        text = text[:15000] + "..."

    await message.answer("🎙 Matn ovozga aylantirilmoqda...")
    out_path = f"{message.from_user.id}_tts.mp3"
    try:
        await text_to_speech(message.from_user.id, text, out_path)
        await message.answer_voice(FSInputFile(out_path), caption="🎧 Mana matningizning ovozli shakli:")

        if message.from_user.id not in ADMIN_IDS and u.get('limit') != -1:
            users[uid]['usage_count'] = u.get('usage_count', 0) + 1
            users[uid]['limit'] = max(0, u.get('limit', 0) - 1)
            save_users()
            remaining_limit = users[uid]['limit']
            await message.answer(f"🔹 Sizda {remaining_limit} ta limit qoldi.", reply_markup=main_keyboard())
    except Exception as e:
        await message.answer(f"❗ Ovoz yaratishda xato: {e}", reply_markup=main_keyboard())
    finally:
        if os.path.exists(out_path):
            try:
                os.remove(out_path)
            except:
                pass

# =================== Callback handleri ===================
@dp.callback_query()
async def handle_callback(query: types.CallbackQuery):
    data = query.data
    if data.startswith("approve|"):
        uid = data.split("|")[1]
        u = users.get(uid)
        if not u:
            await query.answer("❌ Foydalanuvchi topilmadi.")
            return

        limit = u.get("pending_limit", 0)
        if limit == -1:
            expires = datetime.now() + timedelta(days=30)
            u["limit"] = -1
            u["expire_date"] = expires.strftime("%Y-%m-%d %H:%M:%S")
        else:
            u["limit"] = u.get("limit", 0) + limit

        u["pending_tarif"] = None
        u["pending_limit"] = None
        u["pending_price"] = None
        u["payment_status"] = "approved"
        save_users()

        await query.answer("✅ To'lov tasdiqlandi!")
        await query.message.edit_text("✅ To'lov tasdiqlandi!")
        try:
            await bot.send_message(int(uid), "🎉 To'lovingiz tasdiqlandi! Yangi limitingiz faol.")
        except:
            pass

    elif data.startswith("reject|"):
        uid = data.split("|")[1]
        admin_id = query.from_user.id
        user_state[admin_id] = f"awaiting_reject_reason|{uid}"
        await query.answer("✅ Sababni kiriting")
        await query.message.edit_text("📝 Rad etish sababini kiriting:")

# =================== Main ===================
async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    print("bo't ishhga tushdi")
    asyncio.run(main())
