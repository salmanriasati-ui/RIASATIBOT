import os
import re
import json
import random
import string
import asyncio
import base64
import tempfile
from datetime import datetime
from pathlib import Path

from telegram import Update, Bot
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    filters, ContextTypes, ConversationHandler
)
import anthropic
from playwright.async_api import async_playwright
from pypdf import PdfWriter, PdfReader
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ============================================================
# CONFIG
# ============================================================
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "8811395656:AAEV2gOsvymp2sTvnLQjUrrypW3pno0I9Po").strip()
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "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").strip()

print(f"TOKEN starts with: {TELEGRAM_TOKEN[:10] if TELEGRAM_TOKEN else 'EMPTY!'}")
print(f"API KEY starts with: {ANTHROPIC_API_KEY[:15] if ANTHROPIC_API_KEY else 'EMPTY!'}")
EXCEL_FILE = "IRANFLY_tickets.xlsx"
TEMPLATE_FILE = "IRANFLY_TEMPLATE_v1.html"
PDF_PASSWORD = "IRANFLY2025"

# Flight info
FLIGHTS = {
    "qeshm_ika_tbs": {"airline": "Qeshm Airline", "flight": "Q.2273", "time": "11:00", "from": "IKA", "to": "TBS"},
    "varesh_ika_tbs": {"airline": "Varesh Airline", "flight": "VR.6808", "time": "11:00", "from": "IKA", "to": "TBS"},
    "qeshm_tbs_ika": {"airline": "Qeshm Airline", "flight": "Q.2272", "time": "14:50", "from": "TBS", "to": "IKA"},
    "varesh_tbs_ika": {"airline": "Varesh Airline", "flight": "VR.6809", "time": "14:15", "from": "TBS", "to": "IKA"},
}

BAGGAGE = {
    "qeshm_economy": "30 kg",
    "qeshm_business": "40 kg",
    "varesh_economy": "20 kg",
}

# States
WAITING_PASSPORT = 1
WAITING_FLIGHT_INFO = 2

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

# ============================================================
# HELPERS
# ============================================================
def gen_ref():
    """Generate 5-char mixed alphanumeric reference like CJ8H1"""
    while True:
        ref = ''.join(random.choices(string.ascii_uppercase + string.digits, k=5))
        if any(c.isalpha() for c in ref) and any(c.isdigit() for c in ref):
            return ref

def extract_passport_info(image_bytes: bytes) -> dict:
    """Use Claude to extract passport info from image"""
    b64 = base64.b64encode(image_bytes).decode()
    response = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=500,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}
                },
                {
                    "type": "text",
                    "text": "Extract from this passport image: first_name, last_name, passport_number. Return ONLY JSON like: {\"first_name\": \"JOHN\", \"last_name\": \"DOE\", \"passport_number\": \"A12345678\"}"
                }
            ]
        }]
    )
    text = response.content[0].text.strip()
    text = re.sub(r'```json|```', '', text).strip()
    return json.loads(text)

def parse_flight_input(text: str) -> dict:
    """Parse flight info text"""
    text = text.lower().strip()
    result = {
        "type": "oneway",
        "outbound": None,
        "return": None,
        "sell_out": 0,
        "buy_out": 0,
        "sell_ret": 0,
        "buy_ret": 0,
        "voucher": 0,
        "insurance": 0,
        "class": "economy",
    }

    if "business" in text:
        result["class"] = "business"

    # Detect route direction
    has_ika_tbs = "ika" in text and "tbs" in text
    has_tbs_ika = "tbs" in text and "ika" in text

    # Detect airline
    airline = "qeshm" if "qeshm" in text else "varesh"

    # Detect if round trip (two dates or "round" or both directions)
    lines = [l.strip() for l in text.split('\n') if l.strip()]

    # Simple: if single line → one-way
    if len(lines) == 1:
        result["type"] = "oneway"
        # Detect direction
        if text.index("ika") < text.index("tbs"):
            key = f"{airline}_ika_tbs"
        else:
            key = f"{airline}_tbs_ika"
        result["outbound"] = FLIGHTS[key]
    else:
        result["type"] = "roundtrip"

    # Extract prices (pattern: number.number or number)
    prices = re.findall(r'(\d+[\.,]\d+|\d+)\s*(?:sell|فروش)', text)
    buys = re.findall(r'(\d+[\.,]\d+|\d+)\s*(?:buy|خرید)', text)

    def to_num(s):
        return int(float(s.replace(',', '.')) * 1000) if '.' in s or ',' in s else int(s)

    if prices:
        result["sell_out"] = to_num(prices[0])
    if buys:
        result["buy_out"] = to_num(buys[0])

    # Extract date
    months = {"jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,
              "jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12}
    date_match = re.search(r'(\d{1,2})\s*(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)', text)
    if date_match:
        day = date_match.group(1)
        month = date_match.group(2)
        result["date_out"] = f"{day} {month.capitalize()} 2026"

    # Voucher/insurance
    vch = re.search(r'(?:voucher|واچر)[:\s]*(\d+)', text)
    ins = re.search(r'(?:insurance|بیمه)[:\s]*(\d+)', text)
    if vch:
        result["voucher"] = int(vch.group(1))
    if ins:
        result["insurance"] = int(ins.group(1))

    return result

async def generate_pdf(passenger: dict, flight: dict, template_path: str) -> bytes:
    """Generate PDF ticket using Playwright"""
    with open(template_path, 'r') as f:
        html = f.read()

    fname = passenger["first_name"]
    lname = passenger["last_name"]
    pp = passenger["passport_number"]
    ref = gen_ref()

    airline_info = flight.get("outbound") or FLIGHTS.get("varesh_tbs_ika")
    sell = flight.get("sell_out", 0)
    date = flight.get("date_out", "")
    cls = flight.get("class", "economy")
    baggage_key = f"{('qeshm' if 'Qeshm' in airline_info['airline'] else 'varesh')}_{cls}"
    baggage = BAGGAGE.get(baggage_key, "20 kg")

    # Hide return flight for one-way
    is_oneway = flight.get("type") == "oneway"
    if is_oneway:
        html = html.replace(
            '.connector {',
            '.connector { display:none !important; } .flight-card:first-child { display:none !important; } .x_{'
        )

    # Update passenger
    html = html.replace('MAHMOUD<br>ESLAMINOSRATABADI', f'{fname}<br>{lname}')
    html = html.replace('MAHMOUD ESLAMINOSRATABADI', f'{fname} {lname}')
    html = html.replace('B74341095', pp)

    # Update flight details (return leg for TBS→IKA one-way)
    src = airline_info["from"]
    dst = airline_info["to"]
    flight_num = airline_info["flight"]
    flight_time = airline_info["time"]
    airline_name = airline_info["airline"]

    # Update booking ref
    html = html.replace('VR3M84', ref)
    html = html.replace('QA7X29', ref)

    # Update price
    price_str = f"{sell:,}"
    html = re.sub(
        r'<span class="price-value">\$\d+</span>',
        f'<span class="price-value" style="font-size:18px;">{price_str}<br><span style="font-size:10px;font-weight:500;">تومان</span></span>',
        html
    )

    # Update date
    if date:
        html = re.sub(r'(29|26|28) May 2026', date, html)

    # Update baggage
    html = html.replace('>30 kg<', f'>{baggage}<')
    html = html.replace('>20 kg<', f'>{baggage}<')

    # Update class
    if cls == "business":
        html = html.replace('>Economy<', '>Business<')

    # Generate PDF
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        await page.set_content(html, wait_until='networkidle')
        await page.wait_for_timeout(1000)

        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
            tmp_path = tmp.name

        await page.pdf(
            path=tmp_path,
            format='A4',
            margin={'top': '8mm', 'bottom': '8mm', 'left': '8mm', 'right': '8mm'},
            print_background=True
        )
        await browser.close()

    # Lock PDF
    reader = PdfReader(tmp_path)
    writer = PdfWriter()
    for page_obj in reader.pages:
        writer.add_page(page_obj)
    writer.encrypt(user_password="", owner_password=PDF_PASSWORD, permissions_flag=4)

    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as locked:
        locked_path = locked.name
    with open(locked_path, 'wb') as f:
        writer.write(f)

    with open(locked_path, 'rb') as f:
        pdf_bytes = f.read()

    os.unlink(tmp_path)
    os.unlink(locked_path)

    return pdf_bytes, ref

def update_excel(passenger: dict, flight: dict, ref: str):
    """Add row to Excel file"""
    if not os.path.exists(EXCEL_FILE):
        return

    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb['بلیت‌ها']
    next_row = ws.max_row + 1

    light_blue = "EBF3FF"
    thin_gray = Side(style="thin", color="CCCCCC")

    airline_info = flight.get("outbound") or {}
    src = airline_info.get("from", "-")
    dst = airline_info.get("to", "-")
    airline_name = airline_info.get("airline", "-") + " " + airline_info.get("flight", "")
    date = flight.get("date_out", "-")

    is_round = flight.get("type") == "roundtrip"

    row_data = [
        next_row - 1,
        ref,
        f"{passenger['first_name']} {passenger['last_name']}",
        f"{src} → {dst}" if src else "-",
        f"{dst} → {src}" if is_round else "-",
        airline_name if src else "-",
        airline_name if is_round else "-",
        date,
        flight.get("date_ret", "-"),
        flight.get("buy_out", 0),
        flight.get("sell_out", 0),
        flight.get("voucher", 0),
        flight.get("insurance", 0),
        f"=(K{next_row}-J{next_row})+L{next_row}+M{next_row}"
    ]

    for col, val in enumerate(row_data, 1):
        cell = ws.cell(row=next_row, column=col, value=val)
        cell.font = Font(name="Arial", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
        cell.fill = PatternFill("solid", start_color=light_blue)

    for col in [10, 11, 12, 13, 14]:
        ws.cell(row=next_row, column=col).number_format = '#,##0'

    wb.save(EXCEL_FILE)

# ============================================================
# BOT HANDLERS
# ============================================================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "✈️ *سیستم صدور بلیت IRANFLY*\n\n"
        "لطفاً عکس پاسپورت مسافر را ارسال کنید.",
        parse_mode="Markdown"
    )
    return WAITING_PASSPORT

async def receive_passport(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📷 پاسپورت دریافت شد. در حال استخراج اطلاعات...")

    photo = update.message.photo[-1]
    file = await context.bot.get_file(photo.file_id)
    image_bytes = await file.download_as_bytearray()

    try:
        passport_info = extract_passport_info(bytes(image_bytes))
        context.user_data['passport'] = passport_info

        await update.message.reply_text(
            f"✅ اطلاعات استخراج شد:\n"
            f"👤 نام: {passport_info['first_name']} {passport_info['last_name']}\n"
            f"🔢 پاسپورت: {passport_info['passport_number']}\n\n"
            "حالا اطلاعات پرواز را وارد کنید:\n"
            "مثال: `tbs/ika/varesh/26may/2.400sell/1.800buy`",
            parse_mode="Markdown"
        )
        return WAITING_FLIGHT_INFO
    except Exception as e:
        await update.message.reply_text(f"❌ خطا در استخراج اطلاعات: {str(e)}")
        return WAITING_PASSPORT

async def receive_flight_info(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    passport = context.user_data.get('passport')

    if not passport:
        await update.message.reply_text("❌ ابتدا عکس پاسپورت ارسال کنید.")
        return WAITING_PASSPORT

    await update.message.reply_text("⏳ در حال صدور بلیت...")

    try:
        flight = parse_flight_input(text)

        # Determine flight route
        t = text.lower()
        if "qeshm" in t or "قشم" in t:
            airline_key = "qeshm"
        else:
            airline_key = "varesh"

        if t.index("ika") < t.index("tbs"):
            direction = "ika_tbs"
        else:
            direction = "tbs_ika"

        flight["outbound"] = FLIGHTS[f"{airline_key}_{direction}"]

        pdf_bytes, ref = await generate_pdf(passport, flight, TEMPLATE_FILE)

        fname = passport['first_name']
        lname = passport['last_name']
        filename = f"{fname}_{lname}.pdf"

        await update.message.reply_document(
            document=pdf_bytes,
            filename=filename,
            caption=f"🔖 رفرنس: `{ref}`",
            parse_mode="Markdown"
        )

        update_excel(passport, flight, ref)

        await update.message.reply_text(
            "مسافر محترم، مدارک شما در حال صدور بوده و در اسرع وقت ارسال خواهد شد. "
            "ممنون از صبر و شکیبایی شما. 🙏"
        )

        context.user_data.clear()
        return ConversationHandler.END

    except Exception as e:
        await update.message.reply_text(f"❌ خطا: {str(e)}\nلطفاً دوباره تلاش کنید.")
        return WAITING_FLIGHT_INFO

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("❌ لغو شد.")
    return ConversationHandler.END

# ============================================================
# MAIN
# ============================================================
def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            MessageHandler(filters.PHOTO, receive_passport)
        ],
        states={
            WAITING_PASSPORT: [MessageHandler(filters.PHOTO, receive_passport)],
            WAITING_FLIGHT_INFO: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_flight_info)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(conv)
    print("🚀 IRANFLY Bot started...")
    app.run_polling()

if __name__ == "__main__":
    main()
