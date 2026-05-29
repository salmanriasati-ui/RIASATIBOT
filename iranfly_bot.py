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

from telegram import Update, Bot, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    filters, ContextTypes, ConversationHandler, CallbackQueryHandler
)
import anthropic
from playwright.async_api import async_playwright
from pypdf import PdfWriter, PdfReader
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ============================================================
# CONFIG
# ============================================================
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "").strip()
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "").strip()

if not TELEGRAM_TOKEN:
    raise ValueError("TELEGRAM_TOKEN not set!")
if not ANTHROPIC_API_KEY:
    raise ValueError("ANTHROPIC_API_KEY not set!")

EXCEL_FILE = "IRANFLY_tickets.xlsx"
TEMPLATE_FILE = "IRANFLY_TEMPLATE_v1.html"
PDF_PASSWORD = "IRANFLY2025"

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
    "varesh_business": "20 kg",
}

WAITING_PASSPORT = 1
WAITING_FLIGHT_INFO = 2

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

def gen_ref():
    while True:
        ref = ''.join(random.choices(string.ascii_uppercase + string.digits, k=5))
        if any(c.isalpha() for c in ref) and any(c.isdigit() for c in ref):
            return ref

def extract_passport_info(image_bytes: bytes) -> dict:
    b64 = base64.b64encode(image_bytes).decode()
    response = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=500,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}},
            {"type": "text", "text": 'Extract from this passport: first_name, last_name, passport_number. Return ONLY JSON: {"first_name": "JOHN", "last_name": "DOE", "passport_number": "A12345678"}'}
        ]}]
    )
    text = re.sub(r'```json|```', '', response.content[0].text.strip()).strip()
    return json.loads(text)

def to_num(s):
    s = str(s).strip().replace(',', '.')
    if '.' in s:
        return int(float(s) * 1000)
    return int(s)

def parse_date(s):
    months = {"jan":"Jan","feb":"Feb","mar":"Mar","apr":"Apr","may":"May","jun":"Jun",
              "jul":"Jul","aug":"Aug","sep":"Sep","oct":"Oct","nov":"Nov","dec":"Dec"}
    s = s.lower().strip()
    m = re.match(r'(\d{1,2})\s*(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)', s)
    if m:
        return f"{m.group(1)} {months[m.group(2)]} 2026"
    return s

def parse_line(line):
    parts = [p.strip() for p in line.lower().split('/')]
    info = {
        "src": parts[0].upper() if len(parts) > 0 else "",
        "dst": parts[1].upper() if len(parts) > 1 else "",
        "airline": parts[2] if len(parts) > 2 else "varesh",
        "date": parse_date(parts[3]) if len(parts) > 3 else "",
        "sell": 0, "buy": 0, "voucher": 0, "insurance": 0,
        "flight_class": "business" if "business" in line.lower() else "economy"
    }
    for p in parts[4:]:
        p = p.strip()
        if p.startswith("sell"):
            info["sell"] = to_num(p[4:])
        elif p.startswith("buy"):
            info["buy"] = to_num(p[3:])
        elif p.startswith("voucher"):
            info["voucher"] = to_num(p[7:])
        elif p.startswith("insurance"):
            info["insurance"] = to_num(p[9:])
    return info

def parse_flight_input(text: str) -> dict:
    lines = [l.strip() for l in text.strip().split('\n') if l.strip()]
    result = {
        "type": "oneway", "outbound": None, "return_flight": None,
        "sell_out": 0, "buy_out": 0, "sell_ret": 0, "buy_ret": 0,
        "voucher": 0, "insurance": 0, "class": "economy",
        "date_out": "", "date_ret": ""
    }

    if len(lines) >= 1:
        out = parse_line(lines[0])
        airline = "qeshm" if "qeshm" in out["airline"] else "varesh"
        key = f"{airline}_ika_tbs" if out["src"] == "IKA" else f"{airline}_tbs_ika"
        result["outbound"] = FLIGHTS[key]
        result["sell_out"] = out["sell"]
        result["buy_out"] = out["buy"]
        result["voucher"] = out["voucher"]
        result["insurance"] = out["insurance"]
        result["class"] = out["flight_class"]
        result["date_out"] = out["date"]

    if len(lines) >= 2:
        result["type"] = "roundtrip"
        ret = parse_line(lines[1])
        airline_ret = "qeshm" if "qeshm" in ret["airline"] else "varesh"
        key_ret = f"{airline_ret}_ika_tbs" if ret["src"] == "IKA" else f"{airline_ret}_tbs_ika"
        result["return_flight"] = FLIGHTS[key_ret]
        result["sell_ret"] = ret["sell"]
        result["buy_ret"] = ret["buy"]
        result["date_ret"] = ret["date"]

    return result

async def generate_pdf(passenger: dict, flight: dict, template_path: str):
    with open(template_path, 'r') as f:
        html = f.read()

    fname = passenger["first_name"]
    lname = passenger["last_name"]
    pp = passenger["passport_number"]
    ref_out = gen_ref()
    ref_ret = gen_ref()

    is_oneway = flight.get("type") == "oneway"
    flight_class = flight.get("class", "economy")

    outbound = flight.get("outbound") or FLIGHTS["varesh_ika_tbs"]
    return_fl = flight.get("return_flight") or FLIGHTS["varesh_tbs_ika"]

    baggage_key_out = f"{'qeshm' if 'Qeshm' in outbound['airline'] else 'varesh'}_{flight_class}"
    baggage_key_ret = f"{'qeshm' if 'Qeshm' in return_fl['airline'] else 'varesh'}_{flight_class}"
    baggage_out = BAGGAGE.get(baggage_key_out, "20 kg")
    baggage_ret = BAGGAGE.get(baggage_key_ret, "20 kg")

    sell_out = flight.get("sell_out", 0)
    sell_ret = flight.get("sell_ret", 0)
    date_out = flight.get("date_out", "")
    date_ret = flight.get("date_ret", "")

    # For one-way, use outbound data for the visible card
    out_src = outbound.get("from", "IKA")
    if is_oneway and out_src == "TBS":
        # TBS→IKA one-way: show return card, hide outbound
        sell_ret = sell_out
        date_ret = date_out
        sell_out = 0
        date_out = ""
        # Use outbound flight info for return card
        return_fl = outbound

    if is_oneway:
        html = html.replace('.connector {', '.connector { display:none !important; } .flight-card:first-child { display:none !important; } .x_{')

    # Passenger
    html = html.replace('MAHMOUD<br>ESLAMINOSRATABADI', f'{fname}<br>{lname}')
    html = html.replace('MAHMOUD ESLAMINOSRATABADI', f'{fname} {lname}')
    html = html.replace('B74341095', pp)

    # Refs
    html = html.replace('QA7X29', ref_out)
    html = html.replace('VR3M84', ref_ret)

    # Dates
    html = html.replace('DATE_OUT', date_out if date_out else '---')
    html = html.replace('DATE_RET', date_ret if date_ret else '---')

    # Airline names and flight numbers
    html = html.replace('Qeshm Airline', outbound['airline'])
    html = html.replace('QESHM AIRLINE', outbound['airline'].upper())
    html = html.replace('Q.2273', outbound['flight'])
    html = html.replace('Varesh Airline', return_fl['airline'])
    html = html.replace('VARESH AIRLINE', return_fl['airline'].upper())
    html = html.replace('VR.6809', return_fl['flight'])

    # Prices
    price_out_str = f"{sell_out:,}"
    price_ret_str = f"{sell_ret:,}"
    html = html.replace('PRICE_OUT', f'{price_out_str}<br><span style="font-size:10px;font-weight:500;">تومان</span>')
    html = html.replace('PRICE_RET', f'{price_ret_str}<br><span style="font-size:10px;font-weight:500;">تومان</span>')

    # Baggage
    html = re.sub(r'<span class="ic-value">30 kg</span>', f'<span class="ic-value">{baggage_out}</span>', html, count=1)
    html = re.sub(r'<span class="ic-value">20 kg</span>', f'<span class="ic-value">{baggage_ret}</span>', html, count=1)

    # Class
    if flight_class == "business":
        html = html.replace('>Economy<', '>Business<')

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        await page.set_content(html, wait_until='networkidle')
        await page.wait_for_timeout(1000)

        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
            tmp_path = tmp.name

        await page.pdf(path=tmp_path, format='A4',
                       margin={'top':'8mm','bottom':'8mm','left':'8mm','right':'8mm'},
                       print_background=True)
        await browser.close()

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

    return pdf_bytes, ref_out, ref_ret

def update_excel(passenger: dict, flight: dict, ref_out: str, ref_ret: str):
    if not os.path.exists(EXCEL_FILE):
        return

    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb['بلیت‌ها']
    next_row = ws.max_row + 1

    light_blue = "EBF3FF"
    thin_gray = Side(style="thin", color="CCCCCC")

    outbound = flight.get("outbound") or {}
    return_fl = flight.get("return_flight") or {}
    is_round = flight.get("type") == "roundtrip"

    total_sell = flight.get("sell_out", 0) + (flight.get("sell_ret", 0) if is_round else 0)
    total_buy = flight.get("buy_out", 0) + (flight.get("buy_ret", 0) if is_round else 0)

    row_data = [
        next_row - 1,
        ref_out,
        f"{passenger['first_name']} {passenger['last_name']}",
        f"{outbound.get('from','-')} → {outbound.get('to','-')}",
        f"{return_fl.get('from','-')} → {return_fl.get('to','-')}" if is_round else "-",
        f"{outbound.get('airline','-')} {outbound.get('flight','')}",
        f"{return_fl.get('airline','-')} {return_fl.get('flight','')}" if is_round else "-",
        flight.get("date_out", "-"),
        flight.get("date_ret", "-") if is_round else "-",
        total_buy,
        total_sell,
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


ALLOWED_USERS = [178875046]

def is_allowed(update) -> bool:
    return update.effective_user.id in ALLOWED_USERS

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("❌ شما دسترسی به این ربات ندارید.")
        return ConversationHandler.END
    
    reply_keyboard = ReplyKeyboardMarkup([
        ["1️⃣ IRANFLY", "2️⃣ FLYTICKET"],
        ["❌ لغو"]
    ], resize_keyboard=True)

    inline_keyboard = [
        [InlineKeyboardButton("1️⃣ IRANFLY", callback_data="company_iranfly")],
        [InlineKeyboardButton("2️⃣ FLYTICKET", callback_data="company_flyticket")],
        [InlineKeyboardButton("❌ لغو", callback_data="cancel_btn")],
    ]
    inline_markup = InlineKeyboardMarkup(inline_keyboard)

    await update.message.reply_text(
        "✈️ *به سیستم صدور بلیت خوش آمدید!*\n\nلطفاً شرکت مورد نظر را انتخاب کنید:",
        parse_mode="Markdown",
        reply_markup=reply_keyboard
    )
    await update.message.reply_text("انتخاب کنید:", reply_markup=inline_markup)
    return ConversationHandler.END

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "company_iranfly":
        context.user_data['company'] = 'iranfly'
        await query.message.reply_text("✅ IRANFLY انتخاب شد.\n\n📸 لطفاً عکس پاسپورت مسافر را ارسال کنید.")
        return WAITING_PASSPORT
    elif query.data == "company_flyticket":
        context.user_data['company'] = 'flyticket'
        await query.message.reply_text("✅ FLYTICKET انتخاب شد.\n\n📸 لطفاً عکس پاسپورت مسافر را ارسال کنید.")
        return WAITING_PASSPORT
    elif query.data == "cancel_btn":
        await query.message.reply_text("❌ لغو شد.")
        return ConversationHandler.END

async def receive_passport(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("❌ شما دسترسی به این ربات ندارید.")
        return ConversationHandler.END
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
            "حالا اطلاعات پرواز را وارد کنید:\n\n"
            "فرمت:\n`ika/tbs/qeshm/6jun/sell14.000/buy12.000`\n"
            "دو طرفه: خط اول رفت، خط دوم برگشت",
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
        pdf_bytes, ref_out, ref_ret = await generate_pdf(passport, flight, TEMPLATE_FILE)

        fname = passport['first_name']
        lname = passenger['last_name'] if 'passenger' in dir() else passport['last_name']
        filename = f"{passport['first_name']}_{passport['last_name']}.pdf"

        await update.message.reply_document(
            document=pdf_bytes,
            filename=filename,
            caption=f"🔖 رفرنس رفت: `{ref_out}`" + (f"\n🔖 رفرنس برگشت: `{ref_ret}`" if flight['type'] == 'roundtrip' else ""),
            parse_mode="Markdown"
        )

        update_excel(passport, flight, ref_out, ref_ret)

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

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CallbackQueryHandler(button_handler))
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
