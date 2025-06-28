import os
import io
import time
import logging
import asyncio
import threading

import fitz                  # PyMuPDF
import pdfplumber
import pandas as pd
from PIL import Image
from flask import Flask, request
from telegram import (
    Update,
    InputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup
)
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters
)
from docx import Document
from pdf2docx import Converter
import requests

# ---------------------- 1. Конфиг ----------------------
TOKEN = os.getenv("TOKEN")
HOST  = os.getenv("RENDER_EXTERNAL_HOSTNAME")
PORT  = int(os.getenv("PORT", "10000"))
if not TOKEN or not HOST:
    raise RuntimeError("TOKEN и RENDER_EXTERNAL_HOSTNAME обязательны")
WEBHOOK_URL = f"https://{HOST}/{TOKEN}"

# ---------------------- 2. Лог ----------------------
logging.basicConfig(format="%(asctime)s %(levelname)s: %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ---------------------- 3. Flask ----------------------
app = Flask(__name__)

# ---------------------- 4. asyncio-loop ----------------------
telegram_loop = asyncio.new_event_loop()

def start_loop(loop):
    asyncio.set_event_loop(loop)
    loop.run_forever()

threading.Thread(target=start_loop, args=(telegram_loop,), daemon=True).start()

# ---------------------- 5. Telegram-app ----------------------
telegram_app = (
    Application.builder()
    .token(TOKEN)
    .connection_pool_size(100)
    .build()
)
telegram_app.request_kwargs = {
    "read_timeout": 60,
    "connect_timeout": 20,
}

# Инициализация PTB
future = asyncio.run_coroutine_threadsafe(telegram_app.initialize(), telegram_loop)
future.result(timeout=15)
logger.info("✔ Telegram initialized")

# ---------------------- 6. PDF-утилиты ----------------------
def upscale_image(img_bytes: bytes, scale: int = 2) -> bytes:
    bio = io.BytesIO(img_bytes)
    im = Image.open(bio)
    w, h = im.size
    im = im.resize((w*scale, h*scale), Image.LANCZOS)
    out = io.BytesIO()
    im.save(out, format="PNG", dpi=(300,300))
    return out.getvalue()

def extract_pdf_elements(path: str):
    doc = fitz.open(path)
    elems = []
    for page in doc:
        txt = page.get_text().strip()
        if txt:
            elems.append(("text", txt))
        for img in page.get_images(full=True):
            raw = doc.extract_image(img[0])["image"]
            hi_res = upscale_image(raw, scale=2)
            elems.append(("img", hi_res))
    doc.close()
    return elems

def save_txt(elems, out: str):
    with open(out, "w", encoding="utf-8") as f:
        for t, c in elems:
            if t == "text":
                f.write(c + "\n\n")

def convert_to_word(elems, out: str):
    docx = Document()
    for t, c in elems:
        if t == "text":
            docx.add_paragraph(c)
        else:
            bio = io.BytesIO(c)
            bio.name = "image.png"
            docx.add_picture(bio)
    docx.save(out)

def convert_pdf_to_docx(pdf_path: str, docx_path: str):
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()

# ---------------------- 7. Клавиатура ----------------------
def make_main_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("Word: макет 📄",           callback_data="cb_word_layout")],
        [InlineKeyboardButton("Word: текст+картинки 📝", callback_data="cb_word_all")],
        [InlineKeyboardButton("TXT: текст 📄",           callback_data="cb_txt")],
        [InlineKeyboardButton("Excel: таблицы 📊",       callback_data="cb_tables")],
        [InlineKeyboardButton("Чат: текст 📝",           callback_data="cb_text_only")],
        [InlineKeyboardButton("Чат: текст+картинки 🖼️📝",callback_data="cb_chat_all")],
        [InlineKeyboardButton("Новый PDF 🔄",            callback_data="cb_new_pdf")],
    ])

# ---------------------- 8. Handlers ----------------------
async def start_handler(u: Update, c: ContextTypes.DEFAULT_TYPE):
    await u.message.reply_text("Привет! Отправь PDF-файл.")

async def handle_pdf(u: Update, c: ContextTypes.DEFAULT_TYPE):
    doc = u.message.document
    if not doc or doc.mime_type != "application/pdf":
        return await u.message.reply_text("Нужен PDF.")
    tgfile = await doc.get_file()
    path = f"/tmp/{doc.file_unique_id}.pdf"
    await tgfile.download_to_drive(path)
    elems = extract_pdf_elements(path)
    c.user_data["pdf_path"] = path
    c.user_data["elements"] = elems
    await u.message.reply_text("Выбери:", reply_markup=make_main_keyboard())

async def cb_word_layout(u: Update, c: ContextTypes.DEFAULT_TYPE):
    await u.callback_query.answer()
    pdf_path = c.user_data.get("pdf_path")
    if not pdf_path:
        return await c.bot.send_message(u.effective_chat.id, "Сначала PDF.")
    out = f"/tmp/{u.effective_user.id}_layout.docx"
    convert_pdf_to_docx(pdf_path, out)
    with open(out, "rb") as f:
        await c.bot.send_document(u.effective_chat.id, InputFile(f, filename="layout.docx"))

async def cb_word_all(u: Update, c: ContextTypes.DEFAULT_TYPE):
    await u.callback_query.answer()
    elems = c.user_data.get("elements", [])
    if not elems:
        return await c.bot.send_message(u.effective_chat.id, "Нет содержимого.")
    out = f"/tmp/{u.effective_user.id}_all.docx"
    convert_to_word(elems, out)
    with open(out, "rb") as f:
        await c.bot.send_document(u.effective_chat.id, InputFile(f, filename="full_converted.docx"))

async def cb_txt(u: Update, c: ContextTypes.DEFAULT_TYPE):
    await u.callback_query.answer()
    elems = c.user_data.get("elements", [])
    if not elems:
        return await c.bot.send_message(u.effective_chat.id, "Нет текста.")
    out = f"/tmp/{u.effective_user.id}.txt"
    save_txt(elems, out)
    with open(out, "rb") as f:
        await c.bot.send_document(u.effective_chat.id, InputFile(f, filename="converted.txt"))

async def cb_tables(u: Update, c: ContextTypes.DEFAULT_TYPE):
    await u.callback_query.answer()
    pdf_path = c.user_data.get("pdf_path")
    if not pdf_path:
        return await c.bot.send_message(u.effective_chat.id, "Сначала PDF.")
    all_tables = []
    with pdfplumber.open(pdf_path) as pdf:
        settings = {
            "vertical_strategy":   "lines",
            "horizontal_strategy": "lines",
            "intersection_tolerance": 5,
        }
        for i, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables(table_settings=settings)
            for idx, table in enumerate(tables, start=1):
                if len(table) < 2: continue
                df = pd.DataFrame(table[1:], columns=table[0])
                all_tables.append((f"Стр{i}_Т{idx}", df))
    if not all_tables:
        return await c.bot.send_message(u.effective_chat.id, "Таблиц нет.")
    excel = f"/tmp/{u.effective_user.id}_tables.xlsx"
    with pd.ExcelWriter(excel, engine="openpyxl") as writer:
        for name, df in all_tables:
            df.to_excel(writer, sheet_name=name[:31], index=False)
    with open(excel, "rb") as f:
        await c.bot.send_document(u.effective_chat.id, InputFile(f, filename="tables.xlsx"))

async def cb_text_only(u: Update, c: ContextTypes.DEFAULT_TYPE):
    await u.callback_query.answer()
    elems = c.user_data.get("elements", [])
    if not elems:
        return await c.bot.send_message(u.effective_chat.id, "Нет текста.")
    for t, blk in elems:
        if t == "text":
            for i in range(0, len(blk), 4096):
                await c.bot.send_message(u.effective_chat.id, blk[i:i+4096])
                await asyncio.sleep(0.1)

async def cb_chat_all(u: Update, c: ContextTypes.DEFAULT_TYPE):
    await u.callback_query.answer()
    elems = c.user_data.get("elements", [])
    if not elems:
        return await c.bot.send_message(u.effective_chat.id, "Нет содержимого.")
    sent = set()
    for t, content in elems:
        if t == "text":
            for i in range(0, len(content), 4096):
                await c.bot.send_message(u.effective_chat.id, content[i:i+4096])
                await asyncio.sleep(0.1)
        else:
            h = hash(content)
            if h in sent: continue
            sent.add(h)
            bio = io.BytesIO(content)
            bio.name = "image.png"
            await c.bot.send_photo(u.effective_chat.id, photo=bio)
            await asyncio.sleep(0.1)

async def cb_new_pdf(u: Update, c: ContextTypes.DEFAULT_TYPE):
    await u.callback_query.answer()
    c.user_data.clear()
    await c.bot.send_message(u.effective_chat.id, "Отправь новый PDF.")

# ---------------------- 9. Регистрация ----------------------
telegram_app.add_handler(CommandHandler("start", start_handler))
telegram_app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
telegram_app.add_handler(CallbackQueryHandler(cb_word_layout, pattern="cb_word_layout"))
telegram_app.add_handler(CallbackQueryHandler(cb_word_all,    pattern="cb_word_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_txt,         pattern="cb_txt"))
telegram_app.add_handler(CallbackQueryHandler(cb_tables,      pattern="cb_tables"))
telegram_app.add_handler(CallbackQueryHandler(cb_text_only,   pattern="cb_text_only"))
telegram_app.add_handler(CallbackQueryHandler(cb_chat_all,    pattern="cb_chat_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_new_pdf,     pattern="cb_new_pdf"))

# ---------------------- 10. Flask-routes ----------------------
@app.route(f"/{TOKEN}", methods=["POST"])
def webhook():
    data = request.get_json(force=True)
    upd  = Update.de_json(data, telegram_app.bot)
    asyncio.run_coroutine_threadsafe(telegram_app.process_update(upd), telegram_loop)
    return "OK"

@app.route("/ping")
def ping():
    return "pong"

# ---------------------- 11. Запуск ----------------------
if __name__ == "__main__":
    logger.info(f"Setting webhook to {WEBHOOK_URL}")
    r = requests.post(
        f"https://api.telegram.org/bot{TOKEN}/setWebhook",
        data={"url": WEBHOOK_URL}
    )
    logger.info("Webhook set" if r.ok else f"Webhook error: {r.text}")
    logger.info("Run Flask on port %s", PORT)
    app.run(host="0.0.0.0", port=PORT)
