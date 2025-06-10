import os
import io
import logging
import asyncio
import threading

import fitz                  # PyMuPDF
import pdfplumber
import pandas as pd
from flask import Flask, request
from telegram import Update, InputFile, InlineKeyboardButton, InlineKeyboardMarkup
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

# ======================= 1. Конфигурация =======================

TOKEN = os.getenv("TOKEN")
if not TOKEN:
    raise RuntimeError("Environment variable TOKEN is required")

HOST = os.getenv("RENDER_EXTERNAL_HOSTNAME")
if not HOST:
    raise RuntimeError("Environment variable RENDER_EXTERNAL_HOSTNAME is required")

PORT = int(os.getenv("PORT", "10000"))
WEBHOOK_URL = f"https://{HOST}/{TOKEN}"

# ======================= 2. Логирование =======================

logging.basicConfig(format="%(asctime)s %(levelname)s: %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ======================= 3. Flask =======================

app = Flask(__name__)

# ======================= 4. asyncio-loop в фоне =======================

telegram_loop = asyncio.new_event_loop()

def start_loop(loop):
    asyncio.set_event_loop(loop)
    loop.run_forever()

threading.Thread(target=start_loop, args=(telegram_loop,), daemon=True).start()

# ======================= 5. Telegram Application =======================

telegram_app = (
    Application.builder()
    .token(TOKEN)
    .connection_pool_size(100)
    .build()
)
telegram_app.request_kwargs = {
    "read_timeout": 60,
    "connect_timeout": 20
}

# инициализация
future = asyncio.run_coroutine_threadsafe(telegram_app.initialize(), telegram_loop)
future.result(timeout=15)
logger.info("✔ Telegram Application initialized")

# ======================= 6. PDF‐утилиты =======================

def extract_pdf_elements(path: str):
    doc = fitz.open(path)
    elements = []
    for page in doc:
        text = page.get_text().strip()
        if text:
            elements.append(("text", text))
        for img in page.get_images(full=True):
            data = doc.extract_image(img[0])["image"]
            elements.append(("img", data))
    doc.close()
    return elements

def save_text_to_txt(elements, out_path: str):
    with open(out_path, "w", encoding="utf-8") as f:
        for typ, content in elements:
            if typ == "text":
                f.write(content + "\n\n")

def convert_pdf_to_docx(pdf_path: str, docx_path: str):
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()

# ======================= 7. Handlers =======================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Привет! Отправь PDF-файл, и я предложу варианты извлечения ↓")

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    logger.info("Получен документ от %s", update.effective_user.id)
    doc = update.message.document
    if not doc or doc.mime_type != "application/pdf":
        return await update.message.reply_text("Пожалуйста, отправь PDF-файл.")
    tgfile = await doc.get_file()
    local_path = f"/tmp/{doc.file_unique_id}.pdf"
    await tgfile.download_to_drive(local_path)
    elements = extract_pdf_elements(local_path)
    context.user_data["pdf_path"] = local_path
    context.user_data["elements"] = elements
    keyboard = [
        [InlineKeyboardButton("Word: текст+картинки 📄", callback_data="cb_word_all")],
        [InlineKeyboardButton("TXT: текст 📄",           callback_data="cb_txt")],
        [InlineKeyboardButton("Excel: таблицы 📊",        callback_data="cb_tables")],
        [InlineKeyboardButton("Чат: текст 📝",           callback_data="cb_text_only")],
        [InlineKeyboardButton("Чат: текст+картинки 🖼️📝", callback_data="cb_chat_all")],
        [InlineKeyboardButton("Новый PDF 🔄",             callback_data="cb_new_pdf")],
    ]
    await update.message.reply_text(
        "Выбери, что сделать с этим PDF:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def cb_text_only(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await update.callback_query.edit_message_text("Файл не найден. Пришли новый PDF.")
    blocks = [c for t,c in extract_pdf_elements(path) if t=="text"]
    if not blocks:
        return await context.bot.send_message(update.effective_chat.id, "В PDF нет текста.")
    for block in blocks:
        for i in range(0, len(block), 4096):
            await context.bot.send_message(update.effective_chat.id, block[i:i+4096])
            await asyncio.sleep(0.05)

async def cb_chat_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await update.callback_query.edit_message_text("Файл не найден. Пришли новый PDF.")
    sent = set()
    for typ,content in extract_pdf_elements(path):
        if typ=="text":
            for i in range(0, len(content), 4096):
                await context.bot.send_message(update.effective_chat.id, content[i:i+4096])
                await asyncio.sleep(0.05)
        else:
            h=hash(content)
            if h in sent: continue
            sent.add(h)
            bio=io.BytesIO(content); bio.name="image.png"
            await context.bot.send_photo(update.effective_chat.id, photo=bio)
            await asyncio.sleep(0.1)

async def cb_word_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await context.bot.send_message(update.effective_chat.id, "Файл не найден.")
    out = f"/tmp/{update.effective_user.id}_layout.docx"
    convert_pdf_to_docx(path, out)
    with open(out,"rb") as f:
        await context.bot.send_document(update.effective_chat.id, InputFile(f, filename="converted_layout.docx"))

async def cb_txt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    elements = context.user_data.get("elements")
    if not elements:
        return await context.bot.send_message(update.effective_chat.id, "Файл не найден.")
    out = f"/tmp/{update.effective_user.id}.txt"
    save_text_to_txt(elements, out)
    with open(out,"rb") as f:
        await context.bot.send_document(update.effective_chat.id, InputFile(f, filename="converted.txt"))

async def cb_tables(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await context.bot.send_message(update.effective_chat.id, "Файл не найден.")
    tables=[]
    with pdfplumber.open(path) as pdf:
        for pn,page in enumerate(pdf.pages,1):
            for ti,table in enumerate(page.extract_tables(),1):
                if not table or len(table)<2: continue
                df=pd.DataFrame(table[1:], columns=table[0])
                tables.append((f"S{pn}T{ti}"[:31], df))
    if not tables:
        return await context.bot.send_message(update.effective_chat.id, "Таблиц не найдено.")
    excel = f"/tmp/{update.effective_user.id}_tables.xlsx"
    with pd.ExcelWriter(excel, engine="openpyxl") as w:
        for name,df in tables:
            df.to_excel(w, sheet_name=name, index=False)
    with open(excel,"rb") as f:
        await context.bot.send_document(update.effective_chat.id, InputFile(f, filename="tables.xlsx"))

async def cb_new_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    context.user_data.clear()
    await context.bot.send_message(update.effective_chat.id, "Пришлите новый PDF-файл →")

# регистрируем
telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
telegram_app.add_handler(CallbackQueryHandler(cb_text_only, pattern="cb_text_only"))
telegram_app.add_handler(CallbackQueryHandler(cb_chat_all, pattern="cb_chat_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_word_all, pattern="cb_word_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_txt, pattern="cb_txt"))
telegram_app.add_handler(CallbackQueryHandler(cb_tables, pattern="cb_tables"))
telegram_app.add_handler(CallbackQueryHandler(cb_new_pdf, pattern="cb_new_pdf"))

# ======================= 8. Flask-маршруты =======================

@app.route(f"/{TOKEN}", methods=["POST"])
def telegram_webhook():
    data = request.get_json(force=True)
    update = Update.de_json(data, telegram_app.bot)
    asyncio.run_coroutine_threadsafe(
        telegram_app.process_update(update),
        telegram_loop
    )
    return "OK"

@app.route("/ping", methods=["GET"])
def ping():
    return "pong"

# ======================= 9. Запуск =======================

if __name__ == "__main__":
    logger.info(f"Setting webhook to {WEBHOOK_URL}")
    r = requests.post(
        f"https://api.telegram.org/bot{TOKEN}/setWebhook",
        data={"url": WEBHOOK_URL}
    )
    if r.ok:
        logger.info("Webhook установлен успешно")
    else:
        logger.error("Ошибка установки webhook: %s", r.text)

    logger.info("Запускаем Flask на порту %s", PORT)
    app.run(host="0.0.0.0", port=PORT)
