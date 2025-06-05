import os
import io
import logging
import asyncio
import fitz  # PyMuPDF
import pdfplumber
import pandas as pd
import openpyxl
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

TOKEN = os.getenv("TOKEN")
HOST = os.getenv("RENDER_EXTERNAL_HOSTNAME")
PORT = int(os.getenv("PORT", "10000"))
WEBHOOK_URL = f"https://{HOST}/{TOKEN}"

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app_flask = Flask(__name__)

telegram_app = (
    Application.builder()
    .token(TOKEN)
    .connection_pool_size(100)
    .build()
)

# =============== PDF UTILS ====================

def extract_pdf_elements(path):
    doc = fitz.open(path)
    elements = []
    for page in doc:
        text = page.get_text().strip()
        if text:
            elements.append(("text", text))
        for img in page.get_images(full=True):
            xref = img[0]
            data = doc.extract_image(xref)["image"]
            elements.append(("img", data))
    doc.close()
    return elements

def pdf_tables_to_excel(path, out_path):
    tables = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            t = page.extract_table()
            if t:
                tables.append(t)
    if not tables:
        return False
    # Сохраняем первую таблицу в Excel
    df = pd.DataFrame(tables[0])
    df.to_excel(out_path, index=False)
    return True

def save_elements_to_word(elements, out_path):
    docx = Document()
    for typ, content in elements:
        if typ == "text":
            docx.add_paragraph(content)
        else:
            bio = io.BytesIO(content)
            bio.name = "image.png"
            docx.add_picture(bio)
    docx.save(out_path)

def save_elements_to_txt(elements, out_path):
    with open(out_path, "w", encoding="utf-8") as f:
        for typ, content in elements:
            if typ == "text":
                f.write(content + "\n\n")

# =============== BUTTONS ======================

def make_main_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("Word: текст+картинки 📝", callback_data="word_all")],
        [InlineKeyboardButton("TXT: текст 📄", callback_data="txt_text")],
        [InlineKeyboardButton("Excel: таблицы 📊", callback_data="excel_tables")],
        [InlineKeyboardButton("Чат: текст 📝", callback_data="chat_text")],
        [InlineKeyboardButton("Чат: текст+картинки 🖼️📝", callback_data="chat_all")],
        [InlineKeyboardButton("Новый PDF 🔄", callback_data="new_pdf")]
    ])

# ============= HANDLERS =======================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Привет! Отправь PDF-файл.")

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if doc.mime_type != "application/pdf":
        return await update.message.reply_text("Пожалуйста, отправьте PDF.")
    file = await doc.get_file()
    path = f"/tmp/{doc.file_unique_id}.pdf"
    await file.download_to_drive(path)
    context.user_data['pdf_path'] = path
    context.user_data['elements'] = extract_pdf_elements(path)
    await update.message.reply_text(
        "Выбери, что сделать с этим PDF:",
        reply_markup=make_main_keyboard()
    )

async def cb_word_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    elements = context.user_data.get('elements')
    if not elements:
        return await update.callback_query.edit_message_text("Нет данных для Word.")
    out = f"/tmp/{update.effective_user.id}.docx"
    save_elements_to_word(elements, out)
    with open(out, "rb") as f:
        await context.bot.send_document(update.effective_chat.id, InputFile(f, filename="converted.docx"))
    await update.callback_query.edit_message_text("Готово! 📄", reply_markup=make_main_keyboard())

async def cb_txt_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    elements = context.user_data.get('elements')
    if not elements:
        return await update.callback_query.edit_message_text("Нет данных для TXT.")
    out = f"/tmp/{update.effective_user.id}.txt"
    save_elements_to_txt(elements, out)
    with open(out, "rb") as f:
        await context.bot.send_document(update.effective_chat.id, InputFile(f, filename="converted.txt"))
    await update.callback_query.edit_message_text("Готово! 📄", reply_markup=make_main_keyboard())

async def cb_excel_tables(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    pdf_path = context.user_data.get('pdf_path')
    out = f"/tmp/{update.effective_user.id}.xlsx"
    ok = pdf_tables_to_excel(pdf_path, out)
    if ok:
        with open(out, "rb") as f:
            await context.bot.send_document(update.effective_chat.id, InputFile(f, filename="tables.xlsx"))
        await update.callback_query.edit_message_text("Таблицы Excel готовы! 📊", reply_markup=make_main_keyboard())
    else:
        await update.callback_query.edit_message_text("В PDF нет распознаваемых таблиц.", reply_markup=make_main_keyboard())

async def cb_chat_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    elements = context.user_data.get('elements')
    if not elements:
        return await update.callback_query.edit_message_text("Нет данных.")
    for typ, content in elements:
        if typ == "text":
            for i in range(0, len(content), 4096):
                await context.bot.send_message(update.effective_chat.id, content[i:i+4096])
    await update.callback_query.edit_message_text("Готово! 📝", reply_markup=make_main_keyboard())

async def cb_chat_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    elements = context.user_data.get('elements')
    if not elements:
        return await update.callback_query.edit_message_text("Нет данных.")
    sent = set()
    for typ, content in elements:
        if typ == "text":
            for i in range(0, len(content), 4096):
                await context.bot.send_message(update.effective_chat.id, content[i:i+4096])
        else:
            h = hash(content)
            if h in sent:
                continue
            sent.add(h)
            bio = io.BytesIO(content)
            bio.name = "image.png"
            await context.bot.send_photo(update.effective_chat.id, photo=bio)
    await update.callback_query.edit_message_text("Готово! 🖼️📝", reply_markup=make_main_keyboard())

async def cb_new_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    context.user_data.clear()
    await context.bot.send_message(update.effective_chat.id, "Отправьте новый PDF-файл.")

telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
telegram_app.add_handler(CallbackQueryHandler(cb_word_all, pattern="word_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_txt_text, pattern="txt_text"))
telegram_app.add_handler(CallbackQueryHandler(cb_excel_tables, pattern="excel_tables"))
telegram_app.add_handler(CallbackQueryHandler(cb_chat_text, pattern="chat_text"))
telegram_app.add_handler(CallbackQueryHandler(cb_chat_all, pattern="chat_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_new_pdf, pattern="new_pdf"))

# =========== Flask + Webhook =============

loop = asyncio.get_event_loop()

async def init_telegram():
    await telegram_app.initialize()

loop.run_until_complete(init_telegram())

@app_flask.route(f"/{TOKEN}", methods=["POST"])
def webhook():
    update = Update.de_json(request.get_json(force=True), telegram_app.bot)
    # НЕ ЖДЁМ ОТВЕТА, сразу отдаём "ok"!
    asyncio.run_coroutine_threadsafe(telegram_app.process_update(update), loop)
    return "ok"

@app_flask.route("/ping")
def ping():
    return "pong"

if __name__ == "__main__":
    import requests
    requests.post(
        f"https://api.telegram.org/bot{TOKEN}/setWebhook",
        data={"url": WEBHOOK_URL}
    )
    logger.info("Запускаем Flask на порту %s", PORT)
    app_flask.run(host="0.0.0.0", port=PORT)
