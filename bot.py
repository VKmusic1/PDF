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

# Конфигурация
TOKEN = os.getenv("TOKEN")
if not TOKEN:
    raise RuntimeError("Environment variable TOKEN is required")
HOST = os.getenv("RENDER_EXTERNAL_HOSTNAME")
if not HOST:
    raise RuntimeError("Environment variable RENDER_EXTERNAL_HOSTNAME is required")
PORT = int(os.getenv("PORT", "10000"))
WEBHOOK_URL = f"https://{HOST}/{TOKEN}"

logging.basicConfig(format="%(asctime)s %(levelname)s: %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

telegram_app = (
    Application.builder()
    .token(TOKEN)
    .connection_pool_size(100)
    .build()
)

# ================== PDF EXTRACTORS ====================

def extract_pdf_elements(path: str):
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

def extract_text_only(elements):
    return [c for t, c in elements if t == "text"]

def extract_tables(path: str):
    tables = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            page_tables = page.extract_tables()
            for table in page_tables:
                if table:
                    tables.append(table)
    return tables

def save_tables_to_excel(tables, out_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    for table in tables:
        for row in table:
            ws.append(row)
        ws.append([])  # пустая строка между таблицами
    wb.save(out_path)

def elements_to_word(elements, out_path: str):
    docx = Document()
    for typ, content in elements:
        if typ == "text":
            docx.add_paragraph(content)
        elif typ == "img":
            bio = io.BytesIO(content)
            bio.name = "image.png"
            docx.add_picture(bio)
    docx.save(out_path)

def elements_to_txt(elements, out_path: str):
    with open(out_path, "w", encoding="utf-8") as f:
        for typ, content in elements:
            if typ == "text":
                f.write(content + "\n\n")

# ================== HANDLERS =========================

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
    context.user_data['tables'] = extract_tables(path)
    # Кнопки меню
    keyboard = [
        [InlineKeyboardButton("Word: текст+картинки 📝", callback_data="word_all")],
        [InlineKeyboardButton("TXT: текст 📄", callback_data="txt_text")],
        [InlineKeyboardButton("Excel: таблицы 📊", callback_data="excel_tables")],
        [InlineKeyboardButton("Чат: текст 📝", callback_data="chat_text")],
        [InlineKeyboardButton("Чат: текст+картинки 🖼️📝", callback_data="chat_all")],
        [InlineKeyboardButton("Новый PDF ♻️", callback_data="new_pdf")],
    ]
    await update.message.reply_text(
        "Выбери вариант работы с этим PDF:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def cb_word_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    elements = context.user_data.get("elements")
    if not elements:
        return await update.callback_query.edit_message_text("Файл не найден.")
    out = f"/tmp/{update.effective_user.id}.docx"
    elements_to_word(elements, out)
    with open(out, "rb") as f:
        await context.bot.send_document(update.effective_chat.id, InputFile(f, filename="converted.docx"))

async def cb_txt_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    elements = context.user_data.get("elements")
    if not elements:
        return await update.callback_query.edit_message_text("Файл не найден.")
    out = f"/tmp/{update.effective_user.id}.txt"
    elements_to_txt(elements, out)
    with open(out, "rb") as f:
        await context.bot.send_document(update.effective_chat.id, InputFile(f, filename="converted.txt"))

async def cb_excel_tables(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    tables = context.user_data.get("tables")
    if not tables or all(not t for t in tables):
        await update.callback_query.message.reply_text("В PDF нет распознаваемых таблиц.")
        return
    out = f"/tmp/{update.effective_user.id}.xlsx"
    save_tables_to_excel(tables, out)
    with open(out, "rb") as f:
        await context.bot.send_document(update.effective_chat.id, InputFile(f, filename="tables.xlsx"))

async def cb_chat_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    elements = context.user_data.get("elements")
    if not elements:
        return await update.callback_query.edit_message_text("Файл не найден.")
    text_only = [c for t, c in elements if t == "text"]
    if not text_only:
        return await update.callback_query.message.reply_text("В PDF нет текста.")
    for chunk in text_only:
        await context.bot.send_message(update.effective_chat.id, chunk[:4096])

async def cb_chat_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    elements = context.user_data.get("elements")
    if not elements:
        return await update.callback_query.edit_message_text("Файл не найден.")
    sent_imgs = set()
    for typ, content in elements:
        if typ == "text":
            for i in range(0, len(content), 4096):
                await context.bot.send_message(update.effective_chat.id, content[i:i+4096])
        elif typ == "img":
            h = hash(content)
            if h in sent_imgs:
                continue
            sent_imgs.add(h)
            bio = io.BytesIO(content)
            bio.name = "image.png"
            await context.bot.send_photo(update.effective_chat.id, photo=bio)

async def cb_new_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    context.user_data.clear()
    await context.bot.send_message(update.effective_chat.id, "Отправьте новый PDF-файл.")

# ================ HANDLERS REGISTRATION ===============

telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
telegram_app.add_handler(CallbackQueryHandler(cb_word_all, pattern="word_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_txt_text, pattern="txt_text"))
telegram_app.add_handler(CallbackQueryHandler(cb_excel_tables, pattern="excel_tables"))
telegram_app.add_handler(CallbackQueryHandler(cb_chat_text, pattern="chat_text"))
telegram_app.add_handler(CallbackQueryHandler(cb_chat_all, pattern="chat_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_new_pdf, pattern="new_pdf"))

# ================ FLASK ===============================

app_flask = Flask(__name__)

@app_flask.route(f"/{TOKEN}", methods=["POST"])
def webhook():
    update = Update.de_json(request.get_json(force=True), telegram_app.bot)
    asyncio.run(telegram_app.process_update(update))
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
    logging.info("Запускаем Flask на порту %s", PORT)
    app_flask.run(host="0.0.0.0", port=PORT)
