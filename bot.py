import os
import io
import logging
import fitz
import pdfplumber
import pandas as pd
from telegram import (
    Update, InputFile, InlineKeyboardButton, InlineKeyboardMarkup
)
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters
)
from docx import Document
from flask import Flask, request
import asyncio

# ---- Конфиг ----
TOKEN = os.getenv("TOKEN")
HOST = os.getenv("RENDER_EXTERNAL_HOSTNAME")
PORT = int(os.getenv("PORT", "10000"))
WEBHOOK_URL = f"https://{HOST}/{TOKEN}"

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# ---- Telegram App ----
telegram_app = Application.builder().token(TOKEN).build()

# ---- PDF utils ----
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

def convert_to_word(elements, out_path: str):
    docx = Document()
    for typ, content in elements:
        if typ == "text":
            docx.add_paragraph(content)
        else:
            bio = io.BytesIO(content)
            bio.name = "image.png"
            docx.add_picture(bio)
    docx.save(out_path)

# ---- Handlers ----
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Привет! Отправь PDF-файл, и я предложу варианты извлечения.")

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc or doc.mime_type != "application/pdf":
        return await update.message.reply_text("Пожалуйста, отправь PDF-файл.")
    file = await doc.get_file()
    path = f"/tmp/{doc.file_unique_id}.pdf"
    await file.download_to_drive(path)
    context.user_data["pdf_path"] = path

    keyboard = [
        [InlineKeyboardButton("Word: текст+картинки 📄", callback_data="cb_word_all")],
        [InlineKeyboardButton("TXT: текст 📄",            callback_data="cb_txt")],
        [InlineKeyboardButton("Excel: таблицы 📊",         callback_data="cb_tables")],
        [InlineKeyboardButton("Чат: текст 📝",            callback_data="cb_text_only")],
        [InlineKeyboardButton("Чат: текст+картинки 🖼️📝", callback_data="cb_chat_all")],
        [InlineKeyboardButton("Новый PDF 🔄",             callback_data="cb_new_pdf")],
    ]
    await update.message.reply_text(
        "Выбери вариант работы с этим PDF:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def cb_text_only(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await context.bot.send_message(update.effective_chat.id, "Файл не найден. Отправь новый PDF.")
    elements = extract_pdf_elements(path)
    text_only = [c for t, c in elements if t == "text"]
    if not text_only:
        return await context.bot.send_message(update.effective_chat.id, "В PDF нет текста.")
    for block in text_only:
        for i in range(0, len(block), 4096):
            await context.bot.send_message(update.effective_chat.id, block[i:i+4096])
            await asyncio.sleep(0.1)

async def cb_chat_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await context.bot.send_message(update.effective_chat.id, "Файл не найден. Отправь новый PDF.")
    elements = extract_pdf_elements(path)
    if not elements:
        return await context.bot.send_message(update.effective_chat.id, "В PDF нет контента.")
    sent = set()
    for typ, content in elements:
        if typ == "text":
            for i in range(0, len(content), 4096):
                await context.bot.send_message(update.effective_chat.id, content[i:i+4096])
                await asyncio.sleep(0.1)
        else:
            h = hash(content)
            if h in sent:
                continue
            sent.add(h)
            bio = io.BytesIO(content)
            bio.name = "image.png"
            await context.bot.send_photo(update.effective_chat.id, photo=bio)
            await asyncio.sleep(0.1)

async def cb_word_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await context.bot.send_message(update.effective_chat.id, "Файл не найден. Отправь новый PDF.")
    elements = extract_pdf_elements(path)
    if not elements:
        return await context.bot.send_message(update.effective_chat.id, "В PDF нет контента.")
    out = f"/tmp/{update.effective_user.id}_all.docx"
    convert_to_word(elements, out)
    with open(out, "rb") as f:
        await context.bot.send_document(
            update.effective_chat.id,
            document=InputFile(f, filename="full_converted.docx")
        )

async def cb_txt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await context.bot.send_message(update.effective_chat.id, "Файл не найден. Отправь новый PDF.")
    elements = extract_pdf_elements(path)
    text_only = [c for t, c in elements if t == "text"]
    if not text_only:
        return await context.bot.send_message(update.effective_chat.id, "В PDF нет текста.")
    all_text = "\n\n".join(text_only)
    out_path = f"/tmp/{update.effective_user.id}.txt"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(all_text)
    with open(out_path, "rb") as f:
        await context.bot.send_document(
            update.effective_chat.id,
            document=InputFile(f, filename="full_converted.txt")
        )

async def cb_tables(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await context.bot.send_message(update.effective_chat.id, "Файл не найден. Отправь новый PDF.")
    all_tables = []
    with pdfplumber.open(path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            for tbl_idx, table in enumerate(tables, start=1):
                if not table or len(table) < 2:
                    continue
                df = pd.DataFrame(table[1:], columns=table[0])
                sheet_name = f"Стр{page_number}_Таб{tbl_idx}"
                all_tables.append((sheet_name, df))
    if not all_tables:
        return await context.bot.send_message(update.effective_chat.id, "В PDF нет распознаваемых таблиц.")
    excel_path = f"/tmp/{update.effective_user.id}_tables.xlsx"
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        for sheet_name, df in all_tables:
            safe_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)
    with open(excel_path, "rb") as f:
        await context.bot.send_document(
            update.effective_chat.id,
            document=InputFile(f, filename="tables.xlsx")
        )

async def cb_new_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    context.user_data.clear()
    await context.bot.send_message(update.effective_chat.id, "Отправьте новый PDF-файл.")

# ---- Регистрация ----
telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
telegram_app.add_handler(CallbackQueryHandler(cb_text_only,   pattern="cb_text_only"))
telegram_app.add_handler(CallbackQueryHandler(cb_chat_all,    pattern="cb_chat_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_word_all,    pattern="cb_word_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_txt,         pattern="cb_txt"))
telegram_app.add_handler(CallbackQueryHandler(cb_tables,      pattern="cb_tables"))
telegram_app.add_handler(CallbackQueryHandler(cb_new_pdf,     pattern="cb_new_pdf"))

# ---- Flask endpoints ----
@app.route("/ping")
def ping():
    return "pong"

@app.route(f"/{TOKEN}", methods=["POST"])
def telegram_webhook():
    update = Update.de_json(request.get_json(force=True), telegram_app.bot)
    try:
        loop = asyncio.get_event_loop()
    except RuntimeError:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
    loop.run_until_complete(telegram_app.process_update(update))
    return "ok"

# ---- Запуск ----
if __name__ == "__main__":
    import requests
    # Установка вебхука
    requests.post(
        f"https://api.telegram.org/bot{TOKEN}/setWebhook",
        data={"url": WEBHOOK_URL}
    )
    logger.info(f"Запускаем Flask на порту {PORT}")
    app.run(host="0.0.0.0", port=PORT)
