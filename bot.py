import os
import io
import logging
import fitz                  # PyMuPDF
import pdfplumber
import pandas as pd
from pdf2docx import Converter
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
from flask import Flask, request
import asyncio

# ---------------------- 1. Конфигурация из окружения ----------------------
TOKEN = os.getenv("TOKEN")
if not TOKEN:
    raise RuntimeError("Environment variable TOKEN is required")

HOST = os.getenv("RENDER_EXTERNAL_HOSTNAME")
if not HOST:
    raise RuntimeError("Environment variable RENDER_EXTERNAL_HOSTNAME is required")

PORT = int(os.getenv("PORT", "10000"))
WEBHOOK_URL = f"https://{HOST}/{TOKEN}"

# ---------------------- 2. Логирование ----------------------
logging.basicConfig(format="%(asctime)s %(levelname)s: %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ---------------------- 3. Инициализация Flask & Telegram ----------------------
app = Flask(__name__)
telegram_app = (
    Application.builder()
    .token(TOKEN)
    .connection_pool_size(100)
    .build()
)

# ---------------------- 4. PDF-утилиты ----------------------

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

def pdf_tables_to_excel(path: str, out_path: str):
    tables = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            t = page.extract_table()
            if t:
                tables.append(t)
    if not tables:
        return False
    df = pd.DataFrame(tables[0])
    df.to_excel(out_path, index=False)
    return True

def save_elements_to_txt(elements, out_path: str):
    with open(out_path, "w", encoding="utf-8") as f:
        for typ, content in elements:
            if typ == "text":
                f.write(content + "\n\n")

# ---------------------- 5. Клавиатура ----------------------

def make_main_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("Word: макет 📄",     callback_data="cb_word_layout")],
        [InlineKeyboardButton("Word: текст+картинки 📝", callback_data="cb_word_all")],
        [InlineKeyboardButton("TXT: текст 📄",      callback_data="cb_txt")],
        [InlineKeyboardButton("Excel: таблицы 📊",  callback_data="cb_tables")],
        [InlineKeyboardButton("Чат: текст 📝",      callback_data="cb_text_only")],
        [InlineKeyboardButton("Чат: текст+картинки 🖼️📝", callback_data="cb_chat_all")],
        [InlineKeyboardButton("Новый PDF 🔄",       callback_data="cb_new_pdf")],
    ])

# ---------------------- 6. Хендлеры ----------------------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Привет! Отправь PDF-файл.")

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc or doc.mime_type != "application/pdf":
        return await update.message.reply_text("Пожалуйста, отправь PDF.")
    file = await doc.get_file()
    path = f"/tmp/{doc.file_unique_id}.pdf"
    await file.download_to_drive(path)
    context.user_data["pdf_path"] = path
    await update.message.reply_text(
        "Выбери вариант работы с этим PDF:",
        reply_markup=make_main_keyboard()
    )

# ——— Новый хендлер: «макет» через pdf2docx ———

async def cb_word_layout(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    user = update.effective_user.id
    pdf_path = context.user_data.get("pdf_path")
    if not pdf_path:
        return await context.bot.send_message(update.effective_chat.id, "Сначала отправь PDF.")
    out_path = f"/tmp/{user}_layout.docx"
    logger.info("Layout conversion %s → %s", pdf_path, out_path)

    # Конвертация «макета»
    converter = Converter(pdf_path)
    converter.convert(out_path, start=0, end=None)
    converter.close()

    # Отправка
    with open(out_path, "rb") as f:
        await context.bot.send_document(
            update.effective_chat.id,
            document=InputFile(f, filename="layout_converted.docx")
        )

# ——— Остальные коллбеки (Word All, TXT, Excel, Чат) ———

async def cb_word_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    elems = extract_pdf_elements(context.user_data.get("pdf_path", ""))
    out = f"/tmp/{update.effective_user.id}_all.docx"
    convert_to_word(elems, out)
    with open(out, "rb") as f:
        await context.bot.send_document(update.effective_chat.id, document=InputFile(f, "all_converted.docx"))

async def cb_txt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    elems = extract_pdf_elements(context.user_data.get("pdf_path", ""))
    out = f"/tmp/{update.effective_user.id}.txt"
    save_elements_to_txt(elems, out)
    with open(out, "rb") as f:
        await context.bot.send_document(update.effective_chat.id, document=InputFile(f, "converted.txt"))

async def cb_tables(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    pdf_path = context.user_data.get("pdf_path")
    out = f"/tmp/{update.effective_user.id}.xlsx"
    if pdf_tables_to_excel(pdf_path, out):
        with open(out, "rb") as f:
            await context.bot.send_document(update.effective_chat.id, document=InputFile(f, "tables.xlsx"))
    else:
        await context.bot.send_message(update.effective_chat.id, "В PDF нет таблиц.", reply_markup=make_main_keyboard())

async def cb_text_only(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    elems = extract_pdf_elements(context.user_data.get("pdf_path", ""))
    for typ, content in elems:
        if typ == "text":
            for i in range(0, len(content), 4096):
                await context.bot.send_message(update.effective_chat.id, content[i:i+4096])

async def cb_chat_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    sent = set()
    elems = extract_pdf_elements(context.user_data.get("pdf_path", ""))
    for typ, content in elems:
        if typ == "text":
            for i in range(0, len(content), 4096):
                await context.bot.send_message(update.effective_chat.id, content[i:i+4096])
        else:
            h = hash(content)
            if h in sent:
                continue
            sent.add(h)
            bio = io.BytesIO(content); bio.name = "image.png"
            await context.bot.send_photo(update.effective_chat.id, photo=bio)

async def cb_new_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    context.user_data.clear()
    await context.bot.send_message(update.effective_chat.id, "Отправь новый PDF.")

# ——— Регистрация хендлеров ———

telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
telegram_app.add_handler(CallbackQueryHandler(cb_word_layout, pattern="cb_word_layout"))
telegram_app.add_handler(CallbackQueryHandler(cb_word_all,    pattern="cb_word_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_txt,         pattern="cb_txt"))
telegram_app.add_handler(CallbackQueryHandler(cb_tables,      pattern="cb_tables"))
telegram_app.add_handler(CallbackQueryHandler(cb_text_only,   pattern="cb_text_only"))
telegram_app.add_handler(CallbackQueryHandler(cb_chat_all,    pattern="cb_chat_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_new_pdf,     pattern="cb_new_pdf"))

# ---------------------- 7. Ping-pong & Webhook ----------------------

@app.route("/ping")
def ping():
    return "pong"

@app.route(f"/{TOKEN}", methods=["POST"])
def webhook():
    update = Update.de_json(request.get_json(force=True), telegram_app.bot)
    # ставим в текущий loop задачу на обработку, не блокируем Flask
    asyncio.get_event_loop().create_task(telegram_app.process_update(update))
    return "ok"

if __name__ == "__main__":
    logger.info("Устанавливаем webhook %s", WEBHOOK_URL)
    telegram_app.bot.set_webhook(WEBHOOK_URL)
    logger.info("Запускаем Flask на порту %s", PORT)
    app.run(host="0.0.0.0", port=PORT)
