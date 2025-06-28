import os
import io
import logging
import asyncio
import fitz  # PyMuPDF
import pdfplumber
import pandas as pd
from pdf2docx import Converter
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

# ---------------------- Configuration ----------------------
TOKEN = os.getenv("TOKEN")
HOST = os.getenv("RENDER_EXTERNAL_HOSTNAME")
PORT = int(os.getenv("PORT", "10000"))
if not TOKEN or not HOST:
    raise RuntimeError("Environment variables TOKEN and RENDER_EXTERNAL_HOSTNAME are required")
WEBHOOK_URL = f"https://{HOST}/{TOKEN}"

# Chunk size: number of pages per processing batch
CHUNK_SIZE = 30

# ---------------------- Logging ----------------------
logging.basicConfig(format="%(asctime)s %(levelname)s: %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ---------------------- Telegram App ----------------------
telegram_app = (
    Application.builder()
    .token(TOKEN)
    .connection_pool_size(100)
    .build()
)
# Increase internal HTTP timeouts
try:
    telegram_app.request_kwargs = {"read_timeout": 60, "connect_timeout": 20}
except Exception:
    pass

# ---------------------- PDF Utilities ----------------------
def extract_pdf_elements(path: str, start_page: int = 0, end_page: int = None):
    """
    Extracts elements (text, images) from a PDF slice [start_page:end_page).
    """
    doc = fitz.open(path)
    total = doc.page_count
    if end_page is None or end_page > total:
        end_page = total
    elements = []
    for i in range(start_page, end_page):
        page = doc.load_page(i)
        text = page.get_text().strip()
        if text:
            elements.append(("text", text))
        for img in page.get_images(full=True):
            xref = img[0]
            data = doc.extract_image(xref)["image"]
            elements.append(("img", data))
    doc.close()
    return elements

# Word conversion (layout)
def convert_pdf_to_docx(pdf_path: str, docx_path: str, start: int = 0, end: int = None):
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=start, end=end)
    cv.close()

# Plain conversion to Word by elements
def save_elements_to_word(elements, out_path: str):
    docx = Document()
    for typ, content in elements:
        if typ == "text":
            docx.add_paragraph(content)
        else:
            bio = io.BytesIO(content)
            bio.name = "image.png"
            docx.add_picture(bio)
    docx.save(out_path)

# Save text to TXT
def save_elements_to_txt(elements, out_path: str):
    with open(out_path, "w", encoding="utf-8") as f:
        for typ, content in elements:
            if typ == "text":
                f.write(content + "\n\n")

# Extract tables to Excel
def pdf_tables_to_excel(path: str, out_path: str):
    tables = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            tbls = page.extract_tables()
            for table in tbls:
                if table and len(table) > 1:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    tables.append(df)
    if not tables:
        return False
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for idx, df in enumerate(tables, start=1):
            sheet = f"Table{idx}"
            df.to_excel(writer, sheet_name=sheet, index=False)
    return True

# ---------------------- Handlers ----------------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Привет! Отправь PDF-файл и выбери вариант.")

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc or doc.mime_type != "application/pdf":
        return await update.message.reply_text("Пожалуйста, отправь PDF-файл.")
    file = await doc.get_file()
    path = f"/tmp/{doc.file_unique_id}.pdf"
    await file.download_to_drive(path)
    context.user_data["pdf_path"] = path
    # Keyboard
    kb = [
        [InlineKeyboardButton("Word: макет 📄", callback_data="cb_word_layout")],
        [InlineKeyboardButton("Word: текст+картинки 📝", callback_data="cb_word_all")],
        [InlineKeyboardButton("TXT: текст 📄", callback_data="cb_txt")],
        [InlineKeyboardButton("Excel: таблицы 📊", callback_data="cb_tables")],
        [InlineKeyboardButton("Чат: текст 📝", callback_data="cb_text_only")],
        [InlineKeyboardButton("Чат: текст+картинки 🖼️📝", callback_data="cb_chat_all")],
        [InlineKeyboardButton("Новый PDF 🔄", callback_data="cb_new_pdf")],
    ]
    await update.message.reply_text(
        "Выбери вариант работы с PDF:",
        reply_markup=InlineKeyboardMarkup(kb)
    )

async def cb_text_only(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await update.callback_query.edit_message_text("Сначала отправь PDF.")
    doc = fitz.open(path)
    total = doc.page_count
    doc.close()
    # Chunked text
    for start in range(0, total, CHUNK_SIZE):
        end = min(start + CHUNK_SIZE, total)
        elems = extract_pdf_elements(path, start, end)
        text_only = [c for t, c in elems if t == "text"]
        if not text_only:
            continue
        for block in text_only:
            for i in range(0, len(block), 4096):
                await context.bot.send_message(update.effective_chat.id, block[i:i+4096])
        await asyncio.sleep(0.5)

async def cb_chat_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await update.callback_query.edit_message_text("Сначала отправь PDF.")
    doc = fitz.open(path)
    total = doc.page_count
    doc.close()
    sent = set()
    for start in range(0, total, CHUNK_SIZE):
        end = min(start + CHUNK_SIZE, total)
        elems = extract_pdf_elements(path, start, end)
        for typ, content in elems:
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
        await asyncio.sleep(0.5)

async def cb_word_layout(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await context.bot.send_message(update.effective_chat.id, "Сначала отправь PDF.")
    doc = fitz.open(path)
    total = doc.page_count
    doc.close()
    # Convert entire PDF layout at once
    out = f"/tmp/{update.effective_user.id}_layout.docx"
    convert_pdf_to_docx(path, out)
    with open(out, "rb") as f:
        await context.bot.send_document(update.effective_chat.id, InputFile(f, filename="layout.docx"))

async def cb_word_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await context.bot.send_message(update.effective_chat.id, "Сначала отправь PDF.")
    # Extract all elements and save to Word
    elems = extract_pdf_elements(path)
    if not elems:
        return await context.bot.send_message(update.effective_chat.id, "Нет содержимого.")
    out = f"/tmp/{update.effective_user.id}_all.docx"
    save_elements_to_word(elems, out)
    with open(out, "rb") as f:
        await context.bot.send_document(update.effective_chat.id, InputFile(f, filename="full.docx"))

async def cb_txt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await context.bot.send_message(update.effective_chat.id, "Сначала отправь PDF.")
    elems = extract_pdf_elements(path)
    text_only = [c for t, c in elems if t == "text"]
    if not text_only:
        return await context.bot.send_message(update.effective_chat.id, "Нет текста.")
    out = f"/tmp/{update.effective_user.id}.txt"
    save_elements_to_txt(elems, out)
    with open(out, "rb") as f:
        await context.bot.send_document(update.effective_chat.id, InputFile(f, filename="text.txt"))

async def cb_tables(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await context.bot.send_message(update.effective_chat.id, "Сначала отправь PDF.")
    out = f"/tmp/{update.effective_user.id}_tables.xlsx"
    ok = pdf_tables_to_excel(path, out)
    if not ok:
        return await context.bot.send_message(update.effective_chat.id, "Таблиц не найдено.")
    with open(out, "rb") as f:
        await context.bot.send_document(update.effective_chat.id, InputFile(f, filename="tables.xlsx"))

async def cb_new_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    context.user_data.clear()
    await context.bot.send_message(update.effective_chat.id, "Отправь новый PDF-файл.")

# Register handlers
telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
telegram_app.add_handler(CallbackQueryHandler(cb_text_only, pattern="cb_text_only"))
telegram_app.add_handler(CallbackQueryHandler(cb_chat_all, pattern="cb_chat_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_word_layout, pattern="cb_word_layout"))
telegram_app.add_handler(CallbackQueryHandler(cb_word_all, pattern="cb_word_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_txt, pattern="cb_txt"))
telegram_app.add_handler(CallbackQueryHandler(cb_tables, pattern="cb_tables"))
telegram_app.add_handler(CallbackQueryHandler(cb_new_pdf, pattern="cb_new_pdf"))

# ---------------------- Flask Webhook ----------------------
@app_flask.route(f"/{TOKEN}", methods=["POST"])
def telegram_webhook():
    data = request.get_json(force=True)
    update = Update.de_json(data, telegram_app.bot)
    asyncio.create_task(telegram_app.process_update(update))
    return "ok"

@app_flask.route("/ping")
def ping():
    return "pong"

if __name__ == "__main__":
    import requests
    requests.post(f"https://api.telegram.org/bot{TOKEN}/setWebhook", data={"url": WEBHOOK_URL})
    logger.info(f"Запускаем Flask на порту {PORT}, webhook={WEBHOOK_URL}")
    app_flask.run(host="0.0.0.0", port=PORT)
