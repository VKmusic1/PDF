import os
import io
import logging
import asyncio
import fitz  # PyMuPDF
from flask import Flask, request
from telegram import Update, InputFile, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters
)
from docx import Document

# Configuration from environment
TOKEN = os.getenv("TOKEN")
if not TOKEN:
    raise RuntimeError("Environment variable TOKEN is required")

HOST = os.getenv("RENDER_EXTERNAL_HOSTNAME")
if not HOST:
    raise RuntimeError("Environment variable RENDER_EXTERNAL_HOSTNAME is required")

PORT = int(os.getenv("PORT", "10000"))
WEBHOOK_URL = f"https://{HOST}/{TOKEN}"

# Logging setup
logging.basicConfig(
    format="%(asctime)s %(levelname)s: %(message)s", level=logging.INFO
)
logger = logging.getLogger(__name__)

# Initialize Telegram Application
app = Application.builder() \
    .token(TOKEN) \
    .connection_pool_size(100) \
    .build()

# PDF Processing
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

async def send_pdf_content(update: Update, context: ContextTypes.DEFAULT_TYPE, elements):
    sent = set()
    chat_id = update.effective_chat.id
    for typ, content in elements:
        if typ == "text":
            text = content
            for i in range(0, len(text), 4096):
                await context.bot.send_message(chat_id, text[i:i+4096])
                await asyncio.sleep(0.1)
        else:
            h = hash(content)
            if h in sent:
                continue
            sent.add(h)
            bio = io.BytesIO(content)
            bio.name = "image.png"
            await context.bot.send_photo(chat_id, photo=bio)
            await asyncio.sleep(0.1)
    kb = [
        [InlineKeyboardButton("Скачать в Word", callback_data="download_word")],
        [InlineKeyboardButton("Загрузить ещё PDF-файл", callback_data="upload_pdf")]
    ]
    await context.bot.send_message(chat_id, "Ваш текст готов!", reply_markup=InlineKeyboardMarkup(kb))

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

# Handlers
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Привет! Отправь PDF-файл.")

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if doc.mime_type != "application/pdf":
        return await update.message.reply_text("Пожалуйста, отправьте PDF.")
    await update.message.reply_text("Обработка...")
    f = await doc.get_file()
    path = f"/tmp/{doc.file_unique_id}.pdf"
    await f.download_to_drive(path)
    elements = extract_pdf_elements(path)
    context.user_data["elements"] = elements
    await send_pdf_content(update, context, elements)

async def button_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    elements = context.user_data.get("elements", [])
    if data == "download_word":
        out = f"/tmp/{query.from_user.id}.docx"
        convert_to_word(elements, out)
        with open(out, "rb") as f:
            await context.bot.send_document(
                update.effective_chat.id,
                document=InputFile(f, filename="converted.docx")
            )
    else:
        await context.bot.send_message(update.effective_chat.id, "Отправьте новый PDF.")

# Register handlers
app.add_handler(CommandHandler("start", start))
app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
app.add_handler(CallbackQueryHandler(button_callback))

# Start webhook server
if __name__ == "__main__":
    logger.info(f"Setting webhook to {WEBHOOK_URL}")
    app.run_webhook(
        listen="0.0.0.0",
        port=PORT,
        url_path=TOKEN,
        webhook_url=WEBHOOK_URL,
    )
