import os
import io
import logging
import asyncio
from flask import Flask, request
import fitz  # PyMuPDF
from telegram import Update, InputFile, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler, ContextTypes,
    filters, CallbackQueryHandler
)
from docx import Document

# Configuration
token = os.getenv("TOKEN")
if not token:
    raise RuntimeError("Environment variable TOKEN is required")

port = int(os.getenv("PORT", 10000))
host = os.getenv("RENDER_EXTERNAL_HOSTNAME")
if not host:
    raise RuntimeError("Environment variable RENDER_EXTERNAL_HOSTNAME is required")
webhook_url = f"https://{host}/{token}"

# Logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

# Flask app
app_flask = Flask(__name__)

# Telegram Application
telegram_app = (
    Application.builder()
    .token(token)
    .connection_pool_size(200)
    .build()
)

# PDF processing functions
async def process_pdf(file_path):
    doc = fitz.open(file_path)
    elements = []
    for page in doc:
        text = page.get_text().strip()
        if text:
            elements.append(("text", text))
        images = page.get_images(full=True)
        for img in images:
            xref = img[0]
            img_data = doc.extract_image(xref)["image"]
            elements.append(("img", img_data))
    doc.close()
    return elements

async def send_pdf_content(update: Update, context: ContextTypes.DEFAULT_TYPE, elements):
    sent_images = set()
    for elem_type, content in elements:
        if elem_type == "text":
            text = content
            for i in range(0, len(text), 4096):
                await context.bot.send_message(
                    chat_id=update.effective_chat.id,
                    text=text[i:i+4096]
                )
                await asyncio.sleep(0.1)
        elif elem_type == "img":
            h = hash(content)
            if h in sent_images:
                continue
            sent_images.add(h)
            bio = io.BytesIO(content)
            bio.name = "image.png"
            await context.bot.send_photo(
                chat_id=update.effective_chat.id,
                photo=bio
            )
            await asyncio.sleep(0.2)
    # Buttons
    keyboard = [
        [InlineKeyboardButton("Скачать в Word", callback_data="download_word")],
        [InlineKeyboardButton("Загрузить ещё PDF-файл", callback_data="upload_pdf")]
    ]
    markup = InlineKeyboardMarkup(keyboard)
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text="Ваш текст готов! Если хотите скачать в формате Word, нажмите на кнопку ниже.",
        reply_markup=markup
    )

# Convert to Word
def elements_to_word(elements, output_path):
    docx = Document()
    for elem_type, content in elements:
        if elem_type == "text":
            docx.add_paragraph(content)
        elif elem_type == "img":
            bio = io.BytesIO(content)
            bio.name = "image.png"
            docx.add_picture(bio)
    docx.save(output_path)

# Handlers
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! Отправь мне PDF-файл, и я распознаю его содержимое."
    )

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("⏳ Обрабатываю PDF...")
    doc = update.message.document
    if doc.mime_type != "application/pdf":
        return await update.message.reply_text("Пожалуйста, отправьте PDF-файл.")
    file = await doc.get_file()
    path = f"/tmp/{doc.file_unique_id}.pdf"
    await file.download_to_drive(path)
    elements = await process_pdf(path)
    context.user_data["elements"] = elements
    await send_pdf_content(update, context, elements)

async def button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    if data == "download_word":
        elements = context.user_data.get("elements")
        if not elements:
            return await query.edit_message_text("Сначала отправьте PDF-файл.")
        out_path = f"/tmp/{query.from_user.id}_converted.docx"
        elements_to_word(elements, out_path)
        with open(out_path, "rb") as f:
            await context.bot.send_document(
                chat_id=query.message.chat_id,
                document=InputFile(f, filename="converted.docx")
            )
    elif data == "upload_pdf":
        await context.bot.send_message(
            chat_id=query.message.chat_id,
            text="📄 Отправьте новый PDF-файл прямо в этот чат."
        )

# Register handlers
telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
telegram_app.add_handler(CallbackQueryHandler(button))

# Global event loop
loop = asyncio.new_event_loop()
asyncio.set_event_loop(loop)

# Initialize application once
async def init_app():
    if not telegram_app._initialized:
        await telegram_app.initialize()
    logger.info("Telegram application initialized")

loop.run_until_complete(init_app())

# Flask routes
@app_flask.route(f"/{token}", methods=["POST"])
def webhook():
    update = Update.de_json(request.get_json(force=True), telegram_app.bot)
    # Process update asynchronously, respond immediately
    asyncio.run_coroutine_threadsafe(telegram_app.process_update(update), loop)
    return "ok"

@app_flask.route("/ping")
def ping():
    return "pong", 200

# Start
if __name__ == "__main__":
    import requests
    # Set webhook on startup
    requests.post(
        f"https://api.telegram.org/bot{token}/setWebhook",
        data={"url": webhook_url}
    )
    logger.info(f"Starting Flask on port {port}, webhook={webhook_url}")
    app_flask.run(host="0.0.0.0", port=port)
