import os
import io
import logging
import asyncio
import fitz  # PyMuPDF
from telegram import Update, InputFile, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters
)
from docx import Document

# Configuration from environment
token = os.getenv("TOKEN")
if not token:
    raise RuntimeError("Environment variable TOKEN is required")
host = os.getenv("RENDER_EXTERNAL_HOSTNAME")
if not host:
    raise RuntimeError("Environment variable RENDER_EXTERNAL_HOSTNAME is required")
port = int(os.getenv("PORT", "10000"))
webhook_url = f"https://{host}/{token}"

# Logging setup
logging.basicConfig(
    format="%(asctime)s %(levelname)s %(message)s", level=logging.INFO
)
logger = logging.getLogger(__name__)

# Initialize Telegram Application
app = Application.builder() 
                .token(token) 
                .connection_pool_size(100) 
                .build()

# PDF Processing
async def process_pdf(path: str):
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

async def send_content(update: Update, context: ContextTypes.DEFAULT_TYPE, elements):
    # send text and images
    sent = set()
    chat_id = update.effective_chat.id
    for typ, content in elements:
        if typ == "text":
            for i in range(0, len(content), 4096):
                await context.bot.send_message(chat_id, content[i:i+4096])
                await asyncio.sleep(0.1)
        else:  # image
            h = hash(content)
            if h in sent:
                continue
            sent.add(h)
            bio = io.BytesIO(content)
            bio.name = "image.png"
            await context.bot.send_photo(chat_id, photo=bio)
            await asyncio.sleep(0.1)
    # send buttons
    kb = [
        [InlineKeyboardButton("Скачать в Word", callback_data="download_word")],
        [InlineKeyboardButton("Загрузить ещё PDF-файл", callback_data="upload_pdf")]
    ]
    markup = InlineKeyboardMarkup(kb)
    await context.bot.send_message(
        chat_id, "Ваш текст готов!", reply_markup=markup
    )

# Convert to Word
def elements_to_word(elements, out_path: str):
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
        return await update.message.reply_text("Пожалуйста, PDF.")
    await update.message.reply_text("Обработка...")
    f = await doc.get_file()
    path = f"/tmp/{doc.file_unique_id}.pdf"
    await f.download_to_drive(path)
    elems = await process_pdf(path)
    context.user_data["elems"] = elems
    await send_content(update, context, elems)

async def button_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    elems = context.user_data.get("elems") or []
    if data == "download_word":
        out = f"/tmp/{query.from_user.id}.docx"
        elements_to_word(elems, out)
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
app.add_handler(CallbackQueryHandler(button_cb))

# Start webhook
if __name__ == "__main__":
    logger.info(f"Setting webhook: {webhook_url}")
    app.run_webhook(
        listen="0.0.0.0",
        port=port,
        url_path=token,
        webhook_url=webhook_url,
    )
