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

# Получаем токен из переменной окружения
TOKEN = os.getenv("TOKEN")
PORT = int(os.getenv("PORT", 10000))
WEBHOOK_URL = os.getenv("WEBHOOK_URL", f"https://pdf-rc9c.onrender.com/{TOKEN}")

# Проверка токена при старте
print("TOKEN:", TOKEN)
logging.basicConfig(level=logging.INFO)
app_flask = Flask(__name__)

# ========================== PDF PROCESS ===========================

async def process_pdf(file_path):
    doc = fitz.open(file_path)
    elements = []
    for page in doc:
        text = page.get_text().strip()
        if text:
            elements.append(('text', text))
        images = page.get_images(full=True)
        for img in images:
            xref = img[0]
            base_image = doc.extract_image(xref)
            img_bytes = base_image["image"]
            elements.append(('img', img_bytes))
    doc.close()
    return elements

async def send_pdf_content(update, context, elements):
    sent_imgs = set()
    message_ids = []
    # Сначала текст и картинки по порядку
    for elem_type, content in elements:
        if elem_type == 'text':
            for i in range(0, len(content), 4096):
                msg = await context.bot.send_message(
                    chat_id=update.effective_chat.id,
                    text=content[i:i+4096]
                )
                message_ids.append(msg.message_id)
                await asyncio.sleep(0.2)
        elif elem_type == 'img':
            h = hash(content)
            if h in sent_imgs:
                continue
            sent_imgs.add(h)
            bio = io.BytesIO(content)
            bio.name = 'image.png'
            msg = await context.bot.send_photo(
                chat_id=update.effective_chat.id,
                photo=bio
            )
            message_ids.append(msg.message_id)
            await asyncio.sleep(0.5)
    # Кнопки внизу
    keyboard = [
        [InlineKeyboardButton("Скачать в Word", callback_data='download_word')],
        [InlineKeyboardButton("Загрузить ещё PDF-файл", callback_data='upload_pdf')]
    ]
    markup = InlineKeyboardMarkup(keyboard)
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text="Ваш текст готов! Если хотите скачать в формате Word, нажмите на кнопку ниже.",
        reply_markup=markup
    )

def elements_to_word(elements, output_path):
    docx = Document()
    for elem_type, content in elements:
        if elem_type == 'text':
            docx.add_paragraph(content)
        elif elem_type == 'img':
            bio = io.BytesIO(content)
            bio.name = 'image.png'
            docx.add_picture(bio, width=None)
    docx.save(output_path)

# ========================== TELEGRAM HANDLERS ===========================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! Отправь мне PDF-файл, и я распознаю его содержимое."
    )

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file = update.message.document
    if not file.file_name.lower().endswith('.pdf'):
        await update.message.reply_text("Отправь именно PDF-файл.")
        return
    await update.message.reply_text("⏳ Обрабатываю файл, это может занять несколько секунд...")
    file_path = await file.get_file()
    local_path = f"/tmp/{file.file_unique_id}.pdf"
    await file_path.download_to_drive(local_path)
    elements = await process_pdf(local_path)
    context.user_data['elements'] = elements
    await send_pdf_content(update, context, elements)

async def button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == 'download_word':
        elements = context.user_data.get('elements')
        if not elements:
            await query.edit_message_text("Сначала отправьте PDF-файл!")
            return
        output_path = f"/tmp/{query.from_user.id}_converted.docx"
        elements_to_word(elements, output_path)
        with open(output_path, "rb") as f:
            await context.bot.send_document(
                chat_id=query.message.chat_id,
                document=InputFile(f, filename="converted.docx")
            )
    elif query.data == 'upload_pdf':
        await context.bot.send_message(
            chat_id=query.message.chat_id,
            text="📄 Отправьте новый PDF-файл прямо в этот чат."
        )

# ========================== FLASK + WEBHOOK ===========================

telegram_app = Application.builder().token(TOKEN).connection_pool_size(200).build()
telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
telegram_app.add_handler(CallbackQueryHandler(button))

@app_flask.route(f"/{TOKEN}", methods=["POST"])
def webhook():
    update = Update.de_json(request.get_json(force=True), telegram_app.bot)
    asyncio.run(telegram_app.process_update(update))
    return "ok"

@app_flask.route("/ping")
def ping():
    return "pong"

if __name__ == "__main__":
    # Установим webhook (один раз при старте)
    import requests
    requests.post(
        f"https://api.telegram.org/bot{TOKEN}/setWebhook",
        data={"url": WEBHOOK_URL}
    )
    logging.info(f"Запускаем Flask на порту {PORT}, webhook={WEBHOOK_URL}")
    app_flask.run(host="0.0.0.0", port=PORT)
