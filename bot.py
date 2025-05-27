import os
import io
import logging
import asyncio
import fitz  # PyMuPDF
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

# Конфигурация из окружения
TOKEN = os.getenv("TOKEN")
if not TOKEN:
    raise RuntimeError("Environment variable TOKEN is required")

HOST = os.getenv("RENDER_EXTERNAL_HOSTNAME")
if not HOST:
    raise RuntimeError("Environment variable RENDER_EXTERNAL_HOSTNAME is required")

PORT = int(os.getenv("PORT", "10000"))
WEBHOOK_URL = f"https://{HOST}/{TOKEN}"

# Логирование
logging.basicConfig(format="%(asctime)s %(levelname)s: %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# Инициализация Telegram Application
telegram_app = (
    Application.builder()
    .token(TOKEN)
    .connection_pool_size(100)
    .build()
)

# Функция извлечения элементов PDF
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

# Отправка содержимого
async def send_elements(update: Update, context: ContextTypes.DEFAULT_TYPE, elements):
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
    # Кнопки по завершении
    keyboard = [
        [InlineKeyboardButton("Скачать в Word", callback_data="download_word")],
        [InlineKeyboardButton("Новый PDF", callback_data="new_pdf")]
    ]
    await context.bot.send_message(chat_id, "Готово!", reply_markup=InlineKeyboardMarkup(keyboard))

# Конвертация в Word
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

# Обработчики
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Привет! Отправь PDF-файл.")

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if doc.mime_type != "application/pdf":
        return await update.message.reply_text("Пожалуйста, отправьте PDF.")
    # Сохраняем PDF и предлагаем выбор
    file = await doc.get_file()
    path = f"/tmp/{doc.file_unique_id}.pdf"
    await file.download_to_drive(path)
    context.user_data['pdf_path'] = path
    await update.message.reply_text(
        "Выберите, что извлечь:",
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("Только текст", callback_data="only_text")],
            [InlineKeyboardButton("Текст + картинки", callback_data="text_images")]
        ])
    )

async def only_text_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get('pdf_path')
    if not path:
        return await update.callback_query.edit_message_text("Файл не найден.")
    elements = extract_pdf_elements(path)
    text_only = [(t, c) for t, c in elements if t == "text"]
    await send_elements(update, context, text_only)
    context.user_data['elements'] = text_only

async def text_images_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get('pdf_path')
    if not path:
        return await update.callback_query.edit_message_text("Файл не найден.")
    elements = extract_pdf_elements(path)
    await send_elements(update, context, elements)
    context.user_data['elements'] = elements

async def download_word_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    elements = context.user_data.get('elements', [])
    if not elements:
        return await update.callback_query.edit_message_text("Нет данных для конвертации.")
    out = f"/tmp/{update.effective_user.id}.docx"
    convert_to_word(elements, out)
    with open(out, 'rb') as f:
        await context.bot.send_document(update.effective_chat.id, InputFile(f, filename="converted.docx"))

async def new_pdf_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    context.user_data.clear()
    await context.bot.send_message(update.effective_chat.id, "Отправьте новый PDF-файл.")

# Регистрация хендлеров
telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
telegram_app.add_handler(CallbackQueryHandler(only_text_callback, pattern="only_text"))
telegram_app.add_handler(CallbackQueryHandler(text_images_callback, pattern="text_images"))
telegram_app.add_handler(CallbackQueryHandler(download_word_callback, pattern="download_word"))
telegram_app.add_handler(CallbackQueryHandler(new_pdf_callback, pattern="new_pdf"))

# Запуск webhook
if __name__ == "__main__":
    logger.info(f"Setting webhook to {WEBHOOK_URL}")
    telegram_app.run_webhook(
        listen="0.0.0.0",
        port=PORT,
        url_path=TOKEN,
        webhook_url=WEBHOOK_URL
    )
