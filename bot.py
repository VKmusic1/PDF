import os
import io
import logging
import asyncio
import fitz                  # PyMuPDF
import pdfplumber
import pandas as pd
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

# ---------------------- 3. Инициализация Telegram Application ----------------------
telegram_app = (
    Application.builder()
    .token(TOKEN)
    .connection_pool_size(100)
    .build()
)
# Тайм-ауты для send_document (все остальные вызовы идут без timeout)
telegram_app.request_kwargs = {
    "read_timeout": 60,
    "connect_timeout": 20
}

# ---------------------- 4. Функции для работы с PDF ----------------------
def extract_pdf_elements(path: str):
    """
    Открывает PDF через PyMuPDF, возвращает список элементов:
    ('text', строка текста) или ('img', bytes_изображения).
    """
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

async def send_elements(update: Update, context: ContextTypes.DEFAULT_TYPE, elements):
    """
    Проходит по списку элементов и отправляет:
     - каждый текст по 4096 символов
     - каждую картинку
    После этого выводит четыре кнопки:
     «Скачать в Word», «Скачать в TXT», «Скачать таблицы», «Новый PDF»
    и сообщение «Чтобы пользоваться ботом, нажмите /start».
    """
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

    keyboard = [
        [InlineKeyboardButton("Скачать в Word 💾", callback_data="download_word")],
        [InlineKeyboardButton("Скачать в TXT 📄", callback_data="download_txt")],
        [InlineKeyboardButton("Скачать таблицы 📊", callback_data="download_tables")],
        [InlineKeyboardButton("Новый PDF 🔖", callback_data="new_pdf")]
    ]
    await context.bot.send_message(
        chat_id,
        "Готово!",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )
    await context.bot.send_message(chat_id, "Чтобы пользоваться ботом, нажмите /start")

def convert_to_word(elements, out_path: str):
    """
    Конвертирует список элементов в DOCX:
    - текст -> параграфы
    - изображения -> вставляет в документ
    """
    docx = Document()
    for typ, content in elements:
        if typ == "text":
            docx.add_paragraph(content)
        else:
            bio = io.BytesIO(content)
            bio.name = "image.png"
            docx.add_picture(bio)
    docx.save(out_path)

# ---------------------- 5. Обработчики команд и callback ----------------------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    /start
    """
    await update.message.reply_text("Привет! Отправь PDF-файл.")

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    При получении PDF:
     - сохраняет файл во /tmp
     - предлагает кнопку «Только текст» и «Текст + картинки»
    """
    logger.info("Получен документ от %s", update.effective_user.id)
    doc = update.message.document
    if not doc or doc.mime_type != "application/pdf":
        return await update.message.reply_text("Пожалуйста, отправьте PDF-файл.")
    file = await doc.get_file()
    path = f"/tmp/{doc.file_unique_id}.pdf"
    await file.download_to_drive(path)
    context.user_data["pdf_path"] = path

    await update.message.reply_text(
        "Выберите, что извлечь из PDF:",
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("Только текст", callback_data="only_text")],
            [InlineKeyboardButton("Текст + картинки", callback_data="text_images")]
        ])
    )

async def only_text_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    При нажатии «Только текст»:
     - извлекает только текстовые блоки
     - вызывает send_elements()
    """
    logger.info("Callback only_text от %s", update.effective_user.id)
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await update.callback_query.edit_message_text("Файл не найден.")
    elements = extract_pdf_elements(path)
    text_only = [(t, c) for t, c in elements if t == "text"]
    context.user_data["elements"] = text_only
    await send_elements(update, context, text_only)

async def text_images_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    При нажатии «Текст + картинки»:
     - извлекает и текст, и картинки
     - вызывает send_elements()
    """
    logger.info("Callback text_images от %s", update.effective_user.id)
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await update.callback_query.edit_message_text("Файл не найден.")
    elements = extract_pdf_elements(path)
    context.user_data["elements"] = elements
    await send_elements(update, context, elements)

async def download_txt_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    При нажатии «Скачать в TXT»:
     - собирает весь текст
     - записывает в /tmp/USER_ID.txt
     - отправляет через send_document(timeout=60)
    """
    logger.info("Callback download_txt от %s", update.effective_user.id)
    await update.callback_query.answer()
    elements = context.user_data.get("elements", [])
    if not elements:
        return await update.callback_query.edit_message_text("Нет данных для конвертации.")
    all_text = ""
    for typ, content in elements:
        if typ == "text":
            all_text += content + "\n\n"
    if not all_text:
        return await update.callback_query.edit_message_text("В PDF нет текста.")
    out_path = f"/tmp/{update.effective_user.id}.txt"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(all_text)
    with open(out_path, "rb") as f:
        await context.bot.send_document(
            update.effective_chat.id,
            document=InputFile(f, filename="converted.txt"),
            timeout=60
        )

async def download_word_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    При нажатии «Скачать в Word»:
     - конвертирует элементы в DOCX
     - отправляет через send_document(timeout=60)
    """
    logger.info("Callback download_word от %s", update.effective_user.id)
    await update.callback_query.answer()
    elements = context.user_data.get("elements", [])
    if not elements:
        return await update.callback_query.edit_message_text("Нет данных для конвертации.")
    out = f"/tmp/{update.effective_user.id}.docx"
    convert_to_word(elements, out)
    with open(out, "rb") as f:
        await context.bot.send_document(
            update.effective_chat.id,
            document=InputFile(f, filename="converted.docx"),
            timeout=60
        )

async def download_tables_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    При нажатии «Скачать таблицы»:
     - извлекает таблицы через pdfplumber,
     - кладёт их в Excel (openpyxl),
     - отправляет через send_document(timeout=60)
    """
    logger.info("Callback download_tables от %s", update.effective_user.id)
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await update.callback_query.edit_message_text("PDF не найден.")
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
        return await update.callback_query.edit_message_text("В PDF нет распознаваемых таблиц.")
    excel_path = f"/tmp/{update.effective_user.id}_tables.xlsx"
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        for sheet_name, df in all_tables:
            safe_name = sheet_name[:31]  # ограничение Excel: 31 символ
            df.to_excel(writer, sheet_name=safe_name, index=False)
    with open(excel_path, "rb") as f:
        await context.bot.send_document(
            update.effective_chat.id,
            document=InputFile(f, filename="tables.xlsx"),
            timeout=60
        )

async def new_pdf_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    При нажатии «Новый PDF»:
     - очищает user_data,
     - просит загрузить новый файл.
    """
    logger.info("Callback new_pdf от %s", update.effective_user.id)
    await update.callback_query.answer()
    context.user_data.clear()
    await context.bot.send_message(update.effective_chat.id, "Отправьте новый PDF-файл.")

# ---------------------- 6. Регистрация хендлеров ----------------------
telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
telegram_app.add_handler(CallbackQueryHandler(only_text_callback, pattern="only_text"))
telegram_app.add_handler(CallbackQueryHandler(text_images_callback, pattern="text_images"))
telegram_app.add_handler(CallbackQueryHandler(download_word_callback, pattern="download_word"))
telegram_app.add_handler(CallbackQueryHandler(download_txt_callback, pattern="download_txt"))
telegram_app.add_handler(CallbackQueryHandler(download_tables_callback, pattern="download_tables"))
telegram_app.add_handler(CallbackQueryHandler(new_pdf_callback, pattern="new_pdf"))

# ---------------------- 7. Запуск webhook ----------------------
if __name__ == "__main__":
    logger.info(f"Setting webhook to {WEBHOOK_URL}")
    telegram_app.run_webhook(
        listen="0.0.0.0",
        port=PORT,
        url_path=TOKEN,
        webhook_url=WEBHOOK_URL
    )
