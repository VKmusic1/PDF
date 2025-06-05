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
# Тайм-ауты оставляем только для внутренних http-запросов PTB, 
# но не передаём их в сами send_document/send_photo/send_message
telegram_app.request_kwargs = {
    "read_timeout": 60,
    "connect_timeout": 20
}

# ---------------------- 4. Функции для работы с PDF ----------------------
def extract_pdf_elements(path: str):
    """
    Открывает PDF через PyMuPDF и возвращает список элементов:
    - ('text', строка)
    - ('img', bytes изображения)
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

def convert_to_word(elements, out_path: str):
    """
    Конвертирует список элементов в DOCX:
    - текст -> параграфы;
    - картинки -> вставляет в документ.
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

# ---------------------- 5. Обработчики ----------------------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    /start
    """
    await update.message.reply_text("Привет! Отправь PDF-файл, и я предложу варианты извлечения.")

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    При получении PDF:
     - сохраняем во /tmp
     - показываем кнопки выбора:
         • Скачать текст и картинки в Word
         • Скачать текст TXT
         • Скачать просто текст
         • Скачать текст и картинки в чат
         • Скачать таблицу в Excel
         • Новый PDF
    """
    logger.info("Получен документ от %s", update.effective_user.id)
    doc = update.message.document
    if not doc or doc.mime_type != "application/pdf":
        return await update.message.reply_text("Пожалуйста, отправь PDF-файл.")
    file = await doc.get_file()
    path = f"/tmp/{doc.file_unique_id}.pdf"
    await file.download_to_drive(path)
    context.user_data["pdf_path"] = path

    keyboard = [
        [InlineKeyboardButton("Скачать текст и картинки в Word", callback_data="cb_word_all")],
        [InlineKeyboardButton("Скачать текст TXT", callback_data="cb_txt")],
        [InlineKeyboardButton("Скачать просто текст", callback_data="cb_text_only")],
        [InlineKeyboardButton("Скачать текст и картинки в чат", callback_data="cb_chat_all")],
        [InlineKeyboardButton("Скачать таблицу в Excel", callback_data="cb_tables")],
        [InlineKeyboardButton("Новый PDF", callback_data="cb_new_pdf")],
    ]
    await update.message.reply_text(
        "Выбери, что сделать с этим PDF:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def cb_text_only(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Вывести только текст в чат (страница за страницей), без кнопок.
    """
    user = update.effective_user.id
    logger.info("Callback cb_text_only от %s", user)
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
    """
    Вывести в чат текст и картинки (по порядку), без кнопок.
    """
    user = update.effective_user.id
    logger.info("Callback cb_chat_all от %s", user)
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
    """
    Конвертировать весь PDF (текст+картинки) в Word и отправить.
    """
    user = update.effective_user.id
    logger.info("Callback cb_word_all от %s", user)
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await context.bot.send_message(update.effective_chat.id, "Файл не найден. Отправь новый PDF.")
    elements = extract_pdf_elements(path)
    if not elements:
        return await context.bot.send_message(update.effective_chat.id, "В PDF нет контента.")
    out = f"/tmp/{user}_all.docx"
    convert_to_word(elements, out)
    with open(out, "rb") as f:
        await context.bot.send_document(
            update.effective_chat.id,
            document=InputFile(f, filename="full_converted.docx")
        )

async def cb_txt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Положить весь текст в .txt и отправить.
    """
    user = update.effective_user.id
    logger.info("Callback cb_txt от %s", user)
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await context.bot.send_message(update.effective_chat.id, "Файл не найден. Отправь новый PDF.")
    elements = extract_pdf_elements(path)
    text_only = [c for t, c in elements if t == "text"]
    if not text_only:
        return await context.bot.send_message(update.effective_chat.id, "В PDF нет текста.")
    all_text = "\n\n".join(text_only)
    out_path = f"/tmp/{user}.txt"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(all_text)
    with open(out_path, "rb") as f:
        await context.bot.send_document(
            update.effective_chat.id,
            document=InputFile(f, filename="full_converted.txt")
        )

async def cb_tables(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Извлечь таблицы через pdfplumber в Excel и отправить.
    """
    user = update.effective_user.id
    logger.info("Callback cb_tables от %s", user)
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
    excel_path = f"/tmp/{user}_tables.xlsx"
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        for sheet_name, df in all_tables:
            safe_name = sheet_name[:31]  # Excel ограничение: 31 символ
            df.to_excel(writer, sheet_name=safe_name, index=False)
    with open(excel_path, "rb") as f:
        await context.bot.send_document(
            update.effective_chat.id,
            document=InputFile(f, filename="tables.xlsx")
        )

async def cb_new_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Новый PDF: очищаем user_data и просим загрузить снова.
    """
    user = update.effective_user.id
    logger.info("Callback cb_new_pdf от %s", user)
    await update.callback_query.answer()
    context.user_data.clear()
    await context.bot.send_message(update.effective_chat.id, "Отправьте новый PDF-файл.")

# ---------------------- 6. Регистрация хендлеров ----------------------
telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
telegram_app.add_handler(CallbackQueryHandler(cb_text_only, pattern="cb_text_only"))
telegram_app.add_handler(CallbackQueryHandler(cb_chat_all, pattern="cb_chat_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_word_all, pattern="cb_word_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_txt, pattern="cb_txt"))
telegram_app.add_handler(CallbackQueryHandler(cb_tables, pattern="cb_tables"))
telegram_app.add_handler(CallbackQueryHandler(cb_new_pdf, pattern="cb_new_pdf"))

# ---------------------- 7. Запуск webhook ----------------------
if __name__ == "__main__":
    logger.info(f"Setting webhook to {WEBHOOK_URL}")
    telegram_app.run_webhook(
        listen="0.0.0.0",
        port=PORT,
        url_path=TOKEN,
        webhook_url=WEBHOOK_URL
    )
