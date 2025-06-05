import os
import io
import logging
import asyncio
import fitz            # PyMuPDF
import pdfplumber
import pandas as pd
from flask import Flask, request
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
import requests  # нужен для установки webhook

# ---------------------- 1. Проверяем переменные окружения ----------------------

TOKEN = os.getenv("TOKEN")
if not TOKEN:
    raise RuntimeError("Environment variable TOKEN is required")

HOST = os.getenv("RENDER_EXTERNAL_HOSTNAME")
if not HOST:
    raise RuntimeError("Environment variable RENDER_EXTERNAL_HOSTNAME is required")

PORT = int(os.getenv("PORT", "10000"))
WEBHOOK_URL = f"https://{HOST}/{TOKEN}"

# ---------------------- 2. Настраиваем логирование ----------------------

logging.basicConfig(format="%(asctime)s %(levelname)s: %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ---------------------- 3. Создаём Flask-приложение ----------------------

app = Flask(__name__)

# ---------------------- 4. Формируем единый asyncio-loop ----------------------
#     (чтобы Telegram Application и Flask работали в одном loop-е)

loop = asyncio.new_event_loop()
asyncio.set_event_loop(loop)

# ---------------------- 5. Инициализируем Telegram Application ----------------------

telegram_app = (
    Application.builder()
    .token(TOKEN)
    .connection_pool_size(100)
    .build()
)

# Увеличиваем тайм‐ауты HTTP‐запросов внутри python‐telegram‐bot
telegram_app.request_kwargs = {
    "read_timeout": 60,
    "connect_timeout": 20
}

# Дождёмся полной инициализации приложения
loop.run_until_complete(telegram_app.initialize())

# ---------------------- 6. PDF‐утилиты ----------------------

def extract_pdf_elements(path: str):
    """
    Открывает PDF через PyMuPDF и возвращает список элементов:
    - ('text', <строка текста>)
    - ('img', <bytes изображения>)
    """
    doc = fitz.open(path)
    elements = []
    for page in doc:
        txt = page.get_text().strip()
        if txt:
            elements.append(("text", txt))
        for img in page.get_images(full=True):
            xref = img[0]
            data = doc.extract_image(xref)["image"]
            elements.append(("img", data))
    doc.close()
    return elements

def convert_to_word(elements, out_path: str):
    """
    Конвертирует список элементов в DOCX:
    - text → параграфы
    - img → вставляет картинку
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

def save_text_to_txt(elements, out_path: str):
    """
    Записывает из элементов только текстовые блоки в .txt, 
    разделяя каждый абзац двойным переводом строки.
    """
    with open(out_path, "w", encoding="utf-8") as f:
        for typ, content in elements:
            if typ == "text":
                f.write(content + "\n\n")

# ---------------------- 7. Telegram‐обработчики ----------------------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Хендлер /start — присылаем приветственное сообщение.
    """
    await update.message.reply_text("Привет! Отправь PDF-файл, и я предложу варианты извлечения →")

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    При получении PDF:
      1) Скачиваем его в /tmp/<file_unique_id>.pdf
      2) Извлекаем из него текст+изображения и сохраняем в context.user_data
      3) Отправляем меню кнопок
    """
    logger.info("Получен документ от %s", update.effective_user.id)
    doc = update.message.document
    if not doc or doc.mime_type != "application/pdf":
        return await update.message.reply_text("Пожалуйста, отправь именно PDF-файл.")

    # 1) Скачиваем PDF во временный файл
    tg_file = await doc.get_file()
    local_path = f"/tmp/{doc.file_unique_id}.pdf"
    await tg_file.download_to_drive(local_path)
    context.user_data["pdf_path"] = local_path

    # 2) Извлекаем все элементы (text + img) сразу
    elements = extract_pdf_elements(local_path)
    context.user_data["elements"] = elements

    # 3) Отправляем меню кнопок
    keyboard = [
        [InlineKeyboardButton("Word: текст+картинки 📄", callback_data="cb_word_all")],
        [InlineKeyboardButton("TXT: текст 📄",            callback_data="cb_txt")],
        [InlineKeyboardButton("Excel: таблицы 📊",         callback_data="cb_tables")],
        [InlineKeyboardButton("Чат: текст 📝",            callback_data="cb_text_only")],
        [InlineKeyboardButton("Чат: текст+картинки 🖼️📝",  callback_data="cb_chat_all")],
        [InlineKeyboardButton("Новый PDF 🔄",              callback_data="cb_new_pdf")],
    ]
    await update.message.reply_text(
        "Выбери, что сделать с этим PDF:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def cb_text_only(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Кнопка “Чат: текст 📝” — выводим в чат только текстовые блоки (без кнопок).
    """
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await update.callback_query.edit_message_text("Файл не найден. Пришли новый PDF.")

    elements = extract_pdf_elements(path)
    text_blocks = [c for t, c in elements if t == "text"]
    if not text_blocks:
        return await context.bot.send_message(update.effective_chat.id, "В PDF нет текста.")

    for block in text_blocks:
        for i in range(0, len(block), 4096):
            await context.bot.send_message(update.effective_chat.id, block[i:i+4096])
            await asyncio.sleep(0.05)

async def cb_chat_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Кнопка “Чат: текст+картинки 🖼️📝” — выводим всё содержимое (text + img) в порядке,
    без кнопок внизу.
    """
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await update.callback_query.edit_message_text("Файл не найден. Пришли новый PDF.")

    elements = extract_pdf_elements(path)
    if not elements:
        return await context.bot.send_message(update.effective_chat.id, "В PDF нет контента.")

    sent_images = set()
    for typ, content in elements:
        if typ == "text":
            for i in range(0, len(content), 4096):
                await context.bot.send_message(update.effective_chat.id, content[i:i+4096])
                await asyncio.sleep(0.05)
        else:
            h = hash(content)
            if h in sent_images:
                continue
            sent_images.add(h)
            bio = io.BytesIO(content)
            bio.name = "image.png"
            await context.bot.send_photo(update.effective_chat.id, photo=bio)
            await asyncio.sleep(0.1)

async def cb_word_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Кнопка “Word: текст+картинки 📄” — создаём один .docx из всего PDF
    (включая картинки) и отправляем пользователю.
    """
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await context.bot.send_message(update.effective_chat.id, "Файл не найден. Пришли новый PDF.")

    elements = extract_pdf_elements(path)
    if not elements:
        return await context.bot.send_message(update.effective_chat.id, "В PDF нет контента.")

    out = f"/tmp/{update.effective_user.id}_full.docx"
    convert_to_word(elements, out)
    with open(out, "rb") as f:
        await context.bot.send_document(
            update.effective_chat.id,
            document=InputFile(f, filename="converted_full.docx")
        )

async def cb_txt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Кнопка “TXT: текст 📄” — собираем весь текст в один .txt и отправляем.
    """
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await context.bot.send_message(update.effective_chat.id, "Файл не найден. Пришли новый PDF.")

    elements = extract_pdf_elements(path)
    text_blocks = [c for t, c in elements if t == "text"]
    if not text_blocks:
        return await context.bot.send_message(update.effective_chat.id, "В PDF нет текста.")

    all_text = "\n\n".join(text_blocks)
    out_txt = f"/tmp/{update.effective_user.id}_all.txt"
    with open(out_txt, "w", encoding="utf-8") as f:
        f.write(all_text)

    with open(out_txt, "rb") as f:
        await context.bot.send_document(
            update.effective_chat.id,
            document=InputFile(f, filename="converted_text.txt")
        )

async def cb_tables(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Кнопка “Excel: таблицы 📊” — ищем таблицы через pdfplumber,
    складываем их в разные листы .xlsx и отправляем.
    """
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await context.bot.send_message(update.effective_chat.id, "Файл не найден. Пришли новый PDF.")

    all_tables = []
    with pdfplumber.open(path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            for tbl_idx, table in enumerate(tables, start=1):
                if not table or len(table) < 2:
                    continue
                df = pd.DataFrame(table[1:], columns=table[0])
                sheet_name = f"Стр{page_num}_Таб{tbl_idx}"
                all_tables.append((sheet_name[:31], df))

    if not all_tables:
        return await context.bot.send_message(update.effective_chat.id, "В PDF нет распознаваемых таблиц.")

    excel_path = f"/tmp/{update.effective_user.id}_tables.xlsx"
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        for sheet_name, df in all_tables:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    with open(excel_path, "rb") as f:
        await context.bot.send_document(
            update.effective_chat.id,
            document=InputFile(f, filename="converted_tables.xlsx")
        )

async def cb_new_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Кнопка “Новый PDF 🔄” — очищаем user_data и просим прислать файл заново.
    """
    await update.callback_query.answer()
    context.user_data.clear()
    await context.bot.send_message(update.effective_chat.id, "Окей! Пришли новый PDF-файл →")

# ---------------------- 8. Регистрируем хендлеры ----------------------

telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
telegram_app.add_handler(CallbackQueryHandler(cb_text_only,   pattern="cb_text_only"))
telegram_app.add_handler(CallbackQueryHandler(cb_chat_all,    pattern="cb_chat_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_word_all,    pattern="cb_word_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_txt,         pattern="cb_txt"))
telegram_app.add_handler(CallbackQueryHandler(cb_tables,      pattern="cb_tables"))
telegram_app.add_handler(CallbackQueryHandler(cb_new_pdf,     pattern="cb_new_pdf"))

# ---------------------- 9. Flask‐маршруты ----------------------

@app.route(f"/{TOKEN}", methods=["POST"])
def telegram_webhook():
    """
    Этот маршрут вызывается Telegram при каждом обновлении (Update).
    Мы кладём пришедший Update в asyncio‐loop, чтобы обработать его асинхронно.
    """
    data = request.get_json(force=True)
    update = Update.de_json(data, telegram_app.bot)
    loop.create_task(telegram_app.process_update(update))
    return "OK"

@app.route("/ping", methods=["GET"])
def ping():
    """
    Ping‐маршрут, чтобы PingWin или Render опрашивал его каждые N минут
    и не давал сервису «уснуть». Возвращает plain‐text "pong".
    """
    return "pong"

# ---------------------- 10. Запуск ----------------------

if __name__ == "__main__":
    # 1) Устанавливаем webhook в Telegram (https://api.telegram.org/bot<TOKEN>/setWebhook)
    logger.info(f"Setting webhook to {WEBHOOK_URL}")
    resp = requests.post(
        f"https://api.telegram.org/bot{TOKEN}/setWebhook",
        data={"url": WEBHOOK_URL}
    )
    if not resp.ok:
        logger.error("Не удалось установить webhook: %s", resp.text)
    else:
        logger.info("Webhook установлен успешно → %s", WEBHOOK_URL)

    # 2) Запускаем Flask (он слушает порты для /ping и /<TOKEN>)
    logger.info(f"Запускаем Flask на порту {PORT}")
    app.run(host="0.0.0.0", port=PORT)
