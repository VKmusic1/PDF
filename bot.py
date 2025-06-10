import os
import io
import logging
import asyncio

import fitz
import pdfplumber
import pandas as pd
from PIL import Image
from docx import Document
from docx.image.exceptions import UnrecognizedImageError

from telegram import (
    Update,
    InputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
)
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters,
)

# ---------------- Конфигурация ----------------

TOKEN = os.getenv("TOKEN")
if not TOKEN:
    raise RuntimeError("Переменная окружения TOKEN не задана")

# ---------------- Логирование ----------------

logging.basicConfig(
    format="%(asctime)s %(levelname)s: %(message)s", level=logging.INFO
)
logger = logging.getLogger(__name__)

# ---------------- PDF-утилиты ----------------

def extract_pdf_elements(path: str):
    doc = fitz.open(path)
    elems = []
    for page in doc:
        txt = page.get_text().strip()
        if txt:
            elems.append(("text", txt))
        for img in page.get_images(full=True):
            data = doc.extract_image(img[0])["image"]
            elems.append(("img", data))
    doc.close()
    return elems

def save_text_to_txt(elems, out_path):
    with open(out_path, "w", encoding="utf-8") as f:
        for typ, content in elems:
            if typ == "text":
                f.write(content + "\n\n")

def convert_to_word(elems, out_path):
    docx = Document()
    for typ, content in elems:
        if typ == "text":
            docx.add_paragraph(content)
        else:
            bio = io.BytesIO(content)
            bio.name = "image.png"
            try:
                docx.add_picture(bio)
            except UnrecognizedImageError:
                try:
                    im = Image.open(io.BytesIO(content))
                    buf = io.BytesIO()
                    im.save(buf, format="PNG")
                    buf.name = "image.png"
                    buf.seek(0)
                    docx.add_picture(buf)
                except Exception:
                    logger.warning("Пропущена неподдерживаемая картинка")
    docx.save(out_path)

# ---------------- Telegram-хендлеры ----------------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Привет! Отправь PDF-файл, и я предложу варианты.")

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc or doc.mime_type != "application/pdf":
        return await update.message.reply_text("Нужен PDF-файл.")
    await update.message.reply_text("⏳ Скачиваю и обрабатываю…")
    tgfile = await doc.get_file()
    path = f"/tmp/{doc.file_unique_id}.pdf"
    await tgfile.download_to_drive(path)

    elems = extract_pdf_elements(path)
    context.user_data["pdf_path"] = path
    context.user_data["elems"] = elems

    kb = [
        [InlineKeyboardButton("Word: текст+картинки 📄", "cb_word")],
        [InlineKeyboardButton("TXT: текст 📄", "cb_txt")],
        [InlineKeyboardButton("Excel: таблицы 📊", "cb_tables")],
        [InlineKeyboardButton("Чат: текст 📝", "cb_text")],
        [InlineKeyboardButton("Чат: текст+картинки 🖼️📝", "cb_all")],
        [InlineKeyboardButton("Новый PDF 🔄", "cb_new")],
    ]
    await update.message.reply_text("Выбери вариант:", reply_markup=InlineKeyboardMarkup(kb))

# Прогресс‐бар при конвертации Word
async def cb_word(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    elems = context.user_data.get("elems", [])
    if not path or not elems:
        return await update.callback_query.edit_message_text("Сначала PDF.")

    msg = await context.bot.send_message(update.effective_chat.id, "⏳ Конвертация: 0%")
    start = asyncio.get_event_loop().time()

    async def progress_task():
        while True:
            pct = min(int((asyncio.get_event_loop().time() - start) / 4), 99)
            try:
                await context.bot.edit_message_text(
                    f"⏳ Конвертация: {pct}%", update.effective_chat.id, msg.message_id
                )
            except:
                pass
            if pct >= 99:
                break
            await asyncio.sleep(5)

    task = asyncio.create_task(progress_task())

    out = f"/tmp/{update.effective_user.id}_layout.docx"
    convert_to_word(elems, out)

    task.cancel()
    try:
        await context.bot.edit_message_text(
            "✅ Конвертация завершена!", update.effective_chat.id, msg.message_id
        )
    except:
        pass

    with open(out, "rb") as f:
        await context.bot.send_document(
            update.effective_chat.id, InputFile(f, filename="converted.docx")
        )

async def cb_txt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    elems = context.user_data.get("elems", [])
    if not elems:
        return await context.bot.send_message(update.effective_chat.id, "Сначала PDF.")
    out = f"/tmp/{update.effective_user.id}.txt"
    save_text_to_txt(elems, out)
    with open(out, "rb") as f:
        await context.bot.send_document(
            update.effective_chat.id, InputFile(f, filename="converted.txt")
        )

async def cb_tables(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await context.bot.send_message(update.effective_chat.id, "Сначала PDF.")
    tables = []
    with pdfplumber.open(path) as pdf:
        for p in pdf.pages:
            for tbl in p.extract_tables():
                if tbl and len(tbl) > 1:
                    df = pd.DataFrame(tbl[1:], columns=tbl[0])
                    tables.append(df)
    if not tables:
        return await context.bot.send_message(update.effective_chat.id, "Таблиц нет.")
    out = f"/tmp/{update.effective_user.id}_tables.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        for i, df in enumerate(tables, 1):
            df.to_excel(w, sheet_name=f"Таб{i}", index=False)
    with open(out, "rb") as f:
        await context.bot.send_document(
            update.effective_chat.id, InputFile(f, filename="tables.xlsx")
        )

async def cb_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    elems = context.user_data.get("elems", [])
    for typ, content in elems:
        if typ == "text":
            for i in range(0, len(content), 4096):
                await context.bot.send_message(update.effective_chat.id, content[i:i+4096])

async def cb_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    sent = set()
    elems = context.user_data.get("elems", [])
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
            bio.name = "img.png"
            await context.bot.send_photo(update.effective_chat.id, photo=bio)

async def cb_new(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    context.user_data.clear()
    await context.bot.send_message(update.effective_chat.id, "Пришлите новый PDF →")

# Регистрируем
app_builder = ApplicationBuilder().token(TOKEN).build()
app_builder.add_handler(CommandHandler("start", start))
app_builder.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
app_builder.add_handler(CallbackQueryHandler(cb_word,   pattern="cb_word"))
app_builder.add_handler(CallbackQueryHandler(cb_txt,    pattern="cb_txt"))
app_builder.add_handler(CallbackQueryHandler(cb_tables, pattern="cb_tables"))
app_builder.add_handler(CallbackQueryHandler(cb_text,   pattern="cb_text"))
app_builder.add_handler(CallbackQueryHandler(cb_all,    pattern="cb_all"))
app_builder.add_handler(CallbackQueryHandler(cb_new,    pattern="cb_new"))

if __name__ == "__main__":
    logger.info("Запускаем polling…")
    app_builder.run_polling()
