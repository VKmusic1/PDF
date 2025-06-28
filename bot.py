# bot.py
import os
import io
import logging
import threading

import fitz                  # PyMuPDF
import pdfplumber
import pandas as pd
from docx import Document
from flask import Flask
from telegram import (
    Update,
    InputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
)
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters,
)

# ---------------------- config ----------------------
TOKEN = os.getenv("TOKEN")
if not TOKEN:
    raise RuntimeError("Environment variable TOKEN is required")

PORT = int(os.getenv("PORT", "10000"))

# ---------------------- logging ----------------------
logging.basicConfig(
    format="%(asctime)s %(levelname)s: %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# ---------------------- healthcheck ----------------------
app_flask = Flask(__name__)

@app_flask.route("/ping")
def ping():
    return "pong"

def run_flask():
    # Запускаем Flask в отдельном потоке
    app_flask.run(host="0.0.0.0", port=PORT)

# ---------------------- PDF utils ----------------------
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

def save_elements_to_txt(elements, out_path: str):
    with open(out_path, "w", encoding="utf-8") as f:
        for typ, content in elements:
            if typ == "text":
                f.write(content + "\n\n")

def extract_tables_to_excel(path: str, out_path: str):
    all_tables = []
    with pdfplumber.open(path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            for tbl_idx, table in enumerate(tables, start=1):
                if not table or len(table) < 2:
                    continue
                df = pd.DataFrame(table[1:], columns=table[0])
                sheet = f"С{page_number}_T{tbl_idx}"
                all_tables.append((sheet[:31], df))
    if not all_tables:
        return False
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for sheet, df in all_tables:
            df.to_excel(writer, sheet_name=sheet, index=False)
    return True

# ---------------------- keyboard ----------------------
def main_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("Word: текст+картинки 📄", callback_data="word_all")],
        [InlineKeyboardButton("TXT: текст 📄",            callback_data="txt")],
        [InlineKeyboardButton("Excel: таблицы 📊",        callback_data="tables")],
        [InlineKeyboardButton("Чат: текст 📝",            callback_data="chat_text")],
        [InlineKeyboardButton("Чат: текст+картинки 🖼️📝", callback_data="chat_all")],
        [InlineKeyboardButton("Новый PDF 🔄",             callback_data="new_pdf")],
    ])

# ---------------------- handlers ----------------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Привет! Отправь PDF-файл.")

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc or doc.mime_type != "application/pdf":
        return await update.message.reply_text("Пожалуйста, отправь PDF.")
    logger.info("Получен PDF от %s", update.effective_user.id)

    f = await doc.get_file()
    path = f"/tmp/{doc.file_unique_id}.pdf"
    await f.download_to_drive(path)

    context.user_data["pdf_path"] = path
    await update.message.reply_text(
        "Выбери действие:",
        reply_markup=main_keyboard()
    )

async def cb_word_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await update.callback_query.edit_message_text("Сначала отправь PDF.")
    elems = extract_pdf_elements(path)
    out = f"/tmp/{update.effective_user.id}_all.docx"
    convert_to_word(elems, out)
    with open(out, "rb") as f:
        await context.bot.send_document(
            update.effective_chat.id,
            document=InputFile(f, filename="converted.docx")
        )
    await update.callback_query.edit_message_text("Готово! 📄", reply_markup=main_keyboard())

async def cb_txt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await update.callback_query.edit_message_text("Сначала отправь PDF.")
    elems = extract_pdf_elements(path)
    out = f"/tmp/{update.effective_user.id}.txt"
    save_elements_to_txt(elems, out)
    with open(out, "rb") as f:
        await context.bot.send_document(
            update.effective_chat.id,
            document=InputFile(f, filename="converted.txt")
        )
    await update.callback_query.edit_message_text("Готово! 📄", reply_markup=main_keyboard())

async def cb_tables(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await update.callback_query.edit_message_text("Сначала отправь PDF.")
    out = f"/tmp/{update.effective_user.id}_tables.xlsx"
    ok = extract_tables_to_excel(path, out)
    if ok:
        with open(out, "rb") as f:
            await context.bot.send_document(
                update.effective_chat.id,
                document=InputFile(f, filename="tables.xlsx")
            )
        await update.callback_query.edit_message_text("Таблицы готовы! 📊", reply_markup=main_keyboard())
    else:
        await update.callback_query.edit_message_text("Таблиц не найдено.", reply_markup=main_keyboard())

async def cb_chat_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await context.bot.send_message(update.effective_chat.id, "Сначала отправь PDF.")
    elems = extract_pdf_elements(path)
    for typ, content in elems:
        if typ == "text":
            for i in range(0, len(content), 4096):
                await context.bot.send_message(update.effective_chat.id, content[i:i+4096])
    await update.callback_query.edit_message_text("Готово! 📝", reply_markup=main_keyboard())

async def cb_chat_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    path = context.user_data.get("pdf_path")
    if not path:
        return await context.bot.send_message(update.effective_chat.id, "Сначала отправь PDF.")
    elems = extract_pdf_elements(path)
    sent = set()
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
    await update.callback_query.edit_message_text("Готово! 🖼️📝", reply_markup=main_keyboard())

async def cb_new_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    context.user_data.clear()
    await context.bot.send_message(update.effective_chat.id, "Отправь новый PDF.")

# ---------------------- setup bot ----------------------
bot_app = (
    Application.builder()
    .token(TOKEN)
    .connection_pool_size(100)
    .build()
)

bot_app.add_handler(CommandHandler("start", start))
bot_app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
bot_app.add_handler(CallbackQueryHandler(cb_word_all,    pattern="word_all"))
bot_app.add_handler(CallbackQueryHandler(cb_txt,         pattern="txt"))
bot_app.add_handler(CallbackQueryHandler(cb_tables,      pattern="tables"))
bot_app.add_handler(CallbackQueryHandler(cb_chat_text,   pattern="chat_text"))
bot_app.add_handler(CallbackQueryHandler(cb_chat_all,    pattern="chat_all"))
bot_app.add_handler(CallbackQueryHandler(cb_new_pdf,     pattern="new_pdf"))

# ---------------------- main ----------------------
if __name__ == "__main__":
    logger.info("Запускаем Flask (healthcheck) и Telegram (polling)…")
    # Flask для /ping
    threading.Thread(target=run_flask, daemon=True).start()
    # Telegram-бот
    bot_app.run_polling(drop_pending_updates=True)
