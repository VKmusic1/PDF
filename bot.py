import os
import io
import logging
import time
import asyncio
import threading

import fitz
import pdfplumber
import pandas as pd
from flask import Flask, request
from telegram import Update, InputFile, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters
)
from docx import Document
from pdf2docx import Converter
import requests

# 1. Конфиг
TOKEN = os.getenv("TOKEN")
HOST  = os.getenv("RENDER_EXTERNAL_HOSTNAME")
PORT  = int(os.getenv("PORT", "10000"))
if not TOKEN or not HOST:
    raise RuntimeError("TOKEN и RENDER_EXTERNAL_HOSTNAME обязательны")
WEBHOOK_URL = f"https://{HOST}/{TOKEN}"

# 2. Лог
logging.basicConfig(format="%(asctime)s %(levelname)s: %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# 3. Flask
app = Flask(__name__)

# 4. asyncio-loop
telegram_loop = asyncio.new_event_loop()
def start_loop(loop):
    asyncio.set_event_loop(loop)
    loop.run_forever()
threading.Thread(target=start_loop, args=(telegram_loop,), daemon=True).start()

# 5. Telegram-app
telegram_app = (
    Application.builder()
    .token(TOKEN)
    .connection_pool_size(100)
    .build()
)
telegram_app.request_kwargs = {"read_timeout":60,"connect_timeout":20}
future = asyncio.run_coroutine_threadsafe(telegram_app.initialize(), telegram_loop)
future.result(timeout=15)
logger.info("✔ Telegram initialized")

# 6. PDF-утилиты
def extract_pdf_elements(path):
    doc = fitz.open(path)
    elems=[]
    for page in doc:
        txt=page.get_text().strip()
        if txt: elems.append(("text",txt))
        for img in page.get_images(full=True):
            data=doc.extract_image(img[0])["image"]
            elems.append(("img",data))
    doc.close()
    return elems

def save_txt(elems,out):
    with open(out,"w",encoding="utf-8") as f:
        for t,c in elems:
            if t=="text": f.write(c+"\n\n")

def convert_pdf_to_docx(pdf_path, docx_path):
    cv=Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()

# 7. Handlers
async def start(u, c):
    await u.message.reply_text("Привет! Отправь PDF-файл.")

async def handle_pdf(u, c):
    doc=u.message.document
    if not doc or doc.mime_type!="application/pdf":
        return await u.message.reply_text("Нужен PDF.")
    tgfile=await doc.get_file()
    path=f"/tmp/{doc.file_unique_id}.pdf"
    await tgfile.download_to_drive(path)
    elems=extract_pdf_elements(path)
    c.user_data["pdf_path"]=path
    c.user_data["elements"]=elems
    kb=[
      [InlineKeyboardButton("Word: макет 📄",callback_data="cb_word_all")],
      [InlineKeyboardButton("TXT: текст 📄",callback_data="cb_txt")],
      [InlineKeyboardButton("Excel: таблицы 📊",callback_data="cb_tables")],
      [InlineKeyboardButton("Чат: текст 📝",callback_data="cb_text_only")],
      [InlineKeyboardButton("Чат: текст+картинки 🖼️📝",callback_data="cb_chat_all")],
      [InlineKeyboardButton("Новый PDF 🔄",callback_data="cb_new_pdf")],
    ]
    await u.message.reply_text("Выбери:", reply_markup=InlineKeyboardMarkup(kb))

# остальные cb_* без изменений ...

async def cb_word_all(u, c):
    await u.callback_query.answer()
    path=c.user_data.get("pdf_path")
    if not path:
        return await c.bot.send_message(u.effective_chat.id,"Сначала PDF")
    # 1) отправляем сообщение с 0%
    msg = await c.bot.send_message(u.effective_chat.id, "⏳ Конвертация: 0%")
    start_ts = time.time()
    # 2) запускаем прогресс-обновления
    async def updater():
        while True:
            elapsed = time.time()-start_ts
            pct = min(int(elapsed/240*100), 99)
            try:
                await c.bot.edit_message_text(f"⏳ Конвертация: {pct}%", u.effective_chat.id, msg.message_id)
            except:
                pass
            if pct>=99:
                break
            await asyncio.sleep(15)
    task = asyncio.create_task(updater())
    # 3) сам конверт
    out = f"/tmp/{u.effective_user.id}_layout.docx"
    convert_pdf_to_docx(path, out)
    # 4) отменяем прогресс
    task.cancel()
    # 5) редактируем финальное сообщение
    try:
        await c.bot.edit_message_text("✅ Конвертация завершена!", u.effective_chat.id, msg.message_id)
    except:
        pass
    # 6) отправляем файл
    with open(out,"rb") as f:
        await c.bot.send_document(u.effective_chat.id, InputFile(f,filename="layout.docx"))

# ... остальные cb_txt, cb_tables, cb_text_only, cb_chat_all, cb_new_pdf ...

# регистрация
telegram_app.add_handler(CommandHandler("start",start))
telegram_app.add_handler(MessageHandler(filters.Document.PDF,handle_pdf))
telegram_app.add_handler(CallbackQueryHandler(cb_word_all,pattern="cb_word_all"))
# и т.д. для остальных

# 8. Flask‐routes
@app.route(f"/{TOKEN}",methods=["POST"])
def webhook():
    data=request.get_json(force=True)
    upd=Update.de_json(data,telegram_app.bot)
    asyncio.run_coroutine_threadsafe(telegram_app.process_update(upd), telegram_loop)
    return "OK"

@app.route("/ping")
def ping():
    return "pong"

# 9. Запуск
if __name__=="__main__":
    logger.info(f"Setting webhook to {WEBHOOK_URL}")
    r=requests.post(f"https://api.telegram.org/bot{TOKEN}/setWebhook",data={"url":WEBHOOK_URL})
    logger.info("Webhook set" if r.ok else f"Webhook error: {r.text}")
    logger.info("Run Flask on port %s",PORT)
    app.run(host="0.0.0.0",port=PORT)
