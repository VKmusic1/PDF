# bot.py
import os
import io
import logging
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
PORT  = int(os.getenv("PORT","10000"))
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

# 5. Telegram Application
telegram_app = Application.builder().token(TOKEN).connection_pool_size(100).build()
telegram_app.request_kwargs = {"read_timeout":60,"connect_timeout":20}
future = asyncio.run_coroutine_threadsafe(telegram_app.initialize(), telegram_loop)
future.result(timeout=15)
logger.info("✔ Telegram initialized")

# 6. PDF утилиты
def extract_pdf_elements(path):
    doc = fitz.open(path)
    elems=[]
    for page in doc:
        t=page.get_text().strip()
        if t: elems.append(("text",t))
        for img in page.get_images(full=True):
            data=doc.extract_image(img[0])["image"]
            elems.append(("img",data))
    doc.close()
    return elems

def save_text_to_txt(elems,out):
    with open(out,"w",encoding="utf-8") as f:
        for typ,c in elems:
            if typ=="text": f.write(c+"\n\n")

def convert_pdf_to_docx(pdf_path,docx_path):
    cv=Converter(pdf_path)
    cv.convert(docx_path,start=0,end=None)
    cv.close()

# 7. Хендлеры
async def start(update,context):
    await update.message.reply_text("Привет! Пришли PDF-файл.")

async def handle_pdf(update,context):
    doc=update.message.document
    if not doc or doc.mime_type!="application/pdf":
        return await update.message.reply_text("Нужен PDF.")
    tgfile=await doc.get_file()
    path=f"/tmp/{doc.file_unique_id}.pdf"
    await tgfile.download_to_drive(path)
    elems=extract_pdf_elements(path)
    context.user_data["pdf_path"]=path
    context.user_data["elements"]=elems
    kb=[
      [InlineKeyboardButton("Word: макет 📄",callback_data="cb_word_all")],
      [InlineKeyboardButton("TXT: текст 📄",callback_data="cb_txt")],
      [InlineKeyboardButton("Excel: таблицы 📊",callback_data="cb_tables")],
      [InlineKeyboardButton("Чат: текст 📝",callback_data="cb_text_only")],
      [InlineKeyboardButton("Чат: текст+картинки 🖼️📝",callback_data="cb_chat_all")],
      [InlineKeyboardButton("Новый PDF 🔄",callback_data="cb_new_pdf")],
    ]
    await update.message.reply_text("Выбери:",reply_markup=InlineKeyboardMarkup(kb))

async def cb_text_only(update,context):
    await update.callback_query.answer()
    path=context.user_data.get("pdf_path")
    if not path:
        return await update.callback_query.edit_message_text("Сначала PDF.")
    for t in extract_pdf_elements(path):
        if t[0]=="text":
            for i in range(0,len(t[1]),4096):
                await context.bot.send_message(update.effective_chat.id,t[1][i:i+4096])
                await asyncio.sleep(0.05)

async def cb_chat_all(update,context):
    await update.callback_query.answer()
    path=context.user_data.get("pdf_path")
    sent=set()
    for typ,c in extract_pdf_elements(path):
        if typ=="text":
            for i in range(0,len(c),4096):
                await context.bot.send_message(update.effective_chat.id,c[i:i+4096])
        else:
            h=hash(c)
            if h in sent: continue
            sent.add(h)
            bio=io.BytesIO(c); bio.name="img.png"
            await context.bot.send_photo(update.effective_chat.id,photo=bio)
            await asyncio.sleep(0.1)

async def cb_word_all(update,context):
    await update.callback_query.answer()
    path=context.user_data.get("pdf_path")
    if not path:
        return await context.bot.send_message(update.effective_chat.id,"Сначала PDF")
    out=f"/tmp/{update.effective_user.id}_layout.docx"
    convert_pdf_to_docx(path,out)
    with open(out,"rb") as f:
        await context.bot.send_document(update.effective_chat.id,InputFile(f,filename="layout.docx"))

async def cb_txt(update,context):
    await update.callback_query.answer()
    elems=context.user_data.get("elements",[])
    out=f"/tmp/{update.effective_user.id}.txt"
    save_text_to_txt(elems,out)
    with open(out,"rb") as f:
        await context.bot.send_document(update.effective_chat.id,InputFile(f,filename="text.txt"))

async def cb_tables(update,context):
    await update.callback_query.answer()
    path=context.user_data.get("pdf_path")
    tbls=[]
    with pdfplumber.open(path) as pdf:
        for p,page in enumerate(pdf.pages,1):
            for ti,tbl in enumerate(page.extract_tables(),1):
                if not tbl or len(tbl)<2: continue
                df=pd.DataFrame(tbl[1:],columns=tbl[0])
                tbls.append((f"S{p}T{ti}"[:31],df))
    if not tbls:
        return await context.bot.send_message(update.effective_chat.id,"Нет таблиц")
    excel=f"/tmp/{update.effective_user.id}_tables.xlsx"
    with pd.ExcelWriter(excel,engine="openpyxl") as w:
        for nm,df in tbls:
            df.to_excel(w,sheet_name=nm,index=False)
    with open(excel,"rb") as f:
        await context.bot.send_document(update.effective_chat.id,InputFile(f,filename="tables.xlsx"))

async def cb_new_pdf(update,context):
    await update.callback_query.answer()
    context.user_data.clear()
    await context.bot.send_message(update.effective_chat.id,"Пришли новый PDF →")

# рег
telegram_app.add_handler(CommandHandler("start",start))
telegram_app.add_handler(MessageHandler(filters.Document.PDF,handle_pdf))
telegram_app.add_handler(CallbackQueryHandler(cb_text_only,pattern="cb_text_only"))
telegram_app.add_handler(CallbackQueryHandler(cb_chat_all,pattern="cb_chat_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_word_all,pattern="cb_word_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_txt,pattern="cb_txt"))
telegram_app.add_handler(CallbackQueryHandler(cb_tables,pattern="cb_tables"))
telegram_app.add_handler(CallbackQueryHandler(cb_new_pdf,pattern="cb_new_pdf"))

# ======================= 8. Flask-маршруты =======================
@app.route(f"/{TOKEN}",methods=["POST"])
def telegram_webhook():
    data=request.get_json(force=True)
    upd=Update.de_json(data,telegram_app.bot)
    asyncio.run_coroutine_threadsafe(telegram_app.process_update(upd),telegram_loop)
    return "OK"

@app.route("/ping",methods=["GET"])
def ping():
    return "pong"

# ======================= 9. Запуск =======================
if __name__=="__main__":
    logger.info(f"Setting webhook to {WEBHOOK_URL}")
    r=requests.post(f"https://api.telegram.org/bot{TOKEN}/setWebhook",data={"url":WEBHOOK_URL})
    if r.ok: logger.info("Webhook set ✓")
    else: logger.error("Webhook error: %s",r.text)
    logger.info("Running Flask on port %s",PORT)
    app.run(host="0.0.0.0",port=PORT)
