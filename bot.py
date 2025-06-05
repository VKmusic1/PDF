# -------------------------  requirements.txt  --------------------------
# Flask==2.2.5
# numpy==1.23.5
# pandas==2.0.3
# openpyxl==3.1.2
# PyMuPDF==1.22.5
# pdfplumber==0.7.5
# python-docx==1.1.2
# python-telegram-bot[webhooks]==20.3
# requests==2.31.0
# ----------------------------------------------------------------------

import os, io, logging, asyncio, fitz, pdfplumber, pandas as pd
from flask import Flask, request
from telegram import Update, InputFile, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler,
    ContextTypes, filters
)
from docx import Document

# -------------------- 1. конфиг --------------------
TOKEN = os.getenv("TOKEN")
HOST  = os.getenv("RENDER_EXTERNAL_HOSTNAME")
PORT  = int(os.getenv("PORT", "10000"))
if not TOKEN or not HOST:
    raise RuntimeError("Нужны переменные TOKEN и RENDER_EXTERNAL_HOSTNAME")
WEBHOOK_URL = f"https://{HOST}/{TOKEN}"

# -------------------- 2. лог -----------------------
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

# -------------------- 3. Telegram-app --------------
telegram_app = (
    Application.builder()
    .token(TOKEN)
    .connection_pool_size(100)
    .build()
)
telegram_app.request_kwargs = {"read_timeout": 60, "connect_timeout": 20}

# -------------------- 4. PDF helpers ---------------
def extract_pdf_elements(path:str):
    doc = fitz.open(path)
    out = []
    for page in doc:
        txt = page.get_text().strip()
        if txt:
            out.append(("text", txt))
        for img in page.get_images(full=True):
            out.append(("img", doc.extract_image(img[0])["image"]))
    doc.close()
    return out

def convert_to_word(elements, dst):
    docx = Document()
    for t, c in elements:
        if t == "text":
            docx.add_paragraph(c)
        else:
            bio = io.BytesIO(c); bio.name = "img.png"
            docx.add_picture(bio)
    docx.save(dst)

# -------------------- 5. клавиатура ----------------
def main_kb():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("Word: текст+картинки 📄",    callback_data="cb_word_all")],
        [InlineKeyboardButton("TXT: текст 📄",              callback_data="cb_txt")],
        [InlineKeyboardButton("Excel: таблицы 📊",          callback_data="cb_tables")],
        [InlineKeyboardButton("Чат: текст 📝",              callback_data="cb_text_only")],
        [InlineKeyboardButton("Чат: текст+картинки 🖼️📝",   callback_data="cb_chat_all")],
        [InlineKeyboardButton("Новый PDF 🔄",               callback_data="cb_new_pdf")]
    ])

# -------------------- 6. handlers ------------------
async def start(u:Update, c:ContextTypes.DEFAULT_TYPE):
    await u.message.reply_text("Привет! Отправь PDF-файл.")

async def handle_pdf(u:Update, c:ContextTypes.DEFAULT_TYPE):
    doc = u.message.document
    if not doc or doc.mime_type != "application/pdf":
        return await u.message.reply_text("Это не PDF.")
    path = f"/tmp/{doc.file_unique_id}.pdf"
    await (await doc.get_file()).download_to_drive(path)
    c.user_data["pdf_path"] = path
    await u.message.reply_text("Выбери действие:", reply_markup=main_kb())

# ---------- callback: только текст в чат ----------
async def cb_text_only(u:Update, c:ContextTypes.DEFAULT_TYPE):
    await u.callback_query.answer()
    path = c.user_data.get("pdf_path")
    if not path:
        return await u.callback_query.edit_message_text("Файл не найден.")
    parts = [txt for typ,txt in extract_pdf_elements(path) if typ=="text"]
    for block in parts:
        for i in range(0, len(block), 4096):
            await c.bot.send_message(u.effective_chat.id, block[i:i+4096])
            await asyncio.sleep(0.05)

# ---------- callback: текст+картинки в чат ----------
async def cb_chat_all(u:Update, c:ContextTypes.DEFAULT_TYPE):
    await u.callback_query.answer()
    path = c.user_data.get("pdf_path")
    if not path:
        return await u.callback_query.edit_message_text("Файл не найден.")
    sent = set()
    for typ, cnt in extract_pdf_elements(path):
        if typ=="text":
            for i in range(0,len(cnt),4096):
                await c.bot.send_message(u.effective_chat.id,cnt[i:i+4096])
        else:
            h=hash(cnt)
            if h in sent: continue
            sent.add(h)
            bio = io.BytesIO(cnt); bio.name="img.png"
            await c.bot.send_photo(u.effective_chat.id, photo=bio)

# ---------- callback: Word (текст+картинки) ----------
async def cb_word_all(u:Update, c:ContextTypes.DEFAULT_TYPE):
    await u.callback_query.answer()
    path = c.user_data.get("pdf_path")
    if not path:
        return await u.callback_query.edit_message_text("Файл не найден.")
    elems = extract_pdf_elements(path)
    if not elems:
        return await u.callback_query.edit_message_text("Контент не найден.")
    dst=f"/tmp/{u.effective_user.id}_all.docx"
    convert_to_word(elems,dst)
    with open(dst,"rb") as f:
        await c.bot.send_document(u.effective_chat.id, InputFile(f,"converted.docx"))

# ---------- callback: TXT ----------
async def cb_txt(u:Update, c:ContextTypes.DEFAULT_TYPE):
    await u.callback_query.answer()
    path = c.user_data.get("pdf_path")
    if not path:
        return await u.callback_query.edit_message_text("Файл не найден.")
    parts=[t for typ,t in extract_pdf_elements(path) if typ=="text"]
    if not parts:
        return await u.callback_query.edit_message_text("В PDF нет текста.")
    dst=f"/tmp/{u.effective_user.id}.txt"
    with open(dst,"w",encoding="utf-8") as f: f.write("\n\n".join(parts))
    with open(dst,"rb") as f:
        await c.bot.send_document(u.effective_chat.id, InputFile(f,"converted.txt"))

# ---------- callback: таблицы ----------
async def cb_tables(u:Update, c:ContextTypes.DEFAULT_TYPE):
    await u.callback_query.answer()
    path=c.user_data.get("pdf_path")
    if not path:
        return await u.callback_query.edit_message_text("Файл не найден.")
    all_tbl=[]
    with pdfplumber.open(path) as pdf:
        for p_num,page in enumerate(pdf.pages,1):
            for t_num, tbl in enumerate(page.extract_tables(),1):
                if not tbl or len(tbl)<2: continue
                df=pd.DataFrame(tbl[1:],columns=tbl[0])
                all_tbl.append((f"S{p_num}_T{t_num}",df))
    if not all_tbl:
        return await u.callback_query.edit_message_text("Таблиц не найдено.")
    dst=f"/tmp/{u.effective_user.id}_tables.xlsx"
    with pd.ExcelWriter(dst,engine="openpyxl") as w:
        for sheet,df in all_tbl: df.to_excel(w,sheet_name=sheet[:31],index=False)
    with open(dst,"rb") as f:
        await c.bot.send_document(u.effective_chat.id, InputFile(f,"tables.xlsx"))

# ---------- callback: новый PDF ----------
async def cb_new_pdf(u:Update, c:ContextTypes.DEFAULT_TYPE):
    await u.callback_query.answer()
    c.user_data.clear()
    await c.bot.send_message(u.effective_chat.id, "Отправь новый PDF-файл.")

# -------------------- 7. регистрация --------------
telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
telegram_app.add_handler(CallbackQueryHandler(cb_text_only,  pattern="cb_text_only"))
telegram_app.add_handler(CallbackQueryHandler(cb_chat_all,   pattern="cb_chat_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_word_all,   pattern="cb_word_all"))
telegram_app.add_handler(CallbackQueryHandler(cb_txt,        pattern="cb_txt"))
telegram_app.add_handler(CallbackQueryHandler(cb_tables,     pattern="cb_tables"))
telegram_app.add_handler(CallbackQueryHandler(cb_new_pdf,    pattern="cb_new_pdf"))

# -------------------- 8. Flask + webhook + ping ----
app = Flask(__name__)
loop = asyncio.new_event_loop(); asyncio.set_event_loop(loop)
loop.run_until_complete(telegram_app.initialize())

@app.route("/ping")
def ping(): return "pong"

@app.route(f"/{TOKEN}", methods=["POST"])
def telegram_webhook():
    update = Update.de_json(request.get_json(force=True), telegram_app.bot)
    loop.create_task(telegram_app.process_update(update))
    return "ok"

# -------------------- 9. run -----------------------
if __name__ == "__main__":
    import requests
    requests.post(f"https://api.telegram.org/bot{TOKEN}/setWebhook", data={"url": WEBHOOK_URL})
    logger.info("Запускаем Flask на порту %s", PORT)
    app.run(host="0.0.0.0", port=PORT)
