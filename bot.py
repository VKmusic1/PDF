import os
import logging
import re
import fitz
import asyncio
from aiohttp import web
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    filters,
    ContextTypes,
)
from PyPDF2 import PdfReader

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! Пришли мне PDF, и я спрошу что тебе нужно: только текст или текст с картинками."
    )

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc_file = update.message.document
    if doc_file.mime_type != "application/pdf":
        await update.message.reply_text("Это не PDF.")
        return

    file_path = f"/tmp/{doc_file.file_name}"
    new_file = await context.bot.get_file(doc_file.file_id)
    await new_file.download_to_drive(file_path)
    context.user_data['pdf_path'] = file_path

    # Спрашиваем, что нужно делать
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("Только текст", callback_data="only_text")],
        [InlineKeyboardButton("Текст и изображения", callback_data="text_images")],
    ])
    await update.message.reply_text(
        "Что вы хотите получить из файла?",
        reply_markup=keyboard
    )

async def only_text_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    file_path = context.user_data.get('pdf_path')
    if not file_path:
        return await query.edit_message_text("Файл не найден. Пришлите PDF заново.")

    await query.edit_message_text("⏳ Извлекаю только текст...")

    # Извлекаем только текст
    try:
        reader = PdfReader(file_path)
        text = ""
        for page in reader.pages:
            page_text = page.extract_text() or ""
            page_text = re.sub(r"(\w)-\n(\w)", r"\1\2", page_text)
            page_text = page_text.replace("\n", " ")
            page_text = re.sub(r" {2,}", " ", page_text).strip()
            text += page_text + "\n"
        text = text.strip()
    except Exception as e:
        return await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text=f"Ошибка при чтении PDF: {e}"
        )

    if not text:
        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text="Не удалось извлечь текст."
        )
        return

    # Отправляем текст порциями
    for i in range(0, len(text), 4096):
        await context.bot.send_message(chat_id=update.effective_chat.id, text=text[i:i+4096])

    # Кнопка загрузить еще
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("🔄 Загрузить ещё PDF-файл", callback_data="start_over")],
    ])
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text="Если нужно обработать ещё PDF — отправьте новый файл или нажмите кнопку ниже.",
        reply_markup=keyboard
    )

async def text_images_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    file_path = context.user_data.get('pdf_path')
    if not file_path:
        return await query.edit_message_text("Файл не найден. Пришлите PDF заново.")

    await query.edit_message_text("⏳ Извлекаю текст и изображения...")

    # Извлекаем текст и картинки в порядке по страницам
    import tempfile
    reader = PdfReader(file_path)
    pdf_doc = fitz.open(file_path)
    sent_hashes = set()
    num_pages = len(pdf_doc)
    found_content = False

    for i in range(num_pages):
        # текст
        page_text = ""
        try:
            page_text = reader.pages[i].extract_text() or ""
            page_text = re.sub(r"(\w)-\n(\w)", r"\1\2", page_text)
            page_text = page_text.replace("\n", " ")
            page_text = re.sub(r" {2,}", " ", page_text).strip()
        except Exception:
            pass
        if page_text:
            found_content = True
            for j in range(0, len(page_text), 4096):
                await context.bot.send_message(chat_id=update.effective_chat.id, text=page_text[j:j+4096])
        # картинки
        for img in pdf_doc[i].get_images(full=True):
            xref = img[0]
            img_dict = pdf_doc.extract_image(xref)
            img_bytes = img_dict['image']
            ext = img_dict['ext']
            img_hash = hash(img_bytes)
            if img_hash not in sent_hashes:
                found_content = True
                with tempfile.NamedTemporaryFile(delete=False, suffix='.' + ext) as tmp_img:
                    tmp_img.write(img_bytes)
                    tmp_img.flush()
                    await context.bot.send_photo(
                        chat_id=update.effective_chat.id,
                        photo=tmp_img.name
                    )
                sent_hashes.add(img_hash)

    if not found_content:
        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text="В PDF не найдено текста или изображений."
        )

    # Кнопка загрузить еще
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("🔄 Загрузить ещё PDF-файл", callback_data="start_over")],
    ])
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text="Если нужно обработать ещё PDF — отправьте новый файл или нажмите кнопку ниже.",
        reply_markup=keyboard
    )

async def start_over_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    context.user_data.clear()
    await query.edit_message_reply_markup(None)
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text="🔄 Пришлите новый PDF-файл сюда."
    )

# ----------- ПИНГ-ПОНГ ----------
async def ping(request):
    return web.Response(text="pong")

async def run_ping_server():
    app = web.Application()
    app.router.add_get('/ping', ping)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", 8080)
    await site.start()
    print("Ping server running on /ping")
    while True:
        await asyncio.sleep(3600)

# ----------- MAIN -----------
def main():
    token = os.getenv('TELEGRAM_BOT_TOKEN')
    if not token:
        logger.error('TELEGRAM_BOT_TOKEN не задан')
        return

    app = (
        ApplicationBuilder()
        .token(token)
        .build()
    )
    app.add_handler(CommandHandler('start', start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_pdf))
    app.add_handler(CallbackQueryHandler(only_text_callback, pattern='only_text'))
    app.add_handler(CallbackQueryHandler(text_images_callback, pattern='text_images'))
    app.add_handler(CallbackQueryHandler(start_over_callback, pattern='start_over'))

    host = os.getenv('RENDER_EXTERNAL_URL')
    if not host:
        logger.error('RENDER_EXTERNAL_URL не задан')
        return
    port = int(os.getenv('PORT', 5000))
    webhook_url = f"{host}/{token}"

    loop = asyncio.get_event_loop()
    loop.create_task(run_ping_server())
    app.run_webhook(
        listen='0.0.0.0',
        port=port,
        url_path=token,
        webhook_url=webhook_url
    )

if __name__ == '__main__':
    main()
