import os
import logging
import asyncio
from flask import Flask, request
from telegram import Update
from telegram.ext import Application, CommandHandler, ContextTypes

TOKEN = os.getenv("TOKEN")
HOST = os.getenv("RENDER_EXTERNAL_HOSTNAME")
PORT = int(os.getenv("PORT", "10000"))
WEBHOOK_URL = f"https://{HOST}/{TOKEN}"

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
app_flask = Flask(__name__)

telegram_app = (
    Application.builder()
    .token(TOKEN)
    .connection_pool_size(10)
    .build()
)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Бот жив! /start")

telegram_app.add_handler(CommandHandler("start", start))

loop = asyncio.get_event_loop()
async def init_telegram():
    await telegram_app.initialize()
loop.run_until_complete(init_telegram())

@app_flask.route(f"/{TOKEN}", methods=["POST"])
def webhook():
    update = Update.de_json(request.get_json(force=True), telegram_app.bot)
    asyncio.run_coroutine_threadsafe(telegram_app.process_update(update), loop)
    return "ok"   # <-- Telegram должен сразу получить "ok"!

@app_flask.route("/ping")
def ping():
    return "pong"

if __name__ == "__main__":
    import requests
    requests.post(
        f"https://api.telegram.org/bot{TOKEN}/setWebhook",
        data={"url": WEBHOOK_URL}
    )
    logger.info("Запускаем Flask на порту %s", PORT)
    app_flask.run(host="0.0.0.0", port=PORT)
