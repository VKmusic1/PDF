import os
import logging
import asyncio
from flask import Flask, request
from telegram import Update
from telegram.ext import Application, MessageHandler, ContextTypes, filters

TOKEN = os.getenv("TOKEN")
HOST = os.getenv("RENDER_EXTERNAL_HOSTNAME")
PORT = int(os.getenv("PORT", "10000"))
WEBHOOK_URL = f"https://{HOST}/{TOKEN}"

logging.basicConfig(level=logging.INFO)
app_flask = Flask(__name__)

telegram_app = Application.builder().token(TOKEN).build()
loop = asyncio.get_event_loop()

async def echo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("бот жив!")

telegram_app.add_handler(MessageHandler(filters.ALL, echo))

async def init_telegram():
    await telegram_app.initialize()

loop.run_until_complete(init_telegram())

@app_flask.route(f"/{TOKEN}", methods=["POST"])
def webhook():
    update = Update.de_json(request.get_json(force=True), telegram_app.bot)
    asyncio.run_coroutine_threadsafe(telegram_app.process_update(update), loop)
    return "ok"

@app_flask.route("/ping")
def ping():
    return "pong"

if __name__ == "__main__":
    import requests
    requests.post(
        f"https://api.telegram.org/bot{TOKEN}/setWebhook",
        data={"url": WEBHOOK_URL}
    )
    app_flask.run(host="0.0.0.0", port=PORT)
