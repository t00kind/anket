import logging
import asyncio
from aiogram import Bot, Dispatcher, types, F
from aiogram.enums import ParseMode
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.client.default import DefaultBotProperties
from aiogram.filters import Command
from decouple import config
from openpyxl import Workbook, load_workbook
from io import BytesIO
from datetime import datetime

"""
Telegram-бот для проведения опросов по списку пользователей из Excel-файла и сохранения результатов в Excel.

Функциональность:
---------------
1. Пользователь запускает бота командой /start:
   - Бот приветствует и просит прислать Excel-файл с Telegram-username'ами (без символа @).

2. Пользователь отправляет Excel-файл (.xlsx), где в первом столбце указаны username'ы:
   - Бот загружает файл.
   - Считывает username'ы из первой колонки таблицы, начиная со второй строки (предполагается, что первая — заголовок).
   - По каждому username отправляется личное сообщение с опросом.
   - Опрос — неанонимный, с вопросом "Какой язык программирования тебе ближе?" и вариантами ответа.

3. Пользователь, получивший опрос, выбирает один или несколько вариантов:
   - Бот ловит событие ответа на опрос.
   - Сохраняет username (или user_id), выбранный вариант и временную метку в Excel-файл в памяти (BytesIO).

4. Пользователь отправляет команду /get_results:
   - Бот генерирует Excel-файл с результатами.
   - Отправляет Excel-файл в чат.

Архитектура:
-----------
- Aiogram 3.x используется как асинхронный фреймворк для Telegram-бота.
- Все результаты хранятся в оперативной памяти в формате Excel с помощью библиотеки openpyxl.
- Переменные TG_TOKEN (токен бота) и другие данные берутся из файла .env через python-decouple.
- Для сопоставления ответов с пользователями используется маппинг poll_id → username.

Ограничения:
-----------
- Бот работает только с .xlsx-файлами (не .csv).
- Telegram username'ы должны быть публичны и существовать, иначе бот не сможет отправить сообщение.
- Ответы хранятся до перезапуска бота (данные не сохраняются в файл или базу данных автоматически).
"""


# Logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Config
TOKEN = config("TG_TOKEN")
bot = Bot(token=TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher(storage=MemoryStorage())

# Excel results setup
result_wb = Workbook()
result_ws = result_wb.active
result_ws.title = "Poll Results"
result_ws.append(["Username", "Answer", "Timestamp"])
excel_stream = BytesIO()

# Poll setup
QUESTION = "Which programming language do you prefer?"
OPTIONS = ["Python", "JavaScript", "Rust", "Go", "Other"]
poll_id_to_user = {}

@dp.message(Command("start"))
async def start_cmd(message: types.Message):
    await message.reply("👋 Hello! Send me an Excel file (.xlsx) with Telegram usernames.")
    logger.info(f"/start from {message.from_user.id} ({message.from_user.username})")

@dp.message(F.document)
async def handle_excel_file(message: types.Message):
    doc = message.document
    if not doc.file_name.endswith(".xlsx"):
        await message.reply("❌ Please send a valid .xlsx file.")
        return

    file = await bot.download(doc)
    wb = load_workbook(file)
    ws = wb.active

    usernames = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        username = row[0]
        if username:
            usernames.append(username.strip().lstrip("@"))

    await message.reply(f"📤 Sending poll to {len(usernames)} users...")

    for username in usernames:
        try:
            user = await bot.get_chat(username)
            sent_poll = await bot.send_poll(
                chat_id=user.id,
                question=QUESTION,
                options=OPTIONS,
                is_anonymous=False
            )
            poll_id_to_user[sent_poll.poll.id] = username
            logger.info(f"Poll sent to {username}")
        except Exception as e:
            logger.warning(f"Failed to send poll to @{username}: {e}")

@dp.poll_answer()
async def handle_poll_answer(poll: types.PollAnswer):
    username = poll_id_to_user.get(poll.poll_id, "Unknown")
    answer = OPTIONS[poll.option_ids[0]] if poll.option_ids else "No answer"
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    result_ws.append([username, answer, timestamp])
    logger.info(f"{username} answered '{answer}'")

@dp.message(Command("get_results"))
async def get_results(message: types.Message):
    excel_stream.seek(0)
    result_wb.save(excel_stream)
    excel_stream.seek(0)

    await bot.send_document(
        chat_id=message.chat.id,
        document=types.InputFile(excel_stream, filename="poll_results.xlsx"),
        caption="📊 Here are the results!"
    )

async def main():
    logger.info("Bot is starting...")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
