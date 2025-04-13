"""
Telegram bot using aiogram 3.x to conduct polls and save results to an Excel file.

Functionality:
1. /start command — sends a welcome message and prints the chat ID (useful for setting GROUP_CHAT_ID).
2. /start_poll command — sends a non-anonymous poll to the specified group.
3. Poll answer handler — saves user_id, username, selected option, and timestamp to an Excel file.
4. /get_results command — sends the Excel file with poll results to the user.
5. All data is stored in-memory (Excel file inside BytesIO).

Requirements:
- Set TG_TOKEN and GROUP_CHAT_ID in the .env file.
- Install dependencies: aiogram, openpyxl, python-decouple.
"""
"""

1. Запускаем бота и отправляем команду /start:
   - Бот отвечает: "Привет! Chat ID: <ваш chat_id>".
   
2. Запускаем команду /start_poll для создания опроса:
   - Бот отправляет опрос в группу с вопросом "Какой язык программирования тебе ближе?" и вариантами ответа (Python, JavaScript, Rust, Go, Другое).
   - Бот подтверждает, что опрос отправлен в группу.

3. Участники группы отвечают на опрос:
   - Бот сохраняет их ответы и логирует выбор каждого пользователя с их username и временем.

4. Для получения результатов опроса отправляем команду /get_results:
   - Бот отправляет Excel-файл с результатами опроса (пользователь, выбранный ответ и временная метка).

Примечание: Чтобы бот корректно работал, нужно задать TG_TOKEN и GROUP_CHAT_ID в .env файле перед запуском.
"""


# --- Token and chat setup ---
import logging
import asyncio
from aiogram import Bot, Dispatcher, types
from aiogram.types import PollAnswer, InputFile
from aiogram.enums import ParseMode
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.client.default import DefaultBotProperties
from aiogram.filters import Command
from decouple import config
from openpyxl import Workbook
from io import BytesIO
from datetime import datetime

# --- Logging configuration ---
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# --- Configuration ---
TOKEN = config("TG_TOKEN")
GROUP_CHAT_ID = config("GROUP_CHAT_ID", cast=int)

bot = Bot(token=TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher(storage=MemoryStorage())

# --- Poll content ---
QUESTION = "Which programming language do you prefer?"
OPTIONS = ["Python", "JavaScript", "Rust", "Go", "Other"]

# --- Excel setup ---
workbook = Workbook()
sheet = workbook.active
sheet.title = "Poll Results"
sheet.append(["User ID", "Username", "Answer", "Timestamp"])
excel_stream = BytesIO()


@dp.message(Command("start_poll"))
async def start_poll(message: types.Message):
    logger.info("Received /start_poll command")
    await bot.send_poll(
        chat_id=GROUP_CHAT_ID,
        question=QUESTION,
        options=OPTIONS,
        is_anonymous=False
    )
    await message.reply("Poll has been sent to the group.")
    logger.info(f"Poll started by user {message.from_user.id} ({message.from_user.username})")


@dp.poll_answer()
async def handle_poll_answer(poll: PollAnswer):
    user_id = poll.user.id
    username = poll.user.username or "Unknown"
    answers = [OPTIONS[i] for i in poll.option_ids]
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for answer in answers:
        sheet.append([user_id, username, answer, timestamp])
        logger.info(f"Answer received: {username} chose '{answer}' at {timestamp}")


@dp.message(Command("start"))
async def start_cmd(message: types.Message):
    await message.reply(f"Hello! Chat ID: {message.chat.id}")
    logger.info(f"/start command from {message.from_user.username} (Chat ID: {message.chat.id})")


@dp.message(Command("get_results"))
async def get_results(message: types.Message):
    excel_stream.seek(0)
    workbook.save(excel_stream)
    excel_stream.seek(0)

    await bot.send_document(
        chat_id=message.chat.id,
        document=InputFile(excel_stream, filename="poll_results.xlsx"),
        caption="Poll results"
    )
    logger.info(f"Results sent to user {message.from_user.username}")


async def main():
    logger.info("Bot is starting...")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
