import logging
import asyncio
from aiogram import Bot, Dispatcher, types, F
from aiogram.types import FSInputFile 
from aiogram.enums import ParseMode
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.client.default import DefaultBotProperties
from aiogram.filters import Command
from decouple import config
from openpyxl import Workbook
from openpyxl import load_workbook
from io import BytesIO
import os
from datetime import datetime

# Logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Config
TOKEN = config("TG_TOKEN")
bot = Bot(token=TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher(storage=MemoryStorage())

# Runtime state
usernames = []
poll_id_to_user = {}
poll_id_to_data = {}
user_results = {}

# FSM-заглушка — без полноценной машины состояний
current_question = None
current_options = []

@dp.startup()
async def setup_commands(bot: Bot):
    commands = [
        types.BotCommand(command="start", description="Начать работу"),
        types.BotCommand(command="finish", description="Завершить и получить Excel"),
    ]
    await bot.set_my_commands(commands)

@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    await message.reply("👋 Привет! Пришли Excel-файл (.xlsx), где в первом столбце указаны Telegram id пользователей (можно получить через этого бота: @username_to_id_bot)")

@dp.message(Command("finish"))
async def finish(message: types.Message):
    if not user_results:
        await message.reply("❌ Пока нет результатов.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Results"
    ws.append(["Username", "Question", "Answer", "Timestamp"])

    for username, answers in user_results.items():
        for q, a, t in answers:
            ws.append([username, q, a, t])

    # Сохраняем результат во временный файл
    file_path = "/tmp/poll_results.xlsx"
    wb.save(file_path)

    # Используем FSInputFile
    file = FSInputFile(file_path, filename="poll_results.xlsx")
    await bot.send_document(
        chat_id=message.chat.id,
        document=file,
        caption="📊 Вот результаты!"
    )

    # Удаляем временный файл после отправки
    os.remove(file_path)

@dp.message(F.document)
async def handle_excel(message: types.Message):
    global usernames
    doc = message.document
    if not doc.file_name.endswith(".xlsx"):
        await message.reply("❌ Пожалуйста, пришли .xlsx файл.")
        return

    file = await bot.download(doc)
    wb = Workbook()
    wb = wb = Workbook(file) if hasattr(file, "read") else Workbook()
    wb = wb = load_workbook(file)
    ws = wb.active

    usernames = [str(row[0]).strip().lstrip("@") for row in ws.iter_rows(min_row=2, values_only=True) if row[0]]
    logger.info(f"Загружено {len(usernames)} username'ов: {usernames}")
    await message.reply(f"✅ Загружено {len(usernames)} username'ов.\nТеперь отправь вопрос в формате:\n\n<b>Вопрос?</b>\nВариант1\nВариант2\nВариант3", parse_mode="HTML")

@dp.message(F.text)
async def receive_poll_template(message: types.Message):
    global current_question, current_options

    if not usernames:
        await message.reply("⚠️ Сначала отправь Excel-файл с username'ами.")
        return

    lines = message.text.strip().split("\n")
    if len(lines) < 3:
        await message.reply("❌ Формат неверный. Нужно:\nВопрос?\nВариант1\nВариант2\n...")
        return

    question = lines[0].strip()
    if not question.endswith('?'):
        await message.reply("❌ Вопрос должен заканчиваться на '?'")
        return

    options = [line.strip() for line in lines[1:] if line.strip()]
    if len(options) < 2:
        await message.reply("❌ Нужно минимум два варианта ответа.")
        return

    current_question = question
    current_options = options

    await message.reply("📤 Рассылаю опросы...")

    success = 0
    failed = 0
    for username in usernames:
        try:
            chat = await bot.get_chat(username)
            poll = await bot.send_poll(
                chat_id=chat.id,
                question=current_question,
                options=current_options,
                is_anonymous=False
            )
            poll_id_to_user[poll.poll.id] = username
            poll_id_to_data[poll.poll.id] = (current_question, current_options)
            success += 1
        except Exception as e:
            logger.warning(f"Не удалось отправить опрос @{username}: {e}")
            failed += 1

    await message.reply(f"✅ Опрос отправлен {success} пользователям\n⚠️ Не удалось отправить: {failed}")


@dp.poll_answer()
async def handle_poll_answer(poll: types.PollAnswer):
    username = poll_id_to_user.get(poll.poll_id, "Unknown")
    question, options = poll_id_to_data.get(poll.poll_id, ("Unknown", []))
    answer = options[poll.option_ids[0]] if poll.option_ids else "Без ответа"
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if username not in user_results:
        user_results[username] = []

    user_results[username].append((question, answer, timestamp))
    logger.info(f"{username} выбрал '{answer}' на вопрос '{question}'")



async def main():
    logger.info("Bot is starting...")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
