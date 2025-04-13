# Telegram polling bot

As in the case of aiogram 3.x, for using anonymous queries in the Telegram group and sending responses to an Excel file.

## 🚀 Features

- The /start_poll command — requests information in a given group.
- The answers are recorded in Excel: `User ID', `Username', `Response`, `Timestamp'.
- The "/get_results" command — manages an Excel file with ratings.
- The "/start" command — displays the "chat ID" for configuration.

## 🛠️ Installation

1.Clone the repository:

2.Install the dependencies:

```
make setup
```

3.Setup envs (.env)

```
TG\*TOKEN=vah*token*bot
GROUP_CHAT_ID=-1001234567890
```

📦Dependencies

- aiogram>=3.0.0

- openpyxl

- python decoupling

▶️Launch

```
make run
```

📌 Notes

> Make sure that the bot has been added to the group and is an admin. You can find out the GROUP_ID by joining the team/starting work in the group.
