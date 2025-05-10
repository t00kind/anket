# --- 🐍 Local Python commands ---

create-venv:
	python3 -m venv venv

install-deps:
	source venv/bin/activate && pip install -r requirements.txt

setup: create-venv install-deps

run:
	python3 main.py

run-absolute:
	/Library/Frameworks/Python.framework/Versions/3.12/bin/python3 main.py

serve:
	nohup python3 main.py > bot.log 2>&1

log:
	tail -f bot.log

lint:
	ruff check . --fix


# --- 🐳 Docker commands ---

docker-build:
	docker build -t poll-bot .

docker-run:
	docker run --env-file .env poll-bot

docker-run-interactive:
	docker run --rm -it --env-file .env poll-bot

docker-clean:
	-docker rmi poll-bot
