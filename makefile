# --- 🐍 local ---
venv:
	python -m venv venv

setup:
	pip install -r requirements.txt

run:
	python3 main.py

run-absolute:
	/Library/Frameworks/Python.framework/Versions/3.12/bin/python3 main.py

lint:
	ruff check . --fix


# --- 🐳 Docker ---
docker-build:
	docker build -t poll-bot .

docker-run:
	docker run --env-file .env poll-bot

docker-run-interactive:
	docker run --rm -it --env-file .env poll-bot

docker-clean:
	docker rmi poll-bot || true
