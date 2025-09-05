# Schedule Excel Telegram Bot

Parses Ukrainian work-schedule text into an Excel workbook with sheets **week**, **detail**, and **summary** (with **«Перерва»** column).

## One-click Deploy (Heroku)

[![Deploy to Heroku](https://www.herokucdn.com/deploy/button.svg)](https://heroku.com/deploy?template=https://github.com/YOUR_GITHUB_USERNAME/schedule-bot-heroku)

> Replace `YOUR_GITHUB_USERNAME/schedule-bot-heroku` with your public repo URL after you push this code to GitHub.

### Manual (CLI) Deploy

```bash
heroku login
heroku create schedule-bot
heroku git:remote -a schedule-bot
heroku config:set BOT_TOKEN=123456789:ABCyourTelegramToken
git init && git add . && git commit -m "init"
git push heroku HEAD:main   # or master, depending on your default branch
heroku ps:scale worker=1
heroku logs --tail
```

## Local run

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
export BOT_TOKEN=123456789:ABCyourTelegramToken
python bot.py
```

## Format Example

```
Понеділок  08.09

Казидуб 8:00-18:00 (12:30-13:00)
Кіселиця 10:00-20:00 (13:00-13:30)

Вівторок 09.09

Беньковська 8:00-18:00 (12:30-13:00)
Пую 10:00-21:00 (13:00-13:30)
```
