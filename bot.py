#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import io
import re
from dataclasses import dataclass
from datetime import datetime, timedelta, date
from typing import List, Dict, Optional

import pandas as pd
from telegram import Update, InputFile
from telegram.constants import ParseMode
from telegram.ext import (
    Application, CommandHandler, MessageHandler, ContextTypes, filters
)

DAY_NAMES = ["Понеділок","Вівторок","Середа","Cереда","Четвер","П'ятниця","Субота","Неділя"]
DAY_PATTERN = r"^(?P<dow>" + "|".join(map(re.escape, DAY_NAMES)) + r")\s+(?P<date>\d{2}\.\d{2})\s*$"
day_re = re.compile(DAY_PATTERN)

SHIFT_PATTERN = r"""
^(?P<name>[\wА-Яа-яІіЇїЄє'’- ]+?)\s+
(?P<start>\d{1,2}:\d{2})-(?P<end>\d{1,2}:\d{2})
(?:\s*\((?P<bstart>\d{1,2}:\d{2})-(?P<bend>\d{1,2}:\d{2})\))?$
"""
shift_re = re.compile(SHIFT_PATTERN, re.X)

def compute_duration_hours(start, end, bstart=None, bend=None) -> float:
    def hhmm(s):
        h, m = map(int, s.split(":"))
        return h, m
    sh, sm = hhmm(start)
    eh, em = hhmm(end)
    dur = (datetime(2000,1,1,eh,em) - datetime(2000,1,1,sh,sm)).total_seconds()/3600.0
    if bstart and bend:
        bh, bm = hhmm(bstart); ehh, emm = hhmm(bend)
        dur -= (datetime(2000,1,1,ehh,emm) - datetime(2000,1,1,bh,bm)).total_seconds()/3600.0
    return round(dur, 2)

def parse_blocks_for_text(raw_text: str, year: int) -> List[Dict]:
    lines = [ln.strip() for ln in raw_text.splitlines()]
    blocks = []
    current = None
    for ln in lines:
        if not ln:
            continue
        m = day_re.match(ln)
        if m:
            dow = m.group("dow").replace("Cереда","Середа")
            ddmm = m.group("date")
            current = {"dow": dow, "date": ddmm, "entries": []}
            blocks.append(current); continue
        if current is None:
            continue
        m2 = shift_re.match(re.sub(r"\s{2,}", " ", ln))
        if m2:
            d = m2.groupdict()
            d["name"] = re.sub(r"\s{2,}", " ", d["name"]).strip()
            current["entries"].append(d)
    if not blocks:
        raise ValueError("Не знайдено жодного блоку дня. Перевірте формат.")
    for b in blocks:
        dt = datetime.strptime(f"{b['date']}.{year}", "%d.%m.%Y").date()
        b["date_iso"] = dt.isoformat()
        total = 0.0
        for e in b["entries"]:
            e["hours"] = compute_duration_hours(e["start"], e["end"], e.get("bstart"), e.get("bend"))
            total += e["hours"]
        b["total_hours"] = round(total, 2)
    return blocks

@dataclass
class UserSettings:
    year: int
    weeks: int
    anchor: Optional[str] = None

def anchor_monday(min_date: date, user_anchor: Optional[str]) -> date:
    if user_anchor:
        try:
            d = datetime.strptime(user_anchor, "%Y-%m-%d").date()
            if d.weekday() != 0:
                d = d - timedelta(days=d.weekday())
            return d
        except Exception:
            pass
    return min_date - timedelta(days=min_date.weekday())

def assign_weeks(blocks: List[Dict], first_monday: date) -> None:
    for b in blocks:
        d = datetime.fromisoformat(b["date_iso"]).date()
        b["week"] = 1 + ((d - first_monday).days // 7)

def build_week_df(blocks: List[Dict], week_no: int) -> pd.DataFrame:
    rows = []
    week_blocks = sorted([b for b in blocks if b["week"] == week_no], key=lambda x: x["date_iso"])
    for b in week_blocks:
        day_label = f"{b['dow']}  {datetime.fromisoformat(b['date_iso']).strftime('%d.%m')} ({int(round(b['total_hours']))}год.)"
        first = True
        for e in b["entries"]:
            break_str = f"{e['bstart']}-{e['bend']}" if e.get("bstart") and e.get("bend") else ""
            rows.append({
                "День": day_label if first else "",
                "Працівник": e["name"],
                "Робочі години": f"{e['start']}-{e['end']}",
                "Перерва": break_str
            })
            first = False
        rows.append({"День":"","Працівник":"","Робочі години":"","Перерва":""})
    if not rows:
        rows = [{"День":"","Працівник":"","Робочі години":"","Перерва":""}]
    return pd.DataFrame(rows)

def build_wide_weeks(blocks: List[Dict], weeks: int) -> pd.DataFrame:
    week_tables = {w: build_week_df(blocks, w) for w in range(1, weeks+1)}
    max_rows = max((df.shape[0] for df in week_tables.values()), default=0)
    padded = []
    for w in range(1, weeks+1):
        df = week_tables[w]
        if df.shape[0] < max_rows:
            pad = pd.DataFrame([{"День":"","Працівник":"","Робочі години":"","Перерва":""}] * (max_rows - df.shape[0]))
            df = pd.concat([df, pad], ignore_index=True)
        df.columns = [f"{w}й тиждень","Працівник","Робочі години","Перерва"]
        padded.append(df)
    return pd.concat(padded, axis=1) if padded else pd.DataFrame()

def build_detail(blocks: List[Dict]) -> pd.DataFrame:
    rows = []
    for b in blocks:
        for e in b["entries"]:
            rows.append({
                "Дата": datetime.fromisoformat(b["date_iso"]),
                "Тиждень": b["week"],
                "День тижня": b["dow"],
                "Працівник": e["name"],
                "Початок": e["start"],
                "Кінець": e["end"],
                "Перерва початок": e.get("bstart") or "",
                "Перерва кінець": e.get("bend") or "",
                "Тривалість, год": e["hours"]
            })
    return pd.DataFrame(rows).sort_values(["Дата","Працівник"])

def build_summary(detail: pd.DataFrame, weeks: int) -> pd.DataFrame:
    if detail.empty:
        return pd.DataFrame(columns=["Працівник"] + [f"Тиждень {i}" for i in range(1,weeks+1)] + ["Всього за місяць (год)"])
    pv = detail.pivot_table(index="Працівник", columns="Тиждень", values="Тривалість, год", aggfunc="sum", fill_value=0)
    pv = pv.reindex(columns=list(range(1,weeks+1)), fill_value=0)
    pv.columns = [f"Тиждень {c}" for c in pv.columns]
    pv["Всього за місяць (год)"] = pv.sum(axis=1).round(2)
    return pv.reset_index()

def build_excel_bytes(blocks: List[Dict], weeks: int) -> bytes:
    wide = build_wide_weeks(blocks, weeks)
    detail = build_detail(blocks)
    summary = build_summary(detail, weeks)

    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter", options={"strings_to_numbers": False}) as writer:
        if not wide.empty:
            wide.to_excel(writer, sheet_name="week", index=False)
        if not detail.empty:
            detail.to_excel(writer, sheet_name="detail", index=False)
        if not summary.empty:
            summary.to_excel(writer, sheet_name="summary", index=False)
        for sheet_name, ws in writer.sheets.items():
            if sheet_name == "week":
                ws.set_column(0, 0, 26)  # day label
                ws.set_column(1, 100, 18)
            else:
                ws.set_column(0, 100, 18)
    bio.seek(0)
    return bio.read()

# ====== Bot handlers ======

SAMPLE = (
"Понеділок  08.09

"
"Казидуб 8:00-18:00 (12:30-13:00)
"
"Кіселиця 10:00-20:00 (13:00-13:30)

"
"Вівторок 09.09

"
"Беньковська 8:00-18:00 (12:30-13:00)
"
"Пую 10:00-21:00 (13:00-13:30)
"
)

HELP = (
"Надішліть текст розкладу (або .txt файл) у форматі:
"
"<b>Понеділок  08.09</b>
"
"Прізвище 8:00-18:00 (12:30-13:00)
"
"Прізвище 10:00-20:00 (13:00-13:30)

"
"Команди:
"
"/year 2025 — встановити рік для дат
"
"/weeks 4 — кількість тижневих блоків у листі 'week'
"
"/anchor 2025-09-08 — заякорити підрахунок тижнів від цього понеділка
"
"/format — приклад формату
"
"/reset — скинути налаштування чату
"
)

@dataclass
class ChatState:
    settings: 'UserSettings'

def get_settings(context: ContextTypes.DEFAULT_TYPE, chat_id: int) -> UserSettings:
    if 'settings' not in context.bot_data:
        context.bot_data['settings'] = {}
    settings = context.bot_data['settings'].get(chat_id)
    if not settings:
        settings = UserSettings(year=datetime.now().year, weeks=4, anchor=None)
        context.bot_data['settings'][chat_id] = settings
    return settings

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    s = get_settings(context, update.effective_chat.id)
    await update.message.reply_text(
        "Вітаю! Надішліть розклад текстом або .txt файлом — поверну Excel з аркушами "
        "<b>week</b>, <b>detail</b>, <b>summary</b> і колонкою <b>«Перерва»</b>.

"
        + HELP + f"
Поточні налаштування: year={s.year}, weeks={s.weeks}, anchor={s.anchor}",
        parse_mode=ParseMode.HTML
    )

async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    s = get_settings(context, update.effective_chat.id)
    await update.message.reply_text(HELP + f"
Поточні налаштування: year={s.year}, weeks={s.weeks}, anchor={s.anchor}", parse_mode=ParseMode.HTML)

async def format_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(f"Приклад:

<pre>{SAMPLE}</pre>", parse_mode=ParseMode.HTML)

async def year_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    s = get_settings(context, update.effective_chat.id)
    try:
        y = int(context.args[0])
        s.year = y
        await update.message.reply_text(f"✅ Рік встановлено: {y}")
    except Exception:
        await update.message.reply_text("Використання: /year 2025")

async def weeks_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    s = get_settings(context, update.effective_chat.id)
    try:
        w = int(context.args[0])
        if w < 1 or w > 6:
            raise ValueError
        s.weeks = w
        await update.message.reply_text(f"✅ Кількість тижнів у 'week': {w}")
    except Exception:
        await update.message.reply_text("Використання: /weeks 4 (1..6)")

async def anchor_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    s = get_settings(context, update.effective_chat.id)
    if not context.args:
        s.anchor = None
        await update.message.reply_text("Якір скинуто. Тижні рахуються від понеділка на/перед найранішою датою у розкладі.")
        return
    try:
        d = datetime.strptime(context.args[0], "%Y-%m-%d").date()
        if d.weekday() != 0:
            d = d - timedelta(days=d.weekday())
        s.anchor = d.isoformat()
        await update.message.reply_text(f"✅ Якір тижнів встановлено: {s.anchor}")
    except Exception:
        await update.message.reply_text("Використання: /anchor 2025-09-08 (або без аргументу щоб скинути)")

async def reset_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "settings" in context.bot_data:
        context.bot_data["settings"].pop(update.effective_chat.id, None)
    await update.message.reply_text("Налаштування чату скинуто. Використовуйте /start")

def _blocks_with_settings(raw_text: str, s: UserSettings) -> List[Dict]:
    blocks = parse_blocks_for_text(raw_text, s.year)
    min_date = min(datetime.fromisoformat(b["date_iso"]).date() for b in blocks)
    first_monday = anchor_monday(min_date, s.anchor)
    assign_weeks(blocks, first_monday)
    return blocks

async def _handle_text_to_excel(update: Update, context: ContextTypes.DEFAULT_TYPE, text: str):
    s = get_settings(context, update.effective_chat.id)
    try:
        blocks = _blocks_with_settings(text, s)
        excel_bytes = build_excel_bytes(blocks, s.weeks)
        filename = f"work_schedule_{s.year}.xlsx"
        await update.message.reply_document(document=InputFile(io.BytesIO(excel_bytes), filename=filename),
                                            caption="Ось ваш файл 📄")
    except Exception as e:
        await update.message.reply_text(
            "⚠️ Не вдалось згенерувати файл.
"
            f"Помилка: {e}

Спробуйте /format або перевірте форматування.",
            parse_mode=ParseMode.HTML
        )

async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.message or not update.message.text:
        return
    await _handle_text_to_excel(update, context, update.message.text)

async def txt_document_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc or doc.mime_type != "text/plain":
        return
    f = await doc.get_file()
    content = await f.download_as_bytearray()
    text = content.decode("utf-8", errors="replace")
    await _handle_text_to_excel(update, context, text)

def main():
    token = os.environ.get("BOT_TOKEN")
    if not token:
        print("ERROR: Please set BOT_TOKEN environment variable.")
        raise SystemExit(1)

    app = Application.builder().token(token).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("format", format_cmd))
    app.add_handler(CommandHandler("year", year_cmd))
    app.add_handler(CommandHandler("weeks", weeks_cmd))
    app.add_handler(CommandHandler("anchor", anchor_cmd))
    app.add_handler(CommandHandler("reset", reset_cmd))

    app.add_handler(MessageHandler(filters.Document.MimeType("text/plain"), txt_document_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_handler))

    app.run_polling()

if __name__ == "__main__":
    main()
