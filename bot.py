import os, re, sys
import pandas as pd
from datetime import datetime, timedelta, date
from typing import List, Dict
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, CallbackQueryHandler

DAY_NAMES = ["Понеділок","Вівторок","Середа","Cереда","Четвер","П'ятниця","Субота","Неділя"]
DAY_PATTERN = r"^(?P<dow>" + "|".join(map(re.escape, DAY_NAMES)) + r")\s+(?P<date>\d{1,2}\.\d{1,2})\s*$"
day_re = re.compile(DAY_PATTERN)

SHIFT_PATTERN = r"""^(?P<name>[\wА-Яа-яІіЇїЄє'’ -]+?)\s+
                     (?P<start>\d{1,2}:\d{2})-(?P<end>\d{1,2}:\d{2})
                     (?:\s*\((?P<bstart>\d{1,2}:\d{2})-(?P<bend>\d{1,2}:\d{2})\))?$
                  """
shift_re = re.compile(SHIFT_PATTERN, re.X)

# --- Табель helpers ---
def hhmm_to_dt(s: str) -> datetime:
    h, m = map(int, s.split(":"))
    return datetime(2000,1,1,h,m)

def compute_tabulated_hours(start, end, dow: str) -> float:
    sdt, edt = hhmm_to_dt(start), hhmm_to_dt(end)
    if dow in ["Субота","Неділя"]:
        legal_start, legal_end = hhmm_to_dt("09:00"), hhmm_to_dt("19:00")
    else:
        legal_start, legal_end = hhmm_to_dt("08:00"), hhmm_to_dt("20:00")
    sdt = max(sdt, legal_start)
    edt = min(edt, legal_end)
    if edt <= sdt:
        return 0.0
    hours = (edt - sdt).total_seconds()/3600.0
    if hours >= 6:
        hours -= 1
    return round(hours,2)

# --- Parsing schedule text ---
def parse_blocks_for_text(raw_text: str, year: int) -> List[Dict]:
    lines = [ln.strip() for ln in raw_text.splitlines()]
    blocks, current = [], None
    for ln in lines:
        if not ln: continue
        m = day_re.match(ln)
        if m:
            dow = m.group("dow").replace("Cереда","Середа")
            ddmm = m.group("date")
            current = {"dow": dow, "date": ddmm, "entries": []}
            blocks.append(current); continue
        if current is None: continue
        m2 = shift_re.match(re.sub(r"\s{2,}", " ", ln))
        if m2:
            d = m2.groupdict()
            name = re.sub(r"\s{2,}", " ", d["name"]).strip()
            start, end = d["start"], d["end"]
            raw = (hhmm_to_dt(end)-hhmm_to_dt(start)).total_seconds()/3600.0
            tab = compute_tabulated_hours(start,end,current["dow"])
            current["entries"].append({
                "name": name,
                "start": start, "end": end,
                "bstart": d.get("bstart"), "bend": d.get("bend"),
                "raw_hours": round(raw,2),
                "tab_hours": tab
            })
    for b in blocks:
        dt = datetime.strptime(f"{b['date']}.{year}", "%d.%m.%Y").date()
        b["date_iso"] = dt.isoformat()
        b["total_hours"] = sum(e["tab_hours"] for e in b["entries"])
    return blocks

def anchor_monday(min_date: date, user_anchor: str|None) -> date:
    if user_anchor:
        try:
            d = datetime.strptime(user_anchor,"%Y-%m-%d").date()
            if d.weekday()!=0: d -= timedelta(days=d.weekday())
            return d
        except: pass
    return min_date - timedelta(days=min_date.weekday())

def assign_weeks(blocks: List[Dict], first_monday: date):
    for b in blocks:
        d = datetime.fromisoformat(b["date_iso"]).date()
        b["week"] = 1+((d-first_monday).days//7)

# --- Build DataFrames ---
def build_week_df(blocks: List[Dict], week_no:int) -> pd.DataFrame:
    rows=[]
    week_blocks=sorted([b for b in blocks if b["week"]==week_no], key=lambda x:x["date_iso"])
    for b in week_blocks:
        day_label=f"{b['dow']}  {datetime.fromisoformat(b['date_iso']).strftime('%d.%m')} ({int(round(b['total_hours']))}год.)"
        first=True
        for e in b["entries"]:
            brk=f"{e['bstart']}-{e['bend']}" if e.get("bstart") and e.get("bend") else ""
            rows.append({
                "День": day_label if first else "",
                "Працівник": e["name"],
                "Робочі години": f"{e['start']}-{e['end']}",
                "Перерва": brk
            }); first=False
        rows.append({"День":"","Працівник":"","Робочі години":"","Перерва":""})
    if not rows: rows=[{"День":"","Працівник":"","Робочі години":"","Перерва":""}]
    return pd.DataFrame(rows)

def build_wide_weeks(blocks,weeks:int)->pd.DataFrame:
    week_tables={w:build_week_df(blocks,w) for w in range(1,weeks+1)}
    max_rows=max((df.shape[0] for df in week_tables.values()),default=0)
    padded=[]
    for w in range(1,weeks+1):
        df=week_tables[w]
        if df.shape[0]<max_rows:
            pad=pd.DataFrame([{ "День":"","Працівник":"","Робочі години":"","Перерва":""}]*(max_rows-df.shape[0]))
            df=pd.concat([df,pad],ignore_index=True)
        df.columns=[f"{w}й тиждень","Працівник","Робочі години","Перерва"]
        padded.append(df)
    return pd.concat(padded,axis=1) if padded else pd.DataFrame()

def build_detail(blocks: List[Dict])->pd.DataFrame:
    rows=[]
    for b in blocks:
        for e in b["entries"]:
            rows.append({
                "Дата": datetime.fromisoformat(b["date_iso"]),
                "Тиждень": b["week"],
                "День тижня": b["dow"],
                "Працівник": e["name"],
                "Початок": e["start"], "Кінець": e["end"],
                "Перерва початок": e.get("bstart") or "",
                "Перерва кінець": e.get("bend") or "",
                "Фіксовані години": e["raw_hours"],
                "Табельні години": e["tab_hours"]
            })
    return pd.DataFrame(rows).sort_values(["Дата","Працівник"])

def build_summary(detail:pd.DataFrame,weeks:int)->pd.DataFrame:
    if detail.empty:
        return pd.DataFrame(columns=["Працівник"]+[f"Тиждень {i}" for i in range(1,weeks+1)]+["Всього (год)"])
    pv=detail.pivot_table(index="Працівник",columns="Тиждень",values="Табельні години",aggfunc="sum",fill_value=0)
    pv=pv.reindex(columns=list(range(1,weeks+1)),fill_value=0)
    pv.columns=[f"Тиждень {c}" for c in pv.columns]
    pv["Всього (год)"]=pv.sum(axis=1).round(2)
    return pv.reset_index()
# function to calculate working days summary
def build_working_days_summary(blocks: List[Dict]) -> pd.DataFrame:
    rows = []
    for b in blocks:
        for e in b["entries"]:
            rows.append({
                "Працівник": e["name"],
                "Тиждень": b["week"],
                "День": b["dow"],
                "Дата": datetime.fromisoformat(b["date_iso"]),
                "Години": e["tab_hours"]
            })
    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=["Працівник", "Тиждень", "День", "Дата", "Години"])
    
    # Group by employee and week, and aggregate the data
    summary = (
        df.groupby(["Працівник", "Тиждень"])
        .agg(
            Дні=("День", lambda x: ", ".join(sorted(set(x), key=DAY_NAMES.index))),
            Години=("Години", "sum")
        )
        .reset_index()
    )
    return summary
def build_schedule_table(blocks: List[Dict]) -> pd.DataFrame:
    rows = []
    for b in blocks:
        for e in b["entries"]:
            rows.append({
                "Працівник": e["name"],
                "Тиждень": b["week"],
                "День": b["dow"],
                "Дата": datetime.fromisoformat(b["date_iso"]),
                "Години": f"{e['start']}-{e['end']}"
            })
    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=["Працівник"] + DAY_NAMES)
    
    # Pivot the data to create a tabular view
    schedule_table = (
        df.pivot_table(
            index="Працівник",
            columns="День",
            values="Години",
            aggfunc=lambda x: " | ".join(x)  # Combine multiple shifts for the same day
        )
        .reindex(columns=DAY_NAMES, fill_value="")  # Ensure correct day order
        .reset_index()
    )
    return schedule_table

def write_excel(out_path,wide,detail,summary,working_days_summary,schedule_table):
    with pd.ExcelWriter(out_path,engine="xlsxwriter") as writer:
        if not wide.empty: wide.to_excel(writer,sheet_name="week",index=False)
        if not detail.empty: detail.to_excel(writer,sheet_name="detail",index=False)
        if not summary.empty: summary.to_excel(writer,sheet_name="summary",index=False)
        if not working_days_summary.empty:working_days_summary.to_excel(writer, sheet_name="working days summary", index=False)
        if not schedule_table.empty:schedule_table.to_excel(writer, sheet_name="schedule table", index=False)
        for sheet in writer.sheets.values():
            sheet.set_column(0,0,26); sheet.set_column(1,100,18)




# --- Telegram bot state ---
user_settings = {}

# --- Handlers ---
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_settings[update.effective_chat.id] = {"year": datetime.now().year, "weeks": 4, "anchor": None}
    keyboard = [
        [InlineKeyboardButton("ℹ️ Help", callback_data="help")],
        [InlineKeyboardButton("⚙️ Settings", callback_data="settings")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text(
        "👋 Привіт! Я бот для обробки розкладів. Надішліть розклад як .txt файл або скористайтеся кнопками нижче.",
        reply_markup=reply_markup
    )

async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = (
        "ℹ️ <b>Довідка</b>\n\n"
        "📅 <b>Команди:</b>\n"
        "• /year YYYY – встановити рік\n"
        "• /weeks N – кількість тижнів у виводі\n"
        "• /anchor YYYY-MM-DD – задати понеділок як перший тиждень\n"
        "• /reset – скинути налаштування до стандартних\n\n"
        "📂 Надішліть розклад як .txt файл, і я створю таблицю 📊"
    )
    await update.message.reply_text(msg, parse_mode="HTML")

async def year_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):

    try:
        year = int(context.args[0])
        user_settings[update.effective_chat.id]["year"] = year
        await update.message.reply_text(f"✅ Рік змінено на {year} 📅")
    except:
        await update.message.reply_text("❌ Вкажіть рік, наприклад: /year 2025")

async def weeks_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):

    try:
        w = int(context.args[0])
        user_settings[update.effective_chat.id]["weeks"] = w
        await update.message.reply_text(f"✅ Кількість тижнів змінено на {w} 📆")
    except:
        await update.message.reply_text("❌ Вкажіть число, наприклад: /weeks 6")

async def anchor_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):

    try:
        d = datetime.strptime(context.args[0], "%Y-%m-%d").date()
        user_settings[update.effective_chat.id]["anchor"] = d
        await update.message.reply_text(f"✅ Початковий понеділок встановлено: {d} 📌")
    except:
        await update.message.reply_text("❌ Вкажіть дату у форматі YYYY-MM-DD")

async def reset_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_settings[update.effective_chat.id] = {"year": datetime.now().year, "weeks": 4, "anchor": None}
    await update.message.reply_text(
        "⚙️ Налаштування скинуто.\n"
        "Надішліть розклад як .txt файл або скористайтеся /help для довідки."
    )

async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    if text:
        await process_schedule_and_reply(update, context, text)

async def txt_document_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if doc.mime_type != "text/plain":
        await update.message.reply_text("❌ Будь ласка, надішліть .txt файл 📂")
        return
    file = await doc.get_file()
    text = (await file.download_as_bytearray()).decode("utf-8")
    await process_schedule_and_reply(update, context, text)

# --- Callback Query Handler ---
async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "help":
        await help_cmd(update, context)
    elif query.data == "settings":
        await query.edit_message_text(
            "⚙️ <b>Налаштування:</b>\n"
            "• /year YYYY – встановити рік\n"
            "• /weeks N – кількість тижнів у виводі\n"
            "• /anchor YYYY-MM-DD – задати понеділок як перший тиждень\n"
            "• /reset – скинути налаштування до стандартних",
            parse_mode="HTML"
        )

# --- Core ---
async def process_schedule_and_reply(update: Update, context: ContextTypes.DEFAULT_TYPE, text: str):
    settings = user_settings.get(update.effective_chat.id, {"year": datetime.now().year, "weeks": 4, "anchor": None})
    year = settings["year"]
    blocks=parse_blocks_for_text(text,year)
    if not blocks:
        await update.message.reply_text("Не вдалося розпізнати розклад")
        return
    min_date=min(datetime.fromisoformat(b["date_iso"]).date() for b in blocks)
    first_monday=anchor_monday(min_date, settings["anchor"].isoformat() if settings["anchor"] else None)
    assign_weeks(blocks,first_monday)
    weeks=settings["weeks"]
    wide=build_wide_weeks(blocks,weeks)
    detail=build_detail(blocks)
    summary=build_summary(detail,weeks)
    working_days_summary = build_working_days_summary(blocks)
    schedule_table = build_schedule_table(blocks)
    out_path="schedule.xlsx"
    write_excel(out_path,wide,detail,summary,working_days_summary,schedule_table)
    await update.message.reply_document(open(out_path,"rb"), filename="schedule.xlsx")

# --- Main ---
def main():
    # Add logging to debug BOT_TOKEN
    print("Starting bot...")
    token = os.environ.get("BOT_TOKEN")
    if not token:
        print("Error: BOT_TOKEN environment variable is not set.")
        sys.exit(1)
    else:
        print(f"BOT_TOKEN is set: {token[:4]}... (truncated for security)")


    app = Application.builder().token(token).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("year", year_cmd))
    app.add_handler(CommandHandler("weeks", weeks_cmd))
    app.add_handler(CommandHandler("anchor", anchor_cmd))
    app.add_handler(CommandHandler("reset", reset_cmd))
    app.add_handler(MessageHandler(filters.Document.MimeType("text/plain"), txt_document_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_handler))
    app.add_handler(CallbackQueryHandler(button_handler))  # Added callback query handler
    app.run_polling()

if __name__=="__main__":
    main()

