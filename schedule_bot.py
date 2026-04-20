"""
Бот расписания ВУНЦ ВВС (aiogram + pandas)
Поддерживает все группы: 20-21, 20-22, 20-23, 11-21, 26-21, 7-21, 8-21
"""

import pandas as pd
import datetime
import asyncio
import re
import os
import shutil
from aiogram import Bot, Dispatcher, executor, types
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiohttp import web

# --- НАСТРОЙКИ ---
API_TOKEN = os.getenv('TELEGRAM_BOT_TOKEN', '8638757224:AAHNEHCDkl7rPe_4T-cDgW4hIjS21I_PF20')
CHAT_ID = 742954985

# Все доступные группы
AVAILABLE_GROUPS = ['20-21', '20-22', '20-23', '11-21', '26-21', '7-21', '8-21']

# Хранилище состояний и выбранных групп
storage = MemoryStorage()
bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot, storage=storage)

MONTHS_RU = {
    1: "январь", 2: "февраль", 3: "март", 4: "апрель",
    5: "май", 6: "июнь", 7: "июль", 8: "август",
    9: "сентябрь", 10: "октябрь", 11: "ноябрь", 12: "декабрь"
}
DAYS_RU = {
    0: "ПН", 1: "ВТ", 2: "СР", 3: "ЧТ", 4: "ПТ", 5: "СБ", 6: "ВС"
}

MONTHS_RU_LOWER = {v: k for k, v in MONTHS_RU.items()}
for name in list(MONTHS_RU_LOWER.keys()):
    MONTHS_RU_LOWER[name.lower()] = MONTHS_RU_LOWER[name]

# Глобальный DataFrame
df_schedule = None
user_groups = {}  # user_id -> group


# --- Состояния FSM ---
class SearchStates(StatesGroup):
    wait_date = State()
    wait_subject = State()
    wait_excel = State()


# --- Keep-alive сервер ---
async def handle(request):
    return web.Response(text="Бот активен")

async def start_web_server():
    app = web.Application()
    app.router.add_get('/', handle)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, '0.0.0.0', 7860)
    await site.start()


# --- Парсинг метаданных ---
def parse_metadata(meta_str):
    """Извлекает номер темы и занятия из строки типа '1 2 пз'"""
    s = str(meta_str).strip().lower()
    if s == 'nan' or not s or s.isdigit():
        return ""
    match = re.search(r'(\d+)\s+(\d+)\s*([а-яa-z/]+)?', s)
    if match:
        topic = match.group(1)
        lesson = match.group(2)
        ptype = match.group(3) if match.group(3) else ""
        if ptype:
            return f"📌 Тема {topic} | Занятие {lesson} ({ptype.upper()})"
        return f"📌 Тема {topic} | Занятие {lesson}"
    return ""


def get_schedule_for_date(df, target_date, target_group):
    """Получает расписание для конкретной группы на дату"""
    if df is None:
        return None
    
    target_month = MONTHS_RU[target_date.month].lower()
    target_day = str(target_date.day)
    weekday = target_date.weekday()
    
    if weekday > 5:  # Воскресенье
        return None
    
    # Колонки для дня недели (каждый день занимает 3 колонки)
    col_start = 1 + weekday * 3
    day_cols = [col_start, col_start + 1, col_start + 2]
    
    # Ищем строку с датой
    date_row = -1
    for r in range(df.shape[0]):
        row_str = " ".join(df.iloc[r].astype(str).str.lower())
        if target_month in row_str:
            for c in range(col_start - 1, col_start + 3):
                if 0 <= c < df.shape[1]:
                    val = str(df.iloc[r, c]).strip().split('.')[0]
                    if val == target_day:
                        date_row = r
                        break
        if date_row != -1:
            break
    
    if date_row == -1:
        return None
    
    # Ищем строку с группой
    group_row = -1
    for r in range(date_row, min(date_row + 20, df.shape[0])):
        cell_val = str(df.iloc[r, 0])
        if target_group in cell_val:
            group_row = r
            break
    
    if group_row == -1:
        return []
    
    raw_lessons = []
    for i, c in enumerate(day_cols):
        if c >= df.shape[1]:
            continue
        
        subj = str(df.iloc[group_row, c]).strip()
        
        # Поиск метаданных
        this_meta = ""
        for r_off in range(1, 4):
            if group_row - r_off >= 0:
                val = str(df.iloc[group_row - r_off, c]).strip()
                if val.lower() != 'nan' and any(x in val.lower() for x in ['пз', ' л', ' т', 'з/о', ' с', 'вси', 'гз', 'экз']):
                    this_meta = val
                    break
        
        # Если ячейка пустая, но есть предыдущая пара
        if (subj.lower() == 'nan' or not subj or subj.isdigit()):
            if i > 0 and raw_lessons:
                subj = raw_lessons[-1]['subj']
                if not this_meta:
                    this_meta = raw_lessons[-1]['meta_raw']
        
        if subj.lower() != 'nan' and subj and not subj.isdigit() and len(subj) >= 2:
            # Фильтруем мусор
            if subj.upper() in ['ЯНВАРЬ', 'ФЕВРАЛЬ', 'МАРТ', 'АПРЕЛЬ', 'МАЙ', 'ИЮНЬ',
                                 'ИЮЛЬ', 'АВГУСТ', 'СЕНТЯБРЬ', 'ОКТЯБРЬ', 'НОЯБРЬ', 'ДЕКАБРЬ']:
                continue
            
            # Определяем тип занятия
            pair_type = '📖 Лекция'
            meta_lower = this_meta.lower()
            if 'пз' in meta_lower:
                pair_type = '✏️ ПЗ'
            elif 'гз' in meta_lower:
                pair_type = '👥 ГЗ'
            elif 'с' in meta_lower and 'вси' not in meta_lower:
                pair_type = '🗣️ Семинар'
            elif 'вси' in meta_lower:
                pair_type = '🎯 ВСИ'
            elif 'кр' in meta_lower:
                pair_type = '📝 КР'
            elif 'экз' in meta_lower:
                pair_type = '📋 Экзамен'
            elif 'з/о' in meta_lower or 'зач' in meta_lower:
                pair_type = '✅ Зачёт'
            
            raw_lessons.append({
                'idx': i + 1,
                'subj': subj,
                'meta': parse_metadata(this_meta),
                'meta_raw': this_meta,
                'type': pair_type
            })
    
    return raw_lessons


def format_lessons(lessons):
    """Форматирует список пар для вывода"""
    if not lessons:
        return []
    
    formatted = []
    i = 0
    while i < len(lessons):
        subj = lessons[i]['subj']
        meta = lessons[i]['meta']
        ptype = lessons[i]['type']
        j = i + 1
        while j < len(lessons) and lessons[j]['idx'] == lessons[j-1]['idx'] + 1 and lessons[j]['subj'] == subj:
            j += 1
        
        n_start = lessons[i]['idx']
        n_end = lessons[j-1]['idx']
        
        if j - i == 1:
            pair_text = f"• *{n_start} пара*: {subj}"
        else:
            pair_text = f"• *{n_start}-{n_end} пары*: {subj}"
        
        if meta:
            pair_text += f"\n  └ {meta}"
        pair_text += f"\n  └ {ptype}"
        
        formatted.append(pair_text)
        i = j
    
    return formatted


def get_schedule_text(df, target_date, group):
    """Получает отформатированный текст расписания"""
    lessons = get_schedule_for_date(df, target_date, group)
    day_name = DAYS_RU[target_date.weekday()]
    
    if target_date.weekday() == 6:
        return f"📅 *{target_date.day} {MONTHS_RU[target_date.month]} ({day_name})*\n👥 Группа: *{group}*\n\n✨ Воскресенье! Выходной ✨"
    
    if not lessons:
        return f"📅 *{target_date.day} {MONTHS_RU[target_date.month]} ({day_name})*\n👥 Группа: *{group}*\n\n✨ Нет занятий ✨"
    
    formatted = format_lessons(lessons)
    text = f"📅 *{target_date.day} {MONTHS_RU[target_date.month]} ({day_name})*\n👥 Группа: *{group}*\n\n"
    text += "\n".join(formatted)
    
    return text


def find_subject(df, group, query):
    """Ищет все вхождения предмета для группы"""
    if df is None:
        return []
    
    results = []
    query_lower = query.lower().strip()
    
    # Проходим по всем датам с января по июнь
    for month in range(1, 7):
        year = 2026
        for day in range(1, 32):
            try:
                dt = datetime.date(year, month, day)
                if dt.weekday() > 5:  # Пропускаем выходные
                    continue
                
                lessons = get_schedule_for_date(df, dt, group)
                if lessons:
                    for lesson in lessons:
                        if query_lower in lesson['subj'].lower():
                            results.append({
                                'date': dt,
                                'subj': lesson['subj'],
                                'meta': lesson['meta'],
                                'type': lesson['type'],
                                'pair_num': lesson['idx']
                            })
            except ValueError:
                continue
    
    return results


# --- Клавиатуры ---
def get_main_keyboard():
    keyboard = ReplyKeyboardMarkup(resize_keyboard=True)
    keyboard.row(KeyboardButton("📅 Сегодня"), KeyboardButton("📆 На 2 дня"))
    keyboard.row(KeyboardButton("🔍 По дате"), KeyboardButton("🔎 По предмету"))
    keyboard.row(KeyboardButton("👤 Сменить группу"))
    keyboard.row(KeyboardButton("📁 Загрузить Excel"))
    return keyboard


def get_groups_keyboard():
    keyboard = InlineKeyboardMarkup(row_width=3)
    buttons = []
    for g in AVAILABLE_GROUPS:
        buttons.append(InlineKeyboardButton(g, callback_data=f"group_{g}"))
    keyboard.add(*buttons)
    keyboard.add(InlineKeyboardButton("❌ Отмена", callback_data="cancel"))
    return keyboard


# --- Обработчики команд ---
@dp.message_handler(commands=['start'])
async def cmd_start(message: types.Message):
    user_id = message.from_user.id
    if user_id not in user_groups:
        user_groups[user_id] = '20-21'
    
    await message.answer(
        f"🎓 *Бот расписания ВУНЦ ВВС*\n\n"
        f"👤 Ваша группа: *{user_groups[user_id]}*\n"
        f"📊 Статус: {'✅ Загружено' if df_schedule is not None else '⚠️ Ожидание файла'}\n\n"
        f"Используйте кнопки меню:",
        parse_mode="Markdown",
        reply_markup=get_main_keyboard()
    )


@dp.message_handler(lambda m: m.text == "📅 Сегодня")
async def cmd_today(message: types.Message):
    if df_schedule is None:
        await message.answer("⚠️ Сначала загрузите Excel файл!")
        return
    
    user_id = message.from_user.id
    group = user_groups.get(user_id, '20-21')
    today = datetime.date.today()
    
    text = get_schedule_text(df_schedule, today, group)
    await message.answer(text, parse_mode="Markdown")


@dp.message_handler(lambda m: m.text == "📆 На 2 дня")
async def cmd_two_days(message: types.Message):
    if df_schedule is None:
        await message.answer("⚠️ Сначала загрузите Excel файл!")
        return
    
    user_id = message.from_user.id
    group = user_groups.get(user_id, '20-21')
    today = datetime.date.today()
    
    text = f"📆 *Расписание на 2 дня*\n👥 Группа: *{group}*\n\n"
    
    for i in range(2):
        dt = today + datetime.timedelta(days=i)
        lessons = get_schedule_for_date(df_schedule, dt, group)
        day_name = DAYS_RU[dt.weekday()]
        
        text += f"📌 *{dt.day} {MONTHS_RU[dt.month]} ({day_name})*\n"
        if dt.weekday() == 6:
            text += "  ✨ Выходной\n"
        elif not lessons:
            text += "  ✨ Нет занятий\n"
        else:
            for lesson in lessons[:3]:  # Показываем до 3 пар в кратком виде
                text += f"  • {lesson['idx']} пара: {lesson['subj']}\n"
        text += "\n"
    
    await message.answer(text, parse_mode="Markdown")


@dp.message_handler(lambda m: m.text == "🔍 По дате")
async def cmd_search_date_start(message: types.Message):
    if df_schedule is None:
        await message.answer("⚠️ Сначала загрузите Excel файл!")
        return
    
    await SearchStates.wait_date.set()
    await message.answer("📅 Введите дату в формате *ДД.ММ.ГГГГ*\nНапример: `12.01.2026`", parse_mode="Markdown")


@dp.message_handler(state=SearchStates.wait_date)
async def cmd_search_date_handle(message: types.Message, state: FSMContext):
    try:
        dt = datetime.datetime.strptime(message.text.strip(), '%d.%m.%Y').date()
        user_id = message.from_user.id
        group = user_groups.get(user_id, '20-21')
        
        text = get_schedule_text(df_schedule, dt, group)
        await message.answer(text, parse_mode="Markdown")
    except ValueError:
        await message.answer("❌ Неверный формат даты")
    
    await state.finish()


@dp.message_handler(lambda m: m.text == "🔎 По предмету")
async def cmd_search_subject_start(message: types.Message):
    if df_schedule is None:
        await message.answer("⚠️ Сначала загрузите Excel файл!")
        return
    
    await SearchStates.wait_subject.set()
    await message.answer("🔎 Введите название предмета:\nНапример: `ИАО`, `ФП`, `ТВВС`")


@dp.message_handler(state=SearchStates.wait_subject)
async def cmd_search_subject_handle(message: types.Message, state: FSMContext):
    query = message.text.strip()
    user_id = message.from_user.id
    group = user_groups.get(user_id, '20-21')
    
    results = find_subject(df_schedule, group, query)
    
    if not results:
        await message.answer(f"❌ *{query}* не найдено в группе *{group}*", parse_mode="Markdown")
    else:
        text = f"🔎 *Результаты поиска: {query}*\n👥 Группа: *{group}*\n\n"
        for r in results[:15]:
            text += f"📅 {r['date'].strftime('%d.%m')} — *{r['pair_num']} пара*: {r['subj']}\n"
            if r['meta']:
                text += f"  └ {r['meta']}\n"
            text += f"  └ {r['type']}\n\n"
        
        if len(results) > 15:
            text += f"\n... и ещё {len(results) - 15} занятий"
        
        await message.answer(text, parse_mode="Markdown")
    
    await state.finish()


@dp.message_handler(lambda m: m.text == "👤 Сменить группу")
async def cmd_change_group(message: types.Message):
    await message.answer(
        "👤 *Выберите группу:*",
        parse_mode="Markdown",
        reply_markup=get_groups_keyboard()
    )


@dp.callback_query_handler(lambda c: c.data.startswith('group_'))
async def process_group_callback(callback_query: types.CallbackQuery):
    group = callback_query.data.replace('group_', '')
    user_id = callback_query.from_user.id
    user_groups[user_id] = group
    
    await callback_query.message.edit_text(f"✅ Группа изменена на *{group}*", parse_mode="Markdown")
    await callback_query.answer()


@dp.callback_query_handler(lambda c: c.data == 'cancel')
async def process_cancel_callback(callback_query: types.CallbackQuery):
    await callback_query.message.edit_text("❌ Отменено")
    await callback_query.answer()


@dp.message_handler(lambda m: m.text == "📁 Загрузить Excel")
async def cmd_upload_excel(message: types.Message):
    await SearchStates.wait_excel.set()
    await message.answer("📁 Отправьте Excel файл с расписанием")


@dp.message_handler(content_types=['document'], state=SearchStates.wait_excel)
async def handle_excel(message: types.Message, state: FSMContext):
    global df_schedule
    
    if not message.document.file_name.endswith(('.xlsx', '.xls')):
        await message.answer("❌ Нужен файл .xlsx или .xls")
        return
    
    await message.answer("⏳ Обработка файла...")
    
    try:
        file = await bot.get_file(message.document.file_id)
        file_path = "/tmp/schedule.xlsx"
        await file.download(file_path)
        
        df_schedule = pd.read_excel(file_path, header=None)
        
        await message.answer(
            f"✅ *Файл загружен!*\n\n"
            f"📊 Доступные группы:\n`{', '.join(AVAILABLE_GROUPS)}`",
            parse_mode="Markdown",
            reply_markup=get_main_keyboard()
        )
    except Exception as e:
        await message.answer(f"❌ Ошибка: {e}")
    
    await state.finish()


@dp.message_handler(lambda m: m.text and m.text not in ["📅 Сегодня", "📆 На 2 дня", "🔍 По дате", 
                                                          "🔎 По предмету", "👤 Сменить группу", "📁 Загрузить Excel"])
async def cmd_unknown(message: types.Message):
    await message.answer("Используйте кнопки меню для навигации")


# --- Утренняя рассылка ---
async def scheduled_task():
    while True:
        try:
            now = datetime.datetime.now()
            if now.hour == 6 and now.minute == 0:
                if df_schedule is not None:
                    today = datetime.date.today()
                    text = f"🌅 *ДОБРОЕ УТРО!*\n\n"
                    
                    # Отправляем каждому пользователю его группу
                    for user_id, group in user_groups.items():
                        try:
                            schedule_text = get_schedule_text(df_schedule, today, group)
                            await bot.send_message(user_id, text + schedule_text, parse_mode="Markdown")
                        except Exception as e:
                            print(f"Ошибка отправки {user_id}: {e}")
                
                await asyncio.sleep(61)
            await asyncio.sleep(30)
        except Exception as e:
            print(f"Ошибка в scheduled_task: {e}")
            await asyncio.sleep(10)


# --- Запуск ---
if __name__ == '__main__':
    # Пробуем загрузить файл при старте
    if os.path.exists('schedule.xlsx'):
        try:
            df_schedule = pd.read_excel('schedule.xlsx', header=None)
            print("✅ Файл schedule.xlsx загружен при старте")
        except Exception as e:
            print(f"⚠️ Ошибка загрузки файла: {e}")
    
    loop = asyncio.get_event_loop()
    loop.create_task(start_web_server())
    loop.create_task(scheduled_task())
    executor.start_polling(dp, skip_updates=True)
