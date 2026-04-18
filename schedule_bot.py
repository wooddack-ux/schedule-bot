#!/usr/bin/env python3
"""
Бот расписания ВУНЦ ВВС
Поддерживает группы: 20-21, 20-22, 20-23, 11-21, 26-21, 7-21, 8-21
"""

import pandas as pd
import datetime
import asyncio
import re
import os
import shutil
import logging
from aiogram import Bot, Dispatcher, executor, types
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiohttp import web

# ========== НАСТРОЙКА ЛОГИРОВАНИЯ ==========
logging.basicConfig(
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# ========== КОНФИГУРАЦИЯ ==========
API_TOKEN = os.getenv('TELEGRAM_BOT_TOKEN')
if not API_TOKEN:
    logger.error("❌ TELEGRAM_BOT_TOKEN не найден в переменных окружения!")
    exit(1)

# Доступные группы (без "и")
AVAILABLE_GROUPS = ['20-21', '20-22', '20-23', '11-21', '26-21', '7-21', '8-21']

# Хранилище состояний
storage = MemoryStorage()
bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot, storage=storage)

# ========== КОНСТАНТЫ ==========
MONTHS_RU = {
    1: "январь", 2: "февраль", 3: "март", 4: "апрель",
    5: "май", 6: "июнь", 7: "июль", 8: "август",
    9: "сентябрь", 10: "октябрь", 11: "ноябрь", 12: "декабрь"
}

MONTHS_RU_REVERSE = {
    "январь": 1, "февраль": 2, "март": 3, "апрель": 4,
    "май": 5, "июнь": 6, "июль": 7, "август": 8,
    "сентябрь": 9, "октябрь": 10, "ноябрь": 11, "декабрь": 12
}

DAYS_RU = {
    0: "ПН", 1: "ВТ", 2: "СР", 3: "ЧТ", 4: "ПТ", 5: "СБ", 6: "ВС"
}

PAIR_EMOJI = {
    'л': '📖',
    'лек': '📖',
    'пз': '✏️',
    'с': '🗣️',
    'сем': '🗣️',
    'гз': '👥',
    'кр': '📝',
    'экз': '📋',
    'зач': '✅',
    'з/о': '✅',
    'вси': '🎯'
}

# Глобальные переменные
df_schedule = None
user_groups = {}  # user_id -> group


# ========== СОСТОЯНИЯ FSM ==========
class SearchStates(StatesGroup):
    wait_date = State()
    wait_subject = State()
    wait_excel = State()


# ========== KEEP-ALIVE СЕРВЕР (для Render) ==========
async def handle(request):
    return web.Response(text="Бот активен")

async def start_web_server():
    app = web.Application()
    app.router.add_get('/', handle)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, '0.0.0.0', int(os.getenv('PORT', 7860)))
    await site.start()
    logger.info("✅ Keep-alive сервер запущен")


# ========== ПАРСИНГ ==========
def parse_metadata(meta_str):
    """Извлекает тему и занятие из строки типа '1 2 пз'"""
    if pd.isna(meta_str):
        return "", "", ""
    
    s = str(meta_str).strip()
    
    # Ищем паттерн: число пробел число пробел буквы
    match = re.search(r'(\d+)\s+(\d+)\s*([а-яa-z/]+)?', s.lower())
    if match:
        topic = match.group(1)
        lesson = match.group(2)
        ptype = match.group(3) if match.group(3) else ""
        return topic, lesson, ptype
    
    # Если не нашли, возвращаем пустые строки
    return "", "", ""


def get_schedule_for_date(df, target_date, target_group):
    """Получает расписание для группы на конкретную дату"""
    if df is None:
        return None
    
    weekday = target_date.weekday()
    if weekday > 5:  # Воскресенье
        return []
    
    target_month = MONTHS_RU[target_date.month]
    target_day = str(target_date.day)
    
    # Колонки для дня недели (3 колонки на день)
    col_start = 1 + weekday * 3
    day_cols = [col_start, col_start + 1, col_start + 2]
    
    # Ищем строку с датой
    date_row = -1
    for r in range(min(100, df.shape[0])):
        row_values = df.iloc[r].astype(str).str.lower().tolist()
        row_str = " ".join(row_values)
        
        if target_month.lower() in row_str:
            # Проверяем колонки вокруг
            for c in range(max(0, col_start - 2), min(df.shape[1], col_start + 4)):
                val = str(df.iloc[r, c]).strip()
                if val == target_day or val.startswith(target_day):
                    date_row = r
                    break
        if date_row != -1:
            break
    
    if date_row == -1:
        return None
    
    # Ищем строку с группой (после даты)
    group_row = -1
    for r in range(date_row + 1, min(date_row + 25, df.shape[0])):
        cell_val = str(df.iloc[r, 0])
        if target_group in cell_val:
            group_row = r
            break
    
    if group_row == -1:
        return []
    
    lessons = []
    
    for i, col in enumerate(day_cols):
        if col >= df.shape[1]:
            continue
        
        # Предмет (строка с группой)
        subject = df.iloc[group_row, col]
        if pd.isna(subject):
            continue
        
        subject = str(subject).strip()
        
        # Пропускаем служебные слова
        skip_words = ['СР', 'Выходной', 'Праздник', 'Наряд', 'nan', '']
        if subject in skip_words or subject.upper() in [m.upper() for m in MONTHS_RU.values()]:
            continue
        
        # Ищем метаданные (строки ВЫШЕ группы)
        meta_raw = ""
        room = ""
        
        for r_off in range(1, 4):
            if group_row - r_off >= 0:
                val = df.iloc[group_row - r_off, col]
                if not pd.isna(val):
                    val_str = str(val).strip()
                    val_lower = val_str.lower()
                    
                    # Проверяем, это метаданные или аудитория
                    if any(x in val_lower for x in ['пз', ' л', ' т', 'з/о', ' с', 'вси', 'гз', 'экз', 'зач']):
                        meta_raw = val_str
                    elif re.search(r'\d+[дД]', val_str):  # Аудитория (например, 405д)
                        room = val_str
        
        # Определяем тип пары
        pair_type = 'л'
        meta_lower = meta_raw.lower()
        
        if 'пз' in meta_lower:
            pair_type = 'пз'
        elif 'гз' in meta_lower:
            pair_type = 'гз'
        elif 'с' in meta_lower and 'вси' not in meta_lower:
            pair_type = 'с'
        elif 'вси' in meta_lower:
            pair_type = 'вси'
        elif 'кр' in meta_lower:
            pair_type = 'кр'
        elif 'экз' in meta_lower:
            pair_type = 'экз'
        elif 'з/о' in meta_lower or 'зач' in meta_lower:
            pair_type = 'з/о'
        
        topic, lesson, _ = parse_metadata(meta_raw)
        
        pair_num = i + 1
        
        lessons.append({
            'pair_num': pair_num,
            'subject': subject,
            'type': pair_type,
            'topic': topic,
            'lesson': lesson,
            'room': room,
            'meta_raw': meta_raw
        })
    
    # Объединяем одинаковые предметы подряд
    merged = []
    i = 0
    while i < len(lessons):
        current = lessons[i].copy()
        j = i + 1
        while j < len(lessons) and lessons[j]['subject'] == current['subject']:
            current['pair_num'] = f"{current['pair_num']}-{lessons[j]['pair_num']}"
            j += 1
        merged.append(current)
        i = j
    
    return merged


def format_lesson(lesson):
    """Форматирует одну пару для вывода"""
    emoji = PAIR_EMOJI.get(lesson['type'], '📚')
    
    # Название пары
    if isinstance(lesson['pair_num'], str):
        pair_text = f"*{lesson['pair_num']} пары*"
    else:
        pair_text = f"*{lesson['pair_num']} пара*"
    
    text = f"{emoji} {pair_text}: {lesson['subject']}"
    
    # Тема и занятие
    if lesson['topic']:
        text += f"\n  └ 📌 Тема {lesson['topic']}"
        if lesson['lesson']:
            text += f" | Занятие {lesson['lesson']}"
    
    # Тип занятия
    type_names = {'л': 'Лекция', 'пз': 'Практика', 'с': 'Семинар', 
                  'гз': 'Групповое', 'кр': 'Контрольная', 'экз': 'Экзамен', 
                  'з/о': 'Зачёт', 'вси': 'ВСИ'}
    text += f"\n  └ 📝 {type_names.get(lesson['type'], lesson['type'].upper())}"
    
    # Аудитория
    if lesson['room']:
        text += f"\n  └ 🚪 {lesson['room']}"
    
    return text


def get_schedule_text(df, target_date, group):
    """Получает отформатированный текст расписания на дату"""
    lessons = get_schedule_for_date(df, target_date, group)
    day_name = DAYS_RU[target_date.weekday()]
    month_name = MONTHS_RU[target_date.month]
    
    header = f"📅 *{target_date.day} {month_name} ({day_name})*\n👥 Группа: *{group}*\n\n"
    
    if target_date.weekday() == 6:
        return header + "✨ *Воскресенье! Выходной день* ✨"
    
    if not lessons:
        return header + "✨ *Нет занятий* ✨"
    
    text = header
    for lesson in lessons:
        text += format_lesson(lesson) + "\n\n"
    
    return text.strip()


def find_subject(df, group, query):
    """Ищет все вхождения предмета для группы"""
    if df is None:
        return []
    
    results = []
    query_lower = query.lower().strip()
    
    # Ищем с января по июнь 2026
    for month in range(1, 7):
        for day in range(1, 32):
            try:
                dt = datetime.date(2026, month, day)
                if dt.weekday() > 5:
                    continue
                
                lessons = get_schedule_for_date(df, dt, group)
                if lessons:
                    for lesson in lessons:
                        if query_lower in lesson['subject'].lower():
                            results.append({
                                'date': dt,
                                'lesson': lesson
                            })
            except ValueError:
                continue
    
    return results


# ========== КЛАВИАТУРЫ ==========
def get_main_keyboard():
    keyboard = ReplyKeyboardMarkup(resize_keyboard=True)
    keyboard.row(KeyboardButton("📅 Сегодня"), KeyboardButton("📆 На 2 дня"))
    keyboard.row(KeyboardButton("🔍 По дате"), KeyboardButton("🔎 По предмету"))
    keyboard.row(KeyboardButton("👤 Сменить группу"))
    keyboard.row(KeyboardButton("📁 Загрузить Excel"))
    return keyboard


def get_groups_keyboard():
    keyboard = InlineKeyboardMarkup(row_width=3)
    buttons = [InlineKeyboardButton(g, callback_data=f"group_{g}") for g in AVAILABLE_GROUPS]
    keyboard.add(*buttons)
    keyboard.add(InlineKeyboardButton("❌ Отмена", callback_data="cancel"))
    return keyboard


# ========== ОБРАБОТЧИКИ КОМАНД ==========
@dp.message_handler(commands=['start', 'help'])
async def cmd_start(message: types.Message):
    user_id = message.from_user.id
    if user_id not in user_groups:
        user_groups[user_id] = '20-21'
    
    status = "✅ Загружено" if df_schedule is not None else "⚠️ Ожидание файла"
    
    await message.answer(
        f"🎓 *БОТ РАСПИСАНИЯ ВУНЦ ВВС*\n\n"
        f"┌─────────────────────────┐\n"
        f"│ 👤 Группа: *{user_groups[user_id]}*\n"
        f"│ 📊 Статус: {status}\n"
        f"│ 📋 Групп доступно: {len(AVAILABLE_GROUPS)}\n"
        f"└─────────────────────────┘\n\n"
        f"*Доступные команды:*\n"
        f"• 📅 Сегодня — пары на сегодня\n"
        f"• 📆 На 2 дня — расписание на 2 дня\n"
        f"• 🔍 По дате — поиск по дате\n"
        f"• 🔎 По предмету — поиск дисциплины\n"
        f"• 👤 Сменить группу — выбор группы\n"
        f"• 📁 Загрузить Excel — обновить файл",
        parse_mode="Markdown",
        reply_markup=get_main_keyboard()
    )


@dp.message_handler(lambda m: m.text == "📅 Сегодня")
async def cmd_today(message: types.Message):
    if df_schedule is None:
        await message.answer("⚠️ *Сначала загрузите Excel файл!*", parse_mode="Markdown")
        return
    
    user_id = message.from_user.id
    group = user_groups.get(user_id, '20-21')
    today = datetime.date.today()
    
    text = get_schedule_text(df_schedule, today, group)
    await message.answer(text, parse_mode="Markdown")


@dp.message_handler(lambda m: m.text == "📆 На 2 дня")
async def cmd_two_days(message: types.Message):
    if df_schedule is None:
        await message.answer("⚠️ *Сначала загрузите Excel файл!*", parse_mode="Markdown")
        return
    
    user_id = message.from_user.id
    group = user_groups.get(user_id, '20-21')
    today = datetime.date.today()
    
    text = f"📆 *РАСПИСАНИЕ НА 2 ДНЯ*\n👥 Группа: *{group}*\n\n"
    
    for i in range(2):
        dt = today + datetime.timedelta(days=i)
        lessons = get_schedule_for_date(df_schedule, dt, group)
        day_name = DAYS_RU[dt.weekday()]
        month_name = MONTHS_RU[dt.month]
        
        text += f"📌 *{dt.day} {month_name} ({day_name})*\n"
        if dt.weekday() == 6:
            text += "  ✨ Выходной\n"
        elif not lessons:
            text += "  ✨ Нет занятий\n"
        else:
            for lesson in lessons[:4]:
                emoji = PAIR_EMOJI.get(lesson['type'], '📚')
                pair_num = lesson['pair_num']
                text += f"  {emoji} *{pair_num}* — {lesson['subject']}\n"
        text += "\n"
    
    await message.answer(text, parse_mode="Markdown")


@dp.message_handler(lambda m: m.text == "🔍 По дате")
async def cmd_search_date_start(message: types.Message):
    if df_schedule is None:
        await message.answer("⚠️ *Сначала загрузите Excel файл!*", parse_mode="Markdown")
        return
    
    await SearchStates.wait_date.set()
    await message.answer(
        "📅 *Поиск по дате*\n\n"
        "Введите дату в формате:\n`ДД.ММ.ГГГГ`\n\n"
        "Например: `12.01.2026`",
        parse_mode="Markdown"
    )


@dp.message_handler(state=SearchStates.wait_date)
async def cmd_search_date_handle(message: types.Message, state: FSMContext):
    try:
        dt = datetime.datetime.strptime(message.text.strip(), '%d.%m.%Y').date()
        user_id = message.from_user.id
        group = user_groups.get(user_id, '20-21')
        
        text = get_schedule_text(df_schedule, dt, group)
        await message.answer(text, parse_mode="Markdown")
    except ValueError:
        await message.answer("❌ *Неверный формат даты*\nИспользуйте: `ДД.ММ.ГГГГ`", parse_mode="Markdown")
    
    await state.finish()


@dp.message_handler(lambda m: m.text == "🔎 По предмету")
async def cmd_search_subject_start(message: types.Message):
    if df_schedule is None:
        await message.answer("⚠️ *Сначала загрузите Excel файл!*", parse_mode="Markdown")
        return
    
    await SearchStates.wait_subject.set()
    await message.answer(
        "🔎 *Поиск дисциплины*\n\n"
        "Введите название предмета:\n"
        "Например: `ИАО`, `ФП`, `ТВВС`",
        parse_mode="Markdown"
    )


@dp.message_handler(state=SearchStates.wait_subject)
async def cmd_search_subject_handle(message: types.Message, state: FSMContext):
    query = message.text.strip()
    user_id = message.from_user.id
    group = user_groups.get(user_id, '20-21')
    
    results = find_subject(df_schedule, group, query)
    
    if not results:
        await message.answer(
            f"❌ *{query}* не найдено в группе *{group}*",
            parse_mode="Markdown"
        )
    else:
        text = f"🔎 *РЕЗУЛЬТАТЫ ПОИСКА: {query}*\n👥 Группа: *{group}*\n\n"
        
        for r in results[:15]:
            date_str = r['date'].strftime('%d.%m')
            lesson = r['lesson']
            emoji = PAIR_EMOJI.get(lesson['type'], '📚')
            
            text += f"📅 *{date_str}* — {emoji} *{lesson['pair_num']}*: {lesson['subject']}\n"
            if lesson['topic']:
                text += f"  └ 📌 Тема {lesson['topic']}"
                if lesson['lesson']:
                    text += f" | Занятие {lesson['lesson']}"
                text += "\n"
            text += "\n"
        
        if len(results) > 15:
            text += f"\n... и ещё *{len(results) - 15}* занятий"
        
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
    
    await callback_query.message.edit_text(
        f"✅ Группа изменена на *{group}*",
        parse_mode="Markdown"
    )
    await callback_query.answer()


@dp.callback_query_handler(lambda c: c.data == 'cancel')
async def process_cancel_callback(callback_query: types.CallbackQuery):
    await callback_query.message.edit_text("❌ Отменено")
    await callback_query.answer()


@dp.message_handler(lambda m: m.text == "📁 Загрузить Excel")
async def cmd_upload_excel(message: types.Message):
    await SearchStates.wait_excel.set()
    await message.answer("📁 *Отправьте Excel файл с расписанием*", parse_mode="Markdown")


@dp.message_handler(content_types=['document'], state=SearchStates.wait_excel)
async def handle_excel(message: types.Message, state: FSMContext):
    global df_schedule
    
    if not message.document.file_name.endswith(('.xlsx', '.xls')):
        await message.answer("❌ Нужен файл .xlsx или .xls")
        return
    
    msg = await message.answer("⏳ *Обработка файла...*", parse_mode="Markdown")
    
    try:
        file = await bot.get_file(message.document.file_id)
        file_path = "/tmp/schedule.xlsx"
        await file.download(file_path)
        
        df_schedule = pd.read_excel(file_path, header=None)
        
        # Сохраняем файл локально для перезапусков
        shutil.copy(file_path, "schedule.xlsx")
        
        await msg.edit_text(
            f"✅ *Файл успешно загружен!*\n\n"
            f"👥 Доступные группы:\n`{', '.join(AVAILABLE_GROUPS)}`",
            parse_mode="Markdown"
        )
        await message.answer("✅ Готово!", reply_markup=get_main_keyboard())
        
    except Exception as e:
        logger.error(f"Ошибка загрузки: {e}")
        await msg.edit_text(f"❌ *Ошибка:* {str(e)[:100]}", parse_mode="Markdown")
    
    await state.finish()


@dp.message_handler()
async def cmd_unknown(message: types.Message):
    await message.answer("Используйте кнопки меню для навигации", reply_markup=get_main_keyboard())


# ========== УТРЕННЯЯ РАССЫЛКА ==========
async def morning_broadcast():
    """Рассылка расписания в 6:00"""
    while True:
        try:
            now = datetime.datetime.now()
            if now.hour == 6 and now.minute == 0:
                if df_schedule is not None:
                    today = datetime.date.today()
                    logger.info(f"📨 Утренняя рассылка в {now.strftime('%H:%M')}")
                    
                    for user_id, group in user_groups.items():
                        try:
                            text = get_schedule_text(df_schedule, today, group)
                            await bot.send_message(
                                user_id,
                                f"🌅 *ДОБРОЕ УТРО!*\n\n{text}",
                                parse_mode="Markdown"
                            )
                        except Exception as e:
                            logger.error(f"Ошибка отправки {user_id}: {e}")
                
                await asyncio.sleep(61)
            await asyncio.sleep(30)
        except Exception as e:
            logger.error(f"Ошибка в рассылке: {e}")
            await asyncio.sleep(10)


# ========== ЗАГРУЗКА ФАЙЛА ПРИ СТАРТЕ ==========
def load_schedule_on_startup():
    global df_schedule
    if os.path.exists('schedule.xlsx'):
        try:
            df_schedule = pd.read_excel('schedule.xlsx', header=None)
            logger.info("✅ Файл schedule.xlsx загружен при старте")
            return True
        except Exception as e:
            logger.error(f"⚠️ Ошибка загрузки файла: {e}")
    else:
        logger.warning("⚠️ Файл schedule.xlsx не найден")
    return False


# ========== ЗАПУСК ==========
if __name__ == '__main__':
    # Загружаем файл при старте
    load_schedule_on_startup()
    
    # Запускаем keep-alive сервер
    loop = asyncio.get_event_loop()
    loop.create_task(start_web_server())
    
    # Запускаем утреннюю рассылку
    loop.create_task(morning_broadcast())
    
    logger.info("🚀 Бот запущен!")
    
    # Запускаем бота
    executor.start_polling(dp, skip_updates=True)
