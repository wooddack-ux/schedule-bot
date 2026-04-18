"""
Бот расписания ВУНЦ ВВС для Telegram
Автоматическая рассылка расписания с настраиваемыми уведомлениями
"""

import logging
import openpyxl
import re
import os
import shutil
import asyncio
from datetime import datetime, timedelta, time
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler, CallbackQueryHandler

# ========== НАСТРОЙКА ЛОГИРОВАНИЯ ==========
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# ========== СОСТОЯНИЯ ДЛЯ CONVERSATION ==========
SEARCH_DATE, SEARCH_NAME, CUSTOM_SETTINGS = range(3)

# ========== КОНСТАНТЫ ==========
MONTHS_RU = {
    'январь': 1, 'февраль': 2, 'март': 3, 'апрель': 4, 'май': 5, 'июнь': 6,
    'июль': 7, 'август': 8, 'сентябрь': 9, 'октябрь': 10, 'ноябрь': 11, 'декабрь': 12
}

KNOWN_GROUPS = ['20-21', '20-22', '20-23', '11-21', '26-21', '7-21', '8-21', '8и-21', '20и-21', '20и-22']

DEFAULT_SETTINGS = {
    'days_ahead': 2,
    'notify_time': '06:00',
    'warn_pz_s': 3,
    'warn_exam': 5,
    'enabled': True
}


# ========== КЛАСС БОТА ==========
class ScheduleBot:
    def __init__(self):
        self.excel_file = None
        self.workbook = None
        self.schedule_data = {}
        self.groups = {}
        self.user_settings = {}
        self.user_groups = {}
        self.excel_loaded = False
        self.load_excel_on_startup()
    
    def load_excel_on_startup(self):
        try:
            work_dir = os.getcwd()
            files = [f for f in os.listdir(work_dir) if f.endswith(('.xlsx', '.xls')) and not f.startswith('temp_')]
            if not files:
                logger.warning("Excel файл не найден при старте!")
                return False
            self.excel_file = files[0]
            file_path = os.path.join(work_dir, self.excel_file)
            logger.info(f"Загрузка файла: {file_path}")
            self.workbook = openpyxl.load_workbook(file_path, data_only=True)
            self.parse_schedule()
            self.excel_loaded = True
            logger.info(f"✅ Загружено: {len(self.groups)} групп")
            return True
        except Exception as e:
            logger.error(f"Ошибка загрузки: {e}")
            self.excel_loaded = False
            return False
    
    def save_uploaded_excel(self, file_path):
        try:
            work_dir = os.getcwd()
            new_path = os.path.join(work_dir, "schedule.xlsx")
            if os.path.exists(new_path):
                os.remove(new_path)
            shutil.copy2(file_path, new_path)
            try:
                os.remove(file_path)
            except:
                pass
            self.excel_file = "schedule.xlsx"
            self.workbook = openpyxl.load_workbook(new_path, data_only=True)
            self.parse_schedule()
            self.excel_loaded = True
            logger.info(f"✅ Файл сохранен. Групп: {len(self.groups)}")
            return True
        except Exception as e:
            logger.error(f"Ошибка сохранения: {e}")
            return False
    
    def parse_schedule(self):
        self.schedule_data = {}
        self.groups = {}
        for sheet_name in self.workbook.sheetnames:
            if sheet_name == 'Планер':
                continue
            sheet = self.workbook[sheet_name]
            logger.info(f"Обработка листа: '{sheet_name}'")
            sheet_groups = self._get_groups_for_sheet(sheet_name)
            for g in sheet_groups:
                self.groups[g] = g
                if g not in self.schedule_data:
                    self.schedule_data[g] = {}
            self._parse_sheet(sheet, sheet_groups)
        logger.info(f"ИТОГО групп: {len(self.groups)}")
        for g in sorted(self.groups.keys()):
            dates = len(self.schedule_data.get(g, {}))
            pairs = sum(len(p) for p in self.schedule_data.get(g, {}).values())
            logger.info(f"  {g}: {dates} дат, {pairs} пар")
    
    def _get_groups_for_sheet(self, sheet_name):
        if '20,20и' in sheet_name:
            return ['20-21', '20-22', '20-23', '11-21', '20и-21', '20и-22']
        elif '26' in sheet_name:
            return ['26-21']
        elif '7,8,8и' in sheet_name:
            return ['7-21', '8-21', '8и-21']
        return []
    
    def _parse_sheet(self, sheet, sheet_groups):
        day_columns = [
            (2, 'Пн', 1), (3, 'Пн', 2), (4, 'Пн', 3),
            (5, 'Вт', 1), (6, 'Вт', 2), (7, 'Вт', 3),
            (8, 'Ср', 1), (9, 'Ср', 2), (10, 'Ср', 3),
            (11, 'Чт', 1), (12, 'Чт', 2), (13, 'Чт', 3),
            (14, 'Пт', 1), (15, 'Пт', 2), (16, 'Пт', 3),
            (17, 'Сб', 1), (18, 'Сб', 2), (19, 'Сб', 3),
        ]
        current_dates = {}
        for row in range(1, min(500, sheet.max_row + 1)):
            for col in range(1, min(20, sheet.max_column + 1)):
                cell = sheet.cell(row, col).value
                if not cell or not isinstance(cell, str):
                    continue
                cell_lower = cell.lower().strip()
                if cell_lower in MONTHS_RU:
                    day_cell = sheet.cell(row, 1).value
                    if day_cell:
                        try:
                            day = int(day_cell)
                            month = MONTHS_RU[cell_lower]
                            year = 2026 if month <= 6 else 2025
                            current_dates[col] = (day, month, year)
                        except:
                            pass
            col_a = sheet.cell(row, 1).value
            if not col_a:
                continue
            col_a_str = str(col_a).strip()
            found_groups = []
            for g in sheet_groups:
                if g in col_a_str:
                    found_groups.append(g)
            if found_groups:
                for look_back in range(1, 5):
                    data_row = row - look_back
                    if data_row < 1:
                        continue
                    for col, day_name, pair_num in day_columns:
                        if col > sheet.max_column:
                            continue
                        cell_value = sheet.cell(data_row, col).value
                        if not cell_value:
                            continue
                        cell_str = str(cell_value).strip()
                        if cell_str in ['', 'None', '-', 'СР', 'Выходной', 'Праздник', 'Наряд']:
                            continue
                        date_info = None
                        for date_col in sorted(current_dates.keys(), reverse=True):
                            if date_col <= col:
                                date_info = current_dates[date_col]
                                break
                        if not date_info:
                            continue
                        day, month, year = date_info
                        date_obj = datetime(year, month, day)
                        pair_type = 'л'
                        type_cell = sheet.cell(data_row - 1, col).value if data_row > 1 else None
                        if type_cell:
                            type_str = str(type_cell).lower()
                            if 'пз' in type_str:
                                pair_type = 'пз'
                            elif 'гз' in type_str:
                                pair_type = 'гз'
                            elif 'с' in type_str:
                                pair_type = 'с'
                            elif 'кр' in type_str:
                                pair_type = 'кр'
                            elif 'экз' in type_str or 'э' in type_str:
                                pair_type = 'экз'
                            elif 'з/о' in type_str or 'зач' in type_str:
                                pair_type = 'з/о'
                            elif 'вси' in type_str:
                                pair_type = 'вси'
                        room = ''
                        room_cell = sheet.cell(data_row + 1, col).value if data_row + 1 <= sheet.max_row else None
                        if room_cell:
                            room = str(room_cell).strip()
                        topic_num = ''
                        lesson_num = ''
                        if type_cell:
                            type_match = re.search(r'(\d+)\s*([\d\s]+)?\s*[лпзсгкв]', str(type_cell))
                            if type_match:
                                topic_num = type_match.group(1)
                                lesson_num = type_match.group(2).strip() if type_match.group(2) else ''
                        if len(cell_str) > 30 or cell_str[0].isdigit():
                            continue
                        for group in found_groups:
                            if date_obj not in self.schedule_data[group]:
                                self.schedule_data[group][date_obj] = []
                            exists = False
                            for p in self.schedule_data[group][date_obj]:
                                if p.get('pair_num') == pair_num and p.get('day') == day_name:
                                    exists = True
                                    break
                            if not exists:
                                pair_data = {
                                    'subject': cell_str,
                                    'room': room,
                                    'type': pair_type,
                                    'pair_num': pair_num,
                                    'day': day_name,
                                    'topic_num': topic_num,
                                    'lesson_num': lesson_num
                                }
                                self.schedule_data[group][date_obj].append(pair_data)
    
    def get_schedule_for_group(self, group, target_date=None):
        if not self.excel_loaded:
            return None
        if target_date is None:
            target_date = datetime.now()
        group = str(group).strip()
        if group in self.schedule_data:
            data = self.schedule_data[group]
        else:
            for g in self.schedule_data:
                if group.lower() in g.lower():
                    data = self.schedule_data[g]
                    break
            else:
                return None
        result = []
        for date_key, pairs in data.items():
            if isinstance(date_key, datetime) and date_key.date() == target_date.date():
                result.extend(pairs)
        day_order = {'Пн': 0, 'Вт': 1, 'Ср': 2, 'Чт': 3, 'Пт': 4, 'Сб': 5}
        result.sort(key=lambda x: (day_order.get(x.get('day', 'Пн'), 0), x.get('pair_num', 0)))
        return result
    
    def get_schedule_for_days(self, group, days=2):
        if not self.excel_loaded:
            return {}
        result = {}
        today = datetime.now()
        for i in range(days + 1):
            target_date = today + timedelta(days=i)
            schedule = self.get_schedule_for_group(group, target_date)
            if schedule:
                result[target_date] = schedule
        return result
    
    def find_pair_by_name(self, group, name):
        if not self.excel_loaded:
            return None
        target_group = None
        if group in self.schedule_data:
            target_group = group
        else:
            for g in self.schedule_data:
                if group.lower() in g.lower():
                    target_group = g
                    break
        if not target_group:
            return None
        name_lower = name.lower().strip()
        results = []
        for date_obj, pairs in self.schedule_data[target_group].items():
            for pair in pairs:
                if name_lower in pair.get('subject', '').lower():
                    results.append({'date': date_obj, 'pair': pair})
        results.sort(key=lambda x: x['date'])
        return results
    
    def get_upcoming_warnings(self, group):
        if not self.excel_loaded:
            return []
        target_group = None
        if group in self.schedule_data:
            target_group = group
        else:
            for g in self.schedule_data:
                if group.lower() in g.lower():
                    target_group = g
                    break
        if not target_group:
            return []
        today = datetime.now()
        warnings = []
        for date_obj, pairs in self.schedule_data[target_group].items():
            if not isinstance(date_obj, datetime):
                continue
            days_until = (date_obj.date() - today.date()).days
            for pair in pairs:
                pair_type = pair.get('type', '')
                subject = pair.get('subject', '')
                if days_until == 3 and pair_type in ['пз', 'с']:
                    warnings.append({
                        'date': date_obj,
                        'subject': subject,
                        'type': pair_type,
                        'days_until': days_until,
                        'message': f'⚠️ Через 3 дня: {pair_type.upper()} - {subject}'
                    })
                if days_until == 5 and pair_type in ['экз', 'з/о']:
                    warnings.append({
                        'date': date_obj,
                        'subject': subject,
                        'type': pair_type,
                        'days_until': days_until,
                        'message': f'🔥 Через 5 дней: {pair_type.upper()} - {subject}'
                    })
        return warnings


bot = ScheduleBot()


# ========== КЛАВИАТУРЫ ==========
def get_main_keyboard():
    keyboard = [
        [KeyboardButton("📅 Сегодня"), KeyboardButton("📆 На 2 дня")],
        [KeyboardButton("🔍 По дате"), KeyboardButton("🔎 По предмету")],
        [KeyboardButton("👤 Сменить группу"), KeyboardButton("⚙️ Настройки")],
        [KeyboardButton("📁 Загрузить Excel"), KeyboardButton("ℹ️ Помощь")]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)


def get_settings_keyboard():
    keyboard = [
        [KeyboardButton("📊 Стандартные настройки"), KeyboardButton("✏️ Свои настройки")],
        [KeyboardButton("🔄 Вкл/Выкл уведомления")],
        [KeyboardButton("🔙 Назад")]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)


# ========== ОБРАБОТЧИКИ ==========
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = str(update.effective_user.id)
    if user_id not in bot.user_settings:
        bot.user_settings[user_id] = DEFAULT_SETTINGS.copy()
    if user_id not in bot.user_groups:
        bot.user_groups[user_id] = '20-21'
    group = bot.user_groups[user_id]
    status = "✅ Загружено" if bot.excel_loaded else "⚠️ Файл не загружен"
    welcome_text = f"""
🎓 *БОТ РАСПИСАНИЯ ВУНЦ ВВС*

┌─────────────────────────┐
│ 👤 *Группа:* `{group}`
│ 📊 *Статус:* {status}
│ 📋 *Групп в базе:* {len(bot.groups)}
└─────────────────────────┘

*Доступные команды:*
• 📅 *Сегодня* — пары на сегодня
• 📆 *На 2 дня* — расписание на 2 дня
• 🔍 *По дате* — поиск по дате
• 🔎 *По предмету* — поиск дисциплины
• 👤 *Сменить группу* — выбор группы
• ⚙️ *Настройки* — уведомления
• 📁 *Загрузить Excel* — обновить файл
"""
    await update.message.reply_text(welcome_text, parse_mode='Markdown', reply_markup=get_main_keyboard())


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    if not document.file_name.endswith(('.xlsx', '.xls')):
        await update.message.reply_text("❌ *Ошибка!*\nНужен файл .xlsx или .xls", parse_mode='Markdown')
        return
    msg = await update.message.reply_text("⏳ *Обработка файла...*", parse_mode='Markdown')
    try:
        file = await context.bot.get_file(document.file_id)
        temp_path = f"/tmp/temp_{document.file_name}"
        await file.download_to_drive(temp_path)
        if bot.save_uploaded_excel(temp_path):
            groups_count = len(bot.groups)
            groups_list = ", ".join(sorted(bot.groups.keys()))
            total_pairs = 0
            for g in bot.groups:
                total_pairs += sum(len(p) for p in bot.schedule_data.get(g, {}).values())
            await msg.edit_text(
                f"✅ *Файл успешно загружен!*\n\n"
                f"👥 *Групп:* {groups_count}\n"
                f"📊 *Всего пар:* {total_pairs}\n"
                f"📋 `{groups_list}`",
                parse_mode='Markdown'
            )
        else:
            await msg.edit_text("❌ *Ошибка обработки файла*", parse_mode='Markdown')
    except Exception as e:
        logger.error(f"Ошибка загрузки: {e}")
        await msg.edit_text(f"❌ *Ошибка:* {str(e)[:100]}", parse_mode='Markdown')


async def show_today(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ *Сначала загрузите Excel файл!*", parse_mode='Markdown')
        return
    user_id = str(update.effective_user.id)
    group = bot.user_groups.get(user_id, '20-21')
    target_date = datetime.now()
    schedule = bot.get_schedule_for_group(group, target_date)
    days_ru = ['ПН', 'ВТ', 'СР', 'ЧТ', 'ПТ', 'СБ', 'ВС']
    day_name = days_ru[target_date.weekday()]
    if not schedule:
        await update.message.reply_text(
            f"📅 *{target_date.strftime('%d.%m.%Y')} ({day_name})*\n"
            f"👥 Группа: *{group}*\n\n"
            f"✨ *Выходной! Нет занятий* ✨",
            parse_mode='Markdown'
        )
        return
    text = f"📅 *{target_date.strftime('%d.%m.%Y')} ({day_name})*\n👥 Группа: *{group}*\n\n"
    for item in schedule:
        emoji = {'л': '📖', 'пз': '✏️', 'с': '🗣️', 'гз': '👥', 'кр': '📝', 'экз': '📋', 'з/о': '✅'}.get(item.get('type', 'л'), '📚')
        pair_num = item.get('pair_num', '?')
        subject = item.get('subject', '—')
        room = item.get('room', '')
        room_text = f" 📍*{room}*" if room and room != 'None' else ""
        text += f"{emoji} *П{pair_num}* — *{subject}*{room_text}\n"
    await update.message.reply_text(text, parse_mode='Markdown')


async def show_two_days(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ *Сначала загрузите Excel файл!*", parse_mode='Markdown')
        return
    user_id = str(update.effective_user.id)
    group = bot.user_groups.get(user_id, '20-21')
    schedule_dict = bot.get_schedule_for_days(group, 2)
    if not schedule_dict:
        await update.message.reply_text(f"😴 *Нет занятий на ближайшие 2 дня*\n👥 Группа: *{group}*", parse_mode='Markdown')
        return
    text = f"📆 *РАСПИСАНИЕ НА 2 ДНЯ*\n👥 Группа: *{group}*\n"
    text += "═" * 25 + "\n"
    days_ru = ['ПН', 'ВТ', 'СР', 'ЧТ', 'ПТ', 'СБ', 'ВС']
    for date_obj, schedule in sorted(schedule_dict.items()):
        date_str = date_obj.strftime('%d.%m.%Y')
        day_name = days_ru[date_obj.weekday()]
        text += f"\n📌 *{date_str} ({day_name})*\n"
        text += "─" * 20 + "\n"
        if not schedule:
            text += "  ✨ Выходной\n"
        else:
            for item in schedule:
                emoji = {'л': '📖', 'пз': '✏️', 'с': '🗣️', 'гз': '👥'}.get(item.get('type', 'л'), '📚')
                text += f"  {emoji} *П{item.get('pair_num', '?')}* — {item.get('subject', '—')}\n"
    await update.message.reply_text(text, parse_mode='Markdown')


async def search_by_date_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ Сначала загрузите Excel файл!")
        return ConversationHandler.END
    await update.message.reply_text("📅 *Поиск по дате*\n\nВведите дату в формате:\n`ДД.ММ.ГГГГ`\n\nНапример: `12.01.2026`", parse_mode='Markdown')
    return SEARCH_DATE


async def search_by_date_handle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.message.text.strip()
    try:
        target_date = datetime.strptime(query, '%d.%m.%Y')
    except ValueError:
        await update.message.reply_text("❌ Неверный формат. Используйте: `ДД.ММ.ГГГГ`", parse_mode='Markdown')
        return ConversationHandler.END
    user_id = str(update.effective_user.id)
    group = bot.user_groups.get(user_id, '20-21')
    schedule = bot.get_schedule_for_group(group, target_date)
    days_ru = ['ПН', 'ВТ', 'СР', 'ЧТ', 'ПТ', 'СБ', 'ВС']
    day_name = days_ru[target_date.weekday()]
    if not schedule:
        await update.message.reply_text(f"📅 *{target_date.strftime('%d.%m.%Y')} ({day_name})*\n👥 Группа: *{group}*\n\n✨ *Нет занятий*", parse_mode='Markdown')
    else:
        text = f"📅 *{target_date.strftime('%d.%m.%Y')} ({day_name})*\n👥 *{group}*\n\n"
        for item in schedule:
            emoji = {'л': '📖', 'пз': '✏️', 'с': '🗣️', 'гз': '👥', 'кр': '📝', 'экз': '📋', 'з/о': '✅'}.get(item.get('type', 'л'), '📚')
            text += f"{emoji} *П{item.get('pair_num', '?')}* — {item.get('subject', '—')}\n"
        await update.message.reply_text(text, parse_mode='Markdown')
    return ConversationHandler.END


async def search_by_name_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ Сначала загрузите Excel файл!")
        return ConversationHandler.END
    await update.message.reply_text("🔎 *Поиск дисциплины*\n\nВведите название предмета:\nНапример: `ИАО`, `ФП`, `ТВВС`", parse_mode='Markdown')
    return SEARCH_NAME


async def search_by_name_handle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.message.text.strip()
    user_id = str(update.effective_user.id)
    group = bot.user_groups.get(user_id, '20-21')
    results = bot.find_pair_by_name(group, query)
    if not results:
        await update.message.reply_text(f"❌ *{query}* не найдено\n👥 Группа: *{group}*", parse_mode='Markdown')
    else:
        text = f"🔎 *РЕЗУЛЬТАТЫ ПОИСКА: {query}*\n👥 Группа: *{group}*\n" + "═" * 25 + "\n\n"
        for result in results[:15]:
            date_str = result['date'].strftime('%d.%m.%Y')
            pair = result['pair']
            emoji = {'л': '📖', 'пз': '✏️', 'с': '🗣️', 'гз': '👥', 'кр': '📝', 'экз': '📋', 'з/о': '✅'}.get(pair.get('type', 'л'), '📚')
            text += f"📅 *{date_str}* ({pair.get('day', '')})\n"
            text += f"   {emoji} *П{pair.get('pair_num', '?')}* — {pair.get('subject', '')}\n"
            topic = pair.get('topic_num', '')
            lesson = pair.get('lesson_num', '')
            if topic:
                text += f"   └ Тема {topic} | Занятие {lesson}\n"
            text += f"   └ Тип: {pair.get('type', '—').upper()}\n\n"
        if len(results) > 15:
            text += f"\n... и ещё {len(results) - 15} занятий"
        await update.message.reply_text(text, parse_mode='Markdown')
    return ConversationHandler.END


async def select_group(update: Update, context: ContextTypes.DEFAULT_TYPE):
    available_groups = sorted(bot.groups.keys()) if bot.groups else KNOWN_GROUPS
    keyboard = []
    for i in range(0, len(available_groups), 3):
        row = []
        for g in available_groups[i:i+3]:
            row.append(InlineKeyboardButton(f"👥 {g}", callback_data=f"group_{g}"))
        keyboard.append(row)
    keyboard.append([InlineKeyboardButton("🔙 Отмена", callback_data="cancel")])
    await update.message.reply_text("👤 *Выберите группу:*", parse_mode='Markdown', reply_markup=InlineKeyboardMarkup(keyboard))


async def settings_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = str(update.effective_user.id)
    settings = bot.user_settings.get(user_id, DEFAULT_SETTINGS.copy())
    group = bot.user_groups.get(user_id, '20-21')
    status = "✅ ВКЛ" if settings.get('enabled', True) else "❌ ВЫКЛ"
    text = f"""
⚙️ *НАСТРОЙКИ*

┌─────────────────────────────┐
│ 👤 Группа: `{group}`
│ 📊 Уведомления: {status}
│ 📅 Дней вперед: {settings.get('days_ahead', 2)}
│ ⏰ Время: {settings.get('notify_time', '06:00')}
│ ⚠️ ПЗ/С за: {settings.get('warn_pz_s', 3)} дн.
│ 🔥 Экзамен за: {settings.get('warn_exam', 5)} дн.
└─────────────────────────────┘

Выберите действие:
"""
    await update.message.reply_text(text, parse_mode='Markdown', reply_markup=get_settings_keyboard())


async def toggle_notifications(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = str(update.effective_user.id)
    settings = bot.user_settings.get(user_id, DEFAULT_SETTINGS.copy())
    settings['enabled'] = not settings.get('enabled', True)
    bot.user_settings[user_id] = settings
    status = "✅ ВКЛЮЧЕНЫ" if settings['enabled'] else "❌ ВЫКЛЮЧЕНЫ"
    await update.message.reply_text(f"🔔 Уведомления *{status}*", parse_mode='Markdown')


async def custom_settings_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "✏️ *Настройка параметров*\n\n"
        "Введите через пробел:\n"
        "`[дней] [час:минуты] [дней_до_ПЗ] [дней_до_экзамена]`\n\n"
        "Например: `2 07:00 3 5`",
        parse_mode='Markdown'
    )
    return CUSTOM_SETTINGS


async def custom_settings_handle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        parts = update.message.text.strip().split()
        if len(parts) != 4:
            raise ValueError("Нужно 4 параметра")
        days_ahead = int(parts[0])
        notify_time = parts[1]
        warn_pz = int(parts[2])
        warn_exam = int(parts[3])
        datetime.strptime(notify_time, '%H:%M')
        user_id = str(update.effective_user.id)
        bot.user_settings[user_id] = {
            'days_ahead': days_ahead,
            'notify_time': notify_time,
            'warn_pz_s': warn_pz,
            'warn_exam': warn_exam,
            'enabled': True
        }
        await update.message.reply_text("✅ *Настройки сохранены!*", parse_mode='Markdown', reply_markup=get_main_keyboard())
    except Exception as e:
        await update.message.reply_text(f"❌ *Ошибка!*\n{str(e)}", parse_mode='Markdown')
    return ConversationHandler.END


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    help_text = """
ℹ️ *ПОМОЩЬ*

*Основные функции:*
• 📅 *Сегодня* — расписание на сегодня
• 📆 *На 2 дня* — расписание на 2 дня
• 🔍 *По дате* — поиск по конкретной дате
• 🔎 *По предмету* — поиск всех занятий по дисциплине

*Настройки:*
• Уведомления приходят в указанное время
• Можно настроить за сколько дней предупреждать о ПЗ/С и экзаменах
"""
    await update.message.reply_text(help_text, parse_mode='Markdown')


async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    user_id = str(update.effective_user.id)
    if data == "cancel":
        await query.edit_message_text("❌ Отменено")
        return
    if data.startswith("group_"):
        group_code = data[6:]
        bot.user_groups[user_id] = group_code
        await query.edit_message_text(f"✅ Группа изменена на *{group_code}*", parse_mode='Markdown')


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    if text == "📅 Сегодня":
        await show_today(update, context)
    elif text == "📆 На 2 дня":
        await show_two_days(update, context)
    elif text == "🔍 По дате":
        await search_by_date_start(update, context)
    elif text == "🔎 По предмету":
        await search_by_name_start(update, context)
    elif text == "👤 Сменить группу":
        await select_group(update, context)
    elif text == "⚙️ Настройки":
        await settings_menu(update, context)
    elif text == "📊 Стандартные настройки":
        user_id = str(update.effective_user.id)
        bot.user_settings[user_id] = DEFAULT_SETTINGS.copy()
        await update.message.reply_text("✅ *Стандартные настройки применены*", parse_mode='Markdown', reply_markup=get_main_keyboard())
    elif text == "🔄 Вкл/Выкл уведомления":
        await toggle_notifications(update, context)
    elif text == "🔙 Назад":
        await update.message.reply_text("Главное меню", reply_markup=get_main_keyboard())
    elif text == "📁 Загрузить Excel":
        await update.message.reply_text("📁 Отправьте Excel файл с расписанием")
    elif text == "ℹ️ Помощь":
        await help_command(update, context)


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("❌ Отменено", reply_markup=get_main_keyboard())
    return ConversationHandler.END


async def daily_notification(context: ContextTypes.DEFAULT_TYPE):
    for user_id, settings in bot.user_settings.items():
        if not settings.get('enabled', True):
            continue
        try:
            group = bot.user_groups.get(user_id, '20-21')
            days_ahead = settings.get('days_ahead', 2)
            schedule_dict = bot.get_schedule_for_days(group, days_ahead)
            if not schedule_dict:
                continue
            text = f"🌅 *ДОБРОЕ УТРО!*\n📆 Расписание на {days_ahead} дн.\n👥 Группа: *{group}*\n" + "═" * 25 + "\n"
            days_ru = ['ПН', 'ВТ', 'СР', 'ЧТ', 'ПТ', 'СБ', 'ВС']
            for date_obj, schedule in sorted(schedule_dict.items())[:days_ahead]:
                date_str = date_obj.strftime('%d.%m.%Y')
                day_name = days_ru[date_obj.weekday()]
                text += f"\n📌 *{date_str} ({day_name})*\n"
                if not schedule:
                    text += "  ✨ Выходной\n"
                else:
                    for item in schedule:
                        emoji = {'л': '📖', 'пз': '✏️', 'с': '🗣️', 'гз': '👥'}.get(item.get('type', 'л'), '📚')
                        text += f"  {emoji} *П{item.get('pair_num', '?')}* — {item.get('subject', '—')}\n"
            warnings = bot.get_upcoming_warnings(group)
            if warnings:
                text += "\n⚠️ *ВАЖНЫЕ НАПОМИНАНИЯ:*\n"
                for w in warnings[:3]:
                    text += f"• {w['message']}\n"
            await context.bot.send_message(chat_id=int(user_id), text=text, parse_mode='Markdown')
        except Exception as e:
            logger.error(f"Ошибка отправки уведомления {user_id}: {e}")


def main():
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    if not token:
        logger.error("❌ TELEGRAM_BOT_TOKEN не найден!")
        return
    
    logger.info(f"🚀 Запуск бота... Директория: {os.getcwd()}")
    
    app = Application.builder().token(token).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    
    date_conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex(r'^🔍 По дате$'), search_by_date_start)],
        states={SEARCH_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, search_by_date_handle)]},
        fallbacks=[CommandHandler("cancel", cancel)]
    )
    
    name_conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex(r'^🔎 По предмету$'), search_by_name_start)],
        states={SEARCH_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, search_by_name_handle)]},
        fallbacks=[CommandHandler("cancel", cancel)]
    )
    
    custom_conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex(r'^✏️ Свои настройки$'), custom_settings_start)],
        states={CUSTOM_SETTINGS: [MessageHandler(filters.TEXT & ~filters.COMMAND, custom_settings_handle)]},
        fallbacks=[CommandHandler("cancel", cancel)]
    )
    
    app.add_handler(date_conv)
    app.add_handler(name_conv)
    app.add_handler(custom_conv)
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    
    job_queue = app.job_queue
    if job_queue:
        notify_time = time(6, 0)
        job_queue.run_daily(daily_notification, notify_time)
        logger.info(f"✅ Уведомления настроены на {notify_time}")
    
    logger.info("✅ Бот запущен!")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
