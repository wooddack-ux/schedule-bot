import logging
import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler, CallbackQueryHandler
from datetime import datetime, timedelta
import os
import re
from calendar import monthrange

# Настройка логирования
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Состояния для ConversationHandler
SELECT_GROUP, SETTINGS, SEARCH_DATE, SEARCH_NAME, CUSTOM_SETTINGS = range(5)

# Маппинг русских месяцев
MONTHS_RU = {
    'январь': 1, 'февраль': 2, 'март': 3, 'апрель': 4, 'май': 5, 'июнь': 6,
    'июль': 7, 'август': 8, 'сентябрь': 9, 'октябрь': 10, 'ноябрь': 11, 'декабрь': 12
}

class ScheduleBot:
    def __init__(self):
        self.excel_file = None
        self.workbook = None
        self.schedule_data = {}  # {group: {date: [pairs]}}
        self.groups = {}
        self.user_settings = {}
        self.excel_loaded = False
        self.load_excel_on_startup()
    
    def load_excel_on_startup(self):
        """Загружает Excel файл при старте бота"""
        try:
            files = [f for f in os.listdir('.') if f.endswith(('.xlsx', '.xls'))]
            if not files:
                logger.warning("⚠️ Excel файл не найден!")
                return False
                
            self.excel_file = files[0]
            self.workbook = openpyxl.load_workbook(self.excel_file, data_only=True)
            self.parse_schedule()
            self.excel_loaded = True
            logger.info(f"✅ Загружен файл: {self.excel_file}")
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка загрузки Excel: {e}")
            self.excel_loaded = False
            return False
    
    def save_uploaded_excel(self, file_path):
        """Сохраняет загруженный Excel файл"""
        try:
            for f in os.listdir('.'):
                if f.endswith(('.xlsx', '.xls')):
                    os.remove(f)
            
            new_name = "schedule.xlsx"
            os.rename(file_path, new_name)
            
            self.excel_file = new_name
            self.workbook = openpyxl.load_workbook(self.excel_file, data_only=True)
            self.parse_schedule()
            self.excel_loaded = True
            logger.info(f"✅ Загружен новый файл: {new_name}")
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка сохранения Excel: {e}")
            return False
    
    def parse_schedule(self):
        """Парсит расписание из всех листов"""
        self.schedule_data = {}
        self.groups = {}
        
        year = 2025  # Учебный год 2025/2026
        
        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]
            logger.info(f"Обработка листа: {sheet_name}")
            
            # Ищем структуру с неделями и днями
            self._parse_sheet_vunc(sheet, year)
        
        logger.info(f"Найдено групп: {len(self.groups)}")
        for g in self.groups:
            logger.info(f"  - {g}: {len(self.schedule_data.get(g, {}))} дат")
    
    def _parse_sheet_vunc(self, sheet, year):
        """Парсит лист ВУНЦ (твоя структура)"""
        max_row = sheet.max_row
        max_col = sheet.max_column
        
        # Ищем строки с "ЯНВАРЬ", "ФЕВРАЛЬ" и т.д.
        row = 1
        current_month = None
        current_week = None
        
        while row <= max_row:
            cell_value = sheet.cell(row, 1).value
            
            # Проверяем, это месяц?
            if cell_value and isinstance(cell_value, str):
                month_str = cell_value.lower().strip()
                if month_str in MONTHS_RU:
                    current_month = MONTHS_RU[month_str]
                    # Следующая ячейка должна содержать число (номер недели или день)
                    day_cell = sheet.cell(row, 3).value
                    if day_cell and str(day_cell).isdigit():
                        day = int(day_cell)
                        self._process_day(sheet, row, current_month, year, day)
            
            row += 1
    
    def _process_day(self, sheet, start_row, month, year, day):
        """Обрабатывает один день расписания"""
        try:
            # Проверяем валидность даты
            if not (1 <= day <= 31):
                return
            
            date_obj = datetime(year, month, day)
            
            # Ищем строки с группами
            # Обычно структура: строка с номером недели, потом строки с группами
            
            row = start_row + 1
            
            # Пропускаем строки до первой группы
            while row <= sheet.max_row and row < start_row + 15:
                val = sheet.cell(row, 1).value
                if val and isinstance(val, str) and ('-' in str(val) or 'спец' in str(val)):
                    # Нашли группу или блок групп
                    group_block = str(val).strip()
                    
                    # Обрабатываем блок групп (может быть несколько групп подряд)
                    groups_in_block = self._extract_groups(group_block)
                    
                    for group in groups_in_block:
                        if group not in self.schedule_data:
                            self.schedule_data[group] = {}
                            self.groups[group] = group
                        
                        # Парсим пары для этой группы
                        pairs = self._extract_pairs_for_group(sheet, row, date_obj)
                        if pairs:
                            if date_obj not in self.schedule_data[group]:
                                self.schedule_data[group][date_obj] = []
                            self.schedule_data[group][date_obj].extend(pairs)
                    
                    row += 3  # Пропускаем строки этой группы
                else:
                    row += 1
                    
        except ValueError as e:
            logger.debug(f"Неверная дата: {day}.{month}.{year}")
    
    def _extract_groups(self, block_text):
        """Извлекает список групп из текста блока"""
        groups = []
        
        # Форматы: "20-21", "20-22", "20и-21", "26-21", "7-21", "8-21", "8и-21"
        patterns = [
            r'(\d{2}-?\d{2})',  # 20-21, 26-21
            r'(\d{2}и-?\d{2})',  # 20и-21
            r'(\d{1}и?-?\d{2})',  # 7-21, 8-21, 8и-21
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, block_text)
            for match in matches:
                clean = match.replace('-', '-').strip()
                if clean and len(clean) >= 3:
                    groups.append(clean)
        
        # Убираем дубликаты
        return list(set(groups))
    
    def _extract_pairs_for_group(self, sheet, row, date_obj):
        """Извлекает пары для группы из строки"""
        pairs = []
        
        # Колонки пар: 2-3 (1-2), 4-5 (3-4), 6-7 (5-6) и т.д.
        # Смотрим на структуру: номер пары, тип, предмет, аудитория
        
        pair_positions = [
            (2, 3),   # Пара 1-2
            (4, 5),   # Пара 3-4  
            (6, 7),   # Пара 5-6
            (8, 9),   # Пара 1-2 вторник
            (10, 11), # Пара 3-4 вторник
            (12, 13), # Пара 5-6 вторник
            (14, 15), # Среда 1-2
            (16, 17), # Среда 3-4
            (18, 19), # Среда 5-6
            (20, 21), # Четверг 1-2
            (22, 23), # Четверг 3-4
            (24, 25), # Четверг 5-6
            (26, 27), # Пятница 1-2
            (28, 29), # Пятница 3-4
            (30, 31), # Пятница 5-6
            (32, 33), # Суббота 1-2
            (34, 35), # Суббота 3-4
            (36, 37), # Суббота 5-6
        ]
        
        pair_num = 1
        for col_pair in pair_positions:
            col1, col2 = col_pair
            
            val1 = sheet.cell(row, col1).value
            val2 = sheet.cell(row, col2).value
            
            if val1 and str(val1).strip():
                pair_info = self._parse_pair_cell(str(val1), str(val2) if val2 else "")
                if pair_info:
                    pair_info['pair_num'] = pair_num
                    pairs.append(pair_info)
            
            pair_num += 1
        
        return pairs
    
    def _parse_pair_cell(self, cell_text, next_cell):
        """Парсит ячейку с парой"""
        cell_text = str(cell_text).strip()
        if not cell_text or cell_text in ['None', '']:
            return None
        
        # Формат: "1 1 л" или "2 12 л" или "5 5 пз" или просто "ИАО"
        
        # Проверяем, начинается ли с номера пары
        pair_match = re.match(r'(\d+)\s+(\d+)\s+([лпзсгкэо/]+)', cell_text)
        
        if pair_match:
            # Формат "1 1 л" - номер пары, номер занятия, тип
            pair_num = pair_match.group(1)
            lesson_num = pair_match.group(2)
            pair_type_short = pair_match.group(3)
            
            # Следующая ячейка - название предмета
            subject = next_cell if next_cell else "Не указано"
            
            # Определяем тип занятия
            type_map = {
                'л': ('лекция', '📖'),
                'пз': ('практика', '🔧'),
                'с': ('семинар', '🗣️'),
                'гз': ('групповое', '🏫'),
                'кр': ('контр.работа', '📝'),
                'экз': ('экзамен', '📋'),
                'з/о': ('зачёт', '✅'),
                'вси': ('игра', '🔬'),
                'гз': ('групповое', '👥'),
            }
            
            pair_type, emoji = type_map.get(pair_type_short, ('занятие', '📚'))
            
            return {
                'raw': cell_text,
                'subject': subject,
                'type': pair_type_short,
                'type_full': pair_type,
                'emoji': emoji,
                'pair_num': pair_num
            }
        else:
            # Просто название предмета без номера
            return {
                'raw': cell_text,
                'subject': cell_text,
                'type': 'л',
                'type_full': 'занятие',
                'emoji': '📚',
                'pair_num': '1'
            }
    
    def get_schedule_for_group(self, group: str, target_date: datetime = None):
        """Получает расписание для группы на указанную дату"""
        if not self.excel_loaded:
            return None
            
        if target_date is None:
            target_date = datetime.now()
        
        # Нормализуем группу
        group = str(group).strip()
        
        # Ищем точное совпадение или похожее
        if group in self.schedule_data:
            data = self.schedule_data[group]
        else:
            # Пробуем найти похожую группу
            for g in self.schedule_data:
                if group in g or g in group:
                    data = self.schedule_data[g]
                    group = g
                    break
            else:
                return None
        
        # Ищем дату
        for date_key, pairs in data.items():
            if isinstance(date_key, datetime) and date_key.date() == target_date.date():
                return pairs
        
        return []
    
    def find_pair_by_name(self, group: str, name: str):
        """Ищет пару по названию дисциплины"""
        if not self.excel_loaded or group not in self.schedule_data:
            return None
        
        name_lower = name.lower()
        data = self.schedule_data[group]
        
        for date_obj, pairs in data.items():
            for pair in pairs:
                if name_lower in pair.get('subject', '').lower():
                    return {
                        'date': date_obj,
                        'pair': pair
                    }
        return None
    
    def get_upcoming_exams(self, group: str, days_ahead: int = 14):
        """Находит ближайшие зачёты/экзамены"""
        if not self.excel_loaded or group not in self.schedule_data:
            return []
        
        today = datetime.now()
        exams = []
        data = self.schedule_data[group]
        
        for date_obj, pairs in data.items():
            if not isinstance(date_obj, datetime):
                continue
            
            days_diff = (date_obj.date() - today.date()).days
            if 0 <= days_diff <= days_ahead:
                for pair in pairs:
                    pair_type = pair.get('type', '')
                    if pair_type in ['экз', 'з/о', 'кр', 'зач']:
                        exams.append({
                            'date': date_obj,
                            'subject': pair.get('subject', ''),
                            'type': pair_type,
                            'days_until': days_diff
                        })
        
        return sorted(exams, key=lambda x: (x['days_until'], x['date']))

# Инициализация бота
bot = ScheduleBot()

def get_main_keyboard():
    keyboard = [
        [KeyboardButton("📅 Расписание на сегодня"), KeyboardButton("📆 Расписание на 2 дня")],
        [KeyboardButton("🔍 Поиск по дате"), KeyboardButton("🔎 Поиск по названию")],
        [KeyboardButton("⚙️ Настройки"), KeyboardButton("📊 Экзамены/Зачёты")],
        [KeyboardButton("📁 Загрузить расписание")]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Стартовая команда"""
    user_id = update.effective_user.id
    
    if str(user_id) not in bot.user_settings:
        bot.user_settings[str(user_id)] = {
            'group': '20-21',
            'days_ahead': 2,
            'exam_warning_days': [3, 5],
            'notify_time': '06:00',
            'enabled': True
        }
    
    settings = bot.user_settings.get(str(user_id), {})
    current_group = settings.get('group', '20-21')
    
    status = "✅ Расписание загружено" if bot.excel_loaded else "⚠️ Ожидание загрузки расписания"
    groups_list = ", ".join(sorted(bot.groups.keys())[:10]) + "..." if len(bot.groups) > 10 else ", ".join(sorted(bot.groups.keys()))
    
    welcome_text = f"""
🎓 *Бот расписания ВУНЦ ВВС*

👤 Группа: *{current_group}*
📊 Статус: {status}
📋 Группы: `{groups_list}`

⚙️ Настройки: /settings

📋 *Доступные команды:*
• 📅 Расписание на сегодня — показать пары на сегодня
• 📆 Расписание на 2 дня — показать на сегодня + 2 дня вперёд  
• 🔍 Поиск по дате — найти пары на конкретную дату
• 🔎 Поиск по названию — найти пару по названию дисциплины
• 📊 Экзамены/Зачёты — ближайшие контрольные/зачёты
• 📁 Загрузить расписание — отправить Excel файл боту
• ⚙️ Настройки — Настроить группу и уведомления
"""
    
    await update.message.reply_text(
        welcome_text, 
        parse_mode='Markdown',
        reply_markup=get_main_keyboard()
    )

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка загруженного Excel файла"""
    document = update.message.document
    
    if not document.file_name.endswith(('.xlsx', '.xls')):
        await update.message.reply_text("❌ Нужен файл Excel (.xlsx или .xls)")
        return
    
    await update.message.reply_text("⏳ Скачиваю и обрабатываю файл...")
    
    try:
        file = await context.bot.get_file(document.file_id)
        temp_path = f"temp_{document.file_name}"
        await file.download_to_drive(temp_path)
        
        if bot.save_uploaded_excel(temp_path):
            groups_count = len(bot.groups)
            await update.message.reply_text(
                "✅ *Расписание успешно загружено!*\n\n"
                f"📁 Файл: `{document.file_name}`\n"
                f"👥 Найдено групп: *{groups_count}*\n"
                f"📊 Листов: {len(bot.workbook.sheetnames)}\n\n"
                f"Доступные группы: `{', '.join(sorted(bot.groups.keys()))}`\n\n"
                "Теперь можно смотреть расписание!",
                parse_mode='Markdown',
                reply_markup=get_main_keyboard()
            )
        else:
            await update.message.reply_text("❌ Ошибка при обработке файла")
            
    except Exception as e:
        logger.error(f"Ошибка загрузки файла: {e}")
        await update.message.reply_text(f"❌ Ошибка: {str(e)}")

async def show_schedule(update: Update, context: ContextTypes.DEFAULT_TYPE, days_offset: int = 0):
    """Показывает расписание"""
    if not bot.excel_loaded:
        await update.message.reply_text(
            "⚠️ *Расписание не загружено!*\n\n"
            "Отправь Excel файл боту через кнопку 📁 Загрузить расписание",
            parse_mode='Markdown'
        )
        return
    
    user_id = update.effective_user.id
    settings = bot.user_settings.get(str(user_id), {})
    group = settings.get('group', '20-21')
    
    target_date = datetime.now() + timedelta(days=days_offset)
    schedule = bot.get_schedule_for_group(group, target_date)
    
    days_text = "сегодня" if days_offset == 0 else f"+{days_offset} дня"
    header = f"📅 *Расписание на {target_date.strftime('%d.%m.%Y')} ({days_text})*\n"
    header += f"👥 Группа: *{group}*\n\n"
    
    if not schedule:
        text = header + "😴 *Выходной день или пар нет!*"
    else:
        text = header
        for item in schedule:
            emoji = item.get('emoji', '📚')
            pair_num = item.get('pair_num', '?')
            subject = item.get('subject', 'Не указано')
            pair_type = item.get('type_full', 'занятие')
            
            text += f"{emoji} *Пара {pair_num}* — {subject} ({pair_type})\n"
    
    await update.message.reply_text(text, parse_mode='Markdown')

async def search_by_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Поиск по дате"""
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ Сначала загрузите расписание!")
        return
    
    await update.message.reply_text(
        "📅 *Введите дату для поиска*\n\n"
        "Формат: `ДД.ММ.ГГГГ` (например: `20.05.2025`)",
        parse_mode='Markdown'
    )
    return SEARCH_DATE

async def handle_date_search(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка поиска по дате"""
    query = update.message.text.strip()
    
    try:
        target_date = datetime.strptime(query, '%d.%m.%Y')
    except ValueError:
        await update.message.reply_text(
            "❌ *Неверный формат даты!*\nИспользуйте: `ДД.ММ.ГГГГ`",
            parse_mode='Markdown'
        )
        return ConversationHandler.END
    
    user_id = update.effective_user.id
    settings = bot.user_settings.get(str(user_id), {})
    group = settings.get('group', '20-21')
    
    schedule = bot.get_schedule_for_group(group, target_date)
    
    if not schedule:
        await update.message.reply_text(
            f"❌ На *{target_date.strftime('%d.%m.%Y')}* пар не найдено для группы {group}",
            parse_mode='Markdown'
        )
    else:
        text = f"📅 *Пары на {target_date.strftime('%d.%m.%Y')}:*\n\n"
        for item in schedule:
            emoji = item.get('emoji', '📚')
            pair_num = item.get('pair_num', '?')
            subject = item.get('subject', 'Не указано')
            pair_type = item.get('type_full', 'занятие')
            
            text += f"{emoji} *Пара {pair_num}* — {subject} ({pair_type})\n"
        
        await update.message.reply_text(text, parse_mode='Markdown')
    
    return ConversationHandler.END

async def search_by_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Поиск по названию дисциплины"""
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ Сначала загрузите расписание!")
        return
    
    await update.message.reply_text(
        "🔎 *Введите название дисциплины для поиска:*\n"
        "(например: ИАО, ТВВС, СУВ, ОПС)",
        parse_mode='Markdown'
    )
    return SEARCH_NAME

async def handle_name_search(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка поиска по названию"""
    query = update.message.text.strip()
    user_id = update.effective_user.id
    settings = bot.user_settings.get(str(user_id), {})
    group = settings.get('group', '20-21')
    
    result = bot.find_pair_by_name(group, query)
    
    if not result:
        await update.message.reply_text(
            f"❌ Дисциплина *{query}* не найдена в группе {group}",
            parse_mode='Markdown'
        )
    else:
        date_str = result['date'].strftime('%d.%m.%Y')
        pair = result['pair']
        
        await update.message.reply_text(
            f"🔎 *Найдено:*\n\n"
            f"📅 {date_str}\n"
            f"⏰ Пара {pair.get('pair_num', '?')}\n"
            f"📖 {pair.get('subject', 'Не указано')}\n"
            f"📝 Тип: {pair.get('type_full', 'занятие')}",
            parse_mode='Markdown'
        )
    
    return ConversationHandler.END

async def show_exams(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Показывает ближайшие экзамены и зачёты"""
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ Сначала загрузите расписание!")
        return
    
    user_id = update.effective_user.id
    settings = bot.user_settings.get(str(user_id), {})
    group = settings.get('group', '20-21')
    warning_days = settings.get('exam_warning_days', [3, 5])
    
    exams = bot.get_upcoming_exams(group, days_ahead=max(warning_days) + 7)
    
    if not exams:
        await update.message.reply_text(
            "✅ *Ближайшие контрольные работы и зачёты не найдены*",
            parse_mode='Markdown'
        )
        return
    
    text = "📊 *Ближайшие контрольные и зачёты:*\n\n"
    today = datetime.now()
    
    for exam in exams[:15]:  # Показываем максимум 15
        days_until = exam['days_until']
        date_str = exam['date'].strftime('%d.%m.%Y')
        
        warning = " ⚠️" if days_until in warning_days else ""
        days_text = f"через {days_until} дн." if days_until > 0 else "сегодня!"
        
        text += f"📅 {date_str} ({days_text}){warning}\n"
        text += f"📝 {exam['subject']} ({exam['type']})\n\n"
    
    await update.message.reply_text(text, parse_mode='Markdown')

async def settings_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Меню настроек"""
    user_id = update.effective_user.id
    settings = bot.user_settings.get(str(user_id), {})
    
    current_group = settings.get('group', '20-21')
    
    # Получаем список групп для кнопок
    available_groups = sorted(bot.groups.keys()) if bot.groups else ['20-21', '20-22', '26-21', '7-21', '8-21']
    
    keyboard = []
    # Добавляем кнопки групп (по 2 в ряд)
    for i in range(0, len(available_groups), 2):
        row = []
        for g in available_groups[i:i+2]:
            row.append(InlineKeyboardButton(f"👥 {g}", callback_data=f"group_{g}"))
        keyboard.append(row)
    
    keyboard.append([InlineKeyboardButton("🔙 Назад", callback_data="back_to_main")])
    
    await update.message.reply_text(
        "⚙️ *Настройки:*\n\n"
        f"👥 Текущая группа: *{current_group}*\n\n"
        "Выберите группу:",
        parse_mode='Markdown',
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик кнопок"""
    query = update.callback_query
    await query.answer()
    data = query.data
    user_id = update.effective_user.id
    
    if data == "back_to_main":
        await query.edit_message_text("Главное меню")
        return
    
    if data.startswith("group_"):
        group_code = data[6:]
        bot.user_settings[str(user_id)]['group'] = group_code
        await query.edit_message_text(f"✅ Группа изменена на: *{group_code}*", parse_mode='Markdown')
        return

async def morning_job(context: ContextTypes.DEFAULT_TYPE):
    """Утренняя задача для отправки расписания"""
    if not bot.excel_loaded:
        return
    
    for uid, settings in bot.user_settings.items():
        if settings.get('enabled', False):
            try:
                chat_id = int(uid)
                group = settings.get('group', '20-21')
                
                target_date = datetime.now()
                schedule = bot.get_schedule_for_group(group, target_date)
                
                if schedule:
                    text = f"📅 *Расписание на сегодня*\n\n"
                    for item in schedule:
                        emoji = item.get('emoji', '📚')
                        pair_num = item.get('pair_num', '?')
                        subject = item.get('subject', 'Не указано')
                        pair_type = item.get('type_full', 'занятие')
                        
                        text += f"{emoji} *Пара {pair_num}* — {subject} ({pair_type})\n"
                    
                    await context.bot.send_message(chat_id=chat_id, text=text, parse_mode='Markdown')
            except Exception as e:
                logger.error(f"Ошибка отправки: {e}")

def main():
    """Запуск бота"""
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    if not token:
        logger.error("❌ Не найден TELEGRAM_BOT_TOKEN!")
        print("❌ Установите переменную окружения TELEGRAM_BOT_TOKEN")
        return
    
    logger.info(f"🚀 Запуск бота...")
    
    application = Application.builder().token(token).build()
    
    # Обработчики
    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.Document.FileExtension("xlsx") | filters.Document.FileExtension("xls"), handle_document))
    application.add_handler(MessageHandler(filters.Regex(r'^📅'), lambda u, c: show_schedule(u, c, 0)))
    application.add_handler(MessageHandler(filters.Regex(r'^📆'), lambda u, c: show_schedule(u, c, 2)))
    application.add_handler(MessageHandler(filters.Regex(r'^📊'), show_exams))
    application.add_handler(MessageHandler(filters.Regex(r'^⚙️'), settings_menu))
    application.add_handler(MessageHandler(filters.Regex(r'^📁'), lambda u, c: u.message.reply_text("Отправьте Excel файл как документ")))
    application.add_handler(CallbackQueryHandler(button_handler))
    
    # Conversation handlers
    date_conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex(r'^🔍'), search_by_date)],
        states={SEARCH_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_date_search)]},
        fallbacks=[CommandHandler("start", start)]
    )
    
    name_conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex(r'^🔎'), search_by_name)],
        states={SEARCH_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_name_search)]},
        fallbacks=[CommandHandler("start", start)]
    )
    
    application.add_handler(date_conv)
    application.add_handler(name_conv)
    
    # JobQueue
    job_queue = application.job_queue
    if job_queue:
        job_queue.run_daily(morning_job, time=datetime.time(hour=6, minute=0))
        logger.info("✅ Утренняя отправка настроена на 06:00")
    
    logger.info("🚀 Бот запущен!")
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
