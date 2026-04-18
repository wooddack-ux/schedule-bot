import logging
import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler, CallbackQueryHandler
from datetime import datetime, timedelta
import os
import re
import shutil
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
        self.schedule_data = {}
        self.groups = {}
        self.user_settings = {}
        self.excel_loaded = False
        self.load_excel_on_startup()
    
    def load_excel_on_startup(self):
        """Загружает Excel файл при старте бота"""
        try:
            work_dir = os.getcwd()
            files = [f for f in os.listdir(work_dir) if f.endswith(('.xlsx', '.xls')) and not f.startswith('temp_')]
            
            if not files:
                logger.warning("⚠️ Excel файл не найден в рабочей директории!")
                return False
            
            self.excel_file = files[0]
            file_path = os.path.join(work_dir, self.excel_file)
            logger.info(f"Найден файл: {file_path}")
            
            self.workbook = openpyxl.load_workbook(file_path, data_only=True)
            self.parse_schedule()
            self.excel_loaded = True
            logger.info(f"✅ Загружен файл: {self.excel_file}, групп: {len(self.groups)}")
            return True
            
        except Exception as e:
            logger.error(f"❌ Ошибка загрузки Excel: {e}")
            import traceback
            logger.error(traceback.format_exc())
            self.excel_loaded = False
            return False
    
    def save_uploaded_excel(self, file_path):
        """Сохраняет загруженный Excel файл"""
        try:
            work_dir = os.getcwd()
            logger.info(f"Рабочая директория: {work_dir}")
            logger.info(f"Исходный файл: {file_path}, существует: {os.path.exists(file_path)}")
            
            # Удаляем старый schedule.xlsx
            old_file = os.path.join(work_dir, "schedule.xlsx")
            if os.path.exists(old_file):
                try:
                    os.remove(old_file)
                    logger.info(f"Удалён старый файл: {old_file}")
                except Exception as e:
                    logger.warning(f"Не удалось удалить старый файл: {e}")
            
            # Новый путь
            new_path = os.path.join(work_dir, "schedule.xlsx")
            
            # Копируем файл
            if os.path.exists(file_path):
                shutil.copy2(file_path, new_path)
                logger.info(f"Файл скопирован: {new_path}, размер: {os.path.getsize(new_path)}")
            else:
                logger.error(f"Исходный файл не найден: {file_path}")
                return False
            
            # Удаляем временный
            try:
                os.remove(file_path)
                logger.info(f"Временный файл удалён: {file_path}")
            except Exception as e:
                logger.warning(f"Не удалось удалить временный файл: {e}")
            
            # Перезагружаем
            self.excel_file = "schedule.xlsx"
            self.workbook = openpyxl.load_workbook(new_path, data_only=True)
            self.parse_schedule()
            self.excel_loaded = True
            
            logger.info(f"✅ Файл успешно обработан, групп: {len(self.groups)}")
            return True
            
        except Exception as e:
            logger.error(f"❌ Ошибка сохранения Excel: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def parse_schedule(self):
        """Парсит расписание из всех листов"""
        self.schedule_data = {}
        self.groups = {}
        
        year = 2025
        
        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]
            logger.info(f"Обработка листа: {sheet_name}")
            self._parse_sheet_vunc(sheet, year)
        
        logger.info(f"Всего найдено групп: {len(self.groups)}")
        for g in sorted(self.groups.keys())[:10]:
            count = len(self.schedule_data.get(g, {}))
            logger.info(f"  - {g}: {count} дат")
    
    def _parse_sheet_vunc(self, sheet, year):
        """Парсит лист ВУНЦ"""
        max_row = sheet.max_row
        row = 1
        
        while row <= max_row:
            try:
                cell_value = sheet.cell(row, 1).value
                
                if cell_value and isinstance(cell_value, str):
                    month_str = cell_value.lower().strip()
                    
                    if month_str in MONTHS_RU:
                        current_month = MONTHS_RU[month_str]
                        day_cell = sheet.cell(row, 3).value
                        
                        if day_cell:
                            try:
                                day = int(day_cell)
                                if 1 <= day <= 31:
                                    self._process_day(sheet, row, current_month, year, day)
                            except (ValueError, TypeError):
                                pass
            except Exception as e:
                logger.debug(f"Ошибка в строке {row}: {e}")
            
            row += 1
    
    def _process_day(self, sheet, start_row, month, year, day):
        """Обрабатывает один день"""
        try:
            date_obj = datetime(year, month, day)
            row = start_row + 1
            
            while row <= min(start_row + 25, sheet.max_row):
                val = sheet.cell(row, 1).value
                
                if val and isinstance(val, str):
                    val_str = str(val).strip()
                    
                    # Ищем группы по паттернам
                    if any(c in val_str for c in ['20-', '26-', '11-', '7-', '8-', 'спец']):
                        groups = self._extract_groups(val_str)
                        
                        for group in groups:
                            if group not in self.schedule_data:
                                self.schedule_data[group] = {}
                                self.groups[group] = group
                            
                            pairs = self._extract_pairs_for_group(sheet, row, date_obj)
                            if pairs:
                                if date_obj not in self.schedule_data[group]:
                                    self.schedule_data[group][date_obj] = []
                                self.schedule_data[group][date_obj].extend(pairs)
                        
                        row += 3  # Пропускаем строки этой группы
                    else:
                        row += 1
                else:
                    row += 1
                    
        except ValueError as e:
            logger.debug(f"Неверная дата: {day}.{month}.{year} - {e}")
    
    def _extract_groups(self, block_text):
        """Извлекает группы из текста"""
        groups = []
        text = str(block_text).lower()
        
        # Паттерны для групп ВУНЦ
        patterns = [
            (r'(\d{2})-(\d{2})', lambda m: f"{m.group(1)}-{m.group(2)}"),  # 20-21
            (r'(\d{2})и-(\d{2})', lambda m: f"{m.group(1)}и-{m.group(2)}"),  # 20и-21
            (r'([78])-(\d{2})', lambda m: f"{m.group(1)}-{m.group(2)}"),  # 7-21, 8-21
            (r'([78])и-(\d{2})', lambda m: f"{m.group(1)}и-{m.group(2)}"),  # 8и-21
        ]
        
        for pattern, formatter in patterns:
            for match in re.finditer(pattern, text):
                group = formatter(match)
                if group not in groups:
                    groups.append(group)
        
        return groups
    
    def _extract_pairs_for_group(self, sheet, row, date_obj):
        """Извлекает пары для группы"""
        pairs = []
        
        # Колонки B-AK (2-37) — пары по дням недели
        # Пн(2-7), Вт(8-13), Ср(14-19), Чт(20-25), Пт(26-31), Сб(32-37)
        
        days = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб']
        day_start_cols = [2, 8, 14, 20, 26, 32]  # Начальные колонки для каждого дня
        
        for day_idx, start_col in enumerate(day_start_cols):
            for pair_idx, col_offset in enumerate([0, 2, 4]):  # 1-2, 3-4, 5-6 пары
                col = start_col + col_offset
                
                try:
                    # Ячейка с номером пары и типом (например "1 1 л")
                    cell_val = sheet.cell(row, col).value
                    
                    if cell_val and str(cell_val).strip() not in ['', 'None', 'н/д', '-']:
                        val_str = str(cell_val).strip()
                        
                        # Парсим тип занятия из строки
                        pair_type = 'л'  # По умолчанию лекция
                        if 'пз' in val_str.lower():
                            pair_type = 'пз'
                        elif 'с' in val_str.lower() and 'ср' not in val_str.lower():
                            pair_type = 'с'
                        elif 'гз' in val_str.lower():
                            pair_type = 'гз'
                        elif 'кр' in val_str.lower():
                            pair_type = 'кр'
                        elif 'экз' in val_str.lower():
                            pair_type = 'экз'
                        elif 'з/о' in val_str.lower():
                            pair_type = 'з/о'
                        elif 'вси' in val_str.lower():
                            pair_type = 'вси'
                        
                        # Название предмета — следующая строка
                        subject = ""
                        try:
                            subj_val = sheet.cell(row + 1, col).value
                            if subj_val:
                                subject = str(subj_val).strip()
                        except:
                            pass
                        
                        # Аудитория — через строку
                        room = ""
                        try:
                            room_val = sheet.cell(row + 2, col).value
                            if room_val:
                                room = str(room_val).strip()
                        except:
                            pass
                        
                        # Если предмет пустой, используем raw значение
                        if not subject:
                            subject = val_str
                        
                        pairs.append({
                            'raw': val_str,
                            'subject': subject,
                            'room': room,
                            'type': pair_type,
                            'pair_num': pair_idx + 1,
                            'day': days[day_idx],
                            'day_offset': day_idx
                        })
                        
                except Exception as e:
                    logger.debug(f"Ошибка в колонке {col}: {e}")
        
        return pairs
    
    def get_schedule_for_group(self, group, target_date=None):
        """Получает расписание для группы на дату"""
        if not self.excel_loaded:
            return None
        
        if target_date is None:
            target_date = datetime.now()
        
        group = str(group).strip()
        
        # Прямое совпадение
        if group in self.schedule_data:
            data = self.schedule_data[group]
        else:
            # Нечёткий поиск
            for g in self.schedule_data:
                if group.lower() in g.lower() or g.lower() in group.lower():
                    data = self.schedule_data[g]
                    group = g
                    break
            else:
                return None
        
        # Ищем пары на конкретную дату
        result = []
        for date_key, pairs in data.items():
            if isinstance(date_key, datetime) and date_key.date() == target_date.date():
                result.extend(pairs)
        
        return result
    
    def find_pair_by_name(self, group, name):
        """Ищет пару по названию дисциплины"""
        if not self.excel_loaded or group not in self.schedule_data:
            return None
        
        name_lower = name.lower()
        data = self.schedule_data[group]
        
        for date_obj, pairs in data.items():
            for pair in pairs:
                if name_lower in pair.get('subject', '').lower():
                    return {'date': date_obj, 'pair': pair}
        return None
    
    def get_upcoming_exams(self, group, days_ahead=14):
        """Находит ближайшие зачёты/экзамены"""
        if not self.excel_loaded or group not in self.schedule_data:
            return []
        
        today = datetime.now()
        exams = []
        
        for date_obj, pairs in self.schedule_data[group].items():
            if not isinstance(date_obj, datetime):
                continue
            
            days_diff = (date_obj.date() - today.date()).days
            if 0 <= days_diff <= days_ahead:
                for pair in pairs:
                    pt = pair.get('type', '')
                    if pt in ['экз', 'з/о', 'кр', 'зач']:
                        exams.append({
                            'date': date_obj,
                            'subject': pair.get('subject', ''),
                            'type': pt,
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
    groups_list = ", ".join(sorted(bot.groups.keys())[:8]) + "..." if len(bot.groups) > 8 else ", ".join(sorted(bot.groups.keys()))
    
    welcome_text = f"""
🎓 *Бот расписания ВУНЦ ВВС*

👤 Группа: *{current_group}*
📊 Статус: {status}
📋 Группы: `{groups_list or "не найдены"}`

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
        
        # Render разрешает писать в /tmp/ и в текущую директорию
        # Используем /tmp/ для временного файла
        temp_dir = "/tmp"
        safe_name = "".join(c for c in document.file_name if c.isalnum() or c in '._-')
        temp_path = os.path.join(temp_dir, f"temp_{safe_name}")
        
        logger.info(f"Скачивание в: {temp_path}")
        await file.download_to_drive(temp_path)
        logger.info(f"Файл скачан: {os.path.exists(temp_path)}, размер: {os.path.getsize(temp_path) if os.path.exists(temp_path) else 0}")
        
        if bot.save_uploaded_excel(temp_path):
            groups_count = len(bot.groups)
            groups_preview = ", ".join(sorted(bot.groups.keys())[:6]) + "..." if len(bot.groups) > 6 else ", ".join(sorted(bot.groups.keys()))
            
            await update.message.reply_text(
                "✅ *Расписание успешно загружено!*\n\n"
                f"📁 Файл: `{document.file_name}`\n"
                f"👥 Найдено групп: *{groups_count}*\n"
                f"📋 Группы: `{groups_preview or 'не распознаны'}`\n\n"
                "Теперь можно смотреть расписание!",
                parse_mode='Markdown',
                reply_markup=get_main_keyboard()
            )
        else:
            await update.message.reply_text("❌ Ошибка при обработке файла. Проверьте логи на Render.")
            
    except Exception as e:
        logger.error(f"Ошибка загрузки файла: {e}")
        import traceback
        logger.error(traceback.format_exc())
        await update.message.reply_text(f"❌ Ошибка: {str(e)[:200]}")

async def show_schedule(update: Update, context: ContextTypes.DEFAULT_TYPE, days_offset: int = 0):
    """Показывает расписание"""
    if not bot.excel_loaded:
        await update.message.reply_text(
            "⚠️ *Расписание не загружено!*\n\nОтправь Excel файл боту через кнопку 📁 Загрузить расписание",
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
            emoji = {'л': '📖', 'пз': '🔧', 'с': '🗣️', 'гз': '🏫', 
                    'кр': '📝', 'экз': '📋', 'з/о': '✅', 'вси': '🔬'}.get(item.get('type', 'л'), '📚')
            pair_num = item.get('pair_num', '?')
            subject = item.get('subject', 'Не указано')
            room = item.get('room', '')
            room_text = f" 📍{room}" if room else ""
            
            text += f"{emoji} *Пара {pair_num}* — {subject}{room_text}\n"
    
    await update.message.reply_text(text, parse_mode='Markdown')

async def search_by_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Поиск по дате"""
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ Сначала загрузите расписание!")
        return
    
    await update.message.reply_text(
        "📅 *Введите дату для поиска*\n\nФормат: `ДД.ММ.ГГГГ` (например: `20.05.2025`)",
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
            f"❌ На *{target_date.strftime('%d.%м.%Y')}* пар не найдено для группы {group}",
            parse_mode='Markdown'
        )
    else:
        text = f"📅 *Пары на {target_date.strftime('%d.%m.%Y')}:*\n\n"
        for item in schedule:
            emoji = {'л': '📖', 'пз': '🔧', 'с': '🗣️', 'гз': '🏫', 
                    'кр': '📝', 'экз': '📋', 'з/о': '✅'}.get(item.get('type', 'л'), '📚')
            text += f"{emoji} *Пара {item.get('pair_num', '?')}* — {item.get('subject', 'Не указано')}\n"
        
        await update.message.reply_text(text, parse_mode='Markdown')
    
    return ConversationHandler.END

async def search_by_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Поиск по названию дисциплины"""
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ Сначала загрузите расписание!")
        return
    
    await update.message.reply_text(
        "🔎 *Введите название дисциплины для поиска:*\n(например: ИАО, ТВВС, СУВ, ОПС)",
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
        room = pair.get('room', '')
        room_text = f" 📍{room}" if room else ""
        
        await update.message.reply_text(
            f"🔎 *Найдено:*\n\n"
            f"📅 {date_str}\n"
            f"⏰ Пара {pair.get('pair_num', '?')}\n"
            f"📖 {pair.get('subject', 'Не указано')}{room_text}\n"
            f"📝 Тип: {pair.get('type', 'л')}",
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
    
    exams = bot.get_upcoming_exams(group, days_ahead=21)
    
    if not exams:
        await update.message.reply_text(
            "✅ *Ближайшие контрольные работы и зачёты не найдены*",
            parse_mode='Markdown'
        )
        return
    
    text = "📊 *Ближайшие контрольные и зачёты:*\n\n"
    for exam in exams[:12]:
        days_text = f"через {exam['days_until']} дн." if exam['days_until'] > 0 else "сегодня!"
        warning = " ⚠️" if exam['days_until'] <= 3 else ""
        
        text += f"📅 {exam['date'].strftime('%d.%m.%Y')} ({days_text}){warning}\n"
        text += f"📝 {exam['subject']} ({exam['type']})\n\n"
    
    await update.message.reply_text(text, parse_mode='Markdown')

async def settings_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Меню настроек"""
    user_id = update.effective_user.id
    settings = bot.user_settings.get(str(user_id), {})
    current_group = settings.get('group', '20-21')
    
    available_groups = sorted(bot.groups.keys()) if bot.groups else ['20-21', '20-22', '26-21', '7-21', '8-21']
    
    keyboard = []
    for i in range(0, len(available_groups), 2):
        row = [InlineKeyboardButton(f"👥 {g}", callback_data=f"group_{g}") for g in available_groups[i:i+2]]
        keyboard.append(row)
    keyboard.append([InlineKeyboardButton("🔙 Назад", callback_data="back_to_main")])
    
    await update.message.reply_text(
        f"⚙️ *Настройки:*\n\n👥 Текущая группа: *{current_group}*\n\nВыберите группу:",
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
                        emoji = {'л': '📖', 'пз': '🔧', 'с': '🗣️', 'гз': '🏫', 
                                'кр': '📝', 'экз': '📋', 'з/о': '✅'}.get(item.get('type', 'л'), '📚')
                        text += f"{emoji} *Пара {item.get('pair_num', '?')}* — {item.get('subject', '')}\n"
                    
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
    
    logger.info(f"🚀 Запуск бота... Рабочая директория: {os.getcwd()}")
    
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
