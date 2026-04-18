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
            files = [f for f in os.listdir('.') if f.endswith(('.xlsx', '.xls')) and not f.startswith('temp_')]
            if not files:
                logger.warning("⚠️ Excel файл не найден!")
                return False
            
            # Берём самый новый файл
            self.excel_file = files[0]
            file_path = os.path.abspath(self.excel_file)
            logger.info(f"Найден файл: {file_path}")
            
            self.workbook = openpyxl.load_workbook(file_path, data_only=True)
            self.parse_schedule()
            self.excel_loaded = True
            logger.info(f"✅ Загружен файл: {self.excel_file}")
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
            abs_temp_path = os.path.abspath(file_path)
            logger.info(f"Обработка файла: {abs_temp_path}")
            
            # Удаляем старые файлы Excel
            for f in os.listdir('.'):
                if f.endswith(('.xlsx', '.xls')) and not f.startswith('temp_'):
                    try:
                        old_path = os.path.abspath(f)
                        os.remove(old_path)
                        logger.info(f"Удалён старый файл: {old_path}")
                    except Exception as e:
                        logger.warning(f"Не удалось удалить {f}: {e}")
            
            # Целевой путь
            new_name = "schedule.xlsx"
            abs_new_path = os.path.abspath(new_name)
            
            # Копируем файл
            if os.path.exists(abs_temp_path):
                shutil.copy2(abs_temp_path, abs_new_path)
                logger.info(f"Файл скопирован: {abs_new_path}")
            else:
                logger.error(f"Временный файл не найден: {abs_temp_path}")
                return False
            
            # Удаляем временный
            try:
                os.remove(abs_temp_path)
            except:
                pass
            
            # Перезагружаем
            self.excel_file = new_name
            self.workbook = openpyxl.load_workbook(abs_new_path, data_only=True)
            self.parse_schedule()
            self.excel_loaded = True
            
            logger.info(f"✅ Загружен новый файл: {new_name}, групп: {len(self.groups)}")
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
            logger.info(f"  - {g}: {len(self.schedule_data.get(g, {}))} дат")
    
    def _parse_sheet_vunc(self, sheet, year):
        """Парсит лист ВУНЦ"""
        max_row = sheet.max_row
        
        row = 1
        current_month = None
        
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
            
            while row <= min(start_row + 20, sheet.max_row):
                val = sheet.cell(row, 1).value
                
                if val and isinstance(val, str):
                    # Ищем группы
                    if any(c in val for c in ['-', 'спец', 'гр']):
                        groups = self._extract_groups(val)
                        
                        for group in groups:
                            if group not in self.schedule_data:
                                self.schedule_data[group] = {}
                                self.groups[group] = group
                            
                            pairs = self._extract_pairs_for_group(sheet, row, date_obj)
                            if pairs:
                                if date_obj not in self.schedule_data[group]:
                                    self.schedule_data[group][date_obj] = []
                                self.schedule_data[group][date_obj].extend(pairs)
                        
                        row += 2  # Пропускаем строки группы
                    else:
                        row += 1
                else:
                    row += 1
                    
        except ValueError:
            pass
    
    def _extract_groups(self, block_text):
        """Извлекает группы из текста"""
        groups = []
        text = str(block_text).lower()
        
        # Паттерны для групп
        patterns = [
            r'(\d{2}-?\d{2})',      # 20-21, 20-22
            r'(\d{2}и-?\d{2})',     # 20и-21
            r'(\d{1}и?-?\d{2})',    # 7-21, 8-21, 8и-21
            r'(\d{2}-?\d{2}и?)',    # 20-21и
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, text)
            for match in matches:
                clean = match.replace(' ', '').strip()
                if clean and len(clean) >= 3 and clean not in groups:
                    groups.append(clean)
        
        return groups
    
    def _extract_pairs_for_group(self, sheet, row, date_obj):
        """Извлекает пары для группы"""
        pairs = []
        
        # Колонки с парами (Пн-Сб, пары 1-2, 3-4, 5-6)
        pair_cols = list(range(2, 38))  # Колонки B-AK
        
        for i, col in enumerate(pair_cols):
            try:
                val = sheet.cell(row, col).value
                if val and str(val).strip() not in ['', 'None', 'н/д']:
                    
                    # Определяем номер пары и день
                    pair_num = (i % 3) + 1
                    day_num = i // 3  # 0=Пн, 1=Вт, ..., 5=Сб
                    
                    # Проверяем, не пустая ли ячейка с номером пары
                    val_str = str(val).strip()
                    
                    # Парсим тип занятия
                    pair_type = 'л'
                    if 'пз' in val_str.lower():
                        pair_type = 'пз'
                    elif 'с' in val_str.lower():
                        pair_type = 'с'
                    elif 'гз' in val_str.lower():
                        pair_type = 'гз'
                    elif 'кр' in val_str.lower():
                        pair_type = 'кр'
                    elif 'экз' in val_str.lower():
                        pair_type = 'экз'
                    elif 'з/о' in val_str.lower():
                        pair_type = 'з/о'
                    
                    # Следующая строка — название предмета
                    subject = ""
                    try:
                        subject_val = sheet.cell(row + 1, col).value
                        if subject_val:
                            subject = str(subject_val).strip()
                    except:
                        pass
                    
                    # Строка с аудиторией
                    room = ""
                    try:
                        room_val = sheet.cell(row + 2, col).value if row + 2 <= sheet.max_row else None
                        if room_val:
                            room = str(room_val).strip()
                    except:
                        pass
                    
                    if subject or val_str not in ['л', 'пз', 'с', 'гз']:
                        pairs.append({
                            'raw': val_str,
                            'subject': subject if subject else val_str,
                            'room': room,
                            'type': pair_type,
                            'pair_num': pair_num,
                            'day_offset': day_num
                        })
            except Exception as e:
                logger.debug(f"Ошибка в колонке {col}: {e}")
        
        return pairs
    
    def get_schedule_for_group(self, group, target_date=None):
        """Получает расписание для группы"""
        if not self.excel_loaded:
            return None
        
        if target_date is None:
            target_date = datetime.now()
        
        group = str(group).strip()
        
        # Ищем группу
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
        
        # Ищем дату
        result = []
        for date_key, pairs in data.items():
            if isinstance(date_key, datetime) and date_key.date() == target_date.date():
                result.extend(pairs)
        
        return result
    
    def find_pair_by_name(self, group, name):
        """Ищет пару по названию"""
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
    document = update.message.document
    
    if not document.file_name.endswith(('.xlsx', '.xls')):
        await update.message.reply_text("❌ Нужен файл Excel (.xlsx или .xls)")
        return
    
    await update.message.reply_text("⏳ Скачиваю и обрабатываю файл...")
    
    try:
        file = await context.bot.get_file(document.file_id)
        
        # Очищаем имя файла от пробелов и спецсимволов
        safe_name = "".join(c for c in document.file_name if c.isalnum() or c in '._-')
        temp_name = f"temp_{safe_name}"
        temp_path = os.path.abspath(temp_name)
        
        logger.info(f"Скачивание в: {temp_path}")
        await file.download_to_drive(temp_path)
        logger.info(f"Файл скачан, размер: {os.path.getsize(temp_path)} байт")
        
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
    if not bot.excel_loaded:
        await update.message.reply_text(
            "⚠️ *Расписание не загружено!*\n\nОтправь Excel файл боту",
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
                    'кр': '📝', 'экз': '📋', 'з/о': '✅'}.get(item.get('type', 'л'), '📚')
            pair_num = item.get('pair_num', '?')
            subject = item.get('subject', 'Не указано')
            room = item.get('room', '')
            room_text = f" ({room})" if room else ""
            
            text += f"{emoji} *Пара {pair_num}* — {subject}{room_text}\n"
    
    await update.message.reply_text(text, parse_mode='Markdown')

async def search_by_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ Сначала загрузите расписание!")
        return
    
    await update.message.reply_text(
        "📅 *Введите дату для поиска*\n\nФормат: `ДД.ММ.ГГГГ`",
        parse_mode='Markdown'
    )
    return SEARCH_DATE

async def handle_date_search(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.message.text.strip()
    
    try:
        target_date = datetime.strptime(query, '%d.%m.%Y')
    except ValueError:
        await update.message.reply_text("❌ *Неверный формат!*\nИспользуйте: `ДД.ММ.ГГГГ`", parse_mode='Markdown')
        return ConversationHandler.END
    
    user_id = update.effective_user.id
    settings = bot.user_settings.get(str(user_id), {})
    group = settings.get('group', '20-21')
    
    schedule = bot.get_schedule_for_group(group, target_date)
    
    if not schedule:
        await update.message.reply_text(f"❌ На *{target_date.strftime('%d.%m.%Y')}* пар не найдено", parse_mode='Markdown')
    else:
        text = f"📅 *Пары на {target_date.strftime('%d.%m.%Y')}:*\n\n"
        for item in schedule:
            emoji = {'л': '📖', 'пз': '🔧', 'с': '🗣️', 'гз': '🏫', 'кр': '📝', 'экз': '📋', 'з/о': '✅'}.get(item.get('type', 'л'), '📚')
            text += f"{emoji} *Пара {item.get('pair_num', '?')}* — {item.get('subject', 'Не указано')}\n"
        await update.message.reply_text(text, parse_mode='Markdown')
    
    return ConversationHandler.END

async def search_by_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ Сначала загрузите расписание!")
        return
    
    await update.message.reply_text("🔎 *Введите название дисциплины:*\n(например: ИАО, ТВВС, СУВ)", parse_mode='Markdown')
    return SEARCH_NAME

async def handle_name_search(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.message.text.strip()
    user_id = update.effective_user.id
    settings = bot.user_settings.get(str(user_id), {})
    group = settings.get('group', '20-21')
    
    result = bot.find_pair_by_name(group, query)
    
    if not result:
        await update.message.reply_text(f"❌ *{query}* не найдено в группе {group}", parse_mode='Markdown')
    else:
        date_str = result['date'].strftime('%d.%m.%Y')
        pair = result['pair']
        await update.message.reply_text(
            f"🔎 *Найдено:*\n\n📅 {date_str}\n⏰ Пара {pair.get('pair_num', '?')}\n📖 {pair.get('subject', '')}",
            parse_mode='Markdown'
        )
    
    return ConversationHandler.END

async def show_exams(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ Сначала загрузите расписание!")
        return
    
    user_id = update.effective_user.id
    settings = bot.user_settings.get(str(user_id), {})
    group = settings.get('group', '20-21')
    
    exams = bot.get_upcoming_exams(group, days_ahead=21)
    
    if not exams:
        await update.message.reply_text("✅ *Ближайшие зачёты/экзамены не найдены*", parse_mode='Markdown')
        return
    
    text = "📊 *Ближайшие контрольные и зачёты:*\n\n"
    for exam in exams[:10]:
        days_text = f"через {exam['days_until']} дн." if exam['days_until'] > 0 else "сегодня!"
        text += f"📅 {exam['date'].strftime('%d.%m.%Y')} ({days_text})\n📝 {exam['subject']} ({exam['type']})\n\n"
    
    await update.message.reply_text(text, parse_mode='Markdown')

async def settings_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    settings = bot.user_settings.get(str(user_id), {})
    current_group = settings.get('group', '20-21')
    
    available_groups = sorted(bot.groups.keys()) if bot.groups else ['20-21', '20-22', '26-21']
    
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
        await query.edit_message_text(f"✅ Группа: *{group_code}*", parse_mode='Markdown')
        return

async def morning_job(context: ContextTypes.DEFAULT_TYPE):
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
                        emoji = {'л': '📖', 'пз': '🔧', 'с': '🗣️', 'гз': '🏫', 'кр': '📝', 'экз': '📋', 'з/о': '✅'}.get(item.get('type', 'л'), '📚')
                        text += f"{emoji} *Пара {item.get('pair_num', '?')}* — {item.get('subject', '')}\n"
                    await context.bot.send_message(chat_id=chat_id, text=text, parse_mode='Markdown')
            except Exception as e:
                logger.error(f"Ошибка отправки: {e}")

def main():
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    if not token:
        logger.error("❌ Не найден TELEGRAM_BOT_TOKEN!")
        print("❌ Установите переменную окружения TELEGRAM_BOT_TOKEN")
        return
    
    logger.info("🚀 Запуск бота...")
    
    application = Application.builder().token(token).build()
    
    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.Document.FileExtension("xlsx") | filters.Document.FileExtension("xls"), handle_document))
    application.add_handler(MessageHandler(filters.Regex(r'^📅'), lambda u, c: show_schedule(u, c, 0)))
    application.add_handler(MessageHandler(filters.Regex(r'^📆'), lambda u, c: show_schedule(u, c, 2)))
    application.add_handler(MessageHandler(filters.Regex(r'^📊'), show_exams))
    application.add_handler(MessageHandler(filters.Regex(r'^⚙️'), settings_menu))
    application.add_handler(MessageHandler(filters.Regex(r'^📁'), lambda u, c: u.message.reply_text("Отправьте Excel файл как документ")))
    application.add_handler(CallbackQueryHandler(button_handler))
    
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
    
    job_queue = application.job_queue
    if job_queue:
        job_queue.run_daily(morning_job, time=datetime.time(hour=6, minute=0))
        logger.info("✅ Утренняя отправка настроена")
    
    logger.info("🚀 Бот запущен!")
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
