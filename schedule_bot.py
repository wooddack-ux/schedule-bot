import logging
import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler, CallbackQueryHandler
from datetime import datetime, timedelta
import os
import json

# Настройка логирования
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Состояния для ConversationHandler
SELECT_GROUP, SETTINGS, SEARCH_DATE, SEARCH_NAME, CUSTOM_SETTINGS = range(5)

class ScheduleBot:
    def __init__(self):
        self.excel_file = None
        self.workbook = None
        self.groups = {
            '20-21': '20-21',
            '20-22': '20-22', 
            '20-23': '20-23',
            '11-21': '11-21',
            '26-21': '26-21',
            '7-21': '7-21',
            '8-21': '8-21'
        }
        self.user_settings = {}
        self.excel_loaded = False
        self.load_excel_on_startup()
    
    def load_excel_on_startup(self):
        """Загружает Excel файл при старте бота"""
        try:
            files = os.listdir('.')
            excel_files = [f for f in files if f.endswith(('.xlsx', '.xls'))]
            
            if not excel_files:
                logger.warning("⚠️ Excel файл не найден! Бот запустится без расписания.")
                self.excel_loaded = False
                return False
                
            self.excel_file = excel_files[0]
            self.workbook = openpyxl.load_workbook(self.excel_file, data_only=True)
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
            # Удаляем старый файл если есть
            for f in os.listdir('.'):
                if f.endswith(('.xlsx', '.xls')):
                    os.remove(f)
            
            # Переименовываем загруженный файл
            new_name = "schedule.xlsx"
            os.rename(file_path, new_name)
            
            # Перезагружаем
            self.excel_file = new_name
            self.workbook = openpyxl.load_workbook(self.excel_file, data_only=True)
            self.excel_loaded = True
            logger.info(f"✅ Загружен новый файл: {new_name}")
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка сохранения Excel: {e}")
            return False
    
    def get_schedule_for_group(self, group: str, target_date: datetime = None):
        """Получает расписание для группы на указанную дату"""
        if not self.excel_loaded or not self.workbook:
            return None
            
        if target_date is None:
            target_date = datetime.now()
            
        # Ищем лист с названием группы или похожим
        sheet = None
        for sheet_name in self.workbook.sheetnames:
            if group in sheet_name:
                sheet = self.workbook[sheet_name]
                break
        
        if not sheet:
            return None
            
        schedule = []
        
        # Ищем строку с нужной датой
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
                
            cell_date = row[0]
            if isinstance(cell_date, datetime):
                if cell_date.date() == target_date.date():
                    # Нашли нужную дату, собираем информацию
                    for col_idx, cell in enumerate(row[1:], start=1):
                        if cell and str(cell).strip():
                            cell_str = str(cell)
                            pair_type = 'л'  # по умолчанию лекция
                            
                            # Определяем тип занятия
                            if any(x in cell_str.lower() for x in ['пз', 'практ']):
                                pair_type = 'пз'
                            elif any(x in cell_str.lower() for x in ['с ', 'сем']):
                                pair_type = 'с'
                            elif any(x in cell_str.lower() for x in ['гз', 'групп']):
                                pair_type = 'гз'
                            elif any(x in cell_str.lower() for x in ['кр', 'контр']):
                                pair_type = 'кр'
                            elif any(x in cell_str.lower() for x in ['экз', 'экзамен']):
                                pair_type = 'экз'
                            elif any(x in cell_str.lower() for x in ['з/о', 'зач']):
                                pair_type = 'з/о'
                            elif any(x in cell_str.lower() for x in ['вси', 'выпуск']):
                                pair_type = 'вси'
                            
                            schedule.append({
                                'time': f"Пара {col_idx}",
                                'subject': cell_str,
                                'room': '',
                                'type': pair_type
                            })
                    break
        
        return schedule
    
    def find_pair_by_name(self, group: str, name: str):
        """Ищет пару по названию дисциплины"""
        if not self.excel_loaded or not self.workbook:
            return None
            
        sheet = None
        for sheet_name in self.workbook.sheetnames:
            if group in sheet_name:
                sheet = self.workbook[sheet_name]
                break
        
        if not sheet:
            return None
            
        name_lower = name.lower()
        
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
                
            row_date = row[0]
            for col_idx, cell in enumerate(row[1:], start=1):
                if cell and name_lower in str(cell).lower():
                    return {
                        'date': row_date,
                        'time': f"Пара {col_idx}",
                        'subject': str(cell),
                        'room': '',
                        'type': 'л'
                    }
        return None
    
    def get_upcoming_exams(self, group: str, days_ahead: int = 5):
        """Находит ближайшие зачёты/экзамены"""
        if not self.excel_loaded or not self.workbook:
            return []
            
        sheet = None
        for sheet_name in self.workbook.sheetnames:
            if group in sheet_name:
                sheet = self.workbook[sheet_name]
                break
        
        if not sheet:
            return []
            
        today = datetime.now()
        exams = []
        
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
                
            cell_date = row[0]
            if isinstance(cell_date, datetime):
                days_diff = (cell_date.date() - today.date()).days
                if 0 <= days_diff <= days_ahead:
                    for cell in row[1:]:
                        if cell:
                            cell_str = str(cell).lower()
                            if any(x in cell_str for x in ['экз', 'зач', 'кр', 'з/о']):
                                exams.append({
                                    'date': cell_date,
                                    'subject': str(cell),
                                    'type': 'экз/зач'
                                })
                                break
        
        return sorted(exams, key=lambda x: x['date'])

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
    
    welcome_text = f"""
🎓 *Бот расписания ВУНЦ ВВС*

👤 Группа: *{current_group}*
📊 Статус: {status}
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
    
    await update.message.reply_text("⏳ Скачиваю файл...")
    
    try:
        file = await context.bot.get_file(document.file_id)
        temp_path = f"temp_{document.file_name}"
        await file.download_to_drive(temp_path)
        
        if bot.save_uploaded_excel(temp_path):
            await update.message.reply_text(
                "✅ *Расписание успешно загружено!*\n\n"
                f"📁 Файл: `{document.file_name}`\n"
                f"📊 Листов: {len(bot.workbook.sheetnames)}\n\n"
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
            time_str = item.get('time', '')
            subject = item.get('subject', 'Не указано')
            pair_type = item.get('type', 'л')
            
            type_emoji = {
                'л': '📖', 'пз': '🔧', 'с': '🗣️', 'гз': '🏫',
                'кр': '📝', 'экз': '📋', 'з/о': '✅', 'вси': '🔬'
            }.get(pair_type, '📖')
            
            text += f"{type_emoji} *{time_str}* — {subject}\n"
    
    await update.message.reply_text(text, parse_mode='Markdown')

async def search_by_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Поиск по дате"""
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ Сначала загрузите расписание!")
        return
    
    await update.message.reply_text(
        "📅 *Введите дату для поиска*\n\n"
        "Формат: `ДД.ММ.ГГГГ` (например: `20.05.2026`)",
        parse_mode='Markdown'
    )
    return SEARCH_DATE

async def handle_date_search(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка поиска по дате"""
    query = update.message.text.strip()
    
    try:
        if len(query) == 10 and query[2] == '.' and query[5] == '.':
            day, month, year = int(query[:2]), int(query[3:5]), int(query[6:])
            target_date = datetime(year, month, day)
        else:
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
            f"❌ На *{target_date.strftime('%d.%m.%Y')}* пар не найдено",
            parse_mode='Markdown'
        )
    else:
        text = f"📅 *Пары на {target_date.strftime('%d.%m.%Y')}:*\n\n"
        for item in schedule:
            time_str = item.get('time', '')
            subject = item.get('subject', 'Не указано')
            pair_type = item.get('type', 'л')
            
            type_emoji = {
                'л': '📖', 'пз': '🔧', 'с': '🗣️', 'гз': '🏫',
                'кр': '📝', 'экз': '📋', 'з/о': '✅', 'вси': '🔬'
            }.get(pair_type, '📖')
            
            text += f"{type_emoji} *{time_str}* — {subject}\n"
        
        await update.message.reply_text(text, parse_mode='Markdown')
    
    return ConversationHandler.END

async def search_by_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Поиск по названию дисциплины"""
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ Сначала загрузите расписание!")
        return
    
    await update.message.reply_text(
        "🔎 *Введите название дисциплины для поиска:*\n"
        "(например: ИАО, ТВВС, СУВ)",
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
        date_str = result['date'].strftime('%d.%m.%Y') if isinstance(result['date'], datetime) else str(result['date'])
        
        await update.message.reply_text(
            f"🔎 *Найдено:*\n\n"
            f"📅 {date_str}\n"
            f"⏰ {result['time']}\n"
            f"📖 {result['subject']}",
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
    
    exams = bot.get_upcoming_exams(group, days_ahead=max(warning_days))
    
    if not exams:
        await update.message.reply_text(
            "✅ *Ближайшие контрольные работы и зачёты не найдены*",
            parse_mode='Markdown'
        )
        return
    
    text = "📊 *Ближайшие контрольные и зачёты:*\n\n"
    today = datetime.now()
    
    for exam in exams:
        days_until = (exam['date'].date() - today.date()).days
        date_str = exam['date'].strftime('%d.%m.%Y')
        
        warning = f" ⚠️ *Через {days_until} дня!*" if days_until in warning_days else ""
        
        text += f"📅 {date_str}{warning}\n"
        text += f"📝 {exam['subject']}\n\n"
    
    await update.message.reply_text(text, parse_mode='Markdown')

async def settings_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Меню настроек"""
    user_id = update.effective_user.id
    settings = bot.user_settings.get(str(user_id), {})
    
    current_group = settings.get('group', '20-21')
    days_ahead = settings.get('days_ahead', 2)
    warning_days = settings.get('exam_warning_days', [3, 5])
    notify_time = settings.get('notify_time', '06:00')
    enabled = settings.get('enabled', True)
    
    keyboard = [
        [InlineKeyboardButton(f"👥 Группа: {current_group}", callback_data="set_group")],
        [InlineKeyboardButton(f"📆 На сколько дней: {days_ahead}", callback_data="set_days")],
        [InlineKeyboardButton(f"⏰ Время уведомлений: {notify_time}", callback_data="set_time")],
        [InlineKeyboardButton(f"⚠️ Предупреждение: {', '.join(map(str, warning_days))} дн", callback_data="set_warnings")],
        [InlineKeyboardButton(f"{'🔔 Вкл' if enabled else '🔕 Выкл'} авто-отправка", callback_data="toggle_auto")],
        [InlineKeyboardButton("🔙 Назад", callback_data="back_to_main")]
    ]
    
    await update.message.reply_text(
        "⚙️ *Настройки:*\n\n"
        f"👥 Группа: *{current_group}*\n"
        f"📆 Показывать на: *{days_ahead} дня*\n"
        f"⏰ Время отправки: *{notify_time}*\n"
        f"⚠️ Предупреждать за: *{', '.join(map(str, warning_days))} дн*\n"
        f"🔔 Авто-отправка: *{'Вкл' if enabled else 'Выкл'}*",
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
        await query.edit_message_text("Главное меню", reply_markup=get_main_keyboard())
        return
    
    if data == "set_group":
        keyboard = []
        for code, name in bot.groups.items():
            keyboard.append([InlineKeyboardButton(name, callback_data=f"group_{code}")])
        keyboard.append([InlineKeyboardButton("🔙 Назад", callback_data="back_settings")])
        
        await query.edit_message_text(
            "👥 *Выберите группу:*",
            parse_mode='Markdown',
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
        return
    
    if data.startswith("group_"):
        group_code = data[6:]
        bot.user_settings[str(user_id)]['group'] = group_code
        await query.edit_message_text(f"✅ Группа: *{bot.groups[group_code]}*", parse_mode='Markdown')
        return
    
    if data == "toggle_auto":
        current = bot.user_settings[str(user_id)].get('enabled', True)
        bot.user_settings[str(user_id)]['enabled'] = not current
        status = "включена" if not current else "выключена"
        await query.edit_message_text(f"🔔 Авто-отправка {status}!", parse_mode='Markdown')
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
                        time_str = item.get('time', '')
                        subject = item.get('subject', 'Не указано')
                        pair_type = item.get('type', 'л')
                        
                        type_emoji = {
                            'л': '📖', 'пз': '🔧', 'с': '🗣️', 'гз': '🏫',
                            'кр': '📝', 'экз': '📋', 'з/о': '✅', 'вси': '🔬'
                        }.get(pair_type, '📖')
                        
                        text += f"{type_emoji} *{time_str}* — {subject}\n"
                    
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
