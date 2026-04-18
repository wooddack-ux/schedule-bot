import logging
import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler, CallbackQueryHandler
from datetime import datetime, timedelta
import os
import re
import shutil

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

# Известные группы
KNOWN_GROUPS = ['20-21', '20-22', '20-23', '11-21', '26-21', '7-21', '8-21', '8и-21', '20и-21', '20и-22']

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
                logger.warning("⚠️ Excel файл не найден!")
                return False
            
            self.excel_file = files[0]
            file_path = os.path.join(work_dir, self.excel_file)
            logger.info(f"Найден файл: {file_path}")
            
            self.workbook = openpyxl.load_workbook(file_path, data_only=True)
            self.parse_schedule()
            self.excel_loaded = True
            logger.info(f"✅ Загружено: {len(self.groups)} групп")
            return True
            
        except Exception as e:
            logger.error(f"❌ Ошибка загрузки: {e}")
            import traceback
            logger.error(traceback.format_exc())
            self.excel_loaded = False
            return False
    
    def save_uploaded_excel(self, file_path):
        """Сохраняет загруженный Excel файл"""
        try:
            work_dir = os.getcwd()
            old_file = os.path.join(work_dir, "schedule.xlsx")
            
            if os.path.exists(old_file):
                os.remove(old_file)
            
            new_path = os.path.join(work_dir, "schedule.xlsx")
            shutil.copy2(file_path, new_path)
            
            try:
                os.remove(file_path)
            except:
                pass
            
            self.excel_file = "schedule.xlsx"
            self.workbook = openpyxl.load_workbook(new_path, data_only=True)
            self.parse_schedule()
            self.excel_loaded = True
            
            logger.info(f"✅ Обработано: {len(self.groups)} групп")
            return True
            
        except Exception as e:
            logger.error(f"❌ Ошибка сохранения: {e}")
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
            logger.info(f"Обработка листа: '{sheet_name}'")
            self._parse_sheet_vunc(sheet, year)
        
        logger.info(f"ИТОГО: {len(self.groups)} групп")
        for g in sorted(self.groups.keys()):
            count = len(self.schedule_data.get(g, {}))
            logger.info(f"  {g}: {count} дат")
    
        def _parse_sheet_vunc(self, sheet, year):
        """Парсит лист ВУНЦ с детальным логированием"""
        logger.info(f"=== ЛИСТ: '{sheet.title}' ===")
        logger.info(f"Размер: {sheet.max_row} строк × {sheet.max_column} колонок")
        
        # Показываем первые 10 строк первых 10 колонок
        logger.info("Первые 10 строк (колонки A-J):")
        for r in range(1, min(11, sheet.max_row + 1)):
            row_data = []
            for c in range(1, min(11, sheet.max_column + 1)):
                val = sheet.cell(r, c).value
                if val:
                    row_data.append(f"{c}:{str(val)[:15]}")
            if row_data:
                logger.info(f"  Row {r}: {' | '.join(row_data)}")
        
        # Ищем даты (месяца)
        month_rows = []
        for r in range(1, min(100, sheet.max_row + 1)):
            val = sheet.cell(r, 1).value
            if val and isinstance(val, str):
                val_lower = val.lower().strip()
                if val_lower in MONTHS_RU:
                    day_val = sheet.cell(r, 3).value
                    month_rows.append((r, val_lower, day_val))
        
        logger.info(f"Найдено дат (месяцев): {len(month_rows)}")
        for mr in month_rows[:5]:
            logger.info(f"  Строка {mr[0]}: {mr[1]} {mr[2]}")
        
        # Парсим каждую дату
        for row, month_str, day_val in month_rows:
            try:
                if day_val:
                    day = int(day_val)
                    month = MONTHS_RU[month_str]
                    self._process_day(sheet, row, month, year, day)
            except Exception as e:
                logger.error(f"Ошибка парсинга даты в строке {row}: {e}")
    
    def _process_day(self, sheet, start_row, month, year, day):
        """Обрабатывает один день"""
        try:
            date_obj = datetime(year, month, day)
            logger.info(f"--- Дата: {date_obj.strftime('%d.%m.%Y')} (строка {start_row}) ---")
            
            # Ищем группы в следующих 30 строках
            found_groups = []
            for r in range(start_row + 1, min(start_row + 31, sheet.max_row + 1)):
                val = sheet.cell(r, 1).value
                if not val:
                    continue
                
                val_str = str(val).strip()
                
                # Проверяем похоже ли на группу
                is_group = False
                if val_str in KNOWN_GROUPS:
                    is_group = True
                elif re.match(r'^\d{1,2}[и]?-\d{2}$', val_str):
                    is_group = True
                elif any(g in val_str for g in ['20-', '26-', '11-', '7-', '8-']):
                    is_group = True
                
                if is_group:
                    found_groups.append((r, val_str))
                    
                    # Обрабатываем группу
                    groups = self._extract_groups(val_str)
                    for group in groups:
                        if group not in self.schedule_data:
                            self.schedule_data[group] = {}
                            self.groups[group] = group
                        
                        pairs = self._extract_pairs_for_group(sheet, r, date_obj)
                        if pairs:
                            if date_obj not in self.schedule_data[group]:
                                self.schedule_data[group][date_obj] = []
                            self.schedule_data[group][date_obj].extend(pairs)
                            logger.info(f"  Группа {group}: +{len(pairs)} пар")
            
            if not found_groups:
                logger.warning(f"  Группы не найдены после строки {start_row}")
            else:
                logger.info(f"  Найдено групп: {len(found_groups)}")
                
        except Exception as e:
            logger.error(f"Ошибка в _process_day: {e}")
            import traceback
            logger.error(traceback.format_exc())
    
    def _extract_groups(self, text):
        """Извлекает группы из текста"""
        text = str(text).strip()
        groups = []
        
        # Способ 1: Прямое совпадение
        for kg in KNOWN_GROUPS:
            if kg in text and kg not in groups:
                groups.append(kg)
        
        # Способ 2: Регулярка
        matches = re.findall(r'(\d{1,2}(?:и)?-\d{2})', text)
        for m in matches:
            if m not in groups:
                groups.append(m)
        
        return groups
    
    def _extract_pairs_for_group(self, sheet, start_row, date_obj):
        """Извлекает пары для группы"""
        pairs = []
        
        row_type = start_row
        row_subject = start_row + 1
        row_room = start_row + 2
        
        if row_room > sheet.max_row:
            return pairs
        
        # Колонки пар: B(2), D(4), F(6)... до AJ(36)
        pair_cols = [
            (2, 'Пн', 1), (4, 'Пн', 2), (6, 'Пн', 3),
            (8, 'Вт', 1), (10, 'Вт', 2), (12, 'Вт', 3),
            (14, 'Ср', 1), (16, 'Ср', 2), (18, 'Ср', 3),
            (20, 'Чт', 1), (22, 'Чт', 2), (24, 'Чт', 3),
            (26, 'Пт', 1), (28, 'Пт', 2), (30, 'Пт', 3),
            (32, 'Сб', 1), (34, 'Сб', 2), (36, 'Сб', 3),
        ]
        
        for col, day_name, pair_num in pair_cols:
            try:
                type_val = sheet.cell(row_type, col).value
                if not type_val or str(type_val).strip() in ['', 'None', '-', 'н/д']:
                    continue
                
                type_str = str(type_val).strip()
                
                # Определяем тип
                pair_type = 'л'
                tl = type_str.lower()
                if 'пз' in tl:
                    pair_type = 'пз'
                elif 'с' in tl and len(tl) <= 2:
                    pair_type = 'с'
                elif 'гз' in tl:
                    pair_type = 'гз'
                elif 'кр' in tl:
                    pair_type = 'кр'
                elif 'экз' in tl:
                    pair_type = 'экз'
                elif 'з/о' in tl or 'зач' in tl:
                    pair_type = 'з/о'
                elif 'вси' in tl:
                    pair_type = 'вси'
                
                # Предмет
                subject_val = sheet.cell(row_subject, col).value
                subject = str(subject_val).strip() if subject_val else ""
                
                # Аудитория
                room_val = sheet.cell(row_room, col).value
                room = str(room_val).strip() if room_val else ""
                
                if subject and subject not in ['None', '-', '']:
                    pairs.append({
                        'subject': subject,
                        'room': room,
                        'type': pair_type,
                        'pair_num': pair_num,
                        'day': day_name
                    })
                    
            except Exception as e:
                pass
        
        return pairs
    
    def get_schedule_for_group(self, group, target_date=None):
        """Получает расписание для группы"""
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
                    group = g
                    break
            else:
                return None
        
        result = []
        for date_key, pairs in data.items():
            if isinstance(date_key, datetime) and date_key.date() == target_date.date():
                result.extend(pairs)
        
        # Сортируем
        day_order = {'Пн': 0, 'Вт': 1, 'Ср': 2, 'Чт': 3, 'Пт': 4, 'Сб': 5}
        result.sort(key=lambda x: (day_order.get(x.get('day', 'Пн'), 0), x.get('pair_num', 0)))
        
        return result
    
    def find_pair_by_name(self, group, name):
        """Ищет пару по названию"""
        if not self.excel_loaded or group not in self.schedule_data:
            return None
        
        name_lower = name.lower()
        for date_obj, pairs in self.schedule_data[group].items():
            for pair in pairs:
                if name_lower in pair.get('subject', '').lower():
                    return {'date': date_obj, 'pair': pair}
        return None
    
    def get_upcoming_exams(self, group, days_ahead=30):
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
    
    status = "✅ Загружено" if bot.excel_loaded else "⚠️ Ожидание файла"
    groups_list = ", ".join(sorted(bot.groups.keys())) if bot.groups else "не найдены"
    
    welcome_text = f"""
🎓 *Бот расписания ВУНЦ ВВС*

👤 Группа: *{current_group}*
📊 Статус: {status}
📋 Группы: `{groups_list}`

⚙️ Настройки: /settings

📋 *Команды:*
• 📅 Расписание на сегодня
• 📆 Расписание на 2 дня  
• 🔍 Поиск по дате
• 🔎 Поиск по названию
• 📊 Экзамены/Зачёты
• 📁 Загрузить расписание
"""
    
    await update.message.reply_text(welcome_text, parse_mode='Markdown', reply_markup=get_main_keyboard())

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    
    if not document.file_name.endswith(('.xlsx', '.xls')):
        await update.message.reply_text("❌ Нужен Excel (.xlsx или .xls)")
        return
    
    await update.message.reply_text("⏳ Скачиваю и обрабатываю...")
    
    try:
        file = await context.bot.get_file(document.file_id)
        
        temp_dir = "/tmp"
        safe_name = "".join(c for c in document.file_name if c.isalnum() or c in '._-')
        temp_path = os.path.join(temp_dir, f"temp_{safe_name}")
        
        await file.download_to_drive(temp_path)
        
        if bot.save_uploaded_excel(temp_path):
            groups_count = len(bot.groups)
            groups_preview = ", ".join(sorted(bot.groups.keys())[:8])
            
            await update.message.reply_text(
                f"✅ *Загружено!*\n\n"
                f"👥 Групп: *{groups_count}*\n"
                f"📋 `{groups_preview}`",
                parse_mode='Markdown',
                reply_markup=get_main_keyboard()
            )
        else:
            await update.message.reply_text("❌ Ошибка обработки")
            
    except Exception as e:
        logger.error(f"Ошибка: {e}")
        await update.message.reply_text(f"❌ Ошибка: {str(e)[:100]}")

async def show_schedule(update: Update, context: ContextTypes.DEFAULT_TYPE, days_offset: int = 0):
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ *Загрузите расписание!*", parse_mode='Markdown')
        return
    
    user_id = update.effective_user.id
    settings = bot.user_settings.get(str(user_id), {})
    group = settings.get('group', '20-21')
    
    target_date = datetime.now() + timedelta(days=days_offset)
    schedule = bot.get_schedule_for_group(group, target_date)
    
    days_text = "сегодня" if days_offset == 0 else f"+{days_offset}"
    header = f"📅 *{target_date.strftime('%d.%m.%Y')} ({days_text})*\n👥 *{group}*\n\n"
    
    if not schedule:
        await update.message.reply_text(header + "😴 *Нет пар!*", parse_mode='Markdown')
        return
    
    text = header
    current_day = None
    
    for item in schedule:
        day = item.get('day', '')
        if day != current_day:
            text += f"\n📌 *{day}*\n"
            current_day = day
        
        emoji = {'л': '📖', 'пз': '🔧', 'с': '🗣️', 'гз': '🏫', 
                'кр': '📝', 'экз': '📋', 'з/о': '✅', 'вси': '🔬'}.get(item.get('type', 'л'), '📚')
        pair_num = item.get('pair_num', '?')
        subject = item.get('subject', 'Нет')
        room = item.get('room', '')
        room_text = f" 📍{room}" if room else ""
        
        text += f"{emoji} *П{pair_num}* — {subject}{room_text}\n"
    
    await update.message.reply_text(text, parse_mode='Markdown')

async def search_by_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ Загрузите расписание!")
        return
    
    await update.message.reply_text("📅 *Дата:*\n`ДД.ММ.ГГГГ`", parse_mode='Markdown')
    return SEARCH_DATE

async def handle_date_search(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.message.text.strip()
    
    try:
        target_date = datetime.strptime(query, '%d.%m.%Y')
    except ValueError:
        await update.message.reply_text("❌ Формат: `ДД.ММ.ГГГГ`", parse_mode='Markdown')
        return ConversationHandler.END
    
    user_id = update.effective_user.id
    settings = bot.user_settings.get(str(user_id), {})
    group = settings.get('group', '20-21')
    
    schedule = bot.get_schedule_for_group(group, target_date)
    
    if not schedule:
        await update.message.reply_text(f"❌ Нет пар на {target_date.strftime('%d.%m.%Y')}", parse_mode='Markdown')
    else:
        text = f"📅 *{target_date.strftime('%d.%m.%Y')}:*\n\n"
        for item in schedule:
            emoji = {'л': '📖', 'пз': '🔧', 'с': '🗣️', 'гз': '🏫', 
                    'кр': '📝', 'экз': '📋', 'з/о': '✅'}.get(item.get('type', 'л'), '📚')
            text += f"{emoji} *П{item.get('pair_num', '?')}* — {item.get('subject', '')}\n"
        await update.message.reply_text(text, parse_mode='Markdown')
    
    return ConversationHandler.END

async def search_by_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ Загрузите расписание!")
        return
    
    await update.message.reply_text("🔎 *Название:*", parse_mode='Markdown')
    return SEARCH_NAME

async def handle_name_search(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.message.text.strip()
    user_id = update.effective_user.id
    settings = bot.user_settings.get(str(user_id), {})
    group = settings.get('group', '20-21')
    
    result = bot.find_pair_by_name(group, query)
    
    if not result:
        await update.message.reply_text(f"❌ *{query}* не найдено", parse_mode='Markdown')
    else:
        date_str = result['date'].strftime('%d.%m.%Y')
        pair = result['pair']
        await update.message.reply_text(
            f"🔎 *Найдено:*\n📅 {date_str}\n⏰ Пара {pair.get('pair_num', '?')}\n📖 {pair.get('subject', '')}",
            parse_mode='Markdown'
        )
    
    return ConversationHandler.END

async def show_exams(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ Загрузите расписание!")
        return
    
    user_id = update.effective_user.id
    settings = bot.user_settings.get(str(user_id), {})
    group = settings.get('group', '20-21')
    
    exams = bot.get_upcoming_exams(group, days_ahead=30)
    
    if not exams:
        await update.message.reply_text("✅ *Нет ближайших зачётов/экзаменов*", parse_mode='Markdown')
        return
    
    text = "📊 *Ближайшие:*\n\n"
    for exam in exams[:10]:
        days_text = f"{exam['days_until']}д" if exam['days_until'] > 0 else "сегодня!"
        text += f"📅 {exam['date'].strftime('%d.%m.%Y')} ({days_text})\n📝 {exam['subject']}\n\n"
    
    await update.message.reply_text(text, parse_mode='Markdown')

async def settings_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    settings = bot.user_settings.get(str(user_id), {})
    current_group = settings.get('group', '20-21')
    
    available_groups = sorted(bot.groups.keys()) if bot.groups else KNOWN_GROUPS
    
    keyboard = []
    for i in range(0, len(available_groups), 2):
        row = [InlineKeyboardButton(f"👥 {g}", callback_data=f"group_{g}") for g in available_groups[i:i+2]]
        keyboard.append(row)
    keyboard.append([InlineKeyboardButton("🔙 Назад", callback_data="back_to_main")])
    
    await update.message.reply_text(
        f"⚙️ *Настройки*\n\nТекущая: *{current_group}*",
        parse_mode='Markdown',
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    user_id = update.effective_user.id
    
    if data == "back_to_main":
        await query.edit_message_text("Меню")
        return
    
    if data.startswith("group_"):
        group_code = data[6:]
        bot.user_settings[str(user_id)]['group'] = group_code
        await query.edit_message_text(f"✅ *{group_code}*", parse_mode='Markdown')
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
                    text = f"📅 *Сегодня*\n\n"
                    for item in schedule:
                        emoji = {'л': '📖', 'пз': '🔧', 'с': '🗣️', 'гз': '🏫', 
                                'кр': '📝', 'экз': '📋', 'з/о': '✅'}.get(item.get('type', 'л'), '📚')
                        text += f"{emoji} *П{item.get('pair_num', '?')}* — {item.get('subject', '')}\n"
                    await context.bot.send_message(chat_id=chat_id, text=text, parse_mode='Markdown')
            except Exception as e:
                logger.error(f"Ошибка отправки: {e}")

def main():
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    if not token:
        logger.error("❌ Нет TELEGRAM_BOT_TOKEN!")
        return
    
    logger.info(f"🚀 Старт. Директория: {os.getcwd()}")
    
    application = Application.builder().token(token).build()
    
    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.Document.FileExtension("xlsx") | filters.Document.FileExtension("xls"), handle_document))
    application.add_handler(MessageHandler(filters.Regex(r'^📅'), lambda u, c: show_schedule(u, c, 0)))
    application.add_handler(MessageHandler(filters.Regex(r'^📆'), lambda u, c: show_schedule(u, c, 2)))
    application.add_handler(MessageHandler(filters.Regex(r'^📊'), show_exams))
    application.add_handler(MessageHandler(filters.Regex(r'^⚙️'), settings_menu))
    application.add_handler(MessageHandler(filters.Regex(r'^📁'), lambda u, c: u.message.reply_text("Отправьте Excel файл")))
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
    
    logger.info("🚀 Бот запущен!")
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
