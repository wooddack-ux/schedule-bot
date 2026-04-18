import logging
import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler, CallbackQueryHandler
from datetime import datetime, timedelta
import os
import re
import shutil

logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

SELECT_GROUP, SETTINGS, SEARCH_DATE, SEARCH_NAME = range(4)

MONTHS_RU = {
    'январь': 1, 'февраль': 2, 'март': 3, 'апрель': 4, 'май': 5, 'июнь': 6,
    'июль': 7, 'август': 8, 'сентябрь': 9, 'октябрь': 10, 'ноябрь': 11, 'декабрь': 12
}

KNOWN_GROUPS = ['20-21', '20-22', '20-23', '11-21', '26-21', '7-21', '8-21', '8и-21', '20и-21', '20и-22']

class ScheduleBot:
    def __init__(self):
        self.excel_file = None
        self.workbook = None
        self.schedule_data = {}
        self.groups = {}
        self.user_settings = {}
        self.excel_loaded = False
    
    def save_uploaded_excel(self, file_path):
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
            
            self._parse_sheet_new(sheet, sheet_groups)
        
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
    
    def _parse_sheet_new(self, sheet, sheet_groups):
        """Новый парсер, учитывающий реальную структуру файла"""
        
        # Колонки для дней недели и пар
        # Формат: (колонка, день, номер_пары)
        day_columns = [
            (2, 'Пн', 1), (3, 'Пн', 2), (4, 'Пн', 3),
            (5, 'Вт', 1), (6, 'Вт', 2), (7, 'Вт', 3),
            (8, 'Ср', 1), (9, 'Ср', 2), (10, 'Ср', 3),
            (11, 'Чт', 1), (12, 'Чт', 2), (13, 'Чт', 3),
            (14, 'Пт', 1), (15, 'Пт', 2), (16, 'Пт', 3),
            (17, 'Сб', 1), (18, 'Сб', 2), (19, 'Сб', 3),
        ]
        
        current_dates = {}  # колонка -> (день, месяц, год)
        
        for row in range(1, min(500, sheet.max_row + 1)):
            # Проверяем все возможные колонки с датами
            for col in range(1, min(20, sheet.max_column + 1)):
                cell = sheet.cell(row, col).value
                if not cell or not isinstance(cell, str):
                    continue
                
                cell_lower = cell.lower().strip()
                if cell_lower in MONTHS_RU:
                    # Нашли месяц! Ищем число в первой колонке этой строки
                    day_cell = sheet.cell(row, 1).value
                    if day_cell:
                        try:
                            day = int(day_cell)
                            month = MONTHS_RU[cell_lower]
                            year = 2026 if month <= 6 else 2025
                            current_dates[col] = (day, month, year)
                            logger.debug(f"Дата: колонка {col} -> {day:02d}.{month:02d}.{year}")
                        except:
                            pass
            
            # Проверяем первую колонку на наличие группы
            col_a = sheet.cell(row, 1).value
            if not col_a:
                continue
            
            col_a_str = str(col_a).strip()
            
            # Проверяем, есть ли здесь группа
            found_group = None
            for g in sheet_groups:
                if g in col_a_str:
                    found_group = g
                    break
            
            if found_group:
                logger.debug(f"Группа {found_group} в строке {row}")
                
                # Ищем данные ЗА 1-3 строки ДО группы
                for look_back in range(1, 4):
                    data_row = row - look_back
                    if data_row < 1:
                        continue
                    
                    # Проверяем каждую колонку с парами
                    for col, day_name, pair_num in day_columns:
                        if col > sheet.max_column:
                            continue
                        
                        # Получаем предмет
                        subject_cell = sheet.cell(data_row, col).value
                        if not subject_cell:
                            continue
                        
                        subject = str(subject_cell).strip()
                        if subject in ['', 'None', '-', 'СР', 'Выходной', 'Праздник', 'Наряд']:
                            continue
                        
                        # Проверяем, есть ли дата для этой колонки
                        # Ищем ближайшую дату СЛЕВА от текущей колонки
                        date_info = None
                        for date_col in sorted(current_dates.keys(), reverse=True):
                            if date_col <= col:
                                date_info = current_dates[date_col]
                                break
                        
                        if not date_info:
                            continue
                        
                        day, month, year = date_info
                        date_obj = datetime(year, month, day)
                        
                        # Определяем тип занятия (ищем строкой выше)
                        pair_type = 'л'
                        if data_row > 1:
                            type_cell = sheet.cell(data_row - 1, col).value
                            if type_cell:
                                type_str = str(type_cell).lower()
                                if 'пз' in type_str:
                                    pair_type = 'пз'
                                elif 'гз' in type_str:
                                    pair_type = 'гз'
                                elif 'с' in type_str and 'вси' not in type_str:
                                    pair_type = 'с'
                                elif 'кр' in type_str:
                                    pair_type = 'кр'
                                elif 'экз' in type_str:
                                    pair_type = 'экз'
                                elif 'з/о' in type_str or 'зач' in type_str:
                                    pair_type = 'з/о'
                                elif 'вси' in type_str:
                                    pair_type = 'вси'
                        
                        # Ищем аудиторию (строкой ниже)
                        room = ''
                        if data_row + 1 <= sheet.max_row:
                            room_cell = sheet.cell(data_row + 1, col).value
                            if room_cell:
                                room = str(room_cell).strip()
                        
                        # Фильтруем - предметы обычно короткие аббревиатуры
                        if len(subject) > 30 or subject[0].isdigit():
                            continue
                        
                        # Добавляем в данные
                        if date_obj not in self.schedule_data[found_group]:
                            self.schedule_data[found_group][date_obj] = []
                        
                        exists = False
                        for p in self.schedule_data[found_group][date_obj]:
                            if p.get('pair_num') == pair_num and p.get('day') == day_name:
                                exists = True
                                break
                        
                        if not exists:
                            pair_data = {
                                'subject': subject,
                                'room': room,
                                'type': pair_type,
                                'pair_num': pair_num,
                                'day': day_name
                            }
                            self.schedule_data[found_group][date_obj].append(pair_data)
                            logger.debug(f"  + {found_group}: {date_obj.strftime('%d.%m')} {day_name}-{pair_num} {subject}")
    
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
        
        for i in range(days):
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
    
    def get_upcoming_exams(self, group, days_ahead=60):
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
        exams = []
        
        for date_obj, pairs in self.schedule_data[target_group].items():
            if not isinstance(date_obj, datetime):
                continue
            
            days_diff = (date_obj.date() - today.date()).days
            if 0 <= days_diff <= days_ahead:
                for pair in pairs:
                    pt = pair.get('type', '')
                    subj = pair.get('subject', '').lower()
                    if pt in ['экз', 'з/о', 'кр'] or 'экзамен' in subj or 'зачёт' in subj:
                        exams.append({
                            'date': date_obj,
                            'subject': pair.get('subject', ''),
                            'type': pt,
                            'days_until': days_diff
                        })
        
        return sorted(exams, key=lambda x: (x['days_until'], x['date']))
    
    def get_all_subjects(self, group):
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
        
        subjects = set()
        for pairs in self.schedule_data[target_group].values():
            for pair in pairs:
                subject = pair.get('subject', '')
                if subject and len(subject) <= 20:
                    subjects.add(subject)
        
        return sorted(subjects)


bot = ScheduleBot()


def get_main_keyboard():
    keyboard = [
        [KeyboardButton("📅 Сегодня"), KeyboardButton("📆 На 2 дня")],
        [KeyboardButton("🔍 По дате"), KeyboardButton("🔎 По предмету")],
        [KeyboardButton("⚙️ Группа"), KeyboardButton("📊 Экзамены")],
        [KeyboardButton("📁 Загрузить Excel"), KeyboardButton("📋 Предметы")]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if str(user_id) not in bot.user_settings:
        bot.user_settings[str(user_id)] = {'group': '20-21'}
    
    settings = bot.user_settings.get(str(user_id), {})
    current_group = settings.get('group', '20-21')
    
    status = "✅ Загружено" if bot.excel_loaded else "⚠️ Ожидание файла"
    groups_count = len(bot.groups)
    
    welcome_text = f"""
🎓 *Бот расписания ВУНЦ ВВС*

👤 Группа: *{current_group}*
📊 Статус: {status}
📋 Групп: *{groups_count}*
"""
    await update.message.reply_text(welcome_text, parse_mode='Markdown', reply_markup=get_main_keyboard())


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    
    if not document.file_name.endswith(('.xlsx', '.xls')):
        await update.message.reply_text("❌ Нужен Excel файл")
        return
    
    msg = await update.message.reply_text("⏳ Обработка...")
    
    try:
        file = await context.bot.get_file(document.file_id)
        temp_path = f"/tmp/temp_{document.file_name}"
        await file.download_to_drive(temp_path)
        
        if bot.save_uploaded_excel(temp_path):
            groups_count = len(bot.groups)
            
            stats = []
            total_pairs = 0
            for g in sorted(bot.groups.keys()):
                dates = len(bot.schedule_data.get(g, {}))
                pairs = sum(len(p) for p in bot.schedule_data.get(g, {}).values())
                total_pairs += pairs
                if pairs > 0:
                    stats.append(f"{g}: {dates} дат, {pairs} пар")
            
            if total_pairs == 0:
                await msg.edit_text("⚠️ Файл загружен, но занятия не найдены!")
            else:
                stats_text = "\n".join(stats[:10])
                await msg.edit_text(
                    f"✅ *Загружено!*\n\n"
                    f"👥 Групп: {groups_count}\n"
                    f"📊 Всего пар: {total_pairs}\n\n"
                    f"`{stats_text}`",
                    parse_mode='Markdown'
                )
            
            await update.message.reply_text("✅ Готово!", reply_markup=get_main_keyboard())
        else:
            await msg.edit_text("❌ Ошибка обработки")
            
    except Exception as e:
        logger.error(f"Ошибка: {e}")
        await msg.edit_text(f"❌ Ошибка: {str(e)[:100]}")


async def show_today(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ Загрузите Excel файл!")
        return
    
    user_id = update.effective_user.id
    group = bot.user_settings.get(str(user_id), {}).get('group', '20-21')
    
    target_date = datetime.now()
    schedule = bot.get_schedule_for_group(group, target_date)
    
    days_ru = ['пн', 'вт', 'ср', 'чт', 'пт', 'сб', 'вс']
    day_name = days_ru[target_date.weekday()]
    
    header = f"📅 *{target_date.strftime('%d.%m.%Y')} ({day_name})*\n👥 *{group}*\n\n"
    
    if not schedule:
        await update.message.reply_text(header + "😴 Нет занятий", parse_mode='Markdown')
        return
    
    text = header
    for item in schedule:
        emoji = {'л': '📖', 'пз': '✏️', 'с': '🗣️', 'гз': '👥', 'кр': '📝', 'экз': '📋', 'з/о': '✅'}.get(item.get('type', 'л'), '📚')
        room = item.get('room', '')
        room_text = f" 📍{room}" if room and room != 'None' else ""
        text += f"{emoji} *{item.get('pair_num', '?')}* — {item.get('subject', '—')}{room_text}\n"
    
    await update.message.reply_text(text, parse_mode='Markdown')


async def show_two_days(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ Загрузите Excel файл!")
        return
    
    user_id = update.effective_user.id
    group = bot.user_settings.get(str(user_id), {}).get('group', '20-21')
    
    schedule_dict = bot.get_schedule_for_days(group, 2)
    
    if not schedule_dict:
        await update.message.reply_text(f"😴 Нет занятий на 2 дня\n👥 {group}", parse_mode='Markdown')
        return
    
    text = f"📆 *Расписание*\n👥 *{group}*\n"
    days_ru = ['пн', 'вт', 'ср', 'чт', 'пт', 'сб', 'вс']
    
    for date_obj, schedule in schedule_dict.items():
        date_str = date_obj.strftime('%d.%m.%Y')
        day_name = days_ru[date_obj.weekday()]
        text += f"\n📅 *{date_str} ({day_name})*\n"
        
        if not schedule:
            text += "  😴 Нет занятий\n"
        else:
            for item in schedule:
                emoji = {'л': '📖', 'пз': '✏️', 'с': '🗣️', 'гз': '👥'}.get(item.get('type', 'л'), '📚')
                text += f"  {emoji} *{item.get('pair_num', '?')}* — {item.get('subject', '—')}\n"
    
    await update.message.reply_text(text, parse_mode='Markdown')


async def show_all_subjects(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ Загрузите Excel файл!")
        return
    
    user_id = update.effective_user.id
    group = bot.user_settings.get(str(user_id), {}).get('group', '20-21')
    
    subjects = bot.get_all_subjects(group)
    
    if not subjects:
        await update.message.reply_text(f"❌ Предметы не найдены\n👥 {group}")
        return
    
    text = f"📋 *Предметы*\n👥 *{group}*\n\n"
    for s in subjects:
        text += f"• `{s}`\n"
    
    if len(text) > 4000:
        text = text[:4000] + "..."
    
    await update.message.reply_text(text, parse_mode='Markdown')


async def search_by_name_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not bot.excel_loaded:
        await update.message.reply_text("⚠️ Загрузите Excel файл!")
        return ConversationHandler.END
    
    await update.message.reply_text("🔎 Введите название предмета:")
    return SEARCH_NAME


async def search_by_name_handle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.message.text.strip()
    user_id = update.effective_user.id
    group = bot.user_settings.get(str(user_id), {}).get('group', '20-21')
    
    results = bot.find_pair_by_name(group, query)
    
    if not results:
        await update.message.reply_text(f"❌ *{query}* не найдено\n👥 {group}", parse_mode='Markdown')
    else:
        text = f"🔎 *{query}*\n👥 *{group}*\n\n"
        for r in results[:10]:
            date_str = r['date'].strftime('%d.%m.%Y')
            pair = r['pair']
            text += f"📅 {date_str} — *{pair.get('pair_num', '?')}* {pair.get('subject', '')}\n"
        
        if len(results) > 10:
            text += f"\nи ещё {len(results)-10}..."
        
        await update.message.reply_text(text, parse_mode='Markdown')
    
    return ConversationHandler.END


async def settings_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    current_group = bot.user_settings.get(str(user_id), {}).get('group', '20-21')
    
    available = sorted(bot.groups.keys()) if bot.groups else KNOWN_GROUPS
    
    keyboard = []
    for i in range(0, len(available), 3):
        row = []
        for g in available[i:i+3]:
            label = f"✅ {g}" if g == current_group else f"👥 {g}"
            row.append(InlineKeyboardButton(label, callback_data=f"group_{g}"))
        keyboard.append(row)
    
    keyboard.append([InlineKeyboardButton("🔙 Закрыть", callback_data="close")])
    
    await update.message.reply_text(
        f"⚙️ *Выбор группы*\nТекущая: *{current_group}*",
        parse_mode='Markdown',
        reply_markup=InlineKeyboardMarkup(keyboard)
    )


async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    
    if data == "close":
        await query.edit_message_text("✅ Закрыто")
        return
    
    if data.startswith("group_"):
        group_code = data[6:]
        user_id = update.effective_user.id
        if str(user_id) not in bot.user_settings:
            bot.user_settings[str(user_id)] = {}
        bot.user_settings[str(user_id)]['group'] = group_code
        await query.edit_message_text(f"✅ Группа: *{group_code}*", parse_mode='Markdown')


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    
    if text == "📅 Сегодня":
        await show_today(update, context)
    elif text == "📆 На 2 дня":
        await show_two_days(update, context)
    elif text == "🔎 По предмету":
        await search_by_name_start(update, context)
    elif text == "⚙️ Группа":
        await settings_menu(update, context)
    elif text == "📋 Предметы":
        await show_all_subjects(update, context)
    elif text == "📁 Загрузить Excel":
        await update.message.reply_text("Отправьте Excel файл")
    else:
        await update.message.reply_text("Используйте кнопки меню")


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Отменено")
    return ConversationHandler.END


def main():
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    if not token:
        logger.error("Нет токена!")
        return
    
    app = Application.builder().token(token).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.FileExtension("xlsx") | filters.Document.FileExtension("xls"), handle_document))
    
    name_conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex(r'^🔎 По предмету$'), search_by_name_start)],
        states={SEARCH_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, search_by_name_handle)]},
        fallbacks=[CommandHandler("cancel", cancel)]
    )
    app.add_handler(name_conv)
    
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    
    logger.info("Бот запущен!")
    app.run_polling()


if __name__ == "__main__":
    main()
