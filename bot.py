from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler
import asyncio
import nest_asyncio
import pandas as pd
from openpyxl import load_workbook

# Токен вашего бота от BotFather
TOKEN = '7939209383:AAGLF2H2ZqG8S_aTHOKbNOKw3NDUNbV2DU8'

# Этапы диалога
CHOOSING, ENTER_NAME, ENTER_QUANTITY, ENTER_UNIT, ENTER_PROJECT = range(5)

# Клавиатура для блока "Заявки"
REQUESTS_KEYBOARD = [["Сделать заявку", "Принять материал"], ["Назад"]]

# Данные для заполнения заявки
user_data = {}

# Команда /start
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [["Табель", "Отчет"], ["Перемещение", "Заявки"]]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text(
        "Привет! Я бот. Выберите действие:",
        reply_markup=reply_markup
    )

# Блок "Заявки"
async def requests(update: Update, context: ContextTypes.DEFAULT_TYPE):
    reply_markup = ReplyKeyboardMarkup(REQUESTS_KEYBOARD, resize_keyboard=True)
    await update.message.reply_text("Выберите действие:", reply_markup=reply_markup)
    return CHOOSING

# Заполнение заявки: шаг 1 (Наименование товара)
async def make_request(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Введите наименование товара:")
    return ENTER_NAME

# Шаг 2: Ввод количества
async def enter_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_data["Наименование"] = update.message.text
    await update.message.reply_text("Введите количество:")
    return ENTER_QUANTITY

# Шаг 3: Ввод единицы измерения
async def enter_quantity(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        user_data["Количество"] = int(update.message.text)
        await update.message.reply_text("Введите единицу измерения (например, шт., кг, л):")
        return ENTER_UNIT
    except ValueError:
        await update.message.reply_text("Пожалуйста, введите число для количества.")
        return ENTER_QUANTITY

# Шаг 4: Ввод листа проекта
async def enter_unit(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_data["Единица измерения"] = update.message.text
    await update.message.reply_text("Введите лист проекта:")
    return ENTER_PROJECT

# Шаг 5: Формирование Excel-файла
async def enter_project(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_data["Лист проекта"] = update.message.text

    # Создаем Excel-файл
    df = pd.DataFrame([user_data])
    file_path = "заявки.xlsx"
    df.to_excel(file_path, index=False)

    # Настраиваем ширину столбцов
    wb = load_workbook(file_path)
    ws = wb.active
    for column in ws.columns:
        max_length = max(len(str(cell.value)) for cell in column if cell.value)
        ws.column_dimensions[column[0].column_letter].width = max_length + 5
    wb.save(file_path)

    await update.message.reply_text(
        "Заявка сформирована! Вы можете скачать файл.",
        reply_markup=ReplyKeyboardMarkup(REQUESTS_KEYBOARD, resize_keyboard=True)
    )
    await update.message.reply_document(open(file_path, "rb"))
    return CHOOSING

# Обработчик "Назад"
async def go_back(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await start(update, context)
    return ConversationHandler.END

# Основная функция
async def main():
    # Создание приложения
    application = Application.builder().token(TOKEN).build()

    # Обработчик команды /start
    application.add_handler(CommandHandler("start", start))

    # Диалог для блока "Заявки"
    conv_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex("^Заявки$"), requests)],
        states={
            CHOOSING: [
                MessageHandler(filters.Regex("^Сделать заявку$"), make_request),
                MessageHandler(filters.Regex("^Принять материал$"), make_request),
                MessageHandler(filters.Regex("^Назад$"), go_back)
            ],
            ENTER_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, enter_name)],
            ENTER_QUANTITY: [MessageHandler(filters.TEXT & ~filters.COMMAND, enter_quantity)],
            ENTER_UNIT: [MessageHandler(filters.TEXT & ~filters.COMMAND, enter_unit)],
            ENTER_PROJECT: [MessageHandler(filters.TEXT & ~filters.COMMAND, enter_project)],
        },
        fallbacks=[MessageHandler(filters.Regex("^Назад$"), go_back)]
    )
    application.add_handler(conv_handler)

    # Запуск бота
    await application.run_polling()

if __name__ == '__main__':
    nest_asyncio.apply()
    asyncio.run(main())
