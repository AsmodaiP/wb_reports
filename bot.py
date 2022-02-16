import logging
import os
import sys
from logging.handlers import RotatingFileHandler

from dotenv import load_dotenv
from telegram.ext import (CommandHandler, ConversationHandler, Filters,
                          MessageHandler, Updater)

import sheet
import pymorphy2

morph = pymorphy2.MorphAnalyzer()

load_dotenv()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
log_file = os.path.join(BASE_DIR, 'bot.log')
SPREADSHEET_ID = os.environ['SPREADSHEET_ID']

console_handler = logging.StreamHandler()
file_handler = RotatingFileHandler(
    log_file,
    maxBytes=100000,
    backupCount=3,
    encoding='utf-8'
)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s, %(levelname)s, %(message)s',
    handlers=(
        file_handler,
        console_handler
    )
)
try:
    TELEGRAM_TOKEN = os.environ['TELEGRAM_TOKEN']
except KeyError as exc:
    logging.error(exc, exc_info=True)
    sys.exit('Не удалось получить переменные окружения')


def start(bot, update):
    bot.message.reply_text('Начало работы')


def get_and_download_file(bot, update):
    file_id = bot.message.document['file_id']
    name = bot.message.document['file_name']
    file = update.bot.get_file(file_id)
    file.download(name)
    update.user_data['file_name'] = name
    bot.message.reply_text(
        'Введите название листа, на который нужно внести данные')
    return 'get_range'


def get_link(bot, update):
    link = bot.message.text.split('/')
    link.sort(key=len, reverse=True)
    update.user_data['spreadsheet_id'] = link[0]
    bot.message.reply_text(
        'Введите название листа, на который нужно внести данные')
    return 'get_range'


def cancel(bot, update):
    bot.message.reply_text('Операция отменена')
    return ConversationHandler.END


def get_range(bot, update):
    '''Получает таблицу из сообщения'''
    try:
        range_from_message = bot.message.text.strip()
        name = update.user_data['file_name']
        # spreadsheet_id = update.user_data['spreadsheet_id']
        result = sheet.update_table(name, SPREADSHEET_ID, range_from_message)
        errors_articles = result['errors']
        sum_of_refund = result['sum']

        try:
            bot.message.reply_text(f'Нет места для подстановки {len(errors_articles)} {morph.parse("артикулов")[0].make_agree_with_number(len(errors_articles)).word}')
        except:
            bot.message.reply_text(
                f'Нет места для подстановки {len(errors_articles)} артикулов')
        bot.message.reply_text("\n".join(errors_articles))
        bot.message.reply_text(f'Сумма возвратов {round(sum_of_refund,2)}')
    except Exception as ex:
        logging.error(ex, exc_info=ex)
        bot.message.reply_text(
            'Что-то пошло не так. \nУбедитесь, что аккаунт service@sheets-325814.iam.gserviceaccount.com имеет доступ к редактированию таблицы, перепроверьте введенные данные и начните сначала')
    return ConversationHandler.END


updater = Updater(token=TELEGRAM_TOKEN)

start_handler = CommandHandler('start', start)
updater.dispatcher.add_handler(start_handler)

document_handler = MessageHandler(
    Filters.document.file_extension("xlsx"),
    get_and_download_file)

link_handler = MessageHandler(Filters.regex(
    r'[(http(s)?):\/\/(www\.)?a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)'), get_link)


get_current_supplie_handler = ConversationHandler(
    entry_points=[document_handler],
    states={
        # 'get_link': [link_handler],
        'get_range': [MessageHandler(Filters.text & ~Filters.command, get_range)]
    },
    fallbacks=[CommandHandler('cancel', cancel)])
updater.dispatcher.add_handler(get_current_supplie_handler)

updater.start_polling()
