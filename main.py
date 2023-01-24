import os
import telebot
from telebot.types import InlineKeyboardMarkup, InlineKeyboardButton
from dotenv import load_dotenv

from schedule import schedule_workbook, register_schedule_request, current_schedule

# load environment variable
load_dotenv()

# telegram bot api key
API_KEY = os.getenv('API_KEY')

# make constant
LIHAT_JADWAL = "lihat-jadwal"
JADWAL_TERSEDIA = "jadwal-tersedia"
JADWAL_TERISI = "jadwal-terisi"
DAFTAR = "daftar"

# make a bot instance
bot = telebot.TeleBot(API_KEY)


@bot.message_handler(commands=['greet'])
def greet(message):
    bot.send_message(message.chat.id, "Howdy, how are you doing?")


def start_markup():
    markup = InlineKeyboardMarkup()
    markup.row_width = 2

    # lihat jadwal
    lihatJadwal = InlineKeyboardButton("Lihat Jadwal", callback_data=LIHAT_JADWAL)

    # jadwal tersedia
    jadwalTersedia = InlineKeyboardButton("Jadwal Tersedia", callback_data=JADWAL_TERSEDIA)

    # jadwal terisi
    jadwalTerisi = InlineKeyboardButton("Jadwal Terisi", callback_data=JADWAL_TERISI)

    # daftar
    daftar = InlineKeyboardButton("Daftar", callback_data=DAFTAR)

    markup.add(lihatJadwal, jadwalTersedia, jadwalTerisi, daftar)

    return markup


@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, "Pilih apa yang bisa kami bantu", reply_markup=start_markup())

@bot.message_handler(func=register_schedule_request)
def register_schedule(message):
    request = message.text.split("_")
    name = request[1]
    day = request[2]
    session = int(request[3])
    current_schedule[day][session] = name

    if current_schedule[day][session] == name:
        schedule_workbook()
        formatMessage = "Berhasil menambahkan {} pada schedule hari {} sesi {} minggu ini".format(name, day, session)
        bot.send_message(message.chat.id, formatMessage)


@bot.callback_query_handler(func=lambda call:True)
def callback_query(call):
    if call.data == LIHAT_JADWAL:
        bot.send_message(call.message.chat.id, "Nih jadwalnya broooo")
        document = open('current-schedule.xlsx', 'rb')
        bot.send_document(call.message.chat.id, document)
    elif call.data == JADWAL_TERSEDIA:
        bot.answer_callback_query(call.id, "Nihh jadwal tersedia bang")
        bot.send_message(call.message.chat.id, "Nihh jadwal tersedia bang")
    elif call.data == JADWAL_TERISI:
        bot.answer_callback_query(call.id, "Nihhh jadwal yang udah keisi kang")
        bot.send_message(call.message.chat.id, "Nihhh jadwal yang udah keisi kang")
    elif call.data == DAFTAR:
        bot.send_message(call.message.chat.id, "Ketik daftar dengan format Daftar_Nama Anda_Hari_Sesi dengan contoh:\n"
                                               "Daftar_Hilman Taris Muttaqin_Senin_1")


bot.infinity_polling()
