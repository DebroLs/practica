import telebot
from telebot import types
import pandas as pd

Token = "8062391635:AAFtbSMTgRc8WwFxwDCYOFDyBy0UYSqf0n8"

bot=telebot.TeleBot("8062391635:AAFtbSMTgRc8WwFxwDCYOFDyBy0UYSqf0n8")

@bot.message_handler(commands = ["start"])
def send_welcome(message):
    bot.reply_to(message,"\nПриветствую, я бот для поиска студентов с % домашнего задания ниже 50. Кинь мне файл и я скажу, кто провинился!\n")

@bot.message_handler(commands = ["help", "/help"])
def send_welcome(message):
    bot.reply_to(message,"\nКоманда:\n/start - Запускает меня в работу\n/help - Помогает тебе разобраться со мной\n/button - дает доступ к кнопкам управления\n")

@bot.message_handler(commands = ["button"])
def button_message(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    item1 = types.KeyboardButton("Получить ссылку на гитхаб бота")
    item2 = types.KeyboardButton("Загрузить файл Excel")
    markup.add(item1)
    markup.add(item2)
    bot.send_message(message.chat.id, "Что вы хотите сделать?", reply_markup=markup)

@bot.message_handler(content_types = "text")
def message_reply(message):
    if message.text=="Получить ссылку на гитхаб бота":
        bot.send_message(message.chat.id,"https://github.com/DebroLs/practica/blob/main/BotHW.py")
    elif message.text=="Загрузить файл Excel":
        bot.send_message(message.chat.id,"Прикрепляйте файл и я начну анализ")
    else:
        bot.send_message(message.chat.id,"Не понимаю вас. Вы можете воспользоваться командой /help")

@bot.message_handler(content_types = "document")
def document(message):
    if message.document:
        Test = message.document.file_name
        info = bot.get_file(message.document.file_id)
        Tabl = bot.download_file(info.file_path)
        with open(Test, "wb") as f:
            f.write(Tabl)
            df = pd.read_excel(Test)

        if "Percentage Homework" in df.columns:
            df.index = df.index + 2
            itog = df.loc[df["Percentage Homework"] < 50, ["FIO", "Percentage Homework"]]
            final = "Вот студенты, которые сдали менее 50% домашних работ:\n" + itog.to_string()
            bot.reply_to(message, final)
        else:
            bot.reply_to(message, "Ошибка, загружен неверный файл")



bot.infinity_polling()