import telebot
import pandas as pd
import os

BOT_API = telebot.TeleBot('8062391635:AAFtbSMTgRc8WwFxwDCYOFDyBy0UYSqf0n8')

@BOT_API.message_handler(content_types=['text'])
def get_text_messages(message):
    if message.text == "/start":
        BOT_API.send_message(message.from_user.id,
                             "Приветствую! Я умею считать % выполненных дз.\nЕсли я тебе пригожусь, то скинь мне Excel файл\n/help - помощь")
    elif message.text == "Привет":
        BOT_API.send_message(message.from_user.id,
                            "Приветствую! Напиши /help, чтобы узнать обо мне")

    elif message.text == "/help":
        BOT_API.send_message(message.from_user.id,
                             "Функционал:\n/start - запускает меня\n/help - команда для помощи тебе\nБольше у меня функционала нет, я скромный...")
    else:
        BOT_API.send_message(message.from_user.id, "Я тебя не понимаю. Напиши /help.")

def lm(chat_id, text):
    for i in range(0, len(text), 4096):
        BOT_API.send_message(chat_id, text[i:i + 4096])

@BOT_API.message_handler(content_types=['document'])
def get_document_messages(message):
    if message.document:
        if message.document.mime_type in ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                            'application/vnd.ms-excel']:
            try:
                file_info = BOT_API.get_file(message.document.file_id)
                downloaded_file = BOT_API.download_file(file_info.file_path)

                with open('Tablica.xlsx', 'wb') as new_file:
                    new_file.write(downloaded_file)

                tabl = pd.read_excel('Tablica.xlsx')

                if tabl.shape[1] >= 20:
                    row1 = tabl.iloc[:, 0].tolist()
                    row20 = tabl.iloc[:, 19].tolist()

                    neuchi = []

                    for fio, procent in zip(row1, row20):
                        if procent < 50:
                            neuchi.append(f"ФИО: {fio} | Выполнено: {procent}%")

                    if neuchi:
                        mes = "Неучи меньше 50%:\n"
                        mes += "\n".join(neuchi)
                        lm(message.from_user.id, mes)
                    else:
                        BOT_API.send_message(message.from_user.id, "Все учатся!")
                else:
                    BOT_API.send_message(message.from_user.id, "В файле недостаточно столбцов.")

            except Exception as fail:
                BOT_API.send_message(message.from_user.id, f"Ошибка с файлом: {str(fail)}")

BOT_API.polling(none_stop=True, interval=0)