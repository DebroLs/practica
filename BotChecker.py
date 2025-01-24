import telebot
import pandas as pd
import os
import re

BOT_API = telebot.TeleBot('8062391635:AAFtbSMTgRc8WwFxwDCYOFDyBy0UYSqf0n8')


patterns = [
    r'Урок\s№\d*[.]\sТема:',
    r'Урок\s№\d*[.]\sТема\s:',
    r'\sУрок\s№\d*[.]\sТема\s:',
    r'\sУрок\s№\d*[.]\sТема\s:\s',
    r'Урок\s№\d*[.]Тема\s:',
    r'Урок\s№\d*[.]Тема\s:\s',
    #########################################
    r'Урок\s\d*[.]\sТема:',
    r'Урок\s\d*[.]\sТема\s:',
    r'\sУрок\s\d*[.]\sТема\s:',
    r'\sУрок\s\d*[.]\sТема\s:\s',
    r'Урок\s\d*[.]Тема\s:',
    r'Урок\s\d*[.]Тема\s:\s',
    #########################################
    r'Урок\s\d*\sТема\s',
    r'Урок\s№\d*\sТема\s',
    r'\sУрок\s\d*\sТема\s',
    r'\sУрок\s№\d*\sТема\s'
    #########################################
    r'Урок\s№\d*3\sтема\s',
    r'\sрок\s№\d*3\sтема\s',
    r'\sУрок\s№\s\d*3\sтема\s',
    r'\sУрок\s№\s\d*3[.]\sтема\s',
    r'\sУрок\s№\s\d*3[.]\sТема:\s',
    r'Урок\s№\d*3\sтема\s',
    r'Урок\s№\d*3\sтема\s',
    r'Урок\s№\s\d*3\sтема\s',
    r'Урок\s№\s\d*3[.]\sтема\s',
    r'Урок\s№\s\d*3[.]\sТема:\s',
    #########################################
    r'\sУрок\s№\d*\.\sТема\s:\s',
    r'\sУрок\s№\d*\.\sТема\s:',
    r'\sУрок\s№\s\d*\.\sТема:\s',
    r'\sУрок\s№\s\d*\.\sТема:',
    r'\sУрок\s№\d*\.\sтема\s:\s',
    r'\sУрок\s№\d*\.\sтема\s:',
    r'\sУрок\s№\s\d*\.\sтема:\s',
    r'\sУрок\s№\s\d*\.\sтема:',
    r'\sУрок\s№\s\d*\.\sТема:\s',
    r'\sУрок\s№\s\d*\.\sТема:',
    r'\sУрок\s№\s\d*\.Тема:\s',
    r'\sУрок\s№\s\d*\.Тема:',
    r'\sУрок\s№\s\d*\.\sтема:\s',
    r'\sУрок\s№\s\d*\.\тТема:',
    r'\sУрок\s№\s\d*\.тема:\s',
    r'\sУрок\s№\s\d*\.тема:',
    r'Урок\s№\s\d*\.\sТема:\s',
]

@BOT_API.message_handler(content_types=['text'])
def get_text_messages(message):
    if message.text == "/start":
        BOT_API.send_message(message.from_user.id,
                             "Приветствую! Я умею считать количество проведенных пар по разным дисциплинам.\nЕсли я тебе пригожусь, то скинь мне Excel файл\n/help - помощь")

    elif message.text == "/help":
        BOT_API.send_message(message.from_user.id,
                             "Функционал:\n/start - запускает меня\n/help - команда для помощи тебе\nБольше у меня функционала нет, я скромный...")
    else:
        BOT_API.send_message(message.from_user.id, "Я тебя не понимаю. Напиши /help.")

@BOT_API.message_handler(content_types=['document'])
def get_document_messages(message):
    if message.document:
        if message.document.mime_type in ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','application/vnd.ms-excel']:
            try:
                file_info = BOT_API.get_file(message.document.file_id)
                downloaded_file = BOT_API.download_file(file_info.file_path)

                with open('Tablica.xlsx', 'wb') as new_file:
                    new_file.write(downloaded_file)

                tabl = pd.read_excel('Tablica.xlsx')

                if tabl.shape[1] >= 6:
                    row5 = tabl.iloc[:, 4].tolist()
                    row6 = tabl.iloc[:, 5].tolist()

                    notcorrect_tema = []

                    for i, (tema, teacher) in enumerate(zip(row6, row5), 1):
                        if not any(re.match(pattern, tema) for pattern in patterns):
                            notcorrect_tema.append(f"{i+1}. Преподаватель: {teacher} | {tema} ")

                    if notcorrect_tema:
                        tema_message = "Неправильные темы уроков:\n"
                        ml = 4096

                        for tema in notcorrect_tema:
                            if len(tema_message) + len(tema) + 1 > ml:
                                BOT_API.send_message(message.from_user.id, tema_message)
                                tema_message = "Неправильные темы уроков:\n"
                            tema_message += tema + "\n"

                        if tema_message != "Неправильные темы уроков:\n":
                            BOT_API.send_message(message.from_user.id, tema_message)
                    else:
                        BOT_API.send_message(message.from_user.id, "Нет неправильных.")

            except Exception as fail:
                BOT_API.send_message(message.from_user.id, f"Ошибка с файлом: {str(fail)}")

BOT_API.polling(none_stop=True, interval=0)
