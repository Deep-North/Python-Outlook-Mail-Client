import calendar
import os
import shutil
import time
import traceback
import logging
from datetime import datetime, timedelta, date

import win32com
import win32com.client

import Setup as S


def clean():
    try:
        list_of_files = os.listdir(S.PATH_INPUT)
    except (FileNotFoundError, PermissionError):
        list_of_files = None

    if list_of_files is None:
        pass
    else:
        flag = False
        logger = logging.getLogger('MyProject.fileCleaner.Clean')
        for file in list_of_files:
            if ((os.path.isfile(f'{S.PATH_INPUT}{file}')) and not (file.endswith(".xlsx"))):# Проверка на то, файл ли это и не эксель ли он
                try:
                    os.remove(f'{S.PATH_INPUT}{file}') # удаляем все лишние файлы
                    flag = True
                except (FileNotFoundError, PermissionError):
                    print('Ошибка:\n', traceback.format_exc())
                    logger.exception('Exception occurred', exc_info=True)

            else:
                pass
        if flag == True:
            pass
            #logger.info('Папка очищена от лишних файлов.')
            #print('Папка очищена от лишних файлов.')



def LogArchivator():
    if not os.path.isdir(S.LOG_PATH):
        os.makedirs(S.LOG_PATH)
    if not os.path.isdir(S.LOG_ARCH_PATH):
        os.makedirs(S.LOG_ARCH_PATH)
    logger = logging.getLogger('MyProject.fileCleaner.LogArchivator')
    if (datetime.now().day == 1) and (datetime.now().hour == 0) and (datetime.now().minute in range(0,9)):
        today = date.today()
        days = calendar.monthrange(today.year, today.month)[1]
        previous_month_date = today - timedelta(days=days)
        previous_month = previous_month_date.strftime("%Y-%m")
        file = 'log.txt'
        try:
            os.rename(f'{file}', f'{previous_month}_{file}')
            shutil.move(f'{S.LOG_PATH}{previous_month}_{file}', f'{S.LOG_ARCH_PATH}')
            #os.chdir(S.LOG_ARCH_PATH)

            logger.info('Файл лога успешно перемещен в архив.')
            print ('Файл лога успешно перемещен в архив.')
            time.sleep(600)
        except:
            logger.exception('Не удалось переместить файл лога в архив. Возможно файл занят в данный момент.', exc_info=True)
            pass


def saveAttachments(message):
    logger = logging.getLogger('MyProject.mailClient.saveAttachments')
    att = message.Attachments
    if not os.path.isdir(S.PATH_INPUT):
        os.makedirs(S.PATH_INPUT)
    try:
        for i in att:
            i.SaveAsFile(os.path.join(S.PATH_INPUT, i.FileName))
    except:
        print('Ошибка:\n', traceback.format_exc())
        logger.error('Exception occurred', exc_info=True)
    try:
        list_of_files = os.listdir(S.PATH_INPUT)
    except (FileNotFoundError, PermissionError):
        logger.error(S.missingAttachment)
        answer(message, S.missingAttachment)


def getAllNewMessages():
    logger = logging.getLogger('MyProject.mailClient.getAllNewMessages')
    #logger.info('Начало получения почты')
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    #logger.info('Получен почтовый ID')
    #logger.info(S.EMAIL_ACCOUNT)
    box = outlook.Folders(S.EMAIL_ACCOUNT)#Выбираем нужный ящик (аккаунт)
    #logger.info('Выбран нужный ящик')
    folder = box.Folders("Входящие")
    #logger.info('Выбран нужный почтовый ящик')
    messages = folder.Items
    message = messages.GetFirst()
    while message:  # Пока есть письма для чтения, читаем.
        if message.UnRead:
            saveAttachments(message) # Сохраняем вложения
            clean() # Удаляем все файлы кроме экселевских
            message.UnRead = False  # Помечаем письмо прочитанным
            logger.info('Получено новое письмо от ' + str(message.sender) + ' Тема письма: ' + str(message.subject) + '. Запускается обработка вложений.')
            
            # Тут будет происходить обработка экселевского файла
            # поэтому пока ставлю заглушку
            pathToAtt = r'C:\Users\g-luc\YandexDisk\python code\mail client\test.txt'
            message_text = 'test'
            
            # Формируем ответ на письмо
            if pathToAtt is None or pathToAtt == '':
                # Если нет вложений, то ответ без них
                answer(message, message_text)
            else:
                # Если есть, отправляем ответное письмо с вложением
                answerWithAttachment(message, message_text, pathToAtt)
                pathToAtt = '' # Обнуляем путь к файлу вложения во избежание повторного прикрепления к другому письму
        message = messages.GetNext()


def answer(message, text):
    logger = logging.getLogger('MyProject.mailClient.answer')
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.SentOnBehalfOfName = S.EMAIL_SENDER_ACCOUNT # Задается отправитель письма. Нужны права на этот ящик,
    # иначе выдаст ошибку.

    mail.Subject = 'Re: ' + message.subject
    if text != '':
        mail.Body = text  # Помещаем в тело письма текст, переданный в качестве аргумента
    else:
        mail.Body = 'Спасибо, мы получили ваше сообщение.' # Или заранее заданный текст

    #mail.To = str(message.sender)
    #mail.Recipients.Add(str(message.sender))

    # Вместо mail.To лучше использовать следующую функцию, т.к. Outlook иногда тупит и не может распознать
    # пользователя
    messageRecipient = str(message.sender)
    m = mail.Recipients.Add(messageRecipient)
    # Проверяем имя отправителя на корректность. Используем метод из стандартной библиотеки майкрософт.
    m.Resolve()
    if m.Resolved:
        try:
            mail.Send()
            logger.info('Ответ отправлен')
        except:
            print('Ошибка:\n', traceback.format_exc())
            logger.error('Exception occurred', exc_info=True)
    else:
        logger.error('Ошибка Outlook. Невозможно распознать адрес отправителя.')

def answerWithAttachment(message, text, path):
    logger = logging.getLogger('MyProject.mailClient.answer')
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.SentOnBehalfOfName = S.EMAIL_SENDER_ACCOUNT # Задается отправитель письма. Нужны права на этот ящик, иначе выдаст ошибку.
    #mail.To = str(message.sender)
    mail.Subject = 'Re: ' + message.subject
    if text != '':
        mail.Body = text  # Помещаем в тело письма текст, переданный в качестве аргумента
    else:
        mail.Body = 'Спасибо, мы получили ваше сообщение.' # Или заранее заданный текст

    # Прикрепляем вложение
    attachment = path
    mail.Attachments.Add(attachment)

    messageRecipient = str(message.sender)
    m = mail.Recipients.Add(messageRecipient)
    m.Resolve()
    if m.Resolved:
        try:
            mail.Send()
            logger.info('Ответ отправлен')
        except:
            print('Ошибка:\n', traceback.format_exc())
            logger.error('Ошибка отправки сообщения.', exc_info=True)
    else:
        logger.error('Ошибка Outlook. Невозможно распознать адрес отправителя.')

def sendMessageToAdmin(error):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.SentOnBehalfOfName = S.EMAIL_SENDER_ACCOUNT  # Задается отправитель письма. Нужны права на этот ящик, иначе выдаст ошибку.
    mail.To = S.ADMIN_ACCOUNT
    mail.Subject = 'Ошибка робота!'
    mail.Body = f'Внимание!\nПроизошла критическая ошибка робота!\n{error}'
    try:
        mail.Send()
    except:
        pass