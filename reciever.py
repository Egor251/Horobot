# -*- coding: utf-8 -*-
import os
import main
from main import send_email
from minusing import minusing
from minusing_alternative import minusing_alternative
from create_config import create_config
from email.header import decode_header
import time
from enrollment import enrollment

try:
    import poplib                                       # приём pop3 писем
except ImportError:
    try:
        os.system('pip3 install poplib')
        import poplib
    except Exception:
        os.system('pip install poplib')
        import poplib

try:
    import imaplib                                      # приём imap писем
except ImportError:
    try:
        os.system('pip3 install imaplib')
        import imaplib
    except Exception:
        os.system('pip install imaplib')
        import imaplib

try:
    import email                                        # приём pop3 писем
except ImportError:
    try:
        os.system('pip3 install email')
        import email
    except Exception:
        os.system('pip install email')
        import email

try:
    import base64
except ImportError:
    try:
        os.system('pip3 install base64')
        import base64
    except Exception:
        os.system('pip install base64')
        import base64


try:
    import logging                                      # Логгирование
except ImportError:
    try:
        os.system('pip3 install logging')
        import logging
    except Exception:
        os.system('pip install logging')
        import logging

try:
    import configparser
except ImportError:
    import ConfigParser as configparser

path = "settings.ini"
if not os.path.exists(path):
    create_config(path)

config = configparser.ConfigParser()
config.read(path)

my_mail = config.get("Settings", "email")
my_password = config.get("Settings", "password")
POP3_server = config.get("Settings", "POP3_server")
POP3_port = config.get("Settings", "POP3_port")
admin_mail = config.get('Settings', 'admin_mail')
kachanova_mail = config.get("Settings", "kachanova_mail")

def get_imap():
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s %(name)s %(levelname)s:%(message)s', filename='emails.log')
    logger = logging.getLogger(__name__)
    logger.debug('Старт функции get_imap')

    def parse_mail(id):
        result, data = mail.fetch(id, "(RFC822)")   # Получаем тело письма (RFC822) для данного ID

        message = email.message_from_bytes(data[0][1])

        # If the message is multipart, it basically has multiple emails inside
        # so you must extract each "submail" separately.
        '''if message.is_multipart():
            print('Multipart types:')
            for part in message.walk():
                print(f'- {part.get_content_type()}')
            multipart_payload = message.get_payload()
            #for sub_message in multipart_payload:
                # The actual text/HTML email contents, or attachment data
                #print(f'Payload\n{sub_message.get_payload()}')
        else:  # Not a multipart message, payload is simple string
            print(f'Payload\n{message.get_payload()}')'''
        # You could also use `message.iter_attachments()` to get attachments only

        text = message["subject"]

        message_from = []
        try:
            message_from = message["from"].split()
            if len(message_from) > 1:
                sender = message["from"].split()[1][1:-1]   # Выковыриваем адрес отправителя
            else:
                sender = message["from"].split()[0]
        except IndexError:
            sender = 'yandex'
        print(f'message_from = {message_from}')
        print(f'sender = {sender}')
        try:
            subject = decode_header(str(text))
            logger.exception('decode_header: ')
        except Exception:
            try:
                subject = str(base64.b64decode(text[10:-2]), 'utf-8')   # ЗАПОМНИТЬ СТРОКУ!! ДЕКОДИРОВАНИЕ ЗАГОЛОВКОВ ПИСЬМА!
            except Exception:
                subject = text.replace('\r', '')
                subject = subject.replace('\n', '')
                subject = subject.replace('=?UTF-8?B?', '')
                subject = subject.replace('?=', '')
                subject = subject.replace(' ', '')
                subject = str(base64.b64decode(subject), 'utf-8')

        try:
            subject = subject[0][0].decode(subject[0][1])
        except TypeError:
            subject = subject[0][0].decode('utf-8')
            if sender == 'patria@list.ru' and subject == 'Re: ':
                subject = 're: автоматическая рассылка списков детей дтдм хорошёво'

        subject = subject.lower()

        print(f'Получено письмо от:{str(sender)}, тема: {str(subject)}')
        logger.debug(f'Получено письмо от:{str(sender)}, тема: {str(subject)}')

        signal = 0

        if subject[-1] == ' ':
            subject = subject[:-1]
        if subject == 're: aвтоматическая рассылка списков детей дтдм хорошёво' or subject[:11] == 're: автомат'\
                or subject[:12] == 'fwd: автомат' or subject == 're: у вас есть необработанные заявки в дтдм хорошёво':
            signal = 1
        elif subject == 'ошибка' or subject == 'предложение':
            signal = 3
        elif subject == 'рассылка':
            signal = 2
        elif subject == 'приказ':
            signal = 5
        elif subject == 'приказ 2':
            signal = 6
            try:
                os.remove('до.xlsx')
            except Exception:
                pass
            try:
                os.remove('после.xlsx')
            except Exception:
                pass
            try:
                os.remove('result.xlsx')
            except Exception:
                pass
        elif subject == 'заявки':
            signal = 4
        elif subject == 'информация':
            signal = 0
            body = 'Инструкция по использованию:\n' \
                   'Возможные темы письма:\n' \
                   '1) Рассылка (приложить ведомость, при необходимости список программ,можно и' \
                   ' список заявок просто так' \
                   '(всё выгружается из ЕСЗ) - Делает рассылку списков детей ' \
                   'по всем педагогам.\n' \
                   '2) Заявки (приложить список заявок, ведомость и прогаммы ' \
                   'по желанию) - Отправляет письма только тем педагогам,' \
                   ' у кого есть необработанные заявки.\n' \
                   '3) Ошибка или Предложение (описать идею или проблему) - перенаправит письмо на почту' \
                   ' ответственного лица.\n' \
                   '4) Приказ (приложить ведомости за два месяца с именами до.xlsx и после.xlsx) - Вычислит список ' \
                   'отчисленных детей за два месяца и отправит обратно.\n' \
                   '5) Re: Автоматическая рассылка списков детей дтдм хорошёво (приложить изменённый файл из ' \
                   'полученной рассылки) - Сформирует список изменений и отправит ответственному лицу.\n' \
                   '6) Информация - Пришлёт это письмо.\n' \
                   '\n' \
                   'Никаких Fwd: или Re: кроме пункта 5! Темы должны быть именно такими как здесь указано!\n'\
                   'В случае неправильного использования есть вероятность, ' \
                   'что бот словит экзистенциальный кризис и откажется работать!'
            send_email(sender, "Инструкция по использованию бота", body)
        elif subject == 'зачисление':
            signal = 10

        logger.info(f'Signal = {signal}')

        #for string in message.get_payload():        # Чтение содержимого
        #    print(str(base64.b64decode(string.get_payload()), 'utf-8'))

        file_signal = 0
        file_debug_count = 0

        if sender != 'yandex':
            for part in message.walk():   # Находит имя файла. Хз как, не трогать!
                if "application" in part.get_content_type():
                    filename = part.get_filename()
                    filename = str(email.header.make_header(email.header.decode_header(filename)))
                    file_debug_count = len(os.listdir('download'))
                    print(filename)
                    logger.info(f'Файл: {filename}')

                    if signal == 2 or signal == 4:  #рассылка или заявки
                        fp = open(filename, 'wb')
                    else:
                        try:
                            fp = open(f'download/{filename}', 'wb')
                            file_signal = 1
                        except Exception:
                            os.mkdir('download')
                            fp = open(f'download/{filename}', 'wb')
                            file_signal = 1
                    fp.write(part.get_payload(decode=True))
                    fp.close()

        if file_debug_count != len(os.listdir('download')):
            file_signal = 1

        if signal == 1 and file_signal == 0:
            body = 'Прошу не присылать ответные письма без изменения темы и без приложенного файла!' \
                   ' Соблюдайте инструкцию!'
            logger.warning(f'Не соблюдение инструкции {sender}')
            send_email(sender, "Re: Автоматическая рассылка списков детей дтдм хорошёво", body)
            signal = 0

        elif signal == 5:
            try:
                send_email(sender, 'Результат вычитания', ' Система в рабочем режиме! Живые люди больше" \
               " не проверяют этот почтовый ящик!', minusing('download/до.xlsx', 'download/после.xlsx'))
                logger.info('Получен список для формирования приказа, отправил результат обратно')
            except FileNotFoundError:
                send_email(sender, 'Ошибка при формировании списка', 'Перепроверьте отправленные данные')
                logger.exception('Получен список для формирования приказа, словил ошибку: ')
            signal = 0

        elif signal == 6:
            try:
                send_email(sender, 'Результат вычитания', ' Система в рабочем режиме! Живые люди больше" \
               " не проверяют этот почтовый ящик!', minusing_alternative('download/до.xlsx', 'download/после.xlsx'))
                logger.info('Получен список для формирования приказа (альтернатива), отправил результат обратно')
            except FileNotFoundError:
                send_email(sender, 'Ошибка при формировании списка', 'Перепроверьте отправленные данные')
                logger.exception('Получен список для формирования приказа, словил ошибку: ')
            signal = 0

        elif signal == 3:
            try:
                if subject == 'ошибка':
                    base_message_subject = 'Ошибка в рассылке'
                elif subject == 'предложение':
                    base_message_subject = 'Предложение по рассылке'
                else:
                    base_message_subject = 'Кто-то что-то хочет по рассылке'

                error_message = ''
                if message.is_multipart():
                    print('Multipart types:')

                    for part in message.walk():
                        print(f'- {part.get_content_type()}')

                    multipart_payload = message.get_payload()

                    for sub_message in multipart_payload:
                        # The actual text/HTML email contents, or attachment data
                        error_message += sub_message.get_payload()
                else:  # Not a multipart message, payload is simple string
                    error_message = message.get_payload()
                try:
                    error_message = decode_header(str(error_message))
                except Exception:
                    try:
                        error_message = str(base64.b64decode(text[10:-2]),
                                      'utf-8')  # ЗАПОМНИТЬ СТРОКУ!! ДЕКОДИРОВАНИЕ ЗАГОЛОВКОВ ПИСЬМА!
                    except Exception:
                        error_message = error_message.replace('\r', '')
                        error_message = error_message.replace('\n', '')
                        error_message = error_message.replace('=?UTF-8?B?', '')
                        error_message = error_message.replace('?=', '')
                        error_message = error_message.replace(' ', '')
                        error_message = str(base64.b64decode(subject), 'utf-8')
                error_message = error_message[0][0].split('<')[1][4:]
            except Exception:
                error_message = 'Что-то пошло не так при декодировании содержимого письма, смотри сам!'
            logger.debug(f'Сообщение об ошибке: {error_message}')
            #error_message = str(base64.b64decode(error_message), 'utf-8')

            body = f'{sender} Сообщил: \n {error_message}'
            send_email(admin_mail, base_message_subject, body)
            logger.warning(admin_mail, base_message_subject, body)
            signal = 0

        elif signal == 10:
            try:
                send_email(sender, 'Результат на Зачисление', ' Система в рабочем режиме! Живые люди больше" \
               " не проверяют этот почтовый ящик!', enrollment('download/назачисление.xlsx'))
                logger.info('Сформирован список для приказов на Зачисление, результат отправлен обратно')
            except FileNotFoundError:
                send_email(sender, 'Ошибка при формировании списка для приказов на Зачисление', 'Перепроверьте отправленные данные')
                logger.exception('Получен список для формирования приказа на Зачисление, получена ошибка: ')
            signal = 0

        print(signal)
        return(signal)

    while True:
        try:
            mail = imaplib.IMAP4_SSL('imap.yandex.ru')
            mail.login(my_mail, my_password)
            mail.list() # Выводит список папок в почтовом ящике.
            mail.select("inbox")    # Подключаемся к папке "входящие".
            logger.debug('Успешное подключение к почтовому ящику')
            break
        except Exception:
            logger.warning('Не удалось подключиться к почтовому ящику')
            time.sleep(300)
    result, message = mail.search(None, 'UNSEEN')

    ids = message[0]  # Получаем строку номеров писем
    id_list = ids.split()  # Разделяем ID писем

    recive_signal = 0
    send_signal = 0

    for id in id_list:
        signal = parse_mail(id)

        if signal == 1 and recive_signal != 1:
            recive_signal = 1
        elif signal == 2 and send_signal != 1:
            send_signal = 1
        if signal == 4: # заявки
            #main.refresh_db()
            main.ask_mail_bomb()

    return (recive_signal, send_signal)

def get_pop3():
    user = my_mail
    # Connect to the mail box
    Mailbox = poplib.POP3_SSL(POP3_server, POP3_port)
    Mailbox.user(user)
    Mailbox.pass_(my_password)
    pop3info = Mailbox.stat()
    mailcount = pop3info[0]  # total email
    print("Total no. of Email : ", mailcount)
    print("\n\nStart Reading Messages\n\n")
    #for i in range(mailcount):
    #    for message in Mailbox.retr(i + 1)[1]:
    #        print(message)
    print(pop3info)
    Mailbox.quit()

path = "settings.ini"
if not os.path.exists(path):
    create_config(path)
config = configparser.ConfigParser()
config.read(path)
my_mail = config.get("Settings", "email")
my_password = config.get("Settings", "password")
POP3_server = config.get("Settings", "POP3_server")
POP3_port = config.get("Settings", "POP3_port")
admin_mail = config.get('Settings', 'admin_mail')

if __name__ == '__main__':

    path = "settings.ini"
    if not os.path.exists(path):
        create_config(path)
    config = configparser.ConfigParser()
    config.read(path)
    my_mail = config.get("Settings", "email")
    my_password = config.get("Settings", "password")
    POP3_server = config.get("Settings", "POP3_server")
    POP3_port = config.get("Settings", "POP3_port")
    admin_mail = config.get('Settings', 'admin_mail')
    kachanova_mail = config.get("Settings", "kachanova_mail")
    get_imap()
