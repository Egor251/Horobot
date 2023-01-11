# -*- coding: utf-8 -*-
import os
import time
from create_config import create_config
import datetime
try:
    import sqlite3                                      # БД
except ImportError:
    try:
        os.system('pip3 install sqlite3')
        import sqlite3
    except Exception:
        os.system('pip install sqlite3')
        import sqlite3

try:
    import xlrd                                      # Чтение xlsx обязательно версия 1.2.0
except ImportError:
    try:
        os.system('pip3 install xlrd==1.2.0')
        import xlrd
    except Exception:
        os.system('pip install xlrd==1.2.0')
        import xlrd

try:
    import smtplib                                   # отправка писем
except ImportError:
    try:
        os.system('pip3 install smtplib')
        import smtplib
    except Exception:
        os.system('pip install smtplib')
        import smtplib

try:
    import xlsxwriter                               # Создание xlsx
except ImportError:
    try:
        os.system('pip3 install xslxwriter')
        import xlsxwriter
    except Exception:
        os.system('pip install xlsxwriter')
        import xlsxwriter

try:
    from email import encoders  # Модули для формирования письма
    from email.mime.base import MIMEBase
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
except ImportError:
    try:
        os.system('pip3 install email')
        from email import encoders  # Модули для формирования письма
        from email.mime.base import MIMEBase
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText
    except Exception:
        os.system('pip install email')
        from email import encoders  # Модули для формирования письма
        from email.mime.base import MIMEBase
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText

try:
    from transliterate import translit
except ImportError:
    try:
        os.system('pip3 install transliterate')
        from transliterate import translit
    except Exception:
        os.system('pip install transliterate')
        from transliterate import translit

try:
    import zipfile                               # Создание архивов
except ImportError:
    try:
        os.system('pip3 install zipfile')
        import zipfile
    except Exception:
        os.system('pip install zipfile')
        import zipfile

try:
    import shutil                               # Копирование файлов
except ImportError:
    try:
        os.system('pip3 install shutil')
        import shutil
    except Exception:
        os.system('pip install shutil')
        import shutil

try:
    from tqdm import tqdm                               # Копирование файлов
except ImportError:
    try:
        os.system('pip3 install tqdm')
        from tqdm import tqdm
    except Exception:
        os.system('pip install tqdm')
        from tqdm import tqdm
try:
    import logging                               # Логгирование
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
    try:
        import ConfigParser as configparser
    except Exception:
        try:
            os.system('pip3 install ConfigParser')
            import ConfigParser as configparser
        except Exception:
            os.system('pip install ConfigParser')
            import ConfigParser

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s %(name)s %(levelname)s:%(message)s',
                        filename='emails.log')
logger = logging.getLogger(__name__)

path = "settings.ini"
if not os.path.exists(path):
    create_config(path)
config = configparser.ConfigParser()
config.read(path)
admin_mail = config.get("Settings", "admin_mail")
my_mail = config.get("Settings", "email")
my_password = config.get("Settings", "password")
SMTP_server = config.get("Settings", "SMTP_server")
SMTP_port = config.get("Settings", "SMTP_port")
kachanova_mail = config.get("Settings", "kachanova_mail")


def calculate_age(born):
    today = datetime.date.today()
    try:
        birthday = born.replace(year=today.year)
    except ValueError: # raised when birth date is February 29 and the current year is not a leap year
        birthday = born.replace(year=today.year, month=born.month+1, day=1)
    if birthday > today:
        return today.year - born.year - 1
    else:
        return today.year - born.year


def unique(list1):
    unique = []
    for number in list1:
        if number not in unique:
            unique.append(number)
    return unique


# Отправляет письмо с прикреплённым файлом
def send_email(adr, subject, body, file=None):
    # smtpobj = smtplib.SMTP_SSL('smtp.yandex.ru', 465)
    # smtpobj.login('horbot@1dtdm.ru', '12razdva')
    while True:
        try:
            smtpobj = smtplib.SMTP_SSL(SMTP_server, int(SMTP_port))
            smtpobj.login(my_mail, my_password)
            break
        except Exception:
            time.sleep(300)

    sender_email = my_mail
    password = my_password

    # Создание составного сообщения и установка заголовка
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = adr
    message["Subject"] = subject
    # message["Bcc"] = receiver_email   Если у вас несколько получателей

    # Внесение тела письма
    message.attach(MIMEText(body, "plain"))

    if file:
        filename = file  # В той же папке что и код

        # Открытие PDF файла в бинарном режиме
        with open(filename, "rb") as attachment:
            # Заголовок письма application/octet-stream
            # Почтовый клиент обычно может загрузить это автоматически в виде вложения
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        # Шифровка файла под ASCII символы для отправки по почте
        encoders.encode_base64(part)

        # Внесение заголовка в виде пара/ключ к части вложения
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename}",
        )

        # Внесение вложения в сообщение и конвертация сообщения в строку
        message.attach(part)
    text = message.as_string()
    logger.debug(f'Отправляется письмо {adr}')
    logger.debug(f'Отправитель: {sender_email}')
    logger.debug(f'Адресат: {adr}')
    smtpobj.sendmail(sender_email, adr, text)

    logger.exception('Send_mail error: ')
    time.sleep(30)


def make_file(teacher, signal=0):  # формируем файлы для рассылки
    conn = sqlite3.connect("mydatabase.db")  # или :memory: чтобы сохранить в RAM
    cursor = conn.cursor()
    def get_programs(teacher):
        cursor.execute("SELECT program FROM main_db WHERE teacher = ?;", (teacher,))
        list_tmp = cursor.fetchall()
        cursor.execute("SELECT program FROM programs WHERE teacher = ?;", (teacher,))
        list1_tmp = cursor.fetchall()
        list_tmp += list1_tmp
        list_tmp = unique(list_tmp)
        return (list_tmp)
    sig = 0
    list1 = get_programs(teacher)
    head = ['Номер заявления', 'Дата заявления', 'ФИО Ученика', 'Дата рождения', 'ФИО заявителя', 'Телефон',
            'Наименование ДО / образ. программы / услуги', 'Статус заявления', 'ПОМЕТКИ!']
    conn = sqlite3.connect("mydatabase.db")  # или :memory: чтобы сохранить в RAM
    cursor = conn.cursor()
    teacher_tmp = f'{translit(f"{teacher}", "ru", reversed=True)}.xlsx'
    workbook = xlsxwriter.Workbook(f"{teacher_tmp.replace(' ', '_')}")

    bold = workbook.add_format({'bold': True})
    header = workbook.add_format({'bold': True, 'font_size': 11, 'border': True})
    header.set_text_wrap()
    no = workbook.add_format({'font_color': 'Green'})
    usual = workbook.add_format({'border': True})
    usual.set_text_wrap()
    added = workbook.add_format({'bg_color': 'green', 'align': 'center'})
    waiting = workbook.add_format({'bg_color': '95918c', 'align': 'center'})

    # Активные заявки
    if signal == 1 or signal == 0:
        worksheet = workbook.add_worksheet('Заявки')  # Ширина столбцов первого листа
        worksheet.set_column('A:B', 15)
        worksheet.set_column('C:C', 30)
        worksheet.set_column('D:D', 10)
        worksheet.set_column('E:E', 30)
        worksheet.set_column('F:F', 15)
        worksheet.set_column('G:I', 40)

        i = 0
        row = 0
        for i in range(len(head)):  # Заголовок
            worksheet.write(row, i, head[i], header)
        row += 1
        for program in list1:  # Поданы документы будут добавляться только в сентябре и мае
            if datetime.datetime.now().strftime('%m') == '09' or '05':
                sql = f'''SELECT ask_number,  ask_date, child_name, birth_date, parent_name, phone, program, 
                status FROM asks_base WHERE (status ='Ожидание прихода Заявителя для заключения договора' OR status = 
                'Новое' OR status = 'Ожидание подписания электронного договора' OR status = 'Поданы документы') 
                AND program = {"'"+program[0]+"'"};'''
            else:
                sql = f'''SELECT ask_number,  ask_date, child_name, birth_date, parent_name, phone, program, 
                status FROM asks_base WHERE (status ='Ожидание прихода Заявителя для заключения 
                договора' OR status = 'Новое' OR status = 'Ожидание подписания электронного договора') 
                AND program = {"'" + program[0] + "'"};'''
            cursor.execute(sql)
            data = cursor.fetchall()
            row += 1
            worksheet.write(row, 0, program[0], bold)  # Программа
            row += 1
            if len(data) == 0:
                worksheet.write(row, 0, 'Нет заявок', no)
                row += 1
            else:
                if signal != 0:
                    sig = 1
                for i in range(len(data)):
                    j = 0
                    for j in range(len(data[i])):
                        data_tmp = None
                        if j == 1 or j == 3:  # Проверка на наличие даты в ячейке
                            data_tmp = normal_date(data[i][j])
                        else:
                            data_tmp = data[i][j]
                        worksheet.write(row, j, data_tmp, usual)  # Заявки
                    row += 1

    # Все дети
    if signal == 2 or signal == 0:
        head = ['ФИО ребёнка', 'Дата рождения', 'ФИО родителя', 'Телефон', 'Email', 'Школа',
                'Наименование ДО / образ. программы / услуги', 'Группа', 'Дата зачисления', 'ПОМЕТКИ!']
        worksheet = workbook.add_worksheet('Дети')

        worksheet.set_column('A:A', 30)
        worksheet.set_column('B:B', 10)
        worksheet.set_column('C:C', 30)
        worksheet.set_column('D:D', 15)
        worksheet.set_column('E:E', 20)
        worksheet.set_column('F:F', 20)
        worksheet.set_column('G:J', 40)

        i = 0
        row = 0
        for i in range(len(head)):  # Заголовок
            worksheet.write(row, i, head[i], header)
        row += 1
        children = []
        for program in list1:
            sql = f'''SELECT child_name, birth_date, parent_name, phone, email, school, program, group_name, 
            ask_date FROM main_db WHERE  program = {"'" + program[0] + "'"} ORDER BY group_name;'''
            cursor.execute(sql)
            data = cursor.fetchall()
            if len(data) != 0:
                for ii in data:
                    children.append([ii[0], ii[1]])
            worksheet.write(row, 0, program[0], bold)  # Программа
            row += 1
            if len(data) == 0:
                worksheet.write(row, 0, 'Нет детей')
                row += 1
            else:
                for i in range(len(data)):
                    j = 0
                    for j in range(len(data[i])):
                        data_tmp = None
                        if j == 1 or j == 8:
                            data_tmp = normal_date(data[i][j])
                            #logger.exception('Normalize date error: ')
                        else:
                            data_tmp = data[i][j]
                        worksheet.write(row, j, data_tmp, usual)  # Дети
                    row += 1
            row += 1

        # Второй лист

        programs = []
        for item in list1:
            programs.append(item[0].split(',')[0])

        children = sorted(unique(children))
        head = ['ФИО ребёнка', 'Возраст'] + programs

        worksheet = workbook.add_worksheet('Сводная таблица')

        worksheet.set_column('A:A', 30)
        worksheet.set_column('B:B', 8)
        worksheet.set_column('C:Z', 10)

        row = 0
        for i in range(len(head)):  # Заголовок
            worksheet.write(row, i, head[i], header)
        for child in children:

            row += 1
            try:
                age = calculate_age(datetime.date(int(normal_date(child[1]).split('.')[2]),
                                                  int(normal_date(child[1]).split('.')[1]),
                                                  int(normal_date(child[1]).split('.')[0])))
            except ValueError:
                age = 'Упс..'
            worksheet.write(row, 0, child[0], usual)
            worksheet.write(row, 1, age, usual)
            sql = f'''SELECT program FROM main_db WHERE child_name = {"'" + child[0] + "'"} AND 
            teacher = {"'" + teacher + "'"};'''
            cursor.execute(sql)
            data = cursor.fetchall()
            if len(data) != 0:
                for entering in data:
                    worksheet.write(row, list1.index(entering) + 2, '+', added)
            for pr in list1:
                sql = f'''SELECT * FROM asks_base WHERE(status ='Ожидание прихода Заявителя для заключения договора' 
                OR status = 'Новое' OR status = 'Ожидание подписания электронного договора' OR status = 
                'Поданы документы') AND child_name = {"'" + child[0] + "'"} AND program = {"'" + pr[0] + "'"};'''
                cursor.execute(sql)
                data1 = cursor.fetchall()
                if len(data1) != 0:
                    worksheet.write(row, list1.index(pr) + 2, '?', waiting)
        row += 2
        worksheet.write(row, 0, '+', added)
        worksheet.write(row, 1, '-', usual)
        worksheet.write(row, 2, 'Зачислен', usual)
        row += 1
        worksheet.write(row, 0, '?', waiting)
        worksheet.write(row, 1, '-', usual)
        worksheet.write(row, 2, 'Ожидание', usual)
    workbook.close()
    return (sig)


def normal_date(date):     # В ходе парсинга экселя дата выходит неправильная. Исправляем
    test = 0
    try:
        date = int(date)
    except ValueError:
        try:
            date = int(round(float(date)))
        except ValueError:
            pass

    if isinstance(date, str):
        pass
    else:
        str_tmp = str(xlrd.xldate.xldate_as_datetime(date, "%d.%m.%Y"))[:-9]
        date = f'{str(int(str_tmp[-2:])-1).zfill(2)}.{str_tmp[5:7]}.{str(int(str_tmp[0:4])-4)}'
    return date


def get_programs(teacher):
    list_tmp = cursor.execute(f"SELECT program FROM main_db WHERE teacher = ?;", (teacher,))
    list_tmp = unique(list_tmp)
    list_out = []
    for i in list_tmp:
        list_out.append(*i)

    return (list_out)


def ask_mail_bomb():
    refresh_db()
    ask_list = []
    logger.info('Запущена рассылка заявок')
    conn = sqlite3.connect("mydatabase.db")
    cursor = conn.cursor()
    cursor.execute("SELECT date FROM date")
    date = cursor.fetchone()

    subject = "Автоматическая рассылка поступивших заявок ДТДМ Хорошёво"

    body = "Это письмо сгененировано автоматически!\n" \
           "Вы получили это письмо так как у вас есть НЕОБРАБОТАННЫЕ заявки! " \
           "Важно что статус Документы поданы будет выводиться только в сентябре и июне." \
           "В приложенном файле на первом листе заявки, находящиеся в ожидании.\n" \
           "\n" \
           "Удачи и хорошего дня!\n" \
           "\n" \
           "С уважением, ваш Бот ДТДМ Хорошёво."

    cursor.execute("SELECT * FROM emails;")
    result = cursor.fetchall()
    i = 0
    for i in range(len(result)):
        print(result[i][0])
        logger.info(result[i][0])
        print('запуск')
        name = result[i][0] + ' ' + result[i][1] + ' ' + result[i][2]
        name = name.title()
        signal = make_file(name, 1)
        file_name = f'''{translit(f"{name.replace(' ', '_')}", "ru", reversed=True)}.xlsx'''
        if signal == 1:
            ask_list.append(name)
            send_email(result[i][3], subject, body, file_name)
        os.remove(file_name)

    subject = 'Архив с файлами педагогов'

    body = "Это письмо сгененировано автоматически!\n" \
           f"Актуально на {date}\n" \
           f"Заявки есть у следующих педагогов: {ask_list}"
    send_email(kachanova_mail, subject, body)
    logger.error('Ошибка при отправке итогового файла: ', exc_info=True)
    conn.commit()


def mail_bomb(name=None):
    logger.info('Начинаем рассылку детей')
    refresh_db()
    name_check = name
    newzip = zipfile.ZipFile(r'all_teachers_file.zip', 'w')

    conn = sqlite3.connect("mydatabase.db")
    cursor = conn.cursor()
    cursor.execute("SELECT date FROM date")
    date = cursor.fetchone()

    subject = "Автоматическая рассылка списков детей ДТДМ Хорошёво"

    body = "Это письмо сгененировано автоматически! Система всё ещё работает в тестовом режиме " \
           "и переходит полностью в автоматический режим!\n"\
           "Письма, не подходящие под нижеописанные требования, могут быть пропущены!"\
           f"Актуально на {date}" \
           "В приложенном файле на первом листе полный список детей. На втором листе сводная таблица записавшихся к" \
           " вам детей и программ. Является вспомогательной информацией, ничего важного там нет.\n" \
           "Обязательно ознакомьтесь и дайте ответ до 15 числа! " \
           "В случае если что-то пошло не так ТЕХНИЧЕСКИ, то прошу прислать ответное письмо с темой ОШИБКА на этот " \
           "адрес с указанием проблемы. Если вам в списке пришли программы других педагогов то это как раз ошибка!\n" \
           "В случае если каких либо детей нужно отчислить, заявки убрать  и т.д. то прошу прислать ответное письмо" \
           " БЕЗ ИЗМЕНЕНИЯ ТЕМЫ (должно получиться Re: Автоматическая рассылка списков детей ДТДМ Хорошёво или " \
           "что-то вроде) с приложенным исходным файлом. Все ошибки опишите в соседней (справа)" \
           " колонке (Называется ПОМЕТКИ) в файле у соответствующей строки!\n" \
           "Ещё раз, НЕ МЕНЯЙТЕ ИСХОДНЫЕ ДАННЫЕ В ФАЙЛЕ! Только добавляйте комментарии в пустой столбец.\n" \
           "В случае если всё устраивает, то ничего в ответ не присылайте\n" \
           "В случае если есть предолжения, то прошу прислать ответное письмо с темой ПРЕДЛОЖЕНИЕ\n" \
           "В случае если вам пришло несколько одинаковых писем, то не бейте!:)\n" \
           "В случае если вы не хотите получать эту рассылку, то ничем не могу помочь:)\n" \
           "\n" \
           "Удачи и хорошего дня!\n" \
           "\n" \
           "С уважением, ваш Бот ДТДМ Хорошёво."

    if name is None:
        cursor.execute("SELECT * FROM emails;")
        result = cursor.fetchall()
    else:
        name = f"'{name}'"
        cursor.execute(f"SELECT * FROM emails where last_name={name}")
        result = cursor.fetchall()
    i = 0
    logger.info('Запущена рассылка детей')
    for i in range(len(result)):
        print(result[i][0])
        logger.info(result[i][0])
        #print(i)
        print('запуск')
        name = result[i][0] + ' ' + result[i][1] + ' ' + result[i][2]
        name = name.title()
        make_file(name, 2)
        file_name = f'''{translit(f"{name.replace(' ', '_')}", "ru", reversed=True)}.xlsx'''
        send_email(result[i][3], subject, body, file_name)
        newzip.write(file_name)
        os.remove(file_name)
    newzip.close()

    subject = 'Архив с файлами педагогов'

    body = "Это письмо сгененировано автоматически!\n" \
           f"Актуально на (date)"
    if name_check is None:
        send_email(kachanova_mail, subject, body, 'all_teachers_file.zip')
        #send_email('ngushchina@mail.ru', subject, body, 'all_teachers_file.zip')
        logger.error('Ошибка при отправке итогового файла: ', exc_info=True)
    os.remove('all_teachers_file.zip')


def file_bomb():
    newzip = zipfile.ZipFile(r'all_teachers_file.zip', 'w')

    cursor.execute("SELECT * FROM emails;")
    result = cursor.fetchall()
    i = 0
    try:
        os.mkdir('files')
    except FileExistsError:
        pass
    for i in range(len(result)):
        print(result[i][0])
        print(i)
        name = result[i][0] + ' ' + result[i][1] + ' ' + result[i][2]
        name = name.title()

        make_file(name)
        new_name = translit(f"{name.replace(' ', '_')}", "ru", reversed=True)
        newzip.write(f'{new_name}.xlsx')
        try:
            os.rename(f'{new_name}.xlsx', f'files\\{new_name}.xlsx')
        except FileExistsError:
            os.remove(f'files\\{new_name}.xlsx')
            os.rename(f'{new_name}.xlsx', f'files\\{new_name}.xlsx')
            pass
    newzip.close()


def refresh_db():
    def create_db():
        conn = sqlite3.connect("mydatabase.db")  # или :memory: чтобы сохранить в RAM
        cursor = conn.cursor()

        # Создание таблицы
        cursor.execute("""CREATE TABLE IF NOT EXISTS main_db
                          (child_name text, birth_date text, parent_name text, phone text, email text,
                           school text, program text, program_level text, finances text, program_type text,
                            group_name text, period text, teacher text, ask_date text, start_date text, end_date text)
                       """)
        cursor.execute("""CREATE TABLE IF NOT EXISTS asks_base
                              (ask_number integer, source text, ask_date text, child_name text, birth_date text,
                               parent_name text, phone text, program text, organization text, adress text,
                                status text)
                           """)
        cursor.execute("""CREATE TABLE IF NOT EXISTS emails
                                  (last_name text, first_name text, second_name text, email text)
                               """)
        cursor.execute("""CREATE TABLE IF NOT EXISTS date
                                      (date text)
                                   """)
        cursor.execute("""CREATE TABLE IF NOT EXISTS programs
                                  (program text,  teacher text)""")
        conn.commit()

    def parse_excel_main(file):
        conn = sqlite3.connect("mydatabase.db")
        cursor = conn.cursor()
        rb = xlrd.open_workbook(file)
        sheet = rb.sheet_by_index(0)
        for rownum in range(7, sheet.nrows):
            row = sheet.row_values(rownum)
            row_tmp = ''.join(c for c in row[7] if c not in "'")  # помещаем сроку в ковычки
            row[7] = row_tmp

            cursor.executemany("INSERT INTO main_db VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (row[1:],))
        conn.commit()

    def parse_excel_asks(file):
        conn = sqlite3.connect("mydatabase.db")
        cursor = conn.cursor()
        rb = xlrd.open_workbook(file)
        sheet = rb.sheet_by_index(0)
        for rownum in range(21, sheet.nrows):
            row = sheet.row_values(rownum)
            row_tmp = ''.join(c for c in row[7] if c not in "'")  # помещаем сроку в ковычки
            row[7] = row_tmp
            cursor.executemany("INSERT INTO asks_base VALUES (?,?,?,?,?,?,?,?,?,?,?)", (row,))
        conn.commit()

    def parse_excel_email(file):
        conn = sqlite3.connect("mydatabase.db")
        cursor = conn.cursor()
        rb = xlrd.open_workbook(file)
        sheet = rb.sheet_by_index(0)
        for rownum in range(1, sheet.nrows):
            row = sheet.row_values(rownum)
            for i in range(len(row)):
                row[i] = row[i].lower().capitalize()
            cursor.executemany("INSERT INTO emails VALUES (?,?,?,?)", (row,))
        conn.commit()

    def parse_excel_programs(file):
        conn = sqlite3.connect("mydatabase.db")
        cursor = conn.cursor()
        rb = xlrd.open_workbook(file)
        sheet = rb.sheet_by_index(0)
        wrong_list = []
        for rownum in tqdm(range(13, sheet.nrows)):

            # Объявление переменных на всякий случай
            teacher_last_name = ''
            teacher_first_letter = ''
            teacher = ''

            row = sheet.row_values(rownum)
            #print(row[1])
            row[1] = row[1].replace("'", "")  # Госпожа Пономарёва любит использовать в названии программы кавычки
            list_tmp = row[1].split(',')  # разбиение по запятой, ожидаем 3 значения
            if len(list_tmp) == 1 or len(list_tmp) == 0:  # Может не быть запятых вообще
                pass
            else:
                if list_tmp[-1] == '' or ' ':  # Если есть запятая в конце
                    if len(list_tmp[-2].split('.')) != 3:  # может быть не указан педагог (Или указан не в том месте)
                        pass
                    else:
                        teacher = list_tmp[-2]
                else:
                    if len(list_tmp[-1].split(
                            '.')) != 3:  # может быть не указан педагог (совершенно внезапно в нужном месте)
                        pass
                    else:
                        teacher = list_tmp[-1]
            teacher1 = teacher.split(' ')
            #print(teacher1)
            #if len(teacher1) == 0:
            if len(teacher1) <= 1:
                wrong_list.append(row[1])
                pass
            else:
                if teacher1[1] == '':  # Начиная с этого момента я перестаю понимать что происходит
                    teacher_last_name = teacher1[2]
                    teacher_first_letter = teacher1[3][0]
                else:
                    teacher_last_name = teacher1[1]
                    try:
                        teacher_first_letter = teacher1[2][0]
                    except IndexError:
                        try:
                            teacher_first_letter = teacher1[3][0]
                        except Exception:
                            pass

            teacher_last_name = teacher_last_name.replace('ё', 'е')
            # Как не зная sql мутить ультра клёвые решения на sql?
            cursor.execute("SELECT last_name, first_name, second_name FROM emails WHERE last_name = ?;",
                           (teacher_last_name,))
            data = cursor.fetchall()
            string_tmp = ''
            if len(data) == 0:
                wrong_list.append(teacher_last_name)
                pass
            elif len(data) > 1:
                for item in data:
                    if item[1][0] == teacher_first_letter:  # Вот так!
                        data = item
            if isinstance(data, list):
                try:
                    string_tmp = str(data[0][0]) + ' ' + str(data[0][1]) + ' ' + str(data[0][2])
                except Exception:
                    pass
            elif isinstance(data, tuple):
                string_tmp = str(data[0]) + ' ' + str(data[1]) + ' ' + str(data[2])

            program_list = [row[1], string_tmp]
            cursor.executemany("INSERT INTO programs VALUES (?,?)", (program_list,))
        wrong_list = unique(wrong_list)
        body = ''
        for item in wrong_list:
            body += item
            body += '\n'

        send_email(admin_mail, 'Людей нет в рассылке', body)  # Отправляем письмо со списком отсутствущих
        conn.commit()

    dir_files = os.listdir()
    date = ''
    for name in dir_files:
        if name[:-16] == 'Ведомость учащихся':
            for name1 in dir_files:
                if name1[:-16] == 'Backup_Ведомость учащихся':
                    os.remove(name1)
            shutil.copyfile(name, f'Backup_{name}')
            try:
                os.remove('db.xlsx')
            except Exception:
                pass
            os.rename(name, 'db.xlsx')
            date = name[-15:-5]
        elif name[:-16] == 'Реестр заявлений':
            for name1 in dir_files:
                if name1[:-16] == 'Backup_Реестр заявлений':
                    try:
                        os.remove(name1)
                    except FileNotFoundError:
                        pass
            shutil.copyfile(name, f'Backup_{name}')
            try:
                os.remove('asks.xlsx')
            except Exception:
                pass
            os.rename(name, 'asks.xlsx')
        elif name[:-16] == 'Реестр детских объединений':
            for name1 in dir_files:
                if name1[:-16] == 'Backup_Реестр детских объединений':
                    os.remove(name1)
            shutil.copyfile(name, f'Backup_{name}')
            try:
                os.remove('programs.xlsx')
            except Exception:
                pass
            os.rename(name, 'programs.xlsx')

    if date == '':
        for name in dir_files:
            if name[:-16] == 'Backup_Ведомость учащихся':
                date = name[-15:-5]

    if not os.path.exists('mydatabase.db'):
        print('Файла с БД нет. Создаю')
        create_db()

    conn = sqlite3.connect("mydatabase.db")  # или :memory: чтобы сохранить в RAM
    cursor = conn.cursor()
    # Удаляем таблицы
    cursor.execute('DROP TABLE IF EXISTS main_db')
    cursor.execute('DROP TABLE IF EXISTS asks_base')
    cursor.execute('DROP TABLE IF EXISTS emails')
    cursor.execute('DROP TABLE IF EXISTS date')
    cursor.execute('DROP TABLE IF EXISTS programs')
    conn.commit()
    create_db()

    try:
        parse_excel_main('db.xlsx')
    except FileNotFoundError:
        logger.debug('отсутствует файл db.xlsx')
        pass
    try:
        parse_excel_asks('asks.xlsx')
    except FileNotFoundError:
        logger.debug('отсутствует файл asks.xlsx')
        pass
    try:
        parse_excel_email('email.xlsx')
    except FileNotFoundError:
        logger.debug('отсутствует файл email.xlsx')
        raise FileNotFoundError("Отсутствует список почтовых адресов сотрудников")
    try:
        parse_excel_programs('programs.xlsx')
    except FileNotFoundError:
        logger.debug('отсутствует файл programs.xlsx')
        pass

    conn = sqlite3.connect("mydatabase.db")  # или :memory: чтобы сохранить в RAM
    cursor = conn.cursor()
    date = date.replace('-', '.')
    cursor.execute('INSERT INTO date (date) VALUES (?)', (date,))
    conn.commit()


path = "settings.ini"
if not os.path.exists(path):
    create_config(path)
config = configparser.ConfigParser()
config.read(path)
my_mail = config.get("Settings", "email")
my_password = config.get("Settings", "password")
SMTP_server = config.get("Settings", "SMTP_server")
SMTP_port = config.get("Settings", "SMTP_port")


if __name__ == '__main__':

    refresh_db()

    conn = sqlite3.connect("mydatabase.db")  # или :memory: чтобы сохранить в RAM
    cursor = conn.cursor()
    #make_file('Афанасьев Вячеслав Игоревич')
    #mail_bomb()
    #mail_bomb(name)
    # file_bomb()
    #make_file(name)
    # test_mail_bomb()
