# -*- coding: utf-8 -*-
import os
from reciever import get_imap
import main
from create_config import create_config
import time

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
    import xlsxwriter                               # Создание xlsx
except ImportError:
    try:
        os.system('pip3 install xslxwriter')
        import xlsxwriter
    except Exception:
        os.system('pip install xlsxwriter')
        import xlsxwriter

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
    import logging                               # Логгирование
except ImportError:
    try:
        os.system('pip3 install logging')
        import logging
    except Exception:
        os.system('pip install logging')
        import logging

try:
    import schedule                               # cron
except ImportError:
    try:
        os.system('pip3 install schedule')
        import schedule
    except Exception:
        os.system('pip install schedule')
        import schedule

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


def renew_log():
    try:
        size = os.path.getsize('emails.log')
        if size > 5000000:
            os.remove('emails.log')
    except PermissionError:
        pass


def __init__():
    renew_log()
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s %(name)s %(levelname)s:%(message)s',
                        filename='emails.log')
    logger = logging.getLogger(__name__)

    logger.info('Запуск программы!')

    def make_sum_file():
        list_files = os.listdir('download')
        conn = sqlite3.connect("mydatabase.db")
        cursor = conn.cursor()
        cursor.execute("SELECT date FROM date")
        date = cursor.fetchone()
        workbook = xlsxwriter.Workbook(f"Total_file{date[2:-3]}.xlsx")


        worksheet = workbook.add_worksheet('Заявки')  # Ширина столбцов первого листа
        worksheet.set_column('A:B', 15)
        worksheet.set_column('C:C', 30)
        worksheet.set_column('D:D', 10)
        worksheet.set_column('E:E', 30)
        worksheet.set_column('F:F', 15)
        worksheet.set_column('G:I', 40)

        bold = workbook.add_format({'bold': True})
        header = workbook.add_format({'bold': True, 'font_size': 11, 'border': True})
        header.set_text_wrap()
        no = workbook.add_format({'font_color': 'Green'})
        usual = workbook.add_format({'border': True})
        usual.set_text_wrap()

        head = ['Номер заявления', 'Дата заявления', 'ФИО Ученика', 'Дата рождения', 'ФИО заявителя', 'Телефон',
                'Наименование ДО / образ. программы / услуги', 'Статус заявления', 'ПОМЕТКИ!']
        i = 0
        row_iter = 0
        for i in range(len(head)):  # Заголовок
            worksheet.write(row_iter, i, head[i], header)
        row_iter += 1
        for file in list_files:
            row_iter_checker = 0
            row_iter += 1
            worksheet.write(row_iter, 0, str(file)[:-5], bold)
            row_iter += 1
            file = f'download/{file}'
            rb = xlrd.open_workbook(file)
            sheet = rb.sheet_by_index(0)
            program = ''
            if len(sheet.row_values(0)) == 9:
                for rownum in range(1, sheet.nrows):
                    i = 0
                    row = sheet.row_values(rownum)
                    if row[0] == '':
                        program = sheet.row_values(rownum+1)[0]
                    if row[8] != '':
                        if program != '':
                            row_iter += 1
                            row_iter_checker += 1
                            worksheet.write(row_iter, i, program, bold)
                            program = ''
                            row_iter += 1
                        for i in range(len(row)):
                            worksheet.write(row_iter, i, row[i], usual)
                        row_iter += 1
                        row_iter_checker += 1
            if row_iter_checker == 0:
                worksheet.write(row_iter, 0, 'Нет пометок', no)
                row_iter += 1

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
        row_iter = 0
        for i in range(len(head)):  # Заголовок
            worksheet.write(row_iter, i, head[i], header)
        row_iter += 1
        for file in list_files:
            logger.debug(file)
            row_iter_checker = 0
            row_iter += 1
            worksheet.write(row_iter, 0, str(file)[:-5], bold)
            row_iter += 1
            file = f'download/{file}'
            rb = xlrd.open_workbook(file)
            sheet = rb.sheet_by_index(0)
            program = ''
            if len(sheet.row_values(0)) == 10:
                for rownum in range(1, sheet.nrows):
                    i = 0
                    row = sheet.row_values(rownum)
                    if row[0] == '':
                        program = sheet.row_values(rownum+1)[0]
                    if row[9] != '':
                        if program != '':
                            row_iter += 1
                            worksheet.write(row_iter, i, program, bold)
                            program = ''
                            row_iter += 1
                            row_iter_checker += 1
                        for i in range(len(row)):
                            worksheet.write(row_iter, i, row[i], usual)
                        row_iter += 1
                        row_iter_checker += 1

            if row_iter_checker == 0:
                worksheet.write(row_iter, 0, 'Нет пометок', no)
                row_iter += 1
        workbook.close()
        return(f"Total_file{date[2:-3]}.xlsx")

    def clear_folder(folder):
        list_files = os.listdir(folder)
        for file in list_files:
            os.remove(f'{folder}/{file}')

    clear_folder('download')
    clear_folder('files')

    recive_signal = 0
    send_signal = 0
    path = "settings.ini"
    if not os.path.exists(path):
        create_config(path)
    recive_signal, send_signal = get_imap()
    logger.info(f'recive_signal={recive_signal}, send_signal = {send_signal}')
    if recive_signal != 0:
        total_file = make_sum_file()
        main.send_email(kachanova_mail, 'Ответ от педагогов', 'Письмо сгенерированно автоматически!', total_file)
        os.remove(total_file)
    if send_signal != 0:
        main.refresh_db()
        main.mail_bomb()


if __name__ == '__main__':
    print('Запуск')
    renew_log()
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s %(name)s %(levelname)s:%(message)s',
                        filename='emails.log')
    logger = logging.getLogger(__name__)
    path = "settings.ini"
    if not os.path.exists(path):
        main.create_config(path)
    config = configparser.ConfigParser()
    config.read(path)
    admin_mail = config.get('Settings', 'admin_mail')
    kachanova_mail = config.get("Settings", "kachanova_mail")
    schedule.every().day.at("12:00").do(__init__)
    schedule.every().day.at("00:00").do(__init__)
    __init__()
    
    try:
        while True:
            schedule.run_pending()
            time.sleep(1)
        
    except Exception:
        logger.exception('Вся программа накрылась')
        main.send_email(admin_mail, 'Программа крякнулась', 'Проверь!', 'emails.log')
        schedule.CancelJob

