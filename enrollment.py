# -*- coding: utf-8 -*-
import xlrd
import os
import xlsxwriter
from main import normal_date
from datetime import datetime

try:
    from tqdm import tqdm
except ImportError:
    try:
        os.system('pip3 install tqdm')
        from tqdm import tqdm
    except Exception:
        os.system('pip install tqdm')
        from tqdm import tqdm

try:
    from PIL import ImageFont
except ImportError:
    try:
        os.system('pip3 install pillow')
        from PIL import ImageFont
    except Exception:
        os.system('pip install pillow')
        from PIL import ImageFont

def enrollment(list, output='result.xlsx'):

    def make_full_list(file):
        rb = xlrd.open_workbook(file)
        sheet = rb.sheet_by_index(0)
        list1 = []
        for rownum in tqdm(range(7, sheet.nrows)):
            list1.append(sheet.row_values(rownum)[1:])
        return list1

    def in_this_month(date, check_date):
        result = False

        parts = date.split('.')
        if str(check_date.year) == parts[2] and str(check_date.month).zfill(2) == parts[1]:
            result = True

        return result

    current_datetime = datetime.now().date()

    font = ImageFont.truetype('calibri.ttf', 11)
    widths = [24.71,21.43,40.71,20.14,20.00,19.86,21.71]

    full_list = make_full_list(list)
    result_list = []
    for item in tqdm(full_list):
        if in_this_month(normal_date(item[13]), current_datetime):
            result_list.append([item[0],
                                item[1],
                                item[6],
                                item[7],
                                item[9],
                                item[10],
                                item[12],
                                item[8], ])

    result_list.sort(key = lambda x: x[6])

    workbook = xlsxwriter.Workbook(output)

    worksheet_budget = workbook.add_worksheet('БЮДЖЕТ')

    worksheet_budget.set_column('A:A', widths[0])  #ФИО Обучающегося                   178px
    worksheet_budget.set_column('B:B', widths[1])  #Дата рождения                      155px
    worksheet_budget.set_column('C:C', widths[2])  #Услуга (ДО/ОП)                     290px
    worksheet_budget.set_column('D:D', widths[3])  #Уровень программы/ Квалификация    146px
    worksheet_budget.set_column('E:E', widths[4])  #Направленность                     145px
    worksheet_budget.set_column('F:F', widths[5])  #Группа                             144px
    worksheet_budget.set_column('G:G', widths[6])  #ФИО преподавателя                  157px

    header = workbook.add_format({'bold': True, 'font_size': 11, 'border': True, 'align': 'center', 'valign': 'vcenter'})
    header.set_text_wrap()
    page_header = workbook.add_format({'border': False, 'align': 'left'})
    page_header.set_text_wrap()
    usual = workbook.add_format({'border': True, 'align': 'center', 'valign': 'vcenter'})
    usual.set_text_wrap()

    offbudget = []

    row = 0
    if len(result_list) == 0:
        worksheet_budget.write(row, 0, 'Нет зачисленных')
        row += 1
    else:
        head = ['ФИО Обучающегося', 'Дата рождения', 'Услуга (ДО/ОП)','Уровень программы/ Квалификация', 'Направленность', 'Группа', 'ФИО преподавателя', ]
        i = 0
        row = 0

        worksheet_budget.merge_range(row, i, row, i + 3, '  Приложение №1 к приказу №_/з-б от __.__.____', page_header)  #Заголовок листа
        row += 1

        for i in range(len(head)):  # Заголовок
            worksheet_budget.write(row, i, head[i], header)
        worksheet_budget.set_row(row, 45)
        row += 1

        for i in range(len(result_list)):
            j = 0
            if result_list[i][7] == 'бесплатно':
                worksheet_budget.set_row(row, 15)
                sizes = [];
                for j in range(len(result_list[i]) - 1):
                    data_tmp = None
                    if j == 1:
                        data_tmp = normal_date(result_list[i][j])
                    else:
                        data_tmp = result_list[i][j]
                    sizes.append(font.getlength(str(data_tmp)) / (widths[j] * 7.25))
                    worksheet_budget.write(row, j, data_tmp, usual)
                worksheet_budget.set_row(row, (max(sizes) + 2) * 15)
                row += 1
            elif result_list[i][7] == 'платно':
                offbudget.append(result_list[i])
    row += 1

    worksheet_offbudget = workbook.add_worksheet('ВНЕБЮДЖЕТ')

    worksheet_offbudget.set_column('A:A', widths[0])  #ФИО Обучающегося                   178px
    worksheet_offbudget.set_column('B:B', widths[1])  #Дата рождения                      155px
    worksheet_offbudget.set_column('C:C', widths[2])  #Услуга (ДО/ОП)                     290px
    worksheet_offbudget.set_column('D:D', widths[3])  #Уровень программы/ Квалификация    146px
    worksheet_offbudget.set_column('E:E', widths[4])  #Направленность                     145px
    worksheet_offbudget.set_column('F:F', widths[5])  #Группа                             144px
    worksheet_offbudget.set_column('G:G', widths[6])  #ФИО преподавателя                  157px

    row = 0
    if len(offbudget) == 0:
        worksheet_offbudget.write(row, 0, 'Нет зачисленных')
        row += 1
    else:
        head = ['ФИО Обучающегося', 'Дата рождения', 'Услуга (ДО/ОП)','Уровень программы/ Квалификация', 'Направленность', 'Группа', 'ФИО преподавателя', ]
        i = 0
        row = 0

        worksheet_offbudget.merge_range(row, i, row, i + 3, '  Приложение №1 к приказу №_/з-п от __.__.____', page_header)  #Заголовок листа
        
        row += 1

        for i in range(len(head)):  #Заголовок
            worksheet_offbudget.write(row, i, head[i], header)
        worksheet_offbudget.set_row(row, 45)
        row += 1

        for i in range(len(offbudget)):
            j = 0

            worksheet_offbudget.set_row(row, 15)
            sizes = [];
            for j in range(len(offbudget[i]) - 1):
                data_tmp = None
                if j == 1:
                    data_tmp = normal_date(offbudget[i][j])
                else:
                    data_tmp = offbudget[i][j]
                sizes.append(font.getlength(str(data_tmp)) / (widths[j] * 7.25))
                worksheet_offbudget.write(row, j, data_tmp, usual)
            worksheet_offbudget.set_row(row, (max(sizes) + 2) * 15)
            row += 1
    row += 1

    workbook.close()

    return output

if __name__ == '__main__':
    enrollment('назачисление.xlsx')