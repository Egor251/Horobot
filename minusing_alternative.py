# -*- coding: utf-8 -*-
import xlrd
import os
import xlsxwriter
from main import normal_date

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

'''parser = argparse.ArgumentParser(description='Videos to images')
parser.add_argument('old', type=str, help='Input dir for videos')
parser.add_argument('new', type=str, help='Output dir for image')
args = parser.parse_args()
print(args.indir)'''


'''def normal_date(date):
    try:
        date = int(date)
    except ValueError:
        date = int(round(float(date)))

    str_tmp = str(xlrd.xldate.xldate_as_datetime(date, "%d.%m.%Y"))[:-9]
    date = f'{str(int(str_tmp[-2:])-1).zfill(2)}.{str_tmp[5:7]}.{str(int(str_tmp[0:4])-4)}'
    return date'''


def minusing_alternative(old, new, output='result.xlsx'):
    def make_list(file):
        rb = xlrd.open_workbook(file)
        sheet = rb.sheet_by_index(0)
        list1 = []
        for rownum in tqdm(range(7, sheet.nrows)):
            if sheet.row_values(rownum)[7][:2].isdigit():
                prog = sheet.row_values(rownum)[7][len(sheet.row_values(rownum)[7].split(".")[0])+2:]  # убираем номер программы
            else:
                prog = sheet.row_values(rownum)[7]

            list1.append(sheet.row_values(rownum)[1]+prog)
        return list1

    def make_full_list(file):
        rb = xlrd.open_workbook(file)
        sheet = rb.sheet_by_index(0)
        list1 = []
        for rownum in tqdm(range(7, sheet.nrows)):
            list1.append(sheet.row_values(rownum)[1:])
        return list1

    font = ImageFont.truetype('calibri.ttf', 11)
    widths = [24.71,21.43,40.71,20.14,20.00,19.86,21.71]

    list_new = make_list(new)
    list_old = make_list(old)
    full_list = make_full_list(old)
    result_list = []
    for item in tqdm(list_old):
        if item not in list_new:
            result_list.append([full_list[list_old.index(item)][0],
                                full_list[list_old.index(item)][1],
                                full_list[list_old.index(item)][6],
                                full_list[list_old.index(item)][7],
                                full_list[list_old.index(item)][9],
                                full_list[list_old.index(item)][10],
                                full_list[list_old.index(item)][12],
                                full_list[list_old.index(item)][8], ])

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
        worksheet_budget.write(row, 0, 'Нет отчисленных')
        row += 1
    else:
        head = ['ФИО Обучающегося', 'Дата рождения', 'Услуга (ДО/ОП)','Уровень программы/ Квалификация', 'Направленность', 'Группа', 'ФИО преподавателя', ]
        i = 0
        row = 0

        worksheet_budget.merge_range(row, i, row, i + 3, '  Приложение №1 к приказу №_/о-б от __.__.____', page_header)  #Заголовок листа
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
        worksheet_offbudget.write(row, 0, 'Нет отчисленных')
        row += 1
    else:
        head = ['ФИО Обучающегося', 'Дата рождения', 'Услуга (ДО/ОП)','Уровень программы/ Квалификация', 'Направленность', 'Группа', 'ФИО преподавателя', ]
        i = 0
        row = 0

        worksheet_offbudget.merge_range(row, i, row, i + 3, '  Приложение №1 к приказу №_/о-п от __.__.____', page_header)  #Заголовок листа
        
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
    minusing_alternative('до.xlsx', 'после.xlsx')
