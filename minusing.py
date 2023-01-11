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


def minusing(old, new, output='result.xlsx'):
    def make_list(file):
        rb = xlrd.open_workbook(file)
        sheet = rb.sheet_by_index(0)
        list1 = []
        for rownum in tqdm(range(7, sheet.nrows)):
            list1.append(sheet.row_values(rownum)[1:6]+sheet.row_values(rownum)[7:10]+sheet.row_values(rownum)[12:13])
        return list1


    def make_full_list(file):
        rb = xlrd.open_workbook(file)
        sheet = rb.sheet_by_index(0)
        list1 = []
        for rownum in tqdm(range(7, sheet.nrows)):
            list1.append(sheet.row_values(rownum)[1:])
        return list1


    list_new = make_list(new)
    list_old = make_list(old)
    full_list = make_full_list(old)
    result_list = []
    for item in tqdm(list_old):
        if item not in list_new:
            result_list.append([full_list[list_old.index(item)][0], full_list[list_old.index(item)][1],
                                full_list[list_old.index(item)][6], full_list[list_old.index(item)][7],
                                full_list[list_old.index(item)][9], full_list[list_old.index(item)][10],
                                full_list[list_old.index(item)][12], full_list[list_old.index(item)][8], ])

    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet('Дети')

    worksheet.set_column('A:A', 30)
    worksheet.set_column('B:B', 10)
    worksheet.set_column('C:C', 30)
    worksheet.set_column('D:E', 20)
    worksheet.set_column('F:F', 40)
    worksheet.set_column('G:G', 40)
    worksheet.set_column('H:H', 20)
    worksheet.set_column('I:I', 15)
    worksheet.set_column('J:J', 30)
    worksheet.set_column('K:K', 20)
    worksheet.set_column('L:L', 25)
    worksheet.set_column('N:N', 15)
    worksheet.set_column('O:Q', 15)

    header = workbook.add_format({'bold': True, 'font_size': 11, 'border': True})
    header.set_text_wrap()
    usual = workbook.add_format({'border': True})
    usual.set_text_wrap()

    row = 0
    if len(result_list) == 0:
        worksheet.write(row, 0, 'Нет отчисленных')
        row += 1
    else:
        head = ['ФИО Обучающегося', 'Дата рождения', 'Услуга (ДО/ОП)','Уровень программы/ Квалификация', 'Направленность', 'Группа', 'ФИО преподавателя',
                 'Бюджет']
        i = 0
        row = 0
        for i in range(len(head)):  # Заголовок
            worksheet.write(row, i, head[i], header)
        row += 1

        for i in range(len(result_list)):
            j = 0
            for j in range(len(result_list[i])):
                data_tmp = None
                if j  == 1:
                    data_tmp = normal_date(result_list[i][j])
                else:
                    data_tmp = result_list[i][j]
                #data_tmp = result_list[i][j]
                worksheet.write(row, j, data_tmp, usual)
            row += 1
    row += 1
    workbook.close()
    return output

if __name__ == '__main__':
    minusing('до.xlsx', 'после.xlsx')