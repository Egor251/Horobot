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

def enrollment(list, output='result.xlsx'):

    workbook = xlsxwriter.Workbook(output)

    worksheet_budget = workbook.add_worksheet('БЮДЖЕТ')

    workbook.close()

    return output

if __name__ == '__main__':
    enrollment('назачисление.xlsx')