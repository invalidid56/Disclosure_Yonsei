import openpyxl as excel
import re


def parse(fn):
    wb = excel.load_workbook(filename=fn)
    ws = wb.active

    mail = [cell.value for cell in ws['E'] if cell.value]


    f = open('email.txt', 'w')

    for m in mail:
        f.write(m+',')
    f.close()
