import random

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import numbers

wb = openpyxl.load_workbook('numeric.xlsx')
sheet = wb.active

for i in range(1, 11):
    sheet['A' + str(i)] = random.randint(10000, 100000)
    sheet['A' + str(i)].number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1

    sheet['B' + str(i)] = random.randint(10000, 100000)
    sheet['B' + str(i)].number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED2

    sheet['C' + str(i)] = random.randint(10000, 100000)
    sheet['C' + str(i)].number_format = numbers.FORMAT_NUMBER_00

    sheet['D' + str(i)] = random.randint(10000, 100000)
    sheet['D' + str(i)].number_format = numbers.FORMAT_NUMBER  # Default number format

wb.save(filename='numeric.xlsx')
