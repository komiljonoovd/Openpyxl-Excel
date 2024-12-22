import openpyxl
from openpyxl.styles import Alignment

wb = openpyxl.load_workbook(filename='input.xlsx')
sheet = wb.active

horizontal = vertical = 'center'

sheet['D2'] = 'Hello'
sheet['E2'] = 'from'
sheet['F2'] = 'OpenPyXL'

sheet['D2'].alignment = Alignment(horizontal=horizontal, vertical=vertical)
sheet['F2'].alignment = Alignment(text_rotation=90)

wb.save(filename='input.xlsx')
wb.close()
