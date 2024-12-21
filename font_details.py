import openpyxl
from openpyxl.styles import Font

wb = openpyxl.load_workbook(filename='input.xlsx')
sheet = wb.active

cell = sheet['D1']
cell.font = Font(size=12)
cell.value = "Hello"

cell = sheet['E1']
cell.font = Font(name="Arial", size=14)
cell.value = "from"

cell = sheet['F1']
cell.font = Font(name="Tahoma", size=12, color="00FF0000")
cell.value = "OpenPyXL"

wb.save(filename='input.xlsx')
