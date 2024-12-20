import openpyxl

wb = openpyxl.load_workbook(filename='input.xlsx')

sheets = wb.active

# to print all data by sheets
for i in range(1, sheets.max_row + 1):
    print(sheets[f'A{i}'].value, end=' ')
    print(sheets[f'B{i}'].value, end=' ')
    print(sheets[f'C{i}'].value)

wb.save(filename='input.xlsx')
wb.close()
