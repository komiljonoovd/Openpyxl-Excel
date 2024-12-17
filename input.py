import openpyxl
from faker import Faker

fakedata = Faker()

wb = openpyxl.load_workbook(filename='input.xlsx', data_only=True)

sheets = wb.active

# equate sheet to value use it
sheets['A1'] = 'First Name'
sheets['B1'] = 'Last Name'
sheets['C1'] = 'Email'

print(sheets['D1'].value)  # to print data use : .value

# Fill with fake data with Faker
for i in range(2, 32):
    sheets[f'A{i}'] = fakedata.first_name()
    sheets[f'B{i}'] = fakedata.last_name()
    sheets[f'C{i}'] = fakedata.email()

wb.save(filename='input.xlsx')
wb.close()
