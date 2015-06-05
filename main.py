from openpyxl import load_workbook
wb = load_workbook(filename = 'test.xlsx')
sheet1 = wb['Sheet1']
last = 'A1'
for row in sheet1.rows:
    for col in row:
        print col.value
