

from openpyxl import load_workbook

compiled_wb = load_workbook(filename = r'C:\Users\harry\Desktop\Rstatements\formbook.xlsx')
compiled_ws = compiled_wb['Sheet1']

for i in range(1, 30):
    wb = load_workbook(filename = r'C:\Users\harry\Desktop\Rstatements\moo.xlsx'.format(i))
    ws = wb['Sheet1']
    compiled_ws.append() # ignore row 0

compiled_wb.save(r'C:\Users\harry\Desktop\Rstatements\formbook.xlsx')

