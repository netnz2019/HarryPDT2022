from openpyxl import load_workbook

import pyexcel as p
from copy import copy
import os


def delete_col(ws, column_id):
    ws.delete_cols(column_id)


def nights():
    for i in range(ws.max_row):
        cell = ws.cell(row=i + 2, column=5)
        e = i + 2
        e = str(e)
        # print(e)
        cell.value = "=SUM(D" + e + "-C" + e + ")"


def tobold(list):
    for i in list:
        print(i)
        cell = ws.cell(row=1, column=i)
        cell.font = cell.font.copy(bold=True)





try:
    os.remove(r'C:\Users\harry\Desktop\Rstatements\moo.xlsx')
except:
    pass


wb = load_workbook(r'C:\Users\harry\Desktop\Rstatements\Vrbo.xlsx')
ws = wb.active


delete_col(ws, 13)
delete_col(ws, 13)
delete_col(ws, 1)
delete_col(ws, 1)
delete_col(ws, 1)
delete_col(ws, 2)
delete_col(ws, 3)
ws.delete_rows(1)

wb.save(r'C:\Users\harry\Desktop\Rstatements\moo.xlsx')