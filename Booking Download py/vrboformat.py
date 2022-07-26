from openpyxl import load_workbook

import pyexcel as p
from copy import copy
import os


def Main():
    def delete_col(ws, column_id):
        ws.delete_cols(column_id)

    def nights():
        for i in range(ws.max_row):
            cell = ws.cell(row=i + 2, column=5)
            e = i + 2
            e = str(e)
            # print(e)
            cell.value = "=SUM(D" + e + "-C" + e + ")"



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