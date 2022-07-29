#Formats the incoming Files from vrbo.com to the universal formate so it can be
#combined with the files coming from booking.com

from openpyxl import load_workbook
import pyexcel as p
from copy import copy
import os


def Main():
    def delete_col(ws, column_id):
        ws.delete_cols(column_id)

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