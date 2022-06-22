import openpyxl.cell
import openpyxl

from openpyxl import load_workbook

import pyexcel as p
from copy import copy
import os

def main():
    p.save_book_as(file_name=r'C:\Users\harry\Desktop\Rstatements\rstatement.xls',
                   dest_file_name=r'C:\Users\harry\Desktop\Rstatements\booking_com.xlsx')

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

    def removenzd():
        for i in range(ws.max_row - 2):
            # print(i)
            cell = ws.cell(row=i + 2, column=8)
            val = cell.value

            if val == None:
                pass
            else:
                space = val.index(" ")
                # print(space)

                nzdremoved = val[:space]
                print(nzdremoved)
                cell.value = float(nzdremoved)

    try:
        os.remove(r'C:\Users\harry\Desktop\Rstatements\formbook.xlsx')
    except:
        pass

    wb = load_workbook(r'C:\Users\harry\Desktop\Rstatements\booking_com.xlsx')
    ws = wb.active
    delete_col(ws, 1)
    delete_col(ws, 2)
    delete_col(ws, 5)
    delete_col(ws, 6)
    delete_col(ws, 8)
    delete_col(ws, 9)

    ws.insert_cols(1)

    max = ws.max_row
    print(max)
    coll = "E" + str(max)
    ws.move_range("E1:" + coll, cols=-4)
    nights()
    delete_col(ws, 6)
    for i in range(6):
        delete_col(ws, 9)
    du = ws.cell(row=1, column=5)
    du.value = "Nights"

    nz = ws.cell(row=2, column=8)
    print(nz.value)
    removenzd()
    tobold({1, 2, 3, 4, 5, 6, 7, 8})
    wb.save(r"C:\Users\harry\Desktop\Rstatements\formbook.xlsx")

