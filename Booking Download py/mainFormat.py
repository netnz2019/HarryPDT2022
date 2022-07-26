
import openpyxl
import pandas as pd


def main():
    wb3 = openpyxl.load_workbook(r"C:\Users\harry\Desktop\Rstatements\output.xlsx")
    ws3 = wb3.active
    ws3.delete_cols(1)
    wb3.save(r"C:\Users\harry\Desktop\Rstatements\output.xlsx")
    def save():
        wb3 = openpyxl.load_workbook(r"C:\Users\harry\Desktop\Rstatements\output.xlsx")
        ws3 = wb3.active
        wb4 = openpyxl.load_workbook(r"C:\Users\harry\Desktop\Rstatements\input.xlsx")
        ws4 = wb4.active
        for i in ws3["A2:A" + str(ws3.max_row)]:
            for c in i:
                print(c.value)
                ws4.cell(row=c.row, column=1).value = c.value

        for i in ws3["G2:G" + str(ws3.max_row)]:
            for v in i:
                ws4.cell(row=v.row, column=2).value = v.value
        wb4.save(r"C:\Users\harry\Desktop\Rstatements\input.xlsx")

        wb3.save(r"C:\Users\harry\Desktop\Rstatements\output.xlsx")
    save()

    print("we going")
    wb1 = openpyxl.load_workbook(r'C:\Users\harry\Desktop\Rstatements\moo.xlsx')
    wb2 = openpyxl.load_workbook(r"C:\Users\harry\Desktop\Rstatements\formbook.xlsx")

    ws1 = wb1.active
    ws2 = wb2.active

    ws2max = ws2.max_row
    maxcol = str(ws1.max_row)
    print(maxcol)

    x = ws2max
    y = 0
    p = 0
    d = 1
    e = False
    yy=0
    xx=1

    for r in ws1['A1:G' + maxcol]:

        for c in r:
            y += 1


            ws2.cell(row=x, column=y).value = c.value

            if e == False:

                if p / 6 == 1:
                    e = True
                    x += 1
                    p = 1
                    y = 0
            else:
                if p / 8 == 1:
                    e = True
                    x += 1
                    p = 1
                    y = 0

            p += 1
            d += 1

            print(bool(x - 90 / 7))

    wb2.save(r"C:\Users\harry\Desktop\Rstatements\formbook2.xlsx")

    xl = pd.ExcelFile(r"C:\Users\harry\Desktop\Rstatements\formbook2.xlsx")
    df = xl.parse("Sheet1")
    df = df.sort_values("Check-in")

    writer = pd.ExcelWriter(r"C:\Users\harry\Desktop\Rstatements\output.xlsx")
    df.to_excel(writer, sheet_name='Sheet1')
    writer.save()

    wb3 = openpyxl.load_workbook(r"C:\Users\harry\Desktop\Rstatements\output.xlsx")
    ws3 = wb3.active
    ws3.delete_cols(1)

    def nights():
        for i in range(2, ws3.max_row - 1):
            cell = ws3.cell(row=i, column=5)
            e = i
            print(i)
            e = str(e)
            # print(e)
            cell.value = "=SUM(D" + e + "-C" + e + ")"

    nights()

    for i in range(ws3.max_row):
        cell = ws3.cell(row=i + 1, column=2)
        nextcell = ws3.cell(row=i + 2, column=2)
        now = cell.value
        next = nextcell.value
        if next == now:
            ws3.delete_rows(i + 2)
    guest = {}
    wb3.save(r"C:\Users\harry\Desktop\Rstatements\output2.xlsx")

    wb4 = openpyxl.load_workbook(r"C:\Users\harry\Desktop\Rstatements\input.xlsx")
    ws4 = wb4.active

    for cell in ws3["B1:B"+str(ws3.max_row)]:
        for v in cell:
            name=str(v.value)
            #print(name)

            for va in ws4["A1:A"+str(ws4.max_row)]:
                for val in va:
                    #print("Val " + str(val.value))
                    #print(str(val.value) + ": " + name)
                    #print(name)
                    #print(v.value)
                    if val.value == name:
                        print("Hi")

                        if ws4.cell(row=val.row, column=2).value != None:
                            print("Gay")


                            ws3.cell(row=v.row, column=8).value = ws4.cell(row=val.row, column=2).value












    wb3.save(r"C:\Users\harry\Desktop\Rstatements\output.xlsx")

