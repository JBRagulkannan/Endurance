# import datetime
# from tkinter import Button

# from openpyxl import Workbook
#
# def save_excel(Dt,intvol,intsale):
#     wb = Workbook()
#     sheet = wb.active
#     # int_vol = 123.09, 848, 98497
#     sheet['A2'] = "Date"
#     sheet['B2'] = "Initial PPU"
#     sheet['C2'] = "Final PPU"
#     sheet['D2'] = "Initial Vol"
#     sheet['E2'] = "Final Vol"
#     sheet['F2'] = "Vol diff"
#     sheet['G2'] = "Initial sale"
#     sheet['H2'] = "Final sale"
#     sheet['I2'] = "Sale diff"
#     sheet['J2'] = "Initial Bill NO"
#     sheet['K2'] = "Final Bill NO"
#     sheet['L2'] = "Initial ERR"
#     sheet['M2'] = "Final ERR"
#     sheet['N2'] = "Initial CRT_ERR"
#     sheet['O2'] = "Final CRT_ERR"
#     sheet['P2'] = "Initial PPU"
#     sheet['Q2'] = "Initial ERR"
#     sheet['R2'] = "Initial CRT_ERR"
#     # sheet['A1'] = int_vol
#     # sheet.append([Dt,intvol,intsale])
#     sheet.append({'A': Dt, 'D': intvol, 'G': intsale})
#     wb.save('TOT_Save.xlsx')
#
# save_excel(datetime.datetime.now(), 12.34, 34.56)

import datetime
import time

# from two_wire_client_tkinter import save_excel

# today = str(datetime.datetime.now())
#
# today = today.split(" ")[0]
# time.sleep(2)
# tom = str(datetime.datetime.now())
# tom = tom.split(" ")[0]
# print(today)
# print(tom)
# if today != tom:
#     print("do work here")
# else:
#     raise("Same date. Exiting")
#
# import openpyxl
#
# def get_init_val_excel():
#     wb = openpyxl.load_workbook('TOT_Save.xlsx')
#     sheet = wb.active
#     last_row = sheet.max_row
#     print(last_row)
#     init_dateTime = sheet["A" + str(last_row)].value
#     init_ppu = sheet["C" + str(last_row)].value
#     init_vol = sheet["E" + str(last_row)].value
#     init_sale = sheet["H" + str(last_row)].value
#     init_bill = sheet["K" + str(last_row)].value
#     init_err = sheet["M" + str(last_row)].value
#     init_cr_err = sheet["O" + str(last_row)].value
#     return init_dateTime, init_ppu, init_vol, init_sale, init_bill, init_err, init_cr_err
#
# init_dateTime, init_ppu, init_vol, init_sale, init_bill, init_err, init_cr_err = get_init_val_excel()
# print(init_dateTime, init_ppu, init_vol, init_sale, init_bill, init_err, init_cr_err)



#Import the required Libraries
#

from tkinter import *

root = Tk()
root.geometry("300x200")

w = Label(root, text='GeeksForGeeks', font="50")
w.pack()

Checkbutton1 = IntVar()
Checkbutton2 = IntVar()
Checkbutton3 = IntVar()

Button1 = Checkbutton(root, text="Tutorial",
                      variable=Checkbutton1,
                      onvalue=1,
                      offvalue=0,
                      height=2,
                      width=10)

Button2 = Checkbutton(root, text="Student",
                      variable=Checkbutton2,
                      onvalue=1,
                      offvalue=0,
                      height=2,
                      width=10)

Button3 = Checkbutton(root, text="Courses",
                      variable=Checkbutton3,
                      onvalue=1,
                      offvalue=0,
                      height=2,
                      width=10)

Button1.pack()
Button2.pack()
Button3.pack()

mainloop()