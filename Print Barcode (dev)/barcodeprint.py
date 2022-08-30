import ctypes
from unicodedata import name
import openpyxl
from tkinter import filedialog, messagebox
from tkinter import *
import os
import sys
import datetime as dt
import win32com.client
from barcode import Code128
from barcode.writer import ImageWriter

### MAIN
def main():
    global all_data, test_sheet
    root = Tk()
    root.excel = filedialog.askopenfilename(title='เลือกไฟล์ Excel', filetypes=(("Excel", "*.xlsx"),('All Files','*.*')))
    worksheet = openpyxl.load_workbook(root.excel, data_only=True)
    sheet = worksheet.sheetnames
    sticker_front = worksheet[sheet[0]]
    sticker_side = worksheet[sheet[1]]
    all_data = worksheet[sheet[2]]
    test_sheet = worksheet[sheet[3]]

    for data in range(1,all_data.max_row+1):
        numberorder = all_data.cell(row=data, column=1)
        numberorder_return = f'NO.{numberorder}'
        #barcode_generator()
        store_id = all_data.cell(row=data, column=2)
        if store_id:
            sticker_front['A1'].value = numberorder_return
            sticker_front['A2'].value = store_id

        #sticker_front['A4'].value = store_id[data]
        #pass ##TODO : add function
    

def barcode_generator():
    for i in range(1, all_data.max_row+1):
        values = all_data.cell(row=i, column=3).value
        with open(rf"D:\Workstuff\my-work-python-script\Print Barcode\result\{values}.png".replace("/00","-00"), "wb") as files:
            Code128(values, writer=ImageWriter()).write(files)



def test_room():
    global all_data, test_sheet
    root = Tk()
    root.excel = filedialog.askopenfilename(title='เลือกไฟล์ Excel', filetypes=(("Excel", "*.xlsx"),('All Files','*.*')))
    worksheet = openpyxl.load_workbook(root.excel, data_only=True)
    sheet = worksheet.sheetnames
    sticker_front = worksheet[sheet[0]]
    sticker_side = worksheet[sheet[1]]
    all_data = worksheet[sheet[2]]
    test_sheet = worksheet[sheet[3]]
    for i in range(0,4):
        if i != 4:
            for j in range(1,all_data.max_row+1):
                values = all_data.cell(row=j, column=3).value
                value_list = [].append(values)
                test_sheet['A1'].value = value_list[i]
                i += 1
                test_sheet['A2'].value = value_list[i]
                i += 1
                test_sheet['A3'].value = value_list[i]
                i += 1
                test_sheet['A4'].value = value_list[i]
                i += 1
                print(value_list)
                if i == 4:
                    i == 0
        worksheet.save('test.xlsx')
            




### Extra
def MBox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0,text,title,style)

if __name__ in '__main__':
    #main()
    test_room()