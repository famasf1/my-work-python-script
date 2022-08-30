import collections
import ctypes
import openpyxl
from tkinter import filedialog, messagebox
from tkinter import *
import os
import sys
import datetime as dt
import win32com.client
from barcode import Code128
from barcode.writer import ImageWriter
from PIL import Image

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
    col1_tolist = []
    col2_tolist = []
    col3_tolist = []

    for val in range(1, all_data.max_row+1):
        col1 = all_data.cell(row=val, column=1).value
        col1 = f'No.{col1}'
        col2 = all_data.cell(row=val, column=2).value
        col3 = all_data.cell(row=val, column=3).value
        col1_tolist.append(col1)
        col2_tolist.append(col2)
        template_front1()
        template_front2()
        template_front3()
        template_front_final()



    def template_front1(): #top left
        if len(col1_tolist) % 4 == 1:
            sticker_front['A1'].value = col1_tolist[0]
            sticker_front['A2'].value = col2_tolist[0]
            with open(rf"D:\Workstuff\my-work-python-script\Print Barcode\result\{col3}.png".replace("/00","-00"), "wb") as files:
                Code128(col3, writer=ImageWriter()).write(files)
                img1 = openpyxl.drawing.image.Image(files)
                img1.height = 76
                img1.width = 163
                sticker_front.add_image(img1, 'A3')

    def template_front2(): #top right
        if len(col1_tolist) % 4 == 2:
            sticker_front['C1'].value = col1_tolist[1]
            sticker_front['C2'].value = col2_tolist[1]
            with open(rf"D:\Workstuff\my-work-python-script\Print Barcode\result\{col3}.png".replace("/00","-00"), "wb") as files:
                Code128(col3, writer=ImageWriter()).write(files)
                img1 = openpyxl.drawing.image.Image(files)
                img1.height = 76
                img1.width = 163
                sticker_front.add_image(img1, 'C3')

    def template_front3(): #bottom left
        if len(col1_tolist) % 4 == 3:
            sticker_front['A4'].value = col1_tolist[2]
            sticker_front['A5'].value = col2_tolist[2]
            with open(rf"D:\Workstuff\my-work-python-script\Print Barcode\result\{col3}.png".replace("/00","-00"), "wb") as files:
                Code128(col3, writer=ImageWriter()).write(files)
                img1 = openpyxl.drawing.image.Image(files)
                img1.height = 76
                img1.width = 163
                sticker_front.add_image(img1, 'A6')

    def template_front_final(): #bottom right
        if len(col1_tolist) % 4 == 0:
            sticker_front['C4'].value = col1_tolist[3]
            sticker_front['C5'].value = col2_tolist[3]
            with open(rf"D:\Workstuff\my-work-python-script\Print Barcode\result\{col3}.png".replace("/00","-00"), "wb") as files:
                Code128(col3, writer=ImageWriter()).write(files)
                img1 = openpyxl.drawing.image.Image(files)
                img1.height = 76
                img1.width = 163
                sticker_front.add_image(img1, 'C6')
            col1_tolist.clear()
            col2_tolist.clear()
            col3_tolist.clear()
            worksheet.save(root.excel)
            dispatcher = win32com.client.Dispatch('Excel.Application')
            dispatcher.visible = False
            wb = dispatcher.Workbooks.Open(str(root.excel))
            getsheet = wb.Worksheets([0])
            getsheet.PrintOut()
            wb.Close(True)

        #sticker_front['A4'].value = store_id[data]
        #pass ##TODO : add function


### TEST ROOM
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

### Extra
def MBox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0,text,title,style)

def barcode_generator():
    for i in range(1, all_data.max_row+1):
        values = all_data.cell(row=i, column=3).value
        with open(rf"D:\Workstuff\my-work-python-script\Print Barcode\result\{values}.png".replace("/00","-00"), "wb") as files:
            Code128(values, writer=ImageWriter()).write(files)

if __name__ in '__main__':
    main()
    #test_room()