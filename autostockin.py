from tkinter import filedialog, messagebox
from tkinter import *
import pyautogui as pyg
import openpyxl
import pyperclip
import random

try:
    root = Tk()
    root.excel = filedialog.askopenfilename(initialdir='/Desktop',title='เลือกไฟล์ Excel สำหรับ Stock-Out33', filetypes=(('Excel','*.xlsx'),('All Files','*.*')))
    workbook = openpyxl.load_workbook(root.excel, data_only=True)
    root.withdraw()
    sheet = workbook.sheetnames
    worksheet = workbook[sheet[0]]
    numb = [0,1,2]
    receive = ['ครบ / ดิว','ครบ / ก็อต','ครบ / เอก']
    idnumber = ['22073','23017','23267']
except Exception as e:
    messagebox.showerror('Python Error', f'{e}')
    exit()

def pressenter(numberoftimes):
    for i in range(0, numberoftimes):
        pyg.press('enter')
        if i == numberoftimes:
            break

def clickleft(numberoftimes):
    for i in range(0, numberoftimes):
        pyg.leftClick()
        pyg.sleep(0.5)
        if i == numberoftimes:
            break
        

def defaultbeh():
    for i in range(1, worksheet.max_row+1):
        stockoutid = worksheet.cell(row=i, column=1).value
        etc = worksheet.cell(row=i, column=2).value
        istrueor = worksheet.cell(row=i,column=3).value
        if stockoutid:
            pyg.moveTo(235,146)
            pyg.leftClick()
            pyg.sleep(3.5)
            if etc == 'm': 
                pyg.typewrite('7538')
            elif etc == 'โย้':
                pyg.typewrite('22608')
            elif etc == 'j':
                pyg.typewrite('23030')
            elif etc == 'pan':
                pyg.typewrite('22929')
            else: 
                r = random.choice(numb)
                pyg.typewrite(idnumber[r])
            pressenter(2)
            pyg.typewrite(str(stockoutid)) #stockout
            pressenter(2)
            if etc == 'm': # for my queen
                pyperclip.copy('ครบ / เอ็ม')
                pyg.hotkey('ctrl', 'v')
            elif etc == 'โบ้':
                pyperclip.copy('ครบ / โบ้')
                pyg.hotkey('ctrl', 'v')
            elif etc == 'j':
                pyperclip.copy('ครบ / จุ๊ย')
                pyg.hotkey('ctrl', 'v')
            elif etc == 'pan':
                pyperclip.copy('ครบ / ปาน')
                pyg.hotkey('ctrl', 'v')           
            else:
                pyperclip.copy(receive[r])
                pyg.hotkey('ctrl', 'v')
            pyg.moveTo(68,833)
            clickleft(1)
            pressenter(1)
            pyg.press('left')
            pressenter(4)
            pyg.sleep(2)
        else: 
            continue


try:
    defaultbeh()
except Exception as e:
    messagebox.showerror('Python Error', f'{e}')
    exit()








































