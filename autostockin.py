from tkinter import filedialog, messagebox
from tkinter import *
import pyautogui as pyg
import openpyxl
import pyperclip

def tkinterdefaultbox():
    root = Tk()
    root.excel = filedialog.askopenfilename(initialdir='/Desktop',title='เลือกไฟล์ Excel สำหรับ Stock-Out33', filetypes=(('Excel','*.xlsx'),('All Files','*.*')))
    workbook = openpyxl.load_workbook(root.excel, data_only=True)
    root.withdraw()
    sheet = workbook.sheetnames
    worksheet = workbook[sheet[0]]
    receive = 'ครบ / โบ้'

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
    tkinterdefaultbox()
    for i in range(1, worksheet.max_row+1):
        stockoutid = worksheet.cell(row=i, column=1).value
        if stockoutid:
            pyg.moveTo(235,146)
            pyg.leftClick()
            pyg.sleep(3.5)
            pyg.typewrite('22608')
            pressenter(2)
            pyg.typewrite(str(stockoutid)) #stockout
            pressenter(2)
            pyperclip.copy(receive)
            pyg.hotkey('ctrl', 'v')
            pyg.moveTo(68,833)
            clickleft(1)
            pressenter(1)
            pyg.sleep(3)
            pyg.moveTo(871,502)
            clickleft(6)
            pyg.sleep(3)


try:
    defaultbeh()
except Exception as e:
    messagebox.showerror('Python Error', f'{e}')








































