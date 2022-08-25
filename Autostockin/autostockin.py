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
    numb = [0,1,2,3]
    receive = ['ครบ / โบ้','ครบ / ปาน','ครบ / มาร์ค','ครบ / ตั้ม']
    idnumber = ['22608','23947','23800','24179']
    readpicerrorfound = pyg.locateCenterOnScreen('asset/checkerror.png')
    ok = pyg.locateOnScreen('asset/ok.png')
except Exception as e:
    messagebox.showerror('Python Error', f'{e}')
    exit()

def custom_comment(comment):
    pyperclip.copy(comment)
    pyg.hotkey('ctrl', 'v')

def pressenter(numberoftimes):
    for i in range(0, numberoftimes):
        pyg.press('enter')
        if i == numberoftimes:
            break

def clickleft(numberoftimes):
    for i in range(0, numberoftimes):
        pyg.leftClick()
        pyg.sleep(1.5)
        if i == numberoftimes:
            break


def defaultbeh():
    def firstStart(start):
        if start == 1:
            pyg.hotkey('alt','k')
            pyg.press('i')
        else:
            pass
    def nextStart(next):
        if next == 1:
            pass
        else:
            pyg.hotkey('alt','k')
            pyg.press('i')

    for i in range(1, worksheet.max_row+1):
        stockoutid = worksheet.cell(row=i, column=1).value
        etc = worksheet.cell(row=i, column=2).value
        customcomment = 'MEM224076 ตัดยอดคืน supplier'
        if stockoutid:
            nextStart(i)
            firstStart(i)
            print('Start Stock In')
            if etc == 'm': 
                pyg.typewrite('7538')
            elif etc == 'โบ้':
                pyg.typewrite('22608')
            elif etc == 'c':
                pyg.typewrite('22608')
            elif etc == 'pairin':
                pyg.typewrite('1815')
            elif etc == 'ปาน':
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
            elif etc == 'pairin':
                pyperclip.copy('ครบ / ไพรินทร์')
                pyg.hotkey('ctrl', 'v')
            elif etc == 'ปาน':
                pyperclip.copy('ครบ / ปาน')
                pyg.hotkey('ctrl', 'v')
            elif etc == 'c':
                custom_comment(customcomment)           
            else:
                pyperclip.copy(receive[r])
                pyg.hotkey('ctrl', 'v')
            ##ready
                pyg.hotkey('alt','f')
                pyg.press('o')


            try:
                if readpicerrorfound:
                    pressenter(1)
                    print('Found Error')
                    continue
                else:
                    print('Error is not Found')
                    pressenter(1)
                    pyg.press('left')
                    pressenter(1)
                    pyg.sleep(3)
                    pressenter(4)
                    continue
            except Exception as e:
                messagebox.showerror('Python Error', f'{e}')
        else: break


try:
    defaultbeh()
except Exception as e:
    messagebox.showerror('Python Error', f'{e}')
    exit()








































