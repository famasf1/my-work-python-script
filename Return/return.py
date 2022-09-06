from tkinter import filedialog, messagebox
from venv import create
from numpy import place
import pyautogui as pyg
import openpyxl
import pyperclip
import tkinter as tk       

### create label above entry
def createlabel(text1,placex,placey):
    label = tk.Label(text=text1)
    label.place(x=placex,y=placey)

### Create Windows Interface for automated program without
### Changing Code constantly
root = tk.Tk()
root.title("Return Automation")
root.geometry("580x270")

root.rowconfigure(0,minsize=800, weight=1)
root.columnconfigure(1,minsize=800, weight=1)

hello = tk.Label(text="Hello!").pack()
get_supcode = createlabel("Supplier Code",25,70)
sup_code = tk.Entry(master=root)
sup_code.place(x=25,y=90)
get_supname = createlabel("Supplier Name",160,70)
sup_name = tk.Entry(master=root)
sup_name.place(x=160,y=90)
get_doc_Date = createlabel("Doc Date",290,70)
doc_Date = tk.Entry(master=root)
doc_Date.place(x=290,y=90)
get_rts = createlabel("เลข Com 7 RTS",160,120)
getRTS = tk.Entry(master=root)
getRTS.place(x=160,y=140)
get_numberofrow = createlabel("เลขแถว Excel ที่ต้องการให้เริ่ม",420,70)
getnumRow = tk.Entry(master=root)
getnumRow.place(x=420,y=90)

def clear_all_entry():
    sup_code.delete(0,'end')
    sup_name.delete(0,'end')
    doc_Date.delete(0,'end')
    getRTS.delete(0,'end')

### shorten autopress enter function
def press_enter(number):
    for n in range(0,number):
        pyg.press('enter')
### this is where the code launch

def readexcelAgain():
    openpyxl.load_workbook(readData.directoryhere, data_only=True)

def readData():
    try:
        root.state('iconic')
        directoryhere = filedialog.askopenfilename(title="เลือกไฟล์ Excel ที่มีข้อมูล", filetypes=(("Excel","*.xlsx"),('All Files','*.*')))
        root.state('normal')
    except Exception as e:
        messagebox.showerror(title="Error!",message=f"{e}")
        root.state('normal')
    tk.Label(text=f"ไฟล์โหลดเรียบร้อยแล้วที่ {directoryhere}'").place(x=25,y=45)
    data = openpyxl.load_workbook(directoryhere, data_only=True)
    datasheet = data.sheetnames
    readData.datasheet1 = data[datasheet[0]] #VAT default
    readData.datasheet2 = data[datasheet[1]] #Return
    readData.settingsheet = data[datasheet[2]] #Setting
    readData.datasheet3 = data[datasheet[3]] #NOVAT Check
    readData.data66 = data[datasheet[4]] #Not Stock From 18
    ################## Read Excel ##########
    readData.supcode = sup_code.get()
    readData.supname = sup_name.get()
    readData.docdatedata = doc_Date.get()
    readData.com7rts = getRTS.get()
    readData.getnumRow = int(getnumRow.get())

date_sent = tk.Button(master=root, command=readexcelAgain, text="Reinitialized Excel")
date_sent.place(x=25,y=140)
root.state('normal')


## VatBOT start grinding
def Vat_start_here():
    root.state('iconic')
    pyg.sleep(3)
### Vat bot (มี Vat)
    def VATbot_Start():
        pyg.sleep(2)
        for i in range(readData.getnumRow,readData.datasheet1.max_row+1): #skip row 1
            product_Code = readData.datasheet1.cell(row=i,column=2).value
            number = readData.datasheet1.cell(row=i,column=5).value
            billtype = readData.datasheet1.cell(row=i,column=23).value
            serial = readData.datasheet1.cell(row=i, column=4).value
            if billtype == "18":
                pyg.sleep(0.50)
                pyg.write(str(product_Code))
                pyg.press('Down')
                pyg.press('Down')
                pyg.sleep(1)
                if pyg.locateOnScreen('nothing_error.png', confidence=.9): 
                    pyg.press('Esc')
                    press_enter(1)
                    pyg.press('Down')
                    continue
                pyg.moveTo(360,455)
                pyg.leftClick()
                press_enter(1)
                pyg.sleep(1)
                if pyg.locateOnScreen('ret_error.png', grayscale=True):
                    print('error')
                    press_enter(1)
                    pyg.press('Up')
                    pyg.press('Down')
                    pyg.press('Down')
                elif pyg.locateCenterOnScreen('ret_error2.png', grayscale=True):
                    pyg.press('Esc')
                    pyg.press('Esc')
                    pyg.press('Esc')
                    pyg.press('Down')
                    pyg.press('Down')
        root.state('normal')   
    pyg.write('22608')
    press_enter(2)
    pyg.write(readData.supcode)
    press_enter(2)
    pyg.write(readData.com7rts)
    press_enter(1)
    pyg.press('Down')
    pyg.moveTo(166,139)
    pyg.leftClick()
    press_enter(1)
    pyg.sleep(0.5)
    pyperclip.copy(readData.supname)
    pyg.sleep(0.5)
    pyg.write(f"{pyg.hotkey('ctrl','v')} | {readData.docdatedata}")
    pyg.moveTo(124,233)
    pyg.leftClick()
    VATbot_Start()
    pyg.press('Up')
    pyg.hotkey('ctrl','a')
    pyg.hotkey('ctrl','c')
    root.state('normal')


### No VAT bot (ไม่มีภาษี VAT)
def NOVAT_start_here():
    root.state('iconic')
    pyg.sleep(3)
    def bot_Start():
        pyg.sleep(2)
        for i in range(readData.getnumRow,readData.datasheet3.max_row+1): #skip row 1
            product_Code = readData.datasheet3.cell(row=i,column=2).value
            number = readData.datasheet3.cell(row=i,column=5).value
            billtype = readData.datasheet3.cell(row=i,column=23).value
            if billtype == 18:
                pyg.sleep(0.50)
                pyg.write(str(product_Code))
                pyg.press('Down')
                pyg.press('Down')
                pyg.sleep(0.8)
                press_enter(1)
                pyg.sleep(1)
                pyg.locateOnScreen('ret_error.png', grayscale=True)
                pyg.locateCenterOnScreen('ret_error2.png', grayscale=True)
                if pyg.locateOnScreen('ret_error.png', grayscale=True):
                    print('error')
                    press_enter(1)
                    pyg.press('Up')
                    pyg.press('Down')
                    pyg.press('Down')
                elif pyg.locateCenterOnScreen('ret_error2.png', grayscale=True):
                    pyg.press('Esc')
                    pyg.press('Esc')
                    pyg.press('Esc')
                    pyg.press('Down')
                    pyg.press('Down')    
    pyg.write('22608')
    press_enter(2)
    pyg.write(readData.supcode)
    press_enter(2)
    pyg.write(readData.com7rts)
    press_enter(1)
    pyg.moveTo(166,139)
    pyg.leftClick()
    press_enter(1)
    pyg.sleep(0.5)
    pyperclip.copy(readData.supname)
    pyg.sleep(0.5)
    pyg.write(f"{pyg.hotkey('ctrl','v')} | {readData.docdatedata}")
    pyg.moveTo(124,233)
    pyg.leftClick()
    bot_Start()
    pyg.press('Up')
    pyg.hotkey('ctrl','a')
    pyg.hotkey('ctrl','c')
    root.state('normal')


## Put in number of items
def number_Input():
    root.state('iconic')
    pyg.sleep(3)
    pyg.press('Up')
    pyg.press('Right')
    pyg.hotkey('ctrl', 'Up')
    pyg.sleep(2)
    for i in range(readData.getnumRow, readData.datasheet2.max_row+1):
        numberitem = readData.datasheet2.cell(row=i, column=31).value #column 31
        pyg.typewrite(str(numberitem))
        press_enter(1)
        pyg.sleep(0.56)
        if numberitem == None:
            break
    root.state('normal')

# Stock out to 73
def stock_to73():
    root.state('iconic')
    pyg.sleep(3)
    pyg.write('22608')
    press_enter(2)
    pyg.write('73')
    press_enter(2)
    pyg.write(readData.com7rts)
    press_enter(2)
    pyperclip.copy(readData.supname)
    pyg.write(f"{pyg.hotkey('ctrl','v')} | {readData.supcode} | {readData.docdatedata} ")
    pyg.moveTo(231,216)
    pyg.leftClick()
    for i in range(readData.getnumRow, readData.data66.max_row+1):
        product_Code = readData.data66.cell(row=i,column=2).value
        number_Item = readData.data66.cell(row=i, column=5).value
        serial_Item = readData.data66.cell(row=i, column=4).value
        null_list = ['null', 'NULL']
        for id, val in enumerate(null_list):
            if serial_Item != val:
                pyg.write(str(product_Code))
                pyg.press('Right')
                pyg.press('Right')
                pyg.write(str(number_Item))
                pyg.press('Left')
                pyg.press('Left')
                pyg.press('Left')            
                press_enter(1)
    root.state('normal')

def stock_to73_Noserial():
    root.state('iconic')
    pyg.sleep(3)
    pyg.write('22608')
    press_enter(2)
    pyg.write('73')
    press_enter(2)
    pyg.write(readData.com7rts)
    press_enter(2)
    pyperclip.copy(readData.supname)
    pyg.write(f"{pyg.hotkey('ctrl','v')} | {readData.supcode} | {readData.docdatedata} ")
    pyg.moveTo(231,216)
    pyg.leftClick()
    for i in range(readData.getnumRow, readData.data66.max_row+1):
        product_Code = readData.data66.cell(row=i,column=1).value
        number_Item = readData.data66.cell(row=i, column=3).value
        if product_Code:
            pyg.write(str(product_Code))
            pyg.press('Right')
            pyg.press('Right')
            pyg.write(str(number_Item))
            pyg.press('Left')
            pyg.press('Left')
            pyg.press('Left')            
            press_enter(1)
    root.state('normal')

def create_button_tkinter(text1,command,placex,placey):
    tk.Button(text=text1, command=command).place(x=placex,y=placey)

def restart_Bot():
    pyg.sleep(2)
    for i in range(readData.getnumRow,readData.datasheet1.max_row+1): #skip row 1
        product_Code = readData.datasheet1.cell(row=i,column=2).value
        number = readData.datasheet1.cell(row=i,column=5).value
        billtype = readData.datasheet1.cell(row=i,column=23).value
        if billtype == "18":
            pyg.sleep(0.50)
            pyg.write(str(product_Code))
            pyg.press('Down')
            pyg.press('Down')
            pyg.sleep(1)
            if pyg.locateOnScreen('nothing_error.png', confidence=.9): 
                pyg.press('Esc')
                press_enter(1)
                pyg.press('Down')
                continue
            pyg.moveTo(360,455)
            pyg.leftClick()
            press_enter(1)
            pyg.sleep(1)
            if pyg.locateOnScreen('ret_error.png', grayscale=True):
                print('error')
                press_enter(1)
                pyg.press('Up')
                pyg.press('Down')
                pyg.press('Down')
            elif pyg.locateCenterOnScreen('ret_error2.png', grayscale=True):
                pyg.press('Esc')
                pyg.press('Esc')
                pyg.press('Esc')
                pyg.press('Down')
                pyg.press('Down')


greeting = create_button_tkinter("Browse",readData,250,20)
vatstart = create_button_tkinter("เริ่ม (มี VAT)",Vat_start_here,25,190)
novat_start = create_button_tkinter("เริ่ม (ไม่มี VAT)",NOVAT_start_here,110,190)
number_times = create_button_tkinter("ใส่จำนวนที่คืน",number_Input,205,190)
stock_out_73 = create_button_tkinter("โอนบิล 66 ไป 73",stock_to73,300,190)
stock_out_73_noserialize = create_button_tkinter("โอนบิล 66 ไป 73 (ไม่มี Serial)",stock_to73_Noserial,300,150)
clear_all = create_button_tkinter("เคลียร์ช่อง",clear_all_entry,405,190)
renew = create_button_tkinter("Custom Row",restart_Bot,475,190)

if __name__ == "__main__":
    root.mainloop()
    
        