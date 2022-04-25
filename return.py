import pyautogui as pyg
import openpyxl
import pyperclip

def press_enter(number):
    for n in range(0,number):
        pyg.press('enter')

directoryhere = r"C:\Users\jambo\Desktop\Trade In\my work python script\returndataready.xlsx"
data = openpyxl.load_workbook(directoryhere, data_only=True)
datasheet = data.sheetnames
datasheet1 = data[datasheet[0]]
datasheet2 = data[datasheet[1]]
settingsheet = data[datasheet[2]]

supcode = settingsheet['B1'].value
supname = settingsheet['B2'].value
docdatedata = settingsheet['A3'].value
datesent = settingsheet['B3'].value
com7rts = settingsheet['B4'].value

def start_here():
    pyg.write('22608')
    press_enter(2)
    pyg.write(supcode)
    press_enter(2)
    pyg.write(com7rts)
    press_enter(1)
    pyg.press('Down')
    pyg.moveTo(166,139)
    pyg.leftClick()
    press_enter(1)
    pyperclip.copy(supname)
    pyg.hotkey('ctrl', 'v')
    pyg.write(f" | {docdatedata} : {datesent}")
    pyg.moveTo(124,233)
    pyg.leftClick()
    for i in range(2 ,datasheet1.max_row+1): #skip row 1
        product_Code = datasheet1.cell(row=i,column=2).value
        number = datasheet1.cell(row=i,column=5).value
        billtype = datasheet1.cell(row=i,column=23).value
        if billtype == 18:
            pyg.write(str(product_Code))
            press_enter(2)
            pyg.sleep(0.56)
            error = pyg.locateCenterOnScreen('ret_error.png')
            if error:
                press_enter(1)
                pyg.press('Up')
                pyg.press('Down')
                pyg.press('Down')
    pyg.hotkey('ctrl','a')
    pyg.hotkey('ctrl','c')

def stilltest():
    pyg.press('Up')
    pyg.press('Right')
    pyg.hotkey('ctrl', 'Up')
    for i in range(1, datasheet2.max_row+1):
        numberitem = datasheet2.cell(row=i, column=17).value
        pyg.typewrite(str(numberitem))
        press_enter(1)
        pyg.sleep(0.56)

pyg.sleep(1)


start_here()

#stilltest()