import openpyxl
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
import pandas as pd
import pyautogui as pyg
import pyperclip
import xlsxwriter

brand = [403,162]
product_row = [39,412]

def main():
    def goto(x,y):
        pyg.moveTo(x=x,y=y)
        pyg.click()
    i = 1
    while True:
        pyg.sleep(2)
        goto(brand[0],brand[1])
        pyg.click()
        pyg.press('down')
        pyg.press('enter')
        pyg.press('f12')
        pyg.sleep(3)
        goto(product_row[0],product_row[1])
        pyg.hotkey('ctrl','a')
        pyg.hotkey('ctrl','c')
        df = pd.read_clipboard('\t')
        df.to_excel(rf'D:\Workstuff\my-work-python-script\Print_Form_Project\result\{i}.xlsx', engine='xlsxwriter')
        i += 1



if __name__ in '__main__':
    main()
    #pyg.mouseInfo()