from netrc import NetrcParseError
from tkinter import filedialog, messagebox
from tkinter import *
from tracemalloc import start
import pyautogui as pyg
import openpyxl
import pyperclip
import random
############### CLASS ################

### List all employee ###
class Employeelist:
    #เขียนครบ ตามด้วยไอดีพนักงาน
    def __init__(self,index, name, staffid=str):
        self.name = name
        self.staffid = staffid
        self.index = index
    #ขานชื่อ - เรียกชื่อจริง
    #ขานเลขที่ - เรียกเลขที่
    def ขานชื่อ(self):
        return self.name

    def ขานเลขพนักงาน(self):
        return self.staffid

    def get_index(self):
        return self.index

########################################

##### SETTING

def start_script():
    try:
        global workbook, worksheet,readpicerrorfound, ok, employeelist_index, employeelist_name, employeelist_id, checking
        global จิรายุทธ, วรัญญู, มรกต, วุฒิภัทร
        ###################################### LIST EMPLOYEE #####################################
        จิรายุทธ = Employeelist(0,'ครบ / โบ้', '22608') 
        วรัญญู = Employeelist(1,'ครบ / ตั้ม', '24179')
        มรกต = Employeelist(2,'ครบ / ปาน', '23947')
        วุฒิภัทร = Employeelist(3,'ครบ / มาร์ค','23800')
        employeelist_index = [จิรายุทธ.index, วรัญญู.index, มรกต.index, วุฒิภัทร.index]
        ##########################################################################################

        root = Tk()
        root.excel = filedialog.askopenfilename(initialdir='/Desktop',title='เลือกไฟล์ Excel สำหรับ Stock-Out33', filetypes=(('Excel','*.xlsx'),('All Files','*.*')))
        workbook = openpyxl.load_workbook(root.excel, data_only=True)
        root.withdraw()
        sheet = workbook.sheetnames
        worksheet = workbook[sheet[0]] 
        readpicerrorfound = pyg.locateCenterOnScreen('asset/checkerror.png')
        ok = pyg.locateOnScreen('asset/ok.png')

    except Exception as e:
        messagebox.showerror('Python Error', f'{e}')
        exit()

############################
############################
############################
###### MAIN FUNCTION #######

def main(): #short for default behavior
    ### First time starting
    start_script()

## WTF?
## if i remove it, will it break?

    def firstStart(start):
        if start == 1:
            pyg.hotkey('alt','k')
            pyg.press('i')
        else:
            pass
    ### Next

    def nextStart(next):
        if next == 1:
            pass
        else:
            pyg.hotkey('alt','k')
            pyg.press('i')

    def start_in():
        for i in range(1, worksheet.max_row+1):
            stockoutid = worksheet.cell(row=i, column=1).value
            if stockoutid:
                nextStart(i)
                firstStart(i)
                print('Start Stock In')

                #pick staff id
                r = random.choice(employeelist_index)
                match r:
                    case 0:
                        pyg.typewrite(จิรายุทธ.staffid)
                    case 1:
                        pyg.typewrite(วรัญญู.staffid)
                    case 2:
                        pyg.typewrite(มรกต.staffid)
                    case 3:
                        pyg.typewrite(วุฒิภัทร.staffid)
                pressenter(2)
                pyg.typewrite(str(stockoutid)) #stockout
                pressenter(2)

                #name
                match r:
                    case 0:
                        pyperclip.copy(จิรายุทธ.name)
                        pyg.hotkey('ctrl','v')
                    case 1:
                        pyperclip.copy(วรัญญู.name)
                        pyg.hotkey('ctrl','v')
                    case 2:
                        pyperclip.copy(มรกต.name)
                        pyg.hotkey('ctrl','v')
                    case 3:
                        pyperclip.copy(วุฒิภัทร.name)
                        pyg.hotkey('ctrl','v')
                ##ready
                ##now check if value exist
                pyg.hotkey('alt','f')
                pyg.press('o')
                pyg.sleep(5)
                try:
                    if pyg.locateCenterOnScreen(r"D:\Workstuff\my-work-python-script\asset\ret_error.png", grayscale=True):
                        pressenter(1)
                        print('Found Error')
                        pyg.press('esc')
                        continue

                    else:
                        print('Error is not Found')
                        pressenter(1)
                        pyg.press('left')
                        pressenter(1)
                        
                        pressenter(4)
                        continue
                except Exception as e:
                    messagebox.showerror('Python Error', f'{e}')
            else: break

    start_in()
####################################################################
######### Extra Function that doesn't involve with any of the above

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

def test():
    จิรายุทธ = Employeelist(0,'ครบ / โบ้', 22608) 
    วรัญญู = Employeelist(1,'ครบ / ตั้ม', 24179)
    มรกต = Employeelist(2,'ครบ / ปาน', 23947)
    วุฒิภัทร = Employeelist(3,'ครบ / มาร์ค',23800)
    employeelist_index = [จิรายุทธ.index, วรัญญู.index, มรกต.index, วุฒิภัทร.index]
    r = random.choice(employeelist_index)
    print(r)

try:
    #test()
    main()
except Exception as e:
    messagebox.showerror('Python Error', f'{e}')
    exit()








































