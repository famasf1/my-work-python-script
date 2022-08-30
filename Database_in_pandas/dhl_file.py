#import

###############################
# WHAT DAY IS TODAY?
mydayis = 0 #0 = today or 1 = yesterday?
###############################


from imghdr import what
import pandas as pd
from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import askopenfilename
import datetime
today = datetime.datetime.today().strftime("%d-%m-%y")
yesterday = datetime.datetime.today() - datetime.timedelta(days=1)
yesterday = yesterday.strftime("%d-%m-%y")

def read():
    #read
    global sheet
    root = Tk()
    root.excel = filedialog.askopenfilename(title='Open DHL file',filetypes=[('Excel Files', '*.xls'), ('All Files' , '*.*')])
    sheet = pd.read_excel(root.excel)
    remove_word = sheet['CCN'].replace(['PHYIDINSURE','PHYID'],'', regex=True).str.split('-')
    rw_df = pd.DataFrame(remove_word)
    rw_df2 = rw_df[['ID','Branch','Box Num']] = pd.DataFrame(rw_df.CCN.to_list(), index=rw_df.index)
    #combine frame
    frames = [sheet, rw_df2]
    sheet = pd.concat(frames, axis=1)
    whatday(mydayis)

def whatday(whatday):
    if whatday == 0: #today
        name = f'DHL {today}'
    elif whatday == 1: #yesterday
        name = f'DHL{yesterday}'
    print(whatday)
    sheet.to_excel(fr'C:\Users\Comseven\Documents\DHL\Completed\{name}.xlsx',index=False)


if __name__ in '__main__':
    read()