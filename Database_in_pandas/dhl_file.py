#import

###############################
# WHAT DAY IS TODAY?
mydayis = 0 #0 = today or 1 = yesterday or 2 = yesterday of yesterday
###############################

import pstats
import token
import pandas as pd
from tkinter import *
from tkinter import filedialog
import datetime
import openpyxl as pyxl
import bitlyshortener
from token_bitly import list_token

today = datetime.datetime.today().strftime("%d-%m-%y")
yesterday = datetime.datetime.today() - datetime.timedelta(days=1)
yesterday = yesterday.strftime("%d-%m-%y")
dayafteryesterday = datetime.datetime.today() - datetime.timedelta(days=2)
dayafteryesterday = dayafteryesterday.strftime("%d-%m-%y")

def read():

    '''
    Load Excel with Pandas, find and replace PHYIDINSURE | PHYID with empty whitespace. Then split out the data into array with '-'
    once that's done, combine frame and send all the data to 'whatday' function to process into Excel file. 
    '''

    #read
    global sheet
    
    root = Tk()
    root.excel = filedialog.askopenfilename(title='Open DHL file',filetypes=[('Excel Files', '*.xls'), ('All Files' , '*.*')]) #Load Excel
    sheet = pd.read_excel(root.excel)
    remove_word = sheet['CCN'].replace(['PHYIDINSURE','PHYID'],'', regex=True).str.split('-')
    rw_df = pd.DataFrame(remove_word)
    rw_df2 = rw_df[['ID','Branch','Box Num']] = pd.DataFrame(rw_df.CCN.to_list(), index=rw_df.index)

    #combine frame
    frames = [sheet, rw_df2]
    sheet = pd.concat(frames, axis=1)
    insure_only()
    whatday(mydayis)

def whatday(whatday):

    '''
    Take data from 'read' function. Check whatday variable setting before export Pandas dataframe into Excel.
    '''

    global name
    if whatday == 0: #today
        name = f'DHL {today}'
    elif whatday == 1: #yesterday
        name = f'DHL {yesterday}'
    elif whatday == 2:
        name = f'DHL {dayafteryesterday}'
    print(whatday)
    sheet.to_excel(fr'C:\Users\Comseven\Documents\DHL\Completed\{name}.xlsx',index=False)
    sheet_insure_only.to_excel(fr'C:\Users\Comseven\Documents\DHL\Completed\{name}_INSURE_ONLY.xlsx',index=False)
    bitly_api_activate()

def insure_only():
    '''
    After read() function finished. Take data from loaded Excel sheet and flitered out only value that match 'PHYIDINSURE', '33-1' using regex.
    then replace all the 'PHYIDINSURE' into blank, split '-' and return as Pandas Dataframe before sent it to whatday() function to process.
    '''
    global sheet_insure_only
    CCN_Insure = sheet[sheet["CCN"].str.contains(r"(^PHYIDINSURE.*)-(33-1)$",regex=True)]
    remove_word = CCN_Insure['CCN'].replace(['PHYIDINSURE'],'', regex=True).str.split('-')
    df_remove = pd.DataFrame(remove_word)
    df_remove_as_table = df_remove[['ID','Branch','Box Num']] = pd.DataFrame(df_remove.CCN.to_list(), index=df_remove.index)
    frames = [CCN_Insure, df_remove_as_table]
    sheet_insure_only = pd.concat(frames, axis=1)\
    
def bitly_api_activate():

    '''
    Get data from selected row where URL link is located. Then convert them into shortlinks.
    '''

    token_pool = ['2659b7dde7f007b1fe5cbf7784d78905800e3066','691e75c823d2856e59a28b26500bc7492654b85b']
    bitly_connect = bitlyshortener.Shortener(tokens=token_pool, max_cache_size=256)

    load_insure_wb = pyxl.load_workbook(fr'C:\Users\Comseven\Documents\DHL\Completed\{name}_INSURE_ONLY.xlsx', data_only=True)
    load_insure_wb_sh = load_insure_wb.sheetnames
    active_Sheet = load_insure_wb[load_insure_wb_sh[0]]
    for row in range(2, active_Sheet.max_row+1):

        long_link_list = []
        long_link = active_Sheet.cell(row=row, column=11).value
        long_link_list.append(long_link)
    
        short_link = bitly_connect.shorten_urls(long_link_list)
        active_Sheet.cell(row=row, column=15).value = str(short_link[0])
        load_insure_wb.save(fr'C:\Users\Comseven\Documents\DHL\Completed\{name}_INSURE_ONLY_1.xlsx')
    
def bitly_api_activate_once():

    '''
    Get data from selected row where URL link is located. Then convert them into shortlinks.
    ***THIS IS MEANT TO BE USE ONCE***
    '''

    bitly_connect = bitlyshortener.Shortener(tokens=list_token, max_cache_size=256)

    load_insure_wb = pyxl.load_workbook(fr'C:\Users\Comseven\Documents\DHL\Completed\DHL_8+9+10_INSURE_ONLY_1.xlsx', data_only=True)
    load_insure_wb_sh = load_insure_wb.sheetnames
    active_Sheet = load_insure_wb[load_insure_wb_sh[0]]
    for row in range(2, active_Sheet.max_row+1):

        long_link_list = []
        long_link = active_Sheet.cell(row=row, column=11).value
        long_link_list.append(long_link)
    
        short_link = bitly_connect.shorten_urls(long_link_list)
        active_Sheet.cell(row=row, column=15).value = str(short_link[0])
        load_insure_wb.save(fr'C:\Users\Comseven\Documents\DHL\Completed\DHL_8+9+10_INSURE_ONLY_2.xlsx')

if __name__ in '__main__':
    read()
    #bitly_api_activate_once()