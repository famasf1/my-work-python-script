##in-dev

import gspread
import openpyxl

gc = gspread.service_account()

wks = gc.open('รับสินค้าเข้า').get_worksheet_by_id(0)

date = wks.col_values(1)
green_Box = wks.col_values(3)
out_id = wks.col_values(4)
branch = wks.col_values(5)


for id,greenbox_num in enumerate(green_Box):
    if greenbox_num != '':
        for id2,outid in enumerate(out_id):
            if greenbox_num != '':
                dictionary = dict({greenbox_num : outid})
                print(dictionary)
