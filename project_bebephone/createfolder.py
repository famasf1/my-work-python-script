import os, openpyxl

wb = openpyxl.load_workbook(r"C:\Users\Comseven\Downloads\33 สาขา.xlsx")
ws = wb.sheetnames
sheet = wb[ws[0]]

for row in range(3, sheet.max_row+1):
    foldername = sheet.cell(row=row, column=2).value
    os.chdir(r"D:\Workstuff\my-work-python-script\project_bebephone\PRINT")
    os.mkdir(str(foldername))
