arabic_num = {"๙": "+" , "ๅ" : "1" , "/" : "2" , "-" : "3" , "ภ": "4" , "ถ": "5" , "ุ" : "6" , "ึ" : "7" , "ค" : "8" , "ต" : "9" , "จ" : "0", "ข" : "-"}

value_list = []

def calculate(line_):
    for all_data in line_:
        for i in range(0,len(all_data)):
            for key,value in arabic_num.items():
                if all_data[i] in key:
                    value_list.append(value)
    result = ''.join(value_list)
    text.write(f'%s\n' % result)
    value_list.clear()
    

if __name__ in "__main__":
    with open(r"D:\Workstuff\my-work-python-script\thai_text.txt", "r+", encoding='UTF-8') as text:
        lines = text.readlines()
        for index, line in enumerate(lines):
            calculate(line)
            
            

