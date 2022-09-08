import glob
import pandas as pd

def main():
    path = rf"D:\Workstuff\my-work-python-script\Print_Form_Project\result"

    file = glob.glob(path + rf'\*.xlsx')

    excl_list = []

    for i in file:
        excl_list.append(pd.read_excel(i))

    excl_merge = pd.concat(excl_list, ignore_index=True)
    excl_merge.to_excel('all_product.xlsx', index=False, engine="xlsxwriter")

def test():
    pass

if __name__ in '__main__':
    main()
    #test()

