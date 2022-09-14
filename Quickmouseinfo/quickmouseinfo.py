import pyautogui as pyg
import pyperclip

def one():
    image = None
    while (image == None):
        try:
            image = pyg.locateCenterOnScreen(r"D:\Workstuff\my-work-python-script\asset\foundstockbill.png", confidence=.8)
        except Exception as e:
            print(e)
            continue
    print(image)


one()