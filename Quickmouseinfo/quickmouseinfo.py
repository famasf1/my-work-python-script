import pyautogui as pyg
import pyperclip

def one():
    pyg.sleep(1)
    if pyg.moveTo(pyg.locateCenterOnScreen(r"D:\Workstuff\my-work-python-script\asset\ret_error.png", grayscale=True)):
        pyg.moveTo(pyg.locateCenterOnScreen(r"D:\Workstuff\my-work-python-script\asset\ret_error.png", grayscale=True))
    else:
        print(0)
    #63-147

one()