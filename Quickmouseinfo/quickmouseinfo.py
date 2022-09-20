import pyautogui as pyg
import pyperclip

def one():
    image = None
    while (image == None):
        try:
            image = pyg.locateAllOnScreen(r"D:\Workstuff\my-work-python-script\Return\asset\NOT_NULL.png", grayscale=True, confidence=.9)
        except Exception as e:
            print(e)
            continue
    print(image)
    pyg.mouseInfo()



one()