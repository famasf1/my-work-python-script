import pyautogui as pyg
import pyperclip

def one():
    image = None
    while (image == None):
        try:
            image = pyg.locateOnScreen(r"D:\Workstuff\my-work-python-script\Return\asset\NOT_NULL.png", confidence=.7, grayscale=True)
        except Exception as e:
            print(e)
            continue
    pyg.moveTo(image)

one()