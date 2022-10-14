import pyautogui as pyg

def one():
    image = None
    while (image == None):
        try:
            image = pyg.locateOnScreen(rf'D:\Workstuff\my-work-python-script\rotate\asset\ret_error.png', grayscale=True)
            pyg.sleep(2)
        except Exception as e:
            print(e)
            continue
        pyg.moveTo(image)

def two():
    pyg.mouseInfo()

two()