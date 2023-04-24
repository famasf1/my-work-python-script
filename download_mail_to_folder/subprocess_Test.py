import subprocess
import pyautogui as pyg
import pandas as pd


def main():
    #convert into dataframe
    def getdataframe():
            df = pd.read_clipboard(sep="\t")
            print(df)



    ##ID49
    subprocess.Popen("C:\Program Files (x86)\Softbox\ITEC2007_49\ITECStock2007.exe")
    ##ITEC Login Script
    def login(user, pwd):
        pyg.sleep(1)
        def do_your_thing(what_field):
            pyg.click(what_field)
            pyg.write(user)
            pyg.press("enter")
            pyg.write(pwd)
            pyg.press("enter")
        while True:
            pyg.sleep(1)
            user_field = pyg.locateOnScreen(r"D:\Workstuff\my-work-python-script\asset\user_field.png")
            user_field2 = pyg.locateOnScreen(r"D:\Workstuff\my-work-python-script\asset\user_field_is1920.png")
            user_field3 = pyg.locateOnScreen(r"D:\Workstuff\my-work-python-script\asset\your_id_pls.png")
            if user_field:
                do_your_thing(user_field)
                break
            elif user_field2:
                do_your_thing(user_field2)
                break
            elif user_field3:
                do_your_thing(user_field3)
                break
            else:
                pyg.sleep(1)
    login("22608", "22608")
    pyg.hotkey("alt", "k")
    pyg.press("u", 2)
    pyg.press("enter")
    pyg.sleep(2)
    pyg.click(1737,191)
    pyg.press("down")
    pyg.press("enter")
    pyg.sleep(1)
    pyg.press("f12")
    pyg.sleep(1200) #20minute
    pyg.click(596,337)
    pyg.hotkey("ctrl","a")
    pyg.hotkey("ctrl", "c")
    getdataframe()



    ##ID49Insure
    subprocess.Popen("C:\Program Files (x86)\Softbox\ITECInsurance_49\ITECStock2007.exe")
    login("22608", "22608")
    pyg.hotkey("alt", "k")
    pyg.press("u", 2)
    pyg.press("enter")

def mouse():
    pyg.mouseInfo()
    
    


if __name__ in "__main__":
    mouse()
    #main()