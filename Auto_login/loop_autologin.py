import pyautogui as pyg

user = "pairin"
pwd = "rin45822"

##ITEC Login Script
def login():
    pyg.sleep(1)
    def do_your_thing(what_field):
        pyg.click(what_field)
        pyg.write(user)
        pyg.press("enter")
        pyg.write(pwd)
        pyg.press("enter")
    while True:
        pyg.sleep(1)
        user_field = pyg.locateOnScreen(r".\asset\user_field.png")
        user_field2 = pyg.locateOnScreen(r".\asset\user_field_is1920.png")
        user_field3 = pyg.locateOnScreen(r".\asset\your_id_pls.png")
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