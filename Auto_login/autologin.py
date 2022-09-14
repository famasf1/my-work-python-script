from time import sleep
import pyautogui as pyg
import os
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from datetime import date
from subprocess import CREATE_NO_WINDOW
##Customize setting
#path_to_49ins = "C:\Program Files (x86)\Softbox\ITECINSURE_49\ITECStock2007.exe"
#my_own_path = "C:\Program Files (x86)\Softbox\ITECInsurance_49\ITECStock2007.exe"
#p_god_path = "C:\Program Files (x86)\Softbox\ITECINSURE_49\ITECStock2007.exe"
p_dew_path = "C:\Program Files (x86)\Softbox\ITECInsurance\ITECStock2007.exe"
#p_aek_path = "C:\Program Files (x86)\Softbox\ITECInsurance_49\ITECStock2007.exe"
p_mark_path = ""

### ITEC
#p_god = 24021
#p_dew = 22073
p_aek = 23267
user = "22073"
pwd = "22073"

### TECH_TRADE
username = "service"
password = "1234567"

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

### ITEC Launch 
#walk into set path and find ITEC execution file
def launch_49insure(amouth=0):
    while amouth < 2:
        cwd = os.path.abspath(r'C:\Program Files (x86)')
        soft_box = os.path.join(cwd, "Softbox")
        for roots, dirs, files in os.walk(soft_box):
            for filename in files:
                if os.path.join(roots, filename) == str(p_dew_path):
                    os.startfile(p_dew_path)

        login()
        amouth += 1
                

##### TECH TRADE
### A workaround for unable to put driver out of function.

def launch_browser():
    chrome_service = ChromeService(ChromeDriverManager().install())
    chrome_service.creationflags = CREATE_NO_WINDOW
    chrome_option = Options().add_experimental_option("detach",True)
    driver = webdriver.Chrome(options=chrome_option,service=chrome_service)
    driver.maximize_window()
    #พี่ดิว - driver.get("https://docs.google.com/spreadsheets/d/1ePJmJceR37NGA1oxSyen-VTnVnmFsCwNwEGjN9Wi8Jc/edit#gid=0")
    #พี่เอก - driver.get("https://docs.google.com/spreadsheets/d/1I72TCwCa6VMQccQSy5IpoSY-bZlk19m5SEmQlYYFsY8/edit#gid=0")
    #พี่ก็อต - driver.get("https://docs.google.com/spreadsheets/d/1E2FJwhY6WyyME4r44C2XIuGkgAaBG2dxxbpvBoBJ8MM/edit#gid=0")
    #พี่มาร์ค - driver.get("https://docs.google.com/spreadsheets/d/1lCRsY9KdTEDOMMYF4076gz5rVgv-y7tq0H1enm2eHuI/edit#gid=0")
    driver.get("https://docs.google.com/spreadsheets/d/1lCRsY9KdTEDOMMYF4076gz5rVgv-y7tq0H1enm2eHuI/edit#gid=0")
    driver.switch_to.new_window('tab')
    driver.get("https://www.tradein-com7.com/TI/login.aspx")
    return driver

def auto_login_techtrade(user=username, password=password):
    global driver
    driver = launch_browser()
    ## 
    driver.find_element(By.ID,"txtUsername").send_keys(user)
    driver.find_element(By.ID,"txtPassword").send_keys(password)
    driver.find_element(By.ID,"btnSignin").click()
    sleep(1.5)
    driver.get("https://www.tradein-com7.com/Backoffice/Branch_history/branch_history_list.aspx")
    WebDriverWait(driver,5).until(EC.invisibility_of_element_located((By.CLASS_NAME,"modal-backdrop fade show")))
    click_element = WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.ID,"FromPlaceHolder_ddlTradein")))
    ActionChains(driver).move_to_element(click_element).click().perform()
    driver.find_element(By.ID,"FromPlaceHolder_txtStartDate").clear()
    epoch_year = date.today().year
    first_day_of_year = date(epoch_year,1,1).strftime("%d/%m/%Y")
    driver.find_element(By.ID,"FromPlaceHolder_txtStartDate").send_keys(first_day_of_year)
    driver.find_element(By.ID,"FromPlaceHolder_txtEndDate").clear()
    driver.find_element(By.ID,"FromPlaceHolder_txtEndDate").send_keys(date.today().strftime("%d/%m/%Y"))
    select_prt = driver.find_element(By.ID,"FromPlaceHolder_ddlTradein")
    select_prt_object = Select(select_prt)
    select_prt_object.select_by_value("3")
    driver.find_element(By.ID,"btnSearch").click()


#### DEFAULT
if __name__ in "__main__":
    auto_login_techtrade()
    launch_49insure()
    

