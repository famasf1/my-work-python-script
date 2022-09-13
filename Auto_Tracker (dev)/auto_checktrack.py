from optparse import Option
from selenium import webdriver
import pyperclip
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager
from subprocess import CREATE_NO_WINDOW
from selenium.webdriver.chrome.options import Options

##Tried to practice how to use class in Python by create a nested function lmao
class launch_browser():
     ##################################### ##################################### #####################################
    '''
    This class is for launching selected browser. It has 2 methods.
    1. launch_browser_chrome will launch Google Chrome
    2. launch_browser_firefox will launch Firefox
    '''
     ##################################### ##################################### #####################################
    def launch_browser_chrome(self):
        chrome_service = ChromeService(ChromeDriverManager().install())
        chrome_service.creationflags = CREATE_NO_WINDOW
        chrome_option = Options().add_experimental_option("detach",True)
        driver = webdriver.Chrome(options=chrome_option,service=chrome_service)
        driver.maximize_window()
        return driver
    def launch_browser_firefox(self):
        firefox_service = FirefoxService(GeckoDriverManager().install())
        driver = webdriver.Firefox(service=firefox_service)
        driver.maximize_window()
        return driver


def tracker():
    ##################################### ##################################### #####################################
    '''
    Launch Browser, find element in the web and search it.
    '''
    ##################################### ##################################### #####################################
    global driver
    driver = launch_browser().launch_browser_firefox() ##change firefox to chrome to use chrome instead
    driver.get("https://ecommerceportal.dhl.com/track/")
    with fragile(open(r'D:\Workstuff\my-work-python-script\Auto_Tracker (dev)\retrieve_tracking.txt', 'r+')) as text:
        ## First 50 numbers will be in different page.
        for index, value in enumerate(text):
            if index+1 % 49 != 0:
                driver.find_element(By.XPATH,"//textarea[@id='trackItNowForm:trackItNowSearchBox']").send_keys(value)
                print(index)

        driver.find_element(By.ID,'trackItNowForm:searchSkuBtn').click()
        WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.ID,"trackItNowForm:searchSkuBtn")))
        refid = driver.find_element(By.XPATH,"label[@class='ui-outputlabel ui-widget TrackStatus']").text
            #if index > 0:
            #    WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.ID,"trackItNowForm:searchSkuBtn")))
            #    driver.find_element(By.ID,"trackItNowForm:trackItNowSearchBox").send_keys(value)
            #    driver.find_element(By.ID,"trackItNowForm:searchSkuBtn").click()
            #    break
    print(refid)
    status = driver.find_element(By.ID,'trackItNowForm:j_idt523:0:Status').text
    dateandtime = driver.find_element(By.ID,'trackItNowForm:j_idt523:0:dateandtime').text
    ##trackItNowForm:trackSearchBox_content == sidebar search

def retrievetrackingcode():

    ##################################### ##################################### #####################################
    '''
    This function will fetch all value in txt files, then loop it before sending the result into tracker()
    '''
    ##################################### ##################################### #####################################

    with open(r'D:\Workstuff\my-work-python-script\Auto_Tracker (dev)\retrieve_tracking.txt', 'r+') as text:
        for index, value in enumerate(text):
            for v in value:
                print(f'{index}. : {value}')
            if index % 50 == 0 and index != 0:
                print('WE REACH THE END')
                break




if __name__ in "__main__":
    ##tracker is the main function
    tracker()
    #retrievetrackingcode()