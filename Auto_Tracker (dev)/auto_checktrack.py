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
import pandastable
import pandas as pd
import auto_checktrack
import openpyxl

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

def gather_Data():
    ##################################### ##################################### #####################################
    '''
    This function will gather data from the clipboard, Flitered everything out then pass along all the values to calculate in 'tracker' function
    '''
    ##################################### ##################################### #####################################    

    global result

    df_data = pd.DataFrame(pd.read_clipboard(sep='\t'))

    header = list(df_data.columns)
    donotdrop = ['Doc Date','Branch (Name)','Branch To (Name)','Booking ID']
    new_data = df_data[donotdrop]
    new_data['Booking ID'] = new_data['Booking ID'].str.replace('Booking-DHL ID : ','')
    new = new_data.dropna()['Booking ID']   #.to_csv('retrieve_tracking.txt', header=None, index=None)
    new = pd.DataFrame(new)
    new['Booking ID'] = new.apply(lambda x: x['Booking ID'].split(' , '), axis=1)
    result = new.explode('Booking ID').to_clipboard(index=False)
    tracker()


def tracker():
    ##################################### ##################################### #####################################
    '''
    Launch Browser, find element in the web and search it. Then call 'retrievetrackingcode' to fetch all the data into pandastable
    '''
    ##################################### ##################################### #####################################
    global driver
    driver = launch_browser().launch_browser_firefox() ##change firefox to chrome to use chrome instead
    driver.get("https://ecommerceportal.dhl.com/track/")

    def retrievetrackingcode():

        ##################################### ##################################### #####################################
        '''
        This function will fetch all value in txt files, then loop it before sending the result into tracker()
        '''
        ##################################### ##################################### #####################################

        with open(r'D:\Workstuff\my-work-python-script\Auto_Tracker (dev)\retrieve_tracking.txt', 'r+') as text: #fetch all value
            for index, value in enumerate(text): #iterate through all of them first
                print(f'Total Line :{len(text.readlines())}')
                print(f'Index = {index}')
                driver.find_element(By.XPATH,"//textarea[@id='trackItNowForm:trackItNowSearchBox']").send_keys(value)
                if index % 49 == 0 and index != 0:
                    driver.find_element(By.ID,'trackItNowForm:searchSkuBtn').click()  
                    # ^ = start-with
                    # * = contains
                    #Is this regex?
                elif index % len(text.readlines()) == 0 and index != 0:
                    driver.find_element(By.ID,'trackItNowForm:searchSkuBtn').click()
                WebDriverWait(driver,6).until(EC.visibility_of_element_located((By.XPATH, "//label[contains(@id,'trackItNowForm') and(contains(@class,'TrackingNumber'))]"))).text
                for i in range(0,50):
                    tracknumber = driver.find_element(By.CSS_SELECTOR,f"[id^='trackItNowForm'][id*=':{i}:'][class*='TrackingNumber']").text
                    status = driver.find_element(By.CSS_SELECTOR,f"[id^='trackItNowForm'][id*=':{i}:'][class*='TrackingStatus']").text
                    timeanddate = driver.find_element(By.CSS_SELECTOR,f"[id^='trackItNowForm'][id*=':{i}:'][class*='TrackTimeAndDate']").text
                    continue
    retrievetrackingcode()

def test_room():

    ##################################### ##################################### #####################################
    '''
    This function is purely for testing. Delete after production-ready
    '''
    ##################################### ##################################### #####################################
    with open(r'D:\Workstuff\my-work-python-script\Auto_Tracker (dev)\retrieve_tracking.txt', 'r+') as text:
        for index, value in enumerate(text):
            print(index)
            print(value)
            if index % 49 == 0 and index != 0:
                print('Break')


if __name__ in "__main__":
    ##tracker is the main function
    #gather_Data()
    tracker()
    #retrievetrackingcode()

    #test_room()