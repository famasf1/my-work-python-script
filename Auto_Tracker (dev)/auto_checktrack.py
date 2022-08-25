from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from subprocess import CREATE_NO_WINDOW
from selenium.webdriver.chrome.options import Options

def launch_browser():
    chrome_service = ChromeService(ChromeDriverManager().install())
    chrome_service.creationflags = CREATE_NO_WINDOW
    chrome_option = Options().add_experimental_option("detach",True)
    driver = webdriver.Chrome(options=chrome_option,service=chrome_service)
    driver.maximize_window()
    return driver

def tracker():
    global driver
    driver = launch_browser()
    driver.get("https://ecommerceportal.dhl.com/track/")
    driver.find_element(By.XPATH,"//textarea[@id='trackItNowForm:trackItNowSearchBox']").send_keys("test")

if __name__ in "__main__":
    tracker()
