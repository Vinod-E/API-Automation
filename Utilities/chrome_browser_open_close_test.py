import datetime
from selenium import webdriver
from selenium.common import exceptions


class ChromeOpen:
    def __init__(self):
        self.driver = webdriver.Chrome('/home/vinodkumar/PythonProjects/API Automation/Utilities/chromedriver')

    def ui_automation(self):
        # ------------------------------------------------------
        # UI Automation to handel where ever APIs are not present
        # ------------------------------------------------------
        try:
            print("Run started at:: " + str(datetime.datetime.now()))
            print("Environment setup has been Done")
            print("----------------------------------------------------------")
            self.driver.implicitly_wait(10)
            self.driver.maximize_window()
            self.driver.get('www.google.com')

            print("----------------------------------------------------------")
            print("Run completed at:: " + str(datetime.datetime.now()))
            print("Chrome environment Destroyed")
            self.driver.close()
        except exceptions as error:
            print(error)


Object = ChromeOpen()
Object.ui_automation()
