from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

import datetime


class Browser:
    def __init__(self):
        print('')

    def ui_automation(self):

        # ------------------------------------------------------
        # UI Automation to handle where ever APIs are not present
        # ------------------------------------------------------
        # try:
        web_driver = webdriver.Chrome(
            service=Service(executable_path=ChromeDriverManager().install()),
        )
        web_driver.get("http://www.python.org")

Object = Browser()
Object.ui_automation()