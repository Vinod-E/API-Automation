import unittest
import time

import datetime
from sqlite3 import Date

import xlrd
import xlwt
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from urlparse import urlparse
from selenium.common.exceptions import NoSuchElementException, NoAlertPresentException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys


class createSource(unittest.TestCase):
    @classmethod
    def setUp(inst):
        # create a new browser session """
        inst.driver = webdriver.Chrome("/home/rajeshwar/Downloads/chromedriver")
        inst.driver.implicitly_wait(30)
        inst.driver.maximize_window()
        # navigate to the application home page
        inst.driver.get("http://10.0.3.41/rpo")
        print ('\nEntered URL in browser')
        inst.driver.title
        return inst.driver

    def test_createSource(self):
        self.driver.find_element_by_name("clientEmail").send_keys("0003")
        self.driver.find_element_by_name("new-password").send_keys("Ujjwal@123")
        self.driver.find_element_by_xpath("//div[2]/div/div[1]/div/div/div[4]/div[1]/div[1]/div/button").click()
        time.sleep(3)
        self.driver.find_element_by_xpath("//li[5]/a").click()
        time.sleep(2)
        self.driver.find_element_by_xpath("//p[2]").click()
        time.sleep(2)

        SourceType = self.driver.find_element_by_xpath("//div[2]/div/md-input-container/md-select")
        SourceType.send_keys("Online")

        SourceName = self.driver.find_element_by_name("name")
        SourceName.send_keys("Script_Source_One")

        Email = self.driver.find_element_by_xpath("//div[3]/md-input-container/input")
        Email.send_keys("rajeshwar.jadhav@hirepro.in")

        Location = self.driver.find_element_by_xpath("//div[4]/md-input-container/md-select")
        Location.send_keys("Ahmedabad")

        ValidFrom = self.driver.find_element_by_xpath("//div/input")
        ValidFrom.send_keys("1/1/2017")

        ValidTo = self.driver.find_element_by_xpath("//div[6]/md-datepicker/div/input")
        ValidTo.send_keys("26/1/2017")

        Description = self.driver.find_element_by_xpath("//div[2]/div[2]/div/div[2]/div[3]")
        Description.send_keys("Hello Source")

        Save = self.driver.find_element_by_xpath("(//button[@type='button'])[6]")
        Save.click()



        time.sleep(10)



    @classmethod
    def tearDown(inst):
        # close the browser window
        inst.driver.quit()


if __name__ == '__main__':
    unittest.main()