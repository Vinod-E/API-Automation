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


class createCandidate(unittest.TestCase):
    @classmethod
    def setUp(inst):
        # create a new browser session """
        inst.driver = webdriver.Chrome("/home/rajeshwar/Downloads/chromedriver")
        inst.driver.implicitly_wait(30)
        inst.driver.maximize_window()
        # navigate to the application home page
        inst.driver.get("http://amsin.hirepro.in")
        print ('\nEntered URL in browser')
        inst.driver.title
        return inst.driver

    def test_createCandidate(self):
        print ('User reached on project selection screen')
        self.driver.find_element_by_id("ams").click()
        print ('Clicked on AMS project')
        self.driver.find_element_by_name("name").send_keys("at")
        print ('Entered Tenant alias "AT"')
        self.driver.find_element_by_xpath("//*[@class='ui-button-text']").click()
        print ('Clicked on "Next" button to move on next screen')
        time.sleep(3)
        self.LoginName_field = self.driver.find_element_by_id("ctl00_userLogin_txtName")
        self.Password_field = self.driver.find_element_by_id("ctl00_userLogin_txtpssword")
        # enter search keyword and submit
        self.LoginName_field.clear()
        self.LoginName_field.send_keys("admin")
        print ('Entered Login name')
        self.Password_field.clear()
        self.Password_field.send_keys("at@admin")
        abc = self.driver.get_cookies()
        print(abc)
        print ('Entered Password')
        self.driver.find_element_by_id("ctl00_userLogin_lbnEnterButton1").click()
        time.sleep(7)
        print('Clicked to "Login" button')
        self.driver.find_element_by_id("CreateCandidate").click()
        time.sleep(2)
        # ele.send_keys(Keys.ESCAPE)
        self.driver.find_element_by_id("txtName_Candidate").send_keys("Rajeshwar")
        self.driver.find_element_by_id("txtEmail_Candidate").send_keys("rajeshwar.jadhav@hirepro.in")
        self.driver.find_element_by_id("maleRadio_Candidate").click()
        self.driver.find_element_by_id("YesMaritalStatus_Candidate").click()
        self.driver.find_element_by_id("txtContactNo_Candidate").send_keys("1231231231")
        SourceType = self.driver.find_element_by_id('selectSourceType_Candidate')
        for option in SourceType.find_elements_by_tag_name('option'):
            if option.text == "Consultant":
                option.click()  # select() in earlier versions of webdriver
                break
        time.sleep(1)
        self.driver.find_element_by_xpath("//*[@id='trSource_CandidateGeneralInfo']/td[2]/span/span[1]/span/span[2]").click()
        time.sleep(1)
        self.driver.find_element_by_xpath("/html/body/span/span/span[1]/input").send_keys("test")
        self.driver.find_element_by_xpath("/html/body/span/span/span[1]/input").send_keys(Keys.ENTER)
        # self.driver.find_element_by_xpath("//*[@id='select2-selectSource_Candidate-result-mjha-2391']").click()
        time.sleep(1)
        ele = self.driver.find_element_by_id("browse_filedivUploadResumeAttachment_Candidate")
        ele.click()
        time.sleep(1)
        ele1 = self.driver.find_element_by_css_selector("input[type='file']")
        ele1.send_keys("/home/rajeshwar/Downloads/Test_Plan_Template_02.doc")
        self.driver.find_element_by_id("submitDiv").click()
        time.sleep(6)





    @classmethod
    def tearDown(inst):
        # close the browser window
        inst.driver.quit()


if __name__ == '__main__':
    unittest.main()