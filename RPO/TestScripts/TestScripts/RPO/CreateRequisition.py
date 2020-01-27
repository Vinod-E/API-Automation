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
from TestScripts.Config.AllConstants import *
from TestScripts.Config.AllInputDataFilePath import file_Path
from TestScripts.RPO_PageObjects.CreateRequisition_PgObj import createRequisition_PgObj

class createRequisition(unittest.TestCase):
    @classmethod
    def setUp(inst):
        # create a new browser session """
        inst.driver = webdriver.Chrome("/home/rajeshwar/Downloads/chromedriver")
        inst.driver.implicitly_wait(30)
        inst.driver.maximize_window()
        # navigate to the application home page
        inst.driver.get("http://10.0.3.41/rpo")
        inst.driver.title
        return inst.driver

    def test_createRequisition(self):
        self.driver.find_element_by_name("clientEmail").send_keys(CONSTANT.INTERNAL_AMS_LOGIN_NAME)
        self.driver.find_element_by_name("new-password").send_keys(CONSTANT.INTERNAL_AMS_LOGIN_PASSWORD)
        self.driver.find_element_by_xpath("//div[2]/div/div[1]/div/div/div[4]/div[1]/div[1]/div/button").click()
        time.sleep(5)
        self.driver.find_element_by_xpath("//li[3]/a").click()
        time.sleep(2)
        print("clicked on Requisition Tab")
        wb = xlrd.open_workbook(file_Path.internal_AMS_File_Path() + ".xls")
        sheetname = wb.sheet_names()  # Read for XLS Sheet names
        sh1 = wb.sheet_by_index(0)
        # print(sheetname)
        i = 1
        while (i < sh1.nrows):
            rownum = (i)
            rows = sh1.row_values(rownum)
            time.sleep(2)
            self.driver.find_element_by_xpath("//p").click()
            time.sleep(2)
            __createRequisition_PgElement = createRequisition_PgObj.createRequisition_PgElements(self.driver)
            __createRequisition_PgElement["Customer"].send_keys(rows[1])
            __createRequisition_PgElement["Job_Code"].send_keys(rows[2])
            __Req_Title = __createRequisition_PgElement["Req_Title"].text
            __createRequisition_PgElement["Openings"].clear()
            __createRequisition_PgElement["Openings"].send_keys(str(int(rows[3])))
            __createRequisition_PgElement["Location"].send_keys(rows[4])
            __createRequisition_PgElement["Req_Type"].send_keys(rows[5])
            __createRequisition_PgElement["Experience_Range"].send_keys(str(rows[6]))
            __createRequisition_PgElement["Salary_From_LPA"].clear()
            __createRequisition_PgElement["Salary_From_LPA"].send_keys(str(int(rows[7])))
            __createRequisition_PgElement["Salary_To_LPA"].clear()
            __createRequisition_PgElement["Salary_To_LPA"].send_keys(str(int(rows[8])))
            __createRequisition_PgElement["Designation"].send_keys(rows[9])
            __createRequisition_PgElement["Expertise"].send_keys(rows[10])
            __createRequisition_PgElement["Role"].send_keys(rows[11])
            __createRequisition_PgElement["Technology_Text"].send_keys(rows[12])
            time.sleep(1)
            __createRequisition_PgElement["Sensitivity"].send_keys(rows[13])
            time.sleep(2)
            __createRequisition_PgElement["Requisition_Owner"].send_keys(rows[14])
            time.sleep(1)
            __createRequisition_PgElement["Requisition_Approver"].send_keys(rows[15])
            time.sleep(1)
            __createRequisition_PgElement["Recruiter"].send_keys(rows[16])
            time.sleep(1)
            __createRequisition_PgElement["Requisition_Name"].clear()
            __createRequisition_PgElement["Requisition_Name"].send_keys(rows[0])
            __createRequisition_PgElement["Requisition_Name"].send_keys(Keys.ENTER)

            time.sleep(10)
        i = i + 1


    @classmethod
    def tearDown(inst):
        # close the browser window
        inst.driver.quit()


if __name__ == '__main__':
    unittest.main()