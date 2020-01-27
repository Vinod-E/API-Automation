import unittest
import collections
import re
import time
import datetime
import xlwt
import xlrd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

from TestScripts.Config.AllConstants import CONSTANT


class candidateGrid(unittest.TestCase):

    def test_candidateGrid(self):
        now = datetime.datetime.now()
        __current_DateTime = now.strftime("%d-%m-%Y")
        print __current_DateTime
        print type(__current_DateTime)
        #Color Coding Code For XLs.
        __style0 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on')
        __style1 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        __style2 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        __style3 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')
        #Writing XLs Sheet Columns.
        wb_result = xlwt.Workbook()
        ws = wb_result.add_sheet('Extract Resume Result')
        ws.write(0, 0, 'Resume File Name', __style0)
        ws.write(0, 1, 'Extracted Primary Skills', __style0)
        ws.write(0, 2, 'Extracted Secondary Skills', __style0)
        ws.write(0, 3, 'Not Extracted Skills', __style0)
        ws.write(0, 4, 'Name', __style0)
        ws.write(0, 5, 'Email Id', __style0)
        ws.write(0, 6, 'Alternate Email Id', __style0)
        ws.write(0, 7, 'DOB (DD/MM/YYYY)', __style0)
        ws.write(0, 8, 'Gender', __style0)
        ws.write(0, 9, 'Location', __style0)
        ws.write(0, 10, 'Mobile Number', __style0)
        ws.write(0, 11, 'Phone Number', __style0)
        ws.write(0, 12, 'College', __style0)
        ws.write(0, 13, 'Degree', __style0)
        ws.write(0, 14, 'Branch', __style0)
        ws.write(0, 15, 'YOP', __style0)
        ws.write(0, 16, 'CGPAorPercentage', __style0)
        ws.write(0, 17, 'IsFinal', __style0)
        ws.write(0, 18, 'Matched Skills (Expected vs Extracted)', __style0)
        ws.write(0, 19, 'PanNo', __style0)
        ws.write(0, 20, 'PassportNo', __style0)
        ws.write(0, 21, 'ExpectedCompanyName', __style0)
        ws.write(0, 22, 'CompanyName', __style0)
        ws.write(0, 23, 'Designation', __style0)
        ws.write(0, 24, 'ExpFrom', __style0)
        ws.write(0, 25, 'ExpTo', __style0)
        ws.write(0, 26, 'Salary', __style0)
        ws.write(0, 27, 'IsLatest', __style0)
        ws.write(0, 28, 'TotalExpInYears', __style0)
        ws.write(0, 29, 'TatalExpInMonths', __style0)
        #Initiating Chrome Browser.
        self.driver = webdriver.Chrome(CONSTANT.CHROME_DRIVER)
        self.driver.implicitly_wait(30)
        self.driver.maximize_window()
        #Getting Url.
        self.driver.get(CONSTANT.RPO_CRPO_AMS_URL)
        time.sleep(1)
        #Login Applecation For Extract Resume
        self.driver.find_element_by_name("loginName").send_keys(CONSTANT.CRPO_RPO_LOGIN_NAME)
        self.driver.find_element_by_xpath("//div[3]/section/div[1]/div[2]/form/div[2]/input").send_keys(
            CONSTANT.CRPO_RPO_LOGIN_PASSWORD)
        self.driver.find_element_by_xpath("//div[3]/section/div[1]/div[2]/form/div[2]/input").send_keys(Keys.ENTER)
        a = self.driver.find_elements_by_xpath('//*[@id="req-list-view"]/tr[1]/td[4]/a')
        for i in a:

            print i


        # total = len(a)
        # print total

    if __name__ == '__main__':
        unittest.main()
