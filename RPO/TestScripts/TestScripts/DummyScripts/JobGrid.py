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
        # Color Coding Code For XLs.
        __style0 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on')
        __style1 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        __style2 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        __style3 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')
        __style4 = xlwt.easyxf('font: name Times New Roman, color-index blue, bold on')
        # Writing XLs Sheet Columns.
        wb_result = xlwt.Workbook()
        ws = wb_result.add_sheet('Job Grid Data Sheet')
        ws.write(0, 0, 'Status', __style0)
        ws.write(0, 1, 'Job Id', __style0)
        ws.write(0, 2, 'Job Name', __style0)
        ws.write(0, 3, 'Number Of Openings', __style0)
        ws.write(0, 4, 'Job Title', __style0)
        ws.write(0, 5, 'Recruiter', __style0)
        ws.write(0, 6, 'Hiring Manager', __style0)
        ws.write(0, 7, 'Job Type', __style0)
        ws.write(0, 8, 'Number Of Applicants', __style0)
        ws.write(0, 9, 'Job Status', __style0)
        ws.write(0, 10, 'Job Location', __style0)
        ws.write(0, 11, 'Unit', __style0)
        ws.write(0, 12, 'Department', __style0)
        ws.write(0, 13, 'Job Code', __style0)
        ws.write(0, 14, 'Opening Left)', __style0)
        ws.write(0, 15, 'Delay Days', __style0)
        ws.write(0, 16, 'Functional Area', __style0)
        ws.write(0, 17, 'Selection Process', __style0)
        ws.write(0, 18, 'Grade', __style0)
        ws.write(0, 19, 'Number Of Offers', __style0)
        ws.write(0, 20, 'Experience Range', __style0)
        ws.write(0, 21, 'Replacement', __style0)
        ws.write(0, 22, 'Skill Category', __style0)
        ws.write(0, 23, 'CTC Range', __style0)
        ws.write(0, 24, 'Cost centre', __style0)
        ws.write(0, 25, 'Work Mode', __style0)
        ws.write(0, 26, 'Travel required', __style0)
        ws.write(0, 27, 'Notes', __style0)
        ws.write(0, 28, 'Sensitivity', __style0)
        ws.write(0, 29, 'Diversity', __style0)
        ws.write(0, 30, 'Requisition Release Date', __style0)
        ws.write(0, 31, 'Target Applicant On Boarding Date', __style0)
        ws.write(0, 32, 'Created By', __style0)
        ws.write(0, 33, 'Created On', __style0)
        ws.write(0, 34, 'Requirement Name', __style0)
        ws.write(0, 35, 'Requirement Id', __style0)
        ws.write(0, 36, 'Customer', __style0)
        ws.write(0, 37, 'Integer1', __style0)
        ws.write(0, 38, 'Integer2', __style0)
        ws.write(0, 39, 'Integer3', __style0)
        ws.write(0, 40, 'Integer4', __style0)
        ws.write(0, 41, 'Integer5', __style0)
        ws.write(0, 42, 'Text1', __style0)
        ws.write(0, 43, 'Text2', __style0)
        ws.write(0, 44, 'Text3', __style0)
        ws.write(0, 45, 'Text4', __style0)
        ws.write(0, 46, 'Text5', __style0)
        ws.write(0, 47, 'DateCustomField1', __style0)
        ws.write(0, 48, 'DateCustomField2', __style0)

        # Initiating Chrome Browser.
        self.driver = webdriver.Chrome(CONSTANT.CHROME_DRIVER)
        self.driver.implicitly_wait(30)
        self.driver.maximize_window()
        # Getting Url.
        self.driver.get(CONSTANT.RPO_CRPO_AMS_URL_Job)
        time.sleep(1)
        # Login Application
        self.driver.find_element_by_name("loginName").send_keys(CONSTANT.CRPO_RPO_LOGIN_NAME)
        self.driver.find_element_by_xpath("//div[3]/section/div[1]/div[2]/form/div[2]/input").send_keys(
            CONSTANT.CRPO_RPO_LOGIN_PASSWORD_Admin)
        self.driver.find_element_by_xpath("//div[3]/section/div[1]/div[2]/form/div[2]/input").send_keys(Keys.ENTER)

        tds = self.driver.find_elements_by_xpath('//table/tbody/tr')
        wb = xlrd.open_workbook("C:\PythonAutomation\AllGridInput\JobGridVarificationInput.xls")
        sheetname = wb.sheet_names()  # Read for XLS Sheet names
        print(sheetname)
        sh1 = wb.sheet_by_index(0)
        rownum = 1
        irownum = 1
        is_identical = True

        for row in tds:
            rows = sh1.row_values(irownum)
            print type(rows)
            # print row
            # print([td.text for td in row.find_elements_by_xpath("//table/tbody/tr/td")])
            row_data = row.find_elements_by_tag_name('td')
            # print type(row_data)
            # print row_data
            print "============Success Row===================="
            count = 0
            column = 0
            for i in row_data:
                # print "i.text " + str(i.text)
                if count > 1:
                    if i.text == rows[column] or (not i.text and rows[column] == "Empty"):
                        # print rows[column]
                        ws.write(rownum + 1, column + 1, i.text or "Empty", __style3)
                    elif i.text:
                        is_identical = False
                        # print "elif " + str(i.text)
                        ws.write(rownum + 1, column + 1, i.text, __style2)
                    else:
                        is_identical = False
                        # print "else " + str(is_identical)
                        ws.write(rownum + 1, column + 1, "Empty", __style3)
                    ws.write(rownum, column + 1, rows[column], __style0)
                    column += 1
                # print "td    " + i.text
                count += 1
            ws.write(rownum + 1, 0, "Pass" if is_identical else "Fail", __style3 if is_identical else __style2)
            # if rownum == 3:
            #     break
            irownum += 1
            rownum += 2
        wb_result.save('C:\PythonAutomation\JobGridResults\JobGridResult(' + __current_DateTime + ').xls')

    if __name__ == '__main__':
        unittest.main()