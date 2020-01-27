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
        ws = wb_result.add_sheet('Candidate Grid Data Sheet')
        ws.write(0, 0, 'Status', __style0)
        ws.write(0, 1, 'Candidate Id', __style0)
        ws.write(0, 2, 'Candidate Name', __style0)
        ws.write(0, 3, 'First Name', __style0)
        ws.write(0, 4, 'Middle Name', __style0)
        ws.write(0, 5, 'Last Name', __style0)
        ws.write(0, 6, 'Primary Email Id', __style0)
        ws.write(0, 7, 'Secondary Email Id', __style0)
        ws.write(0, 8, 'Primary Mobile Number', __style0)
        ws.write(0, 9, 'Secondary Phone Number', __style0)
        ws.write(0, 10, 'USN', __style0)
        ws.write(0, 11, 'Total Exp', __style0)
        ws.write(0, 12, 'Exp In Year', __style0)
        ws.write(0, 13, 'Location', __style0)
        ws.write(0, 14, 'DOB (DD/MM/YYYY)', __style0)
        ws.write(0, 15, 'Gender', __style0)
        ws.write(0, 16, 'College', __style0)
        ws.write(0, 17, 'Degree', __style0)
        ws.write(0, 18, 'Country', __style0)
        ws.write(0, 19, 'CGPAorPercentage', __style0)
        ws.write(0, 20, 'SSN No', __style0)
        ws.write(0, 21, 'Current CTC', __style0)
        ws.write(0, 22, 'Current Employer', __style0)
        ws.write(0, 23, 'Department', __style0)
        ws.write(0, 24, 'Marital Status', __style0)
        ws.write(0, 25, 'Nationality', __style0)
        ws.write(0, 26, 'Notice Period', __style0)
        ws.write(0, 27, 'Original Source', __style0)
        ws.write(0, 28, 'PanNo', __style0)
        ws.write(0, 29, 'PassportNo', __style0)
        ws.write(0, 30, 'Sensitivity', __style0)
        ws.write(0, 31, 'Skills', __style0)
        ws.write(0, 32, 'PrimaryExpertise', __style0)
        ws.write(0, 33, 'SecondaryExpertise', __style0)
        ws.write(0, 34, 'LinkedinLink', __style0)
        ws.write(0, 35, 'CreatedOn', __style0)
        ws.write(0, 36, 'CreatedBy', __style0)
        ws.write(0, 37, 'ModifiedOn', __style0)
        ws.write(0, 38, 'ModifiedBy', __style0)
        ws.write(0, 39, 'SourceModifiedOn', __style0)
        ws.write(0, 40, 'SourceModifiedBy', __style0)
        ws.write(0, 41, 'PhoneResidence', __style0)
        ws.write(0, 42, 'Technology', __style0)
        ws.write(0, 43, 'Sourcer', __style0)
        ws.write(0, 44, 'Source', __style0)
        ws.write(0, 45, 'OriginatedFrom', __style0)
        ws.write(0, 46, 'Role', __style0)
        ws.write(0, 47, 'Integer1', __style0)
        ws.write(0, 48, 'integer2', __style0)
        ws.write(0, 49, 'Integer3', __style0)
        ws.write(0, 50, 'Integer4', __style0)
        ws.write(0, 51, 'Integer5', __style0)
        ws.write(0, 52, 'Integer6', __style0)
        ws.write(0, 53, 'Integer7', __style0)
        ws.write(0, 54, 'Integer8', __style0)
        ws.write(0, 55, 'Integer9', __style0)
        ws.write(0, 56, 'Integer10', __style0)
        ws.write(0, 57, 'Integer11', __style0)
        ws.write(0, 58, 'Integer12', __style0)
        ws.write(0, 59, 'Integer13', __style0)
        ws.write(0, 60, 'Integer14', __style0)
        ws.write(0, 61, 'Integer15', __style0)
        ws.write(0, 62, 'Text1', __style0)
        ws.write(0, 63, 'Text2', __style0)
        ws.write(0, 64, 'Text3', __style0)
        ws.write(0, 65, 'Text4', __style0)
        ws.write(0, 66, 'Text5', __style0)
        ws.write(0, 67, 'Text6', __style0)
        ws.write(0, 68, 'Text7', __style0)
        ws.write(0, 69, 'Text8', __style0)
        ws.write(0, 70, 'Text9', __style0)
        ws.write(0, 71, 'Text10', __style0)
        ws.write(0, 72, 'Text11', __style0)
        ws.write(0, 73, 'Text12', __style0)
        ws.write(0, 74, 'Text13', __style0)
        ws.write(0, 75, 'Text14', __style0)
        ws.write(0, 76, 'Text15', __style0)
        ws.write(0, 77, 'TrueFalse1', __style0)
        ws.write(0, 78, 'TrueFalse2', __style0)
        ws.write(0, 79, 'TrueFalse3', __style0)
        ws.write(0, 80, 'TrueFalse4', __style0)
        ws.write(0, 81, 'TrueFalse5', __style0)
        ws.write(0, 82, 'TextArea1', __style0)
        ws.write(0, 83, 'TextArea2', __style0)
        ws.write(0, 84, 'TextArea3', __style0)
        ws.write(0, 85, 'TextArea4', __style0)
        ws.write(0, 86, 'DateCustom1', __style0)
        ws.write(0, 87, 'DateCustom2', __style0)
        ws.write(0, 88, 'DateCustom3', __style0)
        ws.write(0, 89, 'DateCustom4', __style0)
        ws.write(0, 90, 'DateCustom5', __style0)
        ws.write(0, 91, 'SecondaryAdd', __style0)
        ws.write(0, 92, 'PrimaryAdd', __style0)
        ws.write(0, 93, 'Notes', __style0)
        ws.write(0, 94, 'RelevantExp', __style0)
        ws.write(0, 95, 'TypeOfCandid', __style0)

        # Initiating Chrome Browser.
        self.driver = webdriver.Chrome(CONSTANT.CHROME_DRIVER)
        self.driver.implicitly_wait(30)
        self.driver.maximize_window()
        # Getting Url.
        self.driver.get(CONSTANT.RPO_CRPO_AMS_URL_Cand)
        time.sleep(1)
        # Login Application For Extract Resume
        self.driver.find_element_by_name("loginName").send_keys(CONSTANT.CRPO_RPO_LOGIN_NAME)
        self.driver.find_element_by_xpath("//div[3]/section/div[1]/div[2]/form/div[2]/input").send_keys(
            CONSTANT.CRPO_RPO_LOGIN_PASSWORD_Admin)
        self.driver.find_element_by_xpath("//div[3]/section/div[1]/div[2]/form/div[2]/input").send_keys(Keys.ENTER)

        tds = self.driver.find_elements_by_xpath('//table/tbody/tr')
        wb = xlrd.open_workbook("C:\PythonAutomation\AllGridInput\CandidateGridVarificationInput.xls")
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
                        ws.write(rownum+1, column + 1, i.text or "Empty", __style3)
                    elif i.text:
                        is_identical = False
                        # print "elif " + str(i.text)
                        ws.write(rownum+1, column + 1, i.text, __style2)
                    else:
                        is_identical = False
                        # print "else " + str(is_identical)
                        ws.write(rownum+1, column + 1, "Empty", __style3)
                    ws.write(rownum, column + 1, rows[column], __style0)
                    column += 1
                # print "td    " + i.text
                count += 1
            ws.write(rownum+1, 0, "Pass" if is_identical else "Fail", __style3 if is_identical else __style2)
            # if rownum == 3:
            #     break
            irownum += 1
            rownum += 2
        wb_result.save('C:\PythonAutomation\CandidateGridResults\CandidateGridResults(' + __current_DateTime + ').xls')

    if __name__ == '__main__':
        unittest.main()
