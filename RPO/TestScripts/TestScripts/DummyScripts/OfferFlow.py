import unittest
import collections
import re
import time
import datetime
from lib2to3.pgen2 import driver
from webbrowser import browser

import xlwt
import xlrd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

from TestScripts.Config.AllConstants import CONSTANT
from selenium.webdriver.support.select import Select


class offerCalculation(unittest.TestCase):

    def test_offerCalculation(self):
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
        ws = wb_result.add_sheet('Offer Flow Break Up Sheet')
        ws.write(0, 0, 'NameOfCTCStructure', __style0)
        ws.write(0, 1, 'OfferedCTC', __style0)
        ws.write(0, 2, 'Status', __style0)
        ws.write(0, 3, 'Component A', __style0)
        ws.write(0, 4, 'Non-Monetary Benefits', __style0)
        ws.write(0, 5, 'CTC', __style0)
        ws.write(0, 6, 'Basic Salary', __style0)
        ws.write(0, 7, 'HRA', __style0)
        ws.write(0, 8, 'Conveyance', __style0)
        ws.write(0, 9, 'Medical Allowance', __style0)
        ws.write(0, 10, 'Statutory Bonus', __style0)
        ws.write(0, 11, 'Performance Linked Bonus', __style0)
        ws.write(0, 12, 'PF Employer', __style0)
        ws.write(0, 13, 'Gratuity', __style0)
        ws.write(0, 14, 'Special Allowance', __style0)
        ws.write(0, 15, 'Gross Salary', __style0)
        ws.write(0, 16, 'ESI Employer', __style0)
        ws.write(0, 17, 'Total Retirement Benefits', __style0)
        ws.write(0, 18, 'Cost To Company(CTC)', __style0)
        ws.write(0, 19, 'Total Cost To Company(TCTC)', __style0)
        # ws.write(0, 20, 'Fuel Allowance', __style0)
        # ws.write(0, 21, 'Car Allowance', __style0)

        # Initiating Chrome Browser.
        self.driver = webdriver.Chrome(CONSTANT.CHROME_DRIVER)
        self.driver.implicitly_wait(30)
        self.driver.maximize_window()
        # Getting Url.
        self.driver.get(CONSTANT.RPO_CRPO_AMS_URL_Offer)
        time.sleep(1)
        # Login Application
        self.driver.find_element_by_name("loginName").send_keys(CONSTANT.CRPO_RPO_LOGIN_NAME_Admin)
        self.driver.find_element_by_xpath("//div[3]/section/div[1]/div[2]/form/div[2]/input").send_keys(
            CONSTANT.CRPO_RPO_LOGIN_PASSWORD)

        self.driver.find_element_by_xpath("//div[3]/section/div[1]/div[2]/form/div[2]/input").send_keys(Keys.ENTER)

        self.driver.find_element_by_xpath('//*[@id="req-list-view"]/tr[1]/td[1]/span[1]/a/i').click()
        time.sleep(5)
        before = self.driver.window_handles[1]
        self.driver.switch_to.window(before)
        time.sleep(1)
        wb = xlrd.open_workbook("C:\PythonAutomation\OfferFlowCTCBrakup\CTCBreakupInputSheet\OfferFlowCTCBreakUpInputAutoTenant.xls")
        sheetname = wb.sheet_names()
        sh1 = wb.sheet_by_index(0)
        xlrow = 1
        xlrow1 = 2
        xlrow2 = 1
        #print sh1.nrows

        while xlrow < sh1.nrows:

            salStrName = sh1.cell(xlrow, 0)
            salStrName1 = str(salStrName.value)
            ctcInputCell = sh1.cell(xlrow, 1)
            ctcInputValue = str(ctcInputCell.value)
            componentA = sh1.cell(xlrow, 2)
            componentAvalue = str(componentA.value)
            s1 = self.driver.find_element_by_xpath("//div[3]/div/div/div[2]/div/div/div[1]/transcluded-input/div/div/div[1]/div/div/ta-dropdown/div/div/input")
            s1.send_keys(salStrName1)
            s1.send_keys(Keys.DOWN)
            s1.send_keys(Keys.ENTER)
            time.sleep(2)
            s2 = self.driver.find_element_by_xpath("//div[3]/div/div/div[2]/div/div/div[1]/transcluded-input/div/div/div[2]/div[1]/div/div[1]/input")
            s2.send_keys(ctcInputValue)
            time.sleep(5)
            calculate = self.driver.find_element_by_xpath("//div[3]/div/div/div[2]/div/div/div[1]/transcluded-input/div/div/div[2]/div[10]/div/button")
            calculate.click()
            time.sleep(9)
            breakUpRow = self.driver.find_elements_by_xpath('//table/tbody/tr')
            #print breakUpRow
            splitele1 = None
            #rownum = 1
            icolumn = 3
            ocolumn = 4
            is_identical = True
            for row in breakUpRow:
                rows = sh1.row_values(xlrow)
                breakUpCol = row.find_elements_by_tag_name('td')
                for i in breakUpCol:
                    try:
                        li = i.text.split(':')[1]
                        li1 = li.split('(')[0]
                        li2 = li1.strip()
                        #print(rows[icolumn])
                        #print li2
                        if str(li2) == str(int(rows[icolumn])):
                            ws.write(xlrow1, ocolumn, li2, __style3)
                            print("Passed")
                        elif li2:
                            is_identical = False
                            ws.write(xlrow1, ocolumn, li2, __style2)
                            print("Not Passed")
                        else:
                            is_identical = False
                            ws.write(xlrow1, ocolumn, 'Failed', __style3)
                            print("Failed")
                        ws.write(xlrow2, ocolumn, (rows[icolumn]), __style1)
                        icolumn += 1
                        ocolumn += 1


                    except:
                        pass

            ws.write(xlrow2, 0, salStrName1, __style1)
            ws.write(xlrow2, 1, ctcInputValue, __style1)
            ws.write(xlrow1, 3, componentAvalue, __style3)
            ws.write(xlrow1-1, 3, componentAvalue, __style1)
            ws.write(xlrow1, 2, "Pass" if is_identical else "Fail", __style3 if is_identical else __style2)


            s1.clear()
            s2.clear()
            xlrow += 1
            xlrow1 += 2
            xlrow2 += 2
        time.sleep(40)

        # self.driver.find_element_by_xpath('//*[@id="req-list-view"]/tr[1]/td[1]/span[3]/a').click()
        # self.driver.find_element_by_xpath('//*[@id="req-list-view"]/tr[1]/td[1]/div/div[3]/div[1]/div/div[2]').click()
        wb_result.save('C:\PythonAutomation\OfferFlowCTCBrakup\OfferFlowCTCBrakupResults\OfferCTCBreakupRes(' + __current_DateTime + ').xls')
        print 'sanjeev'


    if __name__ == '__main__':
        unittest.main()