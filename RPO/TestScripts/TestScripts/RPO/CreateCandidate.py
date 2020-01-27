import unittest
import time
import datetime
import xlrd
import xlwt
from TestScripts.Config.AllConstants import *
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from TestScripts.RPO_PageObjects.CreateCandidate_PgObj import createCandidate_PgObj
from selenium.webdriver.support.ui import Select
from sqlite3 import Date
from urlparse import urlparse
from selenium.common.exceptions import NoSuchElementException, NoAlertPresentException
from selenium.webdriver.common.by import By

from TestScripts.Config.AllInputDataFilePath import file_Path


class createCandidate(unittest.TestCase):
    @classmethod
    def setUp(inst):
        # create a new browser session """
        inst.driver = webdriver.Chrome(CONSTANT.CHROME_DRIVER)
        inst.driver.implicitly_wait(30)
        inst.driver.maximize_window()
        # navigate to the application home page
        inst.driver.get(CONSTANT.INTERNAL_AMS_URL)
        inst.driver.title
        return inst.driver

    def test_createCandidate(self):
        self.driver.find_element_by_name("clientEmail").send_keys(CONSTANT.INTERNAL_AMS_LOGIN_NAME)
        self.driver.find_element_by_name("new-password").send_keys(CONSTANT.INTERNAL_AMS_LOGIN_PASSWORD)
        self.driver.find_element_by_xpath("//div[2]/div/div[1]/div/div/div[4]/div[1]/div[1]/div/button").click()
        time.sleep(5)
        self.driver.find_element_by_xpath("//div[1]/div[2]/div/div[1]/div[1]/div/ul/li[2]/a").click()
        time.sleep(1)
        wb = xlrd.open_workbook(file_Path.internal_AMS_File_Path()+".xls")
        sheetname = wb.sheet_names()  # Read for XLS Sheet names
        sh1 = wb.sheet_by_index(0)
        i = 1
        while (i < sh1.nrows):
            rownum = (i)
            rows = sh1.row_values(rownum)
            time.sleep(2)
            __current_date = Date.today()
            __past_date = __current_date + datetime.timedelta(days=-11315)
            __dOB = __past_date.strftime("%d/%m/%Y")
            self.driver.find_element_by_xpath("//p[2]").click()
            time.sleep(2)
            __createCandidate_PgElement = createCandidate_PgObj.createCandidate_PgElements(self.driver)
            __createCandidate_PgElement["Upload_Resume"].clear()
            __createCandidate_PgElement["Upload_Resume"].send_keys(CONSTANT.RESUME_FILE_PATH)
            __createCandidate_PgElement["Profile_Picture"].clear()
            __createCandidate_PgElement["Profile_Picture"].send_keys(CONSTANT.PROFILE_PIC_FILE_PATH)
            time.sleep(2)



            __extract_Resume = self.driver.find_element_by_xpath("//div[4]/div/button")
            __extract_Resume.click()
            time.sleep(4)

            SkillsOfCandidate = self.driver.find_element_by_xpath("//div[5]/md-dialog/md-dialog-content/div[2]/div/div/div[2]/form/div/div/div[5]/div").text

            Skills = SkillsOfCandidate.replace("\nRemove\nPress delete to remove this chip.", ",")
            Skills = [s.strip() for s in Skills.split(',')]
            print Skills

            __createCandidate_PgElement["Candidate_Name"].clear()
            __createCandidate_PgElement["Candidate_Name"].send_keys(rows[0])
            time.sleep(1)
            __createCandidate_PgElement["Email"].clear()
            __createCandidate_PgElement["Email"].send_keys(rows[1])
            time.sleep(1)
            __createCandidate_PgElement["Alternate_Email"].clear()
            __createCandidate_PgElement["Alternate_Email"].send_keys(rows[2])
            time.sleep(1)
            __createCandidate_PgElement["DOB"].clear()
            __createCandidate_PgElement["DOB"].send_keys("01/01/1985")
            time.sleep(1)
            __createCandidate_PgElement["Gender"].send_keys(rows[3])
            __createCandidate_PgElement["Location"].send_keys(rows[4])
            time.sleep(1)
            __createCandidate_PgElement["Mobile"].clear()
            __createCandidate_PgElement["Mobile"].send_keys(str(int(rows[5])))
            time.sleep(1)
            __createCandidate_PgElement["Phone_No"].clear()
            __createCandidate_PgElement["Phone_No"].send_keys(str(int(rows[6])))
            time.sleep(1)
            # __createCandidate_PgElement["IsFinal_Education"].click()
            __createCandidate_PgElement["IsLatest_Experience"].click()
            time.sleep(1)
            __createCandidate_PgElement["Sensitivity"].send_keys(rows[7])
            __createCandidate_PgElement["Candidate_Status"].send_keys(rows[8])
            __createCandidate_PgElement["Candidate_Sourcer"].send_keys(rows[9])
            __createCandidate_PgElement["Expertise"].send_keys(rows[10])
            time.sleep(1)
            __createCandidate_PgElement["Experience_InYear"].clear()
            __createCandidate_PgElement["Experience_InYear"].send_keys(str(rows[11]))
            time.sleep(1)
            __createCandidate_PgElement["Experience_InMonths"].clear()
            __createCandidate_PgElement["Experience_InMonths"].send_keys(str(rows[12]))
            time.sleep(1)
            __createCandidate_PgElement["Current_Salary_LPA"].clear()
            __createCandidate_PgElement["Current_Salary_LPA"].send_keys(str(rows[13]))
            __createCandidate_PgElement["Source_Type"].send_keys(rows[14])
            __createCandidate_PgElement["Source"].send_keys(rows[15])
            __createCandidate_PgElement["Willing_To_Relocate"].send_keys(rows[16])
            # LocationPreference = self.driver.find_element_by_name("locationpreference")
            # LocationPreference.click()
            __createCandidate_PgElement["Expected_Salary_From_LPA"].clear()
            __createCandidate_PgElement["Expected_Salary_From_LPA"].send_keys(str(rows[17]))
            time.sleep(1)
            __createCandidate_PgElement["Expected_Salary_To_LPA"].clear()
            __createCandidate_PgElement["Expected_Salary_To_LPA"].send_keys(str(rows[18]))
            time.sleep(1)
            __createCandidate_PgElement["Notice_Period_Days"].clear()
            __createCandidate_PgElement["Notice_Period_Days"].send_keys(str(rows[19]))
            time.sleep(1)
            # __createCandidate_PgElement["College"].send_keys(rows[20])
            # __createCandidate_PgElement["Degree"].send_keys(rows[21])
            # __createCandidate_PgElement["Branch"].send_keys(rows[22])
            # __createCandidate_PgElement["YOP"].send_keys(str(int(rows[23])))
            self.driver.find_element_by_name("cgpa").clear()
            self.driver.find_element_by_name("cgpa").send_keys(str(rows[24]))
            # time.sleep(1)
            # __createCandidate_PgElement["Company"].send_keys(rows[25])
            # __createCandidate_PgElement["Designation/Role"].send_keys(rows[26])
            # __createCandidate_PgElement["Experience_From"].send_keys(str(int(rows[27])))
            # __createCandidate_PgElement["Experience_To"].send_keys(str(int(rows[28])))
            # __createCandidate_PgElement["Salary"].clear()
            # __createCandidate_PgElement["Salary"].send_keys(str(rows[29]))
            time.sleep(1)


            __createCandidate_PgElement["Reason_For_Leaving"].clear()
            __createCandidate_PgElement["Reason_For_Leaving"].send_keys(rows[30])
            time.sleep(1)
            __createCandidate_PgElement["Reason_For_Leaving"].send_keys(Keys.ENTER)
            time.sleep(10)
        i = i + 1

    @classmethod
    def tearDown(inst):
        # close the browser window
        inst.driver.quit()

if __name__ == '__main__':
    unittest.main()