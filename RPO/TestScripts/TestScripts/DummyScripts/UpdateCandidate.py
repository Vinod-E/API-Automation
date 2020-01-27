import unittest
import re
from selenium.webdriver.support.ui import WebDriverWait
import time
import datetime
import xlwt
import xlrd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from TestScripts.Config.AllConstants import CONSTANT
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from random import randint
import sys
reload(sys)
sys.setdefaultencoding('utf8')


class UpdateCabdidate():
    now = datetime.datetime.now()
    __current_DateTime = now.strftime("%d-%m-%Y")
    def __init__(self):
        self.input_sheet = xlrd.open_workbook("C:\PythonAutomation\UpdateCandidateInputs\BasicInputs.xls")
        self.input_file = self.input_sheet.sheet_by_index(0)
        self.updatevarification = xlrd.open_workbook(
            "C:\PythonAutomation\UpdateCandidateInputs\UpdateCandidateVarification.xls")
        self.input1 = self.updatevarification.sheet_by_index(0)
        self.failedCount = 0
        self.test_login()
        self.save_results()
        time.sleep(1)
        self.update_resume()
        time.sleep(1)
        self.photo_update()
        time.sleep(1)
        self.update_otheerAttach()
        time.sleep(1)
        self.update_exp_details()
        time.sleep(3)
        self.update_education_deatails()
        self.update_personal_details()

    def test_login(self):
        now = datetime.datetime.now()
        __current_DateTime = now.strftime("%d-%m-%Y")
        self.driver = webdriver.Chrome(CONSTANT.CHROME_DRIVER)
        self.driver.implicitly_wait(30)
        self.driver.maximize_window()
        # Getting Url.
        self.driver.get(CONSTANT.RPO_CRPO_AMS_URL_UPDATE_CAN)
        time.sleep(1)
        # Login Applecation For Extract Resume
        self.driver.find_element_by_name("loginName").send_keys(CONSTANT.CRPO_RPO_LOGIN_NAME_MONU)
        self.driver.find_element_by_xpath("//div[3]/section/div[1]/div[2]/form/div[2]/input").send_keys(
            CONSTANT.CRPO_RPO_LOGIN_PASSWORD_MONU)
        self.driver.find_element_by_xpath("//div[3]/section/div[1]/div[2]/form/div[2]/input").send_keys(Keys.ENTER)
        # pageWait = self.driver.find_element_by_xpath('//*[@ng-if="column.actionNeeded && !column.isPopover"]')
        time.sleep(6)
        self.driver.find_element_by_xpath('//*[@ng-if="column.actionNeeded && !column.isPopover"]').click()
        time.sleep(6)
        moveToNextTab = self.driver.window_handles[1]
        self.driver.switch_to.window(moveToNextTab)
        time.sleep(1)

    def save_results(self):
        __style0 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on')
        self.__style1 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        self.__style2 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        self.__style3 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')
        __style4 = xlwt.easyxf('font: name Times New Roman, color-index blue, bold on')
        self.__style5 = xlwt.easyxf(
            'pattern: pattern solid, fore_colour yellow;' 'font: name Times New Roman, color-index red, bold on;')
        # Writing XLs Sheet Columns.
        self.wb_result = xlwt.Workbook()
        self.ws = self.wb_result.add_sheet('Candidate Updated Sheet', cell_overwrite_ok = True)
        # , cell_overwrite_ok = True
        self.ws.write(0, 0, 'Status', __style0)
        self.ws.write(0, 1, 'InputForResume', __style0)
        self.ws.write(0, 2, 'InputForPhoto', __style0)
        self.ws.write(0, 3, 'InputForOtherAttach', __style0)
        self.ws.write(0, 4, 'ResumeUpdateStatus', __style0)
        self.ws.write(0, 5, 'PhotoUpdateStatus', __style0)
        self.ws.write(0, 6, 'OtherAttachUpdateStatus', __style0)
        self.ws.write(0, 7, 'EducationUpdateStatus', __style0)
        self.ws.write(0, 8, 'Degree', __style0)
        self.ws.write(0, 9, 'Percentage/CGPA', __style0)
        self.ws.write(0, 10, 'College', __style0)
        self.ws.write(0, 11, 'Department', __style0)
        self.ws.write(0, 12, 'YearOFPassing', __style0)
        self.ws.write(0, 13, 'ExperienceUpdateStatus', __style0)
        self.ws.write(0, 14, 'Company', __style0)
        self.ws.write(0, 15, 'Designation', __style0)
        self.ws.write(0, 16, 'Duration', __style0)

        self.ws.write(0, 17, 'CandidateUpdateStatus', __style0)
        self.ws.write(0, 18, 'CandidateName', __style0)
        self.ws.write(0, 19, 'EmailId', __style0)
        self.ws.write(0, 20, 'MobileNumber', __style0)
        self.ws.write(0, 21, 'Location', __style0)
        self.ws.write(0, 22, 'DOB (DD/MM/YYYY)', __style0)
        self.ws.write(0, 23, 'Total Experience', __style0)
        self.ws.write(0, 24, 'Candidate Source', __style0)
        self.ws.write(0, 25, 'USN', __style0)
        self.ws.write(0, 26, 'Gender', __style0)
        self.ws.write(0, 27, 'PassportNumber', __style0)
        self.ws.write(0, 28, 'AadharNumber', __style0)
        self.ws.write(0, 29, 'PanNumber', __style0)
        self.ws.write(0, 30, 'Candidate Status', __style0)
        self.ws.write(0, 31, 'SSNNumber', __style0)
        # self.ws.write(0, 32, 'Degree', __style0)
        # self.ws.write(0, 33, 'Percentage/CGPA', __style0)
        # self.ws.write(0, 34, 'CollegeName', __style0)
        # self.ws.write(0, 35, 'Branch', __style0)
        # self.ws.write(0, 36, 'YOP', __style0)
        # self.ws.write(0, 37, 'Company', __style0)
        # self.ws.write(0, 38, 'Designation', __style0)
        # self.ws.write(0, 39, 'Duration', __style0)
    def update_resume(self):
        actionButton = self.driver.find_element_by_xpath('//*[@class="action-button"]')
        actionButton.click()
        resUpdatebutton = self.driver.find_element_by_xpath('//*[@title="Update Candidate Resume"]')
        resUpdatebutton.click()
        time.sleep(2)
        uploadResume = self.driver.find_element_by_xpath('//*[@file-model="vm.file"]')
        # resfilepath = self.input_file.cell(1, 0).value
        #print ((str(self.input_file.cell(1, 0).value)))
        #time.sleep(5)
        uploadResume.send_keys(str(self.input_file.cell(1, 0).value))
        time.sleep(2)
        updateButton = self.driver.find_element_by_xpath('//*[@ng-click="vm.actionClicked(\'update\');"]')
        updateButton.click()
        time.sleep(7)
        for i in range(10):
            try:
                succMessage = self.driver.find_element_by_xpath('//*[@ng-bind-html="message.text"]').text
                print (succMessage)
                break
            except NoSuchElementException as e:
                print ("Retry in 1 sec")
                time.sleep(1)
        else:
            raise e
        self.ws.write(1, 4, "Resume updated successfully", self.__style1)
        self.ws.write(2, 1, self.input_file.cell(1, 0).value, self.__style1)
        if succMessage == "Resume updated successfully":
            actionButtonn = self.driver.find_element_by_xpath('//*[@class="action-button"]')
            actionButtonn.click()
            resUpdatebuttonn = self.driver.find_element_by_xpath('//*[@title="Update Candidate Resume"]')
            resUpdatebuttonn.click()
            filePath = self.driver.find_element_by_xpath('//*[@ng-if="!vm.data.resume.url"]').text
            cancelButton = self.driver.find_element_by_xpath('//*[@ng-click="vm.actionClicked(\'cancel\')"]')
            cancelButton.click()
            print (filePath)
            if filePath == "Resume2.doc":
                self.ws.write(2, 4, "Resume updated successfully", self.__style3)
                self.ws.write(2, 0, "Test Case Passed", self.__style3)
                print("Pass")
            else:
                self.ws.write(2, 4, "Different File updated successfully", self.__style2)
                self.ws.write(2, 0, "Test Case Failed", self.__style2)
                print ("Different File")
        else:
            self.ws.write(2, 4, "Resume Not updated successfully", self.__style2)
            self.ws.write(2, 0, "Test Case Failed", self.__style2)
            print ("Failed")
    def photo_update(self):
        clickOnEdit = self.driver.find_element_by_xpath('//*[@ng-click="vm.actionClicked(\'editPhoto\');"]')
        clickOnEdit.click()
        uploadPhoto = self.driver.find_element_by_xpath('//*[@file-model="vm.file"]')
        # photofilepath = self.input_file.cell_value(2,1)
        uploadPhoto.send_keys(self.input_file.cell_value(2,1))
        time.sleep(10)
        clickOnUpdate = self.driver.find_element_by_xpath('//*[@ng-click="vm.actionClicked(\'update\');"]')
        clickOnUpdate.click()
        time.sleep(6)
        for i in range(10):
            try:
                succMessage = self.driver.find_element_by_xpath('//*[@ng-bind-html="message.text"]').text
                print (succMessage)
                break
            except NoSuchElementException as e:
                print ("Retry in 1 sec")
                time.sleep(1)
        else:
            raise e
        self.ws.write(3, 5, "Photo updated successfully", self.__style1)
        self.ws.write(4, 2, self.input_file.cell_value(2, 1), self.__style1)
        if succMessage == "Photo updated successfully":
            time.sleep(6)
            clickOnEdit = self.driver.find_element_by_xpath('//*[@ng-click="vm.actionClicked(\'editPhoto\');"]')
            clickOnEdit.click()
            time.sleep(1)
            photofile = self.driver.find_element_by_xpath('//*[@style="font-weight:bold;color:#a97b68"]').text
            close = self.driver.find_element_by_xpath('//*[@ng-click="vm.actionClicked(\'cancel\');"]')
            close.click()
            print (photofile)
            if photofile == "Pic2.jpg":
                self.ws.write(4, 5, "Photo updated successfully", self.__style3)
                self.ws.write(4, 0, "Text Case Passed", self.__style3)
                print ("Pass")
            else:
                self.ws.write(4, 5, "Different File updated successfully", self.__style2)
                self.ws.write(4, 0, "Text Case Failed", self.__style2)
                print ("Different File Updated")
        else:
            self.ws.write(4, 5, succMessage, self.__style2)
            self.ws.write(4, 0, "Text Case Failed", self.__style2)
            print ("Failed")
    def update_otheerAttach(self):
        editPencil = self.driver.find_elements_by_xpath('//*[@class="fa fa-pencil pull-right"]')
        editPencil[3].click()
        uploadFile = self.driver.find_element_by_xpath('//*[@file-model="vm.fileToUpLoad"]')
        uploadFile.send_keys(self.input_file.cell_value(3, 2))
        filetype = self.driver.find_element_by_xpath('//*[@placeholder="Type"][@type="text"]')
        filetype.send_keys("Driving License")
        filetype.send_keys(Keys.DOWN)
        filetype.send_keys(Keys.ENTER)
        attachname = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.attachmentName"]')
        attachname.send_keys("UpdatedOtherAttach")
        clickonupdate = self.driver.find_element_by_xpath('//*[@ng-click="vm.actionClicked(\'update\')"]')
        clickonupdate.click()
        time.sleep(4)
        updatedmessage = self.driver.find_element_by_xpath('//*[@ng-bind-html="message.text"]').text
        closebutton = self.driver.find_element_by_xpath('//*[@ng-click="vm.cancel();"]')
        closebutton.click()
        time.sleep(2)
        self.ws.write(5, 6, "Attachment Uploaded Successfully", self.__style1)
        self.ws.write(6, 3, self.input_file.cell_value(3, 2), self.__style1)
        if updatedmessage == "Attachment Uploaded Successfully":
            self.driver.refresh()
            time.sleep(6)
            updatedfilename = self.driver.find_elements_by_xpath('//*[@ng-if="!row.htmlKey && row.key"]')
            filenamelist = []
            for i in updatedfilename:
                filenamelist.append((i.text).split()[0])
                if "UpdatedOtherAttach" in filenamelist:
                    self.ws.write(6, 6, updatedmessage, self.__style3)
                    self.ws.write(6, 0, "Text Case Passed", self.__style3)
                    print ("Test Case Pass")
                else:
                    self.ws.write(6, 6, "Updated File Not Found", self.__style2)
                    self.ws.write(6, 0, "Text Case Failed", self.__style2)
                    print ("Test case Failed")
            print (filenamelist)
        else:
            self.ws.write(6, 6, updatedmessage, self.__style2)
            self.ws.write(4, 0, "Text Case Failed", self.__style2)

    def update_exp_details(self):
        editExp = self.driver.find_element_by_xpath('//*[@title="Edit experience details"]')
        editExp.click()
        company = self.driver.find_element_by_xpath('//*[@placeholder="Company"][@type="text"]')
        designation = self.driver.find_element_by_xpath('//*[@placeholder="Designation"][@type="text"]')
        expFromMonth = self.driver.find_element_by_xpath('//*[@placeholder="Exp Month From"][@type="text"]')
        expFromYear = self.driver.find_element_by_xpath('//*[@placeholder="Exp From Year"][@type="text"]')
        expToMonth = self.driver.find_element_by_xpath('//*[@placeholder="Exp Month To"][@type="text"]')
        expToYear = self.driver.find_element_by_xpath('//*[@placeholder="Exp To Year"][@type="text"]')
        yearlySalary = self.driver.find_element_by_xpath('//*[@ng-model="vm.addExperienceDetails.YearlySalary"]')
        reasonForLeaving = self.driver.find_element_by_xpath('//*[@ng-model="vm.addExperienceDetails.ReasonForLeaving"]')
        self.driver.execute_script("arguments[0].value='Aztec Software';arguments[1].value='Architect';"
                                   "arguments[2].value='Jan';arguments[3].value='2011';arguments[4].value='Feb';"
                                   "arguments[5].value='2013';arguments[6].value='400000';"
                                   "arguments[7].value='ReasonUpdated';", company,
                                   designation, expFromMonth, expFromYear, expToMonth, expToYear, yearlySalary, reasonForLeaving)
        time.sleep(10)
        company.send_keys(Keys.BACKSPACE)
        company.send_keys(Keys.DOWN)
        company.send_keys(Keys.ENTER)
        designation.send_keys(Keys.BACKSPACE)
        designation.send_keys(Keys.DOWN)
        designation.send_keys(Keys.ENTER)
        expFromMonth.send_keys(Keys.BACKSPACE)
        expFromMonth.send_keys(Keys.DOWN)
        expFromMonth.send_keys(Keys.ENTER)
        expFromYear.send_keys(Keys.BACKSPACE)
        expFromYear.send_keys(Keys.DOWN)
        expFromYear.send_keys(Keys.ENTER)
        expToMonth.send_keys(Keys.BACKSPACE)
        expToMonth.send_keys(Keys.DOWN)
        expToMonth.send_keys(Keys.ENTER)
        expToYear.send_keys(Keys.BACKSPACE)
        expToYear.send_keys(Keys.DOWN)
        expToYear.send_keys(Keys.ENTER)
        yearlySalary.send_keys(Keys.BACKSPACE)
        reasonForLeaving.send_keys(Keys.BACKSPACE)
        addExp = self.driver.find_element_by_xpath('//*[@ng-click="vm.addedExperienceDetails();"]')
        addExp.click()
        clickonupdate = self.driver.find_element_by_xpath('//*[@ng-click="vm.save();"]')
        clickonupdate.click()
        time.sleep(3)
        sucMessage = self.driver.find_element_by_xpath('//*[@ng-bind-html="message.text"]').text
        print (sucMessage)
        if sucMessage == "Experience updated successfully":
            self.ws.write(9, 13, "Experience updated successfully", self.__style1)
            self.ws.write(10, 13, sucMessage, self.__style3)
            expData = self.driver.find_elements_by_class_name("row")
            candidate_experince_index = expData[2].text.index('Experience')
            candidate_experince_index_end = expData[2].text.index('Education')
            experience_block = expData[2].text[candidate_experince_index:candidate_experince_index_end]
            print (len(experience_block))
            exp_detailsxl = []
            columm = 14
            for i in range(85, 88):
                exp_detailsxl.append(str(self.input1.cell(5, i).value))
            for j in exp_detailsxl:
                if re.search(j.encode('ascii', 'ignore'), experience_block, re.IGNORECASE):
                    self.ws.write(10, columm,re.search(j.encode('ascii', 'ignore'), experience_block, re.IGNORECASE).group(),
                                  self.__style3)
                else:
                    self.failedCount += 1
                    self.ws.write(10, columm, "Not Available", self.__style2)
                self.ws.write(9, columm, j, self.__style1)
                columm += 1
            print ("Exppp")
            print (exp_detailsxl)
            print (experience_block)
        else:
            self.failedCount += 1
            self.ws.write(10, 13, sucMessage, self.__style2)
        if self.failedCount>= 1:
            self.ws.write(10, 0, "Test Case Failed", self.__style2)
        else:
            self.ws.write(10, 0, "Test Case Pass", self.__style3)
    def update_education_deatails(self):
        editeducation = self.driver.find_element_by_xpath('//*[@title="Edit education details"]')
        editeducation.click()
        time.sleep(2)
        addeducation = self.driver.find_element_by_xpath('//*[@ng-if="education.isAddMoreAllowed"]')
        addeducation.click()
        degree = self.driver.find_elements_by_xpath('//*[@placeholder="Degree"][@type="text"]')
        degree[1].send_keys("B Arch")
        degree[1].send_keys(Keys.DOWN)
        degree[1].send_keys(Keys.ENTER)
        branch = self.driver.find_elements_by_xpath('//*[@placeholder="Branch"][@type="text"]')
        branch[1].send_keys("Anthropology")
        branch[1].send_keys(Keys.DOWN)
        branch[1].send_keys(Keys.ENTER)
        college = self.driver.find_elements_by_xpath('//*[@placeholder="College"][@type="text"]')
        college[1].send_keys("ACCET-Karaikudi")
        college[1].send_keys(Keys.DOWN)
        college[1].send_keys(Keys.ENTER)
        yearOfPassing = self.driver.find_elements_by_xpath('//*[@placeholder="Year Of Passing"][@type="text"]')
        yearOfPassing[1].send_keys("2008")
        yearOfPassing[1].send_keys(Keys.DOWN)
        yearOfPassing[1].send_keys(Keys.ENTER)
        percentageCGPAs = self.driver.find_elements_by_xpath('//*[@ng-model="education.percentage"]')
        percentageCGPA = percentageCGPAs[1]
        percentageCGPA.send_keys("67")
        clickupdate = self.driver.find_element_by_xpath('//*[@ng-click="vm.actionClicked(\'save\')"]')
        clickupdate.click()
        time.sleep(3)
        successmessage = self.driver.find_element_by_xpath('//*[@ng-bind-html="message.text"]').text
        print (successmessage)
        if successmessage == "Education profiles update successfully":
            self.ws.write(7, 7, "Education profiles update successfully", self.__style1)
            self.ws.write(8, 7, successmessage, self.__style3)
            eduData = self.driver.find_elements_by_class_name("row")
            educationData1 = eduData[2].text.index('Education')
            educationData12 = eduData[2].text.index('Personal and More Details')
            education_block = eduData[2].text[educationData1:educationData12]
            print (education_block)
            education_detailsxl = []

            colmn = 8
            for i in range(80, 85):
                if type(self.input1.cell(5, i).value) == float:
                    education_detailsxl.append(str(int((self.input1.cell(5, i).value))))
                else:
                    education_detailsxl.append(str(self.input1.cell(5, i).value))
            for j in education_detailsxl:
                print("#$$$$$$$$$$$$$$$$$$$$")
                print (j.encode('ascii', 'ignore'))
                if re.search(j.encode('ascii', 'ignore'), education_block, re.IGNORECASE):
                    self.ws.write(8, colmn,re.search(j.encode('ascii', 'ignore'), education_block, re.IGNORECASE).group(),
                                  self.__style3)
                else:
                    self.failedCount += 1
                    self.ws.write(8, colmn, "Not Available", self.__style2)
                self.ws.write(7, colmn, j, self.__style1)
                colmn += 1
                # print (j.encode('ascii','ignore'))
            print ("Educc")
            print (education_detailsxl)
        else:
            self.failedCount += 1
            self.ws.write(8, 7, successmessage, self.__style2)
        if self.failedCount>= 1:
            self.ws.write(8, 0, "Test Case Failed", self.__style2)
        else:
            self.ws.write(8, 0, "Test Case Pass", self.__style3)

    def find(self,xpath):
        element = self.driver.find_element_by_xpath(xpath)
        if element:
            return element
        else:
            return False
    def update_personal_details(self):
        editdetails = self.driver.find_element_by_xpath('//*[@title="Edit personal and more details"]')
        editdetails.click()
        time.sleep(6)
        listOfalldata=[]
        listOfmessage = ["Please specify mobile no","Please specify candidate name","Please specify candidate first name","Please specify candidate middle name","Please specify candidate last name","Personal details updated successfully"]
        xpathlist = ['//*[@name="email"]', '//*[@ng-model="vm.data.personalDetails.dob"]', '//*[@name="primaryMobile"]',
                     '//*[@ng-model="vm.data.personalDetails.name"]','//*[@ng-model="vm.data.personalDetails.firstName"]',
                     '//*[@ng-model="vm.data.personalDetails.middleName"]','//*[@ng-model="vm.data.personalDetails.lastName"]',
                     '//*[@ng-model="vm.data.personalDetails.secondaryEmail"]', '//*[@name="alternatePhone"]',
                     '//*[@ng-model="vm.data.personalDetails.address"]','//*[@ng-model="vm.data.personalDetails.usn"]',
                     '//*[@ng-model="vm.data.personalDetails.pan"]','//*[@ng-model="vm.data.personalDetails.passport"]',
                     '//*[@ng-model="vm.data.personalDetails.aadhaarNo"]','//*[@ng-model="vm.data.socialDetails.ssnNo"]',
                     '//*[@ng-model="vm.data.preference.desiredSalaryFrom"]','//*[@ng-model="vm.data.preference.desiredSalaryTo"]',
                     '//*[@ng-model="vm.data.preference.noticePeriod"]','//*[@ng-model="vm.data.personalDetails.totalExperienceInYears"]',
                     '//*[@ng-model="vm.data.personalDetails.totalExperienceInMonths"]','//*[@ng-model="vm.data.preference.fixedCtc"]',
                     '//*[@ng-model="vm.data.personalDetails.currentCtc"]','//*[@ng-model="vm.data.preference.variablePay"]',
                     '//*[@ng-model="vm.data.preference.stocks"]','//*[@ng-model="vm.data.socialDetails.linkedInLink"]',
                     '//*[@ng-model="vm.data.socialDetails.facebookLink"]','//*[@ng-model="vm.data.socialDetails.twitterLink"]',
                     '//*[@ng-model="vm.data.OtherDetails.text1"]','//*[@ng-model="vm.data.OtherDetails.text2"]',
                     '//*[@ng-model="vm.data.OtherDetails.text3"]','//*[@ng-model="vm.data.OtherDetails.text4"]',
                     '//*[@ng-model="vm.data.OtherDetails.text5"]','//*[@ng-model="vm.data.OtherDetails.text6"]',
                     '//*[@ng-model="vm.data.OtherDetails.text7"]','//*[@ng-model="vm.data.OtherDetails.text8"]',
                     '//*[@ng-model="vm.data.OtherDetails.text9"]','//*[@ng-model="vm.data.OtherDetails.text10"]',
                     '//*[@ng-model="vm.data.OtherDetails.text11"]','//*[@ng-model="vm.data.OtherDetails.text12"]',
                     '//*[@ng-model="vm.data.OtherDetails.text13"]','//*[@ng-model="vm.data.OtherDetails.text14"]',
                     '//*[@ng-model="vm.data.OtherDetails.text15"]','//*[@ng-model="vm.data.OtherDetails.textArea1"]',
                     '//*[@ng-model="vm.data.OtherDetails.textArea2"]','//*[@ng-model="vm.data.OtherDetails.textArea3"]',
                     '//*[@ng-model="vm.data.OtherDetails.textArea4"]',
                     '//*[@placeholder="Gender"][@type="text"]', '//*[@placeholder="MaritalStatus"][@type="text"]',
                     '//*[@placeholder="Location"][@type="text"]', '//*[@placeholder="Sensitivity"][@type="text"]',
                     '//*[@placeholder="Nationality"][@type="text"]', '//*[@placeholder="Country"][@type="text"]',
                     '//*[@placeholder="Sourcer"][@type="text"]','//*[@placeholder="Status"][@type="text"]',
                     '//*[@placeholder="Willing to reallocate"][@type="text"]','//*[@placeholder="Expertise"][@type="text"]',
                     '//*[@placeholder="Hierarchy"][@type="text"]','//*[@placeholder="Integer1"][@type="text"]',
                     '//*[@placeholder="Integer2"][@type="text"]','//*[@placeholder="Integer3"][@type="text"]',
                     '//*[@placeholder="Integer4"][@type="text"]','//*[@placeholder="Integer5"][@type="text"]',
                     '//*[@placeholder="Integer6"][@type="text"]','//*[@placeholder="Integer7"][@type="text"]',
                     '//*[@placeholder="Integer8"][@type="text"]','//*[@placeholder="Integer9"][@type="text"]',
                     '//*[@placeholder="Integer10"][@type="text"]','//*[@placeholder="Integer11"][@type="text"]',
                     '//*[@placeholder="Integer12"][@type="text"]','//*[@placeholder="Integer13"][@type="text"]',
                     '//*[@placeholder="Integer14"][@type="text"]','//*[@placeholder="Integer15"][@type="text"]']
        for xpath in xpathlist:
            listOfalldata.append(self.find(xpath))

        message=[]
        row=11
        col=17
        for i in range(4,self.input_file.nrows):
            listt = []
            for j in range(3, self.input_file.ncols):
                listt.append(self.input_file.cell(i, j).value)

            for ind in range(len(listOfalldata)):
                if ind==2:
                    listOfalldata[ind].clear()
                    if i == 4:
                        listOfalldata[ind].send_keys(listt[ind])
                    else:
                        __number = 10
                        ph_num = [int(line.rstrip()) for line in
                                  open('C:\PythonAutomation\CreateCandidateInput\TextFile1.txt')]
                        while True:
                            ph_number = ''.join(["{}".format(randint(1, 9)) for num in range(0, __number)])
                            if ph_number not in ph_num:
                                break
                            with open("C:\PythonAutomation\CreateCandidateInput\TextFile1.txt", 'a+') as e:
                                e.write('\n')
                                e.write(str(ph_number))
                                e.close()
                        listOfalldata[ind].send_keys(ph_number)
                elif ind==1:
                    dateFormat = listt[ind]
                    try:
                        dateFormat = datetime.datetime.fromordinal(
                            datetime.datetime(1900, 1, 1).toordinal() + int(dateFormat) - 2)
                        dateformatfinal = datetime.datetime.strftime(dateFormat, '%d/%m/%Y')
                        listOfalldata[ind].clear()
                        listOfalldata[ind].send_keys(dateformatfinal)
                        listOfalldata[ind].send_keys(Keys.ENTER)
                    except Exception as e:
                        print str(e)
                elif ind==0:
                    email=listt[ind]
                    try:
                        with open("C:\PythonAutomation\CreateCandidateInput\TextFile.txt", 'r') as f:
                            a = f.readline()
                            f.close()
                        splits = email.split('@')
                        splits[0] += 'sprint' + str(a)
                        email = '@'.join(splits)
                        a = int(a) + 1
                        with open("C:\PythonAutomation\CreateCandidateInput\TextFile.txt", 'w') as e:
                            e.write(str(a))
                    except Exception as e:
                        print str(e)
                    listOfalldata[ind].clear()
                    listOfalldata[ind].send_keys(email)
                elif ind in range(3,46):
                    listOfalldata[ind].clear()
                    try:
                        word=str(int(float(listt[ind])))
                    except:
                        word=str(listt[ind])
                    listOfalldata[ind].send_keys(word)
                else:
                    listOfalldata[ind].clear()
                    listOfalldata[ind].clear()
                    listOfalldata[ind].send_keys(listt[ind])
                    listOfalldata[ind].send_keys(Keys.DOWN)
                    listOfalldata[ind].send_keys(Keys.ENTER)
            clickOnUpdate = self.driver.find_element_by_xpath('//*[@ng-click="vm.update();"]')
            clickOnUpdate.click()
            time.sleep(5)
            updatemessage = self.driver.find_element_by_xpath("//*[@ng-bind-html=\"message.text\"]").text
            message.append(updatemessage)
            print (updatemessage)
            if i == self.input_file.nrows-1:
                self.candidate_details_varification(row+1)
                self.status(row+1)
            if updatemessage == listOfmessage[i-4]:
                self.ws.write(row, col, listOfmessage[i-4], self.__style1)
                row+=1
                self.ws.write(row,col,updatemessage,self.__style3)
                self.ws.write(row, 0, "Test Case Passed", self.__style3)
            else:
                self.failedCount += 1
                self.ws.write(row, col, listOfmessage[i-4], self.__style1)
                row+=1
                self.ws.write(row,col,updatemessage,self.__style2)
                self.ws.write(row, 0, "Test Case Failed", self.__style2)

                print ("Paas")
            if i!= self.input_file.nrows-1:
                if updatemessage == "Personal details updated successfully":
                    time.sleep(5)
                    editdetails.click()
                    listOfalldata=[]
                    for xpath in xpathlist:
                        listOfalldata.append(self.find(xpath))
                    time.sleep(11)
            row+=1


        time.sleep(10)
    def status(self,rownum):
        if self.failedCount>= 1:
            self.ws.write(rownum, 0, "Test Case Failed", self.__style2)
        else:
            self.ws.write(rownum, 0, "Test Case Pass", self.__style3)
        self.wb_result.save('C:\PythonAutomation\UpdateCandidateResults\UpdateCandidateResults(' + self.__current_DateTime + ').xls')

    def candidate_details_varification(self, rownum):
            for i in range(10):
                try:
                    self.driver.execute_script("window.scrollTo(0, 1000);")
                    showMore = self.driver.find_element_by_xpath(
                        '//*[@ng-click="vm.data.basic.showMore = !vm.data.basic.showMore"]')
                    showMore.click()
                    break
                except ElementClickInterceptedException as e:
                    print ("Rety in 1 sec")
                    time.sleep(1)
            else:
                raise e
            time.sleep(2)
            # cand_data = self.driver.find_elements_by_tag_name("p")
            cand_data = self.driver.find_elements_by_class_name("candidate-details-container .ng-binding")
            can_dataxl = []
            for i in range(66, 80):
                can_dataxl.append(str(self.input1.cell(5, i).value))
            print type(cand_data)
            print len(cand_data)
            print(can_dataxl)
            can_dataui = []
            for j in cand_data:
                can_dataui.append(j.text)

            can_dataui.pop(7)
            can_dataui.pop()
            can_dataui.pop()
            col_val = 18
            for i in range(len(can_dataxl)):
                if can_dataxl[i] == can_dataui[i]:
                    self.ws.write(rownum, col_val, can_dataui[i], self.__style3)
                else:
                    #
                    self.failedCount += 1
                    self.ws.write(rownum, col_val, can_dataui[i], self.__style5)
                self.ws.write(rownum - 1, col_val, can_dataxl[i], self.__style1)
                col_val += 1
            print ("Can Ui")
            print(can_dataui)
            print ("===============1=================>")
            cand_details_list = []
            for i in cand_data:
                cand_details_list.append(i.text)
            cand_data_personal = self.driver.find_elements_by_class_name("row")
            lst = cand_data_personal[3].text.split('\n')[:-1]
            res_dct = {lst[i]: lst[i + 1] for i in range(0, len(lst), 2)}
            candidate_social_details = cand_data_personal[4].text.split('\n')[1:]
            social_dict = {candidate_social_details[i]: candidate_social_details[i + 1] for i in
                           range(0, len(candidate_social_details), 2)}
            merge_dict = res_dct.copy()
            merge_dict.update(social_dict)
            can_personal_xlheader = ['Integer2', 'Integer3', 'Integer1', 'Integer6', 'Integer7', 'Integer4', 'Integer5',
                                     'Phone No', 'Integer8', 'TextArea3',
                                     'TrueFalse1', 'Text15', 'TrueFalse3', 'TrueFalse2', 'Text10', 'Text11', 'Text12',
                                     'Integer11', 'Integer9', 'DateCustomField1_1',
                                     'Sourcer', 'Integer13', 'Alternate Email', 'TextArea2', 'TextArea1', 'Integer15',
                                     'Text2', 'Text3', 'Notice Period', 'Text1',
                                     'Text6', 'Text7', 'Text4', 'Address', 'Nationality', 'Text8', 'Stocks',
                                     'DateCustomField2', 'Country', 'Text5', 'Prefered Location(s)',
                                     'DateCustomField5', 'Variable Pay', 'Expected Salary', 'Integer10',
                                     'Current Salary', 'Integer12', 'Willing to Relocate', 'Integer14',
                                     'TextArea4', 'Text14', 'Text9', 'Hierarchy', 'Created On', 'Sensitivity', 'Text13',
                                     'Marital Status', 'DateCustomField3', 'Fixed Ctc',
                                     'Expertise', 'DateCustomField4', 'TrueFalse5', 'TrueFalse4', 'FacebookLink',
                                     'LinkedInLink', 'TwitterLink']
            print ("-------Dict-------")
            listPersonal = []
            for i in range(0, 66):
                listPersonal.append(self.input1.cell(5, i).value)
            dictionary = dict(zip(can_personal_xlheader, listPersonal))
            print (dictionary)
            col = 32
            cols = 32
            for keyy, valuee in dictionary.items():
                self.ws.write(0, cols, keyy, self.__style1)
                cols += 1
            for key, value in merge_dict.items():
                for keyy, valuee in dictionary.items():
                    if key == keyy:
                        if value == valuee:
                            self.ws.write(rownum, col, value, self.__style3)
                            print (value, 'Pass')
                        else:
                            self.failedCount += 1
                            self.ws.write(rownum, col, value, self.__style5)
                            print(value, valuee, 'Failed')
                        self.ws.write(rownum - 1, col, valuee, self.__style1)
                        col += 1
            # self.wb_result.save('C:\PythonAutomation\UpdateCandidateResults\UpdateCandidateResults(' + self.__current_DateTime + ').xls')
            # candidate_education_index = cand_data_personal[2].text.index('Education')
            # candidate_education_index_end = cand_data_personal[2].text.index('Personal and More Details')
            # education_block = cand_data_personal[2].text[candidate_education_index:candidate_education_index_end]
            # print (education_block, "#################")
            # education_detailsxl = []

            # colmn = 24
            # for i in range(80, 85):
            #     if type(self.input1.cell(5, i).value) == float:
            #         education_detailsxl.append(str(int((self.input1.cell(5, i).value))))
            #     else:
            #         education_detailsxl.append(str(self.input1.cell(5, i).value))
            # for j in education_detailsxl:
            #     print("#$$$$$$$$$$$$$$$$$$$$")
            #     print (j.encode('ascii', 'ignore'))
            #     if re.search(j.encode('ascii', 'ignore'), education_block, re.IGNORECASE):
            #         self.ws.write(rownum, colmn,
            #                       re.search(j.encode('ascii', 'ignore'), education_block, re.IGNORECASE).group(),
            #                       self.__style3)
            #     else:
            #         self.failedCount += 1
            #         self.ws.write(rownum, colmn, "Not Found", self.__style5)
            #     self.ws.write(rownum - 1, colmn, j, self.__style1)
            #     colmn += 1
            #     # print (j.encode('ascii','ignore'))
            # print ("Educc")
            # print (education_detailsxl)
            # candidate_experince_index = cand_data_personal[2].text.index('Experience')
            # candidate_experince_index_end = cand_data_personal[2].text.index('Education')
            # experience_block = cand_data_personal[2].text[candidate_experince_index:candidate_experince_index_end]
            # exp_detailsxl = []
            # columm = 29
            # for i in range(85, 88):
            #     exp_detailsxl.append(str(self.input1.cell(5, i).value))
            # for j in exp_detailsxl:
            #     if re.search(j.encode('ascii', 'ignore'), experience_block, re.IGNORECASE):
            #         self.ws.write(rownum, columm,
            #                       re.search(j.encode('ascii', 'ignore'), experience_block, re.IGNORECASE).group(),
            #                       self.__style3)
            #     else:
            #         self.failedCount += 1
            #         self.ws.write(rownum, columm, "Not Found", self.__style5)
            #     self.ws.write(rownum - 1, columm, j, self.__style1)
            #     columm += 1
            # print ("Exppp")
            # print (exp_detailsxl)
            # print (experience_block)

        # self.driver.execute_script("arguments[0].value='{}';arguments[1].value='{}';".format(cName, cFirtName), name, firstName)
        # gender.send_keys(Keys.DOWN)
        # gender.send_keys(Keys.ENTER)
        # maritalStatus.send_keys(Keys.DOWN)
        # maritalStatus.send_keys(Keys.ENTER)
        # name.click()
        # time.sleep(4)


if __name__ == '__main__':
    x = UpdateCabdidate()
