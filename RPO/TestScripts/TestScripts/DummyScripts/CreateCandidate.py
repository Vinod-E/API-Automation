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
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.common.exceptions import NoSuchElementException
from random import randint
import sys
reload(sys)
sys.setdefaultencoding('utf8')

class createCandidate():
    now = datetime.datetime.now()
    __current_DateTime = now.strftime("%d-%m-%Y")
    def __init__(self):
        self.failedCount = 0
        self.dictionaryy = {"0":['File Extension Not Allowed'],"1":['Resume Uploaded Successfully','File Extension Not Allowed'],
                            "2":['Resume Uploaded Successfully','Photo Uploaded Successfully','File Extension Not Allowed'],
                            "3":['Resume Uploaded Successfully','Photo Uploaded Successfully','Other Attachments Uploaded Successfully']}
        self.ws = None
        self.test_login_details()
        self.save_result_inExcel(self.__current_DateTime)
        self.input_excel()
        time.sleep(3)
    def test_login_details(self):
        now = datetime.datetime.now()
        __current_DateTime = now.strftime("%d-%m-%Y")
        self.driver = webdriver.Chrome(CONSTANT.CHROME_DRIVER)
        self.driver.implicitly_wait(30)
        self.driver.maximize_window()
        # Getting Url.
        self.driver.get(CONSTANT.RPO_CRPO_AMS_URL_RPOTestone)
        time.sleep(1)
        # Login Applecation For Extract Resume
        self.driver.find_element_by_name("loginName").send_keys(CONSTANT.CRPO_RPO_LOGIN_NAME_MONU)
        self.driver.find_element_by_xpath("//div[3]/section/div[1]/div[2]/form/div[2]/input").send_keys(
            CONSTANT.CRPO_RPO_LOGIN_PASSWORD_MONU)
        self.driver.find_element_by_xpath("//div[3]/section/div[1]/div[2]/form/div[2]/input").send_keys(Keys.ENTER)
        # try:
        #     loginFailure = self.driver.find_element_by_xpath('//*[@id="mainBodyElement"]/div[1]/div/header/div[3]/div/div/div').text
        # except:
        #     loginFailure = False
        #
        # if loginFailure:
        #     print ("Login Failed and Reason" + loginFailure)
        # else:
        #     logSuccess = self.driver.find_element_by_xpath("//span[@class='header_nam ng-binding']").text
        #     print ("%s loged in Successfully"%logSuccess)
        time.sleep(6)

    def input_excel(self):
        list_function = [self.resume_upload, self.photo_upload, self.upload_other_attachments, self.candidate_personal_details,
                         self.identity_details, self.source_details, self.preferences_details, self.profile_details,self.education_details,
                         self.experience_details, self.social_details, self.custom_details]
        __inputSheet = xlrd.open_workbook("C:\PythonAutomation\CreateCandidateInput\CreateCandidateInputAll.xls")
        __sheetName = __inputSheet.sheet_names()
        print (__sheetName)
        self.input_file = __inputSheet.sheet_by_index(0)
        row = 2
        for i in range(1,self.input_file.nrows):
            listt = []
            for j in range(0, self.input_file.ncols):
                listt.append(self.input_file.cell(i, j).value)
            if i <4:
                listOfxlData = filter(None, listt)
            else:
                listOfxlData=listt
            print listOfxlData
            for ind in range(len(listOfxlData)):
                val = i
                if i > 3:
                    val = 4
                if ind<=2:
                    print (self.dictionaryy[str(val-1)][ind],'@@@@@@@@@@@@@')
                    self.ws.write(row, ind+1, listOfxlData[ind], self.__style1)
                    self.ws.write(row-1, ind+4, self.dictionaryy[str(val-1)][ind], self.__style1)
                if ind==12:
                    break
                if ind==3:
                    list_function[ind](listOfxlData[5:23],i)
                elif ind==4:
                    list_function[ind](listOfxlData[23:28])
                elif ind==5:
                    list_function[ind](listOfxlData[28:30])
                elif ind==6:
                    list_function[ind](listOfxlData[30:34])
                elif ind==7:
                    list_function[ind](listOfxlData[34:42])
                elif ind==8:
                    list_function[ind](listOfxlData[42:47],row)
                elif ind==9:
                    list_function[ind](listOfxlData[47:55],row)
                elif ind==10:
                    list_function[ind](listOfxlData[55:58])
                elif ind==11:
                    list_function[ind](listOfxlData[58:97])
                else:
                    list_function[ind](listOfxlData[ind],row,val)
                    if self.failedCount>=1:
                        self.ws.write(row, 0, "Failed", self.__style2)
                    else:
                        self.ws.write(row, 0, "Passed", self.__style3)

            if i > 3:
                self.duplication_check(row)
                time.sleep(7)
                msg = self.create_candidate_click(row)
                if msg == "Candidate created successfully":
                    personalDetailsxl = xlrd.open_workbook(
                        "C:\PythonAutomation\CreateCandidateInput\CreateCandidatePersonalDetailsInput.xls")
                    personalDetailsSheet = personalDetailsxl.sheet_names()
                    print(personalDetailsSheet)
                    self.input1 = personalDetailsxl.sheet_by_index(0)
                    print (self.input1)
                    time.sleep(10)
                    self.candidate_details_varification(row)
                    self.status(row)
                else:
                    self.driver.refresh()
                    time.sleep(6)
            row += 2

    def resume_upload(self, data, rownum,i):
        dict_key=i-1
        __uploadResume = self.driver.find_element_by_xpath("//div[3]/div/create-update-candidate/section/div[1]/div/div[3]/div[1]/div[1]/div/div[1]/upload-file/div/div/input")
        __uploadResume.send_keys(data)
        time.sleep(7)
        message = ""
        try:
            __succMessage = self.driver.find_element_by_xpath('//*[@ng-if="vm.data.attachments.resumeUrl.length"]').text

            if __succMessage == "Resume1.doc":
                message="Resume Uploaded Successfully"
                self.ws.write(rownum, 4, "Resume Uploaded Successfully", self.__style3)

            else:
                message="Wrong File Uploaded"
                self.ws.write(rownum, 4, "Wrong File Uploaded", self.__style3)

        except NoSuchElementException as e:
            message ="File Extension Not Allowed"
            self.ws.write(rownum, 4, "File Extension Not Allowed", self.__style3)
        print (message,self.dictionaryy[str(dict_key)][0])
        if self.dictionaryy[str(dict_key)][0]!=message:
            self.failedCount+=1
        time.sleep(1)
    def photo_upload(self, data,rownum,i):
        dict_key = i - 1
        __uploadPhoto = self.driver.find_element_by_xpath("//div[3]/div/create-update-candidate/section/div[1]/div/div[3]/div[1]/div[2]/div/upload-file/div/div[1]/input")
        __uploadPhoto.send_keys(data)
        time.sleep(7)
        message = ""
        try:
            __uploaSuccMessage = self.driver.find_element_by_xpath('//*[@ng-if="vm.data.attachments.photoUrl.length"]').text
            print (__uploaSuccMessage)

            if __uploaSuccMessage == "Pic1.png (X)":
                message = "Photo Uploaded Successfully"
                self.ws.write(rownum, 5, "Photo Uploaded Successfully", self.__style3)
            else:
                message = "Wrong File Uploaded"
                self.ws.write(rownum, 5, "Wrong File Uploaded", self.__style3)

        except NoSuchElementException as e:
            message = "File Extension Not Allowed"
            self.ws.write(rownum, 5, "File Extension Not Allowed", self.__style3)
        if self.dictionaryy[str(dict_key)][1]!=message:
            self.failedCount+=1

        time.sleep(1)
    def upload_other_attachments(self,data,rownum,i):
        dict_key = i - 1
        __uploadAttach = self.driver.find_element_by_xpath("//div/create-update-candidate/section/div[1]/div/div[3]/div[1]/div[3]/div/upload-file/div/div/input")
        __uploadAttach.send_keys(data)
        time.sleep(7)
        message = ""
        try:
            __uploaSuccessMessage = self.driver.find_element_by_xpath('//*[@ng-if="vm.data.attachments.otherAttachments.length"]').text
            print(__uploaSuccessMessage)
            if not __uploaSuccessMessage:
                message = "File Extension Not Allowed"
            else:
                if __uploaSuccessMessage == "Resume1.doc (X)":
                    message = "Other Attachments Uploaded Successfully"
                    self.ws.write(rownum, 6, "Other Attachments Uploaded Successfully", self.__style3)
                else:
                    message = "Wrong File Uploaded"
                    self.ws.write(rownum, 6, "Wrong File Uploaded", self.__style3)

        except NoSuchElementException as e:
            message = "File Extension Not Allowed"
            self.ws.write(rownum, 6, "File Extension Not Allowed", self.__style3)
        if self.dictionaryy[str(dict_key)][2]!=message:
            self.failedCount+=1
        time.sleep(1)
    def resume_extract(self):
        __clickExtract = self.driver.find_element_by_xpath('//*[@class="btn btn-default blue"]')
        __clickExtract.click()
        __clickOk = self.driver.find_element_by_xpath('//*[@ng-if="!data.isCommentMandatory"]')
        __clickOk.click()
        print ("Resume Extracted Successfully")
    def candidate_skills(self):
        __primarySkills = self.driver.find_element_by_xpath('//*[@class="pull-left ng-binding"]')
        __primarySkills.click()
        selectSkills = Select(self.driver.find_element_by_xpath('//*[@class="ellipsis_text ng-binding ng-scope"]'))
        selectSkills.select_by_value("2495")
    def candidate_personal_details(self,data,i):
        __candidateName = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.personalDetails.name"]')
        print (data)
        __candidateName.send_keys(data[0])
        __firstName = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.personalDetails.firstName"]')
        __firstName.send_keys(data[1])
        __middleName = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.personalDetails.middleName"]')
        __middleName.send_keys(data[2])
        __lastName = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.personalDetails.lastName"]')
        __lastName.send_keys(data[3])
        __primaryEmail = self.driver.find_element_by_name("email")
        email = data[4]
        try:
            with open("C:\PythonAutomation\CreateCandidateInput\TextFile.txt",'r') as f:
                a = f.readline()
                f.close()
            splits = email.split('@')
            splits[0]+='sprint'+str(a)
            email = '@'.join(splits)
            a=int(a)+1
            with open("C:\PythonAutomation\CreateCandidateInput\TextFile.txt",'w') as e:
                e.write(str(a))
        except:
            e
        __primaryEmail.send_keys(email)
        __secondaryEmail = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.personalDetails.secondaryEmail"]')
        __secondaryEmail.send_keys(data[5])
        __dateOfBirth = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.personalDetails.dob"]')
        dateFormat = data[6]
        try:
            dateFormat = datetime.datetime.fromordinal(datetime.datetime(1900, 1, 1).toordinal() + int(dateFormat) - 2)
            dateformatfinal = datetime.datetime.strftime(dateFormat, '%d/%m/%Y')
            __dateOfBirth.send_keys(dateformatfinal)
            __dateOfBirth.send_keys(Keys.ENTER)
        except:
            e
        # print (dateformatfinal)

        __gender = self.driver.find_element_by_xpath('//*[@placeholder="Gender"][@type="text"]')
        __gender.clear()
        __gender.clear()
        __gender.send_keys(data[7])
        __gender.send_keys(Keys.DOWN)
        __gender.send_keys(Keys.ENTER)
        __maritalStatus = self.driver.find_element_by_xpath('//*[@placeholder="Marital Status"][@type="text"]')
        __maritalStatus.clear()
        __maritalStatus.clear()
        __maritalStatus.send_keys(data[8])
        __maritalStatus.send_keys(Keys.DOWN)
        __maritalStatus.send_keys(Keys.ENTER)
        __secondaryPhone = self.driver.find_element_by_name('alternatePhone')
        secondaryPhone = str(data[9])
        __secondaryPhone.send_keys(secondaryPhone)
        __mobileNumber = self.driver.find_element_by_name('Mobile1')

        if i == 5:
            __mobileNumber.send_keys(data[10])
        else:
            __number = 10
            ph_num = [int(line.rstrip()) for line in open('C:\PythonAutomation\CreateCandidateInput\TextFile1.txt')]
            while True:
                ph_number = ''.join(["{}".format(randint(1, 5)) for num in range(0, __number)])
                if ph_number not in ph_num:
                    break
            with open("C:\PythonAutomation\CreateCandidateInput\TextFile1.txt",'a+') as e:
                e.write('\n')
                e.write(str(ph_number))
                e.close()
            __mobileNumber.send_keys(ph_number)
        __location = self.driver.find_element_by_xpath('//*[@placeholder="Location"][@type="text"]')
        __location.send_keys(data[11])
        __location.send_keys(Keys.DOWN)
        __location.send_keys(Keys.ENTER)
        __sensitivity = self.driver.find_element_by_xpath('//*[@placeholder="Sensitivity"][@type="text"]')
        __sensitivity.send_keys(data[12])
        __sensitivity.send_keys(Keys.DOWN)
        __sensitivity.send_keys(Keys.ENTER)
        __nationality = self.driver.find_element_by_xpath('//*[@placeholder="Nationality"][@type="text"]')
        __nationality.send_keys(data[13])
        __nationality.send_keys(Keys.DOWN)
        __nationality.send_keys(Keys.ENTER)
        __country = self.driver.find_element_by_xpath('//*[@placeholder="Country"][@type="text"]')
        __country.send_keys(data[14])
        __country.send_keys(Keys.DOWN)
        __country.send_keys(Keys.ENTER)
        __sourcer = self.driver.find_element_by_xpath('//*[@placeholder="Sourcer"][@type="text"]')
        __sourcer.send_keys(data[15])
        __sourcer.send_keys(Keys.DOWN)
        __sourcer.send_keys(Keys.ENTER)
        __status = self.driver.find_element_by_xpath('//*[@placeholder="Status"][@type="text"]')
        __status.send_keys(data[16])
        __status.send_keys(Keys.DOWN)
        __status.send_keys(Keys.ENTER)
        __address = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.personalDetails.address"]')
        __address.send_keys(data[17])
    def identity_details(self,data):
        __usn = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.personalDetails.usn"]')
        __usn.send_keys(data[0])
        __panNumber = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.personalDetails.pan"]')
        __panNumber.send_keys(data[1])
        __passport = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.personalDetails.passport"]')
        __passport.send_keys(data[2])
        __aadharNumber = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.personalDetails.aadhaarNo"]')
        aadharnumber = str(data[3])
        __aadharNumber.send_keys(aadharnumber)
        __ssnNumber = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.socialDetails.ssnNo"]')
        ssnnumber = str(data[4])
        __ssnNumber.send_keys(ssnnumber)
    def source_details(self,data):
        __typeOfSource = self.driver.find_element_by_xpath('//*[@placeholder="Type Of Source"][@type="text"]')
        __typeOfSource.send_keys(data[0])
        __typeOfSource.send_keys(Keys.DOWN)
        __typeOfSource.send_keys(Keys.ENTER)
        __source = self.driver.find_element_by_xpath('//*[@placeholder="Source"][@type="text"]')
        __source.send_keys(data[1])
        __source.send_keys(Keys.DOWN)
        __source.send_keys(Keys.ENTER)
    def preferences_details(self,data):
        __expectedCtcFrom = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.preference.desiredSalaryFrom"]')
        __expectedCtcFrom.send_keys(str(data[0]))
        __expectedCtcTo = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.preference.desiredSalaryTo"]')
        __expectedCtcTo.send_keys(str(data[1]))
        __noticePeriod = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.preference.noticePeriod"]')
        __noticePeriod.send_keys(str(data[2]))
        __willingToRelocate = self.driver.find_element_by_xpath('//*[@placeholder="Willing to reallocate"][@type="text"]')
        __willingToRelocate.clear()
        __willingToRelocate.clear()
        __willingToRelocate.send_keys(data[3])
        __willingToRelocate.send_keys(Keys.DOWN)
        __willingToRelocate.send_keys(Keys.ENTER)
    def profile_details(self,data):
        __expertise = self.driver.find_element_by_xpath('//*[@placeholder="Expertise"][@type="text"]')
        __expertise.send_keys(data[0])
        __expertise.send_keys(Keys.DOWN)
        __expertise.send_keys(Keys.ENTER)
        __hierarchy = self.driver.find_element_by_xpath('//*[@placeholder="Hierarchy"][@type="text"]')
        __hierarchy.send_keys(data[1])
        __hierarchy.send_keys(Keys.DOWN)
        __hierarchy.send_keys(Keys.ENTER)
        __expInYears = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.personalDetails.totalExperienceInYears"]')
        __expInYears.send_keys(str(data[2]))
        __expInMonths = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.personalDetails.totalExperienceInMonths"]')
        __expInMonths.send_keys(str(data[3]))
        __currentSalary = self.driver.find_element_by_name("CTC")
        __currentSalary.send_keys(str(data[4]))
        __fixedSalary = self.driver.find_element_by_name("fixedCtc")
        __fixedSalary.send_keys(str(data[5]))
        __variablePay = self.driver.find_element_by_name("variablePay")
        __variablePay.send_keys(str(data[6]))
        __stock = self.driver.find_element_by_name("stocks")
        __stock.send_keys(str(data[7]))
    def education_details(self,data,rownum):
        __college = self.driver.find_element_by_xpath('//*[@placeholder="College"][@type="text"]')
        __college.send_keys(data[0])
        __college.send_keys(Keys.DOWN)
        __college.send_keys(Keys.ENTER)
        __degree = self.driver.find_element_by_xpath('//*[@placeholder="Degree"][@type="text"]')
        __degree.send_keys(data[1])
        __degree.send_keys(Keys.DOWN)
        __degree.send_keys(Keys.ENTER)
        __branch = self.driver.find_element_by_xpath('//*[@placeholder="Branch"][@type="text"]')
        __branch.send_keys(data[2])
        __branch.send_keys(Keys.DOWN)
        __branch.send_keys(Keys.ENTER)
        __yearOfPassing = self.driver.find_element_by_xpath('//*[@placeholder="Year Of Passing"][@type="text"]')
        __yearOfPassing.send_keys(int(data[3]))
        __yearOfPassing.send_keys(Keys.DOWN)
        __yearOfPassing.send_keys(Keys.ENTER)
        __percentageCGPA = self.driver.find_element_by_name("myDecimal")
        __percentageCGPA.send_keys(str(data[4]))
        __isFinal = self.driver.find_element_by_xpath('//*[@ng-model="vm.tempObj.educationDetails.isFinal"]')
        __isFinal.click()
        __addEducationDetail = self.driver.find_element_by_xpath('//*[@ng-click="vm.actionClicked(\'addEducationDetails\');"]')
        __addEducationDetail.click()
        time.sleep(1)
        __addEducMessage = self.driver.find_element_by_class_name("growl-container").text
        print (__addEducMessage)
        self.ws.write(rownum, 7, __addEducMessage, self.__style3)
        self.ws.write(rownum-1, 7, "Education Added Successfully", self.__style1)
    def experience_details(self,data,rownum):
        __company = self.driver.find_element_by_xpath('//*[@placeholder="Company"][@type="text"]')
        __company.send_keys(data[0])
        __company.send_keys(Keys.DOWN)
        __company.send_keys(Keys.ENTER)
        __designation = self.driver.find_element_by_xpath('//*[@placeholder="Designation"][@type="text"]')
        __designation.send_keys(data[1])
        __designation.send_keys(Keys.DOWN)
        __designation.send_keys(Keys.ENTER)
        __yearlySalary = self.driver.find_element_by_xpath('//*[@ng-model="vm.tempObj.experienceDetails.yearlySalary"]')
        __yearlySalary.send_keys(str(data[2]))
        __reasonForLeaving = self.driver.find_element_by_xpath('//*[@ng-model="vm.tempObj.experienceDetails.reasonForLeaving"]')
        __reasonForLeaving.send_keys(data[3])
        __expFromMonth = self.driver.find_element_by_xpath('//*[@placeholder="Exp Month From"][@type="text"]')
        __expFromMonth.send_keys(data[4])
        __expFromMonth.send_keys(Keys.DOWN)
        __expFromMonth.send_keys(Keys.ENTER)
        __expFromYear = self.driver.find_element_by_xpath('//*[@placeholder="Exp From Year"][@type="text"]')
        __expFromYear.send_keys(int(data[5]))
        __expFromYear.send_keys(Keys.DOWN)
        __expFromYear.send_keys(Keys.ENTER)
        __expToMonth = self.driver.find_element_by_xpath('//*[@placeholder="Exp Month To"][@type="text"]')
        __expToMonth.send_keys(data[6])
        __expToMonth.send_keys(Keys.DOWN)
        __expToMonth.send_keys(Keys.ENTER)
        __expToYear = self.driver.find_element_by_xpath('//*[@placeholder="Exp To Year"][@type="text"]')
        __expToYear.send_keys(int(data[7]))
        __expToYear.send_keys(Keys.DOWN)
        __expToYear.send_keys(Keys.ENTER)
        self.driver.execute_script("window.scrollTo(0, 1000);")
        time.sleep(4)
        for i in range(10):
            try:
                __isLatest = self.driver.find_element_by_xpath('//*[@ng-model="vm.tempObj.experienceDetails.isLatest"]')
                time.sleep(2)
                __isLatest.click()
                break
            except ElementClickInterceptedException as e:
                print ("Rety in 1 sec")
                time.sleep(1)
        else:
            raise e
        __aadExpDetails = self.driver.find_element_by_xpath('//*[@ng-click="vm.actionClicked(\'addExperienceDetails\');"]')
        __aadExpDetails.click()
        time.sleep(2)
        __addExpMessage = self.driver.find_element_by_class_name("growl-container").text
        print (__addExpMessage)
        self.ws.write(rownum, 8, __addExpMessage, self.__style3)
        self.ws.write(rownum-1, 8, "Experience Added Successfully", self.__style1)
    def social_details(self,data):
        __linkedInLink = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.socialDetails.linkedInLink"]')
        __linkedInLink.send_keys(data[0])
        __facebookLink = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.socialDetails.facebookLink"]')
        __facebookLink.send_keys(data[1])
        __twitterLink = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.socialDetails.twitterLink"]')
        __twitterLink.send_keys(data[2])
    def custom_details(self,data):
        __intergerOne = self.driver.find_element_by_xpath('//*[@placeholder="Integer1"][@type="text"]')
        __intergerOne.send_keys(data[0])
        __intergerOne.send_keys(Keys.DOWN)
        __intergerOne.send_keys(Keys.ENTER)
        __Integertwo = self.driver.find_element_by_xpath('//*[@placeholder="Integer2"][@type="text"]')
        __Integertwo.send_keys(data[1])
        __Integertwo.send_keys(Keys.DOWN)
        __Integertwo.send_keys(Keys.ENTER)
        __integerThree = self.driver.find_element_by_xpath('//*[@placeholder="Integer3"][@type="text"]')
        __integerThree.send_keys(data[2])
        __integerThree.send_keys(Keys.DOWN)
        __integerThree.send_keys(Keys.ENTER)
        __integerFour = self.driver.find_element_by_xpath('//*[@placeholder="Integer4"][@type="text"]')
        __integerFour.send_keys(data[3])
        __integerFour.send_keys(Keys.DOWN)
        __integerFour.send_keys(Keys.ENTER)
        __integerFive = self.driver.find_element_by_xpath('//*[@placeholder="Integer5"][@type="text"]')
        __integerFive.send_keys(data[4])
        __integerFive.send_keys(Keys.DOWN)
        __integerFive.send_keys(Keys.ENTER)
        __integerSix = self.driver.find_element_by_xpath('//*[@placeholder="Integer6"][@type="text"]')
        __integerSix.send_keys(data[5])
        __integerSix.send_keys(Keys.DOWN)
        __integerSix.send_keys(Keys.ENTER)
        __integerSeven = self.driver.find_element_by_xpath('//*[@placeholder="Integer7"][@type="text"]')
        __integerSeven.send_keys(data[6])
        __integerSeven.send_keys(Keys.DOWN)
        __integerSeven.send_keys(Keys.ENTER)
        __integerEight = self.driver.find_element_by_xpath('//*[@placeholder="Integer8"][@type="text"]')
        __integerEight.send_keys(data[7])
        __integerEight.send_keys(Keys.DOWN)
        __integerEight.send_keys(Keys.ENTER)
        __integerNine = self.driver.find_element_by_xpath('//*[@placeholder="Integer9"][@type="text"]')
        __integerNine.send_keys(data[8])
        __integerNine.send_keys(Keys.DOWN)
        __integerNine.send_keys(Keys.ENTER)
        __integerTen = self.driver.find_element_by_xpath('//*[@placeholder="Integer10"][@type="text"]')
        __integerTen.send_keys(data[9])
        __integerTen.send_keys(Keys.DOWN)
        __integerTen.send_keys(Keys.ENTER)
        __integerEleven = self.driver.find_element_by_xpath('//*[@placeholder="Integer11"][@type="text"]')
        __integerEleven.send_keys(data[10])
        __integerEleven.send_keys(Keys.DOWN)
        __integerEleven.send_keys(Keys.ENTER)
        __integerTwelve = self.driver.find_element_by_xpath('//*[@placeholder="Integer12"][@type="text"]')
        __integerTwelve.send_keys(data[11])
        __integerTwelve.send_keys(Keys.DOWN)
        __integerTwelve.send_keys(Keys.ENTER)
        __integerThirteen = self.driver.find_element_by_xpath('//*[@placeholder="Integer13"][@type="text"]')
        __integerThirteen.send_keys(data[12])
        __integerThirteen.send_keys(Keys.DOWN)
        __integerThirteen.send_keys(Keys.ENTER)
        __integerFourteen = self.driver.find_element_by_xpath('//*[@placeholder="Integer14"][@type="text"]')
        __integerFourteen.send_keys(data[13])
        __integerFourteen.send_keys(Keys.DOWN)
        __integerFourteen.send_keys(Keys.ENTER)
        __integerFifteen = self.driver.find_element_by_xpath('//*[@placeholder="Integer15"][@type="text"]')
        __integerFifteen.send_keys(data[14])
        __integerFifteen.send_keys(Keys.DOWN)
        __integerFifteen.send_keys(Keys.ENTER)
        __textOne = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.text1"]')
        __textOne.send_keys(data[15])
        __textTwo = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.text2"]')
        __textTwo.send_keys(data[16])
        __textThree = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.text3"]')
        __textThree.send_keys(data[17])
        __textFour = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.text4"]')
        __textFour.send_keys(data[18])
        __textFive = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.text5"]')
        __textFive.send_keys(data[19])
        __textSix = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.text6"]')
        __textSix.send_keys(data[20])
        __textSeven = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.text7"]')
        __textSeven.send_keys(data[21])
        __textEight = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.text8"]')
        __textEight.send_keys(data[22])
        __textNine = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.text9"]')
        __textNine.send_keys(data[23])
        __textTen = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.text10"]')
        __textTen.send_keys(data[24])
        __textEleven = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.text11"]')
        __textEleven.send_keys(data[25])
        __textTwelve = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.text12"]')
        __textTwelve.send_keys(data[26])
        __textThirteen = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.text13"]')
        __textThirteen.send_keys(data[27])
        __textFourteen = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.text14"]')
        __textFourteen.send_keys(data[28])
        __textFifteen = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.text15"]')
        __textFifteen.send_keys(data[29])
        __dateCustomOne = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.dateCustomField1"]')
        dateCustomOne = data[30]
        dateCustomOne = datetime.datetime.fromordinal(datetime.datetime(1900, 1, 1).toordinal() + int(dateCustomOne) - 2)
        dateCustomOnefinal = datetime.datetime.strftime(dateCustomOne, '%d/%m/%Y')
        __dateCustomOne.send_keys(dateCustomOnefinal)
        __dateCustomOne.send_keys(Keys.ENTER)
        __dateCustomeTwo = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.dateCustomField2"]')
        dateCustomTwo = data[31]
        dateCustomTwo = datetime.datetime.fromordinal(datetime.datetime(1900, 1, 1).toordinal() + int(dateCustomTwo) - 2)
        dateCustomTwofinal = datetime.datetime.strftime(dateCustomTwo, '%d/%m/%Y')
        __dateCustomeTwo.send_keys(dateCustomTwofinal)
        __dateCustomeTwo.send_keys(Keys.ENTER)
        __dateCustomThree = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.dateCustomField3"]')
        dateCustomThree = data[32]
        dateCustomThree = datetime.datetime.fromordinal(datetime.datetime(1900, 1, 1).toordinal() + int(dateCustomThree) - 2)
        dateCustomThreefinal = datetime.datetime.strftime(dateCustomThree, '%d/%m/%Y')
        __dateCustomThree.send_keys(dateCustomThreefinal)
        __dateCustomThree.send_keys(Keys.ENTER)
        __dateCustomFour = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.dateCustomField4"]')
        dateCustomFour = data[33]
        dateCustomFour = datetime.datetime.fromordinal(datetime.datetime(1900, 1, 1).toordinal() + int(dateCustomFour) - 2)
        dateCustomFourfinal = datetime.datetime.strftime(dateCustomFour, '%d/%m/%Y')
        __dateCustomFour.send_keys(dateCustomFourfinal)
        __dateCustomFour.send_keys(Keys.ENTER)
        __dateCustomFive = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.dateCustomField5"]')
        dateCustomFive = data[34]
        dateCustomFive = datetime.datetime.fromordinal(datetime.datetime(1900, 1, 1).toordinal() + int(dateCustomFive) - 2)
        dateCustomFivefinal = datetime.datetime.strftime(dateCustomFive, '%d/%m/%Y')
        __dateCustomFive.send_keys(dateCustomFivefinal)
        __dateCustomFive.send_keys(Keys.ENTER)
        __textAreaOne = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.textArea1"]')
        __textAreaOne.send_keys(data[35])
        __textAreaTwo = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.textArea2"]')
        __textAreaTwo.send_keys(data[36])
        __textAreaThree = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.textArea3"]')
        __textAreaThree.send_keys(data[37])
        __textAreaFour = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.OtherDetails.textArea4"]')
        __textAreaFour.send_keys(data[38])
    def duplication_check(self,rownum):
        __checkDuplicate = self.driver.find_element_by_xpath('//*[@ng-click="vm.actionClicked(\'checkDuplicate\',vm.data);"]')
        __checkDuplicate.click()
        time.sleep(4)
        __duplicationMessage = self.driver.find_element_by_xpath('//*[@ng-bind-html="message.text"]').text
        print (__duplicationMessage)
        self.ws.write(rownum-1, 9, __duplicationMessage, self.__style1)
        self.ws.write(rownum, 9, __duplicationMessage, self.__style3)
        time.sleep(3)
    def create_candidate_click(self,rownum):
        __clickOnCreate = self.driver.find_element_by_xpath('//*[@ng-click="vm.actionClicked(\'create\');"]')
        __clickOnCreate.click()
        time.sleep(2)
        __succMessage = self.driver.find_element_by_xpath('//*[@ng-bind-html="message.text"]').text
        print (__succMessage, '##############')
        self.ws.write(rownum, 10, __succMessage, self.__style3)
        self.ws.write(rownum-1, 10, __succMessage, self.__style1)
        time.sleep(2)
        return __succMessage
    def save_result_inExcel(self, __current_DateTime):
        # Color Coding Code For XLs.
        __style0 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on')
        self.__style1 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        self.__style2 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        self.__style3 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')
        __style4 = xlwt.easyxf('font: name Times New Roman, color-index blue, bold on')
        self.__style5 = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;' 'font: name Times New Roman, color-index red, bold on;')
        # Writing XLs Sheet Columns.
        self.wb_result = xlwt.Workbook()
        self.ws = self.wb_result.add_sheet('Candidate Grid Data Sheet',cell_overwrite_ok=True)
        self.ws.write(0, 0, 'Status', __style0)
        self.ws.write(0, 1, 'InputForResume', __style0)
        self.ws.write(0, 2, 'InputForPhoto', __style0)
        self.ws.write(0, 3, 'InputForOtherAttach', __style0)
        self.ws.write(0, 4, 'UploadResumeStatus', __style0)
        self.ws.write(0, 5, 'UploadPhotoStatus', __style0)
        self.ws.write(0, 6, 'UploadOtherAttachStatus', __style0)
        self.ws.write(0, 7, 'EducationAddedMessage', __style0)
        self.ws.write(0, 8, 'ExperienceAddedMessage', __style0)
        self.ws.write(0, 9, 'DuplicationCheckMessage', __style0)
        self.ws.write(0, 10, 'CreateCandidateStatus', __style0)
        self.ws.write(0, 11, 'CandidateName', __style0)
        self.ws.write(0, 12, 'EmailId', __style0)
        self.ws.write(0, 13, 'MobileNumber', __style0)
        self.ws.write(0, 14, 'Location', __style0)
        self.ws.write(0, 15, 'DOB (DD/MM/YYYY)', __style0)
        self.ws.write(0, 16, 'Total Experience', __style0)
        self.ws.write(0, 17, 'Candidate Source', __style0)
        self.ws.write(0, 18, 'USN', __style0)
        self.ws.write(0, 19, 'Gender', __style0)
        self.ws.write(0, 20, 'PassportNumber', __style0)
        self.ws.write(0, 21, 'AadharNumber', __style0)
        self.ws.write(0, 22, 'PanNumber', __style0)
        self.ws.write(0, 23, 'Candidate Status', __style0)
        self.ws.write(0, 24, 'SSNNumber', __style0)
        self.ws.write(0, 25, 'Degree', __style0)
        self.ws.write(0, 26, 'Percentage/CGPA', __style0)
        self.ws.write(0, 27, 'CollegeName', __style0)
        self.ws.write(0, 28, 'Branch', __style0)
        self.ws.write(0, 29, 'YOP', __style0)
        self.ws.write(0, 30, 'Company', __style0)
        self.ws.write(0, 31, 'Designation', __style0)
        self.ws.write(0, 32, 'Duration', __style0)

        # self.ws.write(0, 30, 'Integer2', __style0)
        # self.ws.write(0, 31, 'Integer3', __style0)
        # self.ws.write(0, 32, 'Integer1', __style0)
        # self.ws.write(0, 33, 'Integer6', __style0)
        # self.ws.write(0, 34, 'Integer7', __style0)
        # self.ws.write(0, 35, 'Integer4', __style0)
        # self.ws.write(0, 36, 'Integer5', __style0)
        # self.ws.write(0, 37, 'Phone No', __style0)
        # self.ws.write(0, 38, 'Integer8', __style0)
        # self.ws.write(0, 39, 'TextArea3', __style0)
        # self.ws.write(0, 40, 'TrueFalse1', __style0)
        # self.ws.write(0, 41, 'Text15', __style0)
        # self.ws.write(0, 42, 'TrueFalse3', __style0)
        # self.ws.write(0, 43, 'TrueFalse2', __style0)
        # self.ws.write(0, 44, 'Text10', __style0)
        # self.ws.write(0, 45, 'Text11', __style0)
        # self.ws.write(0, 46, 'Text12', __style0)
        # self.ws.write(0, 47, 'Integer11', __style0)
        # self.ws.write(0, 48, 'Integer9', __style0)
        # self.ws.write(0, 49, 'DateCustomField1_1', __style0)
        #
        # self.ws.write(0, 50, 'Sourcer', __style0)
        # self.ws.write(0, 51, 'Integer13', __style0)
        # self.ws.write(0, 52, 'Alternate Email', __style0)
        # self.ws.write(0, 53, 'TextArea2', __style0)
        # self.ws.write(0, 54, 'TextArea1', __style0)
        # self.ws.write(0, 55, 'Integer15', __style0)
        # self.ws.write(0, 56, 'Text2', __style0)
        # self.ws.write(0, 57, 'Text3', __style0)
        # self.ws.write(0, 58, 'Notice Period', __style0)
        # self.ws.write(0, 59, 'Text1', __style0)
        # self.ws.write(0, 60, 'Text6', __style0)
        # self.ws.write(0, 61, 'Text7', __style0)
        # self.ws.write(0, 62, 'Text4', __style0)
        # self.ws.write(0, 63, 'Address', __style0)
        # self.ws.write(0, 64, 'Nationality', __style0)
        # self.ws.write(0, 65, 'Text8', __style0)
        # self.ws.write(0, 66, 'Stocks', __style0)
        # self.ws.write(0, 67, 'DateCustomField2', __style0)
        # self.ws.write(0, 68, 'Country', __style0)
        # self.ws.write(0, 69, 'Text5', __style0)
        # self.ws.write(0, 70, 'Prefered Location(s)', __style0)
        # self.ws.write(0, 71, 'DateCustomField5', __style0)
        # self.ws.write(0, 72, 'Variable Pay', __style0)
        # self.ws.write(0, 73, 'Expected Salary', __style0)
        # self.ws.write(0, 74, 'Integer10', __style0)
        # self.ws.write(0, 75, 'Current Salary', __style0)
        # self.ws.write(0, 76, 'Integer12', __style0)
        # self.ws.write(0, 77, 'Willing to Relocate', __style0)
        # self.ws.write(0, 78, 'Integer14', __style0)
        # self.ws.write(0, 79, 'TextArea4', __style0)
        # self.ws.write(0, 80, 'Text14', __style0)
        # self.ws.write(0, 81, 'Text9', __style0)
        # self.ws.write(0, 82, 'Hierarchy', __style0)
        # self.ws.write(0, 83, 'Created On', __style0)
        # self.ws.write(0, 84, 'Sensitivity', __style0)
        # self.ws.write(0, 85, 'Text13', __style0)
        # self.ws.write(0, 86, 'Marital Status', __style0)
        # self.ws.write(0, 87, 'DateCustomField3', __style0)
        # self.ws.write(0, 88, 'Fixed Ctc', __style0)
        # self.ws.write(0, 89, 'Expertise', __style0)
        # self.ws.write(0, 90, 'DateCustomField4', __style0)

        # self.ws.write(0, 91, 'TrueFalse5', __style0)
        # self.ws.write(0, 92, 'TrueFalse4', __style0)
        # self.ws.write(0, 93, 'LinkedInLink', __style0)
        # self.ws.write(0, 94, 'FacebookLink', __style0)
        # self.ws.write(0, 95, 'TwitterLink', __style0)
        # self.wb_result.save('C:\PythonAutomation\CreateCandidateResults\CreateCandidateResults(' + __current_DateTime + ').xls')
    def candidate_details_varification(self,rownum):
        for i in range(10):
            try:
                self.driver.execute_script("window.scrollTo(0, 1000);")
                showMore = self.driver.find_element_by_xpath('//*[@ng-click="vm.data.basic.showMore = !vm.data.basic.showMore"]')
                showMore.click()
                break
            except ElementClickInterceptedException as e:
                print ("Rety in 1 sec")
                time.sleep(1)
        else:
            raise e
        time.sleep(2)
        #cand_data = self.driver.find_elements_by_tag_name("p")
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
        col_val=11
        for i in range(len(can_dataxl)):
                if can_dataxl[i]==can_dataui[i]:
                    self.ws.write(rownum, col_val, can_dataui[i], self.__style3)
                else:
                    #
                    self.failedCount+=1
                    self.ws.write(rownum, col_val, can_dataui[i], self.__style5)
                self.ws.write(rownum-1, col_val, can_dataxl[i], self.__style1)
                col_val+=1
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
        social_dict = {candidate_social_details[i]: candidate_social_details[i + 1] for i in range(0, len(candidate_social_details), 2)}
        merge_dict = res_dct.copy()
        merge_dict.update(social_dict)
        can_personal_xlheader = ['Integer2', 'Integer3', 'Integer1', 'Integer6', 'Integer7', 'Integer4', 'Integer5', 'Phone No', 'Integer8', 'TextArea3',
                                 'TrueFalse1', 'Text15', 'TrueFalse3', 'TrueFalse2', 'Text10', 'Text11', 'Text12', 'Integer11', 'Integer9', 'DateCustomField1_1',
                                 'Sourcer', 'Integer13', 'Alternate Email', 'TextArea2', 'TextArea1', 'Integer15', 'Text2', 'Text3', 'Notice Period', 'Text1',
                                 'Text6', 'Text7', 'Text4', 'Address', 'Nationality', 'Text8', 'Stocks', 'DateCustomField2', 'Country', 'Text5', 'Prefered Location(s)',
                                 'DateCustomField5', 'Variable Pay', 'Expected Salary', 'Integer10', 'Current Salary', 'Integer12', 'Willing to Relocate', 'Integer14',
                                 'TextArea4', 'Text14', 'Text9', 'Hierarchy', 'Created On', 'Sensitivity', 'Text13', 'Marital Status', 'DateCustomField3', 'Fixed Ctc',
                                 'Expertise', 'DateCustomField4', 'TrueFalse5', 'TrueFalse4', 'FacebookLink', 'LinkedInLink', 'TwitterLink']
        print ("-------Dict-------")
        listPersonal = []
        for i in range(0, 66):
            listPersonal.append(self.input1.cell(5, i).value)
        dictionary = dict(zip(can_personal_xlheader, listPersonal))
        print (dictionary)
        col = 33
        cols=33
        for keyy, valuee in dictionary.items():
            self.ws.write(0, cols, keyy, self.__style1)
            cols+=1
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
                    self.ws.write(rownum-1, col, valuee, self.__style1)
                    col += 1
        candidate_education_index = cand_data_personal[2].text.index('Education')
        candidate_education_index_end = cand_data_personal[2].text.index('Personal and More Details')
        education_block = cand_data_personal[2].text[candidate_education_index:candidate_education_index_end]
        print (education_block, "#################")
        education_detailsxl = []

        colmn = 25
        for i in range(80, 85):
            if type(self.input1.cell(5, i).value) == float:
                education_detailsxl.append(str(int((self.input1.cell(5, i).value))))
            else:
                education_detailsxl.append(str(self.input1.cell(5, i).value))
        for j in education_detailsxl:
            print("#$$$$$$$$$$$$$$$$$$$$")
            print (j.encode('ascii', 'ignore'))
            if re.search(j.encode('ascii','ignore'),education_block,re.IGNORECASE):
                self.ws.write(rownum, colmn, re.search(j.encode('ascii','ignore'),education_block,re.IGNORECASE).group(), self.__style3)
            else:
                self.failedCount += 1
                self.ws.write(rownum, colmn, "Not Found", self.__style5)
            self.ws.write(rownum-1, colmn, j, self.__style1)
            colmn += 1
            #print (j.encode('ascii','ignore'))
        print ("Educc")
        print (education_detailsxl)
        candidate_experince_index = cand_data_personal[2].text.index('Experience')
        candidate_experince_index_end = cand_data_personal[2].text.index('Education')
        experience_block = cand_data_personal[2].text[candidate_experince_index:candidate_experince_index_end]
        exp_detailsxl = []
        columm = 30
        for i in range(85, 88):
            exp_detailsxl.append(str(self.input1.cell(5, i).value))
        for j in exp_detailsxl:
            if re.search(j.encode('ascii','ignore'),experience_block,re.IGNORECASE):
                self.ws.write(rownum, columm, re.search(j.encode('ascii','ignore'),experience_block,re.IGNORECASE).group(),self.__style3)
            else:
                self.failedCount += 1
                self.ws.write(rownum, columm, "Not Found", self.__style5)
            self.ws.write(rownum-1, columm, j, self.__style1)
            columm += 1
        print ("Exppp")
        print (exp_detailsxl)
        print (experience_block)
        #self.wb_result.save('C:\PythonAutomation\CreateCandidateResults\CreateCandidateResults(' + self.__current_DateTime + ').xls')
    def status(self,rownum):
        if self.failedCount>= 1:
            self.ws.write(rownum, 0, "Failed", self.__style2)
        else:
            self.ws.write(rownum, 0, "Passed", self.__style3)
        self.wb_result.save('C:\PythonAutomation\CreateCandidateResults\CreateCandidateResults(' + self.__current_DateTime + ').xls')
if __name__ == '__main__':
    x = createCandidate()
