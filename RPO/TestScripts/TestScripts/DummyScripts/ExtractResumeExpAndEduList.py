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


class extractResume(unittest.TestCase):
    #Methods for Comparing Expexted Education and Exp. details from Extracted.
    def compare_lists(self, extracted, expected):
        extracted = [re.sub('[^a-z0-9]', '', x.lower()) for x in extracted]
        expected = [re.sub('[^a-z0-9]', '', x.lower()) for x in expected]
        for item in expected:
            if item not in extracted:
                return False
        return True

    def test_extractResume(self):
        now = datetime.datetime.now()
        __current_DateTime = now.strftime("%d-%m-%Y")
        print __current_DateTime
        print type(__current_DateTime)
        #Color Coding Code For XLs.
        __style0 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on')
        __style1 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        __style2 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        __style3 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')
        __style4 = xlwt.easyxf('font: name Times New Roman, color-index blue, bold on')
        #Writing XLs Sheet Columns.
        wb_result = xlwt.Workbook()
        ws = wb_result.add_sheet('Extract Resume Result')
        ws.write(0, 0, 'Resume File Name', __style4)
        ws.write(0, 1, 'Extracted Primary Skills', __style4)
        ws.write(0, 2, 'Extracted Secondary Skills', __style4)
        #ws.write(0, 3, 'Not Extracted Skills', __style4)
        ws.write(0, 4, 'Expected Candidate Name', __style4)
        ws.write(0, 5, 'Extracted Candidate Name', __style4)
        ws.write(0, 6, 'Expected Candidate Primary Email', __style4)
        ws.write(0, 7, 'Extracted Candidate Primary Email', __style4)
        ws.write(0, 8, 'Expected Candidate Secondary Email', __style4)
        ws.write(0, 9, 'Extracted Candidate Secondary Email', __style4)
        ws.write(0, 10, 'Expected Candidate DOB (DD/MM/YYYY)', __style4)
        ws.write(0, 11, 'Extracted Candidate DOB (DD/MM/YYYY)', __style4)
        ws.write(0, 12, 'Expected Candidate Gender', __style4)
        ws.write(0, 13, 'Extracted Candidate Gender', __style4)
        ws.write(0, 14, 'Expected Candidate Location', __style4)
        ws.write(0, 15, 'Extracted Candidate Location', __style4)
        ws.write(0, 16, 'Expected Candidate Mobile Number', __style4)
        ws.write(0, 17, 'Extracted Candidate Mobile Number', __style4)
        ws.write(0, 18, 'Expected Candidate Phone Number', __style4)
        ws.write(0, 19, 'Extracted Candidate Phone Number', __style4)
        ws.write(0, 20, 'Expected College List', __style4)
        ws.write(0, 21, 'Extracted College List', __style4)
        ws.write(0, 22, 'Expected Degree List', __style4)
        ws.write(0, 23, 'Extracted Degree List', __style4)
        ws.write(0, 24, 'Expected Branch List', __style4)
        ws.write(0, 25, 'Extracted Branch List', __style4)
        ws.write(0, 26, 'Expected YOP List', __style4)
        ws.write(0, 27, 'Extracted YOP List', __style4)
        ws.write(0, 28, 'Expected CGPAorPercentage List', __style4)
        ws.write(0, 29, 'Extracted CGPAorPercentage List', __style4)
        ws.write(0, 30, 'Expected IsFinal List', __style4)
        ws.write(0, 31, 'Extracted IsFinal List', __style4)
        #ws.write(0, 32, 'Matched Skills (Expected vs Extracted)', __style4)
        ws.write(0, 33, 'Expected Candidate PanNo', __style4)
        ws.write(0, 34, 'Extracted Candidate PanNo', __style4)
        ws.write(0, 35, 'Expected Candidate PassportNo', __style4)
        ws.write(0, 36, 'Extracted Candidate PassportNo', __style4)
        ws.write(0, 37, 'Expected CompanyName List', __style4)
        ws.write(0, 38, 'Extracted CompanyName List', __style4)
        ws.write(0, 39, 'Expected Designation List', __style4)
        ws.write(0, 40, 'Extracted Designation List', __style4)
        ws.write(0, 41, 'Expected ExpFrom List', __style4)
        ws.write(0, 42, 'Extracted ExpFrom List', __style4)
        ws.write(0, 43, 'Expected ExpTo List', __style4)
        ws.write(0, 44, 'Extracted ExpTo List', __style4)
        ws.write(0, 45, 'Expected Salary List', __style0)
        ws.write(0, 46, 'Extracted Salary List', __style4)
        ws.write(0, 47, 'Expected IsLatest List', __style4)
        ws.write(0, 48, 'Extracted IsLatest List', __style4)
        ws.write(0, 49, 'Expected Candidate TotalExpInYears', __style4)
        ws.write(0, 50, 'Extracted Candidate TotalExpInYears', __style4)
        ws.write(0, 51, 'Expected Candidate TatalExpInMonths', __style4)
        ws.write(0, 52, 'Extracted Candidate TatalExpInMonths', __style4)
        #Initiating Chrome Browser.
        self.driver = webdriver.Chrome(CONSTANT.CHROME_DRIVER)
        self.driver.implicitly_wait(30)
        self.driver.maximize_window()
        #Getting Url.
        self.driver.get(CONSTANT.RPO_CRPO_AMS_URL_RPOTestone)
        time.sleep(1)
        #Login Applecation For Extract Resume
        self.driver.find_element_by_name("loginName").send_keys(CONSTANT.CRPO_RPO_LOGIN_NAME_Admin)
        self.driver.find_element_by_xpath("//div[3]/section/div[1]/div[2]/form/div[2]/input").send_keys(
            CONSTANT.CRPO_RPO_LOGIN_PASSWORD_Admin)
        self.driver.find_element_by_xpath("//div[3]/section/div[1]/div[2]/form/div[2]/input").send_keys(Keys.ENTER)
        #Reading Expected Results Sheet.
        #wb = xlrd.open_workbook("/home/sanjeev/TestScripts/TestScripts/ExtractResuWithExpAndEduListInput.xls")
        wb = xlrd.open_workbook("C:\PythonAutomation\TestScripts1\TestScripts\Input_Data\ExtractResumeSheetListOfEduAndExp.xls")
        sheetname = wb.sheet_names()
        print(sheetname)
        sh1 = wb.sheet_by_index(0)  # add login details
        cnt = 1
        while cnt < sh1.nrows:
            rownum = cnt
            rows = sh1.row_values(rownum)
            __attach_Resume = self.driver.find_element_by_xpath(
                "//div[3]/div/create-update-candidate/section/div[1]/div/div[3]/div[1]/div[1]/div/div[1]/upload-file/div/div/input")
            __attach_Resume.send_keys(CONSTANT.RESUME_FILE_PATH + rows[0])
            time.sleep(8)
            #Getting Resume File Name.
            __resume_File_Name = self.driver.find_element_by_xpath('//*[@ng-if="vm.data.attachments.resumeUrl.length"]').text
            #self.driver.implicitly_wait(40)
            print __resume_File_Name
            ws.write(rownum, 0, __resume_File_Name, __style1)
            #Clicking on Extract Resume Button.
            self.driver.implicitly_wait(40)
            __extract_Resume = self.driver.find_element_by_xpath('//*[@ng-click="vm.actionClicked(\'extract\');"]')
            self.driver.implicitly_wait(40)
            time.sleep(5)
            __extract_Resume.click()
            __click_Onextract = self.driver.find_element_by_xpath('//*[@ng-click="data.result=true;$hide();"]')
            __click_Onextract.click()
            self.driver.implicitly_wait(40)
            time.sleep(31)
            #Expected Results Sheet Column Name and Index.
            totalSkills = rows[1]
            name = rows[2]
            email = rows[3]
            altemail = rows[4]
            dob = rows[5]
            dob = dob.strip("'")
            gender = rows[6]
            location = rows[7]
            mobile = rows[8]
            phone = rows[9]
            experienceYear = rows[10]
            experienceMonths = rows[11]
            college = rows[12]
            degree = rows[13]
            branch = rows[14]
            yop = rows[15]
            cgpaorpercentage = rows[16]
            isFinal = rows[17]
            panNumber = rows[24]
            passNo = rows[25]
            compName = rows[18]
            designation = rows[19]
            fromYear = rows[20]
            toYear = rows[21]
            sal = rows[22]
            islatest = rows[23]
            #Start writing and validating Regular Details Of Candidates in Xls.
            time.sleep(1)
            #Writing Expected Data In Result Sheet.
            ws.write(rownum, 4, name, __style0)
            ws.write(rownum, 6, email, __style0)
            ws.write(rownum, 8, altemail, __style0)
            ws.write(rownum, 10, dob, __style0)
            ws.write(rownum, 12, gender, __style0)
            ws.write(rownum, 14, location, __style0)
            ws.write(rownum, 16, mobile, __style0)
            ws.write(rownum, 18, phone, __style0)
            ws.write(rownum, 49, experienceYear, __style0)
            ws.write(rownum, 51, experienceMonths, __style0)
            ws.write(rownum, 20, college, __style0)
            ws.write(rownum, 22, degree, __style0)
            ws.write(rownum, 24, branch, __style0)
            ws.write(rownum, 26, yop, __style0)
            ws.write(rownum, 28, cgpaorpercentage, __style0)
            ws.write(rownum, 30, isFinal, __style0)
            ws.write(rownum, 33, panNumber, __style0)
            ws.write(rownum, 35, passNo, __style0)
            ws.write(rownum, 37, compName, __style0)
            ws.write(rownum, 39, designation, __style0)
            ws.write(rownum, 41, fromYear, __style0)
            ws.write(rownum, 43, toYear, __style0)
            ws.write(rownum, 45, sal, __style0)
            ws.write(rownum, 47, islatest, __style0)
            # Start writing and validating Regular Details Of Candidates in Xls.
            __candidate_Name = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.personalDetails.name"]').get_attribute("value")
            if not __candidate_Name:
                print("Candidate name is not extracted")
                ws.write(rownum, 5, __candidate_Name, __style1)
            else:
                if str(__candidate_Name).lower() == str(name).lower():
                    print("Name extracted as expected")
                    ws.write(rownum, 5, __candidate_Name, __style3)
                else:
                    print('Name "%s"' % str(__candidate_Name) + ' is not extracted as expected "%s"' % str(name))
                    ws.write(rownum, 5, __candidate_Name, __style2)
            __email = self.driver.find_element_by_name('email').get_attribute("value")
            if not __email:
                print("Candidate Email Id is not extracted")
                ws.write(rownum, 7, __email, __style1)
            else:
                if str(__email).lower() == str(email).lower():
                    print("Candidate Email Id extracted as expected")
                    ws.write(rownum, 7, __email, __style3)
                else:
                    print("Candidate Email Id '%s'" % str(__email) + " is not extracted as expected '%s'" % str(email))
                    ws.write(rownum, 7, __email, __style2)
            __alt_Email = self.driver.find_element_by_name('email2').get_attribute("value")
            if not __alt_Email:
                print("Candidate alternate Email Id is not extracted")
                ws.write(rownum, 9, __alt_Email, __style1)
            else:
                if str(__alt_Email).lower() == str(altemail).lower():
                    print("Candidate alternate Email Id extracted as expected")
                    ws.write(rownum, 9, __alt_Email, __style3)
                else:
                    print('Candidate alternate Email Id "%s"' % str(__alt_Email) + ' is not extracted as expected "%s"' % str(altemail))
                    ws.write(rownum, 9, __alt_Email, __style2)
            __dob = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.personalDetails.dob"]').get_attribute("value")
            if not __dob:
                print("DOB is not extracted")
                ws.write(rownum, 11, __dob, __style1)
            else:
                if __dob == dob:
                    print("DOB is extracted as expected")
                    ws.write(rownum, 11, __dob, __style3)
                else:
                    print('DOB "%s"' % __dob + ' is not extracted as expected "%s"' % dob)
                    ws.write(rownum, 11, __dob, __style2)

            __gender = self.driver.find_element_by_xpath('//*[@placeholder="Gender"][@type="text"]').get_attribute("value")
            if not __gender:
                print("Gender is not extracted")
                ws.write(rownum, 13, __gender, __style1)
            else:
                if __gender.lower() == gender.lower():
                    print("Gender is extracted as expected")
                    ws.write(rownum, 13, __gender, __style3)
                else:
                    print('Gender "%s"' % __gender + ' is not extracted as expected "%s"' % gender)
                    ws.write(rownum, 13, __gender, __style2)
            __location = self.driver.find_element_by_xpath('//*[@placeholder="Location"][@type="text"]').get_attribute("value")
            if not __location:
                print("Location is not extracted")
                ws.write(rownum, 15, __location, __style1)
            else:
                if __location.lower() == location.lower():
                    print("Location is extracted as expected")
                    ws.write(rownum, 15, __location, __style3)
                else:
                    print('Location "%s"' % (str(__location)) + ' is not extracted as expected "%s"' % (str(location)))
                    ws.write(rownum, 15, __location, __style2)

            __mobile = self.driver.find_element_by_name('Mobile1').get_attribute("value")
            if not __mobile:
                print("Mobile number is not extracted")
                ws.write(rownum, 17, __mobile, __style1)
            else:
                if int(__mobile) == int(mobile):
                    print("Mobile number extracted as expected")
                    ws.write(rownum, 17, __mobile, __style3)
                else:
                    print(
                        'Mobile number "%s"' % str(__mobile) + ' is not extracted as expected "%s"' % str(int(mobile)))
                    ws.write(rownum, 17, __mobile, __style2)
            __phone = self.driver.find_element_by_name("alternatePhone").get_attribute("value")
            if not __phone:
                print("Phone number is not extracted")
                ws.write(rownum, 19, __phone, __style1)
            else:
                if int(__phone) == int(phone):
                    print("Phone number extracted as expected")
                    ws.write(rownum, 19, __phone, __style3)
                else:
                    print(
                        'Phonr number "%s"' % (str(__phone)) + ' is not extracted as expected "%s"' % (
                            str(int(phone))))
                    ws.write(rownum, 19, __phone, __style2)
            #Checking Extracted Pan And Passpord Number.
            #time.sleep(2)
            __panNo = self.driver.find_element_by_name("pan").get_attribute("value")
            __panNo = str(__panNo).upper()
            if not __panNo:
                print("PanNo is not extracted")
                ws.write(rownum, 34, __panNo, __style1)
            else:
                if __panNo == panNumber:
                    print("PanNo is extracted as expected")
                    ws.write(rownum, 34, __panNo, __style3)
                else:
                    print('PanNo "%s"' % __panNo + ' is not extracted as expected "%s"' % panNumber)
                    ws.write(rownum, 34, __panNo, __style2)
            __passportNo = self.driver.find_element_by_name("passport").get_attribute("value")
            __passportNo = str(__passportNo).upper()
            if not __passportNo:
                print("PassportNo is not extracted")
                ws.write(rownum, 36, __passportNo, __style1)
            else:
                if __passportNo == passNo:
                    print("PassportNo is extracted as expected")
                    ws.write(rownum, 36, __passportNo, __style3)
                else:
                    print('PassportNo "%s"' % __passportNo + ' is not extracted as expected "%s"' % str(passNo))
                    ws.write(rownum, 36, __passportNo, __style2)
            #Getting Candidate Total Exp. In Years And Months
            time.sleep(1)
            __totalExpInYears = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.personalDetails.totalExperienceInYears"]').get_attribute("value")
            if not __totalExpInYears:
                print("Total Exp In Years is not extracted")
                ws.write(rownum, 50, __totalExpInYears, __style1)
            else:
                if int(__totalExpInYears) == int(experienceYear):
                    print("Total Exp In Years extracted as expected")
                    ws.write(rownum, 50, __totalExpInYears, __style3)
                else:
                    print('Total Exp In Years "%s"' % str(__totalExpInYears) + ' is not extracted as expected "%s"' % str(int(experienceYear)))
                    ws.write(rownum, 50, __totalExpInYears, __style2)
            __totalExpInMonths = self.driver.find_element_by_xpath('//*[@ng-model="vm.data.personalDetails.totalExperienceInMonths"]').get_attribute("value")
            if not __totalExpInMonths:
                print("Total Exp In Months is not extracted")
                ws.write(rownum, 52, __totalExpInMonths, __style1)
            else:
                if int(__totalExpInMonths) == int(experienceMonths):
                    print("Total Exp In Months extracted as expected")
                    ws.write(rownum, 52, __totalExpInMonths, __style3)
                else:
                    print('Total Exp In Months "%s"' % str(__totalExpInMonths) + ' is not extracted as expected "%s"' % str(int(experienceMonths)))
                    ws.write(rownum, 52, __totalExpInMonths, __style2)

            #Logic For List Of Extracted Education Details.
            b = self.driver.find_elements_by_xpath('//*[@ng-if="!education.isEditEducation"]')
            print b
            total = len(b)
            print total
            self.all = []
            self.j = 0
            self.collegeName = []
            self.degree = []
            self.branch = []
            self.yearOfPassing = []
            self.percentageCgpa = []
            self.isFinal = []
            self.isActions = []
            for i in b:
                if self.j == 0:
                    self.j = self.j+1
                    self.collegeName.append(i.text)
                elif self.j == 1:
                    self.j = self.j+1
                    self.degree.append(i.text)
                elif self.j == 2:
                    self.j = self.j+1
                    self.branch.append(i.text)
                elif self.j == 3:
                    self.j = self.j+1
                    self.yearOfPassing.append(i.text)
                elif self.j == 4:
                    self.j = self.j+1
                    self.percentageCgpa.append(i.text)
                elif self.j == 5:
                    self.j = self.j+1
                    self.isFinal.append(i.text)
                elif self.j == 6:
                    self.j = 0
                    self.isActions.append(i.text)
            #Getting List Of Colleges.
            if not self.collegeName:
                print("College Name Is Not Extracted")
                ws.write(rownum, 21, ','.join(self.collegeName), __style1)
            else:
                if self.compare_lists(self.collegeName, college.split(',')):
                    print("College Name extracted as expected")
                    print self.collegeName
                    ws.write(rownum, 21, ','.join(self.collegeName), __style3)
                else:
                    print('College Name "%s"' % ','.join(self.collegeName) + ' is not extracted as expected "%s"' % str(college))
                    ws.write(rownum, 21, ','.join(self.collegeName), __style2)
            # Getting List Of Degrees.
            if not self.degree:
                print("Degrees Is Not Extracted")
                ws.write(rownum, 23, ','.join(self.degree), __style1)
            else:
                if self.compare_lists(self.degree, degree.split(',')):
                    print("degrees extracted as expected")
                    print self.degree
                    ws.write(rownum, 23, ','.join(self.degree), __style3)
                else:
                    print('Degrees "%s"' % ','.join(self.degree) + ' is not extracted as expected "%s"' % str(degree))
                    ws.write(rownum, 23, ','.join(self.degree), __style2)
            # Getting List Of Branches.
            if not self.branch:
                print("Branches Is Not Extracted")
                ws.write(rownum, 25, ','.join(self.branch), __style1)
            else:
                if self.compare_lists(self.branch, branch.split(',')):
                    print("Branches Extracted as Expected")
                    print self.branch
                    ws.write(rownum, 25, ','.join(self.branch), __style3)
                else:
                    print('Branches "%s"' % ','.join(self.branch) + ' is not extracted as expected "%s"' % str(branch))
                    ws.write(rownum, 25, ','.join(self.branch), __style2)
            # Getting List Of YesrOfPassing.
            if not self.yearOfPassing:
                print("YOP Is Not Extracted")
                ws.write(rownum, 27, ','.join(self.yearOfPassing), __style1)
            else:
                if self.compare_lists(self.yearOfPassing, (str(yop)).split(',')):
                    print("YOP extracted as expected")
                    print self.yearOfPassing
                    ws.write(rownum, 27, ','.join(self.yearOfPassing), __style3)
                else:
                    print('YOP "%s"' % ','.join(self.yearOfPassing) + ' is not extracted as expected "%s"' % str(yop))
                    ws.write(rownum, 27, ','.join(self.yearOfPassing), __style2)
            # Getting List Of PercentageCGPA.
            if not self.percentageCgpa:
                print("PercentageCGPA Is Not Extracted")
                ws.write(rownum, 29, ','.join(self.percentageCgpa), __style1)
            else:
                if self.compare_lists(self.percentageCgpa, (str(cgpaorpercentage)).split(',')):
                    print("PercentageCGPA extracted as expected")
                    print self.percentageCgpa
                    ws.write(rownum, 29, ','.join(self.percentageCgpa), __style3)
                else:
                    print('PercentageCGPA "%s"' % ','.join(self.percentageCgpa) + ' is not extracted as expected "%s"' % str(cgpaorpercentage))
                    ws.write(rownum, 29, ','.join(self.percentageCgpa), __style2)
            # Getting List Of IsLatestDegree.
            if not self.isFinal:
                print("IsFinal Is Not Extracted")
                ws.write(rownum, 31, ','.join(self.isFinal), __style1)
            else:
                if self.compare_lists(self.isFinal, isFinal.split(',')):
                    print("IsFinal extracted as expected")
                    print self.isFinal
                    ws.write(rownum, 31, ','.join(self.isFinal), __style3)
                else:
                    print('IsFinal "%s"' % ','.join(self.isFinal) + ' is not extracted as expected "%s"' % str(isFinal))
                    ws.write(rownum, 31, ','.join(self.isFinal), __style2)

            # Logic For List Of Extracted Exp. Details.
            a = self.driver.find_elements_by_xpath('//*[@ng-if="!experience.isEditExperience"]')
            print a
            total =  len(a)
            print total
            self.all  = []
            self.j = 0
            self.companyName = []
            self.desigNation = []
            self.expFrom = []
            self.expTo = []
            self.salary = []
            self.reasonForLeaving = []
            self.isLatest = []
            self.isActions = []

            for i in a:
                if self.j==0:
                    self.j = self.j+1
                    self.companyName.append(i.text)

                elif self.j==1:
                    self.j = self.j+1
                    self.desigNation.append(i.text)

                elif self.j==2:
                    self.j = self.j+1
                    self.expFrom.append(i.text)

                elif self.j==3:
                    self.j = self.j+1
                    self.expTo.append(i.text)

                elif self.j==4:
                    self.j = self.j+1
                    self.salary.append(i.text)

                elif self.j==5:
                    self.j = self.j+1
                    self.reasonForLeaving.append(i.text)

                elif self.j==6:
                    self.j = self.j+1
                    self.isLatest.append(i.text)

                elif self.j==7:
                    self.j = 0
                    self.isActions.append(i.text)
            # # Getting List Of Companies.
            # #compName = StringIO.StringIO()
            # #self.compName = compName
            # print compName
            # #self.compName = compName.append(compName.text)
            # # print self.compName
            # #print compName
            # #ws.write(rownum, 21, self.compName, __style1)
            if not self.companyName:
                print("Company Name Is Not Extracted")
                ws.write(rownum, 38, ','.join(self.companyName), __style1)
            else:
                if self.compare_lists(self.companyName, compName.split(',')):
                    print("Company Name extracted as expected")
                    print self.companyName
                    ws.write(rownum, 38, ','.join(self.companyName), __style3)
                else:
                    print('Company Name "%s"' % ','.join(self.companyName) + ' is not extracted as expected "%s"' % str(compName))
                    ws.write(rownum, 38, ','.join(self.companyName), __style2)
            # Getting List Of Designations.
            if not self.desigNation:
                print("Designations Is Not Extracted")
                ws.write(rownum, 40, ','.join(self.desigNation), __style1)
            else:
                if self.compare_lists(self.desigNation, designation.split(',')):
                    print("Designations extracted as expected")
                    print self.desigNation
                    ws.write(rownum, 40, ','.join(self.desigNation), __style3)
                else:
                    print('Designations "%s"' % ','.join(self.desigNation) + ' is not extracted as expected "%s"' % str(designation))
                    ws.write(rownum, 40, ','.join(self.desigNation), __style2)
            # Getting List Of ExpFrom.
            if not self.expFrom:
                print("Exp From Is Not Extracted")
                ws.write(rownum, 42, ','.join(self.expFrom), __style1)
            else:
                if self.compare_lists(self.expFrom, (str(fromYear)).split(',')):
                    print("Exp From extracted as expected")
                    print self.expFrom
                    ws.write(rownum, 42, ','.join(self.expFrom), __style3)
                else:
                    print('Exp From "%s"' % ','.join(self.expFrom) + ' is not extracted as expected "%s"' % str(fromYear))
                    ws.write(rownum, 42, ','.join(self.expFrom), __style2)
            # Getting List Of ExpTo.
            if not self.expTo:
                print("Exp To Is Not Extracted")
                ws.write(rownum, 44, ','.join(self.expTo), __style1)
            else:
                if self.compare_lists(self.expTo, (str(toYear)).split(',')):
                    print("Exp To extracted as expected")
                    print self.expTo
                    ws.write(rownum, 44, ','.join(self.expTo), __style3)
                else:
                    print('Exp To "%s"' % ','.join(self.expTo) + ' is not extracted as expected "%s"' % str(toYear))
                    ws.write(rownum, 44, ','.join(self.expTo), __style2)
            # Getting List Of Salary/CTC.
            if not self.salary:
                print("Salary Is Not Extracted")
                ws.write(rownum, 46, ','.join(self.salary), __style1)
            else:
                if self.compare_lists(self.salary, sal.split(',')):
                    print("Salary extracted as expected")
                    print self.salary
                    ws.write(rownum, 46, ','.join(self.salary), __style3)
                else:
                    print('Salary "%s"' % ','.join(self.salary) + ' is not extracted as expected "%s"' % str(sal))
                    ws.write(rownum, 46, ','.join(self.salary), __style2)
            # Getting List Of IsLatest.
            if not self.isLatest:
                print("IsLatest Is Not Extracted")
                ws.write(rownum, 48, ','.join(self.isLatest), __style1)
            else:
                if self.compare_lists(self.isLatest, islatest.split(',')):
                    print("IsLatest extracted as expected")
                    print self.isLatest
                    ws.write(rownum, 48, ','.join(self.isLatest), __style3)
                else:
                    print('IsLatest "%s"' % ','.join(self.isLatest) + ' is not extracted as expected "%s"' % str(islatest))
                    ws.write(rownum, 48, ','.join(self.isLatest), __style2)
            #Getting Primary Skills.
            primarySkills = self.driver.find_element_by_xpath("//div[3]/div/create-update-candidate/section/div[1]/div/div[3]/div[2]/transcluded-input/div/div/div[1]/div/div/span/span/span[1]").text
            print primarySkills
            total = len(primarySkills)
            print total
            self.__primarySkills = primarySkills
            if not self.__primarySkills:
                print("Primary Skills are Not Extracted")
                ws.write(rownum, 1, self.__primarySkills, __style1)
            else:
                if self.compare_lists(self.__primarySkills, totalSkills.split(',')):
                    print("Primary Skills are extracted as expected")
                    print self.__primarySkills
                    ws.write(rownum, 1, self.__primarySkills, __style3)
                else:
                    print('Primary Skills "%s"' % self.__primarySkills + ' are not extracted as expected "%s"' % str(totalSkills))
                    ws.write(rownum, 1, self.__primarySkills, __style3)
            #Getting Secondary Skills.
            secondarySkills = self.driver.find_element_by_xpath("//div[3]/div/create-update-candidate/section/div[1]/div/div[3]/div[2]/transcluded-input/div/div/div[2]/div/div/span/span/span[1]").text
            print secondarySkills
            total = len(secondarySkills)
            print total
            self.__secondarySkills = secondarySkills
            if not self.__secondarySkills:
                print("Secondary Skills are Not Extracted")
                ws.write(rownum, 2, self.__secondarySkills, __style1)
            else:
                if self.compare_lists(self.__secondarySkills, totalSkills.split(',')):
                    print("Secondary Skills are extracted as expected")
                    print self.__secondarySkills
                    ws.write(rownum, 2, self.__secondarySkills, __style3)
                else:
                    print('Secondary Skills "%s"' % self.__primarySkills + ' are not extracted as expected "%s"' % str(totalSkills))
                    ws.write(rownum, 2, self.__secondarySkills, __style3)
            time.sleep(5)
            wb_result.save('C:\PythonAutomation\AutomationTestResults/ExtractResumeResult(' + __current_DateTime + ').xls')
            #self.driver.quit()
            cnt += 1

    if __name__ == '__main__':
        unittest.main()
