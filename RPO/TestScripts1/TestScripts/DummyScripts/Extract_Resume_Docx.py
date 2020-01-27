import unittest
import xlrd
import time
import datetime
import xlwt
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

from TestScripts.Config.AllConstants import CONSTANT


class extractResume(unittest.TestCase):
    def test_extractResume(self):
        now = datetime.datetime.now()
        __current_DateTime = now.strftime("%d-%m-%Y")
        print __current_DateTime
        print type(__current_DateTime)
        __style0 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on')
        __style1 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        __style2 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        __style3 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')
        wb_result = xlwt.Workbook()
        ws = wb_result.add_sheet('Extract Resume Result')
        ws.write(0, 0, 'Resume File Name', __style0)
        ws.write(0, 1, 'Extracted Skills', __style0)
        ws.write(0, 2, 'Not Extracted Skills', __style0)
        ws.write(0, 3, 'Name', __style0)
        ws.write(0, 4, 'Email Id', __style0)
        ws.write(0, 5, 'Alternate Email Id', __style0)
        ws.write(0, 6, 'DOB (DD/MM/YYYY)', __style0)
        ws.write(0, 7, 'Gender', __style0)
        ws.write(0, 8, 'Location', __style0)
        ws.write(0, 9, 'Mobile Number', __style0)
        ws.write(0, 10, 'Phone Number', __style0)
        ws.write(0, 11, 'College', __style0)
        ws.write(0, 12, 'Degree', __style0)
        ws.write(0, 13, 'Branch', __style0)
        ws.write(0, 14, 'YOP', __style0)
        ws.write(0, 15, 'CGPAorPercentage', __style0)
        ws.write(0, 16, 'Matched Skills (Expected vs Extracted)', __style0)
        ws.write(0, 17, 'PanNo', __style0)
        ws.write(0, 18, 'PassportNo', __style0)
        ws.write(0, 19, 'CompanyName', __style0)
        ws.write(0, 20, 'Designation', __style0)
        ws.write(0, 21, 'ExpFrom', __style0)
        ws.write(0, 22, 'ExpTo', __style0)
        ws.write(0, 23, 'Salary', __style0)
        ws.write(0, 24, 'IsLatest', __style0)
        ws.write(0, 25, 'TotalExpInYears', __style0)
        ws.write(0, 26, 'TatalExpInMonths', __style0)

        wb = xlrd.open_workbook("/home/sanjeev/TestScripts/TestScripts/ExtractResumeDocxInput.xls")
        sheetname = wb.sheet_names()  # Read for XLS Sheet names
        print(sheetname)
        sh1 = wb.sheet_by_index(0)  # add login details
        i = 1
        while i < sh1.nrows:
            rownum = i
            rows = sh1.row_values(rownum)
            self.driver = webdriver.Chrome(CONSTANT.CHROME_DRIVER)
            self.driver.implicitly_wait(30)
            self.driver.maximize_window()
            # navigate to the application home page
            self.driver.get(CONSTANT.RPO_CRPO_AMS_URL)
            # self.driver.title()
            # self.driver.find_element_by_name("alias").send_keys(CONSTANT.CRPO_RPO_TENANT_ALIAS).click()

            time.sleep(1)
            self.driver.find_element_by_name("loginName").send_keys(CONSTANT.CRPO_RPO_LOGIN_NAME)
            self.driver.find_element_by_xpath("//div[3]/section/div[1]/div[2]/form/div[2]/input").send_keys(
                CONSTANT.CRPO_RPO_LOGIN_PASSWORD)
            self.driver.find_element_by_xpath("//div[3]/section/div[1]/div[2]/form/div[2]/input").send_keys(Keys.ENTER)
            # time.sleep(1)
            # self.driver.find_element_by_xpath("//div[3]/section/div[1]/div[2]/form/div[3]").click()
            #
            # __attach_Resume = self.driver.find_element_by_xpath("/html/body/div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[1]/div/div/div[1]/upload-file/div/div/input")
            __attach_Resume = self.driver.find_element_by_xpath(
                "/html/body/div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[1]/div[1]/div/div[1]/upload-file/div/div/input")
            # print CONSTANT.RESUME_FILE_PATH + rows[0]+'.doc'
            # print __attach_Resume.send_keys(CONSTANT.RESUME_FILE_PATH + rows[0] + '.doc')
            # time.sleep(1)
            __attach_Resume.send_keys(CONSTANT.RESUME_FILE_PATH + rows[0])
            time.sleep(5)
            __resume_File_Name = self.driver.find_element_by_xpath(
                "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[1]/div[1]/div/div[2]/div[1]").text
            time.sleep(5)
            print __resume_File_Name
            __extract_Resume = self.driver.find_element_by_xpath(
                "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[1]/div[1]/div/div[1]/button")
            __extract_Resume.click()
            __click_Onextract = self.driver.find_element_by_xpath("//div[6]/div/div/div[3]/div/div[1]")
            __click_Onextract.click()
            time.sleep(14)
            name = rows[3]
            email = rows[4]
            altemail = rows[5]
            dob = rows[6]
            dob = dob.strip("'")
            gender = rows[7]
            location = rows[8]
            mobile = rows[9]
            phone = rows[10]
            experienceYear = rows[11]
            experienceMonths = rows[12]
            college = rows[16]
            degree = rows[17]
            branch = rows[18]
            yop = rows[19]
            cgpaorpercentage = rows[20]
            panNumber = rows[28]
            passNo = rows[29]
            compName = rows[21]
            designation = rows[22]
            fromYear = rows[23]
            toYear = rows[24]
            sal = rows[25]
            islatest = rows[26]
            # reasonForLeaving = rows[34]
            # CANDIDATE PERSONAL DETAILS
            ws.write(rownum, 0, __resume_File_Name, __style1)
            __candidate_Name = self.driver.find_element_by_xpath(
                "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[3]/transcluded-input/div/div/form/div[1]/div/input").get_attribute(
                "value")
            time.sleep(2)
            if not __candidate_Name:
                print("Candidate name is not extracted")
                ws.write(rownum, 3, __candidate_Name, __style1)
            else:
                if str(__candidate_Name).lower() == str(name).lower():
                    print("Name extracted as expected")
                    ws.write(rownum, 3, __candidate_Name, __style3)
                else:
                    print('Name "%s"' % str(__candidate_Name) + ' is not extracted as expected "%s"' % str(name))
                    ws.write(rownum, 3, __candidate_Name, __style2)
            __email = self.driver.find_element_by_name('email').get_attribute("value")
            if not __email:
                print("Candidate Email Id is not extracted")
                ws.write(rownum, 4, __email, __style1)
            else:
                if str(__email).lower() == str(email).lower():
                    print("Candidate Email Id extracted as expected")
                    ws.write(rownum, 4, __email, __style3)
                else:
                    print("Candidate Email Id '%s'" % str(__email) + " is not extracted as expected '%s'" % str(email))
                    ws.write(rownum, 4, __email, __style2)
            __alt_Email = self.driver.find_element_by_name('email2').get_attribute("value")
            if not __alt_Email:
                print("Candidate Email Id is not extracted")
                ws.write(rownum, 5, __alt_Email, __style1)
            else:
                if str(__alt_Email).lower() == str(altemail).lower():
                    print("Candidate alternate Email Id extracted as expected")
                    ws.write(rownum, 5, __alt_Email, __style3)
                else:
                    print(
                        'Candidate alternate Email Id "%s"' % str(
                            __alt_Email) + ' is not extracted as expected "%s"' % str(
                            altemail))
                    ws.write(rownum, 5, __alt_Email, __style2)
            __mobile = self.driver.find_element_by_xpath(
                "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[3]/transcluded-input/div/div/form/div[8]/div/input").get_attribute(
                "value")
            if not __mobile:
                print("Mobile number is not extracted")
                ws.write(rownum, 9, __mobile, __style1)
            else:
                if int(__mobile) == int(mobile):
                    print("Mobile number extracted as expected")
                    ws.write(rownum, 9, __mobile, __style3)
                else:
                    print(
                        'Mobile number "%s"' % str(__mobile) + ' is not extracted as expected "%s"' % str(int(mobile)))
                    ws.write(rownum, 9, __mobile, __style2)
            __phone = self.driver.find_element_by_xpath(
                "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[3]/transcluded-input/div/div/form/div[7]/div/input").get_attribute(
                "value")
            if not __phone:
                print("Phone number is not extracted")
                ws.write(rownum, 10, __phone, __style1)
            else:
                if int(__phone) == int(phone):
                    print("Phone number extracted as expected")
                    ws.write(rownum, 10, __phone, __style3)
                else:
                    print(
                        'Mobile number "%s"' % (str(__phone)) + ' is not extracted as expected "%s"' % (
                            str(int(phone))))
                    ws.write(rownum, 10, __phone, __style2)
            __location = self.driver.find_element_by_xpath(
                "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[3]/transcluded-input/div/div/form/div[9]/div/ta-dropdown/div/div/input").get_attribute(
                "value")
            if not __location:
                print("Location is not extracted")
                ws.write(rownum, 8, __location, __style1)
            else:
                if __location.lower() == location.lower():
                    print("Location is extracted as expected")
                    ws.write(rownum, 8, __location, __style3)
                else:
                    print('Location "%s"' % (str(__location)) + ' is not extracted as expected "%s"' % (str(location)))
                    ws.write(rownum, 8, __location, __style2)
            __dob = self.driver.find_element_by_xpath(
                "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[3]/transcluded-input/div/div/form/div[4]/div/input").get_attribute(
                "value")
            if not __dob:
                print("DOB is not extracted")
                ws.write(rownum, 6, __dob, __style1)
            else:
                if __dob == dob:
                    print("DOB is extracted as expected")
                    ws.write(rownum, 6, __dob, __style3)
                else:
                    print('DOB "%s"' % __dob + ' is not extracted as expected "%s"' % dob)
                    ws.write(rownum, 6, __dob, __style2)
            __gender = self.driver.find_element_by_xpath(
                "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[3]/transcluded-input//form/div[5]/div/ta-dropdown//input").get_attribute(
                "value")

            if not __gender:
                print("Gender is not extracted")
                ws.write(rownum, 7, __gender, __style1)
            else:
                if __gender.lower() == gender.lower():
                    print("Gender is extracted as expected")
                    ws.write(rownum, 7, __gender, __style3)
                else:
                    print('Gender "%s"' % __gender + ' is not extracted as expected "%s"' % gender)
                    ws.write(rownum, 7, __gender, __style2)
            __college = self.driver.find_element_by_xpath(
                "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[8]/transcluded-input/div/div/div[2]/table/tbody/tr/td[1]").text
            if not __college:
                print("College name is not extracted")
                ws.write(rownum, 11, __college, __style1)
            else:
                if __college == "Others":
                    __collegeOthers = self.driver.find_element_by_xpath(
                        "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[8]/transcluded-input/div/div/div[2]/table/tbody/tr/td[1]").get_attribute(
                        "value")
                    if __collegeOthers.lower() == college.lower():
                        print("College name is extracted as expected")
                        ws.write(rownum, 11, __collegeOthers, __style3)
                    else:
                        print('College name "%s"' % __collegeOthers + ' is not extracted as expected "%s"' % college)
                        ws.write(rownum, 11, __collegeOthers, __style2)
                else:
                    if __college.lower() == college.lower():
                        print("College name is extracted as expected")
                        ws.write(rownum, 11, __college, __style3)
                    else:
                        print('College name "%s"' % __college + ' is not extracted as expected "%s"' % college)
                        ws.write(rownum, 11, __college, __style2)
            __degree = self.driver.find_element_by_xpath(
                "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[8]/transcluded-input/div/div/div[2]/table/tbody/tr/td[2]").text
            if not __degree:
                print("Degree is not extracted")
                ws.write(rownum, 12, __degree, __style1)
            else:
                if __degree.lower() == degree.lower():
                    print("Degree is extracted as expected")
                    ws.write(rownum, 12, __degree, __style3)
                else:
                    print('Degree "%s"' % __degree + ' is not extracted as expected "%s"' % degree)
                    ws.write(rownum, 12, __degree, __style2)
            __branch = self.driver.find_element_by_xpath(
                "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[8]/transcluded-input/div/div/div[2]/table/tbody/tr/td[3]").text
            if not __branch:
                print("Branch is not extracted")
                ws.write(rownum, 13, __branch, __style1)
            else:
                if str(__branch).lower == str(branch).lower():
                    print("Branch is extracted as expected")
                    ws.write(rownum, 13, __branch, __style3)
                else:
                    print('Branch "%s"' % str(__branch) + ' is not extracted as expected "%s"' % str(branch))
                    ws.write(rownum, 13, __branch, __style2)
            __yop = self.driver.find_element_by_xpath(
                "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[8]/transcluded-input/div/div/div[2]/table/tbody/tr/td[4]").text
            if not __yop:
                print("YOP is not extracted")
                ws.write(rownum, 14, __yop, __style1)
            else:
                if int(__yop) == int(yop):
                    print("YOP is extracted as expected")
                    ws.write(rownum, 14, __yop, __style3)
                else:
                    print('YOP "%s"' % (str(int(__yop))) + ' is not extracted as expected "%s"' % (str(int(yop))))
                    ws.write(rownum, 14, __yop, __style2)
            __percentageorcgpa = self.driver.find_element_by_xpath(
                "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[8]/transcluded-input/div/div/div[2]/table/tbody/tr/td[5]").text
            if not __percentageorcgpa:
                print("Percentage/CGPA is not extracted")
                ws.write(rownum, 15, __percentageorcgpa, __style1)
            else:
                if __percentageorcgpa == cgpaorpercentage:
                    print("Percentage/CGPA is extracted as expected")
                    ws.write(rownum, 15, __percentageorcgpa, __style3)
                else:
                    print('Percentage/CGPA "%s"' % __percentageorcgpa + ' is not extracted as expected "%s"' % str(
                        int(cgpaorpercentage)))
                    ws.write(rownum, 15, __percentageorcgpa, __style2)
            # time.sleep(2)
            __panNo = self.driver.find_element_by_name("pan").get_attribute("value")
            __panNo = str(__panNo).upper()
            if not __panNo:
                print("PanNo is not extracted")
                ws.write(rownum, 17, __panNo, __style1)
            else:
                if __panNo == panNumber:
                    print("PanNo is extracted as expected")
                    ws.write(rownum, 17, __panNo, __style3)
                else:
                    print('PanNo "%s"' % __panNo + ' is not extracted as expected "%s"' % panNumber)
                    ws.write(rownum, 17, __panNo, __style2)


            __passportNo = self.driver.find_element_by_name("passport").get_attribute("value")
            __passportNo = str(__passportNo).upper()
            if not __passportNo:
                print("PassportNo is not extracted")
                ws.write(rownum, 18, __passportNo, __style1)
            else:
                if __passportNo == passNo:
                    print("PassportNo is extracted as expected")
                    ws.write(rownum, 18, __passportNo, __style3)
                else:
                    print('PassportNo "%s"' % __passportNo + ' is not extracted as expected "%s"' % str(passNo))
                    ws.write(rownum, 18, __passportNo, __style2)

            __totalExpInYears = self.driver.find_element_by_xpath("//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[7]/transcluded-input/div/div/form/div[3]/div/input").get_attribute("value")
            if not __totalExpInYears:
                print("Total Exp In Years is not extracted")
                ws.write(rownum, 25, __totalExpInYears, __style1)
            else:
                if int(__totalExpInYears) == int(experienceYear):
                    print("Total Exp In Years extracted as expected")
                    ws.write(rownum, 25, __totalExpInYears, __style3)
                else:
                    print(
                        'Total Exp In Years "%s"' % str(__totalExpInYears) + ' is not extracted as expected "%s"' % str(int(experienceYear)))
                    ws.write(rownum, 25, __totalExpInYears, __style2)

            __totalExpInMonths = self.driver.find_element_by_xpath(
                "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[7]/transcluded-input/div/div/form/div[4]/div/input").get_attribute(
                "value")
            if not __totalExpInMonths:
                print("Total Exp In Months is not extracted")
                ws.write(rownum, 26, __totalExpInMonths, __style1)
            else:
                if int(__totalExpInMonths) == int(experienceMonths):
                    print("Total Exp In Months extracted as expected")
                    ws.write(rownum, 26, __totalExpInMonths, __style3)
                else:
                    print(
                        'Total Exp In Months "%s"' % str(__totalExpInMonths) + ' is not extracted as expected "%s"' % str(int(experienceMonths)))
                    ws.write(rownum, 26, __totalExpInMonths, __style2)

            __companyName = self.driver.find_element_by_xpath("//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[9]/transcluded-input/div/div/div[2]/table/tbody/tr/td[1]").text
            if not __companyName:
                print("CompanyName is not extracted")
                ws.write(rownum, 19, __companyName, __style1)
            else:
                if __companyName == compName:
                    print("CompanyName is extracted as expected")
                    ws.write(rownum, 19, __companyName, __style3)
                else:
                    print('CompanyName "%s"' % __companyName + ' is not extracted as expected "%s"' % str(compName))
                    ws.write(rownum, 19, __companyName, __style2)

            __desigNation = self.driver.find_element_by_xpath("//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[9]/transcluded-input/div/div/div[2]/table/tbody/tr/td[2]").text
            if not __desigNation:
                print("Designation is not extracted")
                ws.write(rownum, 20, __desigNation, __style1)
            else:
                if __desigNation == designation:
                    print("Designation is extracted as expected")
                    ws.write(rownum, 20, __desigNation, __style3)
                else:
                    print('Designation "%s"' % __desigNation + ' is not extracted as expected "%s"' % str(designation))
                    ws.write(rownum, 20, __desigNation, __style2)

            __expFrom = self.driver.find_element_by_xpath("//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[9]/transcluded-input/div/div/div[2]/table/tbody/tr/td[3]").text
            if not __expFrom:
                print("ExpFrom is not extracted")
                ws.write(rownum, 21, __expFrom, __style1)
            else:
                if __expFrom == fromYear:
                    print("ExpFrom is extracted as expected")
                    ws.write(rownum, 21, __expFrom, __style3)
                else:
                    print('ExpFrom "%s"' % __expFrom + ' is not extracted as expected "%s"' % str(fromYear))
                    ws.write(rownum, 21, __expFrom, __style2)



            __expTo = self.driver.find_element_by_xpath("//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[9]/transcluded-input/div/div/div[2]/table/tbody/tr/td[4]").text
            if not __expTo:
                print("ExpTo is not extracted")
                ws.write(rownum, 22, __expTo, __style1)
            else:
                if __expTo == toYear:
                    print("ExpTo is extracted as expected")
                    ws.write(rownum, 22, __expTo, __style3)
                else:
                    print('ExpTo "%s"' % __expTo + ' is not extracted as expected "%s"' % str(toYear))
                    ws.write(rownum, 22, __expTo, __style2)


            __salary = self.driver.find_element_by_xpath("//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[9]/transcluded-input/div/div/div[2]/table/tbody/tr/td[5]").text
            if not __salary:
                print("Salary is not extracted")
                ws.write(rownum, 23, __salary, __style1)
            else:
                if __salary == sal:
                    print("Salary is extracted as expected")
                    ws.write(rownum, 23, __salary, __style3)
                else:
                    print('Salary "%s"' % __salary + ' is not extracted as expected "%s"' % str(sal))
                    ws.write(rownum, 23, __salary, __style2)


            __isLatest = self.driver.find_element_by_xpath("//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[9]/transcluded-input/div/div/div[2]/table/tbody/tr/td[7]").text
            if not __isLatest:
                print("IsLatest is not extracted")
                ws.write(rownum, 24, __isLatest, __style1)
            else:
                if __isLatest == islatest:
                    print("IsLatest is extracted as expected")
                    ws.write(rownum, 24, __isLatest, __style3)
                else:
                    print('IsLatest "%s"' % __isLatest + ' is not extracted as expected "%s"' % str(islatest))
                    ws.write(rownum, 24, __isLatest, __style2)


            __skillsOfCandidate = self.driver.find_element_by_xpath(
                "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[2]/transcluded-input/div/div/div/div/div/span/span[1]").text
            __extracted_Skills = __skillsOfCandidate.replace("\nRemove\nPress delete to remove this chip.", ",")
            __extracted_Skills1 = [s.strip() for s in __extracted_Skills]
            __extracted_Skills = [s.strip().lower() for s in __extracted_Skills.split(',')]
            # print __extracted_Skills
            __excel_Skills = rows[2]
            # __excel_Skills = [x.strip().lower() for x in str(__excel_Skills).split(',')]
            __excel_Skills = [x.strip().lower() for x in __excel_Skills.split(',')]
            __not_matched = []
            __matched = []
            for skill in __excel_Skills:
                if len(skill) > 0:
                    if skill not in __extracted_Skills:
                        __not_matched.append(skill + ", ")
                    else:
                        __matched.append(skill + ", ")
                else:
                    print("All skill matched")
                    # print len(__excel_Skills), len(__extracted_Skills), len(not_matched)
            ws.write(rownum, 1, __extracted_Skills1, __style1)
            ws.write(rownum, 2, __not_matched, __style1)
            ws.write(rownum, 16, __matched, __style1)
            # sourcetype = self.driver.find_element_by_name('sourcetype')
            # sourcetype.send_keys("Direct")
            # time.sleep(1)
            # SOURCE
            sourcetype = self.driver.find_element_by_xpath(
                "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[5]/transcluded-input/div/div/div/div/ta-dropdown/div/div/input")
            sourcetype.send_keys("ABC Consultant")
            time.sleep(1)
            # __reason_For_Leaving = self.driver.find_element_by_name("addedreasonOfLeaving")
            # __reason_For_Leaving.click()
            # __reason_For_Leaving.send_keys(Keys.ENTER)
            # __msg_Content = self.driver.find_element_by_css_selector("div.growl-message.ng-binding").text
            # time.sleep(5)
            wb_result.save(
                '/home/sanjeev/TestScripts/TestScripts/ResultsOfResumeParsingDocx/ExtractResumeResult(' + __current_DateTime + ').xls')
            # print __msg_Content
            # if __msg_Content == __msg_Content:
            #     try:
            #         conn = mysql.connector.connect(host='10.0.3.35',
            #                                        database='coredbtest',
            #                                        user='root',
            #                                        password='root')
            #         cursor = conn.cursor()
            #         cursor.execute("SELECT email1 from candidates where email1 like '" + __email + "'")
            #         data = cursor.fetchone()
            #         dataResult = [str(s).strip() for s in data[0].split(',')]
            #         print dataResult
            #
            #     except Error as e:
            #         print(e)
            #
            #     finally:
            #         conn.close()
            # else:
            #     print (__msg_Content)
            self.driver.quit()
            i += 1
            # @classmethodR
            # def tearDown(inst):
            #     # close the browser window
            #     inst.driver.quit()

    if __name__ == '__main__':
        unittest.main()
