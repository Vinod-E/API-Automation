import unittest
import xlrd
import time
import datetime
import xlwt
from selenium import webdriver
from TestScripts.Config.AllConstants import CONSTANT


class extractResume(unittest.TestCase):
    # @classmethod
    # def setUp(inst):
    # create a new browser session """
    # inst.driver = webdriver.Chrome(CONSTANT.CHROME_DRIVER)
    # inst.driver.implicitly_wait(30)
    # inst.driver.maximize_window()
    # # navigate to the application home page
    # inst.driver.get(CONSTANT.INTERNAL_AMS_URL)
    # inst.driver.title
    # return inst.driver
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
        ws.write(0, 6, 'DOB (MM/DD/YYYY)', __style0)
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

        wb = xlrd.open_workbook("/home/sanjeev/TestScripts/TestScripts/Input_Data/ExtractResume.xls")
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
            self.driver.get(CONSTANT.INTERNAL_AMS_URL)
            # self.driver.title()
            self.driver.find_element_by_name("clientEmail").send_keys(CONSTANT.INTERNAL_AMS_LOGIN_NAME)
            self.driver.find_element_by_name("new-password").send_keys(CONSTANT.INTERNAL_AMS_LOGIN_PASSWORD)
            self.driver.find_element_by_id("input_2").send_keys(CONSTANT.INTERNAL_AMS_TENANT_ALIAS)
            self.driver.find_element_by_xpath("//div[2]/div/div[1]/div/div/div[4]/div[1]/div[1]/div/button").click()
            time.sleep(5)
            __click_Candidate_Grid = self.driver.find_element_by_xpath(
                '/html/body/div[1]/div[2]/div/div[1]/div[1]/div/ul/li[2]/a')
            __click_Candidate_Grid.click()
            time.sleep(5)
            __add_Candidate = self.driver.find_element_by_xpath('//p[2]')
            __add_Candidate.click()
            time.sleep(2)
            __attach_Resume = self.driver.find_element_by_xpath(".//*[@type='file']")
            # print CONSTANT.RESUME_FILE_PATH + rows[0]+'.doc'
            # print __attach_Resume.send_keys(CONSTANT.RESUME_FILE_PATH + rows[0] + '.doc')
            __attach_Resume.send_keys(CONSTANT.RESUME_FILE_PATH + rows[0])
            __resume_File_Name = self.driver.find_element_by_xpath("//div[2]/ng-include/div/div/div[2]/form/div/div/div[2]/div[2]/div[1]/div[2]/div[4]/div/button/span").text
            time.sleep(5)
            __extract_Resume = self.driver.find_element_by_xpath(".//*[@ng-click='vm.extract()']")
            time.sleep(2)
            __extract_Resume.click()
            time.sleep(10)
            name = rows[3]
            email = rows[4]
            altemail = rows[5]
            dob = rows[6]
            dob = dob.strip("'")
            gender = rows[7]
            location = rows[8]
            mobile = rows[9]
            phone = rows[10]
            # sensitivity = rows[11]
            # candidateStatus = rows[12]
            # candidateSourcer = rows[13]
            # expertise = rows[14]
            # experienceYear = rows[15]
            # experienceMonths = rows[16]
            # currentSalaryLPA = rows[17]
            # sourceType = rows[18]
            # source = rows[19]
            # willingToRelocate = rows[20]
            # locationPreference = rows[21]
            # expectedSalaryFromLPA = rows[22]
            # expectedSalaryToLPA = rows[23]
            college = rows[24]
            degree = rows[25]
            branch = rows[26]
            yop = rows[27]
            cgpaorpercentage = rows[28]
            # company = rows[29]
            # designation = rows[30]
            # fromYear = rows[31]
            # toYear = rows[32]
            # salary = rows[33]
            # reasonForLeaving = rows[34]
            # CANDIDATE PERSONAL DETAILS
            ws.write(rownum, 0, __resume_File_Name, __style1)
            __candidate_Name = self.driver.find_element_by_xpath(
                "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[3]/transcluded-input/div/div/form/div[1]/div/input").get_attribute(
                "value")
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
            "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[3]/transcluded-input/div/div/form/div[11]/div/input").get_attribute(
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
            "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[3]/transcluded-input/div/div/form/div[10]/div/input").get_attribute(
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
            "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[3]/transcluded-input/div/div/form/div[12]/div/ta-dropdown/div/div/input").get_attribute(
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
            "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[3]/transcluded-input/div/div/form/div[7]/div/input").get_attribute(
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
            "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[3]/transcluded-input//form/div[8]/div/ta-dropdown//input").get_attribute(
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
            "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[8]/transcluded-input/div/div/div[1]/div[2]/ta-dropdown/div/div/input").get_attribute(
            "Value")
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
            "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[8]/transcluded-input/div/div/div[1]/div[3]/ta-dropdown/div/div/input").get_attribute(
            "Value")
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
            "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[8]/transcluded-input/div/div/div[1]/div[4]/ta-dropdown/div/div/input").get_attribute(
            "Value")
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
            "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[8]/transcluded-input/div/div/div[1]/div[5]/input").text
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
            '/home/sanjeev/TestScripts/TestScripts/Test_Result/ExtractResumeResult(' + __current_DateTime + ').xls')
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
