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
        ws.write(0, 1, 'Expected Name', __style0)
        ws.write(0, 2, 'Actual Name', __style0)
        # ws.write(0, 4, 'Email Id', __style0)
        # ws.write(0, 5, 'Alternate Email Id', __style0)
        wb = xlrd.open_workbook("/home/sanjeev/TestScripts/TestScripts/Input_Data/ExtractResume_Muthu.xls")
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
            # self.driver.find_element_by_name("alias").send_keys(CONSTANT.CRPO_RPO_TENANT_ALIAS).click()

            time.sleep(2)
            self.driver.find_element_by_name("loginName").send_keys(CONSTANT.INTERNAL_AMS_LOGIN_NAME)
            self.driver.find_element_by_xpath("//div[3]/section/div[1]/div[2]/form/div[2]/input").send_keys(
                CONSTANT.INTERNAL_AMS_LOGIN_PASSWORD)
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
            __attach_Resume.send_keys(CONSTANT.RESUME_FILE_PATH_MUTHU + rows[0])
            time.sleep(2)
            __resume_File_Name = self.driver.find_element_by_xpath(
                "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[1]/div[1]/div/div[2]/div[1]").text
            time.sleep(5)
            print __resume_File_Name
            __extract_Resume = self.driver.find_element_by_xpath(
                "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[1]/div[1]/div/div[1]/button")
            __extract_Resume.click()
            __click_Onextract = self.driver.find_element_by_xpath("//div[6]/div/div/div[3]/div/div[1]")
            __click_Onextract.click()
            time.sleep(5)
            name = rows[3]
            # email = rows[4]
            # altemail = rows[5]

            # CANDIDATE PERSONAL DETAILS
            ws.write(rownum, 0, __resume_File_Name, __style1)
            __candidate_Name = self.driver.find_element_by_xpath(
                "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[3]/transcluded-input/div/div/form/div[1]/div/input").get_attribute(
                "value")
            if not __candidate_Name:
                print("Candidate name is not extracted")
                ws.write(rownum, 1, name, __style3)
                ws.write(rownum, 2, __candidate_Name, __style1)
            else:
                if str(__candidate_Name).lower() == str(name).lower():
                    print("Name extracted as expected")
                    ws.write(rownum, 1, name, __style3)
                    ws.write(rownum, 2, __candidate_Name, __style3)
                else:
                    print('Name "%s"' % str(__candidate_Name) + ' is not extracted as expected "%s"' % str(name))
                    ws.write(rownum, 1, name, __style3)
                    ws.write(rownum, 2, __candidate_Name, __style2)
            # __email = self.driver.find_element_by_name('email').get_attribute("value")
            # if not __email:
            #     print("Candidate Email Id is not extracted")
            #     ws.write(rownum, 4, __email, __style1)
            # else:
            #     if str(__email).lower() == str(email).lower():
            #         print("Candidate Email Id extracted as expected")
            #         ws.write(rownum, 4, __email, __style3)
            #     else:
            #         print("Candidate Email Id '%s'" % str(__email) + " is not extracted as expected '%s'" % str(email))
            #         ws.write(rownum, 4, __email, __style2)
            # __alt_Email = self.driver.find_element_by_name('email2').get_attribute("value")
            # if not __alt_Email:
            #     print("Candidate Email Id is not extracted")
            #     ws.write(rownum, 5, __alt_Email, __style1)
            # else:
            #     if str(__alt_Email).lower() == str(altemail).lower():
            #         print("Candidate alternate Email Id extracted as expected")
            #         ws.write(rownum, 5, __alt_Email, __style3)
            #     else:
            #         print(
            #             'Candidate alternate Email Id "%s"' % str(
            #                 __alt_Email) + ' is not extracted as expected "%s"' % str(
            #                 altemail))
            #         ws.write(rownum, 5, __alt_Email, __style2)
            # __mobile = self.driver.find_element_by_xpath(
            #     "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[3]/transcluded-input/div/div/form/div[11]/div/input").get_attribute(
            #     "value")
            # if not __mobile:
            #     print("Mobile number is not extracted")
            #     ws.write(rownum, 9, __mobile, __style1)
            # else:
            #     if int(__mobile) == int(mobile):
            #         print("Mobile number extracted as expected")
            #         ws.write(rownum, 9, __mobile, __style3)
            #     else:
            #         print(
            #             'Mobile number "%s"' % str(__mobile) + ' is not extracted as expected "%s"' % str(int(mobile)))
            #         ws.write(rownum, 9, __mobile, __style2)
            # __phone = self.driver.find_element_by_xpath(
            #     "//div[3]/div/create-update-candidate/section/div[1]/div/div[2]/div[3]/transcluded-input/div/div/form/div[10]/div/input").get_attribute(
            #     "value")
            # if not __phone:
            #     print("Phone number is not extracted")
            #     ws.write(rownum, 10, __phone, __style1)
            # else:
            #     if int(__phone) == int(phone):
            #         print("Phone number extracted as expected")
            #         ws.write(rownum, 10, __phone, __style3)
            #     else:
            #         print(
            #             'Mobile number "%s"' % (str(__phone)) + ' is not extracted as expected "%s"' % (
            #                 str(int(phone))))
            #         ws.write(rownum, 10, __phone, __style2)

            wb_result.save(
                '/home/sanjeev/TestScripts/TestScripts/Test_Result/ExtractResumeResultmuthu1(' + __current_DateTime + ').xls')
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
