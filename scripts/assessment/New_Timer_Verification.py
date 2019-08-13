import json
import requests
import time
import unittest
import xlrd
import mysql
from mysql import connector
import xlwt
from Config import Api
from scripts.assessment import WebConfig
from selenium import webdriver


class TestTimerCheck(unittest.TestCase):
    @classmethod
    def setUp(cls):
        cls.driver = webdriver.Chrome(WebConfig.CHROME_DRIVER)
        cls.driver.implicitly_wait(30)
        cls.driver.maximize_window()
        cls.driver.get(WebConfig.ONLINE_ASSESSMENT_LOGIN_URL)
        return cls.driver

    def test_TestTimerCheck(self):
        # --------------------------------------------------------------------------------------------------------------
        # CSS to differentiate Correct and Wrong data in Excel
        # --------------------------------------------------------------------------------------------------------------
        self.__style0 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on')
        self.__style1 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        self.__style2 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        self.__style3 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')

        # --------------------------------------------------------------------------------------------------------------
        # Read from Excel
        # --------------------------------------------------------------------------------------------------------------
        wb = xlrd.open_workbook("/home/testingteam/hirepro_automation/API-Automation/Input Data/Assessment/Timer.xls")
        wb_result = xlwt.Workbook()
        ws = wb_result.add_sheet('Timer_Check')
        sh1 = wb.sheet_by_index(0)

        # --------------------------------------------------------------------------------------------------------------
        # Header printing in Output Excel
        # --------------------------------------------------------------------------------------------------------------
        ws.write(0, 0, "Overall Status", self.__style0)
        ws.write(0, 1, "Status", self.__style0)
        ws.write(0, 2, "Test Id", self.__style0)
        ws.write(0, 3, "Candidate Ids", self.__style0)
        ws.write(0, 4, "Login Ids", self.__style0)
        ws.write(0, 5, "Passwords", self.__style0)

        ws.write(0, 6, "Default Time Interval (G1)", self.__style0)
        ws.write(0, 7, "Q1 Start Time Iteration 1", self.__style0)
        ws.write(0, 8, "Q1 End Time Iteration 1", self.__style0)
        ws.write(0, 9, "Q1 Time Spent Iteration 1 (Sec)", self.__style0)
        ws.write(0, 10, "Q1 Start Time Iteration 2", self.__style0)
        ws.write(0, 11, "Q1 End Time Iteration 2", self.__style0)
        ws.write(0, 12, "Q1 Time Spent Iteration 2 (Sec)", self.__style0)
        ws.write(0, 13, "Q1 Expected Time Spent", self.__style0)
        ws.write(0, 14, "Q1 Actual Time Spent UI", self.__style0)
        ws.write(0, 15, "Q1 Actual Time Spent DB", self.__style0)

        ws.write(0, 16, "Q2 Start Time Iteration 1", self.__style0)
        ws.write(0, 17, "Q2 End Time Iteration 1", self.__style0)
        ws.write(0, 18, "Q2 Time Spent Iteration 1 (Sec)", self.__style0)
        ws.write(0, 19, "Q2 Start Time Iteration 2", self.__style0)
        ws.write(0, 20, "Q2 End Time Iteration 2", self.__style0)
        ws.write(0, 21, "Q2 Time Spent Iteration 2 (Sec)", self.__style0)
        ws.write(0, 22, "Q2 Expected Time Spent", self.__style0)
        ws.write(0, 23, "Q2 Actual Time Spent UI", self.__style0)
        ws.write(0, 24, "Q2 Actual Time Spent DB", self.__style0)

        ws.write(0, 25, "Q3 Start Time Iteration 1", self.__style0)
        ws.write(0, 26, "Q3 End Time Iteration 1", self.__style0)
        ws.write(0, 27, "Q3 Time Spent Iteration 1 (Sec)", self.__style0)
        ws.write(0, 28, "Q3 Start Time Iteration 2", self.__style0)
        ws.write(0, 29, "Q3 End Time Iteration 2", self.__style0)
        ws.write(0, 30, "Q3 Time Spent Iteration 2 (Sec)", self.__style0)
        ws.write(0, 31, "Q3 Expected Time Spent", self.__style0)
        ws.write(0, 32, "Q3 Actual Time Spent UI", self.__style0)
        ws.write(0, 33, "Q3 Actual Time Spent DB", self.__style0)

        ws.write(0, 34, "Default Time Interval (G2)", self.__style0)
        ws.write(0, 35, "Q4 Start Time Iteration 1", self.__style0)
        ws.write(0, 36, "Q4 End Time Iteration 1", self.__style0)
        ws.write(0, 37, "Q4 Time Spent Iteration 1 (Sec)", self.__style0)
        ws.write(0, 38, "Q4 Start Time Iteration 2", self.__style0)
        ws.write(0, 39, "Q4 End Time Iteration 2", self.__style0)
        ws.write(0, 40, "Q4 Time Spent Iteration 2 (Sec)", self.__style0)
        ws.write(0, 41, "Q4 Expected Time Spent", self.__style0)
        ws.write(0, 42, "Q4 Actual Time Spent UI", self.__style0)
        ws.write(0, 43, "Q4 Actual Time Spent DB", self.__style0)

        ws.write(0, 44, "Q5 Start Time Iteration 1", self.__style0)
        ws.write(0, 45, "Q5 End Time Iteration 1", self.__style0)
        ws.write(0, 46, "Q5 Time Spent Iteration 1 (Sec)", self.__style0)
        ws.write(0, 47, "Q5 Start Time Iteration 2", self.__style0)
        ws.write(0, 48, "Q5 End Time Iteration 2", self.__style0)
        ws.write(0, 49, "Q5 Time Spent Iteration 2 (Sec)", self.__style0)
        ws.write(0, 50, "Q5 Expected Time Spent", self.__style0)
        ws.write(0, 51, "Q5 Actual Time Spent UI", self.__style0)
        ws.write(0, 52, "Q5 Actual Time Spent DB", self.__style0)

        ws.write(0, 53, "Expected Test Time", self.__style0)
        ws.write(0, 54, "Actual Test Time UI", self.__style0)
        ws.write(0, 55, "Actual Test Time DB", self.__style0)

        n = 1
        row_num = n
        status = []
        while n < sh1.nrows:
            rows = sh1.row_values(row_num)
            test_id = rows[2]
            candidate_id = rows[3]
            login_id = rows[4]
            password = rows[5]
            g1_default_time_spent = rows[6]
            q1_expected_time_spent = rows[13]
            q2_expected_time_spent = rows[22]
            q3_expected_time_spent = rows[31]
            g2_default_time_spent = rows[34]
            q4_expected_time_spent = rows[41]
            q5_expected_time_spent = rows[50]
            total_expected_test_time_spent = rows[53]
            q1_id = rows[56]
            q2_id = rows[57]
            q3_id = rows[58]
            q4_id = rows[59]
            q5_id = rows[60]
            time.sleep(2)
            ws.write(row_num, 2, test_id, self.__style1)
            ws.write(row_num, 3, candidate_id, self.__style1)
            ws.write(row_num, 4, login_id, self.__style1)
            ws.write(row_num, 5, password, self.__style1)
            ws.write(row_num, 6, g1_default_time_spent, self.__style1)
            ws.write(row_num, 13, q1_expected_time_spent, self.__style1)
            ws.write(row_num, 22, q2_expected_time_spent, self.__style1)
            ws.write(row_num, 31, q3_expected_time_spent, self.__style1)
            ws.write(row_num, 34, g2_default_time_spent, self.__style1)
            ws.write(row_num, 41, q4_expected_time_spent, self.__style1)
            ws.write(row_num, 50, q5_expected_time_spent, self.__style1)
            ws.write(row_num, 53, total_expected_test_time_spent, self.__style1)
            self.driver.get(WebConfig.ONLINE_ASSESSMENT_LOGIN_URL)
            time.sleep(2)
            self.driver.find_element_by_xpath("//div[8]/div/div/div[2]/input").send_keys(WebConfig.ALIAS1)
            time.sleep(3)
            self.driver.find_element_by_xpath("//div[8]/div/div/div[3]/div[2]/button").click()
            time.sleep(3)
            self.driver.find_element_by_name("btnLogin").click()
            time.sleep(5)
            self.driver.switch_to.window(self.driver.window_handles[0])
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])
            time.sleep(2)
            self.driver.find_element_by_name("loginUsername").send_keys(login_id)
            self.driver.find_element_by_name("loginPassword").send_keys(password)
            time.sleep(3)
            self.driver.find_element_by_name("btnLogin").click()
            time.sleep(7)
            self.driver.find_element_by_xpath("//*[@id='start-test-col']/div//label").click()
            time.sleep(1)
            self.driver.find_element_by_name("btnStartTest").click()
            print("Start Test")

            # ----------------------------------------------------------------------------------------------------------
            #  Q1 Time calculation - Iteration 1
            # ----------------------------------------------------------------------------------------------------------
            q1_start_time_iteration_1 = ""
            msec = 0.2
            for time_msec in range(0, 50):
                time.sleep(msec)
                if self.driver.find_element_by_name("testTimeCounter").is_displayed():
                    q1_start_time_iteration_1 = self.driver.find_element_by_name("testTimeCounter").text
                    break
            ws.write(row_num, 7, q1_start_time_iteration_1)
            q1_start_iteration_1 = time.strptime(q1_start_time_iteration_1, '%H:%M:%S')
            q1_start_time_in_seconds_iteration_1 = (q1_start_iteration_1.tm_hour * 3600) + (q1_start_iteration_1.tm_min * 60) + q1_start_iteration_1.tm_sec
            print("q1_start_time_iteration_1", q1_start_time_iteration_1, "q1_start_time_in_seconds_iteration_1", q1_start_time_in_seconds_iteration_1)
            print("time.sleep(q1)", g1_default_time_spent)
            time.sleep(g1_default_time_spent)

            q1_end_time_iteration_1 = self.driver.find_element_by_name("testTimeCounter").text
            ws.write(row_num, 8, q1_end_time_iteration_1)
            q1_end_iteration_1 = time.strptime(q1_end_time_iteration_1, '%H:%M:%S')
            q1_end_time_in_seconds_iteration_1 = (q1_end_iteration_1.tm_hour * 3600) + (q1_end_iteration_1.tm_min * 60) + q1_end_iteration_1.tm_sec
            print("q1_end_time_iteration_1", q1_end_time_iteration_1, "q1_end_time_in_seconds_iteration_1", q1_end_time_in_seconds_iteration_1)
            q1_time_spent_iteration_1 = q1_start_time_in_seconds_iteration_1 - q1_end_time_in_seconds_iteration_1
            ws.write(row_num, 9, q1_time_spent_iteration_1)
            print("First iteration q1 time spent", q1_time_spent_iteration_1)

            # ----------------------------------------------------------------------------------------------------------
            #  Q2 Time calculation - Iteration 1
            # ----------------------------------------------------------------------------------------------------------
            self.driver.find_element_by_name("btnNext").click()
            q2_start_time_iteration_1 = self.driver.find_element_by_name("testTimeCounter").text
            ws.write(row_num, 16, q2_start_time_iteration_1)
            q2_start_iteration_1 = time.strptime(q2_start_time_iteration_1, '%H:%M:%S')
            q2_start_time_in_seconds_iteration_1 = (q2_start_iteration_1.tm_hour * 3600) + (
                    q2_start_iteration_1.tm_min * 60) + q2_start_iteration_1.tm_sec
            print("q2_start_time_iteration_1", q2_start_time_iteration_1, "q2_start_time_in_seconds_iteration_1",
                  q2_start_time_in_seconds_iteration_1)
            print("time.sleep(g1)", g1_default_time_spent)
            time.sleep(g1_default_time_spent)
            # self.driver.execute_script('javascript:localStorage.clear();')
            # self.driver.execute_script('javascript:alert(localStorage.length);')
            q2_end_time_iteration_1 = self.driver.find_element_by_name("testTimeCounter").text
            ws.write(row_num, 17, q2_end_time_iteration_1)
            # self.driver.execute_script('javascript:alert(localStorage.length);')
            q2_end_iteration_1 = time.strptime(q2_end_time_iteration_1, '%H:%M:%S')
            # self.driver.execute_script('javascript:localStorage.clear();')
            # self.driver.execute_script('javascript:alert(localStorage.length);')
            q2_end_time_in_seconds_iteration_1 = (q2_end_iteration_1.tm_hour * 3600) + (
                    q2_end_iteration_1.tm_min * 60) + q2_end_iteration_1.tm_sec
            print("q2_end_time_iteration_1", q2_end_time_iteration_1, "q2_end_time_in_seconds_iteration_1",
                  q2_end_time_in_seconds_iteration_1)
            q2_time_spent_iteration_1 = q2_start_time_in_seconds_iteration_1 - q2_end_time_in_seconds_iteration_1
            ws.write(row_num, 18, q2_time_spent_iteration_1)
            print("First iteration q2 time spent", q2_time_spent_iteration_1)

            # ----------------------------------------------------------------------------------------------------------
            #  Q3 Time calculation - Iteration 1
            # ----------------------------------------------------------------------------------------------------------
            self.driver.find_element_by_name("btnNext").click()
            q3_start_time_iteration_1 = self.driver.find_element_by_name("testTimeCounter").text
            ws.write(row_num, 25, q3_start_time_iteration_1)
            q3_start_iteration_1 = time.strptime(q3_start_time_iteration_1, '%H:%M:%S')
            q3_start_time_in_seconds_iteration_1 = (q3_start_iteration_1.tm_hour * 3600) + (
                    q3_start_iteration_1.tm_min * 60) + q3_start_iteration_1.tm_sec
            print("q3_start_time_iteration_1", q3_start_time_iteration_1, "q3_start_time_in_seconds_iteration_1",
                  q3_start_time_in_seconds_iteration_1)
            print("time.sleep(g1)", g1_default_time_spent)
            time.sleep(g1_default_time_spent)

            q3_end_time_iteration_1 = self.driver.find_element_by_name("testTimeCounter").text
            ws.write(row_num, 26, q3_end_time_iteration_1)
            q3_end_iteration_1 = time.strptime(q3_end_time_iteration_1, '%H:%M:%S')
            q3_end_time_in_seconds_iteration_1 = (q3_end_iteration_1.tm_hour * 3600) + (
                    q3_end_iteration_1.tm_min * 60) + q3_end_iteration_1.tm_sec
            print("q3_end_time_iteration_1", q3_end_time_iteration_1, "q3_end_time_in_seconds_iteration_1",
                  q3_end_time_in_seconds_iteration_1)
            q3_time_spent_iteration_1 = q3_start_time_in_seconds_iteration_1 - q3_end_time_in_seconds_iteration_1
            ws.write(row_num, 27, q3_time_spent_iteration_1)
            print("First iteration q3 time spent", q3_time_spent_iteration_1)

            # ----------------------------------------------------------------------------------------------------------
            #  Q4 Time calculation - Iteration 1
            # ----------------------------------------------------------------------------------------------------------
            self.driver.find_element_by_name("btnNext").click()
            q4_start_time_iteration_1 = self.driver.find_element_by_name("testTimeCounter").text
            ws.write(row_num, 35, q4_start_time_iteration_1)
            q4_start_iteration_1 = time.strptime(q4_start_time_iteration_1, '%H:%M:%S')
            q4_start_time_in_seconds_iteration_1 = (q4_start_iteration_1.tm_hour * 3600) + (
                    q4_start_iteration_1.tm_min * 60) + q4_start_iteration_1.tm_sec
            print("q4_start_time_iteration_1", q4_start_time_iteration_1, "q4_start_time_in_seconds_iteration_1",
                  q4_start_time_in_seconds_iteration_1)
            print("time.sleep(g2)", g2_default_time_spent)
            time.sleep(g2_default_time_spent)

            q4_end_time_iteration_1 = self.driver.find_element_by_name("testTimeCounter").text
            ws.write(row_num, 36, q4_end_time_iteration_1)
            q4_end_iteration_1 = time.strptime(q4_end_time_iteration_1, '%H:%M:%S')
            q4_end_time_in_seconds_iteration_1 = (q4_end_iteration_1.tm_hour * 3600) + (
                    q4_end_iteration_1.tm_min * 60) + q4_end_iteration_1.tm_sec
            print("q4_end_time_iteration_1", q4_end_time_iteration_1, "q4_end_time_in_seconds_iteration_1",
                  q4_end_time_in_seconds_iteration_1)
            q4_time_spent_iteration_1 = q4_start_time_in_seconds_iteration_1 - q4_end_time_in_seconds_iteration_1
            ws.write(row_num, 37, q4_time_spent_iteration_1)
            print("First iteration q4 time spent", q4_time_spent_iteration_1)

            # ----------------------------------------------------------------------------------------------------------
            #  Q5 Time calculation - Iteration 1
            # ----------------------------------------------------------------------------------------------------------
            self.driver.find_element_by_name("btnNext").click()
            q5_start_time_iteration_1 = self.driver.find_element_by_name("testTimeCounter").text
            ws.write(row_num, 44, q5_start_time_iteration_1)
            q5_start_iteration_1 = time.strptime(q5_start_time_iteration_1, '%H:%M:%S')
            q5_start_time_in_seconds_iteration_1 = (q5_start_iteration_1.tm_hour * 3600) + (
                    q5_start_iteration_1.tm_min * 60) + q5_start_iteration_1.tm_sec
            print("q5_start_time_iteration_1", q5_start_time_iteration_1, "q5_start_time_in_seconds_iteration_1",
                  q5_start_time_in_seconds_iteration_1)
            print("time.sleep(g2)", g2_default_time_spent)
            time.sleep(g2_default_time_spent)

            q5_end_time_iteration_1 = self.driver.find_element_by_name("testTimeCounter").text
            ws.write(row_num, 45, q5_end_time_iteration_1)
            q5_end_iteration_1 = time.strptime(q5_end_time_iteration_1, '%H:%M:%S')
            q5_end_time_in_seconds_iteration_1 = (q5_end_iteration_1.tm_hour * 3600) + (
                    q5_end_iteration_1.tm_min * 60) + q5_end_iteration_1.tm_sec
            print("q5_end_time_iteration_1", q5_end_time_iteration_1, "q5_end_time_in_seconds_iteration_1",
                  q5_end_time_in_seconds_iteration_1)
            q5_time_spent_iteration_1 = q5_start_time_in_seconds_iteration_1 - q5_end_time_in_seconds_iteration_1
            ws.write(row_num, 46, q5_time_spent_iteration_1)
            print("First iteration q5 time spent", q5_time_spent_iteration_1)

            # ----------------------------------------------------------------------------------------------------------
            #  Refresh Page
            # ----------------------------------------------------------------------------------------------------------
            self.driver.refresh()
            time.sleep(2)
            self.driver.get(WebConfig.ONLINE_ASSESSMENT_LOGIN_URL)
            time.sleep(2)
            self.driver.find_element_by_xpath("//div[8]/div/div/div[2]/input").send_keys(WebConfig.ALIAS1)
            time.sleep(1)
            self.driver.find_element_by_xpath("//div[8]/div/div/div[3]/div[2]/button").click()
            time.sleep(2)
            self.driver.find_element_by_name("btnLogin").click()
            time.sleep(5)
            self.driver.switch_to.window(self.driver.window_handles[0])
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])
            time.sleep(2)
            self.driver.find_element_by_name("loginUsername").send_keys(login_id)
            time.sleep(1)
            self.driver.find_element_by_name("loginPassword").send_keys(password)
            time.sleep(3)
            self.driver.find_element_by_name("btnLogin").click()
            time.sleep(5)

            # ----------------------------------------------------------------------------------------------------------
            #  Login to AMS
            # ----------------------------------------------------------------------------------------------------------
            crpo_login_header = {"content-type": "application/json"}
            data0 = {"LoginName": "admin", "Password": "4LWS-067", "TenantAlias": "automation",
                     "UserName": "admin"}
            response = requests.post(Api.login_user, headers=crpo_login_header, data=json.dumps(data0), verify=True)
            self.TokenVal = response.json()
            self.NTokenVal = self.TokenVal.get("Token")

            # ----------------------------------------------------------------------------------------------------------
            #  Reactivate Test user login and login to test again
            # ----------------------------------------------------------------------------------------------------------
            eval_online_assessment_header = {"content-type": "application/json", "X-AUTH-TOKEN": self.NTokenVal}
            data5 = {"candidateIds": [candidate_id], "testId": test_id}
            request5 = requests.post("https://amsin.hirepro.in/py/assessment/testuser/api/v1/reActivateLogin/",
                                     headers=eval_online_assessment_header,
                                     data=json.dumps(data5, default=str), verify=True)
            time.sleep(4)
            print("Password Reactivate", request5)
            self.driver.find_element_by_name("btnLogin").click()
            time.sleep(7)
            self.driver.find_element_by_xpath("//*[@id='start-test-col']/div//label").click()
            time.sleep(3)
            self.driver.find_element_by_name("btnStartTest").click()
            print("Start Test")

            # ----------------------------------------------------------------------------------------------------------
            #  Q1 Time calculation - Iteration 2
            # ----------------------------------------------------------------------------------------------------------
            q1_start_time_iteration_2 = ""
            msec = 0.2
            for time_msec in range(0, 50):
                time.sleep(msec)
                print("Log", time_msec)
                if self.driver.find_element_by_name("testTimeCounter").is_displayed():
                    q1_start_time_iteration_2 = self.driver.find_element_by_name("testTimeCounter").text
                    break
            ws.write(row_num, 10, q1_start_time_iteration_2)
            q1_start_iteration_2 = time.strptime(q1_start_time_iteration_2, '%H:%M:%S')
            q1_start_time_in_seconds_iteration_2 = (q1_start_iteration_2.tm_hour * 3600) + (
                    q1_start_iteration_2.tm_min * 60) + q1_start_iteration_2.tm_sec
            print("q1_start_time_iteration_2", q1_start_time_iteration_2, "q1_start_time_in_seconds_iteration_2",
                  q1_start_time_in_seconds_iteration_2)
            print("time.sleep(q1)", g1_default_time_spent)
            time.sleep(g1_default_time_spent)

            # self.driver.execute_script('javascript:alert(localStorage.length);')

            q1_end_time_iteration_2 = self.driver.find_element_by_name("testTimeCounter").text
            ws.write(row_num, 11, q1_end_time_iteration_2)
            q1_end_iteration_2 = time.strptime(q1_end_time_iteration_2, '%H:%M:%S')
            q1_end_time_in_seconds_iteration_2 = (q1_end_iteration_2.tm_hour * 3600) + (
                    q1_end_iteration_2.tm_min * 60) + q1_end_iteration_2.tm_sec
            print("q1_end_time_iteration_2", q1_end_time_iteration_2, "q1_end_time_in_seconds_iteration_2",
                  q1_end_time_in_seconds_iteration_2)
            q1_time_spent_iteration_2 = q1_start_time_in_seconds_iteration_2 - q1_end_time_in_seconds_iteration_2
            ws.write(row_num, 12, q1_time_spent_iteration_2)
            print("Second iteration q1 time spent", q1_time_spent_iteration_2)

            total_time_spent_in_q1 = q1_time_spent_iteration_1 + q1_time_spent_iteration_2
            row_status = []
            if q1_expected_time_spent == total_time_spent_in_q1:
                ws.write(row_num, 14, total_time_spent_in_q1, self.__style3)
                row_status.append("Pass")
            else:
                ws.write(row_num, 14, total_time_spent_in_q1, self.__style2)
                row_status.append("Fail")

            # ----------------------------------------------------------------------------------------------------------
            #  Q2 Time calculation - Iteration 2
            # ----------------------------------------------------------------------------------------------------------
            self.driver.find_element_by_name("btnNext").click()
            q2_start_time_iteration_2 = self.driver.find_element_by_name("testTimeCounter").text
            ws.write(row_num, 19, q2_start_time_iteration_2)
            q2_start_iteration_2 = time.strptime(q2_start_time_iteration_2, '%H:%M:%S')
            q2_start_time_in_seconds_iteration_2 = (q2_start_iteration_2.tm_hour * 3600) + (
                    q2_start_iteration_2.tm_min * 60) + q2_start_iteration_2.tm_sec
            print("q2_start_time_iteration_2", q2_start_time_iteration_2, "q2_start_time_in_seconds_iteration_2",
                  q2_start_time_in_seconds_iteration_2)
            print("time.sleep(g1)", g1_default_time_spent)
            time.sleep(g1_default_time_spent)

            q2_end_time_iteration_2 = self.driver.find_element_by_name("testTimeCounter").text
            ws.write(row_num, 20, q2_end_time_iteration_2)
            q2_end_iteration_2 = time.strptime(q2_end_time_iteration_2, '%H:%M:%S')
            q2_end_time_in_seconds_iteration_2 = (q2_end_iteration_2.tm_hour * 3600) + (
                    q2_end_iteration_2.tm_min * 60) + q2_end_iteration_2.tm_sec
            print("q2_end_time_iteration_2", q2_end_time_iteration_2, "q2_end_time_in_seconds_iteration_2",
                  q2_end_time_in_seconds_iteration_2)

            q2_time_spent_iteration_2 = q2_start_time_in_seconds_iteration_2 - q2_end_time_in_seconds_iteration_2
            ws.write(row_num, 21, q2_time_spent_iteration_2)
            print("Second iteration q2 time spent", q2_time_spent_iteration_2)
            total_time_spent_in_q2 = q2_time_spent_iteration_1 + q2_time_spent_iteration_2
            if q2_expected_time_spent == total_time_spent_in_q2 or (
                    q2_expected_time_spent + 1) == total_time_spent_in_q2 or (
                    q2_expected_time_spent + 2) == total_time_spent_in_q2:
                ws.write(row_num, 23, total_time_spent_in_q2, self.__style3)
                row_status.append("Pass")
            else:
                ws.write(row_num, 23, total_time_spent_in_q2, self.__style2)
                row_status.append("Fail")

            # ----------------------------------------------------------------------------------------------------------
            #  Q3 Time calculation - Iteration 2
            # ----------------------------------------------------------------------------------------------------------
            self.driver.find_element_by_name("btnNext").click()
            q3_start_time_iteration_2 = self.driver.find_element_by_name("testTimeCounter").text
            ws.write(row_num, 28, q3_start_time_iteration_2)
            q3_start_iteration_2 = time.strptime(q3_start_time_iteration_2, '%H:%M:%S')
            q3_start_time_in_seconds_iteration_2 = (q3_start_iteration_2.tm_hour * 3600) + (
                    q3_start_iteration_2.tm_min * 60) + q3_start_iteration_2.tm_sec
            print("q3_start_time_iteration_2", q3_start_time_iteration_2, "q3_start_time_in_seconds_iteration_2",
                  q3_start_time_in_seconds_iteration_2)
            print("time.sleep(g1)", g1_default_time_spent)
            time.sleep(g1_default_time_spent)

            q3_end_time_iteration_2 = self.driver.find_element_by_name("testTimeCounter").text
            ws.write(row_num, 29, q3_end_time_iteration_2)
            q3_end_iteration_2 = time.strptime(q3_end_time_iteration_2, '%H:%M:%S')
            q3_end_time_in_seconds_iteration_2 = (q3_end_iteration_2.tm_hour * 3600) + (
                    q3_end_iteration_2.tm_min * 60) + q3_end_iteration_2.tm_sec
            print("q3_end_time_iteration_2", q3_end_time_iteration_2, "q3_end_time_in_seconds_iteration_2",
                  q3_end_time_in_seconds_iteration_2)

            q3_time_spent_iteration_2 = q3_start_time_in_seconds_iteration_2 - q3_end_time_in_seconds_iteration_2
            ws.write(row_num, 30, q3_time_spent_iteration_2)
            print("Second iteration q3 time spent", q3_time_spent_iteration_2)
            total_time_spent_in_q3 = q3_time_spent_iteration_1 + q3_time_spent_iteration_2
            if q3_expected_time_spent == total_time_spent_in_q3 or (
                    q3_expected_time_spent + 1) == total_time_spent_in_q3 or (
                    q3_expected_time_spent + 2) == total_time_spent_in_q3:
                ws.write(row_num, 32, total_time_spent_in_q3, self.__style3)
                row_status.append("Pass")
            else:
                ws.write(row_num, 32, total_time_spent_in_q3, self.__style2)
                row_status.append("Fail")

            # ----------------------------------------------------------------------------------------------------------
            #  Q4 Time calculation - Iteration 2
            # ----------------------------------------------------------------------------------------------------------
            self.driver.find_element_by_name("btnNext").click()
            q4_start_time_iteration_2 = self.driver.find_element_by_name("testTimeCounter").text
            ws.write(row_num, 38, q4_start_time_iteration_2)
            q4_start_iteration_2 = time.strptime(q4_start_time_iteration_2, '%H:%M:%S')
            q4_start_time_in_seconds_iteration_2 = (q4_start_iteration_2.tm_hour * 3600) + (
                    q4_start_iteration_2.tm_min * 60) + q4_start_iteration_2.tm_sec
            print("q4_start_time_iteration_2", q4_start_time_iteration_2, "q4_start_time_in_seconds_iteration_2",
                  q4_start_time_in_seconds_iteration_2)
            print("time.sleep(q4)", g2_default_time_spent)
            time.sleep(g2_default_time_spent)

            q4_end_time_iteration_2 = self.driver.find_element_by_name("testTimeCounter").text
            ws.write(row_num, 39, q4_end_time_iteration_2)
            q4_end_iteration_2 = time.strptime(q4_end_time_iteration_2, '%H:%M:%S')
            q4_end_time_in_seconds_iteration_2 = (q4_end_iteration_2.tm_hour * 3600) + (
                    q4_end_iteration_2.tm_min * 60) + q4_end_iteration_2.tm_sec
            print("q4_end_time_iteration_2", q4_end_time_iteration_2, "q4_end_time_in_seconds_iteration_2",
                  q4_end_time_in_seconds_iteration_2)

            q4_time_spent_iteration_2 = q4_start_time_in_seconds_iteration_2 - q4_end_time_in_seconds_iteration_2
            ws.write(row_num, 40, q4_time_spent_iteration_2)
            print("Second iteration q4 time spent", q4_time_spent_iteration_2)
            total_time_spent_in_q4 = q4_time_spent_iteration_1 + q4_time_spent_iteration_2
            if q4_expected_time_spent == total_time_spent_in_q4:
                ws.write(row_num, 42, total_time_spent_in_q4, self.__style3)
                row_status.append("Pass")
            else:
                ws.write(row_num, 42, total_time_spent_in_q4, self.__style2)
                row_status.append("Fail")

            # ----------------------------------------------------------------------------------------------------------
            #  Q5 Time calculation - Iteration 2
            # ----------------------------------------------------------------------------------------------------------
            self.driver.find_element_by_name("btnNext").click()
            q5_start_time_iteration_2 = self.driver.find_element_by_name("testTimeCounter").text
            ws.write(row_num, 47, q5_start_time_iteration_2)
            q5_start_iteration_2 = time.strptime(q5_start_time_iteration_2, '%H:%M:%S')
            q5_start_time_in_seconds_iteration_2 = (q5_start_iteration_2.tm_hour * 3600) + (
                    q5_start_iteration_2.tm_min * 60) + q5_start_iteration_2.tm_sec
            print("q5_start_time_iteration_2", q5_start_time_iteration_2, "q5_start_time_in_seconds_iteration_2",
                  q5_start_time_in_seconds_iteration_2)
            print("time.sleep(g2)", g2_default_time_spent)
            time.sleep(g2_default_time_spent)
            q5_end_time_iteration_2 = self.driver.find_element_by_name("testTimeCounter").text
            self.driver.find_element_by_name("btnSubmit").click()

            ws.write(row_num, 48, q5_end_time_iteration_2)
            self.driver.find_element_by_xpath("//div[8]/div/div/div[3]/button[1]").click()
            q5_end_iteration_2 = time.strptime(q5_end_time_iteration_2, '%H:%M:%S')
            q5_end_time_in_seconds_iteration_2 = (q5_end_iteration_2.tm_hour * 3600) + (
                    q5_end_iteration_2.tm_min * 60) + q5_end_iteration_2.tm_sec
            print("q5_end_time_iteration_2", q5_end_time_iteration_2, "q5_end_time_in_seconds_iteration_2",
                  q5_end_time_in_seconds_iteration_2)

            q5_time_spent_iteration_2 = q5_start_time_in_seconds_iteration_2 - q5_end_time_in_seconds_iteration_2
            ws.write(row_num, 49, q5_time_spent_iteration_2)
            print("Second iteration q5 time spent", q5_time_spent_iteration_2)
            total_time_spent_in_q5 = q5_time_spent_iteration_1 + q5_time_spent_iteration_2
            if q5_expected_time_spent == total_time_spent_in_q5 or (
                    q5_expected_time_spent + 1) == total_time_spent_in_q5 or (
                    q5_expected_time_spent + 2) == total_time_spent_in_q5:
                ws.write(row_num, 51, total_time_spent_in_q5, self.__style3)
                row_status.append("Pass")
            else:
                ws.write(row_num, 51, total_time_spent_in_q5, self.__style2)
                row_status.append("Fail")

            # ----------------------------------------------------------------------------------------------------------
            #  Evaluate online assessment for candidate
            # ----------------------------------------------------------------------------------------------------------

            header3 = {"content-type": "application/json", "X-AUTH-TOKEN": self.NTokenVal}
            data5 = {"testId": test_id, "candidateIds": [candidate_id]}
            requests.post("https://amsin.hirepro.in/py/assessment/eval/api/v1/eval-online-assessment/",
                          headers=header3,
                          data=json.dumps(data5, default=str), verify=True)
            time.sleep(20)

            # ----------------------------------------------------------------------------------------------------------
            #  Test Time spent calculation
            # ----------------------------------------------------------------------------------------------------------
            print("Time Spent in Q1", total_time_spent_in_q1, "Sec")
            print("Time Spent in Q2", total_time_spent_in_q2, "Sec")
            print("Time Spent in Q3", total_time_spent_in_q3, "Sec")
            print("Time Spent in Q4", total_time_spent_in_q4, "Sec")
            print("Time Spent in Q5", total_time_spent_in_q5, "Sec")
            g1_time_spent = total_time_spent_in_q1 + total_time_spent_in_q2 + total_time_spent_in_q3
            print("Time Spent in G1", g1_time_spent)
            g2_time_spent = total_time_spent_in_q4 + total_time_spent_in_q5
            print("Time Spent in G2", g2_time_spent)
            total_actual_test_time_spent = g1_time_spent + g2_time_spent
            print("Time Spent in Test", total_actual_test_time_spent)
            if total_expected_test_time_spent == total_actual_test_time_spent:
                ws.write(row_num, 54, total_actual_test_time_spent, self.__style3)
                row_status.append("Pass")
            else:
                ws.write(row_num, 54, total_actual_test_time_spent, self.__style2)
                row_status.append("Fail")
            try:
                conn = mysql.connector.connect(host='35.154.36.218',
                                               database='appserver_core',
                                               user='hireprouser',
                                               password='tech@123')
                cursor = conn.cursor()
                cursor.execute("select id from Test_users where test_id = 7518 and candidate_id = %d;" % candidate_id)
                data = cursor.fetchall()
                test_user_id = int(data[0][0])
                cursor.execute("select id, time_spent from Test_users where test_id = 7518 and candidate_id = %d;" % candidate_id)
                data = cursor.fetchall()
                test_time_spent = data[0][1]

                cursor.execute("select question_id, time_spent from test_results where testuser_id = %d;" % test_user_id)
                print(data)
                question_wise_time_spent = cursor.fetchall()
                question_wise_time_spent = dict(question_wise_time_spent)
                print(question_wise_time_spent)

                q1_time_spent_db = question_wise_time_spent.get(int(q1_id))
                print(q1_time_spent_db)
                if int(q1_expected_time_spent) == int(q1_time_spent_db):
                    ws.write(row_num, 15, q1_time_spent_db, self.__style3)
                    row_status.append("Pass")
                else:
                    ws.write(row_num, 15, q1_time_spent_db, self.__style2)
                    row_status.append("Fail")

                q2_time_spent_db = question_wise_time_spent.get(int(q2_id))
                print(q2_time_spent_db)
                if int(q2_expected_time_spent) == int(q2_time_spent_db):
                    ws.write(row_num, 24, q2_time_spent_db, self.__style3)
                    row_status.append("Pass")
                else:
                    ws.write(row_num, 24, q2_time_spent_db, self.__style2)
                    row_status.append("Fail")

                q3_time_spent_db = question_wise_time_spent.get(int(q3_id))
                print(q3_time_spent_db)
                if int(q3_expected_time_spent) == int(q3_time_spent_db):
                    ws.write(row_num, 33, q3_time_spent_db, self.__style3)
                    row_status.append("Pass")
                else:
                    ws.write(row_num, 33, q3_time_spent_db, self.__style2)
                    row_status.append("Fail")

                q4_time_spent_db = question_wise_time_spent.get(int(q4_id))
                print(q4_time_spent_db)
                if int(q4_expected_time_spent) == int(q4_time_spent_db):
                    ws.write(row_num, 43, q4_time_spent_db, self.__style3)
                    row_status.append("Pass")
                else:
                    ws.write(row_num, 43, q4_time_spent_db, self.__style2)
                    row_status.append("Fail")

                q5_time_spent_db = question_wise_time_spent.get(int(q5_id))
                print(q5_time_spent_db)
                if int(q5_expected_time_spent) == int(q5_time_spent_db):
                    ws.write(row_num, 52, q5_time_spent_db, self.__style3)
                    row_status.append("Pass")
                else:
                    ws.write(row_num, 52, q5_time_spent_db, self.__style2)
                    row_status.append("Fail")

                if int(total_expected_test_time_spent) == int(test_time_spent):
                    ws.write(row_num, 55, test_time_spent, self.__style3)
                    row_status.append("Pass")
                else:
                    ws.write(row_num, 55, test_time_spent, self.__style2)
                    row_status.append("Fail")
            except:
                print("Exception")
            finally:
                conn.close()
            if ("Fail" in row_status):
                ws.write(row_num, 1, "Fail", self.__style2)
                status.append("Fail")
            else:
                ws.write(row_num, 1, "Pass", self.__style3)
                status.append("Pass")
            wb_result.save(
                "/home/testingteam/hirepro_automation/API-Automation/Output Data/Assessment/TimerCheckResult.xls")
            n += 1
            row_num += 1
        if ("Fail" in status):
            ws.write(1, 0, "Fail", self.__style2)
        else:
            ws.write(1, 0, "Pass", self.__style3)
        wb_result.save(
            "/home/testingteam/hirepro_automation/API-Automation/Output Data/Assessment/TimerCheckResult.xls")

if __name__ == '__main__':
    unittest.main()
