import unittest
from collections import OrderedDict
import datetime
import xlwt
import time
import requests
import random
import json
from Config import Api
from scripts.assessment import WebConfig
from selenium import webdriver
from Config.read_excel import *


class ClientQuestionRandomization(unittest.TestCase):
    @classmethod
    def setUp(cls):
        cls.driver = webdriver.Chrome(WebConfig.CHROME_DRIVER)
        cls.driver.implicitly_wait(30)
        cls.driver.maximize_window()
        cls.driver.get(WebConfig.ONLINE_ASSESSMENT_LOGIN_URL)
        return cls.driver

    def test_ClientQuestionRandomization(self):
        now = datetime.datetime.now()
        __current_DateTime = now.strftime("%d/%m/%Y")
        answers = ['A', 'B', 'C', 'D']
        # --------------------------------------------------------------------------------------------------------------
        # CSS to differentiate Correct and Wrong data in Excel
        # --------------------------------------------------------------------------------------------------------------
        self.__style0 = xlwt.easyxf(
            'font: name Times New Roman, color-index black, bold on; pattern: pattern solid, fore-colour gold; border: left thin,right thin,top thin,bottom thin')
        self.__style1 = xlwt.easyxf(
            'font: name Times New Roman, color-index black, bold off; border: left thin,right thin,top thin,bottom thin')
        self.__style2 = xlwt.easyxf(
            'font: name Times New Roman, color-index red, bold on; border: left thin,right thin,top thin,bottom thin')
        self.__style3 = xlwt.easyxf(
            'font: name Times New Roman, color-index green, bold on; border: left thin,right thin,top thin,bottom thin')
        self.__style4 = xlwt.easyxf(
            'font: name Times New Roman, color-index black, bold off; pattern: pattern solid, fore-colour light_yellow; border: left thin,right thin,top thin,bottom thin')
        self.__style5 = xlwt.easyxf(
            'font: name Times New Roman, color-index black, bold off; pattern: pattern solid, fore-colour yellow; border: left thin,right thin,top thin,bottom thin')
        # mycursor.execute('delete from test_result_infos where testresult_id in (select id from test_results where testuser_id in (select id from test_users where test_id = '+i+' and login_time is not null));')
        # --------------------------------------------------------------------------------------------------------------
        # Read from Excel
        # --------------------------------------------------------------------------------------------------------------
        excel_read_obj.excel_read(
            '/home/testingteam/hirepro_automation/API-Automation/Input Data/Assessment/ClientQuestionRandomization.xls', 0)
        self.xls_values = excel_read_obj.details
        wb_result = xlwt.Workbook()
        self.ws = wb_result.add_sheet('questionRandomization', cell_overwrite_ok=True)
        col_index = 0
        # -------------------------Need to change file name onle here based on Test level configuration in UI---------------------
        self.file = open(
            "/home/testingteam/hirepro_automation/API-Automation/Output Data/Assessment/Client_Section_Random_Check.html",
            "wt")
        # file name need to be changed based on Test level configuration (
        # Client_Test_Random_Check.html or
        # Client_Group_Random_Check.html or
        # Client_Section_Random_Check.html))
        self.file.write("""<html>
                <head>
                <title>Automation Results</title>
                <style>
                h1 {
                    color: #0e8eab;
                    text-align: left;
                    font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
                }
                .div-h1 {
                    position: absolute;
                    overflow: hidden;
                    top: 0;
                    width: auto;
                    height: 100px;
                    text-align: center;
                }
                .div-overalldata {
                    position: absolute;
                    top: 80px;
                    width: 600px;
                    height: auto;
                    text-align: left;
                    font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
                }
                .label {
                    color: #0e8eab;
                    font-family: Arial;
                    font-size: 14pt;
                    font-weight: bold;
                }
                .value {
                    color: black;
                    font-family: Arial;
                    font-size: 14pt;
                }
                .valuePass {
                    color: green;
                    font-family: Arial;
                    animation: blinkingTextPass 0.8s infinite;
                    font-weight: bold;
                    font-size: 20pt;
                }       
                @keyframes blinkingTextPass{
                    0%{     color: green; font-size: 0pt;  }
                    50%{    color: lightgreen; }
                    100%{   color: green; font-size: 14pt; } 
                }
                .valueFail {
                    color: red;
                    font-family: Arial;
                    animation: blinkingTextFail 0.8s infinite;
                    font-weight: bold;
                    font-size: 20px;
                }
                @keyframes blinkingTextFail{
                    0%{     color: red; font-size: 0pt;   }
                    50%{    color: orange; }
                    100%{   color: red; font-size: 14pt;  }
                }
                .zui-table {
                    border: none;
                    border-right: solid 1px #DDEFEF;
                    border-collapse: separate;
                    border-spacing: 0;
                    font: normal 13px Arial, sans-serif;
                    width: 100%
                }
                .zui-table thead th {
                    border-left: solid 1px white;
                    border-bottom: solid 1px #DDEFEF;
                    background-color: #0e8eab;
                    color: white;
                    padding: 10px;
                    text-align: left;
                    white-space: nowrap;
                }
                .zui-table tbody td {
                    border-left: solid 1px #DDEFEF;
                    border-right: solid 1px #DDEFEF;
                    border-bottom: solid 1px #DDEFEF;
                    padding: 10px;
                    white-space: nowrap;
                }
                .td-pass {
                    color: green;
                }
                .td-fail {
                    color: red;
                    font-weight: bold;
                }
                tr:nth-child(even){background-color: #f2f2f2;}      
                tr:hover {background-color: #ddd; border-collapse: collapse;}
                @media all{
                    table tr th:nth-child(2),
                    table tr td:nth-child(2),
                    table tr th:nth-child(5),
                    table tr td:nth-child(5),
                    table tr th:nth-child(6),
                    table tr td:nth-child(6){
                        display: none;
                    }
                }
                .zui-wrapper {
                    position: relative;
                    top: 180px;
                    width: 100%;
                    height: 100%;
                }
                .zui-scroller {
                    margin-left: 141px;
                    overflow-x: scroll;
                    overflow-y: visible;
                    padding-bottom: 5px;
                }
                .zui-table .zui-sticky-col {
                    border-left: solid 1px #DDEFEF;
                    border-right: solid 1px #DDEFEF;
                    left: 0;
                    position: absolute;
                    top: auto;
                    width: 120px;
                }
                .zui-table .zui-sticky-col-pass {
                    border-left: solid 1px #DDEFEF;
                    border-right: solid 1px #DDEFEF;
                    left: 0;
                    position: absolute;
                    top: auto;
                    width: 120px;
                    color:green;
                    font-weight: bold;
                }
                .zui-table .zui-sticky-col-fail {
                    border-left: solid 1px #DDEFEF;
                    border-right: solid 1px #DDEFEF;
                    left: 0;
                    position: absolute;
                    top: auto;
                    width: 120px;
                    color:red;
                    font-weight: bold;
                }
                .zui-table1 {
                    border: none;
                    border-right: solid 1px #DDEFEF;
                    border-collapse: separate;
                    border-spacing: 0;
                    font: normal 13px Arial, sans-serif;
                    width: 100%
                }
                .zui-table1 thead th {
                    border-left: solid 1px white;
                    border-bottom: solid 1px #DDEFEF;
                    background-color: #0e8eab;
                    color: white;
                    padding: 10px;
                    text-align: left;
                    white-space: nowrap;
                }
                .zui-table1 tbody td {
                    border-left: solid 1px #DDEFEF;
                    border-right: solid 1px #DDEFEF;
                    border-bottom: solid 1px #DDEFEF;
                    padding: 10px;
                    white-space: nowrap;
                }
                .zui-wrapper1 {
                    width: 50%;
                    position: absolute;
                    top: 2%;
                    left: 49%;
                }
                </style>
                <div class="div-h2">
        
                    <h1>Client Side Section Level Randomization Report</h1></div>
                </head>
                <body style="overflow: hidden;">
                <div class="zui-wrapper">
                <div class="zui-scroller"><table class="zui-table"><thead><tr>""")
        #<h1>Client Side Test Level Randomization Report</h1></div> Need to change Title as mentioned in file name
        for xls_headers in excel_read_obj.headers_available_in_excel:
            self.ws.write(0, col_index, xls_headers, self.__style0)
            self.file.write(("""<th>""" + str(xls_headers) + """</th>"""))
            col_index += 1
        self.file.write("""<th class="zui-sticky-col">Status</th>""")
        self.file.write("""</tr></thead><tbody>""")
        self.cellPos = 0
        self.rownum = 1
        overall_status = []
        for login_details in self.xls_values:
            candidateId = int(login_details.get('Candidate Id'))
            testId = int(login_details.get('Test Id'))
            loginId = login_details.get('Login Id')
            password = login_details.get('Password')
            self.ws.write(self.rownum, 0, "UI Ite 1", self.__style4)
            self.ws.write(self.rownum, 2, candidateId, self.__style1)
            self.ws.write(self.rownum, 3, testId, self.__style1)
            self.ws.write(self.rownum, 4, loginId, self.__style1)
            self.ws.write(self.rownum, 5, "******", self.__style1)
            self.file.write("""<tr><td>"UI Ite 1"</td><td></td><td>""" + str(candidateId) + """</td><td>""" + str(
                testId) + """</td><td>""" + str(loginId) + """</td><td>******</td>""")
            self.driver.get(WebConfig.ONLINE_ASSESSMENT_LOGIN_URL)
            time.sleep(2)
            self.driver.find_element_by_xpath("//div[8]/div/div/div[2]/input").send_keys(WebConfig.ALIAS)
            time.sleep(3)
            self.driver.find_element_by_xpath("//div[8]/div/div/div[3]/div[2]/button").click()
            time.sleep(2)
            self.driver.find_element_by_name("btnLogin").click()
            time.sleep(2)
            self.driver.switch_to.window(self.driver.window_handles[0])
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])
            time.sleep(2)
            self.driver.find_element_by_name("loginUsername").send_keys(loginId)
            self.driver.find_element_by_name("loginPassword").send_keys(password)
            time.sleep(1)
            self.driver.find_element_by_name("btnLogin").click()
            time.sleep(7)
            self.driver.find_element_by_xpath("//*[@id='start-test-col']/div//label").click()
            time.sleep(1)
            self.driver.find_element_by_name("btnStartTest").click()
            print("Start Test")
            time.sleep(7)
            current_vs_totalquestion = self.driver.find_element_by_name("currentTotalq").text
            current_question = current_vs_totalquestion.split('/')
            current_question = int(current_question[0])
            print(type(current_question))
            colnum = 6
            Ite1_data = []
            for i in range(current_question, 13):
                group_name = self.driver.find_element_by_name("grpName").text
                Ite1_data.append(group_name)
                self.file.write("""<td>""" + str(group_name) + """</td>""")
                self.ws.write(self.rownum, colnum, group_name, self.__style1)
                print("Group Name - ", group_name)
                section_name = self.driver.find_element_by_name("secName").text
                Ite1_data.append(section_name)
                self.file.write("""<td>""" + str(section_name) + """</td>""")
                self.ws.write(self.rownum, colnum + 1, section_name, self.__style1)
                print("Section Name - ", section_name)
                question_string = self.driver.find_element_by_xpath(
                    "//div[3]/div/question-template/ng-include/div[1]/div[1]/div/div/div/div[2]").text
                Ite1_data.append(question_string)
                self.file.write("""<td>""" + str(question_string) + """</td>""")
                self.ws.write(self.rownum, colnum + 2, question_string, self.__style1)
                print("Question String - ", question_string)
                option1_string = self.driver.find_element_by_xpath("//*[@id='answerSection']/dl[1]/dd/label").text
                Ite1_data.append(option1_string)
                self.file.write("""<td>""" + str(option1_string) + """</td>""")
                self.ws.write(self.rownum, colnum + 3, option1_string, self.__style1)
                print(option1_string)
                option2_string = self.driver.find_element_by_xpath("//*[@id='answerSection']/dl[2]/dd/label").text
                Ite1_data.append(option2_string)
                self.file.write("""<td>""" + str(option2_string) + """</td>""")
                self.ws.write(self.rownum, colnum + 4, option2_string, self.__style1)
                print(option2_string)
                option3_string = self.driver.find_element_by_xpath("//*[@id='answerSection']/dl[3]/dd/label").text
                Ite1_data.append(option3_string)
                self.file.write("""<td>""" + str(option3_string) + """</td>""")
                self.ws.write(self.rownum, colnum + 5, option3_string, self.__style1)
                print(option3_string)
                option4_string = self.driver.find_element_by_xpath("//*[@id='answerSection']/dl[4]/dd/label").text
                Ite1_data.append(option4_string)
                self.file.write("""<td>""" + str(option4_string) + """</td>""")
                self.ws.write(self.rownum, colnum + 6, option4_string, self.__style1)
                print(option4_string)
                candidate_ans = random.choice(answers)
                if candidate_ans == option1_string:
                    self.driver.find_element_by_xpath("//*[@id='answerSection']/dl[1]/dt/label/span").click()
                    self.ws.write(self.rownum, colnum + 7, candidate_ans, self.__style1)
                    self.file.write("""<td>""" + str(option1_string) + """</td>""")
                    Ite1_data.append(option1_string)
                elif candidate_ans == option2_string:
                    self.driver.find_element_by_xpath("//*[@id='answerSection']/dl[2]/dt/label/span").click()
                    self.ws.write(self.rownum, colnum + 7, candidate_ans, self.__style1)
                    self.file.write("""<td>""" + str(option2_string) + """</td>""")
                    Ite1_data.append(option2_string)
                elif candidate_ans == option3_string:
                    self.driver.find_element_by_xpath("//*[@id='answerSection']/dl[3]/dt/label/span").click()
                    self.ws.write(self.rownum, colnum + 7, candidate_ans, self.__style1)
                    self.file.write("""<td>""" + str(option3_string) + """</td>""")
                    Ite1_data.append(option3_string)
                else:
                    self.driver.find_element_by_xpath("//*[@id='answerSection']/dl[4]/dt/label/span").click()
                    self.ws.write(self.rownum, colnum + 7, candidate_ans, self.__style1)
                    self.file.write("""<td>""" + str(option4_string) + """</td>""")
                    Ite1_data.append(option4_string)
                if i < 12:
                    self.driver.find_element_by_name("btnNext").click()
                colnum += 8
            self.file.write("""</tr>""")
            print(Ite1_data)
            self.driver.refresh()
            self.driver.get(WebConfig.ONLINE_ASSESSMENT_LOGIN_URL)
            time.sleep(2)
            self.driver.find_element_by_xpath("//div[8]/div/div/div[2]/input").send_keys(WebConfig.ALIAS)
            time.sleep(2)
            self.driver.find_element_by_xpath("//div[8]/div/div/div[3]/div[2]/button").click()
            time.sleep(3)
            self.driver.find_element_by_name("btnLogin").click()
            time.sleep(1)
            self.driver.switch_to.window(self.driver.window_handles[0])
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])
            time.sleep(2)
            self.driver.find_element_by_name("loginUsername").send_keys(loginId)
            self.driver.find_element_by_name("loginPassword").send_keys(password)
            time.sleep(1)
            self.driver.find_element_by_name("btnLogin").click()
            time.sleep(3)

            # ----------------------------------------------------------------------------------------------------------
            #  Login to AMS
            # ----------------------------------------------------------------------------------------------------------
            crpo_login_header = {"content-type": "application/json"}
            data0 = {"LoginName": "charuk", "Password": "charuk", "TenantAlias": "at", "UserName": "Charuk"}
            response = requests.post(Api.login_user, headers=crpo_login_header, data=json.dumps(data0), verify=True)
            self.TokenVal = response.json()
            self.NTokenVal = self.TokenVal.get("Token")
            print(self.NTokenVal)

            # ----------------------------------------------------------------------------------------------------------
            #  Reactivate Test user login
            # ----------------------------------------------------------------------------------------------------------
            eval_online_assessment_header = {"content-type": "application/json", "X-AUTH-TOKEN": self.NTokenVal}
            data5 = {"candidateIds": [candidateId], "testId": testId}
            requests.post("https://amsin.hirepro.in/py/assessment/testuser/api/v1/reActivateLogin/",
                          headers=eval_online_assessment_header,
                          data=json.dumps(data5, default=str), verify=True)
            time.sleep(2)
            self.ws.write(self.rownum + 1, 0, "UI Ite 2", self.__style5)
            self.ws.write(self.rownum + 1, 2, candidateId, self.__style1)
            self.ws.write(self.rownum + 1, 3, testId, self.__style1)
            self.ws.write(self.rownum + 1, 4, loginId, self.__style1)
            self.ws.write(self.rownum + 1, 5, "******", self.__style1)
            self.file.write("""<tr><td>"UI Ite 2"</td>""")
            self.driver.find_element_by_name("btnLogin").click()
            time.sleep(7)
            self.driver.find_element_by_xpath("//*[@id='start-test-col']/div//label").click()
            time.sleep(1)
            self.driver.find_element_by_name("btnStartTest").click()
            time.sleep(7)

            current_vs_totalquestion = self.driver.find_element_by_name("currentTotalq").text
            current_question = current_vs_totalquestion.split('/')
            current_question = int(current_question[0])
            print(current_question)
            # colnum = 5
            Ite2_data = []
            for i in range(current_question, 13):
                group_name = self.driver.find_element_by_name("grpName").text
                Ite2_data.append(group_name)
                # self.ws.write(self.rownum+1, colnum, group_name, self.__style1)
                print("Group Name - ", group_name)
                section_name = self.driver.find_element_by_name("secName").text
                Ite2_data.append(section_name)
                # self.ws.write(self.rownum+1, colnum+1, section_name, self.__style1)
                print("Section Name - ", section_name)
                question_string = self.driver.find_element_by_xpath(
                    "//div[3]/div/question-template/ng-include/div[1]/div[1]/div/div/div/div[2]").text
                Ite2_data.append(question_string)
                # self.ws.write(self.rownum+1, colnum+2, question_string, self.__style1)
                print("Question String - ", question_string)
                option1_string = self.driver.find_element_by_xpath("//*[@id='answerSection']/dl[1]/dd/label").text
                Ite2_data.append(option1_string)
                # self.ws.write(self.rownum+1, colnum+3, option1_string, self.__style1)
                print(option1_string)
                option2_string = self.driver.find_element_by_xpath("//*[@id='answerSection']/dl[2]/dd/label").text
                Ite2_data.append(option2_string)
                # self.ws.write(self.rownum+1, colnum+4, option2_string, self.__style1)
                print(option2_string)
                option3_string = self.driver.find_element_by_xpath("//*[@id='answerSection']/dl[3]/dd/label").text
                Ite2_data.append(option3_string)
                # self.ws.write(self.rownum+1, colnum+5, option3_string, self.__style1)
                print(option3_string)
                option4_string = self.driver.find_element_by_xpath("//*[@id='answerSection']/dl[4]/dd/label").text
                Ite2_data.append(option4_string)
                # self.ws.write(self.rownum+1, colnum+6, option4_string, self.__style1)
                print(option4_string)
                if self.driver.find_element_by_xpath(
                        "//div[3]/div/question-template/ng-include/div[1]/div[2]/div/div/div/div[2]/dl[1]/dt/label/input").is_selected():
                    # self.ws.write(self.rownum + 1, colnum+7, option1_string, self.__style1)
                    Ite2_data.append(option1_string)
                elif self.driver.find_element_by_xpath(
                        "//div[3]/div/question-template/ng-include/div[1]/div[2]/div/div/div/div[2]/dl[2]/dt/label/input").is_selected():
                    # self.ws.write(self.rownum + 1, colnum+7, option2_string, self.__style1)
                    Ite2_data.append(option2_string)
                elif self.driver.find_element_by_xpath(
                        "//div[3]/div/question-template/ng-include/div[1]/div[2]/div/div/div/div[2]/dl[3]/dt/label/input").is_selected():
                    # self.ws.write(self.rownum + 1, colnum+7, option3_string, self.__style1)
                    Ite2_data.append(option3_string)
                else:
                    # self.ws.write(self.rownum + 1, colnum+7, option4_string, self.__style1)
                    Ite2_data.append(option4_string)
                if i < 12:
                    self.driver.find_element_by_name("btnNext").click()
                # colnum += 8
            print(Ite2_data)

            if Ite1_data == Ite2_data:
                self.ws.write(self.rownum + 1, 1, "Pass", self.__style3)
                overall_status.append("Pass")
                self.file.write(
                    """<td></td><td>""" + str(candidateId) + """</td><td>""" + str(
                        testId) + """</td><td>""" + str(loginId) + """</td><td>******</td>""")
            else:
                self.ws.write(self.rownum + 1, 1, "Fail", self.__style2)
                overall_status.append("Fail")
                self.file.write("""<td></td><td>""" + str(candidateId) + """</td><td>""" + str(
                    testId) + """</td><td>""" + str(loginId) + """</td><td>******</td>""")

            colnum = 6
            for element in Ite2_data:
                if element in Ite1_data:
                    self.ws.write(self.rownum + 1, colnum, element, self.__style3)
                    self.file.write("""<td class="td-pass">""" + str(element) + """</td>""")
                else:
                    self.ws.write(self.rownum + 1, colnum, element, self.__style2)
                    self.file.write("""<td class="td-fail">""" + str(element) + """</td>""")
                colnum += 1

            self.rownum += 3
            if Ite1_data == Ite2_data:
                self.file.write("""<td class="zui-sticky-col-pass"><b>Pass</b></td>""")
            else:
                self.file.write("""<td class="zui-sticky-col-fail"><b>Fail</b></td>""")

            self.file.write("""</tr>""")

            # -------------------------Need to change file name onle here based on Test level configuration in UI---------------------

            wb_result.save(
                "/home/testingteam/hirepro_automation/API-Automation/Output Data/Assessment/Client_Section_Random_Check.xls")

            # file name need to be changed based on Test level configuration (
            # Client_Test_Random_Check.html or
            # Client_Group_Random_Check.html or
            # Client_Section_Random_Check.html))
        self.file.write("""</tbody></table></div></div><div class="div-overalldata"><span class="label">Execution Date:&nbsp;&nbsp;</span><span class="lable value">""" + str(
                __current_DateTime) + """</span></br></br>""")
        if ("Fail" in overall_status):
            self.file.write(
                """<span class="label">Overall Status:&nbsp;&nbsp;</span><span class="lable valueFail">FAIL</span>""")
        else:
            self.file.write(
                """<span class="label">Overall Status:&nbsp;&nbsp;</span><span class="lable valuePass">PASS</span>""")
        self.file.write("""</div>""")
        wb = xlrd.open_workbook("/home/testingteam/hirepro_automation/API-Automation/Output Data/Assessment/Client_Section_Random_Check.xls")
        wb_result = xlwt.Workbook()
        ws = wb_result.add_sheet('questionRandomization', cell_overwrite_ok=True)
        sh1 = wb.sheet_by_index(0)
        first_rows = sh1.row_values(1)
        base_grps = []
        base_secs = []
        base_ques = []
        for base_grp in range(6, len(first_rows), 8):
            base_grps.append(first_rows[base_grp])
        for base_sec in range(7, len(first_rows), 8):
            base_secs.append(first_rows[base_sec])
        for base_que in range(8, len(first_rows), 8):
            base_ques.append(first_rows[base_que])
        base_grp_sec = tuple(zip(base_grps, base_secs))
        base_sec_que = tuple(zip(base_secs, base_ques))
        base_grp_sec_que = tuple(zip(base_grp_sec, base_ques))

        base_grpwise_sec = []
        base_grps_set = list(OrderedDict.fromkeys(base_grps))
        for j in range(0, len(base_grps_set)):
            for i in range(0, len(base_grps)):
                if base_grps_set[j] == base_grp_sec[i][0]:
                    base_grpwise_sec.append(base_grp_sec[i])

        base_grpsecwise_que = []
        base_grpssecs_set = list(OrderedDict.fromkeys(base_grp_sec))
        for j in range(0, len(base_grpssecs_set)):
            for i in range(0, len(base_grp_sec)):
                if base_grpssecs_set[j] == base_grp_sec_que[i][0]:
                    base_grpsecwise_que.append(base_grp_sec_que[i])
        self.file.write("""<div class="zui-wrapper1"><table class="zui-table1"><thead><tr><th>Candidate Id</th><th></th><th>Group</th><th>Section</th><th></th><th></th><th>Question</th></tr></thead><tbody>""")
        grp_result = []
        sec_result = []
        que_result = []
        total_candidate = 0
        for row_n in range(4, sh1.nrows, 3):
            rw = sh1.row_values(row_n)
            candidate_id = int(rw[2])
            self.file.write("""<tr><td>"""+str(candidate_id)+"""</td>""")
            print("C - ", candidate_id)
            grps = []
            secs = []
            ques = []
            random_group = False
            random_section = False
            random_question = False
            row = sh1.row_values(row_n)
            for grp in range(6, len(row), 8):
                grps.append(row[grp])
            for sec in range(7, len(row), 8):
                secs.append(row[sec])
            for que in range(8, len(row), 8):
                ques.append(row[que])
            grp_sec = tuple(zip(grps, secs))
            sec_que = tuple(zip(secs, ques))
            grp_sec_que = tuple(zip(grp_sec, ques))
            for i in range(0, 12):
                if base_grps[i] != grps[i]:
                    random_group = True
            if base_grps != grps:
                grp_result.append(random_group)
                self.file.write("""<td></td><td>""" + str(random_group) + """</td>""")
                print("G - ", random_group)
            else:
                grp_result.append(random_group)
                self.file.write("""<td></td><td>""" + str(random_group) + """</td>""")
                print("G - ", random_group)

            grpwise_sec = []
            for j in range(0, len(base_grps_set)):
                for i in range(0, len(secs)):
                    if base_grps_set[j] == grp_sec[i][0]:
                        grpwise_sec.append(grp_sec[i])
            if base_grpwise_sec != grpwise_sec:
                random_section = True
                sec_result.append(random_section)
                self.file.write("""<td>""" + str(random_section) + """</td>""")
                print("S - ", random_section)
            else:
                sec_result.append(random_section)
                self.file.write("""<td>""" + str(random_section) + """</td>""")
                print("S - ", random_section)

            grpsecwise_que = []
            for j in range(0, len(base_grpssecs_set)):
                for i in range(0, len(grp_sec)):
                    if base_grpssecs_set[j] == grp_sec_que[i][0]:
                        grpsecwise_que.append(grp_sec_que[i])
            if base_grpsecwise_que != grpsecwise_que:
                random_question = True
                que_result.append(random_question)
                self.file.write("""<td></td><td></td><td>""" + str(random_question) + """</td>""")
                print("Q - ", random_question, "\n")
            else:
                que_result.append(random_question)
                self.file.write("""<td></td><td></td><td>""" + str(random_question) + """</td>""")
                print("Q - ", random_question, "\n")
            self.file.write("""</tr>""")
            total_candidate += 1
        self.file.write("""</tbody></table></div></body></html>""")
        for check in range(0, total_candidate):
            if grp_result[check] is True and sec_result[check] is True and que_result[check] is True:
                print("Test level randomization done")
                # ws.write(20, 1, "Test level randomization done")
                break
            elif grp_result[check] is False and sec_result[check] is True and que_result[check] is True:
                print("Group level randomization done")
                # ws.write(20, 1, "Group level randomization done")
                break
            elif grp_result[check] is False and sec_result[check] is False and que_result[check] is True:
                print("Section level randomization done")
                # ws.write(20, 1, "Section level randomization done")
                break
            else:
                print("Randomization not working")
                # ws.write(20, 1, "Randomization not working")
                break
        self.file.close()
