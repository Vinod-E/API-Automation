import json
import requests
import datetime
import random
import xlwt
from Config.read_excel import *
from Config import Api
from collections import OrderedDict


class QuestionRandomization:
    def test_QuestionRandomization(self):
        now = datetime.datetime.now()
        __current_DateTime = now.strftime("%d/%m/%Y")
        # --------------------------------------------------------------------------------------------------------------
        # Global Variables
        # --------------------------------------------------------------------------------------------------------------
        master_qp_que_ids = (
            110094, 110095, 110096, 110097, 110112, 59229, 59230, 105461, 106045, 108401, 109734, 109735, 109736,
            109737,
            109738, 109743, 109744, 109745, 109746, 109747, 109752, 109753, 109754, 109755, 109756, 109761, 109762,
            109763,
            109764, 109765, 109833, 109834, 109835, 109836, 109837, 110392, 110393, 110394, 110395, 110396)
        self.mcq_questions_ans = ("A", "B", "C", "D", "A", "B", "C", "D", "A", "B")
        self.subj_que_ans = (
            "This is answer 1", "This is answer 2", "This is answer 3", "This is answer 4", "This is answer 5",
            "This is answer 6", "This is answer 7", "This is answer 8", "This is answer 9", "This is answer 10")
        self.coding_que_ans = (
            "import java.io.*;\n\npublic class TestClass {\n    public static void main(String[] args) {\n        // Read input from STDIN; write output to STDOUT.\n        a = 10;\n    }\n}",
            "import java.io.*;\n\npublic class TestClass {\n    public static void main(String[] args) {\n        // Read input from STDIN; write output to STDOUT.\n        a = 20;\n    }\n}",
            "import java.io.*;\n\npublic class TestClass {\n    public static void main(String[] args) {\n        // Read input from STDIN; write output to STDOUT.\n        a = 30;\n    }\n}",
            "import java.io.*;\n\npublic class TestClass {\n    public static void main(String[] args) {\n        // Read input from STDIN; write output to STDOUT.\n        a = 40;\n    }\n}",
            "import java.io.*;\n\npublic class TestClass {\n    public static void main(String[] args) {\n        // Read input from STDIN; write output to STDOUT.\n        a = 50;\n    }\n}",
            "import java.io.*;\n\npublic class TestClass {\n    public static void main(String[] args) {\n        // Read input from STDIN; write output to STDOUT.\n        a = 60;\n    }\n}",
            "import java.io.*;\n\npublic class TestClass {\n    public static void main(String[] args) {\n        // Read input from STDIN; write output to STDOUT.\n        a = 70;\n    }\n}",
            "import java.io.*;\n\npublic class TestClass {\n    public static void main(String[] args) {\n        // Read input from STDIN; write output to STDOUT.\n        a = 80;\n    }\n}",
            "import java.io.*;\n\npublic class TestClass {\n    public static void main(String[] args) {\n        // Read input from STDIN; write output to STDOUT.\n        a = 90;\n    }\n}",
            "import java.io.*;\n\npublic class TestClass {\n    public static void main(String[] args) {\n        // Read input from STDIN; write output to STDOUT.\n        a = 100;\n    }\n}")
        self.coding_que_lang = ("C", "Java 8", "Java 7", "C++", "Python 2", "Python 3")

        # --------------------------------------------------------------------------------------------------------------
        # CSS to differentiate Correct and Wrong data in Excel
        # --------------------------------------------------------------------------------------------------------------
        self.__style0 = xlwt.easyxf(
            'font: name Times New Roman, color-index black, bold on; pattern: pattern solid, fore-colour light_yellow; border: left thin,right thin,top thin,bottom thin')
        self.__style1 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        self.__style2 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        self.__style3 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')

        # --------------------------------------------------------------------------------------------------------------
        # Read from Excel
        # --------------------------------------------------------------------------------------------------------------
        excel_read_obj.excel_read('/home/rajeshwar/D Drive/hirepro_automation/API-Automation/Input Data/Assessment/ServerQuestionRandomization.xls', 0)
        self.xls_values = excel_read_obj.details
        wb_result = xlwt.Workbook()
        self.ws = wb_result.add_sheet('questionRandomization', cell_overwrite_ok=True)
        col_index = 0
        self.file = open("/home/rajeshwar/D Drive/hirepro_automation/API-Automation/Output Data/Assessment/ServerQuestionRandomization_Check.html",
            "wt")

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
            top: 60px;
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
        @media all{
            table tr th:nth-child(1),
            table tr td:nth-child(1),
            table tr th:nth-child(2),
            table tr td:nth-child(2){
                display: none;
            }
        }
        tr:nth-child(even){background-color: #f2f2f2;}
                        
        tr:hover {background-color: #ddd; border-collapse: collapse;}
        .zui-wrapper {
            position: relative;
            top: 100px;
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
        </style>
        <div class="div-h2">
            <h1>Server Side Question Randomization Report</h1></div>
        </head>
        <body style="overflow: hidden;">
        <div class="zui-wrapper">
        <div class="zui-scroller"><table class="zui-table"><thead><tr>""")
        for xls_headers in excel_read_obj.headers_available_in_excel:
            self.ws.write(0, col_index, xls_headers, self.__style0)
            self.file.write(("""<th>""" + str(xls_headers) + """</th>"""))
            col_index += 1
        self.file.write("""<th class="zui-sticky-col">Status</th>""")
        self.file.write("""</tr></thead><tbody>""")
        self.loginToAMS()
        self.cellPos = 2
        self.rownum = 1
        self.cellPos1 = 2
        self.rownum1 = 1
        self.rownum2 = 1
        self.cellwriteposition1 = 2
        self.cellwriteposition2 = 3
        status = []
        for login_details in self.xls_values:
            candidateId = int(login_details.get('Candidate Id'))
            testId = int(login_details.get('Test Id'))
            loginId = login_details.get('Login Id')
            password = login_details.get('Password')
            self.ws.write(self.rownum, 2, candidateId, self.__style1)
            self.ws.write(self.rownum, 3, testId, self.__style1)
            self.ws.write(self.rownum, 4, loginId, self.__style1)
            self.ws.write(self.rownum, 5, "******", self.__style1)
            self.file.write("""<tr><td></td><td></td><td>""" + str(candidateId) + """</td><td>""" + str(
                testId) + """</td><td>""" + str(loginId) + """</td><td>******</td>""")
            # ----------------------------------------------------------------------------------------------------------
            # Login to HTML Test/Online Assessment
            # ----------------------------------------------------------------------------------------------------------
            self.loginToTest(loginId, password)
            self.loadTest()
            self.qIdqTypeCollection()
            self.ite1_qids = self.question_ids
            self.submit_test_result()
            self.loginToTest(loginId, password)
            self.reActivate_candidate(candidateId, testId)
            self.loginToTest(loginId, password)
            self.loadTest()
            self.qIdqTypeCollection()
            self.ite2_qids = self.question_ids
            self.ans_Ite1 = self.questionid_vs_answer
            print("ans_Ite1 - ", self.ans_Ite1)
            self.lang_Ite1 = self.questionid_vs_lang
            self.ans_Ite2 = self.get_test_result()
            print("ans_Ite2 - ", self.ans_Ite2)
            self.lang_Ite2 = self.seq_ids_vs_lang
            print("ans_Ite1", self.ans_Ite1)
            print("ans_Ite2", self.ans_Ite2)
            self.qid_writetoxls(self.ite1_qids, self.ite2_qids, self.ans_Ite1, self.ans_Ite2)
            self.codinglang_writexls(self.lang_Ite1, self.lang_Ite2)
            ite1 = []
            for item in range(6, len(self.xls_values), 2):
                ite1.append(self.xls_values[item])
            ite2 = []
            for item in range(7, len(self.xls_values), 2):
                ite2.append(self.xls_values[item])
            if ite1 == ite2:
                self.ws.write(self.rownum, 1, "Pass", self.__style3)
                self.file.write("""<td class="zui-sticky-col-pass">Pass</td></tr>""")
                status.append("Pass")
            else:
                self.ws.write(self.rownum, 1, "Fail", self.__style2)
                self.file.write("""<td class="zui-sticky-col-fail">Fail</td></tr>""")
                status.append("Fail")
            self.rownum += 1
            self.rownum1 += 1
            self.rownum2 += 1
            wb_result.save(
                "/home/rajeshwar/D Drive/hirepro_automation/API-Automation/Output Data/Assessment/Server_Que_Random_Check.xls")

        self.file.write("""</tbody></table></div></div>""")
        wb_result.save(
            "/home/rajeshwar/D Drive/hirepro_automation/API-Automation/Output Data/Assessment/Server_Que_Random_Check.xls")
        self.file.write(
            """<div class="div-overalldata"><span class="label">Execution Date:&nbsp;&nbsp;</span><span class="lable value">""" + str(
                __current_DateTime) + """</span></br></br>""")
        if ("Fail" in status):
            self.ws.write(1, 0, "Fail", self.__style2)
            self.file.write(
                """<span class="label">Overall_Status:&nbsp;&nbsp;</span><span class="lable valueFail">FAIL</span>""")
        else:
            self.ws.write(1, 0, "Pass", self.__style3)
            self.file.write(
                """<span class="label">Overall_Status:&nbsp;&nbsp;</span><span class="lable valuePass">PASS</span>""")
            self.file.write("""</div></body></html>""")

        wb_result.save(
            "/home/rajeshwar/D Drive/hirepro_automation/API-Automation/Output Data/Assessment/Server_Que_Random_Check.xls")

        self.file.close()

    def loginToTest(self, loginId, password):
        login_to_test_header = {"content-type": "application/json"}
        login_to_test_data = {"ClientSystemInfo": "Browser:chrome/60.0.3112.78,OS:Linux x86_64,IPAddress:10.0.3.83",
                              "IPAddress": "10.0.3.83", "IsOnlinePreview": False, "LoginName": loginId,
                              "Password": password,
                              "TenantAlias": "at"}
        login_to_test_request = requests.post(
            "https://amsin.hirepro.in/py/assessment/htmltest/api/v2/login_to_test/",
            headers=login_to_test_header, data=json.dumps(login_to_test_data), verify=True)
        self.Test_Login_response = login_to_test_request.json()
        self.Test_Login_TokenVal = self.Test_Login_response.get("Token")
        print("loginToTest", self.Test_Login_TokenVal)

    def loadTest(self):
        initiate_tua_header = {"content-type": "application/json", "X-AUTH-TOKEN": self.Test_Login_TokenVal}
        initiate_tua_data = {"debugTimeStamp": "2019-04-10T09:07:54.045Z"}
        requests.post("https://amsin.hirepro.in/py/assessment/htmltest/api/v1/initiate-tua/",
                      headers=initiate_tua_header,
                      data=json.dumps(initiate_tua_data, default=str), verify=True)

        loadtest_header = {"content-type": "application/json", "X-AUTH-TOKEN": self.Test_Login_TokenVal}
        loadtest_data = {
            "userAgent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.81 Safari/537.36",
            "isDeviceTypeDesktop": True}
        loadtest_res = requests.post("https://amsin.hirepro.in/py/assessment/htmltest/api/v1/loadtest/",
                                     headers=loadtest_header,
                                     data=json.dumps(loadtest_data, default=str), verify=True)
        self.loadTestResp = loadtest_res.json()
        print(self.loadTestResp['mandatoryGroups'])

    def qIdqTypeCollection(self):
        self.question_type = []
        self.questonwise_section = {}
        self.groups = []
        self.sections = []
        self.question_ids = []
        self.coding_questions_id = []
        for mandatoryGroups_index in range(0, len(self.loadTestResp['mandatoryGroups'])):
            self.groups.append(self.loadTestResp['mandatoryGroups'][mandatoryGroups_index]['id'])
            for sections_index in range(0,
                                        len(self.loadTestResp['mandatoryGroups'][mandatoryGroups_index]['sections'])):
                self.sections.append(
                    self.loadTestResp['mandatoryGroups'][mandatoryGroups_index]['sections'][sections_index]['id'])
                self.sectionid = \
                    self.loadTestResp['mandatoryGroups'][mandatoryGroups_index]['sections'][sections_index]['id']
                for questionDetails_index in range(0, len(
                        self.loadTestResp['mandatoryGroups'][mandatoryGroups_index]['sections'][sections_index][
                            'questionDetails'])):
                    if (self.loadTestResp['mandatoryGroups'][mandatoryGroups_index]['sections'][sections_index][
                        'questionDetails'][questionDetails_index]['typeOfQuestionText'] == "MCQ"
                            or
                            self.loadTestResp['mandatoryGroups'][mandatoryGroups_index]['sections'][
                                sections_index]['questionDetails'][questionDetails_index][
                                'typeOfQuestionText'] == "Boolean"
                            or
                            self.loadTestResp['mandatoryGroups'][mandatoryGroups_index]['sections'][
                                sections_index]['questionDetails'][questionDetails_index][
                                'typeOfQuestionText'] == "QA"
                            or
                            self.loadTestResp['mandatoryGroups'][mandatoryGroups_index]['sections'][
                                sections_index]['questionDetails'][questionDetails_index][
                                'typeOfQuestionText'] == "Coding"):
                        self.questionid = \
                            self.loadTestResp['mandatoryGroups'][mandatoryGroups_index]['sections'][sections_index][
                                'questionDetails'][
                                questionDetails_index]['id']
                        self.questiontype = \
                            self.loadTestResp['mandatoryGroups'][mandatoryGroups_index]['sections'][sections_index][
                                'questionDetails'][
                                questionDetails_index]['typeOfQuestionText']
                        self.question_ids.append(
                            self.loadTestResp['mandatoryGroups'][mandatoryGroups_index]['sections'][sections_index][
                                'questionDetails'][questionDetails_index]['id'])
                        self.question_type.append(
                            self.loadTestResp['mandatoryGroups'][mandatoryGroups_index]['sections'][sections_index][
                                'questionDetails'][questionDetails_index]['typeOfQuestionText'])
                        self.questonwise_section[self.questionid] = self.sectionid
                    else:
                        for childquestionid_index in range(0, len(
                                self.loadTestResp['mandatoryGroups'][mandatoryGroups_index]['sections'][sections_index][
                                    'questionDetails'][questionDetails_index]['childQuestions'])):
                            questionid = self.loadTestResp['mandatoryGroups'][mandatoryGroups_index]['sections'][
                                sections_index]['questionDetails'][questionDetails_index]['childQuestions'][
                                childquestionid_index]['id']
                            self.question_ids.append(
                                self.loadTestResp['mandatoryGroups'][mandatoryGroups_index]['sections'][
                                    sections_index]['questionDetails'][questionDetails_index][
                                    'childQuestions'][
                                    childquestionid_index]['id'])
                            self.question_type.append(
                                self.loadTestResp['mandatoryGroups'][mandatoryGroups_index]['sections'][
                                    sections_index]['questionDetails'][questionDetails_index][
                                    'childQuestions'][
                                    childquestionid_index]['typeOfQuestionText'])
                            self.questonwise_section[questionid] = self.sectionid
        print(self.question_ids)
        print(self.question_type)
        print("new variable for ans")
        print(self.questonwise_section)
        self.questionid_vs_questiontype = OrderedDict(zip(self.question_ids, self.question_type))
        for k, v in self.questionid_vs_questiontype.items():
            if self.questionid_vs_questiontype.get(k) == "Coding":
                self.coding_questions_id.append(k)
        print("questionid_vs_questiontype - ", self.questionid_vs_questiontype)
        print("Coding Question Ids - ", self.coding_questions_id)

    def get_test_result(self):
        get_testresult_header = {"content-type": "application/json", "X-AUTH-TOKEN": self.Test_Login_TokenVal}
        get_testresult_data = {}
        get_testresult_request = requests.post("https://amsin.hirepro.in/py/assessment/htmltest/api/v1/getTestResult/",
                                               headers=get_testresult_header,
                                               data=json.dumps(get_testresult_data, default=str), verify=True)
        self.get_testresult_Resp = get_testresult_request.json()
        self.qid_gettestresult = []
        self.codeqid_gettestresult = []
        self.qans_gettestresult = []
        self.qlang_gettestresult = []
        for qid_index in range(0, len(self.get_testresult_Resp['answers'])):
            self.qid_gettestresult.append(self.get_testresult_Resp['answers'][qid_index]['q'])
            self.qans_gettestresult.append(self.get_testresult_Resp['answers'][qid_index]['a'])
            if self.get_testresult_Resp['answers'][qid_index]['q'] in self.coding_questions_id:
                self.codeqid_gettestresult.append(self.get_testresult_Resp['answers'][qid_index]['q'])
                self.qlang_gettestresult.append(self.get_testresult_Resp['answers'][qid_index]['l'])

        self.questionid_vs_qlangs = OrderedDict(zip(self.codeqid_gettestresult, self.qlang_gettestresult))
        self.questionid_vs_answers = OrderedDict(zip(self.qid_gettestresult, self.qans_gettestresult))
        print("Id vs Ans", self.questionid_vs_answers)
        self.seq_ids_vs_ans = OrderedDict()
        self.seq_ids_vs_lang = OrderedDict()
        for id in self.question_ids:
            self.seq_ids_vs_ans[id] = self.questionid_vs_answers[id]
            if id in self.coding_questions_id:
                self.seq_ids_vs_lang[id] = self.questionid_vs_qlangs[id]
        print("Sequence Id vs Ans - ", self.seq_ids_vs_ans)
        print("Sequence Id vs Lang - ", self.seq_ids_vs_lang)
        return self.seq_ids_vs_ans

    def testResultCollection(self):
        self.candidate_testResultCollection = []
        candidate_qid = []
        candidate_ans_choice = []
        candidate_lang_choice = []
        for k, v in self.questionid_vs_questiontype.items():
            if v == "MCQ":
                candidate_ans = random.choice(self.mcq_questions_ans)
                self.candidate_testResultCollection.append(
                    {"q": k, "timeSpent": 3, "timeSpentOnTicker": 0, "secId": self.questonwise_section[k],
                     "a": candidate_ans})
                candidate_ans_choice.append(candidate_ans)
            elif v == "QA":
                candidate_ans = random.choice(self.subj_que_ans)
                self.candidate_testResultCollection.append(
                    {"q": k, "timeSpent": 8, "timeSpentOnTicker": 0, "secId": self.questonwise_section[k],
                     "a": candidate_ans, "isSubjective": True})
                candidate_ans_choice.append(candidate_ans)
            elif v == "Coding":
                candidate_qid.append(k)
                candidate_ans = random.choice(self.coding_que_ans)
                candidate_ans_lang = random.choice(self.coding_que_lang)
                self.candidate_testResultCollection.append(
                    {"q": k, "timeSpent": 25, "timeSpentOnTicker": 0, "secId": self.questonwise_section[k],
                     "a": candidate_ans, "l": candidate_ans_lang})
                candidate_ans_choice.append(candidate_ans)
                candidate_lang_choice.append(candidate_ans_lang)
            self.questionid_vs_answer = OrderedDict(zip(self.question_ids, candidate_ans_choice))
            self.questionid_vs_lang = OrderedDict(zip(candidate_qid, candidate_lang_choice))

    def submit_test_result(self):
        self.testResultCollection()
        # ----------------------------------------------------------------------------------------------------------
        #  Partial submit test
        # ----------------------------------------------------------------------------------------------------------
        submit_test_header = {"content-type": "application/json", "X-AUTH-TOKEN": self.Test_Login_TokenVal}
        submit_test_data = {"isPartialSubmission": True, "totalTimeSpent": 39,
                            "testResultCollection": self.candidate_testResultCollection,
                            "config": "{\"TimeStamp\":\"2018-03-13T07:28:55.933Z\"}"}

        requests.post("https://amsin.hirepro.in/py/assessment/htmltest/api/v1/submitTestResult/",
                      headers=submit_test_header, data=json.dumps(submit_test_data, default=str), verify=True)

    def loginToAMS(self):
        # ----------------------------------------------------------------------------------------------------------
        #  Login to AMS
        # ----------------------------------------------------------------------------------------------------------
        crpo_login_header = {"content-type": "application/json"}
        data0 = {"LoginName": "admin", "Password": "Assessment@1234", "TenantAlias": "at",
                 "UserName": "admin"}
        response = requests.post(Api.login_user, headers=crpo_login_header, data=json.dumps(data0), verify=True)
        self.TokenVal = response.json()
        self.NTokenVal = self.TokenVal.get("Token")

    def reActivate_candidate(self, candidateId, testId):
        # ----------------------------------------------------------------------------------------------------------
        #  Reactivate Test user login and login to test again
        # ----------------------------------------------------------------------------------------------------------
        eval_online_assessment_header = {"content-type": "application/json", "X-AUTH-TOKEN": self.NTokenVal}
        eval_online_assessment_data = {"candidateIds": [candidateId], "testId": testId}
        requests.post("https://amsin.hirepro.in/py/assessment/testuser/api/v1/reActivateLogin/",
                      headers=eval_online_assessment_header,
                      data=json.dumps(eval_online_assessment_data, default=str), verify=True)

    def qid_writetoxls(self, collection1, collection2, anscollection1, anscollection2):
        self.finalCollection_question = OrderedDict(zip(collection1, collection2))
        cellposition = 2
        for k, v in self.finalCollection_question.items():
            ans1 = anscollection1.get(k)
            ans2 = anscollection2.get(k)
            cellposition_question = cellposition + 4
            if k == v and ans1 == ans2:
                self.ws.write(self.rownum1, cellposition_question, k)
                self.ws.write(self.rownum1, cellposition_question + 1, v, self.__style3)

                self.ws.write(self.rownum1, cellposition_question + 2, ans1)
                self.ws.write(self.rownum1, cellposition_question + 3, ans2, self.__style3)
                self.file.write("""<td>""" + str(k) + """</td><td class="td-pass">""" + str(v) + """</td><td>""" + str(
                    ans1) + """</td><td class="td-pass">""" + str(ans2) + """</td>""")

            elif k == v and ans1 != ans2:
                self.ws.write(self.rownum1, cellposition_question, k)
                self.ws.write(self.rownum1, cellposition_question + 1, v, self.__style3)

                self.ws.write(self.rownum1, cellposition_question + 2, ans1)
                self.ws.write(self.rownum1, cellposition_question + 3, ans2, self.__style2)
                self.file.write("""<td>""" + str(k) + """</td><td class="td-pass">""" + str(v) + """</td><td>""" + str(
                    ans1) + """</td><td class="td-fail">""" + str(ans2) + """</td>""")

            elif k != v and ans1 == ans2:
                self.ws.write(self.rownum1, cellposition_question, k)
                self.ws.write(self.rownum1, cellposition_question + 1, v, self.__style2)

                self.ws.write(self.rownum1, cellposition_question + 2)
                self.ws.write(self.rownum1, cellposition_question + 3, ans2, self.__style3)
                self.file.write("""<td>""" + str(k) + """</td><td class="td-fail">""" + str(v) + """</td><td>""" + str(
                    ans1) + """</td><td class="td-pass">""" + str(ans2) + """</td>""")

            else:
                self.ws.write(self.rownum1, cellposition_question, k)
                self.ws.write(self.rownum1, cellposition_question + 1, v, self.__style2)

                self.ws.write(self.rownum1, cellposition_question + 2)
                self.ws.write(self.rownum1, cellposition_question + 3, ans2, self.__style2)
                self.file.write("""<td>""" + str(k) + """</td><td class="td-fail">""" + str(v) + """</td><td>""" + str(
                    ans1) + """</td><td class="td-fail">""" + str(ans2) + """</td>""")
            cellposition += 4

    def codinglang_writexls(self, langCollection1, langCollection2):
        code_cellposition = 46
        for code_id in self.coding_questions_id:
            lang1 = langCollection1.get(code_id)
            lang2 = langCollection2.get(code_id)
            if lang1 == lang2:
                self.ws.write(self.rownum1, code_cellposition, lang1)
                self.ws.write(self.rownum1, code_cellposition + 1, lang2, self.__style3)
                self.file.write("""<td>""" + str(lang1) + """</td><td class="td-pass">""" + str(lang2) + """</td>""")
            else:
                self.ws.write(self.rownum1, code_cellposition, lang1, self.__style2)
                self.ws.write(self.rownum1, code_cellposition + 1, lang2, self.__style2)
                self.file.write("""<td>""" + str(lang1) + """</td><td class="td-fail">""" + str(lang2) + """</td>""")
            code_cellposition += 2


obj = QuestionRandomization()
obj.test_QuestionRandomization()
