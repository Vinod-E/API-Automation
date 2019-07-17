from hpro_automation import (login, work_book, input_paths, output_paths)
import datetime
import requests
import json
import xlrd


class ManageInterviewers(login.CommonLogin, work_book.WorkBook):

    def __init__(self):

        # ---------------------------- Overall Status Current Run Date -------------------------------------------------
        self.start_time = str(datetime.datetime.now())

        # --------------------------------- Inheritance Class Instance -------------------------------------------------
        super(ManageInterviewers, self).__init__()
        self.common_login('v')

        # --------------------------------- Overall status initialize variables ----------------------------------------
        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 66)))
        self.Actual_Success_case = []

        # --------------------------------- Excel Data initialize variables --------------------------------------------
        self.xl_event_id = []
        self.xl_compositeKey = []
        self.xl_interviewer_id = []
        self.xl_email_id = []
        self.xl_interviewer_decision = []
        self.xl_manager_decision = []

        # --------------------------------- Dictionary initialize variables --------------------------------------------
        self.success_case_01 = {}
        self.success_case_02 = {}
        self.success_case_03 = {}
        self.headers = {}

    def excel_headers(self):

        # --------------------------------- Excel Headers and Cell color, styles ---------------------------------------
        self.main_headers = ['S.No', 'Name', 'Designation']
        self.headers_with_style2 = ['S.No']
        self.headers_with_style9 = ['Name']
        self.headers_with_style19 = ['Designation']
        self.file_headers_col_row()

    def excel_data(self):

        # ------------------------------- Excel Data Read --------------------------------------------------------------
        try:
            workbook = xlrd.open_workbook(input_paths.inputpaths['Manage_Int_Input_sheet'])
            sheet1 = workbook.sheet_by_index(0)
            for i in range(1, sheet1.nrows):
                number = i  # Counting number of rows
                rows = sheet1.row_values(number)

                if not rows[0]:
                    self.xl_event_id.append(None)
                else:
                    self.xl_event_id.append(int(rows[0]))
                if not rows[1]:
                    self.xl_compositeKey.append(None)
                else:
                    self.xl_compositeKey.append(int(rows[1]))
                if not rows[2]:
                    self.xl_interviewer_id.append(None)
                else:
                    self.xl_interviewer_id.append(int(rows[2]))
                if not rows[3]:
                    self.xl_email_id.append(None)
                else:
                    self.xl_email_id.append(str(rows[3]))
                if not rows[4]:
                    self.xl_interviewer_decision.append(None)
                else:
                    self.xl_interviewer_decision.append(int(rows[4]))
                if not rows[5]:
                    self.xl_manager_decision.append(None)
                else:
                    self.xl_manager_decision.append(int(rows[5]))

        except IOError:
            print("File not found or path is incorrect")

    def send_nomination_mails(self, loop):

        self.lambda_function('send_nomination_mails_to_selected_interviewers')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"eventId": self.xl_event_id[loop],
                   "data": {
                       self.xl_compositeKey[loop]: [{
                           "interviewerId": self.xl_interviewer_id[loop],
                           "emailId": self.xl_email_id[loop]
                       }]
                   }}

        send_nomination_api = requests.post(self.webapi, headers=self.headers, data=json.dumps(request, default=str),
                                            verify=False)
        print(send_nomination_api.headers)
        send_nomination_api_response = json.loads(send_nomination_api.content)
        print(send_nomination_api_response)

    def output_report(self, loop):

        # --------------------------------- Writing Input Data ---------------------------------------------------------
        self.ws.write(self.rowsize, self.col, 'Input', self.style4)
        self.ws.write(self.rowsize, 1, self.xl_event_id[loop] if self.xl_event_id[loop] else 'Empty')

        # --------------------------------- Writing Output Data --------------------------------------------------------
        self.rowsize += 1
        self.ws.write(self.rowsize, self.col, 'Output', self.style5)
        self.rowsize += 1

        # ------------------------------------ OutPut File save --------------------------------------------------------
        Object.wb_Result.save(output_paths.outputpaths['MI_output_sheet'])

        if self.success_case_01 == 'Pass':
            self.Actual_Success_case.append(self.success_case_01)
        if self.success_case_02 == 'Pass':
            self.Actual_Success_case.append(self.success_case_02)
        if self.success_case_03 == 'Pass':
            self.Actual_Success_case.append(self.success_case_03)

    def overall_status(self):
        self.ws.write(0, 0, 'Manage Interviewers', self.style23)
        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
        else:
            self.ws.write(0, 1, 'Fail', self.style25)

        self.ws.write(0, 3, 'Start Time', self.style23)
        self.ws.write(0, 4, self.start_time, self.style26)

        # ---------------------------- OutPut File save with Overall Status --------------------------------------------
        Object.wb_Result.save(output_paths.outputpaths['MI_output_sheet'])


Object = ManageInterviewers()
Object.excel_headers()
Object.excel_data()
Total_count = len(Object.xl_event_id)
print("Number Of Rows ::", Total_count)
if Object.login == 'OK':
    for looping in range(0, Total_count):
        print("Iteration Count is ::", looping)
        Object.send_nomination_mails(looping)
        Object.output_report(looping)

        # ----------------- Make Dictionaries clear for each loop ------------------------------------------------------
        Object.success_case_01 = {}
        Object.success_case_02 = {}
        Object.success_case_03 = {}
        Object.headers = {}

# ---------------------------- Call this function at last --------------------------------------------------------------
Object.overall_status()
