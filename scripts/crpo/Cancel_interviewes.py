from hpro_automation import (login, output_paths, input_paths, work_book, db_login)
from hpro_automation.api import *
import json
import requests
import xlrd
import datetime


class CancelInterview(login.CommonLogin, work_book.WorkBook, db_login.DBConnection):

    def __init__(self):
        self.start_time = str(datetime.datetime.now())
        super(CancelInterview, self).__init__()
        self.common_login('admin')
        self.crpo_app_name = self.app_name.strip()
        self.db_connection()
        print(self.crpo_app_name)

        # -----------------------
        # Initialising Excel Data
        # -----------------------
        self.xl_ir_id = []  # [] Initialising data from Excel sheet to the variables
        self.xl_cancel_statusID = 167112
        self.xl_expected_message = []
        self.xl_interviewer_comment = 'Interview cancelled by Admin'

        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 31)))
        self.Actual_Success_case = []

        # ---------------------------------
        # Dictionaries for Interview_cancel
        # ---------------------------------
        self.api_ir = {}
        self.Exception_message_api = {}

        self.success_case_01 = {}
        self.success_case_02 = {}
        self.success_case_03 = {}

    def excel_headers(self):
        self.main_headers = ['Comparison', 'Status', 'Interview_Request_ID', 'Message']
        self.headers_with_style2 = ['Comparison', 'Status']
        self.file_headers_col_row()

    def excel_data_ir1(self):
        try:
            workbook = xlrd.open_workbook(output_paths.outputpaths['Interview_flow_Output_sheet'])
            sheet = workbook.sheet_by_index(0)
            for i in range(2, sheet.nrows):
                number = i
                rows = sheet.row_values(number)

                if rows[0] is not None and rows[3] != '':
                    self.xl_ir_id.append(int(rows[3]))

        except IOError:
            print("File not found or path is incorrect")

    def excel_data_ir2(self):
        try:
            workbook = xlrd.open_workbook(output_paths.outputpaths['Reschdeule_Output_sheet'])
            sheet = workbook.sheet_by_index(0)
            for i in range(2, sheet.nrows):
                number = i
                rows = sheet.row_values(number)

                if rows[0] is not None and rows[2] != '':
                    self.xl_ir_id.append(int(rows[2]))

        except IOError:
            print("File not found or path is incorrect")

        try:
            workbook = xlrd.open_workbook(input_paths.inputpaths['Cancel_interview_input_sheet'])
            sheet = workbook.sheet_by_index(0)
            for i in range(1, sheet.nrows):
                number = i
                rows = sheet.row_values(number)

                if rows[0] is not None and rows[0] != '':
                    self.xl_expected_message.append(rows[0])

        except IOError:
            print("File not found or path is incorrect")

    def cancel_interview_request(self, loop):

        self.lambda_function('cancel')
        self.headers['APP-NAME'] = self.crpo_app_name

        cancel_request = {"interviewRequestIds": [self.xl_ir_id[loop]],
                          "interviewCanceledStatusId": self.xl_cancel_statusID,
                          "applicantStatusItemComment": self.xl_interviewer_comment}

        cancel_request_api = requests.post(self.webapi, headers=self.headers,
                                           data=json.dumps(cancel_request, default=str), verify=False)
        print(cancel_request_api.headers)
        cancel_request_api_response = json.loads(cancel_request_api.content)
        data = cancel_request_api_response['data']

        success = data['success']
        for ir in success:
            self.api_ir = ir

        failure = data['failure']
        for k in failure.keys():
            fail = failure[k]
            for j in fail:
                self.Exception_message_api = j

    def output_excel(self, loop):
        # ------------------
        # Writing Input Data
        # ------------------
        self.ws.write(self.rowsize, self.col, 'Input', self.style4)
        self.ws.write(self.rowsize, 2, self.xl_ir_id[loop])
        self.ws.write(self.rowsize, 3, self.xl_expected_message[loop].format(self.xl_ir_id[loop]))

        # -------------------
        # Writing Output Data
        # -------------------
        self.rowsize += 1  # Row increment
        self.ws.write(self.rowsize, self.col, 'Output', self.style5)

        # --------------------------------------------------------------------------------------------------------------
        if self.api_ir:
            if self.Exception_message_api:
                if self.Exception_message_api and 'has  scheduled or rescheduled  status' \
                        in self.xl_expected_message[loop]:
                    self.ws.write(self.rowsize, 1, 'Pass', self.style26)
                    self.success_case_01 = 'Pass'
                else:
                    self.ws.write(self.rowsize, 1, 'Fail', self.style3)
            else:
                self.ws.write(self.rowsize, 1, 'Pass', self.style26)
                self.success_case_02 = 'Pass'
        elif self.Exception_message_api and 'Unable to cancel interview request id' in self.xl_expected_message[loop]:
            self.ws.write(self.rowsize, 1, 'Pass', self.style26)
            self.success_case_03 = 'Pass'
        else:
            self.ws.write(self.rowsize, 1, 'Fail', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.Exception_message_api:
            if self.Exception_message_api and 'has  scheduled or rescheduled  status' in self.xl_expected_message[loop]:
                self.ws.write(self.rowsize, 3, self.Exception_message_api, self.style8)
            elif self.Exception_message_api and 'Unable to cancel interview request id' \
                    in self.xl_expected_message[loop]:
                self.ws.write(self.rowsize, 3, self.Exception_message_api, self.style8)
            else:
                self.ws.write(self.rowsize, 3, self.Exception_message_api, self.style3)
        elif 'Cancelled Interview' and self.xl_expected_message[loop] in self.xl_expected_message[loop]:
            self.ws.write(self.rowsize, 3, self.xl_expected_message[loop].format(self.api_ir), self.style8)
        else:
            self.ws.write(self.rowsize, 3, self.Exception_message_api, self.style7)
        # --------------------------------------------------------------------------------------------------------------

        if self.api_ir:
            self.ws.write(self.rowsize, 2, self.api_ir, self.style8)
        else:
            self.ws.write(self.rowsize, 2, self.xl_ir_id[loop], self.style8)
        # --------------------------------------------------------------------------------------------------------------

        self.rowsize += 1  # Row increment
        Object.wb_Result.save(output_paths.outputpaths['Cancel_Interview_Output_sheet'])

        if self.success_case_01 == 'Pass':
            self.Actual_Success_case.append(self.success_case_01)
        if self.success_case_02 == 'Pass':
            self.Actual_Success_case.append(self.success_case_02)
        if self.success_case_03 == 'Pass':
            self.Actual_Success_case.append(self.success_case_03)

    def overall_status(self):
        self.ws.write(0, 0, 'Cancel Interview', self.style23)
        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
        else:
            self.ws.write(0, 1, 'Fail', self.style25)

        self.ws.write(0, 2, 'Login Server', self.style23)
        self.ws.write(0, 3, login_server, self.style24)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)
        self.ws.write(0, 6, 'APP Name', self.style23)
        self.ws.write(0, 7, self.crpo_app_name, self.style24)
        self.ws.write(0, 8, 'No.of Test cases', self.style23)
        self.ws.write(0, 9, Total_count, self.style24)
        self.ws.write(0, 10, 'Start Time', self.style23)
        self.ws.write(0, 11, self.start_time, self.style26)
        Object.wb_Result.save(output_paths.outputpaths['Cancel_Interview_Output_sheet'])

    def fetch_null_ir_without_candidate_id(self):
        query = "select ir.id from interview_requests ir " \
                "left join interview_candidates ic on ic.interviewrequest_id=ir.id " \
                "where ir.tenant_id=1787 and ic.interviewrequest_id is null;"
        print(query)
        self.cursor.execute(query)
        c_ir_ids = self.cursor.fetchone()[0]
        print("NULL CANDIDATE IR:: ", c_ir_ids)

        query = "DELETE FROM interview_interviwers WHERE interviewrequest_id = {};".format(c_ir_ids)
        print(query)
        self.cursor.execute(query)

        query = "DELETE FROM interview_requests WHERE id = {};".format(c_ir_ids)
        print(query)
        self.cursor.execute(query)


Object = CancelInterview()
Object.excel_headers()
Object.excel_data_ir1()
Object.excel_data_ir2()
Total_count = len(Object.xl_ir_id)
print("Number of Rows::", Total_count)

try:
    if Object.login == 'OK':
        for looping in range(0, Total_count):
            print("Iteration Count is ::", looping)
            Object.cancel_interview_request(looping)
            Object.output_excel(looping)

            # -------------------------------------
            # Making all dict empty for every loop
            # -------------------------------------
            Object.api_ir = {}
            Object.Exception_message_api = {}
            Object.success_case_01 = {}
            Object.success_case_02 = {}
            Object.success_case_03 = {}

    Object.overall_status()
    Object.fetch_null_ir_without_candidate_id()
    Object.connection.commit()
    Object.connection.close()

except AttributeError as Object_error:
    print(Object_error)
