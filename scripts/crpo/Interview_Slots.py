import requests
import json
import datetime
from hpro_automation.api import *
from hpro_automation import (login, input_paths)
from hpro_automation.Config import read_excel
from scripts.Overall_Status.overall_status_of_usecase import OverallStatus


class InterviewSlots(login.CommonLogin):

    def __init__(self):
        self.start_time = str(datetime.datetime.now())
        self.server = login_server
        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 11)))
        # --------------------------------- Inheritance Class Instance -------------------------------------------------
        super(InterviewSlots, self).__init__()
        self.overall = OverallStatus()

        # --------------------------------- Excel Data initialize variables --------------------------------------------
        self.xl_dict = {}
        self.dict_total = []
        self.choose_response = {}
        self.choose_slot_data = {}
        self.unassign_response = {}
        self.unassign_slot_data = {}

        self.choose_assign_message = ""
        self.choose_failed_message = ""
        self.applicant_id = ""
        self.choose_error = ""
        self.unassign_message = ""
        self.unassign_error_message = ""
        self.unassign_deep_error = ""

    def excel_data(self):

        # ------------------------------- Excel Data Read --------------------------------------------------------------
        try:
            excel = read_excel.ExcelRead()
            if login_server == 'amsin':
                index = 0
            else:
                index = 1
            excel.excel_read(input_paths.inputpaths['interview_slot_input_sheet'], index)
            self.xl_dict = excel.details
            self.dict_total = excel.details

            print("Excel Data:: ", self.xl_dict)
        except IOError:
            print("File not found or path is incorrect")

    def choose_interview_slot(self, loop):
        try:
            self.slot_captcha_login_token('interview')
            self.authenticate(self.xl_dict[loop]['authenticate_request'])
            self.lambda_function('interview_slot_select')

            # ----------------------------------- API request -----------------------------------------------------
            print("------------- Choose Interview slot API Call -----------------")
            request = json.loads(self.xl_dict[loop]['chooseSlot'])
            choose_slot_api = requests.post(self.webapi, headers=self.headers,
                                            data=json.dumps(request), verify=False)
            self.choose_response = json.loads(choose_slot_api.content)
            print(self.choose_response)

            if self.choose_response.get("data"):
                self.choose_assign_message = self.choose_response.get("data").get('success')
                self.choose_failed_message = self.choose_response.get("data").get('failed')
                for key in self.choose_assign_message.keys():
                    self.applicant_id = key
                    self.choose_assign_message = self.choose_response.get("data").get('success').get(key)
                    print(self.applicant_id)
                    break
                for key in self.choose_failed_message.keys():
                    self.applicant_id = key
                    self.choose_failed_message = self.choose_response.get("data").get('failed').get(key)
                    print(self.applicant_id)
                    break
                print(self.choose_assign_message, self.choose_failed_message)
            elif self.choose_response.get('error'):
                self.choose_error = self.choose_response.get('error').get('errorDescription')
                print(self.choose_error)

        except KeyError as e:
            print(e)

    def unassign_slot(self, loop):
        try:
            self.common_login('slot')
            self.lambda_function('interview_unassign_slot')
            self.Non_lambda_headers['Authorization'] = ""

            # ----------------------------------- API request ---------------------------------------------------------
            print("------------- Choose Unassign slot API Call -----------------")
            request = json.loads(self.xl_dict[loop]['UnassignSlot'])

            unassign_slot_api = requests.post(self.webapi, headers=self.headers,
                                              data=json.dumps(request), verify=False)
            self.unassign_response = json.loads(unassign_slot_api.content)
            print(self.unassign_response)

            if self.unassign_response.get('data'):
                self.unassign_message = self.unassign_response.get('data').get('message')
                print(self.unassign_message)
            elif self.unassign_response.get('error'):
                self.unassign_error_message = self.unassign_response.get('error').get('message')
                print(self.unassign_error_message)
                if self.unassign_response.get('error').get('error'):
                    self.unassign_deep_error = self.unassign_response.get('error').get('error').get('errorDescription')
                    print(self.unassign_deep_error)

        except KeyError as e:
            print(e)

    def output_excel_status_headers(self, loop):
        self.overall.output_excel('Interview_slot_output_sheet')

        self.overall.write_in_excel(2, int(self.xl_dict[loop]['Applicant_id']),
                                    int(self.applicant_id), None, 'Null')
        self.overall.write_in_excel(3, self.xl_dict[loop]['choose_Slot_Message'],
                                    self.choose_failed_message, self.choose_assign_message, 'No Message')
        self.overall.write_in_excel(4, self.xl_dict[loop]['choose_error_message'],
                                    self.choose_error, None, 'No Message')
        self.overall.write_in_excel(5, self.xl_dict[loop]['unassign_Slot_Message'],
                                    self.unassign_message, None, 'No Message')
        self.overall.write_in_excel(6, self.xl_dict[loop]['Unassign_error_message'],
                                    self.unassign_error_message, self.unassign_deep_error, 'No Message')


Object = InterviewSlots()
Object.excel_data()

Total_count = len(Object.dict_total)
print("Number of Rows::", Total_count)
for looping in range(0, Total_count):
    print("Iteration Count is ::", looping)
    Object.choose_interview_slot(looping)
    Object.unassign_slot(looping)
    Object.output_excel_status_headers(looping)

    # ------ Remove Dictionaries
    Object.choose_response = {}
    Object.choose_slot_data = {}
    Object.update_response = {}
    Object.update_slot_data = {}
    Object.unassign_response = {}
    Object.unassign_slot_data = {}

    Object.choose_assign_message = ""
    Object.choose_failed_message = ""
    Object.choose_error = ""
    Object.unassign_message = ""
    Object.unassign_error_message = ""
    Object.unassign_deep_error = ""

# ----------------- Overall Output Status ------------------------------------------------------
Object.overall.main_headers = ['Comparison', 'Actual_status', 'Applicant ID',
                               'Choose Slot Message', 'Choose Slot Error Messages',
                               'UnAssign Slot Message', 'Unassign Slot Error Messages']
Object.overall.headers_with_style2 = ['Comparison', 'Actual_status']
Object.overall.file_headers_col_row()
Object.overall.overall_status('INTERVIEW SLOTS', Object.Expected_success_cases,
                              Object.start_time, Object.calling_lambda, 'interviewSlots',
                              Object.server, Total_count, 'Interview_slot_output_sheet')
