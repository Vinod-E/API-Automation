import requests
import json
import datetime
from hpro_automation.api import *
from hpro_automation import (login, input_paths)
from hpro_automation.Config import read_excel
from scripts.Overall_Status.overall_status_of_usecase import OverallStatus


class CandidateSlots(login.CommonLogin):

    def __init__(self):
        self.start_time = str(datetime.datetime.now())
        self.server = login_server
        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 9)))
        # --------------------------------- Inheritance Class Instance -------------------------------------------------
        super(CandidateSlots, self).__init__()
        self.overall = OverallStatus()

        # --------------------------------- Excel Data initialize variables --------------------------------------------
        self.xl_dict = {}
        self.dict_total = []

        self.choose_response = {}
        self.re_choose_response = {}
        self.candidate_status = {}
        self.re_candidate_status = {}
        self.applicant_status = {}
        self.re_applicant_status = {}

        self.choose_assign_data = ""
        self.re_choose_assign_data = ""
        self.context_id = ""
        self.re_context_id = ""
        self.applicant_id = ""
        self.re_applicant_id = ""
        self.applicant_message = ""
        self.re_applicant_message = ""
        self.applicant_info_status = ""
        self.re_applicant_info_status = ""
        self.choose_error_message = ""
        self.re_choose_error_message = ""

    def excel_data(self):

        # ------------------------------- Excel Data Read --------------------------------------------------------------
        try:
            excel = read_excel.ExcelRead()
            if login_server == 'amsin':
                index = 0
            else:
                index = 1
            excel.excel_read(input_paths.inputpaths['candidate_slot_input_sheet'], index)
            self.xl_dict = excel.details
            self.dict_total = excel.details

            print("Excel Data:: ", self.xl_dict)
        except IOError:
            print("File not found or path is incorrect")

    def choose_candidate_slot(self, loop):
        try:
            self.slot_captcha_login_token('verify')
            self.preference(int(self.xl_dict[loop]['Applicant_id']), self.xl_dict[loop]['verify_hash'])
            self.lambda_function('save_slot')

            # ----------------------------------- API request -----------------------------------------------------
            print("------------- Choose Candidate slot API Call -----------------")
            request = json.loads(self.xl_dict[loop]['chooseSlot'])
            choose_slot_api = requests.post(self.webapi, headers=self.headers,
                                            data=json.dumps(request), verify=False)
            self.choose_response = json.loads(choose_slot_api.content)
            print(self.choose_response)

            if self.choose_response.get("data"):
                self.choose_assign_data = self.choose_response.get("data")
                for key in self.choose_assign_data.keys():
                    self.context_id = key
            elif self.choose_response.get('error'):
                self.choose_error_message = self.choose_response.get('error').get('errorDescription')

        except KeyError as e:
            print(e)

    def re_choose_candidate_slot(self, loop):
        try:
            self.slot_captcha_login_token('verify')
            self.preference(int(self.xl_dict[loop]['Applicant_id']), self.xl_dict[loop]['verify_hash'])
            self.lambda_function('save_slot')

            # ----------------------------------- API request -----------------------------------------------------
            print("------------- Choose Candidate slot API Call -----------------")
            request = json.loads(self.xl_dict[loop]['updateSlot'])
            choose_slot_api = requests.post(self.webapi, headers=self.headers,
                                            data=json.dumps(request), verify=False)
            self.re_choose_response = json.loads(choose_slot_api.content)
            print(self.re_choose_response)

            if self.re_choose_response.get("data"):
                self.re_choose_assign_data = self.re_choose_response.get("data")
                for key in self.re_choose_assign_data.keys():
                    self.re_context_id = key
            elif self.re_choose_response.get('error'):
                self.re_choose_error_message = self.re_choose_response.get('error').get('errorDescription')

        except KeyError as e:
            print(e)

    def candidate_slot_status(self, loop):
        try:
            self.common_login('slot')
            self.Non_lambda_headers['Authorization'] = ""
            self.lambda_function('applicant_screen_data')

            # ----------------------------------- API request ---------------------------------------------------------
            print("------------- Choose Unassign slot API Call -----------------")
            request = json.loads(self.xl_dict[loop]['screen_data'])

            status_api = requests.post(self.webapi, headers=self.headers,
                                       data=json.dumps(request), verify=False)
            self.candidate_status = json.loads(status_api.content)
            print(self.candidate_status)

            if self.candidate_status.get('data'):
                self.applicant_info_status = self.candidate_status.get('data').get('getApplicantsInfo').get('Status')

        except KeyError as e:
            print(e)

    def re_candidate_slot_status(self, loop):
        try:
            self.common_login('slot')
            self.Non_lambda_headers['Authorization'] = ""
            self.lambda_function('applicant_screen_data')

            # ----------------------------------- API request ---------------------------------------------------------
            print("------------- Choose Unassign slot API Call -----------------")
            request = json.loads(self.xl_dict[loop]['screen_data'])

            status_api = requests.post(self.webapi, headers=self.headers,
                                       data=json.dumps(request), verify=False)
            self.re_candidate_status = json.loads(status_api.content)
            print(self.re_candidate_status)

            if self.re_candidate_status.get('data'):
                self.re_applicant_info_status = self.re_candidate_status.get('data').get('getApplicantsInfo').get(
                    'Status')

        except KeyError as e:
            print(e)

    def change_applicant_status(self, loop):
        try:
            self.lambda_function('ChangeApplicant_Status')
            self.Non_lambda_headers['Authorization'] = ""

            # ----------------------------------- API request ---------------------------------------------------------
            print("------------- Change Applicant status API Call -----------------")
            request = json.loads(self.xl_dict[loop]['change_status'])

            status_change_api = requests.post(self.webapi, headers=self.headers,
                                              data=json.dumps(request), verify=False)
            self.applicant_status = json.loads(status_change_api.content)
            print(self.applicant_status)

            if self.applicant_status.get('data'):
                if self.applicant_status.get('data').get('success'):
                    success = self.applicant_status.get('data').get('success')
                    for key in success:
                        self.applicant_id = key
                        message = self.applicant_status.get('data').get('success').get(key)
                        for i in message:
                            self.applicant_message = i
                elif self.applicant_status.get('data').get('failure'):
                    failure = self.applicant_status.get('data').get('failure')
                    for key in failure:
                        self.applicant_id = key
                        message = self.applicant_status.get('data').get('failure').get(key)
                        for i in message:
                            self.applicant_message = i

        except KeyError as e:
            print(e)

    def re_change_applicant_status(self, loop):
        try:
            self.lambda_function('ChangeApplicant_Status')
            self.Non_lambda_headers['Authorization'] = ""

            # ----------------------------------- API request ---------------------------------------------------------
            print("------------- Change Applicant status API Call -----------------")
            request = json.loads(self.xl_dict[loop]['change_status'])

            unassign_slot_api = requests.post(self.webapi, headers=self.headers,
                                              data=json.dumps(request), verify=False)
            self.re_applicant_status = json.loads(unassign_slot_api.content)
            print(self.re_applicant_status)

            if self.re_applicant_status.get('data'):
                if self.re_applicant_status.get('data').get('success'):
                    success = self.re_applicant_status.get('data').get('success')
                    for key in success:
                        self.re_applicant_id = key
                        message = self.re_applicant_status.get('data').get('success').get(key)
                        for i in message:
                            self.re_applicant_message = i
                elif self.re_applicant_status.get('data').get('failure'):
                    failure = self.re_applicant_status.get('data').get('failure')
                    for key in failure:
                        self.re_applicant_id = key
                        message = self.re_applicant_status.get('data').get('failure').get(key)
                        for i in message:
                            self.re_applicant_message = i

        except KeyError as e:
            print(e)

    def output_report_excel(self, loop):
        self.overall.output_excel_input_output_header('Candidate_slot_output_sheet')

        self.overall.write_in_excel(2, self.xl_dict[loop]['Applicant_id'],
                                    int(self.applicant_id), None, 'Null')
        self.overall.write_in_excel(3, self.xl_dict[loop]['assigned_slot'],
                                    self.context_id, None, 'Null')
        self.overall.write_in_excel(4, self.xl_dict[loop]['choose_status'],
                                    self.applicant_info_status, None, 'No Message')
        self.overall.write_in_excel(5, self.xl_dict[loop]['applicant_status_message'],
                                    self.applicant_message, None, 'No Message')
        self.overall.write_in_excel(6, self.xl_dict[loop]['re_assigned_error_message'],
                                    self.re_choose_error_message, None, 'No Message')
        self.overall.write_in_excel(7, self.xl_dict[loop]['re_select_slot'],
                                    self.re_context_id, None, 'Null')
        self.overall.write_in_excel(8, self.xl_dict[loop]['re_choose_status'],
                                    self.re_applicant_info_status, None, 'No Message')
        self.overall.write_in_excel(9, self.xl_dict[loop]['re_applicant_status_message'],
                                    self.re_applicant_message, None, 'No Message')
        self.overall.write_in_excel(10, self.xl_dict[loop]['assigned_error_message'],
                                    self.choose_error_message, None, 'No Message')


Object = CandidateSlots()
Object.excel_data()

Total_count = len(Object.dict_total)
print("Number of Rows::", Total_count)
for looping in range(0, Total_count):
    print("Iteration Count is ::", looping)
    Object.choose_candidate_slot(looping)
    Object.candidate_slot_status(looping)
    Object.change_applicant_status(looping)

    Object.re_choose_candidate_slot(looping)
    Object.re_candidate_slot_status(looping)
    Object.re_change_applicant_status(looping)

    Object.output_report_excel(looping)
    Object.overall.test_case_wise_pass_or_fail()

    # ------ Remove Dictionaries
    Object.choose_response = {}
    Object.re_choose_response = {}
    Object.candidate_status = {}
    Object.re_candidate_status = {}
    Object.applicant_status = {}
    Object.re_applicant_status = {}

    Object.choose_assign_data = ""
    Object.re_choose_assign_data = ""
    Object.context_id = ""
    Object.re_context_id = ""
    Object.applicant_id = ""
    Object.re_applicant_id = ""
    Object.applicant_message = ""
    Object.re_applicant_message = ""
    Object.applicant_info_status = ""
    Object.re_applicant_info_status = ""
    Object.choose_error_message = ""
    Object.re_choose_error_message = ""

# ----------------- Overall Output Status ------------------------------------------------------
Object.overall.main_headers = ['Comparison', 'Actual_status', 'Applicant ID', 'Choose Slot Message', 'Applicant Status',
                               'Change Status Message', 'Slot Error Message', 'ReSelect Slot Message',
                               'Applicant Status', 'Change Status Message', 'ReSlot Error Message']
Object.overall.headers_with_style2 = ['Comparison', 'Actual_status']
Object.overall.file_headers_col_row()
Object.overall.overall_status('CANDIDATE SLOTS', Object.Expected_success_cases,
                              Object.start_time, Object.calling_lambda, 'verify',
                              Object.server, Total_count, 'Candidate_slot_output_sheet')
