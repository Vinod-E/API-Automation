import requests
import json
import datetime
from hpro_automation.api import *
from hpro_automation import (login, input_paths)
from hpro_automation.Config import read_excel
from scripts.Overall_Status.overall_status_of_usecase import OverallStatus


class AssessmentSlots(login.CommonLogin):

    def __init__(self):
        self.start_time = str(datetime.datetime.now())
        self.server = login_server
        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 17)))
        # --------------------------------- Inheritance Class Instance -------------------------------------------------
        super(AssessmentSlots, self).__init__()
        self.overall = OverallStatus()

        # --------------------------------- Excel Data initialize variables --------------------------------------------
        self.xl_dict = {}
        self.dict_total = []
        self.choose_response = {}
        self.update_response = {}
        self.unassign_response = {}

        self.choose_data = {}
        self.choose_error = {}
        self.update_data = {}
        self.update_error = {}
        self.dissociate_data = {}
        self.unassign_error = {}

    def excel_data(self):

        # ------------------------------- Excel Data Read --------------------------------------------------------------
        try:
            excel = read_excel.ExcelRead()
            if login_server == 'amsin':
                index = 0
            else:
                index = 1
            excel.excel_read(input_paths.inputpaths['assessment_slot_input_sheet'], index)
            self.xl_dict = excel.details
            self.dict_total = excel.details

            print("Excel Data:: ", self.xl_dict)
        except IOError:
            print("File not found or path is incorrect")

    def choose_assessment_slot(self, loop):
        try:
            self.slot_captcha_login_token('assessment')
            self.verify_hash(self.xl_dict[loop]['verify_hash'])
            self.lambda_function('assessment_slot_select')

            # ----------------------------------- API request -----------------------------------------------------
            print("------------- Choose assessment slot API Call -----------------")
            request = json.loads(self.xl_dict[loop]['chooseSlot'])
            choose_slot_api = requests.post(self.webapi, headers=self.lambda_headers,
                                            data=json.dumps(request), verify=False)
            self.choose_response = json.loads(choose_slot_api.content)
            print(self.choose_data)

            if self.choose_response.get("data"):
                self.choose_data = self.choose_response.get("data")
                print(self.choose_data)
            elif self.choose_response.get("error"):
                self.choose_error = self.choose_response.get("error")
                print(self.choose_error)

        except KeyError as e:
            print(e)

    def update_assessment_slot(self, loop):
        try:
            self.lambda_function('assessment_slot_update')

            # ----------------------------------- API request ------------------------------------------------------
            print("------------- Update assessment slot API Call -----------------")
            request = json.loads(self.xl_dict[loop]['updateSlot'])
            update_slot_api = requests.post(self.webapi, headers=self.lambda_headers,
                                            data=json.dumps(request), verify=False)
            self.update_response = json.loads(update_slot_api.content)
            print(self.update_response)

            if self.update_response.get("data"):
                self.update_data = self.update_response.get("data")
                print(self.update_data)
            elif self.update_response.get("error"):
                self.update_error = self.update_response.get("error")
                print(self.update_error)

        except KeyError as e:
            print(e)

    def unassign_slot(self, loop):
        try:
            self.common_login('slot')
            self.lambda_function('assessment_unassign_slot')

            # ----------------------------------- API request ---------------------------------------------------------
            print("------------- Choose Unassign slot API Call -----------------")
            request = json.loads(self.xl_dict[loop]['UnassignSlot'])

            unassign_slot_api = requests.post(self.webapi, headers=self.headers,
                                              data=json.dumps(request), verify=False)
            self.unassign_response = json.loads(unassign_slot_api.content)
            print(self.unassign_response)

            if self.unassign_response.get('data'):
                unassign_data = self.unassign_response.get('data').get('dissociateSlotDetails')
                for data in unassign_data:
                    self.dissociate_data = data
                    print(self.dissociate_data)
            elif self.unassign_response.get('error'):
                self.unassign_error = self.unassign_response.get('error')

        except KeyError as e:
            print(e)

    def output_excel_status_headers(self, loop):
        self.overall.output_excel('Assessment_slot_output_sheet')

        self.overall.validation(2, self.xl_dict[loop]['Applicant_id'],
                                self.choose_data.get('applicantId') if self.choose_data else 'Null')
        self.overall.validation(3, self.xl_dict[loop]['assigned_slot'],
                                str(self.choose_data.get('isAssigned')) if self.choose_data else 'Null')
        self.overall.validation(4, self.xl_dict[loop]['updated_slot'],
                                str(self.update_data.get('isUpdated')) if self.update_data else 'Null')
        self.overall.validation(5, self.xl_dict[loop]['dissociated'],
                                str(self.dissociate_data.get('isDissociated')) if self.dissociate_data else 'Null')
        self.overall.validation(6, self.xl_dict[loop]['assigned_message'],
                                self.choose_error.get('errorDescription') if self.choose_error else 'No Message')

        if self.update_data:
            self.overall.validation(7, self.xl_dict[loop]['update_message'],
                                    self.update_data.get('error') if self.update_data.get('error') else 'No Message')
        elif self.update_error:
            self.overall.validation(7, self.xl_dict[loop]['update_message'],
                                    self.update_error.get('errorDescription') if self.update_error else 'elif No Message')

        if self.dissociate_data:
            self.overall.validation(8, self.xl_dict[loop]['dissociated_message'],
                                    self.dissociate_data.get('error') if
                                    self.dissociate_data.get('error') else 'if No Message')
        elif self.unassign_error:
            self.overall.validation(8, self.xl_dict[loop]['dissociated_message'],
                                    self.unassign_error.get('errorDescription') if
                                    self.unassign_error else ' elif No Message')


Object = AssessmentSlots()
Object.excel_data()

Total_count = len(Object.dict_total)
print("Number of Rows::", Total_count)
for looping in range(0, Total_count):
    print("Iteration Count is ::", looping)
    Object.choose_assessment_slot(looping)
    Object.update_assessment_slot(looping)
    Object.unassign_slot(looping)
    Object.output_excel_status_headers(looping)

    Object.choose_data = {}
    Object.choose_error = {}
    Object.update_data = {}
    Object.update_error = {}
    Object.dissociate_data = {}
    Object.unassign_error = {}

# ----------------- Overall Output Status ------------------------------------------------------
Object.overall.main_headers = ['Comparison', 'Actual_status', 'Applicant ID', 'Choose Slot', 'Update Slot',
                               'UnAssign Slot', 'Choose Slot Message', 'Update Slot Message',
                               'UnAssign Slot Message']
Object.overall.headers_with_style2 = ['Comparison', 'Actual_status']
Object.overall.file_headers_col_row()
Object.overall.overall_status('ASSESSMENT SLOTS', Object.Expected_success_cases,
                              Object.start_time, Object.calling_lambda, 'assessmentSlots',
                              Object.server, Total_count, 'Assessment_slot_output_sheet')
