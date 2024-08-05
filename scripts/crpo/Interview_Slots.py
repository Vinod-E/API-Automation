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
        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 17)))
        # --------------------------------- Inheritance Class Instance -------------------------------------------------
        super(InterviewSlots, self).__init__()
        self.overall = OverallStatus()

        # --------------------------------- Excel Data initialize variables --------------------------------------------
        self.xl_dict = {}
        self.dict_total = []
        self.choose_response = {}
        self.choose_slot_data = {}
        self.update_response = {}
        self.update_slot_data = {}
        self.unassign_response = {}
        self.unassign_slot_data = {}

        self.choose_data_applicant_id = ""
        self.choose_assign_message = ""
        self.choose_error = ""
        self.choose_slot_assigned = ""
        self.update_data = ""
        self.update_error = ""
        self.update_error_des = ""
        self.dissociate_data = ""
        self.unassign_error = ""
        self.dissociate_data_error = ""

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
            choose_slot_api = requests.post(self.webapi, headers=self.Non_lambda_headers,
                                            data=json.dumps(request), verify=False)
            self.choose_response = json.loads(choose_slot_api.content)
            print(self.choose_response)

            if self.choose_response.get("data"):
                self.choose_assign_message = (self.choose_slot_data.get('success')
                                              .get(self.xl_dict[loop]['Applicant_id']))
                self.choose_data_applicant_id = (self.choose_slot_data.get('success').keys())
                print(self.choose_assign_message, self.choose_data_applicant_id)

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
                self.unassign_slot_data = self.unassign_response.get('data').get('dissociateSlotDetails')
                for data in self.unassign_slot_data:
                    self.dissociate_data = str(data.get('isDissociated'))
                    self.dissociate_data_error = data.get('error')
                    print(self.dissociate_data, self.dissociate_data_error)
            elif self.unassign_response.get('error'):
                error = self.unassign_response.get('error')
                self.unassign_error = error.get('errorDescription')
                print(self.unassign_error)

        except KeyError as e:
            print(e)

    def output_excel_status_headers(self, loop):
        self.overall.output_excel('Assessment_slot_output_sheet')

        self.overall.write_in_excel(2, self.xl_dict[loop]['Applicant_id'],
                                    self.choose_data_applicant_id, None, 'Null')
        self.overall.write_in_excel(3, self.xl_dict[loop]['assigned_slot'],
                                    self.choose_slot_assigned, None, 'Null')
        self.overall.write_in_excel(4, self.xl_dict[loop]['updated_slot'],
                                    self.update_data, None, 'Null')
        self.overall.write_in_excel(5, self.xl_dict[loop]['dissociated'],
                                    self.dissociate_data, None, 'Null')
        self.overall.write_in_excel(6, self.xl_dict[loop]['assigned_message'],
                                    self.choose_error, None, 'No Message')
        self.overall.write_in_excel(7, self.xl_dict[loop]['update_message'],
                                    self.update_error, self.update_error_des, 'No Message')
        self.overall.write_in_excel(8, self.xl_dict[loop]['dissociated_message'],
                                    self.dissociate_data_error, self.unassign_error, 'No Message')


Object = InterviewSlots()
Object.excel_data()

Total_count = len(Object.dict_total)
print("Number of Rows::", Total_count)
for looping in range(0, Total_count):
    print("Iteration Count is ::", looping)
    Object.choose_interview_slot(looping)
    # Object.unassign_slot(looping)
    # Object.output_excel_status_headers(looping)

    # ------ Remove Dictionaries
    Object.choose_response = {}
    Object.choose_slot_data = {}
    Object.update_response = {}
    Object.update_slot_data = {}
    Object.unassign_response = {}
    Object.unassign_slot_data = {}

    Object.choose_data_applicant_id = ""
    Object.choose_error = ""
    Object.choose_slot_assigned = ""
    Object.update_data = ""
    Object.update_error = ""
    Object.update_error_des = ""
    Object.dissociate_data = ""
    Object.unassign_error = ""
    Object.dissociate_data_error = ""

# ----------------- Overall Output Status ------------------------------------------------------
Object.overall.main_headers = ['Comparison', 'Actual_status', 'Applicant ID', 'Choose Slot', 'Update Slot',
                               'UnAssign Slot', 'Choose Slot Message', 'Update Slot Message',
                               'UnAssign Slot Message']
Object.overall.headers_with_style2 = ['Comparison', 'Actual_status']
Object.overall.file_headers_col_row()
Object.overall.overall_status('ASSESSMENT SLOTS', Object.Expected_success_cases,
                              Object.start_time, Object.calling_lambda, 'assessmentSlots',
                              Object.server, Total_count, 'Assessment_slot_output_sheet')
