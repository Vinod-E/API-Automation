from hpro_automation import (login, work_book, input_paths, output_paths)
from hpro_automation.api import *
import datetime
import requests
import json
import xlrd


class CloneEvent(login.CommonLogin, work_book.WorkBook):

    def __init__(self):

        # ---------------------------------- Overall_Status Run Date ---------------------------------------------------
        self.start_time = str(datetime.datetime.now())

        # --------------------------------- Inheritance Class Instance -------------------------------------------------
        super(CloneEvent, self).__init__()
        self.common_login('admin')
        self.crpo_app_name = self.app_name.strip()
        print(self.crpo_app_name)

        # --------------------------------- Overall status initialize variables ----------------------------------------
        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 4)))
        self.Expected_success_test_cases = list(map(lambda x: 'Pass', range(0, 20)))
        self.Actual_Success_case = []
        self.test_case_each_field = []
        self.success_case_01 = {}

        # --------------------------------- Excel Data initialize variables --------------------------------------------
        self.xl_clone_event_request = []
        self.api_clone_event_id = []

        self.xl_expected_event_name = []
        self.xl_expected_event_from = []
        self.xl_expected_event_to = []
        self.xl_expected_req = []
        self.xl_expected_job = []
        self.xl_expected_test = []
        self.xl_expected_event_type = []
        self.xl_expected_event_slot = []
        self.xl_expected_colleges = []
        self.xl_expected_venue = []
        self.xl_expected_city = []
        self.xl_expected_state = []
        self.xl_expected_address = []
        self.xl_expected_owner_name = []
        self.xl_expected_owner_mail = []
        self.xl_expected_ec = []
        self.xl_expected_positivesatge = []
        self.xl_expected_positivestatus = []
        self.xl_expected_negativestage = []
        self.xl_expected_negativestatus = []
        self.xl_expected_registartionfrom = []
        self.xl_expected_registartionto = []

        # --------------------------------- Dictionary initialize variables --------------------------------------------
        self.event_details_dict = {}
        self.assessmentSummarys_dict = {}
        self.owners_dict = {}
        self.Ec_configs_dict = {}
        self.registration_dict = {}

    def excel_headers(self):

        # --------------------------------- Excel Headers and Cell color, styles ---------------------------------------
        self.main_headers = ['Comparision', 'Status', 'Event ID', 'Test ID', 'EventName', 'EventType', 'Requirement',
                             'Job', 'Test', 'EventFrom', 'EventTo', 'EventSlot', 'Colleges', 'Venue', 'City', 'State',
                             'Address', 'OwnerName', 'OwnerEmail', 'EC', 'PositiveStage', 'PositiveStatus',
                             'NegativeStage', 'NegativeStatus', 'RegistrationFrom', 'RegistrationTo']
        self.headers_with_style2 = ['Comparision', 'Status', 'Event ID', 'Test ID']
        self.headers_with_style9 = ['EventName', 'EventType', 'Requirement', 'Job', 'Test', 'EventFrom', 'EventTo',
                                    'EventSlot', 'Colleges', 'Venue', 'City', 'State', 'Address', 'OwnerName',
                                    'OwnerEmail']
        self.headers_with_style19 = ['EC', 'PositiveStage', 'PositiveStatus', 'NegativeStage', 'NegativeStatus']
        self.file_headers_col_row()

    def excel_data(self):

        # ------------------------------- Excel Data Read --------------------------------------------------------------
        try:
            workbook = xlrd.open_workbook(input_paths.inputpaths['CloneEvent_Input_sheet'])
            sheet1 = workbook.sheet_by_index(0)
            for i in range(1, sheet1.nrows):
                number = i  # Counting number of rows
                rows = sheet1.row_values(number)

                if not rows[0]:
                    self.xl_clone_event_request.append(None)
                else:
                    self.xl_clone_event_request.append(rows[0])

            sheet2 = workbook.sheet_by_index(1)
            for i in range(1, sheet2.nrows):
                number = i  # Counting number of rows
                rows = sheet2.row_values(number)

                if not rows[0]:
                    self.xl_expected_event_name.append(None)
                else:
                    self.xl_expected_event_name.append(rows[0])

                if not rows[1]:
                    self.xl_expected_event_from.append(None)
                else:
                    self.xl_expected_event_from.append(rows[1])

                if not rows[2]:
                    self.xl_expected_event_to.append(None)
                else:
                    self.xl_expected_event_to.append(rows[2])

                if not rows[3]:
                    self.xl_expected_req.append(None)
                else:
                    self.xl_expected_req.append(rows[3])

                if not rows[4]:
                    self.xl_expected_job.append(None)
                else:
                    self.xl_expected_job.append(rows[4])

                if not rows[5]:
                    self.xl_expected_test.append(None)
                else:
                    self.xl_expected_test.append(rows[5])

                if not rows[6]:
                    self.xl_expected_event_type.append(None)
                else:
                    self.xl_expected_event_type.append(rows[6])

                if not rows[7]:
                    self.xl_expected_event_slot.append(None)
                else:
                    self.xl_expected_event_slot.append(rows[7])

                if not rows[8]:
                    self.xl_expected_colleges.append(None)
                else:
                    self.xl_expected_colleges.append(rows[8])

                if not rows[9]:
                    self.xl_expected_venue.append(None)
                else:
                    self.xl_expected_venue.append(rows[9])

                if not rows[10]:
                    self.xl_expected_city.append(None)
                else:
                    self.xl_expected_city.append(rows[10])

                if not rows[11]:
                    self.xl_expected_state.append(None)
                else:
                    self.xl_expected_state.append(rows[11])

                if not rows[12]:
                    self.xl_expected_address.append(None)
                else:
                    self.xl_expected_address.append(rows[12])

                if not rows[13]:
                    self.xl_expected_owner_name.append(None)
                else:
                    self.xl_expected_owner_name.append(rows[13])

                if not rows[14]:
                    self.xl_expected_owner_mail.append(None)
                else:
                    self.xl_expected_owner_mail.append(rows[14])

                if not rows[15]:
                    self.xl_expected_ec.append(None)
                else:
                    self.xl_expected_ec.append(rows[15])

                if not rows[16]:
                    self.xl_expected_positivesatge.append(None)
                else:
                    self.xl_expected_positivesatge.append(rows[16])

                if not rows[17]:
                    self.xl_expected_positivestatus.append(None)
                else:
                    self.xl_expected_positivestatus.append(rows[17])

                if not rows[18]:
                    self.xl_expected_negativestage.append(None)
                else:
                    self.xl_expected_negativestage.append(rows[18])

                if not rows[19]:
                    self.xl_expected_negativestatus.append(None)
                else:
                    self.xl_expected_negativestatus.append(rows[19])

                if not rows[20]:
                    self.xl_expected_registartionfrom.append(None)
                else:
                    self.xl_expected_registartionfrom.append(rows[20])

                if not rows[21]:
                    self.xl_expected_registartionto.append(None)
                else:
                    self.xl_expected_registartionto.append(rows[21])

        except IOError:
            print("File not found or path is incorrect")

    def clone_event(self, loop):

        self.lambda_function('cloneEvent')
        self.headers['APP-NAME'] = self.crpo_app_name

        # ----------------------------------- API request --------------------------------------------------------------
        request = self.xl_clone_event_request[loop]

        clone_event_api = requests.post(self.webapi, headers=self.headers,
                                        data=request, verify=False)
        print(clone_event_api.headers)
        clone_event_api_response = json.loads(clone_event_api.content)
        self.api_clone_event_id = clone_event_api_response.get('clonedEventId')
        print(clone_event_api_response)

    def get_all_event(self):
        self.lambda_function('getAllEvent')
        self.headers['APP-NAME'] = self.crpo_app_name

        # ----------------------------------- API request --------------------------------------------------------------
        request = {
            "Paging": {
                "MaxResults": 1,
                "PageNumber": 1
            },
            "isAllEventRequired": True,
            "Search": {
                "Ids": [self.api_clone_event_id],
                "TypeOfEvent": 8
            },
            "Status": 0
        }

        get_all_event_api = requests.post(self.webapi, headers=self.headers,
                                          data=json.dumps(request, default=str), verify=False)
        print(get_all_event_api.headers)
        get_all_event_api_response = json.loads(get_all_event_api.content)

        if get_all_event_api_response.get('status') == 'OK':
            data = get_all_event_api_response.get('data')

            event_details_dict = data.get('Events')
            for k in event_details_dict:
                self.event_details_dict = k

                owners = k.get('Owners')
                for j in owners:
                    self.owners_dict = j

    def get_ec_configs(self):
        self.lambda_function('getEcConfigs')
        self.headers['APP-NAME'] = self.crpo_app_name

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"eventIds": [self.api_clone_event_id]}
        get_ec_configs_api = requests.post(self.webapi, headers=self.headers,
                                           data=json.dumps(request, default=str), verify=False)
        print(get_ec_configs_api.headers)
        get_ec_configs_api_response = json.loads(get_ec_configs_api.content)
        data = get_ec_configs_api_response.get('data')
        if get_ec_configs_api_response['status'] == 'OK':
            ec_config = data.get('EcConfigs')
            for i in ec_config:
                self.Ec_configs_dict = i

    def get_assessment_summary(self):
        self.lambda_function('getAssessmentSummary')
        self.headers['APP-NAME'] = self.crpo_app_name

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"eventId": self.api_clone_event_id}
        get_assessment_summary_api = requests.post(self.webapi, headers=self.headers,
                                                   data=json.dumps(request, default=str), verify=False)
        print(get_assessment_summary_api.headers)
        get_assessment_summary_api_response = json.loads(get_assessment_summary_api.content)
        data = get_assessment_summary_api_response.get('data')
        if get_assessment_summary_api_response['status'] == 'OK':
            ass_summary = data.get('assessmentSummarys')
            for i in ass_summary:
                self.assessmentSummarys_dict = i

    def get_event_registration_dates(self):
        self.lambda_function('getEventRegistrationDates')
        self.headers['APP-NAME'] = self.crpo_app_name

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"EventId": self.api_clone_event_id}
        get_event_registration_dates_api = requests.post(self.webapi, headers=self.headers,
                                                         data=json.dumps(request, default=str), verify=False)
        print(get_event_registration_dates_api.headers)
        get_event_registration_dates_api_response = json.loads(get_event_registration_dates_api.content)
        self.registration_dict = get_event_registration_dates_api_response.get('data')

    def output_report(self, loop):

        # --------------------------------- Writing Input Data ---------------------------------------------------------

        self.ws.write(self.rowsize, self.col, 'Input', self.style4)
        self.ws.write(self.rowsize, 4,
                      self.xl_expected_event_name[loop] if self.xl_expected_event_name[loop] else 'NA')
        self.ws.write(self.rowsize, 5,
                      self.xl_expected_event_type[loop] if self.xl_expected_event_type[loop] else 'NA')
        self.ws.write(self.rowsize, 6,
                      self.xl_expected_req[loop] if self.xl_expected_req[loop] else 'NA')
        self.ws.write(self.rowsize, 7,
                      self.xl_expected_job[loop] if self.xl_expected_job[loop] else 'NA')
        self.ws.write(self.rowsize, 8,
                      self.xl_expected_test[loop] if self.xl_expected_test[loop] else 'NA')
        self.ws.write(self.rowsize, 9,
                      self.xl_expected_event_from[loop] if self.xl_expected_event_from[loop] else 'NA')
        self.ws.write(self.rowsize, 10,
                      self.xl_expected_event_to[loop] if self.xl_expected_event_to[loop] else 'NA')
        self.ws.write(self.rowsize, 11,
                      self.xl_expected_event_slot[loop] if self.xl_expected_event_slot[loop] else 'NA')
        self.ws.write(self.rowsize, 12,
                      self.xl_expected_colleges[loop] if self.xl_expected_colleges[loop] else 'NA')
        self.ws.write(self.rowsize, 13,
                      self.xl_expected_venue[loop] if self.xl_expected_venue[loop] else 'NA')
        self.ws.write(self.rowsize, 14,
                      self.xl_expected_city[loop] if self.xl_expected_city[loop] else 'NA')
        self.ws.write(self.rowsize, 15,
                      self.xl_expected_state[loop] if self.xl_expected_state[loop] else 'NA')
        self.ws.write(self.rowsize, 16,
                      self.xl_expected_address[loop] if self.xl_expected_address[loop] else 'NA')
        self.ws.write(self.rowsize, 17,
                      self.xl_expected_owner_name[loop] if self.xl_expected_owner_name[loop] else 'NA')
        self.ws.write(self.rowsize, 18,
                      self.xl_expected_owner_mail[loop] if self.xl_expected_owner_mail[loop] else 'NA')
        self.ws.write(self.rowsize, 19,
                      self.xl_expected_ec[loop] if self.xl_expected_ec[loop] else 'NA')
        self.ws.write(self.rowsize, 20,
                      self.xl_expected_positivesatge[loop] if self.xl_expected_positivesatge[loop] else 'NA')
        self.ws.write(self.rowsize, 21,
                      self.xl_expected_positivestatus[loop] if self.xl_expected_positivestatus[loop] else 'NA')
        self.ws.write(self.rowsize, 22,
                      self.xl_expected_negativestage[loop] if self.xl_expected_negativestage[loop] else 'NA')
        self.ws.write(self.rowsize, 23,
                      self.xl_expected_negativestatus[loop] if self.xl_expected_negativestatus[loop] else 'NA')
        self.ws.write(self.rowsize, 24,
                      self.xl_expected_registartionfrom[loop] if self.xl_expected_registartionfrom[loop] else 'NA')
        self.ws.write(self.rowsize, 25,
                      self.xl_expected_registartionto[loop] if self.xl_expected_registartionto[loop] else 'NA')

        self.rowsize += 1
        # --------------------------------- Writing Output Data --------------------------------------------------------
        self.ws.write(self.rowsize, self.col, 'Output', self.style5)

        if self.event_details_dict.get('Name') == self.xl_expected_event_name[loop]:
            self.test_case_each_field.append('Pass')
        if self.event_details_dict.get('Type') == self.xl_expected_event_type[loop]:
            self.test_case_each_field.append('Pass')
        if self.event_details_dict.get('ReqName') == self.xl_expected_req[loop]:
            self.test_case_each_field.append('Pass')
        if self.assessmentSummarys_dict.get('jobRoleName') == self.xl_expected_job[loop]:
            self.test_case_each_field.append('Pass')
        if self.assessmentSummarys_dict.get('testName') == self.xl_expected_test[loop]:
            self.test_case_each_field.append('Pass')
        if self.event_details_dict.get('SlotName') == self.xl_expected_event_slot[loop]:
            self.test_case_each_field.append('Pass')
        if self.event_details_dict.get('CText') == self.xl_expected_colleges[loop]:
            self.test_case_each_field.append('Pass')
        if self.event_details_dict.get('Venue') == self.xl_expected_venue[loop]:
            self.test_case_each_field.append('Pass')
        if self.event_details_dict.get('City') == self.xl_expected_city[loop]:
            self.test_case_each_field.append('Pass')
        if self.event_details_dict.get('Province') == self.xl_expected_state[loop]:
            self.test_case_each_field.append('Pass')
        if self.event_details_dict.get('Adds') == self.xl_expected_address[loop]:
            self.test_case_each_field.append('Pass')
        if self.owners_dict.get('UserName') == self.xl_expected_owner_name[loop]:
            self.test_case_each_field.append('Pass')
        if self.owners_dict.get('UserEmail') == self.xl_expected_owner_mail[loop]:
            self.test_case_each_field.append('Pass')
        if self.Ec_configs_dict.get('EcName') == self.xl_expected_ec[loop]:
            self.test_case_each_field.append('Pass')
        if self.Ec_configs_dict.get('PositiveStage') == self.xl_expected_positivesatge[loop]:
            self.test_case_each_field.append('Pass')
        if self.Ec_configs_dict.get('PositiveStatus') == self.xl_expected_positivestatus[loop]:
            self.test_case_each_field.append('Pass')
        if self.Ec_configs_dict.get('NegativeStage') == self.xl_expected_negativestage[loop]:
            self.test_case_each_field.append('Pass')
        if self.Ec_configs_dict.get('NegativeStatus') == self.xl_expected_negativestatus[loop]:
            self.test_case_each_field.append('Pass')
        if self.registration_dict.get('DateFrom') == self.xl_expected_registartionfrom[loop]:
            self.test_case_each_field.append('Pass')
        if self.registration_dict.get('DateTo') == self.xl_expected_registartionto[loop]:
            self.test_case_each_field.append('Pass')
        # --------------------------------------------------------------------------------------------------------------
        if self.test_case_each_field == self.Expected_success_test_cases:
            self.ws.write(self.rowsize, 1, 'Pass', self.style8)
            self.success_case_01 = 'Pass'
        else:
            self.ws.write(self.rowsize, 1, 'Fail', self.style3)
        self.ws.write(self.rowsize, 2, self.api_clone_event_id, self.style8)
        self.ws.write(self.rowsize, 3, self.assessmentSummarys_dict.get('testId'), self.style8)
        # --------------------------------------------------------------------------------------------------------------

        if self.event_details_dict.get('Name') == self.xl_expected_event_name[loop]:
            self.ws.write(self.rowsize, 4, self.event_details_dict.get('Name'), self.style8)
        else:
            self.ws.write(self.rowsize, 4,
                          self.event_details_dict.get('Name') if self.event_details_dict.get('Name')
                          else 'NA', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.event_details_dict.get('Type') == self.xl_expected_event_type[loop]:
            self.ws.write(self.rowsize, 5, self.event_details_dict.get('Type'), self.style8)
        else:
            self.ws.write(self.rowsize, 5,
                          self.event_details_dict.get('Type') if self.event_details_dict.get('Type')
                          else 'NA', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.event_details_dict.get('ReqName') == self.xl_expected_req[loop]:
            self.ws.write(self.rowsize, 6, self.event_details_dict.get('ReqName'), self.style8)
        else:
            self.ws.write(self.rowsize, 6,
                          self.event_details_dict.get('ReqName') if self.event_details_dict.get('ReqName')
                          else 'NA', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.assessmentSummarys_dict.get('jobRoleName') == self.xl_expected_job[loop]:
            self.ws.write(self.rowsize, 7, self.assessmentSummarys_dict.get('jobRoleName'), self.style8)
        else:
            self.ws.write(self.rowsize, 7,
                          self.assessmentSummarys_dict.get('jobRoleName')
                          if self.assessmentSummarys_dict.get('jobRoleName') else 'NA', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.assessmentSummarys_dict.get('testName') == self.xl_expected_test[loop]:
            self.ws.write(self.rowsize, 8, self.assessmentSummarys_dict.get('testName'), self.style8)
        else:
            self.ws.write(self.rowsize, 8,
                          self.assessmentSummarys_dict.get('testName') if self.assessmentSummarys_dict.get('testName')
                          else 'NA', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.event_details_dict.get('From') == self.xl_expected_event_from[loop]:
            self.ws.write(self.rowsize, 9, self.event_details_dict.get('From'), self.style8)
        elif 'T' in self.event_details_dict.get('From'):
            self.ws.write(self.rowsize, 9,
                          self.event_details_dict.get('From') if self.event_details_dict.get('From')
                          else 'NA', self.style7)
        # --------------------------------------------------------------------------------------------------------------

        if self.event_details_dict.get('To') == self.xl_expected_event_to[loop]:
            self.ws.write(self.rowsize, 10, self.event_details_dict.get('To'), self.style8)
        elif 'T' in self.event_details_dict.get('To'):
            self.ws.write(self.rowsize, 10,
                          self.event_details_dict.get('To') if self.event_details_dict.get('To')
                          else 'NA', self.style7)
        # --------------------------------------------------------------------------------------------------------------

        if self.event_details_dict.get('SlotName') == self.xl_expected_event_slot[loop]:
            self.ws.write(self.rowsize, 11, self.event_details_dict.get('SlotName'), self.style8)
        else:
            self.ws.write(self.rowsize, 11,
                          self.event_details_dict.get('SlotName') if self.event_details_dict.get('SlotName')
                          else 'NA', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.event_details_dict.get('CText') == self.xl_expected_colleges[loop]:
            self.ws.write(self.rowsize, 12, self.event_details_dict.get('CText') if self.event_details_dict.get(
                'CText') else 'NA', self.style8)
        else:
            self.ws.write(self.rowsize, 12,
                          self.event_details_dict.get('CText') if self.event_details_dict.get('CText')
                          else 'NA', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.event_details_dict.get('Venue') == self.xl_expected_venue[loop]:
            self.ws.write(self.rowsize, 13, self.event_details_dict.get('Venue') if self.event_details_dict.get(
                'Venue') else 'NA', self.style8)
        else:
            self.ws.write(self.rowsize, 13,
                          self.event_details_dict.get('Venue') if self.event_details_dict.get('Venue')
                          else 'NA', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.event_details_dict.get('City') == self.xl_expected_city[loop]:
            self.ws.write(self.rowsize, 14, self.event_details_dict.get('City'), self.style8)
        else:
            self.ws.write(self.rowsize, 14,
                          self.event_details_dict.get('City') if self.event_details_dict.get('City')
                          else 'NA', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.event_details_dict.get('Province') == self.xl_expected_state[loop]:
            self.ws.write(self.rowsize, 15, self.event_details_dict.get('Province'), self.style8)
        else:
            self.ws.write(self.rowsize, 15,
                          self.event_details_dict.get('Province') if self.event_details_dict.get('Province')
                          else 'NA', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.event_details_dict.get('Adds') == self.xl_expected_address[loop]:
            self.ws.write(self.rowsize, 16, self.event_details_dict.get('Adds'), self.style8)
        else:
            self.ws.write(self.rowsize, 16,
                          self.event_details_dict.get('Adds') if self.event_details_dict.get('Adds')
                          else 'NA', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.owners_dict.get('UserName') == self.xl_expected_owner_name[loop]:
            self.ws.write(self.rowsize, 17, self.owners_dict.get('UserName'), self.style8)
        else:
            self.ws.write(self.rowsize, 17,
                          self.owners_dict.get('UserName') if self.owners_dict.get('UserName')
                          else 'NA', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.owners_dict.get('UserEmail') == self.xl_expected_owner_mail[loop]:
            self.ws.write(self.rowsize, 18, self.owners_dict.get('UserEmail'), self.style8)
        else:
            self.ws.write(self.rowsize, 18,
                          self.owners_dict.get('UserEmail') if self.owners_dict.get('UserEmail')
                          else 'NA', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.Ec_configs_dict.get('EcName') == self.xl_expected_ec[loop]:
            self.ws.write(self.rowsize, 19, self.Ec_configs_dict.get('EcName'), self.style8)
        else:
            self.ws.write(self.rowsize, 19,
                          self.Ec_configs_dict.get('EcName') if self.Ec_configs_dict.get('EcName')
                          else 'NA', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.Ec_configs_dict.get('PositiveStage') == self.xl_expected_positivesatge[loop]:
            self.ws.write(self.rowsize, 20, self.Ec_configs_dict.get('PositiveStage'), self.style8)
        else:
            self.ws.write(self.rowsize, 20,
                          self.Ec_configs_dict.get('PositiveStage') if self.Ec_configs_dict.get('PositiveStage')
                          else 'NA', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.Ec_configs_dict.get('PositiveStatus') == self.xl_expected_positivestatus[loop]:
            self.ws.write(self.rowsize, 21, self.Ec_configs_dict.get('PositiveStatus'), self.style8)
        else:
            self.ws.write(self.rowsize, 21,
                          self.Ec_configs_dict.get('PositiveStatus') if self.Ec_configs_dict.get('PositiveStatus')
                          else 'NA', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.Ec_configs_dict.get('NegativeStage') == self.xl_expected_negativestage[loop]:
            self.ws.write(self.rowsize, 22, self.Ec_configs_dict.get('NegativeStage'), self.style8)
        else:
            self.ws.write(self.rowsize, 22,
                          self.Ec_configs_dict.get('NegativeStage') if self.Ec_configs_dict.get('NegativeStage')
                          else 'NA', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.Ec_configs_dict.get('NegativeStatus') == self.xl_expected_negativestatus[loop]:
            self.ws.write(self.rowsize, 23, self.Ec_configs_dict.get('NegativeStatus'), self.style8)
        else:
            self.ws.write(self.rowsize, 23,
                          self.Ec_configs_dict.get('NegativeStatus') if self.Ec_configs_dict.get('NegativeStatus')
                          else 'NA', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.registration_dict.get('DateFrom') == self.xl_expected_registartionfrom[loop]:
            self.ws.write(self.rowsize, 24, self.registration_dict.get('DateFrom'), self.style8)
        else:
            self.ws.write(self.rowsize, 24,
                          self.registration_dict.get('DateFrom') if self.registration_dict.get('DateFrom')
                          else 'NA', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.registration_dict.get('DateTo') == self.xl_expected_registartionto[loop]:
            self.ws.write(self.rowsize, 25, self.registration_dict.get('DateTo'), self.style8)
        else:
            self.ws.write(self.rowsize, 25,
                          self.registration_dict.get('DateTo') if self.registration_dict.get('DateTo')
                          else 'NA', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        self.rowsize += 1
        # ------------------------------------ OutPut File save --------------------------------------------------------
        Object.wb_Result.save(output_paths.outputpaths['Event_Clone_output_sheet'])

        if self.success_case_01 == 'Pass':
            self.Actual_Success_case.append(self.success_case_01)

    def overall_status(self):
        self.ws.write(0, 0, 'Clone Event', self.style23)
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

        # ---------------------------- OutPut File save with Overall_Status --------------------------------------------
        Object.wb_Result.save(output_paths.outputpaths['Event_Clone_output_sheet'])


Object = CloneEvent()
Object.excel_headers()
Object.excel_data()
Total_count = len(Object.xl_clone_event_request)
print("Number Of Rows ::", Total_count)
if Object.login == 'OK':
    for looping in range(0, Total_count):
        print("Iteration Count is ::", looping)
        Object.clone_event(looping)
        Object.get_all_event()
        Object.get_ec_configs()
        Object.get_assessment_summary()
        Object.get_event_registration_dates()
        Object.output_report(looping)

        # ----------------- Make Dictionaries clear for each loop ------------------------------------------------------
        Object.event_details_dict = {}
        Object.owners_dict = {}
        Object.assessmentSummarys_dict = {}
        Object.Ec_configs_dict = {}
        Object.registration_dict = {}
        Object.owners_dict = {}
        Object.test_case_each_field = []
        Object.success_case_01 = {}

# ---------------------------- Call this function at last --------------------------------------------------------------
Object.overall_status()
