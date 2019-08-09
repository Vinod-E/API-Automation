from hpro_automation import (login, work_book, input_paths, output_paths)
import datetime
import requests
import json
import xlrd


class CloneEvent(login.CommonLogin, work_book.WorkBook):

    def __init__(self):

        # ---------------------------------- Overall Status Run Date ---------------------------------------------------
        self.start_time = str(datetime.datetime.now())

        # --------------------------------- Inheritance Class Instance -------------------------------------------------
        super(CloneEvent, self).__init__()
        self.common_login('crpo')

        # --------------------------------- Overall status initialize variables ----------------------------------------
        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 66)))
        self.Actual_Success_case = []

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
        self.owners_dict = {}
        self.success_case_01 = {}
        self.success_case_02 = {}
        self.success_case_03 = {}
        self.headers = {}

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
        self.headers['APP-NAME'] = 'crpo'

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
        self.headers['APP-NAME'] = 'crpo'

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
        self.lambda_function('getAllEvent')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"eventIds": [self.api_clone_event_id]}
        get_ec_configs_api = requests.post(self.webapi, headers=self.headers,
                                           data=json.dumps(request, default=str), verify=False)
        print(get_ec_configs_api)
        get_ec_configs_api_response = json.loads(get_ec_configs_api.content)
        print(get_ec_configs_api_response)

    def get_assessment_summary(self):
        self.lambda_function('getAllEvent')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"eventIds": [self.api_clone_event_id]}
        get_assessment_summary_api = requests.post(self.webapi, headers=self.headers,
                                                   data=json.dumps(request, default=str), verify=False)
        print(get_assessment_summary_api)
        get_assessment_summary_api_response = json.loads(get_assessment_summary_api.content)
        print(get_assessment_summary_api_response)

    def get_event_registration_dates(self):
        self.lambda_function('getAllEvent')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"eventIds": [self.api_clone_event_id]}
        get_event_registration_dates_api = requests.post(self.webapi, headers=self.headers,
                                                         data=json.dumps(request, default=str), verify=False)
        print(get_event_registration_dates_api)
        get_event_registration_dates_api_response = json.loads(get_event_registration_dates_api.content)
        print(get_event_registration_dates_api_response)

    def output_report(self, loop):

        # --------------------------------- Writing Input Data ---------------------------------------------------------

        self.ws.write(self.rowsize, self.col, 'Input', self.style4)
        self.ws.write(self.rowsize, 2, self.api_clone_event_id if self.api_clone_event_id else 'Empty')
        self.rowsize += 1

        # --------------------------------- Writing Output Data --------------------------------------------------------
        self.ws.write(self.rowsize, self.col, 'Output', self.style5)
        self.ws.write(self.rowsize, 3, self.api_clone_event_id if self.api_clone_event_id else 'Empty')
        self.rowsize += 1

        # ------------------------------------ OutPut File save --------------------------------------------------------
        Object.wb_Result.save(output_paths.outputpaths['Event_Clone_output_sheet'])

        if self.success_case_01 == 'Pass':
            self.Actual_Success_case.append(self.success_case_01)
        if self.success_case_02 == 'Pass':
            self.Actual_Success_case.append(self.success_case_02)
        if self.success_case_03 == 'Pass':
            self.Actual_Success_case.append(self.success_case_03)

    def overall_status(self):
        self.ws.write(0, 0, 'Upload Candidates', self.style23)
        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
        else:
            self.ws.write(0, 1, 'Fail', self.style25)

        self.ws.write(0, 2, 'Start Time', self.style23)
        self.ws.write(0, 3, self.start_time, self.style26)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)

        # ---------------------------- OutPut File save with Overall Status --------------------------------------------
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
        Object.success_case_01 = {}
        Object.success_case_02 = {}
        Object.success_case_03 = {}
        Object.headers = {}

# ---------------------------- Call this function at last --------------------------------------------------------------
Object.overall_status()
