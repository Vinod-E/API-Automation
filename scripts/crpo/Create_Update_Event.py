from hpro_automation import (login, work_book, input_paths, output_paths)
import datetime
import requests
import json
import xlrd


class CreateUpdateEvent(login.CommonLogin, work_book.WorkBook):

    def __init__(self):

        # ---------------------------------- Overall Status Run Date ---------------------------------------------------
        self.start_time = str(datetime.datetime.now())

        # --------------------------------- Inheritance Class Instance -------------------------------------------------
        super(CreateUpdateEvent, self).__init__()
        self.common_login('crpo')

        # --------------------------------- Overall status initialize variables ----------------------------------------
        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 24)))
        self.Expected_success_test_cases = list(map(lambda x: 'Pass', range(0, 10)))
        self.Expected_success_update_test_cases = list(map(lambda x: 'Pass', range(0, 10)))
        self.Actual_Success_case = []
        self.test_case_each_property = []
        self.update_test_case_each_property = []
        self.update_rowsize = 27
        self.update_Col_01 = 0
        self.update_Col_02 = 8
        self.rowsize1 = self.rowsize + 26

        # --------------------------------- Excel Data initialize variables --------------------------------------------
        self.xl_EventName = []
        self.xl_event_type = []
        self.xl_req_id = []
        self.xl_event_status = []
        self.xl_event_from = []
        self.xl_event_to = []
        self.xl_event_em_id = []
        self.xl_campus_venue_id = []
        self.xl_address = []
        self.xl_city_id = []
        self.xl_province_id = []
        self.xl_campuses = []
        self.xl_mjrcs = []
        self.xl_slot_id = []
        self.xl_TypeOfEvent = []

        self.xl_expected_Type = []
        self.xl_expected_req = []
        self.xl_expected_sate = []
        self.xl_expected_city = []
        self.xl_expected_colleges = []
        self.xl_expected_Status = []

        # --------------------------------- Update Data initialize variables -------------------------------------------
        self.xl_update_EventName = []
        self.xl_update_event_type = []
        self.xl_update_event_status = []
        self.xl_update_slot_id = []
        self.xl_update_event_from = []
        self.xl_update_event_to = []
        self.xl_update_event_em_id = []
        self.xl_update_campus_venue_id = []
        self.xl_update_address = []
        self.xl_update_city_id = []
        self.xl_update_province_id = []
        self.xl_update_campuses = []

        self.xl_update_expected_Type = []
        self.xl_update_expected_req = []
        self.xl_update_expected_sate = []
        self.xl_update_expected_city = []
        self.xl_update_expected_colleges = []
        self.xl_update_expected_Status = []
        # --------------------------------- Dictionary initialize variables --------------------------------------------
        self.success_case_01 = {}
        self.success_case_02 = {}
        self.success_case_03 = {}
        self.headers = {}
        self.event_id_dict = {}
        self.event_details_dict = {}
        self.owners_dict = {}
        self.update_event_id_dict = {}

    def excel_headers(self):

        # --------------------------------- Excel Headers and Cell color, styles ---------------------------------------
        self.main_headers = ['Comparison', 'Status', 'EventID', 'EventName', 'EventType', 'City', 'State', 'EventFrom',
                             'EventTo', 'SlotId', 'Requirement Name', 'EventManagerId', 'Campuses', 'Address',
                             'Event_Status']
        self.headers_with_style2 = ['Comparison', 'Status']
        self.headers_with_style9 = ['EventID', 'EventName', 'EventFrom', 'EventTo', 'EventManagerId', 'SlotId',
                                    'EventType', 'City', 'State', 'Requirement Name', 'Campuses', 'Address',
                                    'Event_Status']
        self.file_headers_col_row()

    def excel_data(self):

        # ------------------------------- Excel Data Read --------------------------------------------------------------
        try:
            workbook = xlrd.open_workbook(input_paths.inputpaths['Event_Input_sheet'])
            sheet1 = workbook.sheet_by_index(0)
            for m in range(1, sheet1.nrows):
                number = m  # Counting number of rows
                rows = sheet1.row_values(number)

                if not rows[0]:
                    self.xl_EventName.append(None)
                else:
                    self.xl_EventName.append(rows[0])
                if rows[1] is None and rows[1] == '':
                    self.xl_event_type.append(None)
                else:
                    self.xl_event_type.append(int(rows[1]))
                if not rows[2]:
                    self.xl_req_id.append(None)
                else:
                    self.xl_req_id.append(int(rows[2]))
                if not rows[3]:
                    self.xl_event_status.append(None)
                else:
                    self.xl_event_status.append(int(rows[3]))
                if not rows[4]:
                    self.xl_event_from.append(None)
                else:
                    self.xl_event_from.append(rows[4])
                if not rows[5]:
                    self.xl_event_to.append(None)
                else:
                    self.xl_event_to.append(rows[5])
                if not rows[6]:
                    self.xl_event_em_id.append(None)
                else:
                    self.xl_event_em_id.append(int(rows[6]))
                if not rows[7]:
                    self.xl_campus_venue_id.append(None)
                else:
                    self.xl_campus_venue_id.append(int(rows[7]))
                if not rows[8]:
                    self.xl_address.append(None)
                else:
                    self.xl_address.append(rows[8])
                if not rows[9]:
                    self.xl_city_id.append(None)
                else:
                    self.xl_city_id.append(int(rows[9]))
                if not rows[10]:
                    self.xl_province_id.append(None)
                else:
                    self.xl_province_id.append(int(rows[10]))

                if rows[11] != '':
                    campuses = list(map(int, rows[11].split(',') if isinstance(rows[11], str) else [rows[11]]))
                    self.xl_campuses.append(campuses)
                else:
                    self.xl_campuses.append(list(rows[11]))

                if not rows[12]:
                    self.xl_mjrcs.append(None)
                else:
                    self.xl_mjrcs.append(int(rows[12]))
                if not rows[13]:
                    self.xl_slot_id.append(None)
                else:
                    self.xl_slot_id.append(int(rows[13]))
                if not rows[14]:
                    self.xl_TypeOfEvent.append(None)
                else:
                    self.xl_TypeOfEvent.append(int(rows[14]))
                if not rows[15]:
                    self.xl_expected_Type.append(None)
                else:
                    self.xl_expected_Type.append(rows[15])
                if not rows[16]:
                    self.xl_expected_req.append(None)
                else:
                    self.xl_expected_req.append(rows[16])
                if not rows[17]:
                    self.xl_expected_sate.append(None)
                else:
                    self.xl_expected_sate.append(rows[17])
                if not rows[18]:
                    self.xl_expected_city.append(None)
                else:
                    self.xl_expected_city.append(rows[18])
                if not rows[19]:
                    self.xl_expected_colleges.append(None)
                else:
                    self.xl_expected_colleges.append(rows[19])
                if not rows[20]:
                    self.xl_expected_Status.append(None)
                else:
                    self.xl_expected_Status.append(rows[20])

        except IOError:
            print("File not found or path is incorrect")

    def update_excel_data(self):

        # ------------------------------- Excel Data Read --------------------------------------------------------------
        try:
            workbook = xlrd.open_workbook(input_paths.inputpaths['Event_Input_sheet'])
            sheet1 = workbook.sheet_by_index(1)
            for m in range(1, sheet1.nrows):
                number = m  # Counting number of rows
                rows = sheet1.row_values(number)

                if not rows[0]:
                    self.xl_update_EventName.append(None)
                else:
                    self.xl_update_EventName.append(rows[0])
                if rows[1] is None and rows[1] == '':
                    self.xl_update_event_type.append(None)
                else:
                    self.xl_update_event_type.append(int(rows[1]))
                if not rows[2]:
                    self.xl_update_event_status.append(None)
                else:
                    self.xl_update_event_status.append(int(rows[2]))
                if not rows[3]:
                    self.xl_update_slot_id.append(None)
                else:
                    self.xl_update_slot_id.append(int(rows[3]))
                if not rows[4]:
                    self.xl_update_event_from.append(None)
                else:
                    self.xl_update_event_from.append(rows[4])
                if not rows[5]:
                    self.xl_update_event_to.append(None)
                else:
                    self.xl_update_event_to.append(rows[5])
                if not rows[6]:
                    self.xl_update_event_em_id.append(None)
                else:
                    self.xl_update_event_em_id.append(int(rows[6]))
                if not rows[7]:
                    self.xl_update_campus_venue_id.append(None)
                else:
                    self.xl_update_campus_venue_id.append(int(rows[7]))
                if not rows[8]:
                    self.xl_update_address.append(None)
                else:
                    self.xl_update_address.append(rows[8])
                if not rows[9]:
                    self.xl_update_city_id.append(None)
                else:
                    self.xl_update_city_id.append(int(rows[9]))
                if not rows[10]:
                    self.xl_update_province_id.append(None)
                else:
                    self.xl_update_province_id.append(int(rows[10]))

                if rows[11] != '':
                    campuses = list(map(int, rows[11].split(',') if isinstance(rows[11], str) else [rows[11]]))
                    self.xl_update_campuses.append(campuses)
                else:
                    self.xl_update_campuses.append(list(rows[11]))
                if not rows[12]:
                    self.xl_update_expected_Type.append(None)
                else:
                    self.xl_update_expected_Type.append(rows[12])
                if not rows[13]:
                    self.xl_update_expected_req.append(None)
                else:
                    self.xl_update_expected_req.append(rows[13])
                if not rows[14]:
                    self.xl_update_expected_sate.append(None)
                else:
                    self.xl_update_expected_sate.append(rows[14])
                if not rows[15]:
                    self.xl_update_expected_city.append(None)
                else:
                    self.xl_update_expected_city.append(rows[15])
                if not rows[16]:
                    self.xl_update_expected_colleges.append(None)
                else:
                    self.xl_update_expected_colleges.append(rows[16])
                if not rows[17]:
                    self.xl_update_expected_Status.append(None)
                else:
                    self.xl_update_expected_Status.append(rows[17])

        except IOError:
            print("File not found or path is incorrect")

    def create_event(self, loop):

        self.lambda_function('createEvent')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- API request --------------------------------------------------------------
        if self.xl_event_type[loop] == 2:
            request = {"event": {"name": self.xl_EventName[loop],
                                 "requirementId": self.xl_req_id[loop],
                                 "status": self.xl_event_status[loop],
                                 "slotId": self.xl_slot_id[loop],
                                 "dates": {"from": self.xl_event_from[loop],
                                           "to": self.xl_event_to[loop]},
                                 "eventManagerId": self.xl_event_em_id[loop],
                                 "type": self.xl_event_type[loop],
                                 "campusVenueId": self.xl_campus_venue_id[loop],
                                 "address": self.xl_address[loop],
                                 "cityId": self.xl_city_id[loop],
                                 "isEcEnabled": False,
                                 "provinceId": self.xl_province_id[loop],
                                 "mjrcs": [{"mjrId": self.xl_mjrcs[loop]}]}}
        else:
            request = {"event": {"name": self.xl_EventName[loop],
                                 "requirementId": self.xl_req_id[loop],
                                 "status": self.xl_event_status[loop],
                                 "slotId": self.xl_slot_id[loop],
                                 "dates": {"from": self.xl_event_from[loop],
                                           "to": self.xl_event_to[loop]},
                                 "eventManagerId": self.xl_event_em_id[loop],
                                 "type": self.xl_event_type[loop],
                                 "campusVenueId": self.xl_campus_venue_id[loop],
                                 "address": self.xl_address[loop],
                                 "cityId": self.xl_city_id[loop],
                                 "isEcEnabled": False,
                                 "provinceId": self.xl_province_id[loop],
                                 "campuses": self.xl_campuses[loop],
                                 "mjrcs": [{"mjrId": self.xl_mjrcs[loop]}]}}

        create_event_api = requests.post(self.webapi, headers=self.headers,
                                         data=json.dumps(request, default=str), verify=False)
        print(create_event_api.headers)
        create_event_api_response = json.loads(create_event_api.content)
        self.event_id_dict = create_event_api_response.get('data')
        print(self.event_id_dict)

    def update_event(self, loop):

        self.lambda_function('updateEvent')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"updateEvent": {"name": self.xl_update_EventName[loop],
                                   "status": self.xl_update_event_status[loop],
                                   "slotId": self.xl_update_slot_id[loop],
                                   "dates": {
                                       "from": self.xl_update_event_from[loop],
                                       "to": self.xl_update_event_to[loop]
                                   },
                                   "eventManagerId": self.xl_update_event_em_id[loop],
                                   "type": self.xl_update_event_type[loop],
                                   "campusVenueId": self.xl_update_campus_venue_id[loop],
                                   "address": self.xl_update_address[loop],
                                   "cityId": self.xl_update_city_id[loop],
                                   "isEcEnabled": False,
                                   "provinceId": self.xl_update_province_id[loop],
                                   "campuses": self.xl_update_campuses[loop],
                                   "id": self.event_id_dict['createdId']
                                   }}

        update_event_api = requests.post(self.webapi, headers=self.headers,
                                         data=json.dumps(request, default=str), verify=False)
        print(update_event_api.headers)
        update_event_api_response = json.loads(update_event_api.content)
        self.update_event_id_dict = update_event_api_response.get('data')
        print(self.update_event_id_dict)

    def get_all_event(self, loop):
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
                "Ids": [self.event_id_dict.get('createdId')],
                "TypeOfEvent": self.xl_TypeOfEvent[loop]
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

    def output_report(self, loop):

        # --------------------------------- Writing Input Data ---------------------------------------------------------

        self.ws.write(self.rowsize, self.col, 'Input', self.style4)
        self.ws.write(self.rowsize, 3, self.xl_EventName[loop] if self.xl_EventName[loop] else 'NA')
        self.ws.write(self.rowsize, 4, self.xl_expected_Type[loop] if self.xl_expected_Type[loop] else 'NA')
        self.ws.write(self.rowsize, 5, self.xl_expected_city[loop] if self.xl_expected_city[loop] else 'NA')
        self.ws.write(self.rowsize, 6, self.xl_expected_sate[loop] if self.xl_expected_sate[loop] else 'NA')
        self.ws.write(self.rowsize, 7, self.xl_event_from[loop] if self.xl_event_from[loop] else 'NA')
        self.ws.write(self.rowsize, 8, self.xl_event_to[loop] if self.xl_event_to[loop] else 'NA')
        self.ws.write(self.rowsize, 9, self.xl_slot_id[loop] if self.xl_slot_id[loop] else 'NA')
        self.ws.write(self.rowsize, 10, self.xl_expected_req[loop] if self.xl_expected_req[loop] else 'NA')
        self.ws.write(self.rowsize, 11, self.xl_event_em_id[loop] if self.xl_event_em_id[loop] else 'NA')
        self.ws.write(self.rowsize, 12, self.xl_expected_colleges[loop] if self.xl_expected_colleges[loop] else 'NA')
        self.ws.write(self.rowsize, 13, self.xl_address[loop] if self.xl_address[loop] else 'NA')
        self.ws.write(self.rowsize, 14, self.xl_expected_Status[loop] if self.xl_expected_Status[loop] else 'NA')

        self.rowsize += 1

        # --------------------------------- Writing Output Data --------------------------------------------------------
        self.ws.write(self.rowsize, self.col, 'Output', self.style5)

        # --------------------------------- test_case_property_wise_status ---------------------------------------------
        if self.event_details_dict['Name'] == self.xl_EventName[loop]:
            self.test_case_each_property.append('Pass')
        if self.event_details_dict['Type'] == self.xl_expected_Type[loop]:
            self.test_case_each_property.append('Pass')
        if self.event_details_dict['City'] == self.xl_expected_city[loop]:
            self.test_case_each_property.append('Pass')
        if self.event_details_dict['Province'] == self.xl_expected_sate[loop]:
            self.test_case_each_property.append('Pass')
        if self.event_details_dict['SlotId'] == self.xl_slot_id[loop]:
            self.test_case_each_property.append('Pass')
        if self.event_details_dict['ReqName'] == self.xl_expected_req[loop]:
            self.test_case_each_property.append('Pass')
        if self.event_details_dict['Status'] == self.xl_expected_Status[loop]:
            self.test_case_each_property.append('Pass')
        if self.event_details_dict['Adds'] == self.xl_address[loop]:
            self.test_case_each_property.append('Pass')
        if self.event_details_dict['CText'] == self.xl_expected_colleges[loop]:
            self.test_case_each_property.append('Pass')
        if self.owners_dict['UserId'] == self.xl_event_em_id[loop]:
            self.test_case_each_property.append('Pass')
        # --------------------------------- test_case_property_wise_status ---------------------------------------------

        if self.test_case_each_property == self.Expected_success_test_cases:
            self.ws.write(self.rowsize, 1, 'Pass', self.style8)
            self.success_case_01 = 'Pass'
        else:
            self.ws.write(self.rowsize, 1, 'Fail', self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_id_dict:
            self.ws.write(self.rowsize, 2, self.event_id_dict['createdId'], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            self.ws.write(self.rowsize, 3, self.event_details_dict['Name'], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            self.ws.write(self.rowsize, 4, self.event_details_dict['Type'], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            self.ws.write(self.rowsize, 5, self.event_details_dict['City'], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            self.ws.write(self.rowsize, 6, self.event_details_dict['Province'], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            self.ws.write(self.rowsize, 7, self.event_details_dict['From'], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            self.ws.write(self.rowsize, 8, self.event_details_dict['To'], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            self.ws.write(self.rowsize, 9, self.event_details_dict['SlotId'], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            self.ws.write(self.rowsize, 10, self.event_details_dict['ReqName'], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.owners_dict:
            self.ws.write(self.rowsize, 11, self.owners_dict['UserId'], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            if self.event_details_dict['CText'] == self.xl_expected_colleges[loop]:
                if self.xl_expected_colleges[loop] is None:
                    self.ws.write(self.rowsize, 12, 'NA', self.style8)
                else:
                    self.ws.write(self.rowsize, 12, self.event_details_dict['CText'], self.style8)
            else:
                self.ws.write(self.rowsize, 12, self.event_details_dict['CText'], self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            if self.event_details_dict['Adds'] == self.xl_address[loop]:
                if self.xl_address[loop] is None:
                    self.ws.write(self.rowsize, 13, 'NA', self.style8)
                else:
                    self.ws.write(self.rowsize, 13, self.event_details_dict['Adds'], self.style8)
            else:
                self.ws.write(self.rowsize, 13, self.event_details_dict['Adds'], self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            self.ws.write(self.rowsize, 14, self.event_details_dict['Status'], self.style8)
        # ------------------------------------ OutPut File save --------------------------------------------------------
        self.rowsize += 1
        Object.wb_Result.save(output_paths.outputpaths['Event_output_sheet'])

        if self.success_case_01 == 'Pass':
            self.Actual_Success_case.append(self.success_case_01)

    def overall_status(self):
        self.ws.write(0, 0, 'Create / Update Event', self.style23)
        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
        else:
            self.ws.write(0, 1, 'Fail', self.style25)

        self.ws.write(0, 2, 'Start Time', self.style23)
        self.ws.write(0, 3, self.start_time, self.style26)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)

        # ---------------------------- OutPut File save with Overall Status --------------------------------------------
        Object.wb_Result.save(output_paths.outputpaths['Event_output_sheet'])

    def updated_output_report(self, loop):

        # --------------------------------- Writing Input Data ---------------------------------------------------------

        self.ws.write(self.rowsize1, self.col, 'Input', self.style4)
        self.ws.write(self.rowsize1, 3, self.xl_update_EventName[loop] if self.xl_update_EventName[loop] else 'NA')
        self.ws.write(self.rowsize1, 4,
                      self.xl_update_expected_Type[loop] if self.xl_update_expected_Type[loop] else 'NA')
        self.ws.write(self.rowsize1, 5,
                      self.xl_update_expected_city[loop] if self.xl_update_expected_city[loop] else 'NA')
        self.ws.write(self.rowsize1, 6,
                      self.xl_update_expected_sate[loop] if self.xl_update_expected_sate[loop] else 'NA')
        self.ws.write(self.rowsize1, 7, self.xl_update_event_from[loop] if self.xl_update_event_from[loop] else 'NA')
        self.ws.write(self.rowsize1, 8, self.xl_update_event_to[loop] if self.xl_update_event_to[loop] else 'NA')
        self.ws.write(self.rowsize1, 9, self.xl_update_slot_id[loop] if self.xl_update_slot_id[loop] else 'NA')
        self.ws.write(self.rowsize1, 10,
                      self.xl_update_expected_req[loop] if self.xl_update_expected_req[loop] else 'NA')
        self.ws.write(self.rowsize1, 11, self.xl_update_event_em_id[loop] if self.xl_update_event_em_id[loop] else 'NA')
        self.ws.write(self.rowsize1, 12,
                      self.xl_update_expected_colleges[loop] if self.xl_update_expected_colleges[loop] else 'NA')
        self.ws.write(self.rowsize1, 13, self.xl_update_address[loop] if self.xl_update_address[loop] else 'NA')
        self.ws.write(self.rowsize1, 14,
                      self.xl_update_expected_Status[loop] if self.xl_update_expected_Status[loop] else 'NA')

        self.rowsize1 += 1

        # --------------------------------- Writing Output Data --------------------------------------------------------
        self.ws.write(self.rowsize1, self.col, 'Output', self.style5)

        # --------------------------------- test_case_property_wise_status ---------------------------------------------
        if self.event_details_dict['Name'] == self.xl_update_EventName[loop]:
            self.update_test_case_each_property.append('Pass')
        if self.event_details_dict['Type'] == self.xl_update_expected_Type[loop]:
            self.update_test_case_each_property.append('Pass')
        if self.event_details_dict['City'] == self.xl_update_expected_city[loop]:
            self.update_test_case_each_property.append('Pass')
        if self.event_details_dict['Province'] == self.xl_update_expected_sate[loop]:
            self.update_test_case_each_property.append('Pass')
        if self.event_details_dict['SlotId'] == self.xl_update_slot_id[loop]:
            self.update_test_case_each_property.append('Pass')
        if self.event_details_dict['ReqName'] == self.xl_update_expected_req[loop]:
            self.update_test_case_each_property.append('Pass')
        if self.event_details_dict['Status'] == self.xl_update_expected_Status[loop]:
            self.update_test_case_each_property.append('Pass')
        if self.event_details_dict['Adds'] == self.xl_update_address[loop]:
            self.update_test_case_each_property.append('Pass')
        if self.event_details_dict['CText'] == self.xl_update_expected_colleges[loop]:
            self.update_test_case_each_property.append('Pass')
        if self.owners_dict['UserId'] == self.xl_update_event_em_id[loop]:
            self.update_test_case_each_property.append('Pass')
        # --------------------------------- test_case_property_wise_status ---------------------------------------------

        if self.update_test_case_each_property == self.Expected_success_update_test_cases:
            self.ws.write(self.rowsize1, 1, 'Pass', self.style8)
            self.success_case_02 = 'Pass'
        else:
            self.ws.write(self.rowsize1, 1, 'Fail', self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.update_event_id_dict:
            self.ws.write(self.rowsize1, 2, self.update_event_id_dict['updatedId'], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            self.ws.write(self.rowsize1, 3, self.event_details_dict['Name'], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            self.ws.write(self.rowsize1, 4, self.event_details_dict['Type'], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            self.ws.write(self.rowsize1, 5, self.event_details_dict['City'], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            self.ws.write(self.rowsize1, 6, self.event_details_dict['Province'], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            self.ws.write(self.rowsize1, 7, self.event_details_dict['From'], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            self.ws.write(self.rowsize1, 8, self.event_details_dict['To'], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            self.ws.write(self.rowsize1, 9, self.event_details_dict['SlotId'], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            self.ws.write(self.rowsize1, 10, self.event_details_dict['ReqName'], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.owners_dict:
            self.ws.write(self.rowsize1, 11, self.owners_dict['UserId'], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            if self.event_details_dict['CText'] == self.xl_update_expected_colleges[loop]:
                if self.xl_update_expected_colleges[loop] is None:
                    self.ws.write(self.rowsize1, 12, 'NA', self.style8)
                else:
                    self.ws.write(self.rowsize1, 12, self.event_details_dict['CText'], self.style8)
            else:
                self.ws.write(self.rowsize1, 12, self.event_details_dict['CText'], self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            if self.event_details_dict['Adds'] == self.xl_update_address[loop]:
                if self.xl_update_address[loop] is None:
                    self.ws.write(self.rowsize1, 13, 'NA', self.style8)
                else:
                    self.ws.write(self.rowsize1, 13, self.event_details_dict['Adds'], self.style8)
            else:
                self.ws.write(self.rowsize1, 13, self.event_details_dict['Adds'], self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.event_details_dict:
            self.ws.write(self.rowsize1, 14, self.event_details_dict['Status'], self.style8)
        # ------------------------------------ OutPut File save --------------------------------------------------------
        self.rowsize1 += 1
        Object.wb_Result.save(output_paths.outputpaths['Event_output_sheet'])

        if self.success_case_02 == 'Pass':
            self.Actual_Success_case.append(self.success_case_01)

    def update_row_col_set_1(self):
        self.ws.write(self.update_rowsize, self.update_Col_01, '***', self.style29)
        self.update_Col_01 += 1

    def update_row_col_set_2(self):
        self.ws.write(self.update_rowsize, self.update_Col_02, '***', self.style29)
        self.update_Col_02 += 1


Object = CreateUpdateEvent()
Object.excel_headers()
Object.excel_data()
Object.update_excel_data()
Total_count = len(Object.xl_EventName)
Total_count1 = len(Object.xl_update_EventName)
print("Number Of Create Rows ::", Total_count)
print("Number Of Updated Rows ::", Total_count1)
if Object.login == 'OK':
    for looping in range(0, Total_count):
        print("Iteration Count is ::", looping)
        Object.create_event(looping)
        Object.get_all_event(looping)
        Object.output_report(looping)

        Object.event_details_dict = {}

        if Object.event_id_dict['createdId']:
            Object.update_event(looping)
            Object.get_all_event(looping)
            Object.updated_output_report(looping)

        # ----------------- Make Dictionaries clear for each loop ------------------------------------------------------
        Object.success_case_01 = {}
        Object.success_case_02 = {}
        Object.headers = {}
        Object.event_details_dict = {}
        Object.owners_dict = {}
        Object.test_case_each_property = []
        Object.update_test_case_each_property = []
        Object.update_event_id_dict = {}

for i in range(0, 5):
    Object.update_row_col_set_1()
for i in range(0, 7):
    Object.update_row_col_set_2()
Object.ws.write(Object.update_rowsize, 5, 'Event', Object.style29)
Object.ws.write(Object.update_rowsize, 6, 'Updated', Object.style29)
Object.ws.write(Object.update_rowsize, 7, 'Details', Object.style29)
Object.wb_Result.save(output_paths.outputpaths['Event_output_sheet'])

# ---------------------------- Call this function at last --------------------------------------------------------------
Object.overall_status()
