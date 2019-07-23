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
        self.common_login('v')

        # --------------------------------- Overall status initialize variables ----------------------------------------
        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 66)))
        self.Actual_Success_case = []

        # --------------------------------- Excel Data initialize variables --------------------------------------------
        self.xl_example = []

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
            workbook = xlrd.open_workbook(input_paths.inputpaths['Event_Input_sheet'])
            sheet1 = workbook.sheet_by_index(0)
            for i in range(1, sheet1.nrows):
                number = i  # Counting number of rows
                rows = sheet1.row_values(number)

                if not rows[0]:
                    self.xl_example.append(None)
                else:
                    self.xl_example.append(rows[0])

        except IOError:
            print("File not found or path is incorrect")

    def create_event(self, loop):

        self.lambda_function('createEvent')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"event": {"name": "IIINNCCampus",
                             "requirementId": 782,
                             "status": 1,
                             "slotId": 33778,
                             "dates": {"from": "2019-07-22 00:00:00",
                                       "to": "2019-07-30 00:00:00"},
                             "eventManagerId": 15022,
                             "type": 0,
                             "campusVenueId": 318,
                             "address": None,
                             "cityId": 25151,
                             "isEcEnabled": True,
                             "provinceId": 25075,
                             "campuses": [318],
                             "mjrcs": [{"mjrId": 1390}]}}

        create_event_api = requests.post(self.webapi, headers=self.headers,
                                         data=json.dumps(request, default=str), verify=False)
        create_event_api_response = json.loads(create_event_api.content)
        print(create_event_api_response)

    def output_report(self, loop):

        # --------------------------------- Writing Input Data ---------------------------------------------------------

        self.ws.write(self.rowsize, self.col, 'Input', self.style4)
        self.ws.write(self.rowsize, 1, self.xl_example[loop] if self.xl_example[loop] else 'Empty')
        self.rowsize += 1

        # --------------------------------- Writing Output Data --------------------------------------------------------
        self.ws.write(self.rowsize, self.col, 'Output', self.style5)
        self.rowsize += 1

        # ------------------------------------ OutPut File save --------------------------------------------------------
        Object.wb_Result.save(output_paths.outputpaths['Event_output_sheet'])

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
        Object.wb_Result.save(output_paths.outputpaths['Event_output_sheet'])


Object = CreateUpdateEvent()
Object.excel_headers()
Object.excel_data()
Total_count = len(Object.xl_example)
print("Number Of Rows ::", Total_count)
if Object.login == 'OK':
    for looping in range(0, Total_count):
        print("Iteration Count is ::", looping)
        # Object.create_event(looping)
        Object.output_report(looping)

        # ----------------- Make Dictionaries clear for each loop ------------------------------------------------------
        Object.success_case_01 = {}
        Object.success_case_02 = {}
        Object.success_case_03 = {}
        Object.headers = {}

# ---------------------------- Call this function at last --------------------------------------------------------------
Object.overall_status()
