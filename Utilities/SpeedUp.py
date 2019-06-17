from hpro_automation import (login, work_book, input_paths, output_paths, api)
import datetime
import requests
import json
import xlrd


class ClassName(login.CRPOLogin, work_book.WorkBook):

    def __init__(self):

        # ---------------------------------- Overall Status Run Date ---------------------------------------------------
        self.start_time = str(datetime.datetime.now())

        # --------------------------------- Inheritance Class Instance -------------------------------------------------
        super(ClassName, self).__init__()

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
            workbook = xlrd.open_workbook(input_paths.inputpaths['Example_Input_sheet'])
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

    def api_call(self, loop):

        # ---------------- Passing headers based on API supports to lambda or not --------------------
        if self.calling_lambda == 'On':
            if api.lambda_apis['example_api'] is not None \
                    and api.web_api['example_api'] in api.lambda_apis['example_api']:
                self.headers = self.lambda_headers
            else:
                self.headers = self.Non_lambda_headers
        elif self.calling_lambda == 'Off':
            self.headers = self.lambda_headers
        else:
            self.headers = self.lambda_headers

        # ---------------- Updating headers with app name -----------------
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"ABC": self.xl_example[loop]}

        hit_api = requests.post(api.web_api['example_api'], headers=self.get_token,
                                data=json.dumps(request, default=str), verify=False)
        hitted_api_response = json.loads(hit_api.content)
        print(hitted_api_response)

    def output_report(self, loop):

        # --------------------------------- Writing Input Data ---------------------------------------------------------
        self.rowsize += 1
        self.ws.write(self.rowsize, self.col, 'Input', self.style4)
        self.ws.write(self.rowsize, self.xl_example[loop] if self.xl_example[loop] else 'Empty')

        # --------------------------------- Writing Output Data --------------------------------------------------------
        self.ws.write(self.rowsize, self.col, 'Output', self.style5)
        self.rowsize += 1

        # ------------------------------------ OutPut File save --------------------------------------------------------
        Object.wb_Result.save(output_paths.outputpaths['Example_Output_sheet'])

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

        self.ws.write(0, 3, 'Start Time', self.style23)
        self.ws.write(0, 4, self.start_time, self.style26)

        # ---------------------------- OutPut File save with Overall Status --------------------------------------------
        Object.wb_Result.save(output_paths.outputpaths['Example_Output_sheet'])


Object = ClassName()
Object.excel_headers()
Object.excel_data()
Total_count = len(Object.xl_example)
print("Number Of Rows ::", Total_count)
if Object.login == 'OK':
    for looping in range(0, Total_count):
        print("Iteration Count is ::", looping)
        Object.api_call(looping)
        Object.output_report(looping)

        # ----------------- Make Dictionaries clear for each loop ------------------------------------------------------
        Object.success_case_01 = {}
        Object.success_case_02 = {}
        Object.success_case_03 = {}
        Object.headers = {}

# ---------------------------- Call this function at last --------------------------------------------------------------
Object.overall_status()
