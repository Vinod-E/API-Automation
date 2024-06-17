from hpro_automation import (login, work_book, input_paths, output_paths, db_login)
import datetime
import requests
import json
import xlrd


class ClassName(login.CommonLogin, work_book.WorkBook, db_login.DBConnection):

    def __init__(self):

        # ---------------------------------- Overall_Status Run Date ---------------------------------------------------
        self.start_time = str(datetime.datetime.now())

        # --------------------------------- Inheritance Class Instance -------------------------------------------------
        super(ClassName, self).__init__()
        self.common_login('crpo')
        self.db_connection('amsin')

        # --------------------------------- Overall status initialize variables ----------------------------------------
        self.Expected_overall_success_cases = list(map(lambda x: 'Pass', range(0, 7)))
        self.Actual_overall_success_case = []
        self.Expected_success_test_cases = list(map(lambda x: 'Pass', range(0, 7)))
        self.Actual_success_test_cases = []
        self.all_test_cases = ""

        # --------------------------------- Excel Data variables initialization ----------------------------------------
        self.xl_example = []

        # --------------------------------- Dictionary initialize variables --------------------------------------------
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

        self.lambda_function('example_api')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"ABC": self.xl_example[loop]}

        hit_api = requests.post(self.webapi, headers=self.headers,
                                data=json.dumps(request, default=str), verify=False)
        hitted_api_response = json.loads(hit_api.content)
        print(hitted_api_response)

    def output_report(self, loop):

        # --------------------------------- Writing Input Data ---------------------------------------------------------

        self.ws.write(self.rowsize, self.col, 'Input', self.style4)
        self.ws.write(self.rowsize, 1, self.xl_example[loop] if self.xl_example[loop] else 'Empty')
        self.rowsize += 1

        # --------------------------------- Writing Output Data --------------------------------------------------------
        self.ws.write(self.rowsize, self.col, 'Output', self.style5)

        # --------------------------------- Test cases wise status -----------------------------------------------------
        if self.Expected_success_test_cases == self.Actual_success_test_cases:
            self.all_test_cases = 'Pass'
            self.ws.write(self.rowsize, 1, 'Pass', self.style8)
        else:
            self.ws.write(self.rowsize, 1, 'Fail', self.style3)
        # ------------------------------ OutPut File save with Test cases wise Status ----------------------------------
        self.rowsize += 1
        Object.wb_Result.save(output_paths.outputpaths['Example_Output_sheet'])
        self.Actual_overall_success_case.append(self.all_test_cases)

    def overall_status(self):
        self.ws.write(0, 0, 'Use case Name', self.style23)
        if self.Expected_overall_success_cases == self.Actual_overall_success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
        else:
            self.ws.write(0, 1, 'Fail', self.style25)

        self.ws.write(0, 2, 'Start Time', self.style23)
        self.ws.write(0, 3, self.start_time, self.style26)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)

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
        Object.Actual_success_test_cases = {}
        Object.all_test_cases = ""
        Object.headers = {}

# ---------------------------- Call this function at last --------------------------------------------------------------
Object.overall_status()
