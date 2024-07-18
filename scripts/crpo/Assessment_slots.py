from hpro_automation import (login, work_book, input_paths)
import requests
import json
import xlrd


class AssessmentSlots(login.CommonLogin, work_book.WorkBook):

    def __init__(self):
        # --------------------------------- Inheritance Class Instance -------------------------------------------------
        super(AssessmentSlots, self).__init__()
        # self.common_login('admin')
        # self.slot_captcha_login_token('assessment')

        # --------------------------------- Excel Data initialize variables --------------------------------------------
        self.xl_unslot_request = []
        self.xl_choose_slot_request = []

    def excel_data(self):

        # ------------------------------- Excel Data Read --------------------------------------------------------------
        try:
            workbook = xlrd.open_workbook(input_paths.inputpaths['assessment_slot_input_sheet'])
            sheet1 = workbook.sheet_by_index(0)
            for i in range(1, sheet1.nrows):
                number = i  # Counting number of rows
                rows = sheet1.row_values(number)

                if not rows[0]:
                    self.xl_unslot_request.append(None)
                else:
                    self.xl_unslot_request.append(rows[0])
                if not rows[1]:
                    self.xl_choose_slot_request.append(None)
                else:
                    self.xl_choose_slot_request.append(rows[0])
        except IOError:
            print("File not found or path is incorrect")

    def choose_assessment_slot(self, loop):
        self.slot_captcha_login_token('assessment')
        self.lambda_function('assessment_slot_select')

        # ----------------------------------- API request --------------------------------------------------------------
        print("------------- Choose assessment slot API Call -----------------")
        request = self.xl_choose_slot_request[loop]
        choose_slot_api = requests.post(self.webapi, headers=self.headers,
                                        data=request, verify=False)
        # print(choose_slot_api.headers)
        response = json.loads(choose_slot_api.content)
        print(response)

    def unassign_slot(self, loop):
        self.common_login('admin')
        self.lambda_function('assessment_unassign_slot')

        # ----------------------------------- API request --------------------------------------------------------------
        request = self.xl_unslot_request[loop]

        unassign_slot_api = requests.post(self.webapi, headers=self.headers,
                                          data=request, verify=False)
        # print(unassign_slot_api.headers)
        response = json.loads(unassign_slot_api.content)
        print(response)


Object = AssessmentSlots()
Object.excel_data()
Total_count = len(Object.xl_unslot_request)
print("Number Of Rows ::", Total_count)
for looping in range(0, Total_count):
    print("Iteration Count is ::", looping)
    Object.choose_assessment_slot(looping)
    Object.unassign_slot(looping)

    # ----------------- Make Dictionaries clear for each loop ------------------------------------------------------
