from hpro_automation import (login, work_book, input_paths)
from hpro_automation.Config import read_excel
import requests
import json
import xlrd


class AssessmentSlots(login.CommonLogin, work_book.WorkBook):

    def __init__(self):
        # --------------------------------- Inheritance Class Instance -------------------------------------------------
        super(AssessmentSlots, self).__init__()

        # --------------------------------- Excel Data initialize variables --------------------------------------------
        self.xl_dict = {}
        self.dict_total = []

    def excel_data(self):

        # ------------------------------- Excel Data Read --------------------------------------------------------------
        try:
            excel = read_excel.ExcelRead()
            index = 0
            excel.excel_read(input_paths.inputpaths['assessment_slot_input_sheet'], index)
            self.xl_dict = excel.details
            self.dict_total = excel.details

            print("Excel Data:: ", self.xl_dict)
        except IOError:
            print("File not found or path is incorrect")

    def choose_assessment_slot(self, loop):
        self.slot_captcha_login_token('assessment')
        self.verify_hash(self.xl_dict[loop]['verify_hash'])
        self.lambda_function('assessment_slot_select')

        # ----------------------------------- API request --------------------------------------------------------------
        print("------------- Choose assessment slot API Call -----------------")
        request = json.loads(self.xl_dict[loop]['chooseSlot'])
        choose_slot_api = requests.post(self.webapi, headers=self.lambda_headers,
                                        data=json.dumps(request), verify=False)
        response = json.loads(choose_slot_api.content)
        print(response)
        print(response.get('data'))

    def update_assessment_slot(self, loop):
        self.lambda_function('assessment_slot_update')

        # ----------------------------------- API request --------------------------------------------------------------
        print("------------- Update assessment slot API Call -----------------")
        request = json.loads(self.xl_dict[loop]['updateSlot'])
        update_slot_api = requests.post(self.webapi, headers=self.lambda_headers,
                                        data=json.dumps(request), verify=False)
        response = json.loads(update_slot_api.content)
        print(response)
        print(response.get('data'))

    def unassign_slot(self, loop):
        self.common_login('slot')
        self.lambda_function('assessment_unassign_slot')

        # ----------------------------------- API request --------------------------------------------------------------
        print("------------- Choose Unassign slot API Call -----------------")
        request = json.loads(self.xl_dict[loop]['UnassignSlot'])

        unassign_slot_api = requests.post(self.webapi, headers=self.headers,
                                          data=json.dumps(request), verify=False)
        response = json.loads(unassign_slot_api.content)
        data = response.get('data')
        print(response)
        print(data.get('dissociateSlotDetails'))


Object = AssessmentSlots()
Object.excel_data()

Total_count = len(Object.dict_total)
print("Number of Rows::", Total_count)
for looping in range(0, Total_count):
    print("Iteration Count is ::", looping)
    Object.choose_assessment_slot(looping)
    Object.update_assessment_slot(looping)
    Object.unassign_slot(looping)

# ----------------- Make Dictionaries clear for each loop ------------------------------------------------------
