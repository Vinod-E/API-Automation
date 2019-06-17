import requests
import json
import xlrd
from hpro_automation import (login, api, input_paths, output_paths, db_login)


class DeleteCommunication(login.CRPOLogin, db_login.DBConnection):

    def __init__(self):
        super(DeleteCommunication, self).__init__()

        self.xl_attachment_id = []
        self.xl_CommunicationPurpose = []
        self.xl_applicant_id = []

        self.db_entity_communication_history = []
        self.headers = {}

    def applicant_candidate_excel(self):
        workbook = xlrd.open_workbook(input_paths.inputpaths['Communication_History_Input_sheet'])
        sheet1 = workbook.sheet_by_index(0)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)

            if rows[1]:
                self.xl_applicant_id.append(int(rows[1]))
            else:
                self.xl_applicant_id.append(None)

    def attachment_id_excel(self):
        workbook1 = xlrd.open_workbook(output_paths.outputpaths['Communication_output_sheet'])
        sheet1_1 = workbook1.sheet_by_index(0)
        for j in range(2, sheet1_1.nrows):
            number = j  # Counting number of rows
            rows = sheet1_1.row_values(number)

            if rows[33]:
                self.xl_attachment_id.append(int(rows[33]))

    def delete_attachment(self):

        # ---------------- Passing headers based on API supports to lambda or not --------------------
        if self.calling_lambda == 'On':
            if api.lambda_apis.get('delete_Attachment') is not None \
                    and api.web_api['delete_Attachment'] in api.lambda_apis['delete_Attachment']:
                self.headers = self.lambda_headers
            else:
                self.headers = self.Non_lambda_headers
        elif self.calling_lambda == 'Off':
            self.headers = self.lambda_headers
        else:
            self.headers = self.lambda_headers

        # ---------------- Updating headers with app name -----------------
        self.headers['APP-NAME'] = 'crpo'

        request = {"AttachmentIds": self.xl_attachment_id}
        attachment_api = requests.post(api.web_api['delete_Attachment'], headers=self.headers,
                                       data=json.dumps(request, default=str), verify=False)
        print(attachment_api.headers)
        attachment_api_dict = json.loads(attachment_api.content)
        print(attachment_api_dict)

    def update_communication_history(self):
        app_can_ids = tuple(self.xl_applicant_id)
        self.db_connection()

        query0 = "DELETE FROM entity_communication_history WHERE entitycommunication_id in" \
                 " (SELECT id from entity_communications " \
                 "where entity_id in {} and entity_type = 43);".format(app_can_ids)
        print(query0)
        self.cursor.execute(query0)

        query1 = "DELETE FROM entity_communications WHERE entity_id in {};".format(app_can_ids)
        print(query1)
        self.cursor.execute(query1)


Object = DeleteCommunication()
Object.applicant_candidate_excel()
Object.attachment_id_excel()

Object.delete_attachment()
Object.update_communication_history()
Object.connection.commit()
Object.connection.close()

Object.headers = {}

# total_count = len(Object.xl_applicant_id)
# print("No.of Rows ::", total_count)
#
# for looping in range(0, total_count):
#     print("Iteration Count is ::", looping)
#     Object.delete_attachment(looping)
#     Object.update_communication_history(looping)
#
# Object.connection.commit()
# Object.connection.close()
