from hpro_automation import (login, input_paths, api, output_paths, work_book)
import requests
import json
import xlrd
import datetime


class ActivityCallBack(login.CRPOLogin, work_book.WorkBook):

    def __init__(self):

        self.start_time = str(datetime.datetime.now())
        super(ActivityCallBack, self).__init__()
        # -----------------------------------------
        # Activity Call Back data set initialsation
        # -----------------------------------------
        self.xl_candidateId = []
        self.xl_applicantId = []
        self.xl_eventId = []
        self.xl_jobId = []
        self.xl_statusId = []
        self.xl_mjrId = []
        self.xl_comment = []
        self.xl_Activity_01 = []
        self.xl_A1_task_01 = []
        self.xl_A1_t1_status = []
        self.xl_A1_controlId = []
        self.xl_A1_controlvalue = []
        self.xl_A1_controltype = []
        self.xl_Activity_02 = []
        self.xl_A2_task_01 = []
        self.xl_A2_t1_status = []
        self.xl_A2_task_02 = []
        self.xl_A2_t2_status = []
        self.xl_A2_task_03 = []
        self.xl_A2_t3_status = []
        self.xl_Activity_03 = []
        self.xl_A3_task_01 = []
        self.xl_A3_t1_status = []
        self.xl_A3_task_02 = []
        self.xl_A3_t2_status = []
        self.xl_common_controlId = []
        self.xl_common_controlvalue = []
        self.xl_common_controltype = []
        # ------------------
        # Expected Results
        # ------------------
        self.xl_Expected_Stage = []
        self.xl_Expected_Status = []
        self.xl_Expected_A1_T1 = []
        self.xl_Expected_A1_T1_Status = []
        self.xl_Expected_A2_T1 = []
        self.xl_Expected_A2_T1_Status = []
        self.xl_Expected_A2_T2 = []
        self.xl_Expected_A2_T2_Status = []
        self.xl_Expected_A2_T3 = []
        self.xl_Expected_A2_T3_Status = []
        self.xl_Expected_A3_T1 = []
        self.xl_Expected_A3_T1_Status = []
        self.xl_Expected_A3_T2 = []
        self.xl_Expected_A3_T2_Status = []
        self.xl_Expected_Task_ids = []

        # --------------
        # Others values
        # --------------
        self.other_change_applicant_statusid = []
        self.other_change_applicant_comment = []
        self.other_activity_01 = []
        self.other_activity_02 = []
        self.other_activity_03 = []
        self.other_task_01 = []
        self.other_task_02 = []
        self.other_task_03 = []
        self.other_task_04 = []
        self.other_task_05 = []
        self.other_task_06 = []
        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 14)))
        self.Actual_Success_case = []

        self.Actual_Task_ids = ""
        self.success_case_01 = {}
        self.headers = {}

        # -----------
        # Dictionary
        # -----------
        self.success = {}
        self.suc = self.success = {}
        self.chng_app_status = {}
        self.c_a_s = self.chng_app_status
        self.applicant_dict_details = {}
        self.a_d_d = self.applicant_dict_details
        self.applicant_mjr_dict = {}
        self.app_d1 = self.applicant_mjr_dict

        self.ticket_number_A1_t1 = {}
        self.ticket1 = self.ticket_number_A1_t1

        self.ticket_number_A2_t1 = {}
        self.ticket2 = self.ticket_number_A2_t1
        self.ticket_number_A2_t2 = {}
        self.ticket3 = self.ticket_number_A2_t2
        self.ticket_number_A2_t3 = {}
        self.ticket4 = self.ticket_number_A2_t3

        self.ticket_number_A3_t1 = {}
        self.ticket5 = self.ticket_number_A3_t1
        self.ticket_number_A3_t2 = {}
        self.ticket6 = self.ticket_number_A3_t2

        self.A1 = {}
        self.a1 = self.A1 = {}
        self.A1_t1 = {}
        self.a1t1 = self.A1_t1 = {}
        self.A2 = {}
        self.a2 = self.A2 = {}
        self.A2_t1 = {}
        self.a2t1 = self.A2_t1 = {}
        self.A2_t2 = {}
        self.a2t2 = self.A2_t2 = {}
        self.A2_t3 = {}
        self.a2t3 = self.A2_t3 = {}
        self.A3 = {}
        self.a3 = self.A3 = {}
        self.A3_t1 = {}
        self.a3t1 = self.A3_t1 = {}
        self.A3_t2 = {}
        self.a3t2 = self.A3_t2 = {}

        self.A1_t1_status = {}
        self.a1_t1_s = self.A1_t1_status = {}
        self.A2_t1_status = {}
        self.a2_t1_s = self.A2_t1_status = {}
        self.A2_t2_status = {}
        self.a2_t2_s = self.A2_t2_status = {}
        self.A2_t3_status = {}
        self.a2_t3_s = self.A2_t3_status = {}
        self.A3_t1_status = {}
        self.a3_t1_s = self.A3_t1_status = {}
        self.A3_t2_status = {}
        self.a3_t2_s = self.A3_t2_status = {}

    def excel_headers(self):

        # --------------------------------- Excel Headers and Cell color, styles ---------------------------------------
        self.main_headers = ['Comparision', 'Expected_Status', 'candidateId', 'ApplicantId', 'Current_stage',
                             'Current_Status', 'A1_T1', 'A1_T1_Status', 'A2_T1', 'A2_T1_Status', 'A2_T2',
                             'A2_T2_Status', 'A2_T3(OT)', 'A2_T3_Status', 'A3_T1', 'A3_T1_Status', 'A3_T2',
                             'A3_T2_Status', 'Task_IDs']
        self.headers_with_style2 = ['Comparision', 'Expected_Status', 'candidateId', 'ApplicantId', 'Current_stage',
                                    'Current_Status']
        self.headers_with_style10 = ['Task_IDs']
        self.headers_with_style21 = ['A2_T1', 'A2_T1_Status', 'A2_T2', 'A2_T2_Status', 'A2_T3(OT)', 'A2_T3_Status']
        self.headers_with_style22 = ['A1_T1', 'A1_T1_Status']
        self.file_headers_col_row()

    def excel_data(self):

        # ----------------
        # Excel Data Read
        # ----------------
        workbook = xlrd.open_workbook(input_paths.inputpaths['Activity_C_back_Input_sheet'])
        sheet1 = workbook.sheet_by_index(0)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)

            if rows[0] is not None and rows[0] == '':
                self.xl_candidateId.append(None)
            else:
                self.xl_candidateId.append(int(rows[0]))

            if rows[1] is not None and rows[1] == '':
                self.xl_applicantId.append(None)
            else:
                self.xl_applicantId.append(int(rows[1]))

            if rows[2] is not None and rows[2] == '':
                self.xl_eventId.append(None)
            else:
                self.xl_eventId.append(int(rows[2]))

            if rows[3] is not None and rows[3] == '':
                self.xl_jobId.append(None)
            else:
                self.xl_jobId.append(int(rows[3]))

            if rows[4] is not None and rows[4] == '':
                self.xl_statusId.append(None)
            else:
                self.xl_statusId.append(int(rows[4]))

            if rows[5] is not None and rows[5] == '':
                self.xl_mjrId.append(None)
            else:
                self.xl_mjrId.append(int(rows[5]))

            if rows[6] is not None and rows[6] == '':
                self.xl_comment.append(None)
            else:
                self.xl_comment.append(rows[6])

            if rows[7] is not None and rows[7] == '':
                self.xl_Activity_01.append(None)
            else:
                self.xl_Activity_01.append(int(rows[7]))

            if rows[8] is not None and rows[8] == '':
                self.xl_A1_task_01.append(None)
            else:
                self.xl_A1_task_01.append(int(rows[8]))

            if rows[9] is not None and rows[9] == '':
                self.xl_A1_t1_status.append(None)
            else:
                self.xl_A1_t1_status.append(rows[9])

            if rows[10] is not None and rows[10] == '':
                self.xl_A1_controlvalue.append(None)
            else:
                self.xl_A1_controlvalue.append(rows[10])

            if rows[11] is not None and rows[11] == '':
                self.xl_A1_controlId.append(None)
            else:
                self.xl_A1_controlId.append(int(rows[11]))

            if rows[12] is not None and rows[12] == '':
                self.xl_A1_controltype.append(None)
            else:
                self.xl_A1_controltype.append(int(rows[12]))

            if rows[13] is not None and rows[13] == '':
                self.xl_Activity_02.append(None)
            else:
                self.xl_Activity_02.append(int(rows[13]))

            if rows[14] is not None and rows[14] == '':
                self.xl_A2_task_01.append(None)
            else:
                self.xl_A2_task_01.append(int(rows[14]))

            if rows[15] is not None and rows[15] == '':
                self.xl_A2_t1_status.append(None)
            else:
                self.xl_A2_t1_status.append(int(rows[15]))

            if rows[16] is not None and rows[16] == '':
                self.xl_A2_task_02.append(None)
            else:
                self.xl_A2_task_02.append(int(rows[16]))

            if rows[17] is not None and rows[17] == '':
                self.xl_A2_t2_status.append(None)
            else:
                self.xl_A2_t2_status.append(int(rows[17]))

            if rows[18] is not None and rows[18] == '':
                self.xl_A2_task_03.append(None)
            else:
                self.xl_A2_task_03.append(int(rows[18]))

            if rows[19] is not None and rows[19] == '':
                self.xl_A2_t3_status.append(None)
            else:
                self.xl_A2_t3_status.append(int(rows[19]))

            if rows[20] is not None and rows[20] == '':
                self.xl_Activity_03.append(None)
            else:
                self.xl_Activity_03.append(int(rows[20]))

            if rows[21] is not None and rows[21] == '':
                self.xl_A3_task_01.append(None)
            else:
                self.xl_A3_task_01.append(int(rows[21]))

            if rows[22] is not None and rows[22] == '':
                self.xl_A3_t1_status.append(None)
            else:
                self.xl_A3_t1_status.append(int(rows[22]))

            if rows[23] is not None and rows[23] == '':
                self.xl_A3_task_02.append(None)
            else:
                self.xl_A3_task_02.append(int(rows[23]))

            if rows[24] is not None and rows[24] == '':
                self.xl_A3_t2_status.append(None)
            else:
                self.xl_A3_t2_status.append(int(rows[24]))

            if rows[25] is not None and rows[25] == '':
                self.xl_common_controlvalue.append(None)
            else:
                self.xl_common_controlvalue.append(rows[25])

            if rows[26] is not None and rows[26] == '':
                self.xl_common_controlId.append(None)
            else:
                self.xl_common_controlId.append(int(rows[26]))

            if rows[27] is not None and rows[27] == '':
                self.xl_common_controltype.append(None)
            else:
                self.xl_common_controltype.append(int(rows[27]))

            if rows[29] is not None and rows[29] == '':
                self.xl_Expected_Stage.append(None)
            else:
                self.xl_Expected_Stage.append(rows[29])

            if rows[30] is not None and rows[30] == '':
                self.xl_Expected_Status.append(None)
            else:
                self.xl_Expected_Status.append(rows[30])

            if rows[31] is not None and rows[31] == '':
                self.xl_Expected_A1_T1.append(None)
            else:
                self.xl_Expected_A1_T1.append(int(rows[31]))

            if rows[32] is not None and rows[32] == '':
                self.xl_Expected_A1_T1_Status.append(None)
            else:
                self.xl_Expected_A1_T1_Status.append(rows[32])

            if rows[33] is not None and rows[33] == '':
                self.xl_Expected_A2_T1.append(None)
            else:
                self.xl_Expected_A2_T1.append(int(rows[33]))

            if rows[34] is not None and rows[34] == '':
                self.xl_Expected_A2_T1_Status.append(None)
            else:
                self.xl_Expected_A2_T1_Status.append(rows[34])

            if rows[35] is not None and rows[35] == '':
                self.xl_Expected_A2_T2.append(None)
            else:
                self.xl_Expected_A2_T2.append(int(rows[35]))

            if rows[36] is not None and rows[36] == '':
                self.xl_Expected_A2_T2_Status.append(None)
            else:
                self.xl_Expected_A2_T2_Status.append(rows[36])

            if rows[37] is not None and rows[37] == '':
                self.xl_Expected_A2_T3.append(None)
            else:
                self.xl_Expected_A2_T3.append(int(rows[37]))

            if rows[38] is not None and rows[38] == '':
                self.xl_Expected_A2_T3_Status.append(None)
            else:
                self.xl_Expected_A2_T3_Status.append(rows[38])

            if rows[39] is not None and rows[39] == '':
                self.xl_Expected_A3_T1.append(None)
            else:
                self.xl_Expected_A3_T1.append(int(rows[39]))

            if rows[40] is not None and rows[40] == '':
                self.xl_Expected_A3_T1_Status.append(None)
            else:
                self.xl_Expected_A3_T1_Status.append(rows[40])

            if rows[41] is not None and rows[41] == '':
                self.xl_Expected_A3_T2.append(None)
            else:
                self.xl_Expected_A3_T2.append(int(rows[41]))

            if rows[42] is not None and rows[42] == '':
                self.xl_Expected_A3_T2_Status.append(None)
            else:
                self.xl_Expected_A3_T2_Status.append(rows[42])

            if rows[43] is not None and rows[43] == '':
                self.xl_Expected_Task_ids.append(None)
            else:
                self.xl_Expected_Task_ids.append(rows[43])

        workbook = xlrd.open_workbook(input_paths.inputpaths['Activity_C_back_Input_sheet'])
        sheet1 = workbook.sheet_by_index(1)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)

            if rows[0]:
                self.other_change_applicant_statusid.append(int(rows[0]))
            if rows[1]:
                self.other_change_applicant_comment.append(rows[1])
            if rows[2]:
                self.other_activity_01.append(int(rows[2]))
            if rows[3]:
                self.other_activity_02.append(int(rows[3]))
            if rows[4]:
                self.other_activity_03.append(int(rows[4]))
            if rows[5]:
                self.other_task_01.append(int(rows[5]))
            if rows[6]:
                self.other_task_02.append(int(rows[6]))
            if rows[7]:
                self.other_task_03.append(int(rows[7]))
            if rows[8]:
                self.other_task_04.append(int(rows[8]))
            if rows[9]:
                self.other_task_05.append(int(rows[9]))
            if rows[10]:
                self.other_task_06.append(int(rows[10]))

    def change_applicant_status(self, loop):

        # ---------------- Passing headers based on API supports to lambda or not --------------------
        if self.calling_lambda == 'On':
            if api.lambda_apis.get('ChangeApplicant_Status') is not None \
                    and api.web_api['ChangeApplicant_Status'] in api.lambda_apis.get('ChangeApplicant_Status'):
                self.headers = self.lambda_headers
            else:
                self.headers = self.Non_lambda_headers
        elif self.calling_lambda == 'Off':
            self.headers = self.lambda_headers
        else:
            self.headers = self.lambda_headers

        # ---------------- Updating headers with app name -----------------
        self.headers['APP-NAME'] = 'crpo'

        request = {"ApplicantIds": [self.xl_applicantId[loop]],
                   "EventId": self.xl_eventId[loop],
                   "JobRoleId": self.xl_jobId[loop],
                   "ToStatusId": self.xl_statusId[loop],
                   "Sync": "True",
                   "Comments": self.xl_comment[loop],
                   "InitiateStaffing": False,
                   "MjrId": self.xl_mjrId[loop]
                   }
        change_status_api = requests.post(api.web_api['ChangeApplicant_Status'], headers=self.headers,
                                          data=json.dumps(request, default=str), verify=False)
        response = json.loads(change_status_api.content)
        self.chng_app_status = response.get('status')

        data = response['data']
        self.success = data.get('success')

    def applicant_info(self, loop):

        # ---------------- Passing headers based on API supports to lambda or not --------------------
        if self.calling_lambda == 'On':
            if api.lambda_apis.get('getApplicantsInfo') is not None \
                    and api.web_api['getApplicantsInfo'] in api.lambda_apis['getApplicantsInfo']:
                self.headers = self.lambda_headers
            else:
                self.headers = self.Non_lambda_headers
        elif self.calling_lambda == 'Off':
            self.headers = self.lambda_headers
        else:
            self.headers = self.lambda_headers

        # ---------------- Updating headers with app name -----------------
        self.headers['APP-NAME'] = 'crpo'

        if self.xl_Activity_01:
            request = {
                "CandidateIds": [self.xl_candidateId[loop]]
            }
            applicant_info_api = requests.post(api.web_api['getApplicantsInfo'], headers=self.headers,
                                               data=json.dumps(request, default=str), verify=False)
            response = json.loads(applicant_info_api.content)
            data = response.get('data')

            for i in data:
                self.applicant_dict_details = i
                applicant_dict = self.applicant_dict_details['ApplicantDetails']
                for j in applicant_dict:
                    if self.xl_mjrId[loop] == j.get('MjrId'):
                        self.applicant_mjr_dict = j

    def get_ticket_number(self, loop):

        # ---------------- Passing headers based on API supports to lambda or not --------------------
        if self.calling_lambda == 'On':
            if api.lambda_apis.get('gettaskbycandidate') is not None \
                    and api.web_api['gettaskbycandidate'] in api.lambda_apis['gettaskbycandidate']:
                self.headers = self.lambda_headers
            else:
                self.headers = self.Non_lambda_headers
        elif self.calling_lambda == 'Off':
            self.headers = self.lambda_headers
        else:
            self.headers = self.lambda_headers

        # ---------------- Updating headers with app name -----------------
        self.headers['APP-NAME'] = 'crpo'

        request = {
            "CandidateId": self.xl_candidateId[loop]
        }
        activity_details_api = requests.post(api.web_api['gettaskbycandidate'], headers=self.headers,
                                             data=json.dumps(request, default=str), verify=False)
        response = json.loads(activity_details_api.content)

        if response['status'] == 'OK':
            task_collection = response.get('CandidateTaskCollection')
            for i in task_collection:

                if self.xl_Activity_01[loop] == i['ActivityId']:
                    if self.xl_A1_task_01[loop] == i['TaskId']:
                        self.ticket_number_A1_t1 = i['TicketNumber']

                if self.xl_Activity_02[loop] == i['ActivityId']:
                    if self.xl_A2_task_01[loop] == i['TaskId']:
                        self.ticket_number_A2_t1 = i['TicketNumber']
                    if self.xl_A2_task_02[loop] == i['TaskId']:
                        self.ticket_number_A2_t2 = i['TicketNumber']
                    if self.xl_A2_task_03[loop] == i['TaskId']:
                        self.ticket_number_A2_t3 = i['TicketNumber']

                if self.xl_Activity_03[loop] == i['ActivityId']:
                    if self.xl_A3_task_01[loop] == i['TaskId']:
                        self.ticket_number_A3_t1 = i['TicketNumber']
                    if self.xl_A3_task_02[loop] == i['TaskId']:
                        self.ticket_number_A3_t2 = i['TicketNumber']

    def submit_form_a1(self, loop):

        # ---------------- Passing headers based on API supports to lambda or not --------------------
        if self.calling_lambda == 'On':
            if api.lambda_apis.get('submitform') is not None \
                    and api.web_api['submitform'] in api.lambda_apis['submitform']:
                self.headers = self.lambda_headers
            else:
                self.headers = self.Non_lambda_headers
        elif self.calling_lambda == 'Off':
            self.headers = self.lambda_headers
        else:
            self.headers = self.lambda_headers

        # ---------------- Updating headers with app name -----------------
        self.headers['APP-NAME'] = 'staffing'

        if self.ticket_number_A1_t1:
            request = {
                "FormControlValues": [
                    {
                        "ControlId": self.xl_A1_controlId[loop],
                        "FormControlType": self.xl_A1_controltype[loop],
                        "DisclaimerCheck": False,
                        "ControlValue": self.xl_A1_controlvalue[loop]}
                ],
                "TicketId": self.ticket_number_A1_t1,
                "IsCoordinatorSubmittingBehalfOfCandidate": True
            }
            submit_form_api = requests.post(api.web_api['submitform'], headers=self.headers,
                                            data=json.dumps(request, default=str), verify=False)
            response1 = json.loads(submit_form_api.content)
            print(response1)

    def submit_form_a2(self, loop):

        # ---------------- Passing headers based on API supports to lambda or not --------------------
        if self.calling_lambda == 'On':
            if api.lambda_apis.get('submitform') is not None \
                    and api.web_api['submitform'] in api.lambda_apis['submitform']:
                self.headers = self.lambda_headers
            else:
                self.headers = self.Non_lambda_headers
        elif self.calling_lambda == 'Off':
            self.headers = self.lambda_headers
        else:
            self.headers = self.lambda_headers

        # ---------------- Updating headers with app name -----------------
        self.headers['APP-NAME'] = 'staffing'

        if self.ticket_number_A2_t1:
            request = {
                "FormControlValues": [
                    {
                        "ControlId": self.xl_common_controlId[loop],
                        "FormControlType": self.xl_common_controltype[loop],
                        "DisclaimerCheck": False,
                        "ControlValue": self.xl_common_controlvalue[loop]}
                ],
                "TicketId": self.ticket_number_A2_t1,
                "IsCoordinatorSubmittingBehalfOfCandidate": True
            }
            submit_form_api = requests.post(api.web_api['submitform'], headers=self.headers,
                                            data=json.dumps(request, default=str), verify=False)
            response1 = json.loads(submit_form_api.content)
            print(response1)

        if self.ticket_number_A2_t2:
            request = {
                "FormControlValues": [
                    {
                        "ControlId": self.xl_common_controlId[loop],
                        "FormControlType": self.xl_common_controltype[loop],
                        "DisclaimerCheck": False,
                        "ControlValue": self.xl_common_controlvalue[loop]}
                ],
                "TicketId": self.ticket_number_A2_t2,
                "IsCoordinatorSubmittingBehalfOfCandidate": True
            }
            submit_form_api = requests.post(api.web_api['submitform'], headers=self.headers,
                                            data=json.dumps(request, default=str), verify=False)
            response2 = json.loads(submit_form_api.content)
            print(response2)

        if self.ticket_number_A2_t3:
            request = {
                "FormControlValues": [
                    {
                        "ControlId": self.xl_common_controlId[loop],
                        "FormControlType": self.xl_common_controltype[loop],
                        "DisclaimerCheck": False,
                        "ControlValue": self.xl_common_controlvalue[loop]}
                ],
                "TicketId": self.ticket_number_A2_t3,
                "IsCoordinatorSubmittingBehalfOfCandidate": True
            }
            submit_form_api = requests.post(api.web_api['submitform'], headers=self.headers,
                                            data=json.dumps(request, default=str), verify=False)
            response3 = json.loads(submit_form_api.content)
            print(response3)

        # ---------------------------
        # Approve task by AEE owner
        # ---------------------------

        # ---------------- Passing headers based on API supports to lambda or not --------------------
        if self.calling_lambda == 'On':
            if api.lambda_apis.get('Approve_task') is not None \
                    and api.web_api['Approve_task'] in api.lambda_apis['Approve_task']:
                self.headers = self.lambda_headers
            else:
                self.headers = self.Non_lambda_headers
        elif self.calling_lambda == 'Off':
            self.headers = self.lambda_headers
        else:
            self.headers = self.lambda_headers

        # ---------------- Updating headers with app name -----------------
        self.headers['APP-NAME'] = 'crpo'

        req1 = {"AssignedUserTaskIds": [self.ticket_number_A2_t1],
                "TaskStatus": self.xl_A2_t1_status[loop],
                "Comments": self.xl_common_controlvalue[loop]}
        approve_api = requests.post(api.web_api['Approve_task'], headers=self.headers,
                                    data=json.dumps(req1, default=str), verify=False)
        res1 = json.loads(approve_api.content)
        print(res1)

        req2 = {"AssignedUserTaskIds": [self.ticket_number_A2_t2],
                "TaskStatus": self.xl_A2_t2_status[loop],
                "Comments": self.xl_common_controlvalue[loop]}
        approve_api = requests.post(api.web_api['Approve_task'], headers=self.headers,
                                    data=json.dumps(req2, default=str), verify=False)
        res2 = json.loads(approve_api.content)
        print(res2)

        req3 = {"AssignedUserTaskIds": [self.ticket_number_A2_t3],
                "TaskStatus": self.xl_A2_t3_status[loop],
                "Comments": self.xl_common_controlvalue[loop]}
        approve_api = requests.post(api.web_api['Approve_task'], headers=self.headers,
                                    data=json.dumps(req3, default=str), verify=False)
        res3 = json.loads(approve_api.content)
        print(res3)

    def submit_form_a3(self, loop):

        # ---------------- Passing headers based on API supports to lambda or not --------------------
        if self.calling_lambda == 'On':
            if api.lambda_apis.get('submitform') is not None \
                    and api.web_api['submitform'] in api.lambda_apis['submitform']:
                self.headers = self.lambda_headers
            else:
                self.headers = self.Non_lambda_headers
        elif self.calling_lambda == 'Off':
            self.headers = self.lambda_headers
        else:
            self.headers = self.lambda_headers

        # ---------------- Updating headers with app name -----------------
        self.headers['APP-NAME'] = 'staffing'

        if self.ticket_number_A3_t1:
            request = {
                "FormControlValues": [
                    {
                        "ControlId": self.xl_common_controlId[loop],
                        "FormControlType": self.xl_common_controltype[loop],
                        "DisclaimerCheck": False,
                        "ControlValue": self.xl_common_controlvalue[loop]}
                ],
                "TicketId": self.ticket_number_A3_t1,
                "IsCoordinatorSubmittingBehalfOfCandidate": True
            }
            submit_form_api = requests.post(api.web_api['submitform'], headers=self.headers,
                                            data=json.dumps(request, default=str), verify=False)
            response1 = json.loads(submit_form_api.content)
            print(response1)

        if self.ticket_number_A3_t2:
            request = {
                "FormControlValues": [
                    {
                        "ControlId": self.xl_common_controlId[loop],
                        "FormControlType": self.xl_common_controltype[loop],
                        "DisclaimerCheck": False,
                        "ControlValue": self.xl_common_controlvalue[loop]}
                ],
                "TicketId": self.ticket_number_A3_t2,
                "IsCoordinatorSubmittingBehalfOfCandidate": True
            }
            submit_form_api = requests.post(api.web_api['submitform'], headers=self.headers,
                                            data=json.dumps(request, default=str), verify=False)
            response2 = json.loads(submit_form_api.content)
            print(response2)

        # ---------------------------
        # Approve task by AEE owner
        # ---------------------------

        # ---------------- Passing headers based on API supports to lambda or not --------------------
        if self.calling_lambda == 'On':
            if api.lambda_apis.get('Approve_task') is not None \
                    and api.web_api['Approve_task'] in api.lambda_apis['Approve_task']:
                self.headers = self.lambda_headers
            else:
                self.headers = self.Non_lambda_headers
        elif self.calling_lambda == 'Off':
            self.headers = self.lambda_headers
        else:
            self.headers = self.lambda_headers

        # ---------------- Updating headers with app name -----------------
        self.headers['APP-NAME'] = 'crpo'

        req = {"AssignedUserTaskIds": [self.ticket_number_A3_t1],
               "TaskStatus": self.xl_A3_t1_status[loop],
               "Comments": self.xl_common_controlvalue[loop]}
        approve_api = requests.post(api.web_api['Approve_task'], headers=self.headers,
                                    data=json.dumps(req, default=str), verify=False)
        res = json.loads(approve_api.content)
        print(res)
        req1 = {"AssignedUserTaskIds": [self.ticket_number_A3_t2],
                "TaskStatus": self.xl_A3_t2_status[loop],
                "Comments": self.xl_common_controlvalue[loop]}
        approve_api = requests.post(api.web_api['Approve_task'], headers=self.headers,
                                    data=json.dumps(req1, default=str), verify=False)
        res1 = json.loads(approve_api.content)
        print(res1)

    def get_activity_task_details(self, loop):

        # ---------------- Passing headers based on API supports to lambda or not --------------------
        if self.calling_lambda == 'On':
            if api.lambda_apis.get('gettaskbycandidate') is not None \
                    and api.web_api['gettaskbycandidate'] in api.lambda_apis['gettaskbycandidate']:
                self.headers = self.lambda_headers
            else:
                self.headers = self.Non_lambda_headers
        elif self.calling_lambda == 'Off':
            self.headers = self.lambda_headers
        else:
            self.headers = self.lambda_headers

        # ---------------- Updating headers with app name -----------------
        self.headers['APP-NAME'] = 'crpo'

        request = {
            "CandidateId": self.xl_candidateId[loop]
        }
        activity_details_api = requests.post(api.web_api['gettaskbycandidate'], headers=self.headers,
                                             data=json.dumps(request, default=str), verify=False)
        response = json.loads(activity_details_api.content)

        if response['status'] == 'OK':
            task_collection = response.get('CandidateTaskCollection')

            if task_collection:
                for i in task_collection:
                    # self.get_activity_task_info = i
                    # print self.get_activity_task_info

                    if i['ActivityId'] == self.other_activity_01[0]:
                        self.A1 = i.get('ActivityId')
                        if i['TaskId'] == self.other_task_01[0]:
                            self.A1_t1 = i.get('TaskId')
                            self.A1_t1_status = i['StatusText']

                    if i['ActivityId'] == self.other_activity_02[0]:
                        self.A2 = i.get('ActivityId')
                        if i['TaskId'] == self.other_task_02[0]:
                            self.A2_t1 = i.get('TaskId')
                            self.A2_t1_status = i.get('StatusText')
                        if i['TaskId'] == self.other_task_03[0]:
                            self.A2_t2 = i.get('TaskId')
                            self.A2_t2_status = i.get('StatusText')
                        if i['TaskId'] == self.other_task_04[0]:
                            self.A2_t3 = i.get('TaskId')
                            self.A2_t3_status = i.get('StatusText')

                    if i['ActivityId'] == self.other_activity_03[0]:
                        self.A3 = i.get('ActivityId')
                        if i['TaskId'] == self.other_task_05[0]:
                            self.A3_t1 = i.get('TaskId')
                            self.A3_t1_status = i.get('StatusText')
                        if i['TaskId'] == self.other_task_06[0]:
                            self.A3_t2 = i.get('TaskId')
                            self.A3_t2_status = i.get('StatusText')

                    if i['CandidateId'] == self.xl_candidateId[loop]:
                        actual_task_ids = i['TaskId']
                        self.Actual_Task_ids += ",%s" % actual_task_ids
            self.Actual_Task_ids = self.Actual_Task_ids.lstrip(',')

    def reset_applicant_status(self, loop):

        # ---------------- Passing headers based on API supports to lambda or not --------------------
        if self.calling_lambda == 'On':
            if api.lambda_apis.get('ChangeApplicant_Status') is not None \
                    and api.web_api['ChangeApplicant_Status'] in api.lambda_apis.get('ChangeApplicant_Status'):
                self.headers = self.lambda_headers
            else:
                self.headers = self.Non_lambda_headers
        elif self.calling_lambda == 'Off':
            self.headers = self.lambda_headers
        else:
            self.headers = self.lambda_headers

        # ---------------- Updating headers with app name -----------------
        self.headers['APP-NAME'] = 'crpo'

        request = {"ApplicantIds": [self.xl_applicantId[loop]],
                   "EventId": self.xl_eventId[loop],
                   "JobRoleId": self.xl_jobId[loop],
                   "ToStatusId": self.other_change_applicant_statusid[0],
                   "Sync": "True",
                   "Comments": self.other_change_applicant_comment[0],
                   "InitiateStaffing": False,
                   "MjrId": self.xl_mjrId[loop]
                   }
        q = requests.post(api.web_api['ChangeApplicant_Status'], headers=self.headers,
                          data=json.dumps(request, default=str), verify=False)
        qw = json.loads(q.content)
        print(qw)

    def output_excel(self, loop):

        # ------------------
        # Writing Input Data
        # ------------------
        self.ws.write(self.rowsize, self.col, 'Input', self.style4)
        self.ws.write(self.rowsize, 2, self.xl_candidateId[loop])
        self.ws.write(self.rowsize, 3, self.xl_applicantId[loop])
        self.ws.write(self.rowsize, 4, self.xl_Expected_Stage[loop])
        self.ws.write(self.rowsize, 5, self.xl_Expected_Status[loop])
        self.ws.write(self.rowsize, 6, self.xl_Expected_A1_T1[loop] if self.xl_Expected_A1_T1[loop] else 'NA')
        self.ws.write(self.rowsize, 7,
                      self.xl_Expected_A1_T1_Status[loop] if self.xl_Expected_A1_T1_Status[loop] else 'NA')
        self.ws.write(self.rowsize, 8, self.xl_Expected_A2_T1[loop] if self.xl_Expected_A2_T1[loop] else 'NA')
        self.ws.write(self.rowsize, 9,
                      self.xl_Expected_A2_T1_Status[loop] if self.xl_Expected_A2_T1_Status[loop] else 'NA')
        self.ws.write(self.rowsize, 10, self.xl_Expected_A2_T2[loop] if self.xl_Expected_A2_T2[loop] else 'NA')
        self.ws.write(self.rowsize, 11,
                      self.xl_Expected_A2_T2_Status[loop] if self.xl_Expected_A2_T2_Status[loop] else 'NA')
        self.ws.write(self.rowsize, 12, self.xl_Expected_A2_T3[loop] if self.xl_Expected_A2_T3[loop] else 'NA')
        self.ws.write(self.rowsize, 13,
                      self.xl_Expected_A2_T3_Status[loop] if self.xl_Expected_A2_T3_Status[loop] else 'NA')
        self.ws.write(self.rowsize, 14, self.xl_Expected_A3_T1[loop] if self.xl_Expected_A3_T1[loop] else 'NA')
        self.ws.write(self.rowsize, 15,
                      self.xl_Expected_A3_T1_Status[loop] if self.xl_Expected_A3_T1_Status[loop] else 'NA')
        self.ws.write(self.rowsize, 16, self.xl_Expected_A3_T2[loop] if self.xl_Expected_A3_T2[loop] else 'NA')
        self.ws.write(self.rowsize, 17,
                      self.xl_Expected_A3_T2_Status[loop] if self.xl_Expected_A3_T2_Status[loop] else 'NA')
        self.ws.write(self.rowsize, 18,
                      self.xl_Expected_Task_ids[loop] if self.xl_Expected_Task_ids[loop] else 'NA')

        # -------------------
        # Writing Output Data
        # -------------------
        self.rowsize += 1  # Row increment

        self.ws.write(self.rowsize, self.col, 'Output', self.style5)
        # --------------------------------------------------------------------------------------------------------------

        if self.success:
            if self.applicant_mjr_dict['Stage'] == self.xl_Expected_Stage[loop]:
                if self.applicant_mjr_dict['Status'] == self.xl_Expected_Status[loop]:
                    a_id = self.Actual_Task_ids.split(',')
                    e_id = self.xl_Expected_Task_ids[loop].split(',')
                    e_id.sort()
                    a_id.sort()
                    if e_id == a_id:
                        self.ws.write(self.rowsize, 1, 'Pass', self.style26)
                        self.success_case_01 = 'Pass'
                    else:
                        self.ws.write(self.rowsize, 1, 'Fail', self.style3)
                else:
                    self.ws.write(self.rowsize, 1, 'Fail', self.style3)
            else:
                self.ws.write(self.rowsize, 1, 'Fail', self.style3)
        else:
            self.ws.write(self.rowsize, 1, 'Fail', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.applicant_dict_details:
            if self.applicant_dict_details['CandidateId'] == self.xl_candidateId[loop]:
                self.ws.write(self.rowsize, 2, self.applicant_dict_details['CandidateId'], self.style14)
            else:
                self.ws.write(self.rowsize, 2, self.applicant_dict_details['CandidateId'], self.style3)
        else:
            self.ws.write(self.rowsize, 2, self.applicant_dict_details.get('CandidateId'))
        # --------------------------------------------------------------------------------------------------------------

        if self.applicant_mjr_dict:
            if self.applicant_mjr_dict['Id'] == self.xl_applicantId[loop]:
                self.ws.write(self.rowsize, 3, self.applicant_mjr_dict['Id'], self.style14)
            else:
                self.ws.write(self.rowsize, 3, self.applicant_mjr_dict['Id'], self.style3)
        else:
            self.ws.write(self.rowsize, 3, self.applicant_mjr_dict.get('Id'))
        # --------------------------------------------------------------------------------------------------------------

        if self.applicant_mjr_dict:
            if self.applicant_mjr_dict['Stage'] == self.xl_Expected_Stage[loop]:
                self.ws.write(self.rowsize, 4, self.applicant_mjr_dict['Stage'], self.style14)
            else:
                self.ws.write(self.rowsize, 4, self.applicant_mjr_dict['Stage'], self.style3)
        else:
            self.ws.write(self.rowsize, 4, self.applicant_mjr_dict.get('Stage'))
        # --------------------------------------------------------------------------------------------------------------

        if self.applicant_mjr_dict:
            if self.applicant_mjr_dict['Status'] == self.xl_Expected_Status[loop]:
                self.ws.write(self.rowsize, 5, self.applicant_mjr_dict['Status'], self.style14)
            else:
                self.ws.write(self.rowsize, 5, self.applicant_mjr_dict['Status'], self.style3)
        else:
            self.ws.write(self.rowsize, 5, self.applicant_mjr_dict.get('Status'))
        # --------------------------------------------------------------------------------------------------------------

        if self.A1_t1:
            if self.xl_Expected_A1_T1[loop] == self.A1_t1:
                self.ws.write(self.rowsize, 6, self.A1_t1, self.style14)
            else:
                self.ws.write(self.rowsize, 6, self.A1_t1, self.style3)
        elif self.xl_Expected_A1_T1[loop] is None:
            self.ws.write(self.rowsize, 6, 'NA', self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.A1_t1_status:
            if self.xl_Expected_A1_T1_Status[loop] == self.A1_t1_status:
                self.ws.write(self.rowsize, 7, self.A1_t1_status, self.style14)
            else:
                self.ws.write(self.rowsize, 7, self.A1_t1_status, self.style3)
        elif self.xl_Expected_A1_T1_Status[loop] is None:
            self.ws.write(self.rowsize, 7, 'NA', self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.A2_t1:
            if self.xl_Expected_A2_T1[loop] == self.A2_t1:
                self.ws.write(self.rowsize, 8, self.A2_t1, self.style14)
            else:
                self.ws.write(self.rowsize, 8, self.A2_t1, self.style3)
        elif self.xl_Expected_A2_T1[loop] is None:
            self.ws.write(self.rowsize, 8, 'NA', self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.A2_t1_status:
            if self.xl_Expected_A2_T1_Status[loop] == self.A2_t1_status:
                self.ws.write(self.rowsize, 9, self.A2_t1_status, self.style14)
            else:
                self.ws.write(self.rowsize, 9, self.A2_t1_status, self.style3)
        elif self.xl_Expected_A2_T1_Status[loop] is None:
            self.ws.write(self.rowsize, 9, 'NA', self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.A2_t2:
            if self.xl_Expected_A2_T2[loop] == self.A2_t2:
                self.ws.write(self.rowsize, 10, self.A2_t2, self.style14)
            else:
                self.ws.write(self.rowsize, 10, self.A2_t2, self.style3)
        elif self.xl_Expected_A2_T2[loop] is None:
            self.ws.write(self.rowsize, 10, 'NA', self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.A2_t2_status:
            if self.xl_Expected_A2_T2_Status[loop] == self.A2_t2_status:
                self.ws.write(self.rowsize, 11, self.A2_t2_status, self.style14)
            else:
                self.ws.write(self.rowsize, 11, self.A2_t2_status, self.style3)
        elif self.xl_Expected_A2_T2_Status[loop] is None:
            self.ws.write(self.rowsize, 11, 'NA', self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.A2_t3:
            if self.xl_Expected_A2_T3[loop] == self.A2_t3:
                self.ws.write(self.rowsize, 12, self.A2_t3, self.style14)
            else:
                self.ws.write(self.rowsize, 12, self.A2_t3, self.style3)
        elif self.xl_Expected_A2_T3[loop] is None:
            self.ws.write(self.rowsize, 12, 'NA', self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.A2_t3_status:
            if self.xl_Expected_A2_T3_Status[loop] == self.A2_t3_status:
                self.ws.write(self.rowsize, 13, self.A2_t3_status, self.style14)
            else:
                self.ws.write(self.rowsize, 13, self.A2_t3_status, self.style3)
        elif self.xl_Expected_A2_T3_Status[loop] is None:
            self.ws.write(self.rowsize, 13, 'NA', self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.A3_t1:
            if self.xl_Expected_A3_T1[loop] == self.A3_t1:
                self.ws.write(self.rowsize, 14, self.A3_t1, self.style14)
            else:
                self.ws.write(self.rowsize, 14, self.A3_t1, self.style3)
        elif self.xl_Expected_A3_T1[loop] is None:
            self.ws.write(self.rowsize, 14, 'NA', self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.A3_t1_status:
            if self.xl_Expected_A3_T1_Status[loop] == self.A3_t1_status:
                self.ws.write(self.rowsize, 15, self.A3_t1_status, self.style14)
            else:
                self.ws.write(self.rowsize, 15, self.A3_t1_status, self.style3)
        elif self.xl_Expected_A3_T1_Status[loop] is None:
            self.ws.write(self.rowsize, 15, 'NA', self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.A3_t2:
            if self.xl_Expected_A3_T2[loop] == self.A3_t2:
                self.ws.write(self.rowsize, 16, self.A3_t2, self.style14)
            else:
                self.ws.write(self.rowsize, 16, self.A3_t2, self.style3)
        elif self.xl_Expected_A3_T2[loop] is None:
            self.ws.write(self.rowsize, 16, 'NA', self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.A3_t2_status:
            if self.xl_Expected_A3_T2_Status[loop] == self.A3_t2_status:
                self.ws.write(self.rowsize, 17, self.A3_t2_status, self.style14)
            else:
                self.ws.write(self.rowsize, 17, self.A3_t2_status, self.style3)
        elif self.xl_Expected_A3_T2_Status[loop] is None:
            self.ws.write(self.rowsize, 17, 'NA', self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.Actual_Task_ids:
            actual_id = self.Actual_Task_ids.split(',')
            expected_id = self.xl_Expected_Task_ids[loop].split(',')
            expected_id.sort()
            actual_id.sort()
            if expected_id == actual_id:
                self.ws.write(self.rowsize, 18, self.Actual_Task_ids, self.style14)
            else:
                self.ws.write(self.rowsize, 18, self.Actual_Task_ids, self.style3)
        else:
            self.ws.write(self.rowsize, 18, self.Actual_Task_ids, self.style3)
        # --------------------------------------------------------------------------------------------------------------

        self.rowsize += 1  # Row increment
        Object.wb_Result.save(output_paths.outputpaths['Activity_CallBack_Output_sheet'])

        if self.success_case_01 == 'Pass':
            self.Actual_Success_case.append(self.success_case_01)

    def over_status(self):
        self.ws.write(0, 0, 'Activity CallBack', self.style23)
        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
        else:
            self.ws.write(0, 1, 'Fail', self.style25)

        self.ws.write(0, 2, 'Start Time', self.style23)
        self.ws.write(0, 3, self.start_time, self.style26)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)
        Object.wb_Result.save(output_paths.outputpaths['Activity_CallBack_Output_sheet'])


Object = ActivityCallBack()
Object.excel_headers()
Object.excel_data()
Total_count = len(Object.xl_applicantId)
print("Number of rows::", Total_count)

if Object.login == 'OK':
    for looping in range(0, Total_count):
        print("Iteration Count is ::", looping)
        Object.change_applicant_status(looping)
        if Object.success:
            Object.get_ticket_number(looping)
            Object.submit_form_a1(looping)
            if Object.xl_Activity_02[looping]:
                Object.get_ticket_number(looping)
                Object.submit_form_a2(looping)
            if Object.xl_Activity_03[looping]:
                Object.get_ticket_number(looping)
                Object.submit_form_a3(looping)

            Object.applicant_info(looping)
            Object.get_activity_task_details(looping)
        Object.output_excel(looping)
        Object.reset_applicant_status(looping)

        # ----------------------------------------
        # Making all dicts are empty for each loop
        # ----------------------------------------
        Object.success = {}
        Object.chng_app_status = {}
        Object.applicant_dict_details = {}
        Object.applicant_mjr_dict = {}
        Object.ticket_number_A1_t1 = {}
        Object.ticket_number_A2_t1 = {}
        Object.ticket_number_A2_t2 = {}
        Object.ticket_number_A2_t3 = {}
        Object.ticket_number_A3_t1 = {}
        Object.ticket_number_A3_t2 = {}

        Object.A1 = {}
        Object.A1_t1 = {}
        Object.A2 = {}
        Object.A2_t1 = {}
        Object.A2_t2 = {}
        Object.A2_t3 = {}
        Object.A3 = {}
        Object.A3_t1 = {}
        Object.A3_t2 = {}

        Object.A1_t1_status = {}
        Object.A2_t1_status = {}
        Object.A2_t2_status = {}
        Object.A2_t3_status = {}
        Object.A3_t1_status = {}
        Object.A3_t2_status = {}

        Object.success_case_01 = {}
        Object.Actual_Task_ids = ""
        Object.headers = {}
Object.over_status()
