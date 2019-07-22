from hpro_automation import (login, work_book, input_paths, output_paths)
import datetime
import requests
import json
import xlrd


class ManageInterviewers(login.CommonLogin, work_book.WorkBook):

    def __init__(self):

        # ---------------------------- Overall Status Current Run Date -------------------------------------------------
        self.start_time = str(datetime.datetime.now())

        # --------------------------------- Inheritance Class Instance -------------------------------------------------
        super(ManageInterviewers, self).__init__()
        self.common_login('crpo')

        # --------------------------------- Overall status initialize variables ----------------------------------------
        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 16)))
        self.Actual_Success_case = []

        # --------------------------------- Excel Data initialize variables --------------------------------------------
        self.xl_event_id = []
        self.xl_compositeKey = []
        self.xl_interviewer_id = []
        self.xl_email_id = []
        self.xl_interviewer_decision = []
        self.xl_withdrawn_After_int = []
        self.xl_manager_decision = []
        self.xl_withdrawn_After_manager = []
        self.xl_manager_rejected = []
        self.xl_sync_interviewers = []
        self.xl_withdrawn_After_sync = []
        self.xl_success_message = []
        self.xl_waring_message = []
        self.xl_tagged_interviewers = []

        self.xl_nomi_sent = []
        self.xl_nomi_total_response = []
        self.xl_nomi_confirm_int = []
        self.xl_nomi_withdrawn = []
        self.xl_nomi_approve = []
        self.xl_nomi_rejected = []
        self.xl_report_event_id = []

        # --------------------------------- Dictionary initialize variables --------------------------------------------
        self.success_case_01 = {}
        self.success_case_02 = {}
        self.success_case_03 = {}
        self.success_case_04 = {}
        self.success_case_05 = {}
        self.success_case_06 = {}
        self.success_case_07 = {}
        self.success_case_08 = {}
        self.success_case_09 = {}
        self.headers = {}
        self.invited_interviewers_dict = {}
        self.invited_interviewers_dict_after_rejected = {}
        self.sync_data = {}
        self.owners_dict = {}
        self.composite_key_1 = {}

    def excel_headers(self):

        # --------------------------------- Excel Headers and Cell color, styles ---------------------------------------
        self.main_headers = ['Comparision', 'Status', 'InterviewerID', 'EmailID', 'InterviewerDecision',
                             'Withdraw-AfterInterviewerRequest', 'ManagerDecision', 'Withdraw-AfterManagerRequest',
                             'ApproveAndRejectedByManager', 'Sync', 'Withdraw-AfterSyncInterviewers',
                             'Tagged InterviewIDs', 'successMessage', 'WaringMessage']
        self.headers_with_style2 = ['Comparision', 'Status']
        self.headers_with_style9 = ['InterviewerID', 'EmailID', 'InterviewerDecision',
                                    'Withdraw-AfterInterviewerRequest', 'ManagerDecision',
                                    'Withdraw-AfterManagerRequest', 'Sync', 'Withdraw-AfterSyncInterviewers',
                                    'Tagged InterviewIDs', 'ApproveAndRejectedByManager', 'successMessage',
                                    'WaringMessage']
        self.file_headers_col_row()

    def excel_data(self):

        # ------------------------------- Excel Data Read --------------------------------------------------------------
        try:
            workbook = xlrd.open_workbook(input_paths.inputpaths['Manage_Int_Input_sheet'])
            sheet1 = workbook.sheet_by_index(0)
            for i in range(1, sheet1.nrows):
                number = i  # Counting number of rows
                rows = sheet1.row_values(number)

                if not rows[0]:
                    self.xl_event_id.append(None)
                else:
                    self.xl_event_id.append(int(rows[0]))

                if not rows[1]:
                    self.xl_compositeKey.append(None)
                else:
                    self.xl_compositeKey.append(int(rows[1]))

                if not rows[2]:
                    self.xl_interviewer_id.append(None)
                else:
                    self.xl_interviewer_id.append(int(rows[2]))

                if not rows[3]:
                    self.xl_email_id.append(None)
                else:
                    self.xl_email_id.append(str(rows[3]))

                if rows[4] is not None and rows[4] == '':
                    self.xl_interviewer_decision.append(None)
                else:
                    self.xl_interviewer_decision.append(int(rows[4]))

                if rows[5] is not None and rows[5] == '':
                    self.xl_withdrawn_After_int.append(None)
                else:
                    self.xl_withdrawn_After_int.append(int(rows[5]))

                if rows[6] is not None and rows[6] == '':
                    self.xl_manager_decision.append(None)
                else:
                    self.xl_manager_decision.append(int(rows[6]))

                if rows[7] is not None and rows[7] == '':
                    self.xl_withdrawn_After_manager.append(None)
                else:
                    self.xl_withdrawn_After_manager.append(int(rows[7]))

                if rows[8] is not None and rows[8] == '':
                    self.xl_manager_rejected.append(None)
                else:
                    self.xl_manager_rejected.append(int(rows[8]))

                if not rows[9]:
                    self.xl_sync_interviewers.append(None)
                else:
                    self.xl_sync_interviewers.append(str(rows[9]))

                if rows[10] is not None and rows[10] == '':
                    self.xl_withdrawn_After_sync.append(None)
                else:
                    self.xl_withdrawn_After_sync.append(int(rows[10]))

                if rows[11] is not None and rows[11] == '':
                    self.xl_success_message.append(None)
                else:
                    self.xl_success_message.append(str(rows[11]))

                if rows[12] is not None and rows[12] == '':
                    self.xl_waring_message.append(None)
                else:
                    self.xl_waring_message.append(str(rows[12]))

                if rows[13] is not None and rows[13] == '':
                    self.xl_tagged_interviewers.append(None)
                else:
                    self.xl_tagged_interviewers.append(str(rows[13]))

        except IOError:
            print("File not found or path is incorrect")

        try:
            workbook1 = xlrd.open_workbook(input_paths.inputpaths['Manage_Int_Input_sheet'])
            sheet1 = workbook1.sheet_by_index(1)
            for i in range(1, sheet1.nrows):
                number = i  # Counting number of rows
                rows = sheet1.row_values(number)

                self.xl_nomi_sent.append(int(rows[0]))
                self.xl_nomi_total_response.append(int(rows[1]))
                self.xl_nomi_confirm_int.append(int(rows[2]))
                self.xl_nomi_withdrawn.append(int(rows[3]))
                self.xl_nomi_approve.append(int(rows[4]))
                self.xl_nomi_rejected.append(int(rows[5]))
                self.xl_report_event_id.append(int(rows[6]))

        except IOError:
            print("File not found or path is incorrect")

    def send_nomination_mails(self, loop):

        self.lambda_function('send_nomination_mails_to_selected_interviewers')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"eventId": self.xl_event_id[loop],
                   "data": {
                       self.xl_compositeKey[loop]: [{
                           "interviewerId": self.xl_interviewer_id[loop],
                           "emailId": self.xl_email_id[loop]
                       }]
                   }}

        send_nomination_api = requests.post(self.webapi, headers=self.headers, data=json.dumps(request, default=str),
                                            verify=False)
        print(send_nomination_api.headers)
        send_nomination_api_response = json.loads(send_nomination_api.content)
        print(send_nomination_api_response)

    def interviewer_request(self, loop):

        self.lambda_function('update_interviewer_nomination_status')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"eventId": self.xl_event_id[loop],
                   "data": {
                       self.xl_compositeKey[loop]: [{
                           "interviewerId": self.xl_interviewer_id[loop],
                           "emailId": self.xl_email_id[loop],
                           "decisionByInterviewer": self.xl_interviewer_decision[loop]
                       }]
                   }}

        interviewer_request_api = requests.post(self.webapi, headers=self.headers,
                                                data=json.dumps(request, default=str), verify=False)
        print(interviewer_request_api.headers)
        interviewer_request_api_response = json.loads(interviewer_request_api.content)
        print(interviewer_request_api_response)

    def withdraw_after_int_decision(self, loop):
        self.lambda_function('update_interviewer_nomination_status')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"eventId": self.xl_event_id[loop],
                   "data": {
                       self.xl_compositeKey[loop]: [{
                           "interviewerId": self.xl_interviewer_id[loop],
                           "emailId": self.xl_email_id[loop],
                           "decisionByInterviewer": self.xl_withdrawn_After_int[loop]
                       }]
                   }}

        interviewer_req_wd_api = requests.post(self.webapi, headers=self.headers,
                                               data=json.dumps(request, default=str), verify=False)
        interviewer_req_wd_api_response = json.loads(interviewer_req_wd_api.content)
        print(interviewer_req_wd_api_response)

    def withdraw_after_sync(self, loop):
        self.lambda_function('update_interviewer_nomination_status')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"eventId": self.xl_event_id[loop],
                   "data": {
                       self.xl_compositeKey[loop]: [{
                           "interviewerId": self.xl_interviewer_id[loop],
                           "emailId": self.xl_email_id[loop],
                           "decisionByInterviewer": self.xl_withdrawn_After_sync[loop]
                       }]
                   }}

        withdraw_after_sync_api = requests.post(self.webapi, headers=self.headers,
                                                data=json.dumps(request, default=str), verify=False)
        withdraw_after_sync_api_response = json.loads(withdraw_after_sync_api.content)
        print(withdraw_after_sync_api_response)

    def event_manager_request(self, loop):

        self.lambda_function('update_interviewer_nomination_status')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"eventId": self.xl_event_id[loop],
                   "data": {
                       self.xl_compositeKey[loop]: [{
                           "interviewerId": self.xl_interviewer_id[loop],
                           "emailId": self.xl_email_id[loop],
                           "compositeKey": self.xl_compositeKey[loop],
                           "statusCode": 4,
                           "decisionByEventManager": self.xl_manager_decision[loop]
                       }]
                   }}

        event_manager_api = requests.post(self.webapi, headers=self.headers, data=json.dumps(request, default=str),
                                          verify=False)
        print(event_manager_api.headers)
        event_manager_api_response = json.loads(event_manager_api.content)
        print(event_manager_api_response)

    def reject_event_manager_request(self, loop):

        self.lambda_function('update_interviewer_nomination_status')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"eventId": self.xl_event_id[loop],
                   "data": {
                       self.xl_compositeKey[loop]: [{
                           "interviewerId": self.xl_interviewer_id[loop],
                           "emailId": self.xl_email_id[loop],
                           "compositeKey": self.xl_compositeKey[loop],
                           "statusCode": 4,
                           "decisionByEventManager": self.xl_manager_rejected[loop]
                       }]
                   }}

        reject_event_manager_api = requests.post(self.webapi, headers=self.headers,
                                                 data=json.dumps(request, default=str), verify=False)
        print(reject_event_manager_api.headers)
        reject_event_manager_api_response = json.loads(reject_event_manager_api.content)
        print(reject_event_manager_api_response)

    def withdraw_after_manager_decision(self, loop):
        self.lambda_function('update_interviewer_nomination_status')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"eventId": self.xl_event_id[loop],
                   "data": {
                       self.xl_compositeKey[loop]: [{
                           "interviewerId": self.xl_interviewer_id[loop],
                           "emailId": self.xl_email_id[loop],
                           "decisionByInterviewer": self.xl_withdrawn_After_manager[loop]
                       }]
                   }}

        manager_req_wd_api = requests.post(self.webapi, headers=self.headers,
                                           data=json.dumps(request, default=str), verify=False)
        manager_req_wd_api_response = json.loads(manager_req_wd_api.content)
        print(manager_req_wd_api_response)

    def sync_interviewers(self, loop):

        self.lambda_function('sync_interviewers')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"eventId": self.xl_event_id[loop]}

        sync_interviewers_api = requests.post(self.webapi, headers=self.headers,
                                              data=json.dumps(request, default=str), verify=False)
        sync_interviewers_api_response = json.loads(sync_interviewers_api.content)
        self.sync_data = sync_interviewers_api_response.get('data')
        print(self.sync_data)

    def get_all_invited_interviewers(self, loop):

        self.lambda_function('get_all_invited_interviewers')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"eventId": self.xl_event_id[loop],
                   "pagingCriteria": {"pageSize": 100, "pageNumber": 1}
                   }

        invited_interviewers_api = requests.post(self.webapi, headers=self.headers,
                                                 data=json.dumps(request, default=str), verify=False)
        print(invited_interviewers_api.headers)
        invited_interviewers_response = json.loads(invited_interviewers_api.content)
        data = invited_interviewers_response.get('data')
        for i in data:
            if i['interviewerId'] == self.xl_interviewer_id[loop]:
                self.invited_interviewers_dict = i
                print(self.invited_interviewers_dict)

    def get_all_invited_interviewers_after_manager_reject(self, loop):

        self.lambda_function('get_all_invited_interviewers')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"eventId": self.xl_event_id[loop],
                   "pagingCriteria": {"pageSize": 100, "pageNumber": 1}
                   }

        invited_interviewers_api = requests.post(self.webapi, headers=self.headers,
                                                 data=json.dumps(request, default=str), verify=False)
        print(invited_interviewers_api.headers)
        invited_interviewers_response = json.loads(invited_interviewers_api.content)
        data = invited_interviewers_response.get('data')
        for i in data:
            if i['interviewerId'] == self.xl_interviewer_id[loop]:
                self.invited_interviewers_dict_after_rejected = i

    def get_partial_get_event_id(self, loop):

        self.lambda_function('getPartialGetEventForId')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"EId": self.xl_event_id[loop], "EOptions": [5]}

        get_event_id_api = requests.post(self.webapi, headers=self.headers,
                                         data=json.dumps(request, default=str), verify=False)
        print(get_event_id_api.headers)
        get_event_id_response = json.loads(get_event_id_api.content)
        owners = get_event_id_response.get('Owners')
        for i in owners:
            if i['UserId'] == self.xl_interviewer_id[loop]:
                self.owners_dict = i

    def get_event_wise_nominations_summary_count(self, loop):

        self.lambda_function('get_event_wise_nominations_summary_count')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- API request --------------------------------------------------------------
        request = {"eventId": self.xl_report_event_id[loop]}

        event_wise_nominations_api = requests.post(self.webapi, headers=self.headers,
                                                   data=json.dumps(request, default=str), verify=False)
        print(event_wise_nominations_api.headers)
        event_wise_nominations_response = json.loads(event_wise_nominations_api.content)
        data = event_wise_nominations_response.get('data')
        event = data[str(self.xl_report_event_id[loop])]
        composite_key = event[str(self.xl_compositeKey[loop])]
        self.composite_key_1 = composite_key

    def output_report(self, loop):

        # --------------------------------- Writing Input Data ---------------------------------------------------------
        self.ws.write(self.rowsize, self.col, 'Input', self.style4)
        self.ws.write(self.rowsize, 2, self.xl_interviewer_id[loop] if self.xl_interviewer_id[loop] else 'Empty')
        self.ws.write(self.rowsize, 3, self.xl_email_id[loop] if self.xl_email_id[loop] else 'Empty')
        self.ws.write(self.rowsize, 9, self.xl_sync_interviewers[loop])
        self.ws.write(self.rowsize, 12, self.xl_success_message[loop] if self.xl_success_message[loop] else 'Empty')
        self.ws.write(self.rowsize, 13, self.xl_waring_message[loop] if self.xl_waring_message[loop] else 'Empty')
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_interviewer_decision[loop] == 1:
            int_decision = 'Confirm'
            self.ws.write(self.rowsize, 4, int_decision)
        elif self.xl_interviewer_decision[loop] == 0:
            int_decision = 'Decline'
            self.ws.write(self.rowsize, 4, int_decision)
        elif self.xl_interviewer_decision[loop] == 2:
            int_decision = 'Withdrawn'
            self.ws.write(self.rowsize, 4, int_decision)
        else:
            self.ws.write(self.rowsize, 4, 'Pending')
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_withdrawn_After_int[loop] is None:
            self.ws.write(self.rowsize, 5, 'NA')
        else:
            if self.xl_withdrawn_After_int[loop] == 2:
                withdrawn_decision_by_int = 'Withdrawn'
                self.ws.write(self.rowsize, 5, withdrawn_decision_by_int)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_manager_decision[loop] is None:
            self.ws.write(self.rowsize, 6, 'NA')
        else:
            if self.xl_manager_decision[loop] == 1:
                manager_decision = 'Approved'
                self.ws.write(self.rowsize, 6, manager_decision)
            else:
                manager_decision = 'Rejected'
                self.ws.write(self.rowsize, 6, manager_decision)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_withdrawn_After_manager[loop] is None:
            self.ws.write(self.rowsize, 7, 'NA')
        else:
            if self.xl_withdrawn_After_manager[loop] == 2:
                withdrawn_decision_by_man = 'Withdrawn'
                self.ws.write(self.rowsize, 7, withdrawn_decision_by_man)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_manager_rejected[loop] is None:
            self.ws.write(self.rowsize, 8, 'NA')
        else:
            if self.xl_manager_rejected[loop] == 0:
                approve_rejected_manager = 'Rejected'
                self.ws.write(self.rowsize, 8, approve_rejected_manager)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_withdrawn_After_sync[loop] is None:
            self.ws.write(self.rowsize, 10, 'NA')
        else:
            if self.xl_withdrawn_After_sync[loop] == 2:
                withdrawn_after_sync = 'Withdrawn'
                self.ws.write(self.rowsize, 10, withdrawn_after_sync)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_tagged_interviewers[loop] == 'Not Tagged':
            self.ws.write(self.rowsize, 11, 'Not Tagged')
        elif self.owners_dict:
            self.ws.write(self.rowsize, 11, self.xl_tagged_interviewers[loop].format(self.xl_interviewer_id[loop]))

        # --------------------------------------------------------------------------------------------------------------
        # ---------------------------------------- Writing Output Data -------------------------------------------------
        # --------------------------------------------------------------------------------------------------------------
        self.rowsize += 1

        self.ws.write(self.rowsize, self.col, 'Output', self.style5)

        if self.invited_interviewers_dict['interviewerId'] == self.xl_interviewer_id[loop]:
            if self.invited_interviewers_dict['emailId'] == self.xl_email_id[loop]:
                self.ws.write(self.rowsize, 1, 'Pass', self.style26)
                self.success_case_01 = 'Pass'
            else:
                self.ws.write(self.rowsize, 1, 'Fail', self.style3)
        else:
            self.ws.write(self.rowsize, 1, 'Fail', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.invited_interviewers_dict['interviewerId'] == self.xl_interviewer_id[loop]:
            self.ws.write(self.rowsize, 2, self.invited_interviewers_dict['interviewerId'], self.style8)
        else:
            self.ws.write(self.rowsize, 2, self.invited_interviewers_dict.get('interviewerId', None), self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.invited_interviewers_dict['emailId'] == self.xl_email_id[loop]:
            self.ws.write(self.rowsize, 3, self.invited_interviewers_dict['emailId'], self.style8)
        else:
            self.ws.write(self.rowsize, 3, self.invited_interviewers_dict.get('emailId', None), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_interviewer_decision[loop] == 0 or self.xl_interviewer_decision[loop] == 1\
                or self.xl_interviewer_decision[loop] == 2 or self.xl_interviewer_decision[loop] is None:
            if self.invited_interviewers_dict['isNominatedSelf']:
                self.ws.write(self.rowsize, 4, 'Confirm', self.style8)
            elif self.invited_interviewers_dict['isNominatedSelf'] is None:
                self.ws.write(self.rowsize, 4, 'Pending', self.style8)
            else:
                self.ws.write(self.rowsize, 4, 'Decline', self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_withdrawn_After_int[loop] == 2:
            if self.invited_interviewers_dict['isRefusedByInterviewerAfterAcceptance']:
                self.ws.write(self.rowsize, 5, 'Withdraw', self.style8)
        elif self.xl_withdrawn_After_int[loop] is None:
            self.ws.write(self.rowsize, 5, 'NA', self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_manager_decision[loop] == 1:
            if self.invited_interviewers_dict['isNominationAccepted']:
                self.ws.write(self.rowsize, 6, 'Approved', self.style8)
        elif self.xl_manager_decision[loop] == 0:
            self.ws.write(self.rowsize, 6, 'Rejected', self.style8)
        else:
            self.ws.write(self.rowsize, 6, 'NA', self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_withdrawn_After_manager[loop] == 2:
            if self.invited_interviewers_dict['isRefusedByInterviewerAfterAcceptance']:
                self.ws.write(self.rowsize, 7, 'Withdraw', self.style8)
        elif self.xl_withdrawn_After_manager[loop] is None:
            self.ws.write(self.rowsize, 7, 'NA', self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_manager_rejected[loop] == 0:
            if self.invited_interviewers_dict_after_rejected['isNominationAccepted']:
                self.ws.write(self.rowsize, 8, 'Approved', self.style3)
            elif self.invited_interviewers_dict_after_rejected['isNominationAccepted'] is None:
                self.ws.write(self.rowsize, 8, 'Pending', self.style3)
            else:
                self.ws.write(self.rowsize, 8, 'Rejected', self.style8)
        else:
            self.ws.write(self.rowsize, 8, 'NA', self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.sync_data:
            if self.sync_data['successMessage']:
                self.ws.write(self.rowsize, 9, 'Yes', self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_withdrawn_After_sync[loop] == 2:
            if self.invited_interviewers_dict['isRefusedByInterviewerAfterAcceptance']:
                self.ws.write(self.rowsize, 10, 'Withdraw', self.style8)
        elif self.xl_withdrawn_After_sync[loop] is None:
            self.ws.write(self.rowsize, 10, 'NA', self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_tagged_interviewers[loop]:
            if self.owners_dict:
                if self.xl_interviewer_id[loop] == self.owners_dict['UserId']:
                    self.ws.write(self.rowsize, 11, 'Tagged_{}'.format(self.owners_dict['UserId']), self.style8)
                else:
                    self.ws.write(self.rowsize, 11, 'Tagged_{}'.format(self.owners_dict['UserId']), self.style3)
            else:
                self.ws.write(self.rowsize, 11, self.xl_tagged_interviewers[loop], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.sync_data:
            if self.xl_success_message[loop] == self.sync_data['successMessage']:
                self.ws.write(self.rowsize, 12, self.sync_data['successMessage'], self.style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.sync_data:
            if self.xl_waring_message[loop] == self.sync_data.get('warningMessage'):
                self.ws.write(self.rowsize, 13, self.sync_data.get('warningMessage', 'Empty'), self.style8)

        # ------------------------------------ OutPut File save --------------------------------------------------------
        self.rowsize += 1
        Object.wb_Result.save(output_paths.outputpaths['MI_output_sheet'])

        if self.success_case_01 == 'Pass':
            self.Actual_Success_case.append(self.success_case_01)
        if self.success_case_02 == 'Pass':
            self.Actual_Success_case.append(self.success_case_02)
        if self.success_case_03 == 'Pass':
            self.Actual_Success_case.append(self.success_case_03)

    def report_count(self, loop):
        self.ws.write(24, self.col, 'Comparison', self.style23)
        self.ws.write(24, 1, 'Status', self.style23)
        self.ws.write(24, 2, 'Nominations Sent', self.style27)
        self.ws.write(24, 3, 'Total Response from Interviewer', self.style27)
        self.ws.write(24, 4, 'Confirm by Interviewer', self.style27)
        self.ws.write(24, 5, 'Withdrawn by Interviewer', self.style27)
        self.ws.write(24, 6, 'Approve by EM', self.style27)
        self.ws.write(24, 7, 'Rejected by EM', self.style27)

        self.ws.write(25, self.col, 'Input', self.style4)
        self.ws.write(25, 2, self.xl_nomi_sent[loop], self.style28)
        self.ws.write(25, 3, self.xl_nomi_total_response[loop], self.style28)
        self.ws.write(25, 4, self.xl_nomi_confirm_int[loop], self.style28)
        self.ws.write(25, 5, self.xl_nomi_withdrawn[loop], self.style28)
        self.ws.write(25, 6, self.xl_nomi_approve[loop], self.style28)
        self.ws.write(25, 7, self.xl_nomi_rejected[loop], self.style28)
        # --------------------------------------------------------------------------------------------------------------
        self.ws.write(26, self.col, 'Output', self.style5)

        if self.xl_nomi_sent[loop] == self.composite_key_1['nominationsSentCount']:
            self.ws.write(26, 2, self.composite_key_1['nominationsSentCount'], self.style24)
            self.success_case_04 = 'Pass'
        else:
            self.ws.write(26, 2, self.composite_key_1['nominationsSentCount'], self.style25)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_nomi_total_response[loop] == self.composite_key_1['selfNominatedRespondCount']:
            self.ws.write(26, 3, self.composite_key_1['selfNominatedRespondCount'], self.style24)
            self.success_case_05 = 'Pass'
        else:
            self.ws.write(26, 3, self.composite_key_1['selfNominatedRespondCount'], self.style25)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_nomi_confirm_int[loop] == self.composite_key_1['selfNominatedCount']:
            self.ws.write(26, 4, self.composite_key_1['selfNominatedCount'], self.style24)
            self.success_case_06 = 'Pass'
        else:
            self.ws.write(26, 4, self.composite_key_1['selfNominatedCount'], self.style25)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_nomi_withdrawn[loop] == self.composite_key_1['refusedByInterviewerAfterAcceptanceCount']:
            self.ws.write(26, 5, self.composite_key_1['refusedByInterviewerAfterAcceptanceCount'], self.style24)
            self.success_case_07 = 'Pass'
        else:
            self.ws.write(26, 5, self.composite_key_1['refusedByInterviewerAfterAcceptanceCount'], self.style25)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_nomi_approve[loop] == self.composite_key_1['nominationApprovedByEMCount']:
            self.ws.write(26, 6, self.composite_key_1['nominationApprovedByEMCount'], self.style24)
            self.success_case_08 = 'Pass'
        else:
            self.ws.write(26, 6, self.composite_key_1['nominationApprovedByEMCount'], self.style25)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_nomi_rejected[loop] == self.composite_key_1['nominationRejectedByEMCount']:
            self.ws.write(26, 7, self.composite_key_1['nominationRejectedByEMCount'], self.style24)
            self.success_case_09 = 'Pass'
        else:
            self.ws.write(26, 7, self.composite_key_1['nominationRejectedByEMCount'], self.style25)

        if self.success_case_04 == 'Pass':
            self.Actual_Success_case.append(self.success_case_04)
        if self.success_case_05 == 'Pass':
            self.Actual_Success_case.append(self.success_case_05)
        if self.success_case_06 == 'Pass':
            self.Actual_Success_case.append(self.success_case_06)
        if self.success_case_07 == 'Pass':
            self.Actual_Success_case.append(self.success_case_07)
        if self.success_case_08 == 'Pass':
            self.Actual_Success_case.append(self.success_case_08)
        if self.success_case_09 == 'Pass':
            self.Actual_Success_case.append(self.success_case_09)

        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(26, 1, 'Pass', self.style26)
        else:
            self.ws.write(26, 1, 'Fail', self.style3)

        Object.wb_Result.save(output_paths.outputpaths['MI_output_sheet'])

    def overall_status(self):
        self.ws.write(0, 0, 'Manage Interviewers', self.style23)
        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
        else:
            self.ws.write(0, 1, 'Fail', self.style25)

        self.ws.write(0, 2, 'Start Time', self.style23)
        self.ws.write(0, 3, self.start_time, self.style26)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)
        # ---------------------------- OutPut File save with Overall Status --------------------------------------------
        Object.wb_Result.save(output_paths.outputpaths['MI_output_sheet'])


Object = ManageInterviewers()
Object.excel_headers()
Object.excel_data()
Total_count = len(Object.xl_event_id)
print("Number Of Rows ::", Total_count)
if Object.login == 'OK':
    for looping in range(0, Total_count):
        print("Iteration Count is ::", looping)
        Object.send_nomination_mails(looping)
        Object.interviewer_request(looping)

        if Object.xl_withdrawn_After_int[looping] == 2:
            Object.withdraw_after_int_decision(looping)

        Object.event_manager_request(looping)

        if Object.xl_withdrawn_After_manager[looping] == 2:
            Object.withdraw_after_manager_decision(looping)

        Object.get_all_invited_interviewers(looping)

        if Object.xl_manager_rejected[looping] == 0:
            Object.reject_event_manager_request(looping)
            Object.get_all_invited_interviewers_after_manager_reject(looping)

        Object.sync_interviewers(looping)

        if Object.xl_withdrawn_After_sync[looping] == 2:
            Object.withdraw_after_sync(looping)
            Object.get_all_invited_interviewers(looping)

        Object.get_partial_get_event_id(looping)
        Object.output_report(looping)

        # ----------------- Make Dictionaries clear for each loop ------------------------------------------------------
        Object.success_case_01 = {}
        Object.success_case_02 = {}
        Object.success_case_03 = {}
        Object.headers = {}
        Object.invited_interviewers_dict = {}
        Object.invited_interviewers_dict_after_rejected = {}
        Object.sync_data = {}
        Object.owners_dict = {}

Total_count1 = len(Object.xl_report_event_id)
for looping in range(0, Total_count1):
    print("Report Iteration Count is ::", looping)
    Object.get_event_wise_nominations_summary_count(looping)
    Object.report_count(looping)

    Object.composite_key_1 = {}
    Object.success_case_04 = {}
    Object.success_case_05 = {}
    Object.success_case_06 = {}
    Object.success_case_07 = {}
    Object.success_case_08 = {}
    Object.success_case_09 = {}

Object.overall_status()
