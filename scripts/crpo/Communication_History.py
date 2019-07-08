from hpro_automation import (login, input_paths, output_paths, work_book)
import requests
import json
import xlrd
import datetime
import time
from selenium import webdriver
from selenium.common import exceptions


class CommunicationHistory(login.CRPOLogin, work_book.WorkBook):

    def __init__(self):
        self.start_time = str(datetime.datetime.now())
        super(CommunicationHistory, self).__init__()

        self.driver = ""
        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 14)))
        self.Actual_Success_case = []

        self.purpose = []
        self.xl_applicant_id = []
        self.xl_label_id = []
        self.xl_s3_file = []
        self.xl_CommunicationPurpose = []
        self.xl_verification_type = []

        self.xl_expected_app_id = []
        self.xl_expected_event_id = []
        self.xl_expected_admit = []
        self.xl_expected_rl = []
        self.xl_expected_disable_rl = []
        self.xl_expected_enable_rl = []
        self.xl_expected_rl_done = []
        self.xl_expected_re_rl_not_allowed = []
        self.xl_expected_re_rl_allowed = []
        self.xl_expected_re_rl_done = []
        self.xl_expected_GDPR = []
        self.xl_expected_approved_admit = []
        self.xl_expected_score = []
        self.xl_expected_offer = []
        self.xl_expected_email_mobile = []
        self.xl_expected_email = []
        self.xl_API_hits = []

        # -------------
        # Dicitionores
        # -------------
        self.applicant_dict = {}
        self.communication_statuses = {}
        self.api_admitcard = {}
        self.api_rl = {}
        self.api_approve_admitcard = {}
        self.api_score_sheet = {}
        self.api_offer = {}
        self.api_email_v = {}
        self.api_mobile_v = {}
        self.api_disable_rl = {}
        self.api_enable_rl = {}

        self.success_case_01 = {}
        self.data = {}
        self.AttachmentId = {}
        self.api_get_rl = {}
        self.api_rl_done = {}
        self.api_re_rl_allowed = {}
        self.api_re_rl_Not_allowed = {}
        self.api_GDPR = {}
        self.api_re_rl_done = {}
        self.headers = {}

        self.excel_headers()

    def excel_headers(self):
        self.main_headers = ['Actual_Status', 'Name of Flag', 'Event_Id', 'Expected Applicant_Id',
                             'Actual_Applicant_Id', 'Admit_card', 'Actual_Admit_card', 'RL_sent', 'Actual_RL_sent',
                             'Disable_RL', 'Actual_Disable_RL', 'Enable_RL', 'Actual_Enable_RL', 'RL_Done',
                             'Actual_RL_Done', 'Re_RL', 'Actual_Re_RL', 'Re_RL_Done', 'Actual_Re_Rl_Done',
                             'Re_RL_NotAllowed', 'Actual_Re_RL_NotAllowed', 'GDPR', 'Actual_GDPR',
                             'Approved_admit_card', 'Actual_Approved_admit_card', 'score_sheet', 'Actual_score_sheet',
                             'Offer', 'Actual_Offer', 'Email/Mobile', 'Actual_Email/Mobile',
                             'Email', 'Actual_Email', 'Attachment_Id']
        self.headers_with_style2 = ['Actual_Status', 'Name of Flag', 'Event_Id']
        self.headers_with_style9 = ['Actual_Applicant_Id', 'Actual_Admit_card', 'Actual_RL_sent', 'Actual_Disable_RL',
                                    'Actual_Enable_RL', 'Actual_RL_Done', 'Actual_Re_RL', 'Actual_Re_Rl_Done',
                                    'Actual_Re_RL_NotAllowed', 'Actual_GDPR', 'Actual_Approved_admit_card',
                                    'Actual_score_sheet', 'Actual_Offer', 'Actual_Email/Mobile', 'Actual_Email']
        self.file_headers_col_row()

    def excel_data(self):

        # -----------------------------
        # Communication Excel Data Read
        # -----------------------------
        workbook = xlrd.open_workbook(input_paths.inputpaths['Communication_History_Input_sheet'])
        sheet1 = workbook.sheet_by_index(0)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)

            if rows[0]:
                self.purpose.append(rows[0])
            else:
                self.purpose.append(None)
            if rows[1]:
                self.xl_applicant_id.append(int(rows[1]))
            else:
                self.xl_applicant_id.append(None)
            if rows[2]:
                self.xl_s3_file.append(rows[2])
            else:
                self.xl_s3_file.append(None)
            if rows[3]:
                self.xl_label_id.append(int(rows[3]))
            else:
                self.xl_label_id.append(None)
            if rows[4]:
                self.xl_CommunicationPurpose.append(int(rows[4]))
            else:
                self.xl_CommunicationPurpose.append(None)
            if rows[5]:
                self.xl_verification_type.append(int(rows[5]))
            else:
                self.xl_verification_type.append(None)
            if rows[8]:
                self.xl_expected_app_id.append(int(rows[8]))
            else:
                self.xl_expected_app_id.append(None)
            if rows[9]:
                self.xl_expected_event_id.append(int(rows[9]))
            else:
                self.xl_expected_event_id.append(None)

            self.xl_expected_admit.append(rows[10])
            self.xl_expected_rl.append(rows[11])
            self.xl_expected_disable_rl.append(rows[12])
            self.xl_expected_enable_rl.append(rows[13])
            self.xl_expected_rl_done.append(rows[14])
            self.xl_expected_re_rl_allowed.append(rows[15])
            self.xl_expected_re_rl_done.append(rows[16])
            self.xl_expected_re_rl_not_allowed.append(rows[17])
            self.xl_expected_GDPR.append(rows[18])
            self.xl_expected_approved_admit.append(rows[19])
            self.xl_expected_score.append(rows[20])
            self.xl_expected_offer.append(rows[21])
            self.xl_expected_email_mobile.append(rows[22])
            self.xl_expected_email.append(rows[23])
            self.xl_API_hits.append(int(rows[24]))

    def admit_card(self, loop):
        self.lambda_function('sendAdmitCardsToApplicants')
        self.headers['APP-NAME'] = 'crpo'

        # ------------------------------------- Admit Card -------------------------------------------------------------
        request = {"Sync": "False",
                   "EventId": '',
                   "StatusId": '',
                   "ForceSend": "True",
                   "EnsureSuccess": "True",
                   "Applicants": [self.xl_applicant_id[loop]]}
        admitcard = requests.post(self.webapi, headers=self.headers, data=json.dumps(request, default=str),
                                  verify=False)
        print(admitcard.headers)
        admitcard_response = json.loads(admitcard.content)
        print(admitcard_response)

    def registration_link(self, loop):

        self.lambda_function('sendRegistrationLinkToApplicants')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- Registration Link --------------------------------------------------------
        request = {"ApplicantIds": [self.xl_applicant_id[loop]],
                   "Sync": "False",
                   "ForceSend": True
                   }
        rl = requests.post(self.webapi, headers=self.headers, data=json.dumps(request, default=str), verify=False)
        print(rl.headers)
        rl_response = json.loads(rl.content)
        print(rl_response)

    def registration_link_disable(self, loop):

        self.lambda_function('setApplicantCommunicationStatus')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- registration link disable -----------------------------------------------
        request1 = {"ApplicantId": self.xl_applicant_id[loop],
                    "CommunicationPurpose": self.xl_CommunicationPurpose[loop],
                    "CommunicationStatus": True}
        disable_api = requests.post(self.webapi,
                                    headers=self.headers, data=json.dumps(request1, default=str), verify=False)
        print(disable_api.headers)
        disable_api_dict = json.loads(disable_api.content)
        print(disable_api_dict)

    def registration_link_enable(self, loop):

        self.lambda_function('setApplicantCommunicationStatus')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- registration link enable-------------------------------------------------
        request1 = {"ApplicantId": self.xl_applicant_id[loop],
                    "CommunicationPurpose": self.xl_CommunicationPurpose[loop],
                    "CommunicationStatus": False}
        enable_api = requests.post(self.webapi, headers=self.headers, data=json.dumps(request1, default=str),
                                   verify=False)
        print(enable_api.headers)
        enable_api_dict = json.loads(enable_api.content)
        print(enable_api_dict)

    def approved_admit_card(self, loop):

        self.lambda_function('Create_Attachment')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- Approved admit card ------------------------------------------------------
        request = {
            "Entity": "Applicant",
            "EntityId": self.xl_applicant_id[loop],
            "Url": self.xl_s3_file[loop],
            "LabelId": self.xl_label_id[loop],
            "LabelOther": "{\"Name\":\"Approved Admitcard\"}"
        }

        approved_api = requests.post(self.webapi, headers=self.headers, data=json.dumps(request, default=str),
                                     verify=False)
        print(approved_api.headers)
        approved_api_dict = json.loads(approved_api.content)
        print(approved_api_dict)
        data = approved_api_dict.get('data')
        self.AttachmentId = data.get('AttachmentId')

        # -------------------------------- set communication status ----------------------------------------------------
        self.lambda_function('setApplicantCommunicationStatus')
        self.headers['APP-NAME'] = 'crpo'

        request1 = {"ApplicantId": self.xl_applicant_id[loop],
                    "CommunicationPurpose": self.xl_CommunicationPurpose[loop],
                    "CommunicationStatus": True}
        communicationstatus_api = requests.post(self.webapi, headers=self.headers,
                                                data=json.dumps(request1, default=str), verify=False)
        print(communicationstatus_api.headers)
        communicationstatus_api_dict = json.loads(communicationstatus_api.content)
        print(communicationstatus_api_dict)

    def score_sheet(self, loop):

        self.lambda_function('Create_Attachment')
        self.headers['APP-NAME'] = 'crpo'

        # ------------------------------------- Score sheet ------------------------------------------------------------
        request = {
            "Entity": "Applicant",
            "EntityId": self.xl_applicant_id[loop],
            "Url": self.xl_s3_file[loop],
            "LabelId": self.xl_label_id[loop],
            "LabelOther": "{\"Name\":\"Score Sheet\"}"
        }

        score_api = requests.post(self.webapi, headers=self.headers,
                                  data=json.dumps(request, default=str), verify=False)
        print(score_api.headers)
        score_api_dict = json.loads(score_api.content)
        print(score_api_dict)
        data = score_api_dict.get('data')
        self.AttachmentId = data.get('AttachmentId')

        # -------------------------------- set communication status ------------------------------------------------
        self.lambda_function('setApplicantCommunicationStatus')
        self.headers['APP-NAME'] = 'crpo'

        request1 = {"ApplicantId": self.xl_applicant_id[loop],
                    "CommunicationPurpose": self.xl_CommunicationPurpose[loop],
                    "CommunicationStatus": True}
        communicationstatus_api = requests.post(self.webapi, headers=self.headers,
                                                data=json.dumps(request1, default=str), verify=False)
        print(communicationstatus_api.headers)
        communicationstatus_api_dict = json.loads(communicationstatus_api.content)
        print(communicationstatus_api_dict)

    def mobile_email_verification(self, loop):

        self.lambda_function('sendVerificationNotification')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- Mobile/Email Verification ------------------------------------------------
        request = {"candidateIds": [self.xl_applicant_id[loop]],
                   "sync": True,
                   "verificationType": self.xl_verification_type[loop],
                   "isForceSend": True}
        mobile_email_api = requests.post(self.webapi, headers=self.headers,
                                         data=json.dumps(request, default=str), verify=False)
        print(mobile_email_api.headers)
        mobile_email_api_dict = json.loads(mobile_email_api.content)
        print(mobile_email_api_dict)

    def email_verification(self, loop):

        self.lambda_function('sendVerificationNotification')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- Email Verification ------------------------------------------------------
        request = {"candidateIds": [self.xl_applicant_id[loop]],
                   "sync": True,
                   "verificationType": self.xl_verification_type[loop],
                   "isForceSend": True}
        email_api = requests.post(self.webapi, headers=self.headers, data=json.dumps(request, default=str),
                                  verify=False)
        print(email_api.headers)
        email_api_dict = json.loads(email_api.content)
        print(email_api_dict)

    def flag(self, loop):

        self.lambda_function('setApplicantCommunicationStatus')
        self.headers['APP-NAME'] = 'crpo'

        # ----------------------------------- flag set for offer pending -----------------------------------------------
        request1 = {"ApplicantId": self.xl_applicant_id[loop],
                    "CommunicationPurpose": self.xl_CommunicationPurpose[loop],
                    "CommunicationStatus": True}
        flag_api = requests.post(self.webapi, headers=self.headers, data=json.dumps(request1, default=str),
                                 verify=False)
        print(flag_api.headers)
        flag_api_dict = json.loads(flag_api.content)
        print(flag_api_dict)

    def fetch_r_link(self, loop):

        self.lambda_function('getRegistrationLinkForApplicants')
        self.headers['APP-NAME'] = 'crpo'

        request = {"ApplicantIds": [self.xl_applicant_id[loop]]}
        get_rl_api = requests.post(self.webapi, headers=self.headers, data=json.dumps(request, default=str),
                                   verify=False)
        print(get_rl_api.headers)
        get_rl_api_dict = json.loads(get_rl_api.content)
        data = get_rl_api_dict['data']
        success = data.get('success')
        app_id = success.get(str(self.xl_applicant_id[loop]))
        self.api_get_rl = app_id.get('RegistrationLink')
        print(self.api_get_rl)

    def ui_automation(self):

        # ------------------------------------------------------
        # UI Automation to handel where ever APIs are not present
        # ------------------------------------------------------
        try:
            self.driver = webdriver.Chrome(input_paths.driver['chrome'])
            print("Run started at:: " + str(datetime.datetime.now()))
            print("Environment setup has been Done")
            print("----------------------------------------------------------")
            self.driver.implicitly_wait(10)
            self.driver.maximize_window()
            self.driver.get(self.api_get_rl)
            time.sleep(5)
            self.driver.find_element_by_id("lbl_terms_yes").click()
            time.sleep(2)
            self.driver.find_element_by_id("declaration").click()
            time.sleep(2)
            self.driver.find_element_by_id("registerbtndiv").click()
            time.sleep(3)

            print("----------------------------------------------------------")
            print("Run completed at:: " + str(datetime.datetime.now()))
            print("Chrome environment Destroyed")
            self.driver.close()

        except exceptions.WebDriverException as Environment_Error:
            print(Environment_Error)

    def re_registration_link(self, loop):

        self.lambda_function('applicantRe-Registration')
        self.headers['APP-NAME'] = 'crpo'

        if self.purpose[loop] == 'Re-Registration Allowed':
            request = {"applicantIds": [self.xl_applicant_id[loop]],
                       "reasonId": 44609,
                       "comment": "Sending registration link again"}

            re_rl_api = requests.post(self.webapi,
                                      headers=self.headers, data=json.dumps(request, default=str), verify=False)
            print(re_rl_api.headers)
            re_rl_api_dict = json.loads(re_rl_api.content)
            print(re_rl_api_dict)

    def get_applicants(self, loop):
        self.lambda_function('getAllEventApplicant')
        self.headers['APP-NAME'] = 'crpo'

        # --------------------------------------
        # Hitting login API based on input value
        # --------------------------------------
        for i in range(self.xl_API_hits[loop]):

            request = {"PagingCriteriaType": {"MaxResults": 100,
                                              "ObjectState": 0, "PageNo": 1, "SortOrder": "1", "SortParameter": "1",
                                              "PageNumber": 1},
                       "RecruitEventId": self.xl_expected_event_id[loop]}
            applicant_api = requests.post(self.webapi, headers=self.headers,
                                          data=json.dumps(request, default=str), verify=False)
            print(applicant_api.headers)
            applicant_api_dict = json.loads(applicant_api.content)
            self.data = applicant_api_dict.get('data')

        for k in self.data:
            if self.xl_expected_app_id[loop] == k['ApplicantId']:
                self.applicant_dict = k
                communication_status_type = k.get('CandidateCommunicationStatusType')
                print(communication_status_type)

                if k.get('CandidateCommunicationStatusType'):
                    for j in communication_status_type:
                        self.communication_statuses = j
                        if self.communication_statuses.get('Admit Card Sent'):
                            self.api_admitcard = self.communication_statuses.get('Admit Card Sent')

                        if self.communication_statuses.get('Registration Link Sent'):
                            self.api_rl = self.communication_statuses.get('Registration Link Sent')

                        if self.communication_statuses.get('Registration Done'):
                            self.api_rl_done = self.communication_statuses.get('Registration Done')

                        if self.communication_statuses.get('GDPR Acceptance Done'):
                            self.api_GDPR = self.communication_statuses.get('GDPR Acceptance Done')

                        if self.communication_statuses.get('Re-Registration Done'):
                            self.api_re_rl_done = self.communication_statuses.get('Re-Registration Done')

                        if self.communication_statuses.get('Re-Registration Allowed'):
                            self.api_re_rl_allowed = self.communication_statuses.get('Re-Registration Allowed')
                        else:
                            self.api_re_rl_Not_allowed = 1

                        if self.communication_statuses.get('Approved Admitcard Uploaded'):
                            self.api_approve_admitcard = self.communication_statuses.get(
                                'Approved Admitcard Uploaded')

                        if self.communication_statuses.get('Score Sheet Uploaded'):
                            self.api_score_sheet = self.communication_statuses.get('Score Sheet Uploaded')

                        if self.communication_statuses.get('Pending For Offer'):
                            self.api_offer = self.communication_statuses.get('Pending For Offer')

                        if self.communication_statuses.get('Email Verification Sent'):
                            self.api_email_v = self.communication_statuses.get('Email Verification Sent')

                        if self.communication_statuses.get('Mobile Verification Sent'):
                            self.api_mobile_v = self.communication_statuses.get('Mobile Verification Sent')

                        if self.communication_statuses.get('Registration Disabled'):
                            self.api_disable_rl = self.communication_statuses.get('Registration Disabled')
                        else:
                            self.api_enable_rl = 1

    def output_report(self, loop):
        # ------------------
        # Writing Input Data
        # ------------------

        if self.communication_statuses:
            self.ws.write(self.rowsize, 0, 'Pass', self.style26)
            self.success_case_01 = 'Pass'
        else:
            self.ws.write(self.rowsize, 0, 'Fail', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        self.ws.write(self.rowsize, 1, self.purpose[loop], self.style12)
        # --------------------------------------------------------------------------------------------------------------

        if self.applicant_dict.get('EventId'):
            self.ws.write(self.rowsize, 2, self.xl_expected_event_id[loop], self.style12)
        else:
            self.ws.write(self.rowsize, 2, self.xl_expected_event_id[loop], self.style3)
        # --------------------------------------------------------------------------------------------------------------

        self.ws.write(self.rowsize, 3, self.xl_expected_app_id[loop], self.style12)
        # --------------------------------------------------------------------------------------------------------------

        if self.applicant_dict.get('ApplicantId'):
            self.ws.write(self.rowsize, 4, self.applicant_dict.get('ApplicantId'), self.style14)
        else:
            self.ws.write(self.rowsize, 4, self.applicant_dict.get('ApplicantId'), self.style3)
        # --------------------------------------------------------------------------------------------------------------

        self.ws.write(self.rowsize, 5, self.xl_expected_admit[loop], self.style12)
        if self.communication_statuses:
            if self.api_admitcard:
                if self.xl_expected_admit[loop] == 'Yes':
                    self.ws.write(self.rowsize, 6, 'Yes', self.style14)
                elif self.xl_expected_admit[loop] == 'No':
                    self.ws.write(self.rowsize, 6, 'No', self.style14)
            elif self.xl_expected_admit[loop] == 'No':
                self.ws.write(self.rowsize, 6, 'No', self.style14)
            else:
                self.ws.write(self.rowsize, 6, 'No', self.style3)
        else:
            self.ws.write(self.rowsize, 6, 'No_com', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        self.ws.write(self.rowsize, 7, self.xl_expected_rl[loop], self.style12)
        if self.communication_statuses:
            if self.api_rl:
                if self.xl_expected_rl[loop] == 'Yes':
                    self.ws.write(self.rowsize, 8, 'Yes', self.style14)
                elif self.xl_expected_rl[loop] == 'No':
                    self.ws.write(self.rowsize, 8, 'No', self.style14)
            elif self.xl_expected_rl[loop] == 'No':
                self.ws.write(self.rowsize, 8, 'No', self.style14)
            else:
                self.ws.write(self.rowsize, 8, 'No', self.style3)
        else:
            self.ws.write(self.rowsize, 8, 'No_com', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        self.ws.write(self.rowsize, 9, self.xl_expected_disable_rl[loop], self.style12)
        if self.communication_statuses:
            if self.api_disable_rl:
                if self.xl_expected_disable_rl[loop] == 'Yes':
                    self.ws.write(self.rowsize, 10, 'Yes', self.style14)
                elif self.xl_expected_disable_rl[loop] == 'No':
                    self.ws.write(self.rowsize, 10, 'No', self.style14)
            elif self.xl_expected_disable_rl[loop] == 'No':
                self.ws.write(self.rowsize, 10, 'No', self.style14)
            else:
                self.ws.write(self.rowsize, 10, 'No', self.style3)
        else:
            self.ws.write(self.rowsize, 10, 'No_com', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        self.ws.write(self.rowsize, 11, self.xl_expected_enable_rl[loop], self.style12)
        if self.communication_statuses:
            if self.api_enable_rl:
                if self.xl_expected_enable_rl[loop] == 'Yes':
                    self.ws.write(self.rowsize, 12, 'Yes', self.style14)
                elif self.xl_expected_enable_rl[loop] == 'No':
                    self.ws.write(self.rowsize, 12, 'No', self.style14)
            elif self.xl_expected_enable_rl[loop] == 'No':
                self.ws.write(self.rowsize, 12, 'No', self.style14)
            else:
                self.ws.write(self.rowsize, 12, 'No', self.style3)
        else:
            self.ws.write(self.rowsize, 12, 'No_com', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        self.ws.write(self.rowsize, 13, self.xl_expected_rl_done[loop], self.style12)
        if self.communication_statuses:
            if self.api_rl_done:
                if self.xl_expected_rl_done[loop] == 'Yes':
                    self.ws.write(self.rowsize, 14, 'Yes', self.style14)
                elif self.xl_expected_rl_done[loop] == 'No':
                    self.ws.write(self.rowsize, 14, 'No', self.style14)
            elif self.xl_expected_rl_done[loop] == 'No':
                self.ws.write(self.rowsize, 14, 'No', self.style14)
            else:
                self.ws.write(self.rowsize, 14, 'No', self.style3)
        else:
            self.ws.write(self.rowsize, 14, 'No_com', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        self.ws.write(self.rowsize, 15, self.xl_expected_re_rl_allowed[loop], self.style12)
        if self.communication_statuses:
            if self.api_re_rl_allowed:
                if self.xl_expected_re_rl_allowed[loop] == 'Yes':
                    self.ws.write(self.rowsize, 16, 'Yes', self.style14)
                elif self.xl_expected_re_rl_allowed[loop] == 'No':
                    self.ws.write(self.rowsize, 16, 'No', self.style14)
            elif self.xl_expected_re_rl_allowed[loop] == 'No':
                self.ws.write(self.rowsize, 16, 'No', self.style14)
            else:
                self.ws.write(self.rowsize, 16, 'No', self.style3)
        else:
            self.ws.write(self.rowsize, 16, 'No_com', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        self.ws.write(self.rowsize, 17, self.xl_expected_re_rl_done[loop], self.style12)
        if self.communication_statuses:
            if self.api_re_rl_done:
                if self.xl_expected_re_rl_done[loop] == 'Yes':
                    self.ws.write(self.rowsize, 18, 'Yes', self.style14)
                elif self.xl_expected_re_rl_done[loop] == 'No':
                    self.ws.write(self.rowsize, 18, 'No', self.style14)
            elif self.xl_expected_re_rl_done[loop] == 'No':
                self.ws.write(self.rowsize, 18, 'No', self.style14)
            else:
                self.ws.write(self.rowsize, 18, 'No', self.style3)
        else:
            self.ws.write(self.rowsize, 18, 'No_com', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        self.ws.write(self.rowsize, 19, self.xl_expected_re_rl_not_allowed[loop], self.style12)
        if self.communication_statuses:
            if self.api_re_rl_Not_allowed:
                if self.xl_expected_re_rl_not_allowed[loop] == 'Yes':
                    self.ws.write(self.rowsize, 20, 'Yes', self.style14)
                elif self.xl_expected_re_rl_not_allowed[loop] == 'No':
                    self.ws.write(self.rowsize, 20, 'No', self.style14)
            elif self.xl_expected_re_rl_not_allowed[loop] == 'No':
                self.ws.write(self.rowsize, 20, 'No', self.style14)
            else:
                self.ws.write(self.rowsize, 20, 'No', self.style3)
        else:
            self.ws.write(self.rowsize, 20, 'No_com', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        self.ws.write(self.rowsize, 21, self.xl_expected_GDPR[loop], self.style12)
        if self.communication_statuses:
            if self.api_GDPR:
                if self.xl_expected_GDPR[loop] == 'Yes':
                    self.ws.write(self.rowsize, 22, 'Yes', self.style14)
                elif self.xl_expected_GDPR[loop] == 'No':
                    self.ws.write(self.rowsize, 22, 'No', self.style14)
            elif self.xl_expected_GDPR[loop] == 'No':
                self.ws.write(self.rowsize, 22, 'No', self.style14)
            else:
                self.ws.write(self.rowsize, 22, 'No', self.style3)
        else:
            self.ws.write(self.rowsize, 22, 'No_com', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        self.ws.write(self.rowsize, 23, self.xl_expected_approved_admit[loop], self.style12)
        if self.communication_statuses:
            if self.api_approve_admitcard:
                if self.xl_expected_approved_admit[loop] == 'Yes':
                    self.ws.write(self.rowsize, 24, 'Yes', self.style14)
                elif self.xl_expected_approved_admit[loop] == 'No':
                    self.ws.write(self.rowsize, 24, 'No', self.style14)
            elif self.xl_expected_approved_admit[loop] == 'No':
                self.ws.write(self.rowsize, 24, 'No', self.style14)
            else:
                self.ws.write(self.rowsize, 24, 'No', self.style3)
        else:
            self.ws.write(self.rowsize, 24, 'No_com', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        self.ws.write(self.rowsize, 25, self.xl_expected_score[loop], self.style12)
        if self.communication_statuses:
            if self.api_score_sheet:
                if self.xl_expected_score[loop] == 'Yes':
                    self.ws.write(self.rowsize, 26, 'Yes', self.style14)
                elif self.xl_expected_score[loop] == 'No':
                    self.ws.write(self.rowsize, 26, 'No', self.style14)
            elif self.xl_expected_score[loop] == 'No':
                self.ws.write(self.rowsize, 26, 'No', self.style14)
            else:
                self.ws.write(self.rowsize, 26, 'No', self.style3)
        else:
            self.ws.write(self.rowsize, 26, 'No_com', self.style3)
            # --------------------------------------------------------------------------------------------------------------

        self.ws.write(self.rowsize, 27, self.xl_expected_offer[loop], self.style12)
        if self.communication_statuses:
            if self.api_offer:
                if self.xl_expected_offer[loop] == 'Yes':
                    self.ws.write(self.rowsize, 28, 'Yes', self.style14)
                elif self.xl_expected_offer[loop] == 'No':
                    self.ws.write(self.rowsize, 28, 'No', self.style14)
            elif self.xl_expected_offer[loop] == 'No':
                self.ws.write(self.rowsize, 28, 'No', self.style14)
            else:
                self.ws.write(self.rowsize, 28, 'No', self.style3)
        else:
            self.ws.write(self.rowsize, 28, 'No_com', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        self.ws.write(self.rowsize, 29, self.xl_expected_email_mobile[loop], self.style12)
        if self.communication_statuses:
            if self.api_email_v and self.api_mobile_v:
                if self.xl_expected_email_mobile[loop] == 'Yes':
                    self.ws.write(self.rowsize, 30, 'Yes', self.style14)
                elif self.xl_expected_email_mobile[loop] == 'No':
                    self.ws.write(self.rowsize, 30, 'No', self.style14)
            elif self.xl_expected_email_mobile[loop] == 'No':
                self.ws.write(self.rowsize, 30, 'No', self.style14)
            else:
                self.ws.write(self.rowsize, 30, 'No', self.style3)
        else:
            self.ws.write(self.rowsize, 30, 'No_com', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        self.ws.write(self.rowsize, 31, self.xl_expected_email[loop], self.style12)
        if self.communication_statuses:
            if self.api_email_v:
                if self.xl_expected_email[loop] == 'Yes':
                    self.ws.write(self.rowsize, 32, 'Yes', self.style14)
                elif self.xl_expected_email[loop] == 'No':
                    self.ws.write(self.rowsize, 32, 'No', self.style14)
            elif self.xl_expected_email[loop] == 'No':
                self.ws.write(self.rowsize, 32, 'No', self.style14)
            else:
                self.ws.write(self.rowsize, 32, 'No', self.style3)
        else:
            self.ws.write(self.rowsize, 32, 'No_com', self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.AttachmentId:
            if self.purpose[loop] == 'ApprovedAdmitcard':
                self.ws.write(self.rowsize, 33, self.AttachmentId, self.style12)
            elif self.purpose[loop] == 'ScoreSheet':
                self.ws.write(self.rowsize, 33, self.AttachmentId, self.style12)

        self.rowsize += 1  # Row increment
        self.col += 1  # Column increment
        Object.wb_Result.save(output_paths.outputpaths['Communication_output_sheet'])

        if self.success_case_01 == 'Pass':
            self.Actual_Success_case.append(self.success_case_01)

    def overall_status(self):
        self.ws.write(0, 0, 'Communication status', self.style23)
        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
        else:
            self.ws.write(0, 1, 'Fail', self.style25)
        self.ws.write(0, 2, 'Start Time', self.style23)
        self.ws.write(0, 3, self.start_time, self.style26)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)
        Object.wb_Result.save(output_paths.outputpaths['Communication_output_sheet'])


Object = CommunicationHistory()
Object.excel_data()
Total_count = len(Object.purpose)
print("Number of Rows ::", Total_count)
if Object.login == 'OK':
    for looping in range(0, Total_count):
        print("Iteration Count is ::", looping)

        if Object.purpose[looping] == 'AdmitCard':
            Object.admit_card(looping)
        if Object.purpose[looping] == 'RL':
            Object.registration_link(looping)
        if Object.purpose[looping] == 'RL-Disable':
            Object.registration_link_disable(looping)
        if Object.purpose[looping] == 'RL-Enable':
            Object.registration_link_enable(looping)
        if Object.purpose[looping] == 'RL-Done':
            Object.registration_link(looping)
            Object.fetch_r_link(looping)
            Object.ui_automation()
            time.sleep(5)

        if Object.purpose[looping] == 'Re-Registration Allowed':
            Object.registration_link(looping)
            Object.fetch_r_link(looping)
            Object.ui_automation()
            Object.re_registration_link(looping)

        if Object.purpose[looping] == 'Re-Registration Done':
            Object.registration_link(looping)
            Object.fetch_r_link(looping)
            Object.ui_automation()
            Object.re_registration_link(looping)
            Object.ui_automation()
            time.sleep(5)

        if Object.purpose[looping] == 'Re-Registration NotAllowed':
            try:
                Object.registration_link(looping)
                Object.fetch_r_link(looping)
                Object.ui_automation()
                Object.re_registration_link(looping)
                Object.fetch_r_link(looping)
                Object.ui_automation()
            except exceptions.WebDriverException as fail:
                print(fail)

        if Object.purpose[looping] == 'GDPR Acceptance Done':
            Object.registration_link(looping)
            Object.fetch_r_link(looping)
            Object.ui_automation()

        if Object.purpose[looping] == 'ApprovedAdmitcard':
            Object.approved_admit_card(looping)
        if Object.purpose[looping] == 'ScoreSheet':
            Object.score_sheet(looping)
        if Object.purpose[looping] == 'Mobile/Email Verification':
            Object.mobile_email_verification(looping)
        if Object.purpose[looping] == 'Email Verification':
            Object.email_verification(looping)
        if Object.purpose[looping] == 'Flag':
            Object.flag(looping)

        time.sleep(4)
        Object.get_applicants(looping)
        Object.output_report(looping)

        # -----------------
        # Making Dict empty
        # -----------------
        Object.applicant_dict = {}
        Object.communication_statuses = {}
        Object.api_admitcard = {}
        Object.api_rl = {}
        Object.api_approve_admitcard = {}
        Object.api_score_sheet = {}
        Object.api_offer = {}
        Object.api_email_v = {}
        Object.api_mobile_v = {}
        Object.api_disable_rl = {}
        Object.api_enable_rl = {}

        Object.success_case_01 = {}
        Object.data = {}
        Object.AttachmentId = {}
        Object.api_get_rl = {}
        Object.api_rl_done = {}
        Object.api_re_rl_allowed = {}
        Object.api_re_rl_done = {}
        Object.api_re_rl_Not_allowed = {}
        Object.api_GDPR = {}
        Object.headers = {}

Object.overall_status()
