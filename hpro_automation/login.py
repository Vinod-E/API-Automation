import urllib3
import datetime
import json
import time
import requests
from hpro_automation.api import *


class CommonLogin(object):

    def __init__(self):
        super(CommonLogin, self).__init__()
        # self.app_name = input("APP-NAME: crpo or pyappe1 or py3app:: ")
        self.app_name = "crpo"
        # self.calling_lambda = str(input("Lambda On/Off:: "))
        self.calling_lambda = "On"
        self.lambda_headers = {"content-type": "application/json", 'X-APPLMA': 'true'}
        self.Non_lambda_headers = {"content-type": "application/json"}
        self.header = {"content-type": "application/json", 'APP-NAME': "crpo", 'X-APPLMA': 'true'}
        self.get_token = ""
        self.integrationGuid = ""
        self.login = ""
        self.webapi = ""
        self.headers = {}

    def common_login(self, login_user):
        # -------------------------------- CRPO LOGIN APPLICATION ------------------------------------------------------

        print("-------------------------------------------------")
        print("Run Started at :", str(datetime.datetime.now()))
        print("-------------------------------------------------")
        print('That you are running in this server :: ', generic_domain)
        print("-------------------------------------------------")

        try:
            urllib3.disable_warnings()
            if login_server == 'amsin':
                if login_user == 'admin':
                    login_data = credentials.login_details['crpo']
                elif login_user == 'slot':
                    login_data = credentials.login_details['amsin_slot']
                else:
                    login_data = credentials.login_details['int']

            else:
                if login_user == 'admin':
                    login_data = credentials.login_details['ams_crpo']
                elif login_user == 'slot':
                    login_data = credentials.login_details['ams_slot']
                else:
                    login_data = credentials.login_details['ams_int']

            self.headers['APP-NAME'] = self.app_name
            self.headers['X-APPLMA'] = 'true'
            login_api = requests.post(lambda_apis.get("Loginto_CRPO"), headers=self.header, data=json.dumps(login_data),
                                      verify=False)
            print(login_data)
            response = login_api.json()
            print(response)
            print(login_api.headers)
            self.get_token = response.get("Token")

            time.sleep(1)
            resp_dict = json.loads(login_api.content)
            status = resp_dict['status']
            if status == 'OK':
                self.login = 'OK'
                print("CRPO Login successfully")
            else:
                self.login = 'KO'
                print("CRPO Login Failed")
        except ValueError as login_error:
            print(login_error)
        self.lambda_check()

    def performance_login(self, login_user):
        # -------------------------------- CRPO LOGIN APPLICATION ------------------------------------------------------

        print("-------------------------------------------------")
        print("Run Started at :", str(datetime.datetime.now()))
        print("-------------------------------------------------")

        try:
            urllib3.disable_warnings()
            if login_server == 'amsin':
                if login_user == 'amsin_eu':
                    print('That you are running in this server :: ', eu_amsin_domain)
                    print("-------------------------------------------------")
                    login_data = credentials.login_details['amsin_eu']
                    api = lambda_apis.get("eu_amsin_login")
                    print(api)
                else:
                    print('That you are running in this server :: ', generic_domain)
                    print("-------------------------------------------------")
                    login_data = credentials.login_details['amsin_non_eu']
                    api = lambda_apis.get("Loginto_CRPO")
                    print(api)

            else:
                if login_user == 'live_eu':
                    print('That you are running in this server :: ', eu_ams_domain)
                    print("-------------------------------------------------")
                    login_data = credentials.login_details['live_eu']
                    api = lambda_apis.get("eu_ams_login")
                    print(api)
                else:
                    print('That you are running in this server :: ', generic_domain)
                    print("-------------------------------------------------")
                    login_data = credentials.login_details['live_non_eu']
                    api = lambda_apis.get("Loginto_CRPO")
                    print(api)

            self.headers['APP-NAME'] = self.app_name
            self.headers['X-APPLMA'] = 'true'
            login_api = requests.post(api, headers=self.header, data=json.dumps(login_data),
                                      verify=False)
            print(login_data)
            response = login_api.json()
            print(login_api.headers)
            self.get_token = response.get("Token")

            time.sleep(1)
            resp_dict = json.loads(login_api.content)
            status = resp_dict['status']
            if status == 'OK':
                self.login = 'OK'
                print("CRPO Login successfully")
            else:
                self.login = 'KO'
                print("CRPO Login Failed")
        except ValueError as login_error:
            print(login_error)

    def slot_captcha_login_token(self, login_user):
        # -------------------------------- CRPO LOGIN APPLICATION ------------------------------------------------------

        print("-------------------------------------------------")
        print("Run Started at :", str(datetime.datetime.now()))

        try:
            urllib3.disable_warnings()
            if login_server == 'amsin':
                if login_user == 'verify':
                    oauth_data = credentials.login_details['amsin_choose_slot']
                    self.integrationGuid = credentials.amsin_verify_slot_guid
                    self.headers['APP-NAME'] = "verify"
                elif login_user == 'assessment':
                    oauth_data = credentials.login_details['amsin_assessment_slot']
                    self.integrationGuid = credentials.amsin_assessment_slot_guid
                    self.headers['APP-NAME'] = "assessmentSlots"
                else:
                    oauth_data = credentials.login_details['amsin_interview_slot']
                    self.integrationGuid = credentials.amsin_interview_slot_guid

            else:
                if login_user == 'verify':
                    oauth_data = credentials.login_details['ams_choose_slot']
                    self.integrationGuid = credentials.ams_verify_slot_guid
                    self.headers['APP-NAME'] = "verify"
                elif login_user == 'assessment':
                    oauth_data = credentials.login_details['ams_assessment_slot']
                    self.integrationGuid = credentials.ams_assessment_slot_guid
                    self.headers['APP-NAME'] = "assessmentSlots"
                else:
                    oauth_data = credentials.login_details['ams_interview_slot']
                    self.integrationGuid = credentials.ams_interview_slot_guid

            # ------------------ API Call -------------------------------------
            oauth_api = requests.post(slot_app.get("access_token").format(self.integrationGuid), headers=self.headers,
                                      data=json.dumps(oauth_data), verify=False)
            response = oauth_api.json()
            # print(oauth_api.headers)
            print(response)
            self.get_token = response.get("access_token")

            time.sleep(1)
            token_type = response.get("token_type")
            if token_type == 'bearer':
                self.login = 'OK'
                print("Oauth Token Generated Successfully")
            else:
                self.login = 'KO'
                print("Oauth Token Generation Failed")
        except ValueError as oauth_error:
            print(oauth_error)

    def verify_hash(self, request):
        try:
            urllib3.disable_warnings()
            self.lambda_headers['Authorization'] = 'bearer ' + self.get_token
            self.lambda_headers['App-Name'] = 'assessmentSlots'
            request = json.loads(request)

            # ------------------ API Call -------------------------------------
            verify_hash_api = requests.post(lambda_apis.get("verfiyhash"), headers=self.lambda_headers,
                                            data=json.dumps(request), verify=False)
            response = verify_hash_api.json()
            data = response.get('data')
            if data.get('message') == 'Authorized.':
                print("slot captcha login api token:: ", data['message'])
            else:
                print(response)
        except ValueError as oauth_error:
            print(oauth_error)

    def authenticate(self, request):
        try:
            urllib3.disable_warnings()
            self.Non_lambda_headers['Authorization'] = 'bearer ' + self.get_token
            request = json.loads(request)

            # ------------------ API Call -------------------------------------
            authenticate_api = requests.post(non_lambda_apis.get("authenticate"), headers=self.Non_lambda_headers,
                                             data=json.dumps(request), verify=False)
            response = authenticate_api.json()
            data = response.get('data')
            if data.get('message') == 'Authorized.':
                print("slot captcha login api token:: ", data['message'])
                self.Non_lambda_headers['Authorization'] = 'bearer ' + data.get('token')
            else:
                print(response)
        except ValueError as oauth_error:
            print(oauth_error)

    def lambda_check(self):
        # ------------------------------- getAllAppPreference / Lambda verification ------------------------------------
        try:
            self.header['X-AUTH-TOKEN'] = self.get_token
            request = {"pagingCriteria": {"pageSize": 1000,
                                          "pageNumber": 1,
                                          "sortOn": "id",
                                          "sortBy": "desc"}
                       }

            app_preference_api = requests.post(lambda_apis['getAllAppPreference'], headers=self.header,
                                               data=json.dumps(request), verify=False)
            res = json.loads(app_preference_api.content)
            data = res.get('data')
            for i in data:
                app_preference = i.get('typeText')
                if app_preference == 'crpo.tenantConfigurations':
                    content_text = i.get('contentText')
                    is_lambda = content_text.get('isLambdaRequired')
                    if is_lambda:
                        print("**----------------------Lambda is enabled in the tenant---------------------------**")
                    else:
                        print("**----------------------Lambda is disabled in the tenant--------------------------**")
        except ValueError as app:
            print(app)

    def lambda_function(self, api_key):

        # ---------------------- Passing headers based on API supports to lambda or not --------------------------------
        if self.calling_lambda.lower() == 'on':
            if lambda_apis.get(api_key) is not None:
                self.headers = self.lambda_headers
                self.headers['X-AUTH-TOKEN'] = self.get_token
                self.webapi = lambda_apis[api_key]
            else:
                self.headers = self.Non_lambda_headers
                self.headers['X-AUTH-TOKEN'] = self.get_token
                self.webapi = non_lambda_apis[api_key]

        elif self.calling_lambda.lower() == 'off':
            self.headers = self.Non_lambda_headers
            self.headers['X-AUTH-TOKEN'] = self.get_token
            if non_lambda_apis.get(api_key) is not None:
                self.webapi = non_lambda_apis[api_key]
            else:
                self.webapi = lambda_apis[api_key]

        else:
            if lambda_apis.get(api_key) is not None:
                self.headers = self.lambda_headers
                self.headers['X-AUTH-TOKEN'] = self.get_token
                self.webapi = lambda_apis[api_key]
            else:
                self.headers = self.Non_lambda_headers
                self.headers['X-AUTH-TOKEN'] = self.get_token
                self.webapi = non_lambda_apis[api_key]
