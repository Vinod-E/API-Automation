import urllib3
import datetime
import json
import time
import requests
from hpro_automation.api import *


class CommonLogin(object):

    def __init__(self):
        super(CommonLogin, self).__init__()
        self.calling_lambda = str(input("Lambda On/Off:: "))
        self.lambda_headers = {"content-type": "application/json", 'X-APPLMA': 'true'}
        self.Non_lambda_headers = {"content-type": "application/json"}
        self.header = {"content-type": "application/json", 'APP-NAME': "crpo", 'X-APPLMA': 'true'}
        self.get_token = ""
        self.login = ""
        self.webapi = ""
        self.headers = {}

    def common_login(self, login_user):
        # -------------------------------- CRPO LOGIN APPLICATION ------------------------------------------------------

        print("-------------------------------------------------")
        print("Run Started at :", str(datetime.datetime.now()))

        try:
            urllib3.disable_warnings()
            login_data = credentials.login_details[login_user]
            login_api = requests.post(lambda_apis.get("Loginto_CRPO"), headers=self.header, data=json.dumps(login_data),
                                      verify=False)
            response = login_api.json()
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
        if self.calling_lambda == 'On' or self.calling_lambda == 'on':
            if lambda_apis.get(api_key) is not None:
                self.headers = self.lambda_headers
                self.headers['X-AUTH-TOKEN'] = self.get_token
                self.webapi = lambda_apis[api_key]
            else:
                self.headers = self.Non_lambda_headers
                self.headers['X-AUTH-TOKEN'] = self.get_token
                self.webapi = non_lambda_apis[api_key]

        elif self.calling_lambda == 'Off' or self.calling_lambda == 'off':
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
