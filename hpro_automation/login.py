import urllib3
import datetime
import json
import time
import requests
from hpro_automation.identity import credentials
from hpro_automation.api import *


class CRPOLogin(object):

    def __init__(self):
        super(CRPOLogin, self).__init__()

        # ------------------------
        # CRPO LOGIN APPLICATION
        # ------------------------

        print("-------------------------------------------------")
        print("Run Started at :", str(datetime.datetime.now()))

        try:
            urllib3.disable_warnings()
            self.header = {"content-type": "application/json", 'APP-NAME': "crpo", 'X-APPLMA': 'true'}
            self.type_of_user = str(input("Type of User/application_name:: "))
            self.calling_lambda = str(input("Lambda On/Off:: "))

            login_data = credentials.login_details[self.type_of_user]

            login_api = requests.post(web_api.get("Loginto_CRPO"),
                                      headers=self.header,
                                      data=json.dumps(login_data),
                                      verify=False)
            self.response = login_api.json()
            self.get_token = {"content-type": "application/json",
                              "X-AUTH-TOKEN": self.response.get("Token")}

            self.header = {"content-type": "application/json",
                           'APP-NAME': "crpo",
                           'X-APPLMA': 'true',
                           "X-AUTH-TOKEN": self.response.get("Token")
                           }

            self.var = None
            time.sleep(1)
            resp_dict = json.loads(login_api.content)
            self.status = resp_dict['status']
            if self.status == 'OK':
                self.login = 'OK'
                print("Login successfully")
                print("Status is", self.status)
                time.sleep(1)
            else:
                self.login = 'KO'
                print("Failed to login")
                print("Status is", self.status)
        except ValueError as login_error:
            print(login_error)

    def lambda_api(self):
        request = {"pagingCriteria": {"pageSize": 1000,
                                      "pageNumber": 1,
                                      "sortOn": "id",
                                      "sortBy": "desc"}
                   }

        api = requests.post("https://amsin.hirepro.in/py/common/common_app_utils/api/v1/getAllAppPreference/",
                            headers=self.header, data=json.dumps(request), verify=False)

        res = json.loads(api.content)
        data = res.get('data')
        for i in data:
            app_preference = i.get('typeText')

            if app_preference == 'crpo.tenantConfigurations':
                content_text = i.get('contentText')
                is_lambda = content_text.get('isLambdaRequired')

                if is_lambda:
                    print("**----------------------Lambda is enable in tenant---------------------------**")
                else:
                    print("**----------------------Lambda is disable in tenant--------------------------**")


ob = CRPOLogin()
ob.lambda_api()
