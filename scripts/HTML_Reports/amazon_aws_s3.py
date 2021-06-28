import json
import requests
import urllib3
from hpro_automation.login import CommonLogin
from hpro_automation import api


class AWS:
    def __init__(self, save_file_name, path):
        urllib3.disable_warnings()
        self.__html_path = path
        self.__file_name = save_file_name
        self.one_day_link = ''
        self.login = CommonLogin()

    def file_handler(self):
        """
        ===================>> Login for Token <<======================
        """
        self.login.common_login('crpo')
        url = api.non_lambda_apis['s3']
        headers = {
            'X-AUTH-TOKEN': self.login.get_token,
        }
        with open(self.__html_path, "rb") as a_file:
            file_context = a_file.read()
            if 'html' in self.__file_name:
                file_dict = {'{}_.html'.format(self.__file_name): file_context}
                response = requests.post(url, headers=headers, files=file_dict)
                # print(response.text)
            elif 'xls' in self.__file_name:
                file_dict = {'{}_.xls'.format(self.__file_name): file_context}
                response = requests.post(url, headers=headers, files=file_dict)
                # print(response.text)
        res = json.loads(response.content)
        data = res.get('data')
        self.one_day_link = data.get('fileUrl')
        print('Expire In One Day:: ', self.one_day_link)
