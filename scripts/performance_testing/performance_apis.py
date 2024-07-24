import json
import time
import requests
from datetime import date
from hpro_automation import (login)
from scripts.performance_testing import Performance_excel


class PerformanceTesting(login.CommonLogin, Performance_excel.PerformanceExcel):
    def __init__(self):
        t = time.localtime()
        self.run_date = str(date.today())
        self.run_time = time.strftime("%I:%M:%p", t)
        super(PerformanceTesting, self).__init__()
        self.response = ''
        self.request = {}
        self.Average_Time = ''
        self.Average_Time_tenant_details = ''
        self.Average_Time_entity = ''
        self.Average_Time_catalog = ''
        self.Average_Time_candidates = ''
        self.Average_Time_testuser = ''
        self.Average_Time_interview = ''
        self.Average_Time_new_interview = ''

        self.excel_read_by_index()

    def get_response_time_api(self, api):
        self.lambda_function(api)

        Iterate_time = 0
        for i in range(0, 5):
            self.headers['APP-NAME'] = self.app_name
            response_time_api = requests.post(self.webapi, headers=self.headers,
                                              data=json.dumps(self.request, default=str),
                                              verify=False)
            response_time = response_time_api.elapsed.total_seconds()
            self.response = json.loads(response_time_api.content)
            print(response_time_api.headers)
            # print(self.response)

            # print('Response Time ::', response_time)
            Iterate_time = Iterate_time + response_time
            # print(Iterate_time)
            time.sleep(1)
        self.Average_Time = Iterate_time/5
        print("Average_Time :: ", self.Average_Time)

    def get_tenant_details(self, api):
        self.headers['X-AUTH-TOKEN'] = None
        self.request = {"TenantAlias": self.tenant_name}
        self.get_response_time_api(api)
        self.Average_Time_tenant_details = self.Average_Time
        print('API is:: get_tenant_details')
        # print(self.response)

    def get_all_entity_properties(self, api):
        self.request = {"EntityType": "2"}
        self.get_response_time_api(api)
        self.Average_Time_entity = self.Average_Time
        print('API is:: get_all_entity_properties')
        # print(self.response)

    def group_by_catalog_masters(self, api):
        self.request = {"catalogMasterNames": ["GetCandidateGridProperties"]}
        self.get_response_time_api(api)
        self.Average_Time_catalog = self.Average_Time
        print('API is:: group_by_catalog_masters')
        # print(self.response)

    def get_all_candidates(self, api):
        self.request = {"PagingCriteria": {"IsRefresh": False, "IsSpecificToUser": False, "MaxResults": 20, "PageNo": 1,
                                           "SortParameter": "0", "SortOrder": "0",
                                           "PropertyIds": self.property_ids,
                                           "ObjectState": 0,
                                           "IsCountRequired": False
                                           }, "OrderBy": {"FieldName": "Default", "Order": "desc"},
                        "IsNotCacheRequired": True
                        }
        self.get_response_time_api(api)
        self.Average_Time_candidates = self.Average_Time
        print('API is:: get_all_candidates')
        # print(self.response)

    def getTestUsersForTest(self, api):
        self.request = {"isProctroingInfo": True, "isPartnerTestUserInfo": True, "testId": self.test_id,
                        "paging": {"maxResults": 20, "pageNumber": 1}}
        self.get_response_time_api(api)
        self.Average_Time_testuser = self.Average_Time
        print('API is:: getTestUsersForTest')
        # print(self.response)

    def interviews(self, api):
        self.request = {"pagingCriteria": {"pageSize": 20, "pageNumber": 1}, "isRecordedFileUrlsRequired": False,
                        "isAllInterviewRequired": True, "status": 0, "search": {}}
        self.get_response_time_api(api)
        self.Average_Time_interview = self.Average_Time
        print('API is:: interviews')
        # print(self.response)

    def new_interviews(self, api):
        self.request = {"pagingCriteria": {"pageSize": 20, "pageNumber": 1}, "isRecordedFileUrlsRequired": False,
                        "isAllInterviewRequired": True, "status": 0, "search": {}}
        self.get_response_time_api(api)
        self.Average_Time_new_interview = self.Average_Time
        print('API is:: new_interviews')
        # print(self.response)
