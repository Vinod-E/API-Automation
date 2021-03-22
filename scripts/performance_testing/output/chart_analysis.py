import time
import datetime
import pandas as pd
from pandas import ExcelFile
from selenium import webdriver
from selenium.common import exceptions
from selenium.webdriver.common.keys import Keys
from hpro_automation import (output_paths, input_paths)


class Chart(object):
    def __init__(self):
        self.output_file = output_paths.outputpaths['performance_testing']
        self.chart_file = output_paths.outputpaths['chart_analysis']

        self.value_dict = {}
        self.frame1 = {}
        self.frame2 = {}
        self.frame3 = {}
        self.frame4 = {}

        self.get_tenant_details_dict = {}
        self.get_all_entity_properties_dict = {}
        self.group_by_catalog_masters_dict = {}
        self.get_all_candidates_dict = {}
        self.getTestUsersForTest_dict = {}
        self.interview_dict = {}
        self.new_interview_dict = {}

        self.summation = 0
        self.get_tenant_details_sum = 0
        self.get_all_entity_properties_sum = 0
        self.group_by_catalog_masters_sum = 0
        self.get_all_candidates_sum = 0
        self.getTestUsersForTest_sum = 0
        self.interview_sum = 0
        self.new_interview_sum = 0

    def read_data_from_excel(self, sheet_name):
        xlsx = ExcelFile(self.output_file)
        df = xlsx.parse(sheet_name)
        dict_sheet_values = df.to_dict()
        # print(dict_sheet_values)

        keys = []
        for i in dict_sheet_values:
            keys.append(i)
        print(keys)

        for j in keys:
            self.value_dict = dict_sheet_values.get(j)
            self.get_tenant_details_dict = dict_sheet_values.get('get_tenant_details')
            self.get_all_entity_properties_dict = dict_sheet_values.get('get_all_entity_properties')
            self.group_by_catalog_masters_dict = dict_sheet_values.get('group_by_catalog_masters')
            self.get_all_candidates_dict = dict_sheet_values.get('get_all_candidates')
            self.getTestUsersForTest_dict = dict_sheet_values.get('getTestUsersForTest')
            self.interview_dict = dict_sheet_values.get('interviews')
            self.new_interview_dict = dict_sheet_values.get('interview_new')
        print(self.get_tenant_details_dict)
        print(len(self.get_tenant_details_dict))

        self.summation_data(self.get_tenant_details_dict)
        self.get_tenant_details_sum = self.summation / len(self.get_tenant_details_dict)

        self.summation_data(self.get_all_entity_properties_dict)
        self.get_all_entity_properties_sum = self.summation / len(self.get_all_entity_properties_dict)

        self.summation_data(self.group_by_catalog_masters_dict)
        self.group_by_catalog_masters_sum = self.summation / len(self.group_by_catalog_masters_dict)

        self.summation_data(self.get_all_candidates_dict)
        self.get_all_candidates_sum = self.summation / len(self.get_all_candidates_dict)

        self.summation_data(self.getTestUsersForTest_dict)
        self.getTestUsersForTest_sum = self.summation / len(self.getTestUsersForTest_dict)

        # self.summation_data(self.interview_dict)
        # self.interview_sum = self.summation / len(self.interview_dict)
        #
        # self.summation_data(self.new_interview_dict)
        # self.new_interview_sum = self.summation / len(self.new_interview_dict)

        if sheet_name == 'AMSIN_NON_EU':
            self.frame1 = {'get_tenant_details': self.get_tenant_details_sum,
                           'get_all_entity_properties': self.get_all_entity_properties_sum,
                           'group_by_catalog_masters': self.group_by_catalog_masters_sum,
                           'get_all_candidates': self.get_all_candidates_sum,
                           'getTestUsersForTest': self.getTestUsersForTest_sum,
}
            print(self.frame1)

        if sheet_name == 'AMSIN_EU':
            self.frame2 = {'get_tenant_details': self.get_tenant_details_sum,
                           'get_all_entity_properties': self.get_all_entity_properties_sum,
                           'group_by_catalog_masters': self.group_by_catalog_masters_sum,
                           'get_all_candidates': self.get_all_candidates_sum,
                           'getTestUsersForTest': self.getTestUsersForTest_sum,
}
            print(self.frame2)

        if sheet_name == 'LIVE_NON_EU':
            self.frame3 = {'get_tenant_details': self.get_tenant_details_sum,
                           'get_all_entity_properties': self.get_all_entity_properties_sum,
                           'group_by_catalog_masters': self.group_by_catalog_masters_sum,
                           'get_all_candidates': self.get_all_candidates_sum,
                           'getTestUsersForTest': self.getTestUsersForTest_sum,
}
            print(self.frame3)

        if sheet_name == 'LIVE_EU':
            self.frame4 = {'get_tenant_details': self.get_tenant_details_sum,
                           'get_all_entity_properties': self.get_all_entity_properties_sum,
                           'group_by_catalog_masters': self.group_by_catalog_masters_sum,
                           'get_all_candidates': self.get_all_candidates_sum,
                           'getTestUsersForTest': self.getTestUsersForTest_sum,
}
            print(self.frame4)

    def summation_data(self, api_dict_time):
        item = 0
        self.summation = 0
        for k in range(0, len(api_dict_time)):
            if str(api_dict_time.get(item)) != 'nan':
                self.summation = self.summation + api_dict_time.get(item)
            item += 1
        print(self.summation)

    def chart_sheets(self, sheet_name):

        data = [self.frame1, self.frame2, self.frame3, self.frame4]
        index = ['AMSIN_NON_EU', 'AMSIN_EU', 'LIVE_NON_EU', 'LIVE_EU']
        df = pd.DataFrame(data, index=index)
        writer = pd.ExcelWriter(self.chart_file, engine='xlsxwriter')
        df.to_excel(writer, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        a = len(self.frame4)
        chart = workbook.add_chart({'type': 'column'})
        for col_num in range(1, a + 1):
            chart.add_series({
                'name': [sheet_name, 0, col_num],
                'categories': [sheet_name, 1, 0, 4, 0],
                'values': [sheet_name, 1, col_num, 4, col_num],
                'overlap': -10,
                'line': {'color': 'yellow'},
                'overlay': True,
            })

        chart.set_x_axis({'name': 'Server(s)_Tenant(s)'})
        chart.set_y_axis({'name': 'Time(seconds)', 'major_gridlines': {'visible': False}})

        worksheet.insert_chart('A7', chart)
        writer.save()

    def merge_2_excels(self):
        # -------- Opening Online website to merge excels -----------------
        try:
            driver = webdriver.Chrome(input_paths.driver['chrome'])
            print("Run started at:: " + str(datetime.datetime.now()))
            print("Environment setup has been Done")
            print("----------------------------------------------------------")
            driver.implicitly_wait(10)
            driver.maximize_window()
            driver.get('https://products.aspose.app/cells/merger')
            time.sleep(5)
            driver.find_element_by_xpath('//*[@id="select2-saveAs-container"]').click()
            time.sleep(2)
            xlsx = driver.find_element_by_xpath('//*[@type="search"]')
            xlsx.send_keys('xlsx', Keys.ENTER)
            driver.find_element_by_xpath('//*[@type="file"]').send_keys(self.chart_file)
            time.sleep(2)
            driver.find_element_by_xpath('//*[@type="file"]').send_keys(self.output_file)
            time.sleep(2)
            driver.find_element_by_id("uploadButton").click()
            time.sleep(3)
            driver.find_element_by_id("DownloadButton").click()
            time.sleep(5)

            print("----------------------------------------------------------")
            print("Run completed at:: " + str(datetime.datetime.now()))
            print("Chrome environment Destroyed")
            driver.close()

        except exceptions.WebDriverException as Environment_Error:
            print(Environment_Error)
