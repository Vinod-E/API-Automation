import os.path
from hpro_automation import (work_book, output_paths, api)
from scripts.performance_testing import performance_apis
from openpyxl import load_workbook
import pandas as pd
from pandas import ExcelFile
from vincent.colors import brews


class AmsinNonEuOutput(work_book.WorkBook, performance_apis.PerformanceTesting):
    def __init__(self):
        super(AmsinNonEuOutput, self).__init__()
        self.output_file = output_paths.outputpaths['performance_testing']
        self.all_data = []
        self.value_dict = {}

        self.get_tenant_details_dict = {}
        self.get_all_entity_properties_dict = {}
        self.group_by_catalog_masters_dict = {}
        self.get_all_candidates_dict = {}
        self.getTestUsersForTest_dict = {}

        self.summation = 0
        self.get_tenant_details_sum = 0
        self.get_all_entity_properties_sum = 0
        self.group_by_catalog_masters_sum = 0
        self.get_all_candidates_sum = 0
        self.getTestUsersForTest_sum = 0

    def create_pandas_excel(self, sheet_name):
        # ----------------------- Headers initialization ----------------------------
        h1 = 'Run Date'
        h2 = 'Run Time'
        h3 = 'get_tenant_details'
        h4 = 'get_all_entity_properties'
        h5 = 'group_by_catalog_masters'
        h6 = 'get_all_candidates'
        h7 = 'getTestUsersForTest'
        headers = [h1, h2, h3, h4, h5, h6, h7]

        # ------------------ Validation for File exists  ------------------------------
        local_path = os.path.exists(self.output_file)
        if local_path:
            print('**----->> File exists in your machine')
        else:
            vinod = pd.ExcelWriter(self.output_file, engine='xlsxwriter')
            sheetsList = ['AMSIN_NON_EU', 'AMSIN_EU', 'LIVE_NON_EU', 'LIVE_EU']

            for new_sheet in sheetsList:
                df_dynamic = pd.DataFrame(columns=headers)
                df_dynamic.to_excel(vinod, sheet_name=new_sheet, startrow=1, header=False, index=False)
                workbook = vinod.book
                header_format = workbook.add_format({'bold': True, 'valign': 'top', 'fg_color': '#00FA9A',
                                                     'font_size': 10.5})
                worksheet = vinod.sheets[new_sheet]
                for col_num, value in enumerate(df_dynamic.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    col_num += 1
            vinod.save()
            print('**----->> File has been created successfully')

        # ------------- Appending the values into their columns --------------------
        df = pd.DataFrame(columns=headers)
        df.loc[1, h1] = self.run_date
        df.loc[1, h2] = self.run_time
        df.loc[1, h3] = self.Average_Time_tenant_details
        df.loc[1, h4] = self.Average_Time_entity
        df.loc[1, h5] = self.Average_Time_catalog
        df.loc[1, h6] = self.Average_Time_candidates
        df.loc[1, h7] = self.Average_Time_testuser

        writer = pd.ExcelWriter(self.output_file, engine='openpyxl')
        writer.book = load_workbook(self.output_file)
        writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
        reader = pd.read_excel(self.output_file, sheet_name=sheet_name)
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=len(reader) + 1)

        writer.close()

    def read_data_from_excel(self, sheet_name):
        xlsx = ExcelFile(self.output_file)
        df = xlsx.parse(sheet_name)
        dict_sheet_values = df.to_dict()
        print(dict_sheet_values)

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

        self.summation_data(self.get_tenant_details_dict)
        self.get_tenant_details_sum = self.summation

        self.summation_data(self.get_all_entity_properties_dict)
        self.get_all_entity_properties_sum = self.summation

        self.summation_data(self.group_by_catalog_masters_dict)
        self.group_by_catalog_masters_sum = self.summation

        self.summation_data(self.get_all_candidates_dict)
        self.get_all_candidates_sum = self.summation

        self.summation_data(self.getTestUsersForTest_dict)
        self.getTestUsersForTest_sum = self.summation

        frame = {'get_tenant_details': self.get_tenant_details_sum,
                 'get_all_entity_properties': self.get_all_entity_properties_sum,
                 'group_by_catalog_masters': self.group_by_catalog_masters_sum,
                 'get_all_candidates': self.get_all_candidates_sum,
                 'getTestUsersForTest': self.getTestUsersForTest_sum}
        print(frame)

    def summation_data(self, api_dict_time):
        item = 0
        self.summation = 0
        for k in range(0, len(api_dict_time)):
            if str(api_dict_time.get(item)) != 'nan':
                self.summation = self.summation + api_dict_time.get(item)
            item += 1
        print(self.summation)

    def chart_sheets(self, sheet_name):
        self.output_file = output_paths.outputpaths['performance_testing']
        data = [self.value_dict]
        index = ['Farm 1', 'Farm 2', 'Farm 3', 'Farm 4']
        df = pd.DataFrame(data, index=index)
        writer = pd.ExcelWriter(self.output_file, engine='xlsxwriter')
        df.to_excel(writer, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        chart = workbook.add_chart({'type': 'column'})
        for col_num in range(1, len(self.run_date) + 1):
            chart.add_series({
                'name': [sheet_name, 0, col_num],
                'categories': [sheet_name, 1, 0, 4, 0],
                'values': [sheet_name, 1, col_num, 4, col_num],
                'overlap': -10,
            })

        chart.set_x_axis({'name': 'Total Produce'})
        chart.set_y_axis({'name': 'Farms', 'major_gridlines': {'visible': False}})

        worksheet.insert_chart('A2', chart)
        writer.save()
