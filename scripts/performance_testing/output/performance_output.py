import os.path
from Utilities import excelRead
from hpro_automation import (work_book, output_paths, api)
from scripts.performance_testing import performance_apis
from openpyxl import load_workbook
import pandas as pd


class AmsinNonEuOutput(work_book.WorkBook, performance_apis.PerformanceTesting):
    def __init__(self):
        super(AmsinNonEuOutput, self).__init__()
        self.run_number = 1
        self.output_file = output_paths.outputpaths['performance_testing']
        self.all_data = []

    def create_pandas_excel(self, sheet_name):
        # ----------------------- Headers initialization ----------------------------
        h1 = 'Run Number'
        h2 = 'Run Date'
        h3 = 'Run Time'
        h4 = 'get_tenant_details'
        h5 = 'get_all_entity_properties'
        h6 = 'group_by_catalog_masters'
        h7 = 'get_all_candidates'
        h8 = 'getTestUsersForTest'
        h9 = 'Interview'
        h10 = 'New_Interview'
        headers = [h1, h2, h3, h4, h5, h6, h7, h8, h9, h10]

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
            vinod._save()
            print('**----->> File has been created successfully')

        # ------------- Run Number increase --------------------
        if sheet_name == 'AMSIN_NON_EU':
            index = 0
        elif sheet_name == 'AMSIN_EU':
            index = 1
        elif sheet_name == 'LIVE_NON_EU':
            index = 2
        elif sheet_name == 'LIVE_EU':
            index = 3
        status_excel = excelRead.ExcelRead()
        status_excel.read(self.output_file, index=index)
        xl = status_excel.excel_dict
        run_num = xl['Run Number']
        number = max(run_num)
        vinod = number + 1
        # ------------- Appending the values into their columns --------------------
        df = pd.DataFrame(columns=headers)
        df.loc[1, h1] = vinod
        df.loc[1, h2] = self.run_date
        df.loc[1, h3] = self.run_time
        df.loc[1, h4] = self.Average_Time_tenant_details
        df.loc[1, h5] = self.Average_Time_entity
        df.loc[1, h6] = self.Average_Time_catalog
        df.loc[1, h7] = self.Average_Time_candidates
        df.loc[1, h8] = self.Average_Time_testuser
        df.loc[1, h9] = self.Average_Time_interview
        df.loc[1, h10] = self.Average_Time_new_interview

        book = load_workbook(self.output_file)
        writer = pd.ExcelWriter(self.output_file, engine='openpyxl')
        writer._book = book
        writer._sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
        writer._save()
        reader = pd.read_excel(self.output_file, sheet_name=sheet_name, engine='openpyxl')
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=len(reader) + 1)

        writer.close()
