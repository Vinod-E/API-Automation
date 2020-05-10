import os.path
from hpro_automation import (work_book, output_paths, api)
from scripts.performance_testing import performance_apis
from openpyxl.workbook import Workbook
from openpyxl import load_workbook
import pandas as pd


class AmsinNonEuOutput(work_book.WorkBook, performance_apis.PerformanceTesting):
    def __init__(self):
        super(AmsinNonEuOutput, self).__init__()
        self.all_data = []

    def create_pandas_excel(self, sheet_name):
        # ------------------------ Output file path --------------------------------
        output_file = output_paths.outputpaths['performance_testing']

        # ----------------------- Headers initialization ----------------------------
        h1 = 'Run Date'
        h2 = 'Run Time'
        h3 = api.lambda_apis['get_tenant_details']
        h4 = api.lambda_apis['get_all_entity_properties']
        h5 = api.lambda_apis['group_by_catalog_masters']
        h6 = api.lambda_apis['get_all_candidates']
        h7 = api.lambda_apis['getTestUsersForTest']
        headers = [h1, h2, h3, h4, h5, h6, h7]

        # ------------------ Validation for File exists  ------------------------------
        local_path = os.path.exists(output_file)
        if local_path:
            print('**----->> File exists in your machine')
        else:
            vinod = pd.ExcelWriter(output_file, engine='xlsxwriter')
            sheetsList = ['AMSIN_NON_EU', 'AMSIN_EU', 'LIVE_NON_EU', 'LIVE_EU']

            for new_sheet in sheetsList:
                df_dynamic = pd.DataFrame(columns=headers)
                df_dynamic.to_excel(vinod, sheet_name=new_sheet, startrow=1, header=False, index=False)
                workbook = vinod.book
                header_format = workbook.add_format({'bold': True, 'valign': 'top', 'fg_color': '#008000'})
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

        writer = pd.ExcelWriter(output_file, engine='openpyxl')
        writer.book = load_workbook(output_file)
        writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
        reader = pd.read_excel(output_file, sheet_name=sheet_name)
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=len(reader) + 1)

        writer.close()
