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

    def chart_sheets(self):
        # Some sample data to plot.
        farm_1 = {'apples': 10, 'berries': 32, 'squash': 21, 'melons': 13, 'corn1': 18, 'corn2': 18, 'corn': 18}
        farm_2 = {'apples': 15, 'berries': 43, 'squash': 17, 'melons': 10, 'corn': 22}
        farm_3 = {'apples': 6, 'berries': 24, 'squash': 22, 'melons': 16, 'corn': 30}
        farm_4 = {'apples': 12, 'berries': 30, 'squash': 15, 'melons': 9, 'corn': 15}

        data = [farm_1, farm_2, farm_3, farm_4]
        index = ['Farm 11', 'Farm 2', 'Farm 3', 'Farm 4']

        # Create a Pandas dataframe from the data.
        df = pd.DataFrame(data, index=index)

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        excel_file = 'grouped_column_farms.xlsx'
        sheet_name = 'Sheet1'

        writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
        df.to_excel(writer, sheet_name=sheet_name)

        # Access the XlsxWriter workbook and worksheet objects from the dataframe.
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Create a chart object.
        chart = workbook.add_chart({'type': 'column'})

        # Configure the series of the chart from the dataframe data.
        for col_num in range(1, len(farm_1) + 1):
            chart.add_series({
                'name': ['Sheet1', 0, col_num],
                'categories': ['Sheet1', 1, 0, 4, 0],
                'values': ['Sheet1', 1, col_num, 4, col_num],
                'fill': {'color': brews['Set1'][col_num - 1]},
                'overlap': -10,
            })

        # Configure the chart axes.
        chart.set_x_axis({'name': 'Total Produce'})
        chart.set_y_axis({'name': 'Farms', 'major_gridlines': {'visible': False}})

        # Insert the chart into the worksheet.
        worksheet.insert_chart('H2', chart)

        # Close the Pandas Excel writer and output the Excel file.
        writer.save()

