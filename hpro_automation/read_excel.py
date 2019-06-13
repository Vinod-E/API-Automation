import xlrd


class ExcelRead:

    # ----------
    # Notes:-
    # ----------
    # 1. Excel Header Name Should be in camel case and should not have space. ex:- HeaderName
    # 2. This Script converts excel data in to dictionary, Key is 0th Row, Value is 1st row to End of Row
    # 3. Need to call excel_read method with file_path and sheet index from where ever you want.
    # 4. Don't do any changes in the below script. process your data in your child script.

    # Data Processing Ex:-
    # self.details will give as list so if you want to process your data please iterate your data and process it.
    # To use :- self.details.get('HeaderName').
    # convert to int - int(self.details.get('HeaderName')).
    # convert not none :- int(self.current_data.get('HeaderName')) if self.current_data.get('HeaderName') else None
    # Convert to date :-    convert_date_of_birth = self.current_data.get('HeaderName')
    #                       self.date_of_birth = datetime.datetime(*xlrd.xldate_as_tuple \
    #                       (convert_date_of_birth, excel_read_obj.excel_file.datemode))
    #                       self.date_of_birth = self.date_of_birth.strftime("%Y-%m-%d")
    # -------------------------------------------------------------------------------------------------------------------

    def __init__(self):
        self.complete_excel_data = []
        self.details1 = {}
        # file_path = input_paths.inputpaths['Duplication_rule_Input_sheet']
        # sheet_index = 0
        # file_path = '/home/vinod/PycharmProjects/API_AUTOMATION/Input Data/Crpo/Duplication_rule/Duplication_Rule.xls'
        # sheet_index = 0
        # self.excel_read(file_path, sheet_index)

    def excel_read(self, excel_file_path, sheet_index):
        self.excel_file = xlrd.open_workbook(excel_file_path)
        self.excel_sheet_index = self.excel_file.sheet_by_index(sheet_index)
        self.headers_available_in_excel = self.excel_sheet_index.row_values(0)
        exp_headers = {}
        for exp_headers_dictionary in self.headers_available_in_excel:
            d = {}
            d = {str(exp_headers_dictionary): str(exp_headers_dictionary)}
            exp_headers.update(d)
        for row_by_index in range(1, self.excel_sheet_index.nrows):
            column_by_index = 0
            self.details1 = {}
            excel_row_data = self.excel_sheet_index.row_values(row_by_index)
            for excel_header_value in self.headers_available_in_excel:
                for key, value in exp_headers.items():
                    if value == excel_header_value:
                        data = {key: excel_row_data[column_by_index]}
                        self.details1.update(data)
                        column_by_index = column_by_index + 1
            self.complete_excel_data.append(self.details1)
        # print self.complete_excel_data


excel_read_obj = ExcelRead()
