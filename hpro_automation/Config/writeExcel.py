import xlrd
import xlsxwriter
import datetime


class Excel:

    def __init__(self):
        self.now = datetime.datetime.now()
        self.started_time = self.now.strftime("%d-%m-%Y %H:%M")

    def save_result(self, save_excel_path):
        self.write_excel = xlsxwriter.Workbook(save_excel_path)
        self.ws = self.write_excel.add_worksheet()
        self.black_color = self.write_excel.add_format({'font_color': 'black', 'font_size': 9})
        self.red_color = self.write_excel.add_format({'font_color': 'red', 'font_size': 9})
        self.green_color = self.write_excel.add_format({'font_color': 'green', 'font_size': 9})
        self.black_color_bold = self.write_excel.add_format({'font_color': 'black', 'bold': True, 'font_size': 9})

    def excelReadExpectedSheet(self, excepted_sheet_path):
        self.expected_excel = xlrd.open_workbook(excepted_sheet_path)
        self.expected_excel_sheet1 = self.expected_excel.sheet_by_index(0)

    def excelReadActualSheet(self, actual_sheet_path):
        self.actual_excel = xlrd.open_workbook(actual_sheet_path)
        self.actual_excel_sheet1 = self.actual_excel.sheet_by_index(0)

    def excelWriteHeaders(self, hierarchy_headers_count):
        for i in range(0, hierarchy_headers_count):
            expected_sheet_rows = self.expected_excel_sheet1.row_values(i)
            self.ws.write(i+1, 0, "Header - " + str(i+1), self.black_color_bold)
            for j in range(1, self.expected_excel_sheet1.ncols):
                self.ws.write(i+1, j + 2, expected_sheet_rows[j], self.black_color_bold)

    def excelMatchValues(self, usecase_name, comparision_required_from_index, total_testcase_count):
        self.ws.write(0, 0, usecase_name, self.black_color_bold)
        self.write_position = comparision_required_from_index + 1
        self.overall_status = 'Pass'
        self.overall_status_color = self.green_color
        for row_indx in range(comparision_required_from_index, self.expected_excel_sheet1.nrows):
            expected_sheet_rows = self.expected_excel_sheet1.row_values(row_indx)
            actual_sheet_rows = self.actual_excel_sheet1.row_values(row_indx)
            self.ws.write(self.write_position, 0, "Expected ", self.black_color)
            self.ws.write(self.write_position + 1, 0, "Actual ", self.black_color)
            self.status = 'Pass'
            self.color = self.green_color
            for col_indx in range(0, self.expected_excel_sheet1.ncols):
                if expected_sheet_rows[col_indx] == actual_sheet_rows[col_indx]:
                    self.ws.write(self.write_position, col_indx + 2, expected_sheet_rows[col_indx], self.black_color)
                    self.ws.write(self.write_position + 1, col_indx + 2, actual_sheet_rows[col_indx], self.green_color)
                else:
                    self.ws.write(self.write_position, col_indx + 2, expected_sheet_rows[col_indx], self.black_color)
                    if not actual_sheet_rows[col_indx] or actual_sheet_rows[col_indx] == ' ':
                        self.ws.write(self.write_position + 1, col_indx + 2, "EMPTY",
                                      self.red_color)
                    else:
                        self.ws.write(self.write_position + 1, col_indx + 2, actual_sheet_rows[col_indx],
                                      self.red_color)
                    self.status = 'Fail'
                    self.color = self.red_color
                    self.overall_status = 'Fail'
                    self.overall_status_color = self.red_color
                self.ws.write(self.write_position, 1, self.status, self.color)
            self.write_position += 3
        self.ws.write(0, 1, "OverAll Status:- " + self.overall_status, self.overall_status_color)
        self.ws.write(0, 2, "Total Testcase Count:- " + str(total_testcase_count), self.black_color_bold)
        self.ws.write(0, 3, "Started :- " + str(self.started_time), self.black_color_bold)
        self.now = datetime.datetime.now()
        self.endeded_time = self.now.strftime("%d-%m-%Y %H:%M")
        self.ws.write(0, 4, "Ended :- " + str(self.endeded_time), self.black_color_bold)
        self.write_excel.close()

excel_object = Excel()
