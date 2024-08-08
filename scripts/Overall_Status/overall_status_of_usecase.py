from hpro_automation import (work_book, output_paths)


class OverallStatus(work_book.WorkBook):

    def __init__(self):
        super(OverallStatus, self).__init__()
        self.calculate_actual_pass = []
        self.Actual_Success_case = []

    # -------------------------------
    # Writing overall status to excel
    # -------------------------------
    def overall_status(self, case_name, expected_success_cases, start_time,
                       calling_lambda, crpo_app_name, login_server, total_count, output_sheet):
        self.ws.write(0, 0, case_name, self.style23)
        if expected_success_cases == self.Actual_Success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
        else:
            self.ws.write(0, 1, 'Fail', self.style25)

        self.ws.write(0, 2, 'Login Server', self.style23)
        self.ws.write(0, 3, login_server, self.style24)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, calling_lambda, self.style24)
        self.ws.write(0, 6, 'APP Name', self.style23)
        self.ws.write(0, 7, crpo_app_name, self.style24)
        self.ws.write(0, 8, 'No.of Test cases', self.style23)
        self.ws.write(0, 9, total_count, self.style24)
        self.ws.write(0, 10, 'Start Time', self.style23)
        self.ws.write(0, 11, start_time, self.style26)
        self.wb_Result.save(output_paths.outputpaths[output_sheet])

        print("Actual Pass Count:: ", len(self.Actual_Success_case))
        print("Expected Pass Count:: ", len(expected_success_cases))

    def test_case_wise_pass_or_fail(self):
        row = self.rowsize - 1
        # print(self.calculate_actual_pass)
        if 'Fail' in self.calculate_actual_pass:
            self.ws.write(row, 1, "Fail", self.style3)
            self.calculate_actual_pass = []
            self.wb_Result.save(output_paths.outputpaths['Candidate_slot_output_sheet'])
        else:
            self.ws.write(row, 1, "Pass", self.style26)
            self.Actual_Success_case.append('Pass')
            self.calculate_actual_pass = []
            self.wb_Result.save(output_paths.outputpaths['Candidate_slot_output_sheet'])

    def output_excel_input_output_header(self, output_file_key):

        self.ws.write(self.rowsize, self.col, 'Input', self.style4)
        self.rowsize += 1  # Row increment

        self.ws.write(self.rowsize, self.col, 'Output', self.style5)
        self.rowsize += 1  # Row increment

        self.wb_Result.save(output_paths.outputpaths[output_file_key])

    def validation(self, col, input_value, output_value):
        # --------- Input
        row = self.rowsize - 2
        col_row = self.rowsize - 1
        self.ws.write(row, col, input_value)

        # --------- Output
        if output_value == input_value:
            self.ws.write(col_row, col, output_value, self.style14)
            self.calculate_actual_pass.append('Pass')
        else:
            self.ws.write(col_row, col, output_value, self.style3)
            self.calculate_actual_pass.append('Fail')

    def write_in_excel(self, column, excel_dict, validation_api_key1,
                       validation_api_key2, additional_message):

        if validation_api_key1:
            self.validation(column, excel_dict, validation_api_key1)
        elif validation_api_key2:
            self.validation(column, excel_dict, validation_api_key2)
        else:
            self.validation(column, excel_dict, additional_message)
