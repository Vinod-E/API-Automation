from hpro_automation import (input_paths, db_login)
import xlrd


class DeleteManageInterviewers(db_login.DBConnection):

    def _init__(self):
        super(DeleteManageInterviewers, self).__init__()
        self.db_connection('amsin')

        # -----------------------
        # Initialising the values
        # -----------------------
        self.xl_deleted_event_id = []

    def mi_excel_data(self):
        try:
            workbook = xlrd.open_workbook(input_paths.inputpaths['Manage_Int_Input_sheet'])
            sheet1 = workbook.sheet_by_index(1)
            for i in range(1, sheet1.nrows):
                number = i  # Counting number of rows
                rows = sheet1.row_values(number)

                self.xl_deleted_event_id.append(int(rows[6]))

        except IOError:
            print("File not found or path is incorrect")

    def delete_rows(self):

        query = "DELETE FROM interview_nominees_status_for_events WHERE event_id={};".format(self.xl_deleted_event_id)
        self.cursor.execute(query)
        print(query)


Object = DeleteManageInterviewers()
Object.mi_excel_data()
Object.delete_rows()
Object.connection.commit()
Object.connection.close()
