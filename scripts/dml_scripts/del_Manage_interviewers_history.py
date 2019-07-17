from hpro_automation import (input_paths, db_login)
import xlrd


class DeleteManageInterviewers(db_login.DBConnection):

    def _init__(self):
        super(DeleteManageInterviewers, self).__init__()
        self.db_connection('amsin')

        # -----------------------
        # Initialising the values
        # -----------------------
        self.xl_EventId = []

    def mi_excel_data(self):
        workbook = xlrd.open_workbook(input_paths.inputpaths[''])
        sheet1 = workbook.sheet_by_index(0)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)
            if rows[0]:
                self.xl_EventId.append(int(rows[0]))

    def delete_rows(self):

        query = "DELETE FROM interview_nominees_status_for_events WHERE event_id={};".format(self.xl_EventId)
        self.cursor.execute(query)
        print(query)


Object = DeleteManageInterviewers()
Object.delete_rows()
Object.connection.commit()
Object.connection.close()
