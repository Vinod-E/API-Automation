from hpro_automation import (input_paths, db_login)
import xlrd


class DeleteManageInterviewers(db_login.DBConnection):

    def __init__(self):
        super(DeleteManageInterviewers, self).__init__()
        self.db_connection()

        # -----------------------
        # Initialising the values
        # -----------------------
        self.xl_deleted_event_id = []
        self.xl_role_id = []

    def mi_excel_data(self):
        try:
            workbook = xlrd.open_workbook(input_paths.inputpaths['Manage_Int_Input_sheet'])
            sheet1 = workbook.sheet_by_index(1)
            for i in range(1, sheet1.nrows):
                number = i  # Counting number of rows
                rows = sheet1.row_values(number)

                self.xl_deleted_event_id.append(int(rows[6]))
                self.xl_deleted_event_id.append(0)
                self.xl_role_id.append(int(rows[7]))
                self.xl_role_id.append(0)

        except IOError:
            print("File not found or path is incorrect")

    def delete_nomination_rows(self):

        event_id = tuple(self.xl_deleted_event_id)
        query = "DELETE FROM interview_nominees_status_for_events WHERE event_id in{};".format(event_id)
        self.cursor.execute(query)
        print(query)

    def untag_interviewers(self):

        role_id = tuple(self.xl_role_id)
        query = 'DELETE FROM recruit_event_owners WHERE role_id in {};'.format(role_id)
        self.cursor.execute(query)
        print(query)


Object = DeleteManageInterviewers()
Object.mi_excel_data()
Object.delete_nomination_rows()
Object.untag_interviewers()
Object.connection.commit()
Object.connection.close()
