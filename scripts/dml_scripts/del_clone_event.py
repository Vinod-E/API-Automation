from hpro_automation import (output_paths, db_login)
import xlrd


class DeleteCloneEvents(db_login.DBConnection):

    def __init__(self):
        super(DeleteCloneEvents, self).__init__()
        self.db_connection('amsin')

        # -----------------------
        # Initialising the values
        # -----------------------
        self.xl_Cloned_event_ids = []
        self.xl_Cloned_test_ids = []

    def event_excel_data(self):
        workbook = xlrd.open_workbook(output_paths.outputpaths['Event_Clone_output_sheet'])
        sheet1 = workbook.sheet_by_index(0)
        for i in range(2, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)
            if rows[2]:
                self.xl_Cloned_event_ids.append(int(rows[2]))
            if rows[2]:
                self.xl_Cloned_test_ids.append(int(rows[2]))

    def archive_event(self):
        cloned_event_ids = tuple(self.xl_Cloned_event_ids)
        query = "UPDATE recruit_events SET is_archived=1 WHERE id in {};" .format(cloned_event_ids)
        print(query)
        self.cursor.execute(query)

    def archive_test(self):
        clone_test_ids = tuple(self.xl_Cloned_test_ids)
        query = "UPDATE tests SET is_archived=1 WHERE id in {};" .format(clone_test_ids)
        print(query)
        self.cursor.execute(query)


Object = DeleteCloneEvents()
Object.event_excel_data()
Object.archive_event()
Object.archive_test()
Object.connection.commit()
Object.connection.close()
