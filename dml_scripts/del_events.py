from hpro_automation import (output_paths, db_login)
import xlrd


class DeleteEvents(db_login.DBConnection):

    def __init__(self):
        super(DeleteEvents, self).__init__()
        self.db_connection('amsin')

        # -----------------------
        # Initialising the values
        # -----------------------
        self.xl_output_event_ids = []

    def event_excel_data(self):
        workbook = xlrd.open_workbook(output_paths.outputpaths['Event_output_sheet'])
        sheet1 = workbook.sheet_by_index(0)
        for i in range(2, 26):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)
            if rows[2]:
                self.xl_output_event_ids.append(int(rows[2]))

    def archive_event(self):
        event_ids = tuple(self.xl_output_event_ids)
        query = "UPDATE recruit_events SET is_archived=1 WHERE id in {};" .format(event_ids)
        print(query)
        self.cursor.execute(query)


Object = DeleteEvents()
Object.event_excel_data()
Object.archive_event()
Object.connection.commit()
Object.connection.close()
