import xlrd
from hpro_automation import (output_paths, db_login)


class DeleteQuery(db_login.DBConnection):
    def __init__(self):
        super(DeleteQuery, self).__init__()
        self.db_connection()

        # ------------------------
        # Initialising Excel Data
        # ------------------------
        self.xl_userId = []

    def user_excel_data(self):

        workbook = xlrd.open_workbook(output_paths.DMLOutput['User_output_sheet'])
        sheet1 = workbook.sheet_by_index(0)
        for i in range(2, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)
            if rows[2]:
                self.xl_userId.append(int(rows[2]))

        xl_userid1 = tuple(self.xl_userId)
        user_query = "UPDATE appserver_core.users SET tenant_id='0', is_archived='1'," \
                     " is_deleted='1' WHERE id in{}".format(xl_userid1)
        self.cursor.execute(user_query)
        print(user_query)


Object = DeleteQuery()
Object.user_excel_data()
Object.connection.commit()
Object.connection.close()
