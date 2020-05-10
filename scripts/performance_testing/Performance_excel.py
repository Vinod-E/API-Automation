import xlrd
import input_paths


class PerformanceExcel(object):
    def __init__(self):
        self.login_server = str(input('Server :: '))
        self.eu = str(input("EU Yes(or)No :: "))

        # ------------ Data initialization ---------------------
        self.xl_tenant_name = []
        self.xl_property_ids = []
        self.xl_test_id = []

        self.tenant_name = ''
        self.property_ids = []
        self.test_id = ''

        # ------------- file reader index -------------------
        workbook = xlrd.open_workbook(input_paths.performance['performance'])
        if self.login_server == 'amsin':
            if self.eu.lower() == 'no':
                self.performance_sheet = workbook.sheet_by_index(0)
            elif self.eu.lower() == 'yes':
                self.performance_sheet = workbook.sheet_by_index(1)

        elif self.login_server == 'ams':
            if self.eu.lower() == 'no':
                self.performance_sheet = workbook.sheet_by_index(2)
            elif self.eu.lower() == 'yes':
                self.performance_sheet = workbook.sheet_by_index(3)

    def excel_read_by_index(self):
        # --------------- Excel Read ----------------------------
        for i in range(1, self.performance_sheet.nrows):
            number = i  # Counting number of rows
            rows = self.performance_sheet.row_values(number)

            if rows[0]:
                self.xl_tenant_name.append(rows[0])
            if rows[1]:
                self.xl_property_ids.append(rows[1])
            if rows[2]:
                self.xl_test_id.append(rows[2])

        for i in self.xl_tenant_name:
            self.tenant_name = i

        for j in self.xl_property_ids:
            ids = j
            ids = ids.split(",")
            q = map(int, ids)
            self.property_ids.extend(q)

        for k in self.xl_test_id:
            self.test_id = k
