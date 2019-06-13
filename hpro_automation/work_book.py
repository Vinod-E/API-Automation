import xlwt
from hpro_automation import styles


class WorkBook(styles.FontColor):

    def __init__(self):

        super(WorkBook, self).__init__()

        # ------------------------------------ generic excel headers part ----------------------------------------------
        self.main_headers = []
        self.headers_with_style2 = []  # Here can decide which color of the column to be use
        self.headers_with_style9 = []
        self.headers_with_style10 = []
        self.headers_with_style15 = []
        self.headers_with_style19 = []
        self.headers_with_style20 = []
        self.headers_with_style21 = []
        self.headers_with_style22 = []

        # -------------------- Create an new Excel file and add a worksheet. -------------------------------------------
        self.wb_Result = xlwt.Workbook()
        self.ws = self.wb_Result.add_sheet('API_Automation')

        self.final_status_rowsize = 0
        self.rowsize = 2
        self.col = 0

    def file_headers_col_row(self):

        header_column = 0
        excelheaders = self.main_headers
        for headers in excelheaders:
            if headers in self.headers_with_style2:
                self.ws.write(1, header_column, headers, self.style2)
            elif headers in self.headers_with_style9:
                self.ws.write(1, header_column, headers, self.style9)
            elif headers in self.headers_with_style10:
                self.ws.write(1, header_column, headers, self.style10)
            elif headers in self.headers_with_style15:
                self.ws.write(1, header_column, headers, self.style15)
            elif headers in self.headers_with_style19:
                self.ws.write(1, header_column, headers, self.style19)
            elif headers in self.headers_with_style20:
                self.ws.write(1, header_column, headers, self.style20)
            elif headers in self.headers_with_style21:
                self.ws.write(1, header_column, headers, self.style21)
            elif headers in self.headers_with_style22:
                self.ws.write(1, header_column, headers, self.style22)
            else:
                self.ws.write(1, header_column, headers, self.style0)
            header_column += 1

        # Set up some formats to use.
        # self.cell_format = self.wb_Result.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})

        # ------------We can only write simple types to merged ranges so we write a blank string.-----------------------
        # self.merge_cell = self.ws.merge_range('A1:B1', "", self.cell_format)
