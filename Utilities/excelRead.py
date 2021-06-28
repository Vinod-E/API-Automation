from pandas import *


class ExcelRead(object):
    """
    Excel read from the excel and store all values in dict.

        :Notes:
        - excel_read :: Using pandas package to read the excel and making the header / values
                        as dict.

    """
    def __init__(self):
        super(ExcelRead, self).__init__()
        self.excel_dict = {}

    def read(self, file_path, index):
        try:

            xls = ExcelFile(file_path)
            dataframe = xls.parse(xls.sheet_names[index])
            self.excel_dict = dataframe.to_dict()

        except Exception as ExcelReadError:
            print(ExcelReadError)
