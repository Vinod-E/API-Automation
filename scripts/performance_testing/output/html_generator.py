from scripts.HTML_Reports.history_data_html_generator import HistoryDataHTMLGenerator
from scripts.HTML_Reports.amazon_aws_s3 import AWS


class HtmlGenerator:
    def __init__(self):
        self.amazon_s3 = AWS('/home/vinod/hirepro_automation/API-Automation/Output Data/Crpo/Common_folder/vv.html',
                             '/home/vinod/hirepro_automation/API-Automation/Output Data/Crpo/Common_folder/vvv.html')

    def history_html_generator(self):
        HistoryDataHTMLGenerator.html_report_generation(self.server, self.version, self.start_date_time,
                                                        self.use_case_name, self.xlw.result,
                                                        self.xlw.total_cases, self.xlw.pass_cases,
                                                        self.xlw.failure_cases, self.xlw.percentage,
                                                        self.xlw.minutes, self.time, self.xlw.date_now,
                                                        self.__path)
        self.amazon_s3.file_handler()
