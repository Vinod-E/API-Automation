import subprocess
from Config.stackranking_config import *
from Config.writeExcel import *
import datetime
import time


class StackRanking:

    def __init__(self):
        now = datetime.datetime.now()
        self.current_date = now.strftime("%d-%m-%Y")

    @staticmethod
    def download_stack_ranking_report():

        generate_stack_ranking_report = requests.post(config_obj.getall_applicant_api, headers=config_obj.headers,
                                                      data=json.dumps(config_obj.stack_ranking_report_payload,
                                                                      default=str), verify=False)
        resp_dict = json.loads(generate_stack_ranking_report.content)
        api_status = resp_dict['status']
        if api_status == 'OK':
            time.sleep(5)
            subprocess.check_output(['wget', '-O', config_obj.download_path, resp_dict['data']['downloadLink']])
            print("Download Success")
        else:
            print("Download Failed Check Manually")


air = StackRanking()
try:
    config_obj.filePath(air.current_date)
    air.download_stack_ranking_report()
    excel_object.save_result(config_obj.save_path)
    excel_object.excelReadExpectedSheet(config_obj.expected_excel_sheet_path)
    excel_object.excelReadActualSheet(config_obj.download_path)
    excel_object.excelWriteHeaders(hierarchy_headers_count=2)
    excel_object.excelMatchValues(usecase_name=' Stack Ranking Report ', comparision_required_from_index=2,
                                  total_testcase_count=4)
except Exception as e:
    print("------------------------------------------------------------------------")
    print("Please verify it manually")
    print("------------------------------------------------------------------------")
    print(e)
