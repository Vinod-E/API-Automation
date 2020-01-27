import unittest
import collections
import time
import datetime
import xlwt
import xlrd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

from TestScripts.Config.AllConstants import CONSTANT


class createNewReq(unittest.TestCase):
    def test_create_req(self):

            self.driver = webdriver.Chrome(CONSTANT.CHROME_DRIVER)
            self.driver.implicitly_wait(30)
            self.driver.maximize_window()
            # navigate to the application home page
            self.driver.get(CONSTANT.RPO_CRPO_AMS_URL)
            # time.sleep(1)
            self.driver.find_element_by_name("loginName").send_keys(CONSTANT.CRPO_RPO_LOGIN_NAME)
            self.driver.find_element_by_xpath("//div[3]/section/div[1]/div[2]/form/div[2]/input").send_keys(
                CONSTANT.CRPO_RPO_LOGIN_PASSWORD)
            self.driver.find_element_by_xpath("//div[3]/section/div[1]/div[2]/form/div[2]/input").send_keys(Keys.ENTER)
            self.driver.find_element_by_xpath("//ul/li[3]/a").click()
            self.driver.find_element_by_xpath("/div[3]/section/div/div/div[1]/div[2]/a/i").click()

    if __name__ == '__main__':
        unittest.main()
