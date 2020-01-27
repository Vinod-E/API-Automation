import unittest
import requests


class MyTestCase(unittest.TestCase):
    def test_something(self):
        host = 'amsin.hirepro.in'
        Token  = X-AUTH-TOKEN
        self.ApiUrl = "https://amsin.hirepro.in/py/rpo/calculate_salary_structure/"
        req = requests.get(self.ApiUrl)
        print (req.text)


if __name__ == '__main__':
    unittest.main()
