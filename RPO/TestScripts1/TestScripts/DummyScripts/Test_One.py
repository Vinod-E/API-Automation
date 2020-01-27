import unittest
import sys
import json
import urllib2


class RestApiTest(unittest.TestCase):
    def setUp(self):
        # define Api URL and API Key
        self.ApiUrl = "http://api.openweathermap.org/data/2.5/weather"
        self.ApiKey = "70926ddfd37fdf454548b8db13695995"

    def test_weather_api_by_city_name1(self):
        # define api response
        testurl = (self.ApiUrl + "?q=Baltimore,us" + "&" + "APPID=" + self.ApiKey)
        print testurl
        response = urllib2.urlopen(testurl)
        # read response
        html = response.read()
        # print response
        print(html)
        # assert response
        self.assertTrue("Baltimore" in html)

    def test_weather_api_by_city_name2(self):
        # define api response
        testurl = (self.ApiUrl + "?q=Baltimore,us" + "&" + "APPID=" + self.ApiKey)
        print testurl
        response = urllib2.urlopen(testurl)
        # read response
        html = response.read()
        # print response
        print(html)
        # loads response as json
        json_data = json.loads(html)
        # get the key "name" value
        city_name = json_data["name"]
        print("city name is:" + city_name)
        # assert city name
        self.assertTrue(city_name == "Baltimore")

    def tearDown(self):
        print "------- test is over -------"


if __name__ == "__main__":
    unittest.main()