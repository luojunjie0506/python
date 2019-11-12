

import unittest
from selenium import webdriver


class MyunitTest(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.driver = webdriver.Firefox()
        cls.driver.get('http://cswl2016.3322.org:8081/agent/index.jhtml')
        cls.driver.maximize_window()


    @classmethod
    def tearDownClass(cls):
        cls.driver.close()


if __name__ == '__main__':
    unittest.main()
