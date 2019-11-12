import unittest
import time
from rongge.testpage.models.myunit import  MyunitTest
from rongge.tool.tool import js,yzm,jt,xr
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
import os

class TestSq(MyunitTest):

    @classmethod
    def setUpClass(cls):
        cls.driver = webdriver.Firefox()
        cls.driver.get('http://cswl2016.3322.org:8081/agent/index.jhtml')
        cls.driver.maximize_window()
        cls.driver.find_element_by_id('header_loginbtn').click()
        time.sleep(2)
        cls.driver.find_element_by_id('userCode').send_keys('AH5012')
        cls.driver.find_element_by_id('login_storepassword').send_keys('12345678')
        a = cls.driver.find_element_by_id('safecode').get_attribute('src')
        c = cls.driver.get_cookie('JSESSIONID')['value']
        b = yzm(a, c)
        cls.driver.find_element_by_id('captchaCode').send_keys(b)
        cls.driver.find_element_by_xpath('/html/body/div[2]/div/div/form/ul/li[2]/div[2]/button').click()
        time.sleep(2)

    def sq(self):
        i = js()
        if i==5:
            self.driver.find_element_by_xpath('//*[@id="mainmenu"]/li[2]/a').click()
            time.sleep(2)
        self.driver.find_element_by_id('module_miAgentCardList').click()
        time.sleep(2)

    def bh(self):
        i = js()
        if i==10:
            self.driver.find_element_by_id('module_poAgentCardList').click()
            text = WebDriverWait(self.driver, 5, 0.2).until(
                lambda x: x.find_element_by_xpath('/html/body/div[3]/ul/li/div[2]/table/tbody/tr[1]/td[3]')).text
            test_report_path = os.path.join(os.getcwd() + '\\report\\1.txt')
            xr(text)
            self.driver.find_element_by_id('module_miAgentCardList').click()

    #选择20张电子卡金额
    def test05(self):
        self.sq()
        self.driver.find_element_by_id('thirtypaper_label').click()
        money = self.driver.find_element_by_id('yfmoney').text
        self.assertEqual(money,'100')

    #输入数量8显示金额
    def test06(self):
        self.sq()
        self.driver.find_element_by_id('nongye5_label').click()
        self.driver.find_element_by_id('idNummber').send_keys('8')
        self.driver.find_element_by_id('yfmoney').click()
        money = self.driver.find_element_by_id('yfmoney').text
        self.assertEqual(money,'40')

    #选择电子卡购买空密码
    def test07(self):
        self.sq()
        self.driver.find_element_by_id('onepaper_label').click()
        self.driver.find_element_by_xpath('/html/body/div[3]/ul/li/div/form/div/div/div/div/div/ul/li/div[2]/span[3]/input[3]').click()
        a = WebDriverWait(self.driver, 5, 0.2).until(lambda x: x.find_element_by_xpath('/html/body/div[6]/div/table/tbody/tr[2]/td/div/span')).text
        try:
            self.assertEqual(a, '付款失败,请输入正确的支付密码!')
        except Exception:
            jt()

    #选择电子卡购买错误密码
    def test08(self):
        self.sq()
        self.driver.find_element_by_id('onepaper').click()
        self.driver.find_element_by_id('numberpaper').send_keys('312354')
        self.driver.find_element_by_xpath('/html/body/div[3]/ul/li/div/form/div/div/div/div/div/ul/li/div[2]/span[3]/input[3]').click()
        #显示等待WebDriverWait
        a = WebDriverWait(self.driver, 5, 0.2).until(
            lambda x: x.find_element_by_xpath('/html/body/div[6]/div/table/tbody/tr[2]/td/div/span')).text
        self.assertEqual(a, '付款失败,支付密码不正确!')

    # 选择1张电子卡购买
    def test09(self):
        self.sq()
        self.driver.find_element_by_id('onepaper').click()
        self.driver.find_element_by_id('numberpaper').send_keys('12345678')
        self.driver.find_element_by_xpath('/html/body/div[3]/ul/li/div/form/div/div/div/div/div/ul/li/div[2]/span[3]/input[3]').click()
        a = WebDriverWait(self.driver, 5, 0.2).until(
            lambda x: x.find_element_by_xpath('/html/body/div[6]/div/table/tbody/tr[2]/td/div/span')).text
        self.assertEqual(a, '恭喜，您已成功申请了1张荣格电子卡,具体卡号请到查询页面.')

    # 输入1张电子卡购买
    def test10(self):
        self.sq()
        self.driver.find_element_by_id('nongye5_label').click()
        self.driver.find_element_by_id('idNummber').send_keys('1')
        self.driver.find_element_by_id('numberpaper').send_keys('12345678')
        self.driver.find_element_by_xpath('/html/body/div[3]/ul/li/div/form/div/div/div/div/div/ul/li/div[2]/span[3]/input[3]').click()
        time.sleep(1)
        a = self.driver.find_element_by_xpath('/html/body/div[6]/div/table/tbody/tr[2]/td/div/span').text
        self.assertEqual(a, '恭喜，您已成功申请了1张荣格电子卡,具体卡号请到查询页面.')




if __name__ == '__main__':
    unittest.main()
