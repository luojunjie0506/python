import unittest
from selenium import webdriver
import time
from rongge.tool.tool import yzm,sfz,sjh
from rongge.testpage.models.myunit import MyunitTest
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select

class TestOperation(MyunitTest):
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

    def ss(self,*u):
        self.driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div/ul/li[2]/a').click()
        self.driver.find_element_by_id('module_applyManageList').click()
        if u ==():
            WebDriverWait(self.driver, 5, 0.2).until(lambda x: x.find_element_by_xpath('/html/body/div[3]/ul/li/div/div/table/tbody/tr[1]/td[2]/a')).click()

    #修改经销商名字
    def test25(self):
        self.ss()
        self.driver.find_element_by_id('dealername').clear()
        self.driver.find_element_by_id('dealername').send_keys('你好')
        self.driver.find_element_by_id('regeditBtn').click()
        text2 = WebDriverWait(self.driver, 5, 0.2).until(
            lambda x: x.find_element_by_xpath('/html/body/div[3]/ul/li/div/div/table/tbody/tr[1]/td[3]')).text
        self.assertEquals(text2,'你好')


    #修改已存在的身份证号
    def test26(self):
        self.ss()
        WebDriverWait(self.driver, 5, 0.2).until(
            lambda x: x.find_element_by_id('idcard')).clear()
        self.driver.find_element_by_id('idcard').send_keys('441624199505065297')
        self.driver.find_element_by_id('mobile').click()
        text1 = WebDriverWait(self.driver, 5, 0.2).until(
            lambda x: x.find_element_by_id('paperNoMessage')).text
        self.assertEquals(text1,'您使用的身份证已被注册！')

    #修改为已存在的手机号
    def test27(self):
        self.ss()
        WebDriverWait(self.driver, 5, 0.2).until(
            lambda x: x.find_element_by_id('mobile')).clear()
        self.driver.find_element_by_id('mobile').send_keys('13168090784')
        self.driver.find_element_by_id('regeditBtn').click()
        text = WebDriverWait(self.driver,5,0.2).until(lambda x:x.find_element_by_xpath('/html/body/div[6]/div/table/tbody/tr[2]/td/div/span')).text
        self.assertEquals(text,'此手机号已经注册，请重新填写')

    #修改为不存在的身份证号
    def test28(self):
        self.ss()
        a= sfz()
        WebDriverWait(self.driver, 5, 0.2).until(
            lambda x: x.find_element_by_id('idcard')).clear()
        self.driver.find_element_by_id('idcard').send_keys(a)
        self.driver.find_element_by_id('mobile').click()
        self.driver.find_element_by_id('regeditBtn').click()
        WebDriverWait(self.driver,5,0.2).until(lambda x: x.find_element_by_xpath('/html/body/div[3]/ul/li/div/div/table/tbody/tr[1]/td[2]/a')).click()
        text1 = WebDriverWait(self.driver, 5, 0.2).until(lambda x: x.find_element_by_id('idcard')).get_attribute('value')
        time.sleep(2)
        self.assertEquals(a,text1)

    #修改为不存在的手机号
    def test29(self):
        self.ss()
        b= sjh()
        WebDriverWait(self.driver, 5, 0.2).until(
            lambda x: x.find_element_by_id('mobile')).clear()
        self.driver.find_element_by_id('mobile').send_keys(b)
        self.driver.find_element_by_id('regeditBtn').click()
        text1 = WebDriverWait(self.driver, 5, 0.2).until(
            lambda x: x.find_element_by_xpath('/html/body/div[3]/ul/li/div/div/table/tbody/tr[1]/td[8]')).text
        self.assertEquals(b,text1)

    #作废
    def test30(self):
        self.ss(2)
        time.sleep(2)
        self.driver.find_element_by_xpath('/html/body/div[3]/ul/li/div/div/table/tbody/tr/td[1]/input').click()
        self.driver.find_element_by_id('cancel_sure').click()
        text1 = WebDriverWait(self.driver, 5, 0.2).until(lambda x: x.find_element_by_xpath('/html/body/div[7]/div/table/tbody/tr[2]/td/div/span')).text
        self.driver.find_element_by_xpath('/html/body/div[7]/div/table/tbody/tr[1]/td/button').click()
        self.assertIn('作废成功',text1)

    #激活
    def test31(self):
        self.ss(2)
        s1 = Select(self.driver.find_element_by_id('status'))
        s1.select_by_index(1)
        self.driver.find_element_by_id('grid_querybtn').click()
        time.sleep(1)
        self.driver.find_element_by_xpath('/html/body/div[3]/ul/li/div/div/table/tbody/tr/td[1]/input').click()
        self.driver.find_element_by_id('ok_sure').click()
        text1 = WebDriverWait(self.driver, 5, 0.2).until(
            lambda x: x.find_element_by_xpath('/html/body/div[7]/div/table/tbody/tr[2]/td/div/span')).text
        self.assertEquals('激活经销商卡号成功!',text1)
