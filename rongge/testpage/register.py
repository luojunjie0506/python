from rongge.testpage.models.myunit import  MyunitTest
from selenium import webdriver
from rongge.tool.tool import yzm,js,sfz,mz,jt,sjh
import unittest
import time
from selenium.webdriver.support.ui import WebDriverWait

class TestRegister(MyunitTest):

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

    def register(self):
        self.driver.find_element_by_xpath('//*[@id="mainmenu"]/li[2]/a').click()
        WebDriverWait(self.driver, 5, 0.2).until(lambda x: x.find_element_by_id('module_agentNewMember')).click()
        time.sleep(2)

    #联系人卡号为空
    def test11(self):
        self.register()
        self.driver.find_element_by_id('referenceno').click()
        self.driver.find_element_by_id('memberno').click()
        a = self.driver.find_element_by_css_selector('span.g-c-red4').text
        try:
            self.assertEqual(a,'推荐人卡号不能为空')
        except Exception:
            self.driver.save_screenshot(jt())
            raise

    #经销商卡号为空
    def test12(self):
        self.register()
        self.driver.find_element_by_id('referenceno').send_keys('5475966')
        self.driver.find_element_by_id('memberno').click()
        self.driver.find_element_by_id('dealername').click()
        a = self.driver.find_element_by_css_selector('span.g-c-red4').text
        try:
            self.assertEqual(a,'卡号不能为空')
        except Exception:
            self.driver.save_screenshot(jt())
            raise

    #经销商名字为空
    def test13(self):
        self.register()
        self.driver.find_element_by_id('dealername').click()
        self.driver.find_element_by_id('idcard').click()
        a = self.driver.find_element_by_css_selector('span.g-c-red4').text
        try:
            self.assertEqual(a,'经销商姓名不能为空')
        except Exception:
            self.driver.save_screenshot(jt())
            raise

    #身份证为空
    def test14(self):
        self.register()
        self.driver.find_element_by_id('idcard').click()
        self.driver.find_element_by_id('mobile').click()
        a = self.driver.find_element_by_css_selector('span.g-c-red4').text
        try:
            self.assertEqual(a,'身份证不能为空')
        except Exception:
            self.driver.save_screenshot(jt())
            raise

    #手机号码为空
    def test15(self):
        self.register()
        self.driver.find_element_by_id('mobile').click()
        self.driver.find_element_by_id('mobile2').click()
        a = self.driver.find_element_by_css_selector('span.g-c-red4').text
        try:
            self.assertEqual(a,'手机号码不能为空')
        except Exception:
            self.driver.save_screenshot(jt())
            raise

    #错误的联系人卡号
    def test16(self):
        self.register()
        self.driver.find_element_by_id('referenceno').send_keys('45645')
        self.driver.find_element_by_id('dealername').click()
        a =self.driver.find_element_by_id('recommendCardNoName').text
        try:
            self.assertEqual(a,'adas卡号错误或紧急联系人不存在。')
        except Exception:
            self.driver.save_screenshot(jt())
            raise

    #错误的经销商名字
    def test17(self):
        self.register()
        self.driver.find_element_by_id('dealername').send_keys('1121')
        self.driver.find_element_by_id('idcard').click()
        a = self.driver.find_element_by_css_selector('span.g-c-red4').text
        try:
            self.assertEqual(a,'经销商姓名必须是汉字或 点组成,2-20个字')
        except Exception:
            self.driver.save_screenshot(jt())
            raise

    #错误的身份证
    def test18(self):
        self.register()
        self.driver.find_element_by_id('idcard').send_keys('1121')
        self.driver.find_element_by_id('mobile').click()
        a =self.driver.find_element_by_id('paperNoMessage').text
        try:
            self.assertEqual(a,'证件号错误！')
        except Exception:
            self.driver.save_screenshot(jt())
            raise

    #错误的手机号
    def test19(self):
        self.register()
        self.driver.find_element_by_id('mobile').send_keys('5475966')
        self.driver.find_element_by_id('mobile2').click()
        a = self.driver.find_element_by_css_selector('span.g-c-red4').text
        try:
            self.assertEqual(a,'手机号码格式不合法')
        except Exception:
            self.driver.save_screenshot(jt())
            raise

    #错误的卡号
    def test20(self):
        self.register()
        self.driver.find_element_by_id('referenceno').send_keys('5475966')
        self.driver.find_element_by_id('memberno').send_keys('66666666')
        self.driver.find_element_by_id('dealername').send_keys(mz())
        self.driver.find_element_by_id('idcard').send_keys(sfz())
        self.driver.find_element_by_id('mobile').send_keys(sjh())
        self.driver.find_element_by_id('regeditBtn').click()
        a = WebDriverWait(self.driver, 5, 0.2).until(lambda x: x.find_element_by_xpath('/html/body/div[6]/div/table/tbody/tr[2]/td/div/span')).text
        try:
            self.assertEqual(a,'对不起，您填写的卡号不存在,请重新输入')
        except Exception:
            self.driver.save_screenshot(jt())
            raise

    #被注册过的身份证
    def test21(self):
        self.register()
        self.driver.find_element_by_id('idcard').send_keys('220102199003073417')
        self.driver.find_element_by_id('mobile').click()
        a = self.driver.find_element_by_id('paperNoMessage').text
        try:
            self.assertEqual(a,'您使用的身份证已被注册！')
        except Exception:
            self.driver.save_screenshot(jt())
            raise

    #被注册过的手机号
    def test22(self):
        self.register()
        self.driver.find_element_by_id('referenceno').send_keys('5475966')
        self.driver.find_element_by_xpath('/html/body/div[3]/ul/li/form/ul/li[1]/div[2]/div/a').click()
        self.driver.find_element_by_xpath('/html/body/div[3]/ul/li/form/ul/li[1]/div[2]/div/div/div/div/div[2]/ul/li[1]/input').click()
        self.driver.find_element_by_id('madalbtn').click()
        self.driver.find_element_by_id('dealername').send_keys(mz())
        self.driver.find_element_by_id('idcard').send_keys(sfz())
        self.driver.find_element_by_id('mobile').send_keys('13168090784')
        self.driver.find_element_by_id('regeditBtn').click()
        a = WebDriverWait(self.driver, 5, 0.2).until(lambda x: x.find_element_by_xpath('/html/body/div[6]/div/table/tbody/tr[2]/td/div/span')).text
        try:
            self.assertEqual(a,'此手机号已经注册，请重新填写')
        except Exception:
            self.driver.save_screenshot(jt())
            raise

    #选择卡号注册
    def test23(self):
        self.register()
        self.driver.find_element_by_id('referenceno').send_keys('5475966')
        self.driver.find_element_by_xpath('/html/body/div[3]/ul/li/form/ul/li[1]/div[2]/div/a').click()
        self.driver.find_element_by_xpath('/html/body/div[3]/ul/li/form/ul/li[1]/div[2]/div/div/div/div/div[2]/ul/li[2]/input').click()
        self.driver.find_element_by_id('madalbtn').click()
        a = mz()
        self.driver.find_element_by_id('dealername').send_keys(a)
        self.driver.find_element_by_id('idcard').send_keys(sfz())
        self.driver.find_element_by_id('mobile').send_keys(sjh())
        self.driver.find_element_by_id('regeditBtn').click()
        time.sleep(2)
        hm = WebDriverWait(self.driver, 5, 0.2).until(
            lambda x: x.find_element_by_xpath('/html/body/div[3]/ul/li/div/div/table/tbody/tr[1]/td[3]')).text
        try:
            self.assertEqual(hm,a)
        except Exception:
            self.driver.save_screenshot(jt())
            raise

    #输入卡号注册
    def test24(self):
        self.register()
        self.driver.find_element_by_id('referenceno').send_keys('5475966')
        self.driver.find_element_by_id('memberno').send_keys('80222615')
        a = mz()
        self.driver.find_element_by_id('dealername').send_keys(a)
        self.driver.find_element_by_id('idcard').send_keys(sfz())
        self.driver.find_element_by_id('mobile').send_keys(sjh())
        self.driver.find_element_by_id('regeditBtn').click()
        time.sleep(2)
        hm = WebDriverWait(self.driver, 5, 0.2).until(
            lambda x: x.find_element_by_xpath('/html/body/div[3]/ul/li/div/div/table/tbody/tr[1]/td[3]')).text
        try:
            self.assertEqual(hm,a)
        except Exception:
            self.driver.save_screenshot(jt())
            raise


if __name__ == '__main__':
    unittest.main()
