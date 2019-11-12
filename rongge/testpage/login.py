import unittest
import time
from rongge.testpage.models.myunit import  MyunitTest
from rongge.tool.tool import js,jt,yzm

class TestLogin(MyunitTest):

    def login(self,user,paw,*u):
        i = js()
        if i==1:
            self.driver.find_element_by_id('header_loginbtn').click()
            time.sleep(2)
        self.driver.find_element_by_id('userCode').send_keys(user)
        self.driver.find_element_by_id('login_storepassword').send_keys(paw)
        a = self.driver.find_element_by_id('safecode').get_attribute('src')
        c = self.driver.get_cookie('JSESSIONID')['value']
        b = yzm(a, c)
        if u ==():
            self.driver.find_element_by_id('captchaCode').send_keys(b)
        else:
            self.driver.find_element_by_id('captchaCode').send_keys(u)
        self.driver.find_element_by_xpath('/html/body/div[2]/div/div/form/ul/li[2]/div[2]/button').click()
        time.sleep(2)

    #不存在的经销商
    def test01(self):
        self.login('555333','12312')
        a = self.driver.find_element_by_xpath('/html/body/div[2]/div/div/form/ul/li[2]/div[2]/div/span').text
        print(a)
        try:
            self.assertEqual(a,'asdas经销商编号对应的经销商不存在')
        except Exception:
            self.driver.save_screenshot(jt())
            raise
        time.sleep(2)

    #错误的验证码
    def test02(self):
        self.login('AH5012','12345678','15646')
        b = self.driver.find_element_by_xpath('/html/body/div[2]/div/div/form/ul/li[2]/div[2]/div/span').text
        self.assertEqual(b,'验证码不正确!')
        time.sleep(2)

    #错误的密码
    def test03(self):
        self.login('AH5012','456456')
        b = self.driver.find_element_by_xpath('/html/body/div[2]/div/div/form/ul/li[2]/div[2]/div/span').text
        self.assertEqual(b,'用户帐号或登录密码错误')
        time.sleep(2)

    #成功登陆
    def test04(self):
        self.login('AH5012','12345678')
        time.sleep(3)
        s = self.driver.find_element_by_xpath('/html/body/div[1]/div[1]/div[5]/ul/li/span').text
        self.assertEqual(s, '您好,杨光（#）店长!欢迎登录,祝您生活愉快!')

if __name__ == '__main__':
    unittest.main()








