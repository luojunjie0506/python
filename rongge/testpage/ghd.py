import unittest
from selenium import webdriver
import time
from rongge.testpage.models.myunit import MyunitTest
from rongge.tool.tool import yzm
from selenium.webdriver.support.ui import WebDriverWait
from pymouse import PyMouse


class TestGhd(MyunitTest):
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

    def shop(self):
        self.driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div/ul/li[3]/a').click()


    #搜索框功能
    def test32(self):
        self.shop()
        self.driver.find_element_by_id('findtxt').send_keys('K-C007')
        self.driver.find_element_by_id('findBtn').click()
        time.sleep(1)
        text1 = self.driver.find_element_by_id('orderQty_K-C007').get_attribute('style')
        self.assertIn('background',text1)

    #不选择商品直接下单
    def test33(self):
        self.shop()
        time.sleep(2)
        tijiao = self.driver.find_element_by_xpath('/html/body/div[3]/form/div[3]/button[3]')
        self.driver.execute_script("arguments[0].click();", tijiao)
        time.sleep(2)
        text1 = self.driver.switch_to.alert.text
        self.driver.switch_to.alert.accept()
        self.assertEqual(text1,'至少选择一个商品数量大于0的商品!')

    #保存订单功能
    def test34(self):
        self.shop()
        self.driver.find_element_by_id('recName').clear()
        self.driver.find_element_by_id('recName').send_keys('ljj')
        self.driver.find_element_by_id('orderQty_K-W001').send_keys('5')
        self.driver.find_element_by_id('orderQty_K-W002').send_keys('5')
        time.sleep(1)
        self.driver.find_element_by_id('amount').click()
        time.sleep(1)
        bc = self.driver.find_element_by_id('save')
        self.driver.execute_script("arguments[0].click();",bc)
        time.sleep(2)
        self.driver.switch_to.window(self.driver.window_handles[1])
        time.sleep(2)
        self.driver.find_element_by_id('module_poOrderSaleList').click()
        text = WebDriverWait(self.driver,5,0.2).until(lambda x:x.find_element_by_xpath('/html/body/div[3]/ul/li/div[2]/table/tbody/tr[1]/td[2]')).text
        self.assertEqual(text,'ljj')

    #支付已保存订单
    def test35(self):
        self.shop()
        self.driver.find_element_by_id('module_poOrderSaleList').click()
        self.driver.find_element_by_xpath('/html/body/div[3]/ul/li/div[2]/table/tbody/tr/td[11]/a[1]').click()
        WebDriverWait(self.driver,5,0.2).until(self.driver.find_element_by_link_text('充值账户余额：')).click()
        fk = self.driver.find_element_by_xpath('/html/body/div[3]/form/div[3]/button[3]')
        self.driver.execute_script('arguments[0].click():',fk)
        self.driver.find_element_by_id('selfPay').click()
        self.driver.find_element_by_id('payPassword').send_keys('12345678')
        self.driver.find_element_by_id('formBtnOk').click()

    #删除待支付订单选为否
    def test36(self):
        self.shop()
        self.driver.find_element_by_id('module_poOrderSaleList').click()
        text = self.driver.find_element_by_xpath('/html/body/div[3]/ul/li/div[2]/table/tbody/tr[1]/td[2]').text
        self.driver.find_element_by_xpath('/html/body/div[3]/ul/li/div[2]/table/tbody/tr[1]/td[11]/a[2]').click()
        self.driver.switch_to.alert.dismiss()
        time.sleep(2)
        text1 = self.driver.find_element_by_xpath('/html/body/div[3]/ul/li/div[2]/table/tbody/tr[1]/td[2]').text
        self.assertEqual(text,text1)

    #删除待支付订单选为是
    def test37(self):
        self.shop()
        self.driver.find_element_by_id('module_poOrderSaleList').click()
        text = self.driver.find_element_by_xpath('/html/body/div[3]/ul/li/div[2]/table/tbody/tr[1]/td[2]').text
        self.driver.find_element_by_xpath('/html/body/div[3]/ul/li/div[2]/table/tbody/tr[1]/td[11]/a[2]').click()
        self.driver.switch_to.alert.accept()
        time.sleep(2)
        text1 = self.driver.find_element_by_xpath('/html/body/div[3]/ul/li/div[2]/table/tbody/tr[1]/td[2]').text
        self.assertNotEqual(text,text1)

    #找他人支付错误店号
    def test38(self):
        self.shop()
        #新建鼠标对象
        mouse = PyMouse()
        self.driver.find_element_by_id('module_poOrderSale1').click()
        WebDriverWait(self.driver, 5, 0.2).until(lambda x: x.find_element_by_id('orderQty_K-W001')).send_keys('5')
        self.driver.find_element_by_id('orderQty_K-W002').send_keys('5')
        fk = self.driver.find_element_by_xpath('/html/body/div[3]/form/div[3]/button[2]')
        self.driver.execute_script('arguments[0].click();',fk)
        self.driver.find_element_by_id('NewcardNo').send_keys('asdasd')
        #移动坐标
        mouse.move(670, 437)
        time.sleep(1)
        mouse.move(500, 500)
        time.sleep(1)
        text = WebDriverWait(self.driver, 5, 0.2).until(lambda x: x.find_element_by_xpath(
            '/html/body/div[3]/form/div[4]/div[2]/table/tbody/tr[3]/td/span/font')).text
        self.assertEqual(text,'专卖店不存在或错误')

    #找他人支付非正常状态店号
    def test39(self):
        self.shop()
        mouse = PyMouse()
        self.driver.find_element_by_id('module_poOrderSale1').click()
        WebDriverWait(self.driver,5,0.2).until(lambda x:x.find_element_by_id('orderQty_K-W001')).send_keys('5')
        self.driver.find_element_by_id('orderQty_K-W002').send_keys('5')
        fk = self.driver.find_element_by_xpath('/html/body/div[3]/form/div[3]/button[2]')
        self.driver.execute_script('arguments[0].click();',fk)
        self.driver.find_element_by_id('NewcardNo').send_keys('GD5555')
        mouse.move(670, 437)
        time.sleep(1)
        mouse.move(500, 500)
        time.sleep(1)
        text = WebDriverWait(self.driver,5,0.2).until(lambda x:x.find_element_by_xpath('/html/body/div[3]/form/div[4]/div[2]/table/tbody/tr[3]/td/span/font')).text
        self.assertEqual(text,'不是正常状态的专卖店')

    #找他人支付功能输入本店
    def test40(self):
        self.shop()
        self.driver.find_element_by_id('module_poOrderSale1').click()
        self.driver.find_element_by_id('orderQty_M048').send_keys('5')
        self.driver.find_element_by_id('orderQty_M049').send_keys('5')
        self.driver.find_element_by_xpath('/html/body/div[3]/form/div[3]/button[2]').click()
        self.driver.find_element_by_id('NewcardNo').send_keys('AH5012')
        self.driver.find_element_by_id('orderQty_ZL112').click()
        time.sleep(2)
        self.driver.switch_to.window(self.driver.window_handles[1])
        time.sleep(2)
        text = self.driver.find_element_by_xpath('/html/body/div[9]/div/table/tbody/tr[2]/td/div/span').text
        self.assertEqual(text,'订单代付专卖店不能为自己!')


    #找他人支付功能
    def test(self):
        self.shop()
        self.driver.find_element_by_id('module_poOrderSale1').click()
        self.driver.find_element_by_id('orderQty_M048').send_keys('5')
        self.driver.find_element_by_id('orderQty_M049').send_keys('5')
        self.driver.find_element_by_xpath('/html/body/div[3]/form/div[3]/button[2]').click()
        self.driver.find_element_by_id('NewcardNo').send_keys('GD5555')



    #撤销代支付人
    def test(self):
        pass

    #不满3000提示
    def test(self):
        pass

    #自付运费错误密码
    def test(self):
        pass

    #满3000包邮错误密码
    def test(self):
        pass

    #支付不满3000订单
    def test(self):
        pass

    #支付满3000订单
    def test(self):
        pass


if __name__ == '__main__':
    unittest.main()