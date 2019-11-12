import  unittest
from rongge.testpage.login import TestLogin
from rongge.testpage.sq  import TestSq
from rongge.testpage.register  import TestRegister
from rongge.testpage.operation import TestOperation
from rongge.testpage.ghd import TestGhd
import  os
import time
import HTMLTestRunner

class RunTcScript(object):
    def __init__(self):
        self.suite = unittest.TestSuite()

    #载入登陆模块
    def load_login(self,testcase):
        self.suite.addTest(TestLogin(testcase))

    #载入申请模块
    def load_sq(self,testcase):
        self.suite.addTest(TestSq(testcase))

    #载入注册模块
    def load_register(self,testcase):
        self.suite.addTest(TestRegister(testcase))

    #载入注册模块
    def load_operation(self,testcase):
        self.suite.addTest(TestOperation(testcase))

    #载入购货单模块
    def load_ghd(self,testcase):
        self.suite.addTest(TestGhd(testcase))

if __name__ =='__main__':
    suite_tc = RunTcScript()

    '''
    #登陆用例
    suite_tc.load_login('test01')
    suite_tc.load_login('test02')
    suite_tc.load_login('test03')
    suite_tc.load_login('test04')
    # 申请电子卡用例
    suite_tc.load_sq('test05')
    suite_tc.load_sq('test06')
    suite_tc.load_sq('test07')
    suite_tc.load_sq('test08')
    suite_tc.load_sq('test09')
    suite_tc.load_sq('test10')
    #注册用例
    suite_tc.load_register('test11')
    suite_tc.load_register('test12')
    suite_tc.load_register('test13')
    suite_tc.load_register('test14')
    suite_tc.load_register('test15')
    suite_tc.load_register('test16')
    suite_tc.load_register('test17')
    suite_tc.load_register('test18')
    suite_tc.load_register('test19')
    suite_tc.load_register('test20')
    suite_tc.load_register('test21')
    suite_tc.load_register('test22')   
    suite_tc.load_register('test23')
    suite_tc.load_register('test24')
          
    #经销商操作用例
    suite_tc.load_operation('test25')
    suite_tc.load_operation('test26')
    suite_tc.load_operation('test27')
    suite_tc.load_operation('test28')
    suite_tc.load_operation('test29')
    suite_tc.load_operation('test30')
    suite_tc.load_operation('test31')

    #购货单用例
    suite_tc.load_ghd('test32')
    suite_tc.load_ghd('test33')

    suite_tc.load_ghd('test34')
    suite_tc.load_ghd('test35')
    suite_tc.load_ghd('test36')
    suite_tc.load_ghd('test37')
    suite_tc.load_ghd('test40')
                '''

    suite_tc.load_ghd('test38')
    suite_tc.load_ghd('test39')






    now =time.strftime('%Y-%m-%d', time.localtime(time.time()))
    test_report_path = os.path.join(os.getcwd()+'\\report\\'+now+'.html')
    with open(test_report_path, "wb") as report_file:
        runner = HTMLTestRunner.HTMLTestRunner(stream=report_file, title="自动化测试报告", description="XX应用功能测试")
        runner.run(suite_tc.suite)