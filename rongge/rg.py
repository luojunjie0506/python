from selenium import webdriver
import time
import random
import requests
from lxml import etree
from selenium.webdriver.support.ui import Select
import  os

#获取验证码
def yzm(w, c):
    headers = {'Host': 'cswl2016.3322.org:8081',
               'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:69.0) Gecko/20100101 Firefox/69.0',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                'Cookie': 'cUserCode = "";cRememberMe = "";JSESSIONID = 0F746537AD71ED915CFDF9FBE079E909.213',
                'Upgrade - Insecure - Requests': '1', }
    headers['Cookie'] = 'cUserCode = "";cRememberMe = "";JSESSIONID = ' + c
    response = requests.get(w, headers=headers)
    head = response.headers
    b = head['captchaCode']
    return b

#手机号
def sjh():
    global sj
    a=['131','181','134','135','136','138']
    b=random.randint(0, 5)
    sj = ''.join(random.sample('123456789', 8))
    sj = a[b] + sj
    return sj

#身份证
def sfz():
    global  id
    r = requests.get('http://www.360doc.com/content/19/0228/08/23109265_818055024.shtml')
    h = etree.HTML(r.text)
    f=random.randint(1, 10)
    q=random.randint(1, 1)
    aa = '//*[@id="artContent"]/div[2]/div[1]/div['+str(f)
    bb= ']/text()['+str(q)
    qym= h.xpath(aa+bb+']')
    # 出生日期
    csr = 19000000 + random.randint(5, 9) * 100000 + random.randint(0, 9) * 10000 + random.randint(1,12) * 100 + random.randint(1, 28)
    # 性别
    sex = random.randint(1, 2)
    num = random.randint(10, 80)
    url = 'http://sfz.uzuzuz.com/'
    response = requests.get(url,
    params={
    'region': qym[0].strip(),
    'birthday': csr,
    'sex': sex,
    'num': 20,
    'r': num
                            })
    html = etree.HTML(response.text)
    id = []
    s = html.xpath('/html/body/div[2]/div[1]/div/div[2]/table[2]/tbody/tr[3]/td/text()')
    for i in range(1, 20):
        a = html.xpath('/html/body/div[2]/div[1]/div/div[2]/table[2]/tbody/tr[' + str(i) + ']/td/text()')
        id.append(a)
    b=  random.randint(1, 10)
    return id[b][0]

#名字
def mz():
    global name
    a = random.randint(2, 5)
    name = ''.join(random.sample('你是撒大声地大神啊大苏打', a))
    return name

#截图
def jt():
    #日期文件夹
    date1 = time.localtime()
    a = ''.join([str(date1.tm_year), str(date1.tm_mon), str(date1.tm_mday)])
    #日期文件名
    b = '-'.join([str(date1.tm_hour), str(date1.tm_min), str(date1.tm_sec)])
    dateDir = os.path.join(os.getcwd() + '\\image\\', a)
    # 如果当前日期目录不存的话就创建
    if not os.path.exists(dateDir):
        os.mkdir(dateDir)
    path = os.path.join(dateDir, b + '.png')
    driver.get_screenshot_as_file(path)

#登陆
def login():
    global  driver
    driver = webdriver.Firefox()
    driver.get('http://cswl2016.3322.org:8081/member/index.jhtml')
    driver.maximize_window()
    time.sleep(2)
    driver.find_element_by_id('header_loginbtn').click()
    driver.find_element_by_id('userCode').send_keys('AH5012')
    driver.find_element_by_id('login_storepassword').send_keys('12345678')
    a = driver.find_element_by_id('safecode').get_attribute('src')
    c = driver.get_cookie('JSESSIONID')['value']
    b = yzm(a,c)
    driver.find_element_by_id('captchaCode').send_keys(b)
    driver.find_element_by_xpath('/html/body/div[2]/div/div/form/ul/li[2]/div[2]/button').click()
    jt()

#申请电子卡
def sq():
    global bh
    driver.find_element_by_xpath('//*[@id="mainmenu"]/li[2]/a').click()
    time.sleep(2)
    driver.find_element_by_id('module_miAgentCardList').click()
    driver.find_element_by_id('nongye5_label').click()
    num = random.randint(1, 1)
    driver.find_element_by_id('idNummber').send_keys(num)
    driver.find_element_by_id('yfmoney').click()
    a= driver.find_element_by_id('yfmoney').text
    if int(a) == num*5 :
        driver.find_element_by_id('numberpaper').send_keys('12345678')
        driver.find_element_by_xpath('/html/body/div[3]/ul/li/div/form/div/div/div/div/div/ul/li/div[2]/span[3]/input[3]').click()
        time.sleep(1)
        ss = driver.find_element_by_xpath('/html/body/div[6]/div/table/tbody/tr[2]/td/div/span').text
        if str(num) in ss :
            print('购买成功！')
            jt()
            driver.find_element_by_id('module_poAgentCardList').click()
            time.sleep(2)
            bh = driver.find_element_by_xpath('/html/body/div[3]/ul/li/div[2]/table/tbody/tr[1]/td[3]').text
            print(bh)

#注册
def regist(*u):
    global name,sfz
    driver.find_element_by_id('module_agentNewMember').click()
    time.sleep(1)
    driver.find_element_by_id('referenceno').send_keys('5475966')
    #判断传参个数
    if u == ():
        driver.find_element_by_xpath('//*[@id="register_wrap"]/li[1]/div[2]/div/a').click()
        driver.find_element_by_xpath('//*[@id="check-box-list"]/li[1]/input').click()
        driver.find_element_by_id('madalbtn').click()
    else:
        driver.find_element_by_id('memberno').send_keys(u[0])
    name = mz()
    driver.find_element_by_id('dealername').send_keys(name)
    sfz1 =sfz()
    driver.find_element_by_id('idcard').send_keys(sfz1)
    sjh1 = sjh()
    driver.find_element_by_id('mobile').send_keys(sjh1)
    driver.find_element_by_id('regeditBtn').click()
    time.sleep(2)
    xm = driver.find_element_by_xpath('/html/body/div[7]/div/table/tbody/tr[2]/td/div/span/p[3]').text
    hm = driver.find_element_by_xpath('/html/body/div[3]/ul/li/div/div/table/tbody/tr[1]/td[4]').text
    if  name in xm:
        jt()
        print('注册成功！')

#查询
def cx():
    #卡号查询
    hm = driver.find_element_by_xpath('/html/body/div[3]/ul/li/div/div/table/tbody/tr[1]/td[4]').text
    driver.find_element_by_id('card_no').send_keys(hm)
    driver.find_element_by_id('grid_querybtn').click()
    ss = driver.find_element_by_xpath('/html/body/div[3]/ul/li/div/div/table/tbody/tr/td[4]').text
    if ss == hm:
        jt()
        print('卡号查询无异常！')
    #身份证查询
'''
    driver.find_element_by_id('paper_no').send_keys(sfz)
    driver.find_element_by_id('grid_querybtn').click()
    ss = driver.find_element_by_xpath('/html/body/div[3]/ul/li/div/div/table/tbody/tr/td[3]').text
    if name in ss :
        print('身份证查询无异常！')
    #姓名查询
    driver.find_element_by_id('name').send_keys(name)
    driver.find_element_by_id('grid_querybtn').click()
    ss = driver.find_element_by_xpath('/html/body/div[3]/ul/li/div/div/table/tbody/tr/td[3]').text
    if name in ss :
        print('姓名查询无异常！')
    # 手机号查询
    driver.find_element_by_id('mobile').send_keys(sjh)
    driver.find_element_by_id('grid_querybtn').click()
    ss = driver.find_element_by_xpath('/html/body/div[3]/ul/li/div/div/table/tbody/tr/td[8]').text
    if sjh == ss :
        print('手机号查询无异常！')
    #状态+手机号查询
    driver.find_element_by_id('mobile').send_keys(sjh)
    s1 = Select(driver.find_element_by_id('status'))
    s1.select_by_index(1)
    ss = driver.find_element_by_xpath('/html/body/div[3]/ul/li/div/div/table/tbody/tr/td[8]').text
    if sjh == ss :
        print('状态+手机号查询无异常！')
    #状态+手机号查询
    driver.find_element_by_id('mobile').send_keys(sjh)
    s1 = Select(driver.find_element_by_id('status'))
    s1.select_by_index(2)
    ss = driver.find_element_by_xpath('/html/body/div[3]/ul/li/div/div/table/tbody/tr/td[8]').text
    if sjh != ss :
        print('状态+手机号查询无异常！')
'''

#经销商作废
def sc():
    driver.find_element_by_id('module_applyManageList').click()
    time.sleep(2)
    driver.find_element_by_xpath('/html/body/div[3]/ul/li/div/div/table/tbody/tr[1]/td[1]/input').click()
    driver.find_element_by_id('cancel_sure').click()
    time.sleep(2)
    bb = driver.find_element_by_xpath('/html/body/div[7]/div/table/tbody/tr[2]/td/div/span').text
    if  name in bb:
        jt()
        print('删除成功!')

#经销商激活
def jh():
    regist(bh)
    s1 = Select(driver.find_element_by_id('status'))
    s1.select_by_index(1)
    driver.find_element_by_id('grid_querybtn').click()
    time.sleep(1)
    jj = driver.find_element_by_xpath('/html/body/div[3]/ul/li/div/div/table/tbody/tr[1]/td[3]').text
    if  name in jj:
        jt()
        driver.find_element_by_xpath('/html/body/div[3]/ul/li/div/div/table/tbody/tr[1]/td[1]').click()
        driver.find_element_by_id('ok_sure').click()
        time.sleep(1)

#购货单
def ghd():
    driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div/ul/li[3]/a').click()
    time.sleep(1)
    a = random.randint(1,5)
    driver.find_element_by_id('orderQty_K-C007').send_keys(a)
    time.sleep(2)
    tijiao = driver.find_element_by_name('paymentMethod')
    driver.execute_script("arguments[0].click();", tijiao)
    driver.find_element_by_xpath('/html/body/a[2]').click()
    time.sleep(1)
    element = driver.find_element_by_xpath('/html/body/div[3]/form/div[3]/button[3]')
    driver.execute_script("arguments[0].click();", element)
    time.sleep(1)
    driver.find_element_by_id('otherPay').click()
    driver.find_element_by_id('payPassword').send_keys('12345678')
    driver.find_element_by_id('formBtnOk').click()
    time.sleep(1)
    jt()

#业绩申报
def yjsb():
    time.sleep(1)
    #页面跳转
    driver.switch_to.window(driver.window_handles[0])
    time.sleep(2)
    driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div/ul/li[4]/a').click()
    time.sleep(1)
    #后期使用上面注册的经销商bh
    driver.find_element_by_id('memberNo').send_keys(bh)
    driver.find_element_by_id('inputAddr_FV').send_keys('11')
    driver.find_element_by_id('btnsave').click()
    time.sleep(1)
    #操作alter窗口
    driver.switch_to.alert.accept()
    time.sleep(1)
    jt()
    a = driver.find_element_by_xpath('/html/body/div[6]/div/table/tbody/tr[2]/td/div/span').text
    if '申报成功' in a :
        print('申报成功')

#购货业绩单
def ghyjd():
    driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div/ul/li[5]/a').click()
    time.sleep(1)
    driver.find_element_by_xpath('/html/body/div[3]/ul/li[2]/ul/li[2]/div[1]/div/span').click()
    driver.find_element_by_xpath('/html/body/div[3]/ul/li[2]/ul/li[3]/div[1]/div/span').click()
    driver.find_element_by_xpath('/html/body/ul/li[2]/a').click()
    time.sleep(1)
    driver.find_element_by_id('shopping_buy').click()
    time.sleep(1)
    driver.find_element_by_id('cardNo').send_keys(bh)
    time.sleep(1)
    tijiao = driver.find_element_by_xpath('/html/body/div[3]/form/div[2]/table/tbody/tr[5]/td[2]/input')
    driver.execute_script("arguments[0].click();", tijiao)
    time.sleep(1)
    element = driver.find_element_by_xpath('/html/body/div[3]/form/div[3]/button[3]')
    driver.execute_script("arguments[0].click();", element)
    time.sleep(1)
    driver.find_element_by_id('otherPay').click()
    driver.find_element_by_id('payPassword').send_keys('12345678')
    driver.find_element_by_id('formBtnOk').click()
    time.sleep(1)
    jt()
    time.sleep(1)
    driver.switch_to.window(driver.window_handles[0])
    time.sleep(1)

#全拓
def qt():
    driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div/ul/li[6]/a').click()
    time.sleep(1)
    driver.find_element_by_id('cardNo').send_keys(bh)
    time.sleep(1)
    driver.find_element_by_id('orderQty_K-C016').send_keys('40')
    driver.find_element_by_id('zp').click()
    driver.find_element_by_id('orderQty2_AF011新').send_keys('100')
    driver.find_element_by_xpath('/html/body/a[2]').click()
    time.sleep(1)
    tijiao = driver.find_element_by_xpath('/html/body/div[3]/form/div[2]/table[3]/tbody/tr/td[2]/input')
    driver.execute_script("arguments[0].click();", tijiao)
    time.sleep(1)
    element = driver.find_element_by_xpath('//*[@id="check"]')
    driver.execute_script("arguments[0].click();", element)
    time.sleep(1)
    driver.find_element_by_id('payPassword').send_keys('12345678')
    driver.find_element_by_id('formBtnOk').click()
    time.sleep(1)
    jt()
    time.sleep(1)
    driver.switch_to.window(driver.window_handles[0])

login()
sq()
regist()
cx()
sc()
jh()
ghd()
yjsb()
ghyjd()
qt()


#页面报错的截图后退