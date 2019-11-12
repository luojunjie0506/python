import requests
import random
import time
from lxml import etree
import  os
import csv

jihe = {}
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
    #cs(sjh,sj)
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
    #cs(sfz, id[b][0])
    return id[b][0]

#名字
def mz():
    global name
    a = random.randint(2, 5)
    name = ''.join(random.sample('你是撒大声地大神啊大苏打', a))
    #cs(mz, name)
    return name

#计数器
i = 0
def js():
    global i
    i = i + 1
    return i

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
    return path

#记录注册信息
def cs(a,b,*u):
    jihe[a] = b
    datas = [jihe]
    if u == 1:
        xr(datas)

#写入csv
def xr(a):
    headers = ['mz', 'sjh', 'sfz']
    now = time.strftime('%Y-%m-%d', time.localtime(time.time()))
    test_report_path = os.path.join(os.getcwd() + '\\report\\' + now + '.json')
    with open(test_report_path, 'a', newline='') as f:
        # 标头在这里传入，作为第一行数据
        writer = csv.DictWriter(f, headers)
        writer.writeheader()
        for row in a:
            writer.writerow(row)

#读取csv
def dq(a):
    now = time.strftime('%Y-%m-%d', time.localtime(time.time()))
    test_report_path = os.path.join(os.getcwd() + '\\report\\' + now + '.json')
    with open(test_report_path, "r") as f:
        reader = csv.DictReader(f)
        print(type(reader))
        print(reader)
        for row in reader:
            print(row[a])
