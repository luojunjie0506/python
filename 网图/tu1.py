import pandas as pd
import numpy as np
from graphviz import Digraph


path = 'D:\\'  #文档路径
js = 1   #层数
a = 0
dy = []     #每层经销商
btlist = [] #表头字典

#值转为列表
def zh(aa):
    citydaima = np.array(aa)
    citydaima = citydaima.tolist()
    for i in citydaima:
        for j in i:
            return j

def xh(zz):
    global js,a
    print(zz)
    f = []
    #循环表头
    bt = date.ix[0].values
    #表头长度
    btL = len(bt)
    if len(f)==0:
        for i in range(0,btL):
            li = bt[i].strip()
            btlist.append(li)

    #循环整个文档行
    for i in range(1, hs):
        zw = zh(date.loc[[i], [1]]) #执委
        hy = zh(date.loc[[i], [2]])  # 会员编号
        name = zh(date.loc[[i], [3]])  # 姓名
        age = zh(date.loc[[i], [4]])  # 年龄
        sex = zh(date.loc[[i], [5]])  # 性别
        leg = zh(date.loc[[i], [6]]) #月均2.8万业绩leg数
        qg = zh(date.loc[[i], [7]]) #19年业绩（切割）
        bqg = zh(date.loc[[i], [8]]) #19年业绩
        tjr = zh(date.loc[[i], [9]])    #紧急联系人卡号

        #取第一层
        if hy == zz:
            A = zw + '-' + hy + ' ' + name + '\n' + '年龄:' + str(age) + ' 性别:' + sex + '\n' + '月均2.8万业绩leg数:' + str(leg) + '\n' + '19年业绩(切割):' + str('%.1f' % (qg / 10000)) + '万\n' + '19年业绩:' + str('%.1f' % (bqg / 10000)) + '万\n'
            filename = zw
        # 查询每层符合条件的人
        if tjr in zz:
            #网图每个框显示的内容
            B = zw + '-' + hy + ' ' + name + '\n' + '年龄:' + str(age) + ' 性别:' + sex + '\n' + '月均2.8万业绩leg数:' + str(
                leg) + '\n' + '19年业绩(切割):' + str('%.1f' % (qg / 10000)) + '万\n' + '19年业绩:' + str(
                '%.1f' % (bqg / 10000)) + '万\n'
            f.append(hy)
            #绘制整个网图
            if True:
                if js==1:
                    if sex =="男":
                        dot.node(tjr, A, fontname="SimHei", color='LightSkyBlue')
                    else:
                        dot.node(tjr, A, fontname="SimHei", color='Pink')
                if sex =="男":
                    dot.node(tjr, fontname="SimHei")
                    dot.node(hy, B, fontname="SimHei",color ='LightSkyBlue')
                    dot.edge(tjr, hy)
                else:
                    dot.node(tjr, fontname="SimHei")
                    dot.node(hy, B, fontname="SimHei",color ='Pink')
                    dot.edge(tjr, hy)
    #判断
    if f != zz[0]:
        js = js + 1
        print(js)
        xh(f)
    else:
        dot.render('test-output/'+filename+'.gv', view=True)


cc = input('请输入要查询的卡号:')
dd = input('请输入你的文件名(存储在D盘根目录):')
path = path +dd
#打开excel
date = pd.read_excel(path,header=None)
dy.append(cc)
#创建网图
dot = Digraph(name='shop',node_attr={ 'style': 'filled','shape':'box'})
#获取行数
hs = len(date)
xh(dy)

