import pandas as pd
import numpy as np
from graphviz import Digraph


path = 'D:\\'  #文档路径
js = 1   #层数
dy = []     #第一层经销商

yanse = 4


#新市场11月申报业绩_20191121 14点43.xlsx
# 0881198
#值转为列表
def zh(aa):
    citydaima = np.array(aa)
    citydaima = citydaima.tolist()
    for i in citydaima:
        for j in i:
            return j

def xh():
    global js,yanse
     # 存储相同L的行
    #循环整个文档行
    js = 0
    for l in range(1, 5):
        ls = []
        # 存储相同L的行
        for i in range (1,hs):
            cs = zh(date.loc[[i], [9]])  # 层数
            if l == cs :
                ls.append(i)
        print (ls)
        a = len(ls)
        for i  in range(1,a):
            tjr = zh(date.loc[[ls[i]], [5]])  # 紧急联系人编号
            bh = zh(date.loc[[ls[i]], [1]])  # 会员编号
            kh = zh(date.loc[[ls[i]], [2]])  # 会员卡号
            name = zh(date.loc[[ls[i]], [3]])  # 姓名
            jb = zh(date.loc[[ls[i]], [4]])  # 职级
            qt = zh(date.loc[[ls[i]], [7]])  # 全拓.
            qtdw = zh(date.loc[[ls[i]], [8]])  # 全拓档位
            fqt = zh(date.loc[[ls[i]], [6]])  # 非全拓
            '''
            if  js ==0:
                A = kh + ',' + name + '\n' + jb + '\n' + str(qt) + '\n' + str(fqt)
                dot.node(tjr, A, fontname="SimHei", color='Snow4')
                '''
            A = kh + ',' + name + '\n' + jb + '\n' + str(qt) + '\n' + str(fqt)
            B = kh + ',' + name + '\n' + jb + '\n' + str(qt) + '\n' + str(fqt)
            if True:
                print(1)
                dot.node(tjr,A, fontname="SimHei")
                dot.node(bh, B, fontname="SimHei")
                dot.edge(tjr, bh)

    dot.render('test-output/round-table.gv',view=True)


#dd = input('请输入你的文件名(存储在D盘根目录):')
dd = '新市场11月申报业绩_20191121 14点43.xlsx'
path = path +dd
#打开excel
date = pd.read_excel(path,header=None)
#创建网图
dot = Digraph(name='shop',node_attr={ 'style': 'filled','shape':'box'})
#获取行数
hs = len(date)
#获取最大的l
a = 0
for i in range(1, hs):
    cs = zh(date.loc[[i], [9]])  # 层数
    if  cs >a :
        a = cs
print(a)
xh()



'''
        #取第一层
        if cs == 1:
            if qt == 0:
                A = kh+','+name +'\n'+jb +'\n'+qt +'\n'+fqt

        # 查询每层符合条件的人
        if tjr in c:
            #网图每个框显示的内容
            B =  hy+'\n'+a +'\n'+str('%.1f'%ys)
            zyj = zyj + ys
            f.append(hy)
            #绘制整个网图
            if True:
                if js==1:
                    dot.node(tjr, A, fontname="SimHei", color='Snow4')
                dot.node(tjr, fontname="SimHei")
                dot.node(hy, B, fontname="SimHei",color =yanse2)
'''


