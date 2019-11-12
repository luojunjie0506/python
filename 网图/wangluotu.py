import pandas as pd
import numpy as np
from graphviz import Digraph


path = 'D:\\'  #文档路径
js = 1   #层数
zyj =0   #总业绩
zfc=''  #每层业绩字符串
yj5=0   #5层业绩
yj10 = 0    #10层业绩
zrs = 0     #总人数
ds = 0      #市场数
rs5 = 0     #5层人数
dy = []     #每层经销商
yanse = 4


#201909领导奖压缩网(剔除未达标)L.xlsx
# 0881198
#值转为列表
def zh(aa):
    citydaima = np.array(aa)
    citydaima = citydaima.tolist()
    for i in citydaima:
        for j in i:
            return j

def xh(c):
    global zyj,js,zfc,yj10,zrs,yj5,dy,ds,rs5,yanse
    f = []
    rs = 0
    mcyj = 0
    #循环整个文档行
    for i in range(1, hs):
        ys = zh(date.loc[[i], [7]])  #压缩业绩
        tjr = zh(date.loc[[i], [8]])    #紧急联系人编号
        hy = zh(date.loc[[i], [2]])     #会员编号
        jb = zh(date.loc[[i],[5]])      #级别
        sc = zh(date.loc[[i], [4]])     #市场数
        yanse2 = 'Snow' + str(yanse-1)

        #取第一层
        if hy == cc:
            # 级别的判断
            if jb==8:
                a='高经'
            else:
                a='经理'
            A = hy+'\n'+a +'\n'+str('%.1f'%ys)

        # 查询每层符合条件的人
        if tjr in c:
            #取出要查询人的市场数
            if rs ==0:
                rr = zh(dy)
                if hy == rr :
                    ds =sc
            zrs =zrs +1

            if jb==8:
                a='高经'
            else:
                a='经理'
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
                dot.edge(tjr, hy)
                if js <=10:
                    yj10 = yj10 + ys
                    if js<=5:
                        mcyj= mcyj +ys
                        yj5 = yj5 +ys
                        rs = rs + 1
                        rs5 =rs5 +1


    if f != c:
        if js <= 5:
            zfc = zfc + '第' + str(js) + '层:人数' + str(rs) + ',业绩' + str('%.2f' % (mcyj / 10000)) + '万\n'
            print(zfc)
        js = js + 1
        if yanse >2:
            yanse = yanse - 1
        print(zyj)
        xh(f)
    else:
        a1='5L职级为经理或高经:共'+str(rs5) +'\n'
        a2 = '合格市场数:'+str(sc)+'个' + '\n'
        a3='5L:'+str('%.2f' % (yj5/10000))+'万'+'/全网:'+str('%.2f' % (zyj/10000))+'万'+str('{:.0%}'.format(yj5/zyj))+'\n'
        a4 = '10L:'+str('%.2f' % (yj10/10000))+'万'+'/全网:'+str('%.2f' % (zyj/10000))+'万'+str('{:.0%}'.format(yj10/zyj))+'\n'
        zfc =a1 +a2 +a3 +a4 +zfc
        dot.node(zfc, fontname="SimHei", color="White",fontsize='15')
        dot.render('test-output/round-table.gv',view=True)



cc = input('请输入要查询的卡号:')
dd = input('请输入你的文件名(存储在D盘根目录):')
path = path +dd
dy.append(cc)
#打开excel
date = pd.read_excel(path,header=None)
#创建网图
dot = Digraph(name='shop',node_attr={ 'style': 'filled','shape':'box'})
#获取行数
hs = len(date)
xh(dy)