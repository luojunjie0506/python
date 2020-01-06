import pandas as pd
import numpy as np
from graphviz import Digraph


path = 'D:\\'  #文档路径
js = 1   #层数
dy = []
rs= 0
#值转为列表
def zh(aa):
    citydaima = np.array(aa)
    citydaima = citydaima.tolist()
    for i in citydaima:
        for j in i:
            return j

def xh(c):
    global js,rs
    f = []
    #循环整个文档行
    for i in range(1, hs):
        tjr = zh(date.loc[[i], [5]])  # 紧急联系人编号
        bh = zh(date.loc[[i], [1]])  # 会员编号
        kh = zh(date.loc[[i], [2]])  # 会员卡号
        name = zh(date.loc[[i], [3]])  # 姓名
        jb = zh(date.loc[[i], [4]])  # 职级
        qt = zh(date.loc[[i], [7]])  # 全拓.
        qtdw = zh(date.loc[[i], [8]])  # 全拓档位
        fqt = zh(date.loc[[i], [6]])  # 非全拓

        for i in range(1, hs+1):
            if bh == '20190306327':
                if qt == 0 and fqt !=0:
                    A = kh + ',' + name + '\n' + jb + '\n' + '全拓: -'  + '\n' + '非全拓:' + str(fqt)
                elif qt != 0 and fqt ==0:
                    A = kh + ',' + name + '\n' + jb + '\n' + '全拓:' + str(qt) + '  档位:' + qtdw + '\n' + '非全拓: -'
                elif qt == 0 and fqt == 0:
                    A = kh + ',' + name + '\n' + jb + '\n' + '全拓: -'   + '\n' + '非全拓: -'
                else:
                    A = kh + ',' + name + '\n' + jb + '\n' + '全拓:' + str(qt) + '  档位:' + qtdw + '\n' + '非全拓:' + str(fqt)
            break

        # 查询每层符合条件的人
        if tjr in c:
            rs =rs +1
            #网图每个框显示的内容
            if qt == 0 and fqt !=0:
                B = kh + ',' + name + '\n' + jb + '\n' + '全拓: -'  + '\n' + '非全拓:' + str(fqt)
            elif qt != 0 and fqt ==0:
                B = kh + ',' + name + '\n' + jb + '\n' + '全拓:' + str(qt) + '  档位:' + qtdw + '\n' + '非全拓: -'
            elif qt == 0 and fqt == 0:
                B = kh + ',' + name + '\n' + jb + '\n' + '全拓: -'   + '\n' + '非全拓: -'
            else:
                B = kh + ',' + name + '\n' + jb + '\n' + '全拓:' + str(qt) + '  档位:' + qtdw + '\n' + '非全拓:' + str(fqt)
            f.append(bh)
            #绘制整个网图

            if True:
                if js==1:
                    if fqt==0 and qt== 0:
                        dot.node(tjr, A, fontname="SimHei",color='GhostWhite')
                    else:
                        dot.node(tjr, A, fontname="SimHei",color='SkyBlue')

                if fqt==0 and qt ==0:
                    dot.node(tjr, fontname="SimHei")
                    dot.node(bh, B, fontname="SimHei",color='GhostWhite')
                    dot.edge(tjr, bh)
                else :
                    dot.node(tjr, fontname="SimHei")
                    dot.node(bh, B, fontname="SimHei",color='SkyBlue')
                    dot.edge(tjr, bh)

    if f != c:
        js = js + 1
        print(js)
        xh(f)
    else:
        dot.render('D:\\round-table.gv',view=True)


#cc = input('请输入要查询的卡号:')
#dd = input('请输入你的文件名(存储在D盘根目录):')

cc='20190306327'
dd = '新市场12月业绩_20200101.xlsx'
dy.append(cc)
path = path +dd
#打开excel
date = pd.read_excel(path,header=None)
#创建网图
dot = Digraph(name='shop',node_attr={ 'style': 'filled','shape':'box'})
#获取行数
hs = len(date)
xh(dy)
print(rs)