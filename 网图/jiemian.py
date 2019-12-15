import tkinter as tk
import pandas as pd
import numpy as np
from graphviz import Digraph
from tkinter import *
import threading

js = 1   #层数
dy = []     #每层经销商
btlist = [] #表头字典
xslie= [] #要显示的列
xslist = [] #要显示的字段
xssj = [] #要显示的字段序号
cf = False

def thread_it(func, *args):
    '''将函数打包进线程'''
    # 创建
    t = threading.Thread(target=func, args=args)
    # 守护 !!!
    t.setDaemon(True)
    # 启动
    t.start()
    # 阻塞--卡死界面！
    # t.join()

#写入表头信息框架frame4数据
def func():
    filePath = var_path.get()
    # 打开excel
    date = pd.read_excel(filePath, header=None)
    f = []
    # 循环表头
    bt = date.ix[0].values
    # 表头长度
    btL = len(bt)
    if len(f) == 0:
        for i in range(0, btL):
            li = bt[i].strip()
            btlist.append(li)
    #写入frame4中
    for a in range(0,len(btlist)):
        v.append(tk.IntVar())
        b = tk.Checkbutton(frame4, text=btlist[a], variable=v[-1])
        b.grid(row=0, column=a)

#D:\咨委推荐名单调取数据2_20191127（20191210）(1).xlsx 0038501
#C:\Users\jjfly\Desktop\咨委推荐名单调取数据2_20191127（20191210）.xlsx

def func1():
    global date,hs,dot,card1,hylie,tjrlie
    card1 = var_card.get()
    filePath = var_path.get()
    hylie = int(var_hylie.get())
    tjrlie = int(var_trjlie.get())
    for a in range(0,len(btlist)):
        if v[a].get() ==1:
            xslie.append(a)
            xslist.append(btlist[a])
    # 打开excel
    date = pd.read_excel(filePath, header=None)
    #获取行数
    hs = len(date)
    dy.append(card1)
    #创建网图
    dot = Digraph(name='shop',node_attr={ 'style': 'filled','shape':'box'})
    xh(dy)

def xh(zz):
    global js,cf
    if cf == False:
        for xx in range(0,len(btlist)):
            if v[xx].get() == True:
                xssj.append(xx)
        cf = True
    f = []
    #循环整个文档行
    for i in range(1, hs):
        mhz = []  # 存储每行显示值列表
        str1 = '' #存储显示内容
        for xx in range(0, len(xssj)):
            b =  xssj[xx]
            hy = zh(date.loc[[i], [hylie-1]])
            tjr = zh(date.loc[[i], [tjrlie-1]])
            mhz.append(zh(date.loc[[i], [b]]))

        #取第一层
        if  hy == card1:
            for xx in range(0, len(xssj)):
                b = xssj[xx]
                str1.join(str(btlist[b]) + ':' +str(mhz[xx]) + '\n')
                print(btlist[b],mhz[xx])
            filename = card1
        # 查询每层符合条件的人
        if tjr in zz:
            #网图每个框显示的内容
            for xx in range(0, len(xssj)):
                b = xssj[xx]
                str1.join(str(btlist[b]) + ':' + str(mhz[xx]) + '\n')
            f.append(hy)
            #绘制整个网图
            if True:
                if js==1:
                        dot.node(tjr, str1, fontname="SimHei", color='LightSkyBlue')
                dot.node(tjr, fontname="SimHei")
                dot.node(hy, str1, fontname="SimHei",color ='Pink')
                dot.edge(tjr, hy)

            '''
            if True:
                if js==1:
                    if sex =="男":
                        dot.node(tjr, str1, fontname="SimHei", color='LightSkyBlue')
                    else:
                        dot.node(tjr, str1, fontname="SimHei", color='Pink')
                if sex =="男":
                    dot.node(tjr, fontname="SimHei")
                    dot.node(hy, str1, fontname="SimHei",color ='LightSkyBlue')
                    dot.edge(tjr, hy)
                else:
                    dot.node(tjr, fontname="SimHei")
                    dot.node(hy, str1, fontname="SimHei",color ='Pink')
                    dot.edge(tjr, hy)
            '''
    #判断
    if f != []:
        js = js + 1
        xh(f)
    else:
        path = 'd://'+ filename +'.gv'
        dot.render(path, view=False)
        Entry_savePath= tk.Entry(frame3, show=None,state ='readonly',textvariable=path, font=('黑体', 8), width=20)
        Entry_savePath.pack(side=LEFT, fill=Y, expand=YES)


#值转为列表
def zh(aa):
    citydaima = np.array(aa)
    citydaima = citydaima.tolist()
    for i in citydaima:
        for j in i:
            return j

if __name__ == '__main__':
    top = tk.Tk()
    top.title('窗口')

    #创建主框架frame
    frame = tk.LabelFrame(top)
    frame.pack()


    # 创建基本信息框架frame3
    frame3 = tk.Frame(frame)
    frame3.pack()
    label_path = tk.Label(frame3, text='文档位置:', font=('黑体', 10)).pack(side=LEFT, fill=Y, expand=YES)
    var_path = tk.StringVar()
    filePath = tk.Entry(frame3,show=None,textvariable=var_path,font=('黑体', 8),width = 30).pack(side=LEFT, fill=Y, expand=YES)
    label_card = tk.Label(frame3, text='卡号:', font=('黑体', 10)).pack(side=LEFT, fill=Y, expand=YES)
    var_card = tk.StringVar()
    hylie = tk.Entry(frame3,show=None,textvariable=var_card,font=('黑体', 8),width = 10).pack(side=LEFT, fill=Y, expand=YES)
    label_hylie = tk.Label(frame3, text='卡号所属列:', font=('黑体', 10)).pack(side=LEFT, fill=Y, expand=YES)
    var_hylie = tk.StringVar()
    card = tk.Entry(frame3,show=None,textvariable=var_hylie,font=('黑体', 8),width = 5).pack(side=LEFT, fill=Y, expand=YES)
    label_trjlie = tk.Label(frame3, text='推荐人所属列:', font=('黑体', 10)).pack(side=LEFT, fill=Y, expand=YES)
    var_trjlie = tk.StringVar()
    tjrlie = tk.Entry(frame3,show=None,textvariable=var_trjlie,font=('黑体', 8),width = 5).pack(side=LEFT, fill=Y, expand=YES)


    button1 = tk.Button(frame, text='获取表头信息',command = lambda :thread_it(func))
    button1.pack()

    # 创建表头信息框架frame4
    frame4 = tk.LabelFrame(frame,text = '表头信息：',height= '50',width='700')
    frame4.pack()
    v = []
    button2 = tk.Button(frame, text='生成图',command = lambda :thread_it(func1))
    button2.pack()

    #创建进度条框架frame5
    frame5 = tk.Frame(frame)
    frame5.pack()
    tk.Label(frame5, text='下载进度:', ).pack(side='left')
    canvas = tk.Canvas(frame5, width=465, height=22, bg="white")
    canvas.pack(side='right')

    #创建文件储存位置框架frame1
    frame1 = tk.Frame(frame)
    frame1.pack()
    label1 = tk.Label(frame1,text='文件储存位置:',font=('黑体',10)).pack(side=LEFT, fill=Y,expand=YES)


    top.mainloop()
