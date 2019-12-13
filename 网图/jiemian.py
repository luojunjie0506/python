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
a = False

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

def func():
    filePath = var2.get()
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
    for a in range(0,len(btlist)):
        v.append(tk.IntVar())
        b = tk.Checkbutton(frame4, text=btlist[a], variable=v[-1])
        b.grid(row=0, column=a)

#D:\咨委推荐名单调取数据2_20191127（20191210）(1).xlsx 0038501

def func1():
    global date,hs,dot,card1
    card1 = var1.get()
    filePath = var2.get()
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
    global js
    f = []
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
        if hy == card1:
            A = zw + ' ' + hy + ' ' + name  + ' ' + str(age) + ' ' + sex  + '月均2.8万业绩leg数:' + str(leg) + '\n' + '19年业绩(切割):' + str('%.1f' % (qg / 10000)) + '万\n' + '19年业绩:' + str('%.1f' % (bqg / 10000)) + '万\n'
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
    if f != []:
        js = js + 1
        xh(f)
    else:
        path = 'd://'+ filename +'.gv'
        dot.render(path, view=False)
        label_savePath = tk.Label(frame1, text=path, font=('黑体', 8), bg='red')
        label_savePath.pack(side='right')


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

    # 创建框架frame7
    frame7 = tk.Frame(frame)
    frame7.pack()

    # 创建第二层框架frame2
    frame2 = tk.Frame(frame7)
    frame2.pack(side='left')
    label2 = tk.Label(frame2, text='卡号:', font=('黑体', 10))
    var1 = tk.StringVar()
    card = tk.Entry(frame2,show=None,textvariable=var1,font=('黑体', 8),width = 10)
    label2.pack(side='left')
    card.pack(side='right')

    # 创建第三层框架frame3
    frame3 = tk.Frame(frame7)
    frame3.pack(side='right')
    label3 = tk.Label(frame3, text='文档位置:', font=('黑体', 10))
    var2 = tk.StringVar()
    filePath = tk.Entry(frame3,show=None,textvariable=var2,font=('黑体', 8),width = 30)
    label3.pack(side='left')
    filePath.pack(side='right')

    button1 = tk.Button(frame, text='点击',command = lambda :thread_it(func))
    button1.pack()

    # 创建第四层框架frame4
    frame4 = tk.LabelFrame(frame,text = '表头信息：',height= '50',width='700')
    frame4.pack()
    v = []
    button2 = tk.Button(frame, text='点击',command = lambda :thread_it(func1))
    button2.pack()

    #创建进度条框架frame5
    frame5 = tk.Frame(frame)
    frame5.pack()
    tk.Label(frame5, text='下载进度:', ).pack(side='left')
    canvas = tk.Canvas(frame5, width=465, height=22, bg="white")
    canvas.pack(side='right')

    #创建第一层框架frame1
    frame1 = tk.Frame(frame)
    frame1.pack()
    label1 = tk.Label(frame1,text='文件储存位置:',font=('黑体',10))
    label1.pack(side='left')


    top.mainloop()
