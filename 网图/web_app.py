import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import pandas as pd
import numpy as np
from graphviz import Digraph
import threading
import os
import datetime

class Example(QWidget):
    def __init__(self):
        super().__init__()
        self.btlist = []  # 表头字典
        self.v = []
        self.xssj = [] #要显示的字段序号
        self.b = []  # 存多选框的变量列表
        self.cf = False
        self.js = 1   #层数
        self.initUI()
        self.step = 0
        self.kaiguan = 0

    def thread_it(self,func, *args):
        '''将函数打包进线程'''
        # 创建
        t = threading.Thread(target=func, args=args)
        # 守护 !!!
        t.setDaemon(True)
        # 启动
        t.start()
        # 阻塞--卡死界面！
        # t.join()

    def initUI(self):
        failPath = QLabel('文档位置:')
        self.failPathText = QLabel()
        cardNo = QLabel('制图卡号:')
        self.cardNoEdit = QLineEdit()
        cardLie = QLabel('卡号列 :')
        self.cardLieEdit = QLineEdit()
        self.cardLieEdit.setValidator(QIntValidator())
        tjrLie = QLabel('推荐人列:')
        self.tjrLieEdit = QLineEdit()
        self.tjrLieEdit.setValidator(QIntValidator())
        btxxLabel = QLabel('表头信息:')
        # 表头信息框
        self.btLayout = QGridLayout()
        self.square = QWidget()
        self.square.setLayout(self.btLayout)
        ztLabel = QLabel('制图进度:')
        savePath = QLabel('文件储存位置:')
        self.savePathText = QLabel()
        btButton = QPushButton('获取表头信息')
        ztButton = QPushButton('制图')
        self.jdtBar = QProgressBar()
        failButton = QPushButton('打开')
        Button1 = QPushButton('Button1')
        Button2 = QPushButton('Button2')
        Button3 = QPushButton('Button3')

        #主布局
        zLayout = QHBoxLayout(self)
        left = QFrame(self)
        right = QFrame(self)
        splitter1 = QSplitter(Qt.Horizontal)
        splitter1.addWidget(left)
        splitter1.setSizes([20,]) #设置分隔条位置
        splitter1.addWidget(right)
        zLayout.addWidget(splitter1)
        self.setLayout(zLayout)

        # 左布局
        LeftLayout = QVBoxLayout()
        LeftLayout.addWidget(Button1)
        LeftLayout.addWidget(Button2)
        LeftLayout.addWidget(Button3)
        LeftWidget = QWidget(left)
        LeftWidget.setLayout(LeftLayout)


        #右布局1
        self.RightLayout1 = QGridLayout()
        self.RightLayout1.setSpacing(17)
        self.RightLayout1.addWidget(failPath,1,0,1,1)
        self.RightLayout1.addWidget(self.failPathText, 1,1,1,3)
        self.RightLayout1.addWidget(failButton, 1,4,1,1)
        self.RightLayout1.addWidget(cardNo, 2,0)
        self.RightLayout1.addWidget(self.cardNoEdit, 2,1)
        self.RightLayout1.addWidget(cardLie, 2,2)
        self.RightLayout1.addWidget(self.cardLieEdit, 2,3)
        self.RightLayout1.addWidget(tjrLie, 2,4)
        self.RightLayout1.addWidget(self.tjrLieEdit, 2,5)
        self.RightLayout1.addWidget(btButton, 3,4,1,1)
        self.RightLayout1.addWidget(btxxLabel, 4, 0)
        self.RightLayout1.addWidget(self.square, 5, 0, 3, 6)
        self.RightLayout1.addWidget(ztButton, 10, 4, 1, 1)
        self.RightLayout1.addWidget(ztLabel, 11, 0,)
        self.RightLayout1.addWidget(self.jdtBar, 11, 1,1,5)
        self.RightLayout1.addWidget(savePath, 12, 0)
        self.RightLayout1.addWidget(self.savePathText, 12, 1,1,3)


        RightWidget1 = QWidget()
        RightWidget1.setLayout(self.RightLayout1)


        # 右布局2
        RightWidget2 = QWidget()
        self.RightLayout2 = QGridLayout()
        self.RightLayout2.setSpacing(15)
        RightWidget2.setLayout(self.RightLayout2)

        #界面切换
        self.stackedWidget = QStackedWidget(right)
        self.stackedWidget.addWidget(RightWidget1)
        self.stackedWidget.addWidget(RightWidget2)
        Button1.clicked.connect(self.tiaozhuan1)
        Button2.clicked.connect(self.tiaozhuan2)

        failButton.clicked.connect(self.showDialog)
        btButton.clicked.connect(self.huoqubiaotou)
        ztButton.clicked.connect(lambda :self.thread_it(self.zhitu))

        self.square.setStyleSheet("QWidget { background-color: Blue}")
        self.setFixedSize(800,400)
        self.setWindowTitle('制图')
        self.setWindowIcon(QIcon('D:\\py\\web\\7559\\q1.ico'))
        self.show()

    def tiaozhuan1(self):
        self.stackedWidget.setCurrentIndex(0)
    def tiaozhuan2(self):
        self.stackedWidget.setCurrentIndex(1)

    def showDialog(self):
        fname = QFileDialog.getOpenFileName(self,'Open file','/home')
        self.failPathText.setText(fname[0])

    def huoqubiaotou(self):
        filePath = self.failPathText.text()
        #打开excel
        date = pd.read_excel(filePath, header=None)
        f = []
        #循环表头
        bt = date.ix[0].values
        #表头长度
        btL = len(bt)

        #q清空显示
        if self.kaiguan == 1:
            # 再次点击清空表头信息框显示和进度条设置为o
            self.jdtBar.reset()
            self.savePathText.setText('')
            for a in range(0, len(self.btlist)):
                self.b[a].deleteLater()
            self.b.clear()
            self.btlist.clear()

        if len(f) == 0:
            for i in range(0, btL):
                li = bt[i].strip()
                self.btlist.append(li)

        #新建变量存多选框
        for a in range(0, len(self.btlist)):
            self.b.append('checkBox_'+str(a))

        # 写入frame4中
        for a in range(0, len(self.btlist)):
            if a > 5 :
                self.b[a] = QCheckBox(self.btlist[a], self)
                self.btLayout.addWidget(self.b[a],2,a-6)
            else:
                self.b[a] = QCheckBox(self.btlist[a], self)

                self.btLayout.addWidget(self.b[a],1,a)

        self.kaiguan = 1

    def zhitu(self):
        filePath1 = self.failPathText.text()
        self.card1 = self.cardNoEdit.text()
        self.cardLie1 = self.cardLieEdit.text()
        self.tjrLie1 = self.tjrLieEdit.text()
        # 打开excel
        self.date = pd.read_excel(filePath1, header=None)
        # 获取行数
        self.hs = len(self.date)
        dy = []
        dy.append(self.card1)
        # 创建网图
        self.dot = Digraph(name='shop', node_attr={'style': 'filled', 'shape': 'box'})
        self.xh(dy)

    def xh(self,zz):
        if self.cf == False:
            for a in range(0, len(self.btlist)):
                if self.b[a].isChecked() == True:
                    self.xssj.append(a)
            self.cf = True
        f = []
        # 循环整个文档行
        for i in range(1, self.hs):
            mhz = []  # 存储每行显示值列表
            hy = self.zh(self.date.loc[[i], [int(self.cardLie1) - 1]])
            tjr = self.zh(self.date.loc[[i], [int(self.tjrLie1) - 1]])
            for xx in range(0, len(self.xssj)):
                abc = self.xssj[xx]
                mhz.append(self.zh(self.date.loc[[i], [abc]]))

            # 取第一层
            if hy == self.card1:
                str1 = ''
                for xx in range(0, len(self.xssj)):
                    abc = self.xssj[xx]
                    # 判断是否为空值
                    if pd.isna(mhz[xx]):
                        ww = '-'
                    else:
                        ww = str(mhz[xx])
                    str1 = str1 + str(self.btlist[abc]) + ':' + ww  + '\n'

            # 查询每层符合条件的人
            if tjr in zz:
                # 网图每个框显示的内容
                str2 = ''
                for xx in range(0, len(self.xssj)):
                    abc = self.xssj[xx]
                    #判断是否为空值
                    if pd.isna(mhz[xx]):
                        ww = '-'
                    else:
                        ww = str(mhz[xx])
                    str2 = str2 + str(self.btlist[abc]) + ':' + ww + '\n'
                f.append(hy)
                # 绘制整个网图
                if True:
                    if self.js == 1:
                        self.dot.node(tjr, str1, fontname="SimHei", color='LightSkyBlue')
                    self.dot.node(tjr, fontname="SimHei")
                    self.dot.node(hy, str2, fontname="SimHei", color='LightSkyBlue')
                    self.dot.edge(tjr, hy)

        # 判断
        if f != []:
            self.js = self.js + 1
            if self.step <100:
                self.step = self.step + 10
                self.jdtBar.setValue(self.step)
            self.xh(f)

        else:
            self.jdtBar.setValue(100)
            nowtime = datetime.datetime.now().strftime('%Y%m%d')
            path = os.getcwd() +'\\'+self.card1 + '  '+nowtime
            self.dot.render(path, view=False)
            self.savePathText.setText(str(path))

    def zh(self,aa):
        citydaima = np.array(aa)
        citydaima = citydaima.tolist()
        for i in citydaima:
            for j in i:
                return j

if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = Example()
    sys.exit(app.exec_())


#20190306327