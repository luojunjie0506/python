import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
import pandas as pd
import numpy as np
from graphviz import Digraph
import threading
import os


class Example(QWidget):
    def __init__(self):
        super().__init__()
        self.btlist = []  # 表头字典
        self.v = []
        self.xssj = [] #要显示的字段序号
        self.cf = False
        self.js = 1   #层数
        self.initUI()

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
        #表头信息框
        self.btLayout = QHBoxLayout()
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
        zLayout = QHBoxLayout()
        #左布局
        LeftLayout = QVBoxLayout()
        LeftLayout.addWidget(Button1)
        LeftLayout.addWidget(Button2)
        LeftLayout.addWidget(Button3)
        LeftWidget = QWidget()
        LeftWidget.setLayout(LeftLayout)

        #右布局
        RightLayout = QGridLayout()
        RightLayout.setSpacing(15)
        RightLayout.addWidget(failPath,1,0,1,1)
        RightLayout.addWidget(self.failPathText, 1,1,1,3)
        RightLayout.addWidget(failButton, 1,4,1,1)
        RightLayout.addWidget(cardNo, 2,0)
        RightLayout.addWidget(self.cardNoEdit, 2,1)
        RightLayout.addWidget(cardLie, 2,2)
        RightLayout.addWidget(self.cardLieEdit, 2,3)
        RightLayout.addWidget(tjrLie, 2,4)
        RightLayout.addWidget(self.tjrLieEdit, 2,5)
        RightLayout.addWidget(btButton, 3,4,1,1)
        RightLayout.addWidget(btxxLabel, 4, 0)
        RightLayout.addWidget(self.square, 5,0,2,6)
        RightLayout.addWidget(ztButton, 9, 4, 1, 1)
        RightLayout.addWidget(ztLabel, 10, 0,)
        RightLayout.addWidget(self.jdtBar, 10, 1,1,5)
        RightLayout.addWidget(savePath, 11, 0)
        RightLayout.addWidget(self.savePathText, 11, 1,1,3)

        RightWidget = QWidget()
        RightWidget.setLayout(RightLayout)


        failButton.clicked.connect(self.showDialog)
        btButton.clicked.connect(self.huoqubiaotou)
        #ztButton.clicked.connect(self.zhitu)

        zLayout.addWidget(LeftWidget)
        zLayout.addWidget(RightWidget)
        self.setLayout(zLayout)

        self.square.setStyleSheet("QWidget { background-color: Blue}")
        self.resize(700,400)
        self.setWindowTitle('制图')
        self.setWindowIcon(QIcon('D:\\py\\web\\7559\\q1.ico'))
        self.show()

    def showDialog(self):
        fname = QFileDialog.getOpenFileName(self,'Open file','/home')
        self.failPathText.setText(fname[0])

    def huoqubiaotou(self):
        filePath = self.failPathText.text()
        #打开excel
        date = pd.read_excel(filePath, header=None)
        f = []
        b= []
        #循环表头
        bt = date.ix[0].values
        #表头长度
        btL = len(bt)
        if len(f) == 0:
            for i in range(0, btL):
                li = bt[i].strip()
                self.btlist.append(li)

        #新建变量
        for a in range(0, len(self.btlist)):
            b.append('checkBox_'+str(a))

        # 写入frame4中
        for a in range(0, len(self.btlist)):
            b[a] = QCheckBox(self.btlist[a], self)
            self.btLayout.addWidget(b[a])



    '''
    def zhitu(self):
        filePath1 = self.failPathText.text()
        card1 = self.cardNoEdit.text()
        self.cardLie1 = int(self.cardLieEdit.text())
        self.tjrLie1 = int(self.tjrLieEdit.text())
        # 打开excel
        date = pd.read_excel(filePath1, header=None)
        # 获取行数
        self.hs = len(date)
        dy = []
        dy.append(card1)
        # 创建网图
        self.dot = Digraph(name='shop', node_attr={'style': 'filled', 'shape': 'box'})
        self.xh(dy)


    def xh(self,zz):
        if self.cf == False:
            for a in range(0, len(self.btlist)):
                m = getattr(self,"checkBox_%d" % a)
                if m.isChecked() == True:
                    self.xssj.append(a)
            print(self.xssj)
            self.cf = True
            self.savePathText.setText(str(self.cardLie1 - 1))


        
        f = []
        # 循环整个文档行
        for i in range(1, self.hs):
            mhz = []  # 存储每行显示值列表
            for xx in range(0, len(self.xssj)):
                b = self.xssj[xx]
                hy = self.zh(self.date.loc[[i], [self.cardLie1- 1]])
                tjr = self.zh(self.date.loc[[i], [self.tjrLie1 - 1]])
                mhz.append(self.zh(self.date.loc[[i], [b]]))

            # 取第一层
            if hy == self.card1:
                str1 = ''
                for xx in range(0, len(self.xssj)):
                    b = self.xssj[xx]
                    str1 = str1 + str(self.btlist[b]) + ':' + str(mhz[xx]) + '\n'
                filename = self.card1
            # 查询每层符合条件的人
            if tjr in zz:
                # 网图每个框显示的内容
                str2 = ''
                for xx in range(0, len(self.xssj)):
                    b = self.xssj[xx]
                    str2 = str2 + str(self.btlist[b]) + ':' + str(mhz[xx]) + '\n'
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
            js = self.js + 1
            self.xh(f)

        else:
            path = os.getcwd() + filename + '.gv'
            self.dot.render(path, view=False)
            self.failPathText.setText(path)

    def zh(self,aa):
        citydaima = np.array(aa)
        citydaima = citydaima.tolist()
        for i in citydaima:
            for j in i:
                return j
    '''

if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = Example()
    sys.exit(app.exec_())


#0038501