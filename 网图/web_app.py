import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
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
        self.cf = False
        self.js = 1   #层数
        self.b = []   #存多选框的变量列表
        self.initUI()
        self.step = 0

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
        ztButton.clicked.connect(lambda :self.thread_it(self.zhitu))

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
        #再次点击清空表头信息框显示和进度条设置为o
        self.jdtBar.setValue(0)
        self.savePathText.setText('')

        for i in range(self.btLayout.count()):
            self.btLayout.removeWidget(self.btLayout.itemAt(i).widget())

        filePath = self.failPathText.text()
        #打开excel
        date = pd.read_excel(filePath, header=None)
        f = []
        #循环表头
        bt = date.ix[0].values
        #表头长度
        btL = len(bt)
        if len(f) == 0:
            for i in range(0, btL):
                li = bt[i].strip()
                self.btlist.append(li)

        #新建变量存多选框
        for a in range(0, len(self.btlist)):
            self.b.append('checkBox_'+str(a))

        # 写入frame4中
        for a in range(0, len(self.btlist)):
            self.b[a] = QCheckBox(self.btlist[a], self)
            self.btLayout.addWidget(self.b[a])

    def zhitu(self):
        filePath1 = self.failPathText.text()
        self.card1 = self.cardNoEdit.text()
        self.cardLie1 = int(self.cardLieEdit.text())
        self.tjrLie1 = int(self.tjrLieEdit.text())
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
            hy = self.zh(self.date.loc[[i], [self.cardLie1 - 1]])
            tjr = self.zh(self.date.loc[[i], [self.tjrLie1 - 1]])
            for xx in range(0, len(self.xssj)):
                abc = self.xssj[xx]
                mhz.append(self.zh(self.date.loc[[i], [abc]]))

            # 取第一层
            if hy == self.card1:
                str1 = ''
                for xx in range(0, len(self.xssj)):
                    abc = self.xssj[xx]
                    str1 = str1 + str(self.btlist[abc]) + ':' + str(mhz[xx]) + '\n'
                filename = self.card1

            # 查询每层符合条件的人
            if tjr in zz:
                # 网图每个框显示的内容
                str2 = ''
                for xx in range(0, len(self.xssj)):
                    abc = self.xssj[xx]
                    str2 = str2 + str(self.btlist[abc]) + ':' + str(mhz[xx]) + '\n'
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
            path = os.getcwd() +'\\'+filename + '  '+nowtime
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