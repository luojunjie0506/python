import xlrd
from graphviz import Digraph


def ddd():
    global q
    q = q +1
    name = 'tabel' + str(q)
    dot = Digraph(name=name, node_attr={'style': 'filled', 'shape': 'box'})
    return dot

def zh1(xx):
    A='' #存储最上点的文字显示
    dot = ddd()
    list1_tjr = []
    list1_tjr.append(xx)
    a = True
    kz = 0

    for a1 in range(1,row1):
        list1_kh = [] #每层清空推荐人列表

        for i in range(1, row1):
            xh = sheet1.cell(i, 0).value
            kh = sheet1.cell(i, 1).value
            tjr = sheet1.cell(i, 5).value
            name = sheet1.cell(i, 2).value
            zj = sheet1.cell(i, 3).value
            lxfx = sheet1.cell(i, 4).value

            # 在sheet1中找到图最上点的文字显示
            if kh == xx and kz ==0 :
                A = name
                kz =1

            #判断推荐人是否在推荐人列表里
            if tjr in list1_tjr:
                if a:
                    dot.node(tjr, A, fontname="SimHei", color='GhostWhite')
                    a = False
                list1_kh.append(kh)
                B = kh + '\n' + name + '\n' + zj + '\n' + str(lxfx) + '\n'
                if xh >26:
                    dot.node(kh, B, fontname="SimHei", color='LightSkyBlue')
                else:
                    dot.node(kh, B, fontname="SimHei", color='GhostWhite')
                dot.edge(tjr, kh)

        if len(list1_kh)>0:
            list1_tjr = list1_kh
        else:
            break

    save_path = 'D:\\图\\'+xx
    dot.render(save_path)

def zh2(xx):
    dot = ddd()
    list2_tjr = []
    list2_tjr.append(xx)
    a = True
    ks = 0
    kz = 0
    A = ''

    #在sheet2中的开始行数并构造图最上点的文字显示
    for b in range(1, row2):
        kh = sheet2.cell(b, 1).value
        if kh == xx:
            ks = b
            break

    for a1 in range(ks, row2):
        list2_kh = []  # 每层清空推荐人列表

        for i in range(ks, row2):
            kh = sheet2.cell(i, 1).value
            tjr = sheet2.cell(i, 5).value
            name = sheet2.cell(i, 2).value
            zj = sheet2.cell(i, 3).value
            lxfx = sheet2.cell(i, 4).value

            if kh == xx and kz ==0 :
                A = kh + '\n' + name + '\n' + zj + '\n' + str(lxfx) + '\n'
                kz =1

            # 判断推荐人是否在推荐人列表里
            if tjr in list2_tjr:
                if a:
                    dot.node(tjr, A, fontname="SimHei", color='GhostWhite')
                    a = False
                list2_kh.append(kh)
                B = kh + '\n' + name + '\n' + zj + '\n' + str(lxfx) + '\n'
                dot.node(kh, B, fontname="SimHei", color='GhostWhite')
                dot.edge(tjr, kh)

        if len(list2_kh) > 0:
            list2_tjr = list2_kh
        else:
            break

    save_path = 'D:\\图\\' + xx
    dot.render(save_path)




if __name__ == '__main__':
    q = 0
    m = 0 #用于存储sheet1中的卡号在sheet2中匹配的数量
    ks = 0  # 制图卡号在表中开始的行数
    kh_lish1 = [] #存储sheet1中卡号
    kh_lish=[] #存储在sheet1中并在sheet2中有匹配数量的卡号
    #dd = input('请输入你的文件名(存储在D盘根目录):')
    dd='前海回归（重点扶持数据）.xlsx'
    path = 'D:\\'
    path = path +dd
    data = xlrd.open_workbook(path)
    #打开sheet
    sheet1 = data.sheet_by_index(0)
    sheet2 = data.sheet_by_index(1)
    #获取行数
    row1 = sheet1.nrows
    row2 = sheet2.nrows


    #把sheet1中的卡号存入kh_list1中
    for i in range(1,27):
        kh = sheet1.cell(i, 1).value
        if i == 1:
            kh_lish.append(kh)
        else:
            kh_lish1.append(kh)

    #将sheet1的卡号去匹配二表的数据，如果数据存在2条以上，就存在制表卡号中kh_list
    for card in kh_lish1:
        for i in range(1,row2):
            kh = sheet2.cell(i, 1).value
            tjr = sheet2.cell(i, 5).value
            if kh == card or tjr == card:
                m =m +1
            if m > 1 :
                kh_lish.append(card)
                break
        m = 0


    #判断用哪个表制图
    for a in range(0,len(kh_lish)):
        if a==0:
            zh1(kh_lish[a])

        else:
            zh2(kh_lish[a])




#前海回归（重点扶持数据）.xlsx