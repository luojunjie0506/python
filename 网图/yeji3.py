import xlrd
from graphviz import Digraph


def xx(a):
    dy = [] #存放每层推荐人
    rs = 0 #存放五层人数
    zj = '' #职级
    zt_card = '' #存放制图卡号
    yeji5 = 0.0 #存放五层业绩
    hz = '' #存放大框中每行数据

    # 创建网图
    dot = Digraph(name='shop', node_attr={'style': 'filled', 'shape': 'box'})
    #打开指定sheet
    sheet = data.sheet_by_index(a)
    #获取行数
    rows = sheet.nrows
    qici = int(sheet.cell(1, 0).value)
    print(qici)
    zt_card = sheet.cell(1, 2).value
    name = sheet.cell(1, 3).value
    num = int(sheet.cell(1, 4).value)
    td_yeji = sheet.cell(1, 8).value
    pgv = sheet.cell(1, 9).value
    if num == 8:
        zj = '高经'
    elif num == 7:
        zj = '经理'
    A = zt_card + '\n ' + name + '\n' + zj + '\n' + 'PGV:' + str(pgv)
    file_name = name + ' ' + str(qici) + '期次'  #文件存放名字
    dy.append(zt_card)


    #做六层大框内容
    for i  in range(0,6):
        yeji = 0.00 #存放每层业绩
        mc = []  # 每层经销商(中间值)
        for row in range(0, rows):
            tjr = sheet.cell(row, 6).value
            if tjr in dy:
                card = sheet.cell(row, 2).value
                yeji = yeji + float(sheet.cell(row, 9).value)
                if i <5:
                    yeji5 = yeji5 + float(sheet.cell(row, 9).value)
                mc.append(card)
            else:
                continue
        dy = mc
        if i < 5:
            rs = rs + len(mc)
        hz = hz + '第' + str(i + 1) + '层:人数' + str(len(mc)) + ',业绩' + str('%.1f' % yeji) + '\n'

    dy = []
    dy.append(zt_card )
    #做六层神经图
    for i  in range(0,6):
        if i < 3:
            yanse2 = 'LightBlue' + str(i + 1)
        else:
            yanse2 = 'LightBlue3'
        mc = []  # 每层经销商(中间值)
        for row in range(0, rows):
            tjr = sheet.cell(row, 6).value
            if tjr in dy:
                card = sheet.cell(row, 2).value
                name = sheet.cell(row, 3).value
                num = int(sheet.cell(row, 4).value)
                pgv = sheet.cell(row, 9).value
                if num == 8:
                    zj = '高经'
                elif num == 7:
                    zj = '经理'
                B = str(card) + '\n' + str(name) + '\n' + zj + '\n' + 'PGV:' + str(pgv)
                mc.append(card)
                if i == 0:
                    dot.node(tjr, A, fontname="SimHei", color='LightBlue1')
                dot.node(tjr, fontname="SimHei")
                dot.node(card, B, fontname="SimHei", color=yanse2)
                dot.edge(tjr, card)
            else:
                continue
        dy = mc

    dot1_str ='5L职级为经理或高经:共' + str(rs) +'\n' + '5L:'+str('%.1f' % (yeji5))+'/' + str('%.1f' % (td_yeji)) +str(' {:.0%}'.format(yeji5/td_yeji)) + '\n'+hz
    dot.node(dot1_str , fontname="SimHei", color="LightBlue1", fontsize='20')
    save_path = 'D:\\图\\咨委\\' + file_name
    dot.render(save_path)



if __name__ == '__main__':
    dd ='咨委组织架构数据.xlsx'
    path = 'D:\\' + dd  #文件路径
    data = xlrd.open_workbook(path) #打开excel
    #获取sheet数量
    count = len(data.sheets())
    # 打开sheet1
    sheet1 = data.sheet_by_index(0)
    #获取列数
    cow = sheet1.ncols

    #循环sheet
    for a in range(0,count):
        xx(a)

