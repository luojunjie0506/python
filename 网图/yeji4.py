import csv
from graphviz import Digraph


def xh(dy):
    global a ,A,file_name,rs,hz,yeji5
    hz = ''
    rs = 0
    zj = ''
    for yw in range(0,hang1):
        if list1[yw][1] == cc:
            num = int(list1[yw][3])
            if num == 7:
                zj = '销售经理'
            elif num == 8:
                zj = '高级销售经理'
            A = list1[yw][1] + ' ' + list1[yw][2] + '\n' + zj  + '\n' + '压缩团队业绩:' + str(list1[yw][6])
            file_name = list1[yw][2] + ' '+ list1[yw][0] +'期次'
            break

    for i  in range(0,6):
        yeji = 0.0
        mc = []  # 每层经销商
        for yw in range(0,hang1):
            if list1[yw][4]  in dy :
                yeji = yeji + float(list1[yw][6])
                if i <5:
                    yeji5 = yeji5 + float(list1[yw][6])
                mc.append(list1[yw][1])
            else:
                continue
        dy = mc
        if i < 5:
            rs = rs + len(mc)
        hz = hz + '第' + str(i+1) + '层:人数' + str(len(mc)) + ',业绩' + str('%.2f' % yeji) +'\n'

def zt(dy):
    dot1_str = ''
    for i in range(0, 6):
        if i < 3:
            yanse2 = 'LightBlue' + str(i + 1)
        else:
            yanse2 = 'LightBlue3'
        mc = []  # 每层经销商
        for yw in range(0, hang1):
            if list1[yw][4] in dy:
                num = int(list1[yw][3])
                if num == 8:
                    zj = '高级销售经理'
                elif num == 7:
                    zj = '销售经理'
                B = str(list1[yw][1]) + ' ' + str(list1[yw][2]) + '\n' + zj + '\n' + '压缩团队业绩:' + str(list1[yw][6])
                mc.append(list1[yw][1])
                if i == 0:
                    dot.node(list1[yw][4], A, fontname="SimHei", color='LightBlue1')
                dot.node(list1[yw][4], fontname="SimHei")
                dot.node(list1[yw][1], B, fontname="SimHei", color=yanse2)
                dot.edge(list1[yw][4], list1[yw][1])
            else:
                continue
        dy = mc
        print(dy)

    zz = 1001567.85
    dot1_str ='5L职级为经理或高经:共' + str(rs) +'\n' + '5L:'+str('%.2f' % (yeji5))+'/' + str(zz) +str(' {:.0%}'.format(yeji5/zz)) + '\n'+hz
    dot.node(dot1_str , fontname="SimHei", color="LightBlue1", fontsize='20')
    save_path = 'test-output/' + file_name
    dot.render(save_path, view=True)


if __name__ == '__main__':
    dy = []  # 第一层经销商
    yeji5 = 0.0
    a = 0
    cc = '3179768'
    dd = '孙果林6月.csv'
    path = 'D:\\'  + dd

    with open(path, 'r') as csv_file:
        reader = csv.reader(csv_file)
        next(csv_file, None)
        list1 = (list(reader))
    hang1 = len(list1) #一维列表长度
    hang2 = len(list1[1]) #二维列表长度


    dy.append(cc)
    # 创建网图
    dot = Digraph(name='shop', node_attr={'style': 'filled', 'shape': 'box'})
    xh(dy)
    zt(dy)










