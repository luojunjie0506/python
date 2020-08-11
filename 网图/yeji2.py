import csv
from graphviz import Digraph





def xh(dy):
    global a ,yeji,A,file_name
    zj = ''
    for yw in range(0,hang1):
        if list1[yw][2] == cc:
            num = int(list1[yw][4])
            if num == 7:
                zj = '销售经理'
            elif num == 8:
                zj = '高级销售经理'
            A = list1[yw][2] + ' ' + list1[yw][3] + '\n' + zj + '\n' + '团队业绩:' + str(list1[yw][8]) + '\n' + '个人压缩业绩:' + str(list1[yw][9])
            file_name = list1[yw][3] + ' '+ list1[yw][0] +'期次'
            break

    for i  in range(0,5):
        mc = []  # 每层经销商
        for yw in range(0,hang1):
            if list1[yw][6]  in dy :
                yeji = yeji + float(list1[yw][9])
                mc.append(list1[yw][2])
            else:
                continue
        dy = mc

    A = A + '\n' + '5层团队总业绩:' + str('%.0f' %yeji)


def zt(dy):
    for i in range(0, 6):
        if i < 3:
            yanse2 = 'LightBlue' + str(i + 1)
        else:
            yanse2 = 'LightBlue3'
        mc = []  # 每层经销商
        for yw in range(0, hang1):
            if list1[yw][6] in dy:
                num = int(list1[yw][4])
                if num == 8:
                    zj = '高级销售经理'
                elif num == 7:
                    zj = '销售经理'
                B = str(list1[yw][2]) + ' ' + str(list1[yw][3]) + '\n' + zj + '\n' + '压缩团队业绩:' + str(list1[yw][9])
                mc.append(list1[yw][2])
                if i == 0:
                    dot.node(list1[yw][6], A, fontname="SimHei", color='LightBlue1')
                dot.node(list1[yw][6], fontname="SimHei")
                dot.node(list1[yw][2], B, fontname="SimHei", color=yanse2)
                dot.edge(list1[yw][6], list1[yw][2])
            else:
                continue
        dy = mc
        print(dy)

    save_path = 'test-output/' + file_name
    dot.render(save_path, view=True)


if __name__ == '__main__':
    dy = []  # 第一层经销商
    yeji = 0.0
    a = 0
    cc = '0881198'
    dd = '202007领导奖压缩网.csv'
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










