# coding=gbk
import datetime
import os
import pandas as pd
import xlrd
from openpyxl import *
import numpy as np

def getmonth():
    # 获取当前月份-1
    year = datetime.datetime.now().year
    month = datetime.datetime.now().month
    if month ==1 :
        ym = str(year-1) + str(12)
    elif month < 10:
        ym = str(year) + '0' + str(month - 1)
    else:
        ym = str(year) + str(month - 1)
    return ym

# 整合文件夹中的表到一个excel中
def zhenghe():
    global sheets_list
    path = 'D:\\fuwufei\\c\\' # 文件夹路径
    save_path = 'D:\\fuwufei\\xx.xlsx'  # 合并后excel的存放路径
    file_list = os.listdir(path)  # 获取文件夹中的所有文件名的列表
    writer = pd.ExcelWriter(save_path)  # 使用ExcelWriter函数，可以写入数据时不会覆盖sheet
    # 遍历文件夹中的所有文件名的列表
    for a in file_list:
        file_path = path + a  # 拼接完整路径
        data = pd.read_excel(file_path, sheet_name=0)  # 读取文件夹中的表
        df = pd.DataFrame(data)  # 转换格式

        # 循环每个表中表头的第一个字段，并用来排序
        for btname in data:
            #判断表是否是服务费清单原始数据表，是的话就要表中第一个字段和顺序字段排列
            # if '原始数据' in a:
            #     df.sort_values(by=[btname, '顺序'], ascending=True, inplace=True)  # 使用第一个字段来升序。ascending排序，inplace替代
            # else:
            #     df.sort_values(by=btname, ascending=True, inplace=True)  # 使用第一个字段来升序。ascending排序，inplace替代
            df.sort_values(by=btname, ascending=True, inplace=True)
            df.to_excel(writer, encoding='utf-8', sheet_name=a[:-5], index=None)  # 写入sheet中 index无索引
            writer.save()
            writer.close()
            sheets_list.append(a[:-5])
            break
    print(sheets_list)
    return save_path

def xhs(dh):
    global b0, b1, b2, b3, b4
    global sheet0, sheet1, sheet2, sheet3, sheet4
    global row_list, col_list
    # '202012 服务费清单原始数据','202012 门店信息表', '202012 小微店补', '202012 补扣款备注',  '202012 活动明细表'
    list1 = []  # 存储匹配服务费清单原始数据
    list2 = []  # 存储匹配补扣款备注,小微店补,门店信息表
    list3 = []  # 存储匹配活动明细表

    # 找到每个表的行数，如有存在行数就追加到list1,如不存在，在list1中加0后到后一个表匹配的行数
    # b参数是记录每次查询到的位置
    for a0 in range(b0, row_list[0]):
        mh0 = []
        # b的位置跟长度相等时，说明表中数据一匹配完，直接在list中加0
        if a0 == row_list[0]:
            break
        else:
            cell_0 = sheet0.cell(a0, 0).value
            if cell_0 != dh:
                break
            else:
                for index in range(0,col_list[0]):
                    value_0 = sheet0.cell(a0, index).value
                    mh0.append(value_0)
                list1.append(mh0)
                b0 = a0 + 1

    # for循环中如果range中的两个参数都一样时,那这个循环就不执行,为了让循环继续执行所以+1
    for a1 in range(b1, row_list[1]+1):
        if a1 == row_list[1]:
            list2.append('')
            break
        else:
            cell_1 = sheet1.cell(a1, 0).value
            if cell_1 == dh:
                for index in range(0,col_list[1]):
                    value_1 = sheet1.cell(a1, index).value
                    list2.append(value_1)
                b1 = a1 + 1
                break

    for a2 in range(b2, row_list[2]+1):
        if a2 == row_list[2]:
            list2.append('')
            break
        else:
            cell_2 = sheet2.cell(a2, 0).value
            if cell_2 == dh:
                value_2 = sheet2.cell(a2, 2).value
                list2.append(value_2)
                b2 = a2 + 1
                break

    for a3 in range(b3, row_list[3]+1):
        if a3 == row_list[3]:
            list2.append('')
            break
        else:
            cell_3 = sheet3.cell(a3, 0).value
            if cell_3 == dh:
                value_3 = sheet3.cell(a3, 1).value
                list2.append(value_3)
                b3 = a3 + 1
                break

    for a4 in range(b4, row_list[4]):
        mh1 = []
        if a4 == row_list[4]:
            break
        else:
            cell_4 = sheet4.cell(a4, 0).value
            if cell_4 != dh:
                break
            else:
                for index in range(0,col_list[4]):
                    value_4 = sheet4.cell(a4, index).value
                    mh1.append(value_4)
                list3.append(mh1)
                b4 = a4 + 1

    #数据处理
    list1  = dataCl1(list1,list2)
    list3 = dataCl3(list3)
    # print(list3)
    # print(len(list3))
    # 传入xr方法
    xr(list1,list2,list3)

# list1中数据的处理
def dataCl1(list1,list2):
    def take2(elem):
        return elem[1]
    list1.sort(key=take2)
    sum = 0
    for yiwei in range(0, len(list1)):
        # 判断第一层
        if yiwei == 0:
            # 判断list1的第一行是否店长
            if list1[yiwei][1] == '店长':
                if list2[5] != '':
                    # list2中小薇店补不为空,加一行小薇,一行合计
                    sss = []
                    sss.append(list2[0])
                    sss.append(list2[1])
                    sss.append(list2[4])
                    sss.append('小微店补')
                    sss.append(list2[5])
                    list1.insert(1, sss)
                    sss1 = ['电子积分个人账户入账合计']
                    nuM = float(list1[0][16]) + float(list2[5])
                    sss1.append(nuM)
                    list1.insert(2, sss1)
                else:
                    # list2中小薇店补为空,加一行合计
                    sss1 = ['电子积分个人账户入账合计']
                    sss1.append(list1[0][16])
                    list1.insert(1, sss1)
            else:
                if list2[5] != '':
                    sss = []
                    sss.append(list2[0])
                    sss.append(list2[1])
                    sss.append(list2[4])
                    sss.append(list2[5])
                    list1.insert(0, sss)
                    sss1 = ['电子积分个人账户入账合计']
                    sss1.append(list2[5])
                    list1.insert(1, sss1)
        else:
            if list1[yiwei][1] == '门店代发＜200元':
                if yiwei+1 == len(list1):
                    break
                elif list1[yiwei][1] == list1[yiwei + 1][1]:
                    sum += float(list1[yiwei][16])
                    continue
                elif list1[yiwei][1] != list1[yiwei + 1][1]:
                    sum += float(list1[yiwei][16])
                    sss1 = ['电子积分代发账户入账合计']
                    sss1.append(sum)
                    list1.insert(yiwei + 1, sss1)
                    break
    for i in range(0,len(list1)):
        list1[i].insert(0, i+1)
    return list1

# list3数据处理
def dataCl3(list3):
    if len(list3) > 0:
        list4 = np.array(list3)
        idex = np.lexsort([list4[:, 1], list4[:, 5]])
        list3 = list4[idex, :]
        list3 = list3.tolist()
        sum = 0.0 #存每个人的合计
        ss = [] #存第几行插入合计
        for i in range(0, len(list3)):
            if i == len(list4) - 1:
                sum += float(list3[i][6])
                ss.append([i + 1, sum])
                break
            elif list3[i][5] != list3[i + 1][5]:
                sum += float(list3[i][6])
                ss.append([i + 1, sum])
                sum = 0.0
            else:
                sum += float(list3[i][6])
        for y in range(0, len(ss)):
            cc = ['合计：']
            cc.append(ss[y][1])
            # 因为循环了几次就插入了几行,行数变化所以加上y
            list3.insert(ss[y][0] + y, cc)
        return list3
    else:
        return list3


# 新建模板并写入数据
def xr(list1, list2,list3):
    global yue,wcnum
    wb = load_workbook('D:\\fuwufei2\\清单模板.xlsx')
    ws = wb["清单"]
    ws1 = wb['活动奖项明细']
    a = '注：服务费发放说明：\n1、活动奖：全拓直接服务奖 + 分享有礼直接服务奖 + 翻倍服务奖（明细见附件2\n2、收入合计 = 个人销售奖 + 销售服务奖 + 活动奖 + 领导奖 + 卓越奖\n3、秒结入账：符合秒结规则的分享有礼直接服务奖和对应个人销售奖，已实时入账相应电子积分秒结账户，不再做月结入账\n4、月结入账 = 收入合计 - 秒结入账 - 保险代扣 - 其他扣除\n5、绿色代表收入，粉色代表支出。\n'
    n = 6  #清单sheet从哪行开始写

    #清单sheet写入数据
    for yiwei in range(0, len(list1)):
        if len(list1[yiwei]) >10:
            ss = 1#清单sheet从哪列开始写
            for i in range(0,len(list1[yiwei])):
                if i==1:
                    continue
                else:
                    ws.cell(row=n, column=ss).value = list1[yiwei][i]
                    ss += 1
            n +=1
        elif len(list1[yiwei]) == 6:
            ws.cell(row=n, column=1).value = list1[yiwei][0]
            ws.cell(row=n, column=3).value = list1[yiwei][1]
            ws.cell(row=n, column=4).value = list1[yiwei][2]
            ws.cell(row=n, column=6).value = list1[yiwei][3]
            ws.cell(row=n, column=8).value = list1[yiwei][4]
            ws.cell(row=n, column=17).value = list1[yiwei][5]
            n += 1
        elif len(list1[yiwei]) == 3:
            ws.cell(row=n, column=1).value = list1[yiwei][0]
            ws.cell(row=n, column=3).value = list1[yiwei][1]
            ws.cell(row=n, column=17).value = list1[yiwei][2]
            n += 1

    ws['C3'] = list2[0]
    ws['H3'] = list2[1]
    ws['M3'] = yue
    ws['Q3'] = list2[2]
    ws['T6'] = list2[6]
    ws.cell(row=n, column=1).value = a

    #活动奖项明细sheet写入数据
    ss1 = 3  # 活动奖项明细sheet从哪列开始写
    if len(list3) == 1:
        if len(list3[0]) != 0:
            for i in range(0, len(list3[0])-1):
                ws1.cell(row=3, column=2).value = i + 1
                ws1.cell(row=3, column=i+3).value = list3[0][i+1]
    else:
        for i in range(0, len(list3)):
            if len(list3[i]) >2:
                for c in range(0, len(list3[i])-1):
                    ws1.cell(row=i+3, column=2).value = i + 1
                    ws1.cell(row=i+3, column=c+3).value = list3[i][c+1]
            else:
                ws1.cell(row=i + 3, column=2).value = list3[i][0]
                ws1.cell(row=i + 3, column=8).value = list3[i][1]


    file_name = dir_path + '\\'+ list2[0] + '.xlsx'
    wb.save(file_name)
    wb.close()
    wcnum += 1
    print(wcnum)

if __name__ == '__main__':
    b0 = 1
    b1 = 1
    b2 = 1
    b3 = 1
    b4 = 1
    wcnum = 0 #完成的数量
    row_list = []  # 存放每个表的行数
    col_list = []  # 存放每个表的列数
    yue = getmonth()
    # sheets_list = []  # 存放合并后的sheet名列表
    sheets_list = ['202012 小微店补', '202012 服务费清单原始数据', '202012 活动明细表', '202012 补扣款备注', '202012 门店信息表']
    # zhenghe_path = zhenghe()

    #创建文件夹存放各店信息
    dir_path = 'D:\\fuwufei2\\' + getmonth() + '服务费清单'
    # os.mkdir(dir_path)

    data = xlrd.open_workbook("D:\\fuwufei2\\xx.xlsx")
    # 按顺序打开各个sheet表并获取行列
    for num in range(0, len(sheets_list)):
        for sheet_list in sheets_list:
            if ('原始数据' in sheet_list) and num == 0:
                sheet0 = data.sheet_by_name(sheet_list)
                rows0 = sheet0.nrows
                cols0 = sheet0.ncols
                row_list.append(rows0)
                col_list.append(cols0)
                break
            elif ('门店信息' in sheet_list) and num == 1:
                sheet1 = data.sheet_by_name(sheet_list)
                rows1 = sheet1.nrows
                cols1 = sheet1.ncols
                row_list.append(rows1)
                col_list.append(cols1)
                break
            elif ('小微店补' in sheet_list) and num == 2:
                sheet2 = data.sheet_by_name(sheet_list)
                rows2 = sheet2.nrows
                cols2 = sheet2.ncols
                row_list.append(rows2)
                col_list.append(cols2)
                break
            elif ('补扣款备注' in sheet_list) and num == 3:
                sheet3 = data.sheet_by_name(sheet_list)
                rows3 = sheet3.nrows
                cols3 = sheet3.ncols
                row_list.append(rows3)
                col_list.append(cols3)
                break
            elif ('活动明细表' in sheet_list) and num == 4:
                sheet4 = data.sheet_by_name(sheet_list)
                rows4 = sheet4.nrows
                cols4 = sheet4.ncols
                row_list.append(rows4)
                col_list.append(cols4)
                break

    print(row_list)
    for xhdh in range(1, row_list[0] ):
        dh1 = sheet0.cell(xhdh-1, 0).value
        dh2 = sheet0.cell(xhdh, 0).value # 需要制作清单的店号
        if dh1 != dh2 :
            xhs(dh2)

