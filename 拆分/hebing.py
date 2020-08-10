import pandas as pd
import os
import xlrd, xlwt
from xlwt import *
import datetime
from xlutils.copy import copy

'''
前提：先把每个表中的数据按字母降序排序，这样每次匹配行数的时候会少执行很多次
1.打开服务费清单原始表，获取店号。
1.1.用店号去匹配每个表中对应数据的行数（用列表存进来）
1.2.增加写模板，用存起来的行数去查询对应表中的数据并写入复制的模板中并店号或名字保存
2.再次执行前面步骤
'''


# 循环表
def xhs(dh):
    global b0, b1, b2, b3, b4, b5, ym,dir_path
    global sheet0, sheet1, sheet2, sheet3, sheet4, sheet5
    global list3, list4
    list1 = []  # 存储匹配前五个表中对应数据的行数
    list2 = []  # 存储匹配最后一个表中对应数据的行数

    # 找到每个表的行数，如有存在行数就追加到list1,如不存在，在list1中加0后到后一个表匹配的行数
    # b参数是记录每次查询到的位置
    for a0 in range(b0, list3[0] + 1):
        len0 = list3[0]
        # b的位置跟长度相等时，说明表中数据一匹配完，直接在list1中加1
        if b0 == len0:
            list1.append(0)
            break
        else:
            cell_0 = sheet0.cell(a0, 0).value
            # 前五个表有且最多只能存在一条数据，判断一次是否等于匹配值就可以
            if cell_0 == dh:
                list1.append(a0)
                b0 = a0 + 1
                break

    for a1 in range(b1, list3[1] + 1):
        len1 = list3[1]
        if a1 == len1:
            list1.append(0)
            break
        else:
            cell_1 = sheet1.cell(a1, 0).value
            if cell_1 == dh:
                list1.append(a1)
                b1 = a1 + 1
                break

    for a2 in range(b2, list3[2] + 1):
        len2 = list3[2]
        if a2 == len2:
            list1.append(0)
            break
        else:
            cell_2 = sheet2.cell(a2, 0).value
            if cell_2 == dh:
                list1.append(a2)
                b2 = a2 + 1
                break

    for a3 in range(b3, list3[3] + 1):
        len3 = list3[3]
        if a3 == len3:
            list1.append(0)
            break
        else:
            cell_3 = sheet3.cell(a3, 0).value
            if cell_3 == dh:
                list1.append(a3)
                b3 = a3 + 1
                break

    for a4 in range(b4, list3[4] + 1):
        len4 = list3[4]
        if a4 == len4:
            list1.append(0)
            break
        else:
            cell_4 = sheet4.cell(a4, 0).value
            if cell_4 == dh:
                list1.append(a4)
                b4 = a4 + 1
                break

    for a5 in range(b5, list3[5]):
        len5 = list3[5]
        if b5 == len5:
            list2.append(0)
            break
        else:
            cell_5 = sheet5.cell(a5, 0).value
            if cell_5 != dh:
                if len(list2) > 0:
                    break
                else:
                    list2.append(0)
                    break
            else:
                list2.append(a5)
                b5 = a5 + 1

    xr(list1, list2, dh)


# 新建模板并写入数据
def xr(list1, list2, dh):
    global cs

    #复制模板并写入数据保存
    oldWb = xlrd.open_workbook('D:\\fuwufei\\模板.xls',formatting_info=True)  # 先打开已存在的表
    newWb = copy(oldWb)  # 复制
    sheet6 = newWb.get_sheet(0)  # 取sheet表

    # 新建样式
    style = XFStyle()
    style2 = XFStyle()
    style4 = XFStyle()
    style5 = XFStyle()

    # 设置边框
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    style.borders = borders
    style4.borders = borders
    style5.borders = borders

    # 设置对其al（居中）
    al = xlwt.Alignment()
    al.horz = 0x02  # 设置水平居中
    al.vert = 0x01  # 设置垂直居中
    style.alignment = al
    style2.alignment = al
    style4.alignment = al
    style5.alignment = al

    # 设置对其al1（靠左）
    al1 = xlwt.Alignment()
    al1.horz = 0x01
    al1.vert = 0x01
    al1.wrap = 1  # 设置自动换行
    style4.alignment = al1

    # 设置字体格式fnt
    fnt = Font()
    fnt.name = u'宋体'
    fnt.bold = True
    fnt.height = 220
    style.font = fnt
    style5.font = fnt

    # 设置字体格式fnt1（红色）
    fnt1 = Font()
    fnt1.name = u'宋体'
    fnt1.bold = True
    fnt1.height = 220
    fnt1.colour_index = 2

    # 设置字体格式fnt2（蓝色）
    fnt2 = Font()
    fnt2.name = u'宋体'
    fnt2.bold = True
    fnt2.height = 220
    fnt2.colour_index = 4
    style4.font = fnt2
    style5.font = fnt2

    #写入月份
    sheet6.write(1,8,  ym, style)

    # 从表中取值
    for i in range(0, len(list1)):
        # 在第一个sheet中找数据并填入前三个需要填的值
        if i == 0:
            # 循环表的行数
            for l in range(0, list4[0]):
                xx = sheet0.cell(list1[0], l).value  # 获取对应行的全部信息
                # 如果取值为空，就在对应的框填入空白，否则填入对应值
                if xx == '':
                    if l == 0:
                        sheet6.write(0, 4, '',style)
                    elif l == 1:
                        sheet6.write(1, 2, '',style)
                    else:
                        sheet6.write(2, 2, '',style)
                    continue
                else:
                    if l == 0:
                        sheet6.write(0, 4, xx,style)
                    elif l == 1:
                        sheet6.write(1, 2, xx,style)
                    else:
                        sheet6.write(2, 2, xx,style)


        # 在第二个sheet中找数据并填入4，5需要填的值
        elif i == 1:
            # 如果list1中对应表中的值为0，说明该表中不存在匹配值的数据，就直接在对应的框填入空白
            if list1[1] == 0:
                sheet6.write(3, 2, ' ', style)
                sheet6.write(3, 8, ' ', style)

            else:
                #
                for l in range(1, list4[1]):
                    xx = sheet1.cell(list1[1], l).value  # 获取对应行的全部信息
                    if xx == '':
                        if l == 1:
                            sheet6.write(3, 2,' ', style)
                        else:
                            sheet6.write(3, 8, ' ', style)
                    else:
                        if l == 1:
                            sheet6.write(3,  2,  xx, style)
                        else:
                            sheet6.write(3, 8,  xx, style)

        # 在第三个sheet中找数据并填入6需要填的值
        elif i == 2:
            if list1[2] == 0:
                sheet6.write( 5, 2,  '', style4)
                continue
            else:
                for l in range(1, list4[2]):
                    xx = sheet2.cell(list1[2], l).value  # 获取对应行的全部信息
                    if xx == '':
                        sheet6.write( 5, 2,  '', style4)
                        continue
                    else:
                        sheet6.write(5,  2,  xx, style4)

        # 在第四个sheet中找数据并填入7需要填的值
        elif i == 3:
            if list1[3] == 0:
                sheet6.write(6, 2,  '', style)
                continue
            else:
                for l in range(1, list4[3]):
                    xx = sheet3.cell(list1[3], l).value  # 获取对应行的全部信息
                    if xx == '':
                        sheet6.write(6, 2, '', style)
                        continue
                    else:
                        sheet6.write( 6, 2,  xx, style)

        # 在第五个sheet中找数据并填入7需要填的值
        else:
            if list1[4] == 0:
                sheet6.write(7, 2,  '', style)
                continue
            else:
                for l in range(1, list4[4]):
                    xx = sheet4.cell(list1[4], l).value  # 获取对应行的全部信息
                    if xx == '':
                        sheet6.write(7, 2,  '', style)
                        continue
                    else:
                        sheet6.write(7, 2, xx, style)

    # 在第六个sheet中找数据并填入
    for l in range(0, len(list2)):
        if list2[0] == 0:
            break
        else:
            for cc in range(0, list4[5]-1):
                xx = sheet5.cell(list2[l], cc).value
                if xx == '':
                    continue
                else:
                    sheet6.write(9 + l, cc, xx, style2)


    file_name = dir_path + '\\'+ dh + '.xls'
    newWb.save(file_name)  # 保存

    cs = cs + 1
    print(cs)


if __name__ == '__main__':
    b0 = 1
    b1 = 1
    b2 = 1
    b3 = 1
    b4 = 1
    b5 = 1
    cs = 0
    list3 = []  # 存放每个表的行数
    list4 = []  # 存放每个表的列数
    sheets_list = []  # 存放合并后的sheet名列表
    save_path = 'D:\\fuwufei\\xx.xls'  # 合并后excel的存放路径

    # 获取当前月份-1
    year = datetime.datetime.now().year
    month = datetime.datetime.now().month
    if month < 10:
        ym = str(year) + '0' + str(month - 1)
    else:
        ym = str(year) + str(month - 1)

    #创建文件夹存放各店信息
    dir_path = 'D:\\fuwufei\\' + ym + '服务费清单'
    os.mkdir(dir_path)

    # 整合排序多个excel的多个sheet为一个excel的多个sheet
    ##门店信息中有数字框要设置为文本格式
    path = 'D:\\fuwufei\\c\\'
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
            if '原始数据' in a:
                df.sort_values(by=[btname, '顺序'], ascending=True, inplace=True)  # 使用第一个字段来升序。ascending排序，inplace替代
            else:
                df.sort_values(by=btname, ascending=True, inplace=True)  # 使用第一个字段来升序。ascending排序，inplace替代
            df.to_excel(writer, encoding='utf-8', sheet_name=a[:-5], index=None)  # 写入sheet中 index无索引
            writer.save()
            writer.close()
            sheets_list.append(a[:-5])
            break

    print(sheets_list)

    data = xlrd.open_workbook(save_path)
    # 按顺序打开各个sheet表并获取行列
    for num in range(0, len(sheets_list)):
        for sheet_list in sheets_list:
            if ('门店信息' in sheet_list) and num == 0:
                sheet0 = data.sheet_by_name(sheet_list)
                rows = sheet0.nrows
                cols = sheet0.ncols
                list3.append(rows)
                list4.append(cols)
                break
            elif ('业绩汇总' in sheet_list) and num == 1:
                sheet1 = data.sheet_by_name(sheet_list)
                rows = sheet1.nrows
                cols = sheet1.ncols
                list3.append(rows)
                list4.append(cols)
                break
            elif ('补扣款备注' in sheet_list) and num == 2:
                sheet2 = data.sheet_by_name(sheet_list)
                rows = sheet2.nrows
                cols = sheet2.ncols
                list3.append(rows)
                list4.append(cols)
                break
            elif ('200元以下合计' in sheet_list) and num == 3:
                sheet3 = data.sheet_by_name(sheet_list)
                rows = sheet3.nrows
                cols = sheet3.ncols
                list3.append(rows)
                list4.append(cols)
                break
            elif ('店补发放' in sheet_list) and num == 4:
                sheet4 = data.sheet_by_name(sheet_list)
                rows = sheet4.nrows
                cols = sheet4.ncols
                list3.append(rows)
                list4.append(cols)
                break
            elif ('服务费清单原始数据' in sheet_list) and num == 5:
                sheet5 = data.sheet_by_name(sheet_list)
                rows = sheet5.nrows
                cols = sheet5.ncols
                list3.append(rows)
                list4.append(cols)
                break

    print(list3, list4)

    # 循环服务费清单原始数据表中店号
    for xhdh in range(0, list3[5] - 1):
        dh1 = sheet5.cell(xhdh, 0).value  # 获取每行店号
        dh2 = sheet5.cell(xhdh + 1, 0).value
        if dh1 != dh2:
            xhs(dh2)
        else:
            continue




