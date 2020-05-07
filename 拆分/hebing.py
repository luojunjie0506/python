import  xlrd,xlwt
from xlwt import *
import datetime

'''
前提：先把每个表中的数据按字母降序排序，这样每次匹配行数的时候会少执行很多次
1.打开服务费清单原始表，获取店号。
1.1.用店号去匹配每个表中对应数据的行数（用列表存进来）
1.2.增加写模板，用存起来的行数去查询对应表中的数据并写入复制的模板中并店号或名字保存
2.再次执行前面步骤
'''

#循环表
def xhs(dh):
    global b0,b1, b2, b3, b4, b5,ym
    global sheet0,sheet1, sheet2, sheet3, sheet4, sheet5
    global list3,list4
    list1 =[] #存储匹配前五个表中对应数据的行数
    list2 =[]  # 存储匹配最后一个表中对应数据的行数

    # 找到每个表的行数，如有存在行数就追加到list1,如不存在，在list1中加0后到后一个表匹配的行数
    # b参数是记录每次查询到的位置
    for a0 in range(b0,list3[0]+1):
        len0 = list3[0]
        if b0 == len0:
            list1.append(0)
            break
        elif a0 == len0-1:
            list1.append(0)
            break
        else:
            cell_0 = sheet0.cell(a0, 0).value
            # 前五个表有且最多只能存在一条数据，判断一次是否等于匹配值就可以
            if cell_0 == dh:
                list1.append(a0)
                b0 = a0+1
                break

    for a1 in range(b1,list3[1]+1):
        len1 = list3[1]
        if a1 == len1:
            list1.append(0)
            break
        elif b1 == len1:
            list1.append(0)
            break
        else:
            cell_1 = sheet1.cell(a1, 0).value
            if cell_1 == dh:
                list1.append(a1)
                b1 = a1+1
                break

    for a2 in range(b2,list3[2]+1):
        len2 = list3[2]
        if a2 == len2:
            list1.append(0)
            break
        elif b2 == len2:
            list1.append(0)
            break
        else:
            cell_2 = sheet2.cell(a2, 0).value
            if cell_2 == dh:
                list1.append(a2)
                b2 = a2 + 1
                break

    for a3 in range(b3,list3[3]+1):
        len3 = list3[3]
        if a3 == len3:
            list1.append(0)
            break
        elif b3 == len3:
            list1.append(0)
            break
        else:
            cell_3 = sheet3.cell(a3, 0).value
            if cell_3 == dh:
                list1.append(a3)
                b3 = a3+1
                break

    for a4 in range(b4,list3[4]+1):
        len4 = list3[4]
        if a4 == len4:
            list1.append(0)
            break
        elif b4 == len4:
            list1.append(0)
            break
        else:
            cell_4 = sheet4.cell(a4, 0).value
            if cell_4 == dh:
                list1.append(a4)
                b4 = a4+1
                break

    for a5 in range(b5,list3[5]):
        len5 = list3[5]
        if b5 == len5:
            list2.append(0)
            break
        else:
            cell_5 = sheet5.cell(a5, 0).value
            if cell_5 != dh:
                if len(list2) >0:
                    break
                else:
                    list2.append(0)
                    break
            else:
                list2.append(a5)
                b5 = a5+1

    xr(list1,list2,dh)




#新建模板并写入数据
def xr(list1,list2,dh):
    global style2 ,cs

    a = "一、服务费清单发放说明：服务费清单只提供电子档一种格式，如需图片格式，可联系公司客服。\n二、服务费发放说明：服务费清单中应发服务费金额包含直接服务奖金额。\n三、服务费发放说明：服务费清单中增加“转货款”列，应发服务费金额已扣减转货款金额。"
    workbook = xlwt.Workbook(encoding='utf-8')
    sheet6 = workbook.add_sheet("表格")  # 新建sheet

    #新建样式
    style = XFStyle()
    style2 = XFStyle()
    style3 = XFStyle()
    style4 = XFStyle()
    style5 = XFStyle()

    # 设置边框
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    style.borders = borders
    style3.borders = borders
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
    style3.alignment = al1
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
    style3.font = fnt1

    # 设置字体格式fnt2（蓝色）
    fnt2 = Font()
    fnt2.name = u'宋体'
    fnt2.bold = True
    fnt2.height = 220
    fnt2.colour_index = 4
    style4.font = fnt2
    style5.font = fnt2

    # 写入数据
    sheet6.write_merge(0, 0, 0, 3, "专卖店编号：", style)
    sheet6.write_merge(1, 1, 0, 1, "店主姓名：", style)
    sheet6.write(1, 7, "月份：", style)
    sheet6.write_merge(1, 1, 8, 10, ym, style)
    sheet6.write_merge(2, 2, 0, 1, "Email：", style)
    sheet6.write_merge(3, 3, 0, 1, "报单数量：", style)
    sheet6.write(3, 7, "报单额：", style)
    sheet6.write_merge(4, 4, 0, 10, a, style3)
    sheet6.write_merge(5, 5, 0, 1, "备注：", style5)
    sheet6.write_merge(6, 6, 0, 1, "200元以下合计", style)
    sheet6.write_merge(7, 7, 0, 1, "店补合计", style)


    #从表中取值
    for i in range(0, len(list1)):
        #在第一个sheet中找数据并填入前三个需要填的值
        if i == 0:
            #循环表的行数
            for l in range(0, list4[0] ):
                xx = sheet0.cell(list1[0], l).value  # 获取对应行的全部信息
                #如果取值为空，就在对应的框填入空白，否则填入对应值
                if xx == '':
                    if l==0:
                        sheet6.write_merge(0, 0, 4, 10, '', style)
                    elif l == 1:
                        sheet6.write_merge(1, 1, 2, 6, '', style)
                    else:
                        sheet6.write_merge(2, 2, 2, 10, '', style)
                    continue
                else:
                    if l==0:
                        sheet6.write_merge(0, 0, 4, 10, xx, style)
                    elif l == 1:
                        sheet6.write_merge(1, 1, 2, 6, xx, style)
                    else:
                        sheet6.write_merge(2, 2, 2, 10, xx, style)

        # 在第二个sheet中找数据并填入4，5需要填的值
        elif i == 1:
            # 如果list1中对应表中的值为0，说明该表中不存在匹配值的数据，就直接在对应的框填入空白
            if list1[1] == 0:
                    sheet6.write_merge(3, 3, 2, 6, ' ', style)
                    sheet6.write_merge(3, 3, 8, 10, ' ', style)

            else:
                #
                for l in range(1, list4[1]):
                    xx = sheet1.cell(list1[1], l).value  # 获取对应行的全部信息
                    if xx == '':
                        if l == 1:
                            sheet6.write_merge(3, 3, 2, 6, ' ', style)
                        else:
                            sheet6.write_merge(3, 3, 8, 10, ' ', style)
                    else:
                        if l == 1:
                            sheet6.write_merge(3, 3, 2, 6, xx, style)
                        else:
                            sheet6.write_merge(3, 3, 8, 10, xx, style)

        # 在第三个sheet中找数据并填入6需要填的值
        elif i == 2:
            if list1[2] == 0:
                sheet6.write_merge(5, 5, 2, 10, '', style4)
                continue
            else:
                for l in range(1, list4[2] ):
                    xx = sheet2.cell(list1[2], l).value  # 获取对应行的全部信息
                    if xx == '':
                        sheet6.write_merge(5, 5, 2, 10, '', style4)
                        continue
                    else:
                        sheet6.write_merge(5, 5, 2, 10, xx, style4)

        # 在第四个sheet中找数据并填入7需要填的值
        elif i == 3:
            if list1[3] == 0:
                sheet6.write_merge(6, 6, 2, 10, '', style)
                continue
            else:
                for l in range(1, list4[3]):
                    xx = sheet3.cell(list1[3], l).value  # 获取对应行的全部信息
                    if xx == '':
                        sheet6.write_merge(6, 6, 2, 10, '', style)
                        continue
                    else:
                        sheet6.write_merge(6, 6, 2, 10, xx, style)

        # 在第五个sheet中找数据并填入7需要填的值
        else:
            if list1[4] == 0:
                sheet6.write_merge(7, 7, 2, 10, '', style)
                continue
            else:
                for l in range(1, list4[4]):
                    xx = sheet4.cell(list1[4], l).value # 获取对应行的全部信息
                    if xx == '':
                        sheet6.write_merge(7, 7, 2, 10, '', style)
                        continue
                    else:
                        sheet6.write_merge(7, 7, 2, 10, xx, style)

    # 在第六个sheet中找数据并填入
    for l in range(0, len(list2)):
        if list2[0] == 0:
            break
        else:
            for cc in range(0,11):
                xx = sheet5.cell(list2[l], cc).value
                if xx == '':
                    continue
                else:
                    sheet6.write(9+l, cc, xx, style2)


    #设置最后一行字段显示
    list1 = ["店号", "卡号", "姓名", "农行卡号", "职级", "职称", "市场业绩", "直接服务奖", "转货款", "应发服务费", "服务费发放状态"]
    for i in range(0, 11):
        sheet6.write(8, i, list1[i], style2)

    # 设置1到3行的高度
    for i in range(0, 4):
        tall_style = xlwt.easyxf('font:height 216;')
        first_row = sheet6.row(i)
        first_row.set_style(tall_style)

    tall_style = xlwt.easyxf('font:height 1056;')
    first_row = sheet6.row(4)
    first_row.set_style(tall_style)
    tall_style = xlwt.easyxf('font:height 2416;')
    first_row = sheet6.row(5)
    first_row.set_style(tall_style)
    tall_style = xlwt.easyxf('font:height 470;')
    first_row = sheet6.row(6)
    first_row.set_style(tall_style)
    tall_style = xlwt.easyxf('font:height 470;')
    first_row = sheet6.row(7)
    first_row.set_style(tall_style)

    # 设置列宽
    for i in range(0, 3):
        first_col = sheet6.col(0)
        first_col.width = 256 * 8

    first_col = sheet6.col(3)
    first_col.width = 256 * 23
    first_col = sheet6.col(4)
    first_col.width = 256 * 10
    first_col = sheet6.col(5)
    first_col.width = 256 * 14
    first_col = sheet6.col(6)
    first_col.width = 256 * 11
    first_col = sheet6.col(7)
    first_col.width = 256 * 11
    first_col = sheet6.col(8)
    first_col.width = 256 * 11
    first_col = sheet6.col(9)
    first_col.width = 256 * 11
    first_col = sheet6.col(10)
    first_col.width = 256 * 18

    file_name = 'D:\\zz\\table\\'+dh+'.xls'
    workbook.save(file_name)  # 保存

    cs =cs +1
    print(cs)

if __name__ == '__main__':
    b0 = 1
    b1 = 1
    b2 = 1
    b3 = 1
    b4 = 1
    b5 = 1
    cs = 0
    list3 = [] #存放每个表的行数
    list4 = []  # 存放每个表的列数

    #获取当前月份-1
    year = datetime.datetime.now().year
    month = datetime.datetime.now().month
    if month <10:
        ym = str(year) +'0'+ str(month-1)
    else:
        ym = str(year) + str(month - 1)

    file = 'D:\\zz\\cc1.xls'
    data = xlrd.open_workbook(file)
    #打开sheet表
    sheet0 = data.sheets()[0] #门店信息表
    sheet1 = data.sheets()[1] #业绩汇总表
    sheet2 = data.sheets()[2] #补扣款备注
    sheet3 = data.sheets()[3] #200元以下合计
    sheet4 = data.sheets()[4] #店补发放
    sheet5 = data.sheets()[5] #服务费清单原始数据

    #获取每表的行数
    for i in range(0,6):
        rows = data.sheets()[i].nrows
        list3.append(rows)

    # 获取每表的列数
    for i in range(0,6):
        cols = data.sheets()[i].ncols
        list4.append(cols)

    
    #循环服务费清单原始数据表中店号
    for xhdh in range(0,list3[5]-1):
        dh1 = sheet5.cell(xhdh, 0).value  #获取每行店号
        dh2 = sheet5.cell(xhdh+1, 0).value
        if dh1 != dh2:
            xhs(dh2)
        else:
            continue






























