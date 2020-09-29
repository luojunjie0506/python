import xlrd
import xlwt
from xlutils.copy import copy
from xlwt import XFStyle, Font


def xx(list):
    n = 5
    num = 1
    oldWb = xlrd.open_workbook('D:\\分红\\模板.xls', formatting_info=True)  # 先打开已存在的表
    newWb = copy(oldWb)  # 复制
    sheet6 = newWb.get_sheet(0)

    # 新建样式
    style1 = XFStyle()
    style2 = XFStyle()
    style3 = XFStyle()
    style4 = XFStyle()
    style5 = XFStyle()

    # 设置边框1
    borders1 = xlwt.Borders()
    borders1.left = xlwt.Borders.THIN
    borders1.right = xlwt.Borders.THIN
    borders1.top = xlwt.Borders.THIN
    borders1.bottom = xlwt.Borders.THIN

    # 设置边框2
    borders2 = xlwt.Borders()
    borders2.left = xlwt.Borders.THIN
    borders2.right = xlwt.Borders.HAIR
    borders2.top = xlwt.Borders.HAIR

    # 设置边框3
    borders3 = xlwt.Borders()
    borders3.left = xlwt.Borders.HAIR
    borders3.right = xlwt.Borders.HAIR
    borders3.top = xlwt.Borders.HAIR

    # 设置边框4
    borders4 = xlwt.Borders()
    borders4.left = xlwt.Borders.HAIR
    borders4.right = xlwt.Borders.THIN
    borders4.top = xlwt.Borders.HAIR

    # 设置边框5
    borders5 = xlwt.Borders()
    borders5.top = xlwt.Borders.THIN

    # 设置字体格式fnt
    fnt = Font()
    fnt.name = u'微软雅黑'
    fnt.height = 220

    # 设置对其al（居中）
    al = xlwt.Alignment()
    al.horz = 0x02  # 设置水平居中
    al.vert = 0x01  # 设置垂直居中
    al.wrap = 1

    style1.alignment = al
    style1.font = fnt
    style1.borders = borders1

    style2.alignment = al
    style2.font = fnt
    style2.borders = borders2

    style3.alignment = al
    style3.font = fnt
    style3.borders = borders3

    style4.alignment = al
    style4.font = fnt
    style4.borders = borders4

    style5.borders = borders5

    #写入第一个sheet数据
    for i in range(1,len(list)+1):
        sheet6.write(2, i, list[i-1],style1)

    #写入第二个sheet数据
    for r1 in range(2,row1):
        card4 = sheet1.cell(r1, 2).value
        if  card4 == list[0]:
            card6 = sheet1.cell(r1, 4).value
            name4 = sheet1.cell(r1, 5).value
            pgv = sheet1.cell(r1, 6).value
            db = sheet1.cell(r1, 7).value
            sheet6.write(n, 1, num,style2)
            sheet6.write(n, 2, card6,style3)
            sheet6.write(n, 3, name4,style3)
            sheet6.write(n, 4, pgv,style3)
            sheet6.write(n, 5, db,style4)
            num = num + 1
            n = n + 1

    sheet6.write(n, 1, '',style5)
    sheet6.write(n, 2, '',style5)
    sheet6.write(n, 3, '',style5)
    sheet6.write(n, 4, '',style5)
    sheet6.write(n, 5, '',style5)



    file_name = 'D:\\分红\\fh1\\' + name3 + '.xls'
    newWb.save(file_name)  # 保存


if __name__ == '__main__':
    dd ='全球分红（截止8期次累计PGV不小于80万）(1).xlsx'
    path = 'D:\\分红\\' + dd  #文件路径
    data = xlrd.open_workbook(path) #打开excel
    #获取sheet数量
    count = len(data.sheets())
    # 打开sheet0
    sheet0 = data.sheet_by_index(0)
    # 打开sheet1
    sheet1 = data.sheet_by_index(1)
    #获取行数
    row0 = sheet0.nrows
    #获取行数
    row1 = sheet1.nrows




    #循环sheet
    for a in range(6,row0):
        list1 = []
        card3 = sheet0.cell(a, 2).value
        list1.append(card3)
        name3 = sheet0.cell(a, 3).value
        list1.append(name3)
        old_fh = sheet0.cell(a, 4).value
        list1.append(old_fh)
        now_fh = sheet0.cell(a, 5).value
        list1.append(now_fh)
        jl_fh = sheet0.cell(a, 6).value
        list1.append(jl_fh)
        xx(list1)


