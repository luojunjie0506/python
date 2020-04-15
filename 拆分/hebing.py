import  xlrd,xlwt
from xlutils.copy import copy

'''
前提：先把每个表中的数据按字母降序排序，这样每次匹配行数的时候会少执行很多次
1.打开门店信息表，获取行数，循环行数，获取店号。
1.1.用店号去匹配每个表中对应数据的行数（用字典或者列表存进来）
1.2.复制模板后，用存起来的行数去查询对应表中的数据并写入复制的模板中并店号或名字保存
2.再次执行前面步骤


'''

'''
#循环表
def xhs(dh,hs):
    global b1, b2, b3, b4, b5
    global sheet1, sheet2, sheet3, sheet4, sheet5
    list1 =[] #存储匹配前五个表中对应数据的行数
    list2 =[]  # 存储匹配最后一个表中对应数据的行数
    list1.append(hs)

    nrows1 = sheet1.nrows  # 获取sheets的行数
    for a1 in range(b1,nrows1):
        cell_1 = sheet1.cell(a1, 0).value
        if b1+1 == nrows1:
            list1.append(0)
            break
        else:
            if cell_1 != dh:
                list1.append(0)
                break
            else:
                list1.append(a1)
                b1 = a1 + 1
                break


    nrows2 = sheet2.nrows  # 获取sheets的行数
    for a2 in range(b2,nrows2):
        cell_2 = sheet2.cell(a2, 0).value
        if b2+1 == nrows2:
            list1.append(0)
            break
        else:
            if cell_2 != dh:
                list1.append(0)
                break
            else:
                list1.append(a2)
                b2 = a2+1
                break


    nrows3 = sheet3.nrows  # 获取sheets的行数
    for a3 in range(b3,nrows3):
        cell_3 = sheet3.cell(a3, 0).value
        if b3+1 == nrows3:
            list1.append(0)
            break
        else:
            if cell_3 != dh:
                list1.append(0)
                break
            else:
                list1.append(a3)
                b3 = a3+1
                break

    nrows4 = sheet4.nrows  # 获取sheets的行数
    for a4 in range(b4,nrows4):
        cell_4= sheet4.cell(a4, 0).value
        if b4+1 == nrows4:
            list1.append(0)
            break
        else:
            if cell_4 != dh:
                list1.append(0)
                break
            else:
                list1.append(a4)
                b4 = a4+1
                break

    nrows5 = sheet5.nrows  # 获取sheets的行数
    print(nrows5)
    for a in range(b5,nrows5):
        cell_5 = sheet5.cell(a, 0).value
        if b5+1 == nrows5:
            list2.append(0)
            break
        else:
            if cell_5 != dh:
                if len(list2) >0:
                    break
                else:
                    list2.append(0)
                    break
            else:
                list2.append(a)
                b5 = a+1

    #复制模板,formatting_info保存原格式
    file = 'D:\\sad.xls'
    data = xlrd.open_workbook(file, formatting_info=True)
    new_workbook = copy(data)

    ff = r'asdas.xls'
    new_workbook.save(ff)



if __name__ == '__main__':
    file = 'D:\\sad.xls'
    data = xlrd.open_workbook(file)
    sheet0 = data.sheets()[0] #打开第一个sheets
    sheet1 = data.sheets()[1]
    sheet2 = data.sheets()[2]
    sheet3 = data.sheets()[3]
    sheet4 = data.sheets()[4]
    sheet5 = data.sheets()[5]
    nrows0 = sheet0.nrows # 获取第一个sheets的行数
    b1 = 1
    b2 = 1
    b3 = 1
    b4 = 1
    b5 = 1

    rows = table.row_values(0)  # 获取第二行内容 列表
    cell_A1 = table.cell(0, 0).value #获取第一行第一列的值



    #循环第一个表中店号
    for xhdh in range(1,nrows0):
        dh = sheet0.cell(xhdh, 0).value  #获取每行店号
        #从第二个sheets开始循环sheets
        xhs(dh, xhdh)

    '''



























#复制模板,formatting_info保存原格式
file = 'D:\\模板.xls'
data = xlrd.open_workbook(file, formatting_info=True)
new_workbook = copy(data)
new_workbook.get_sheet(0).write(0,4,label = 'this is test5645645645645')
ff = r'D:\\asdas.xls'
new_workbook.save(ff)
