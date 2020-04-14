import  xlrd,xlwt
from xlutils.copy import copy
import pandas as pd

'''
设想：可以先把每个表中的数据按字母降序排序，这样每次匹配行数的时候会少执行很多次
1.打开门店信息表，获取行数，循环行数，获取店号。
1.1.用店号去匹配每个表中对应数据的行数（用字典或者列表存进来）
1.2.复制模板后，用存起来的行数去查询对应表中的数据并写入复制的模板中并店号或名字保存
2.再次执行前面步骤


'''

if __name__ == '__main__':
    #排序未完成
    for i in range(0,2):
        stexcel = pd.read_excel('D:/test.xls', sheet_name=i)
        stexcel.sort_values(by='姓名', inplace=True, ascending=True)
        stexcel.to_excel('D:/test.xls', index=False, encoding='utf-8_sig',sheet_name=str(i))




    '''
    file = 'D:\\sad.xls'
    data = xlrd.open_workbook(file, formatting_info=True)
    #获取第一个表
    table = data.sheets()[0]
    #获取行数
    nrows = table.nrows
    
    #复制模板
    new_workbook = copy(data)
    ff = r'asdas.xls'
    new_workbook.save(ff)
    '''''