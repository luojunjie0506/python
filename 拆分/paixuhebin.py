import  pandas as pd
import numpy as np
import xlwt,xlrd
from xlwt import *
import os

###门店信息中有数字的列表不能用来排序ss

path='D:\\zz\\'
file_list=os.listdir(path) #获取文件夹中的所有文件名的列表
writer = pd.ExcelWriter('D:\\xx.xls')  #使用ExcelWriter函数，可以写入数据时不会覆盖sheet

#遍历文件夹中的所有文件名的列表
for a  in file_list:
    file_path = path + a  #拼接完整路径
    data = pd.read_excel(file_path, sheet_name=0) #读取文件夹中的表
    df = pd.DataFrame(data) #转换格式

    #循环每个表中表头的第一个字段，并用来排序
    for btname  in data:
        df.sort_values(by=btname, ascending=True, inplace=True)  #使用第一个字段来升序。ascending排序，inplace替代
        df.to_excel(writer, encoding='utf-8', sheet_name=a[:-4],index=None) #写入sheet中 index无索引
        writer.save()
        writer.close()
        print(a[:-4])
        break



