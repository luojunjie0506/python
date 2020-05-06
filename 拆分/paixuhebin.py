import  pandas as pd
import numpy as np
import xlwt,xlrd
from xlwt import *
import os



path='D:\\c\\'
file_list=os.listdir(path) #遍历文件夹中的所有文件

'''
for a  in file_list:
    file_path = path + a
    print(file_path)

    data = pd.read_excel(file_path, sheet_name=0)
    df = pd.DataFrame(data)
    df.sort_values('店号',ascending=True)
'''

data = pd.read_excel('D:\\c\\202004 店补发放.xlsx', sheet_name=0)
df = pd.DataFrame(data)
df.sort_values(by=['店号'],ascending=True,inplace=True)
df.to_excel('D:\\a.xls',encoding='utf-8', index=False, header=False)