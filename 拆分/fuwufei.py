import copy
import datetime
import openpyxl
import os
import pandas as pd
import xlrd
from openpyxl.utils import get_column_letter

def getmonth():
    # 获取当前月份-1
    year = datetime.datetime.now().year
    month = datetime.datetime.now().month
    if month < 10:
        ym = str(year) + '0' + str(month - 1)
    else:
        ym = str(year) + str(month - 1)
    return ym

# 整合文件夹中的表到一个excel中
def zhenghe():
    global sheets_list
    path = 'D:\\fuwufei2\\c\\' # 文件夹路径
    save_path = 'D:\\fuwufei2\\xx.xlsx'  # 合并后excel的存放路径
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

# 复制模板
def copyMode():
    path = 'D:\\Book1.xlsx' # 模板路径
    save_path = 'D:\\Book2.xlsx' #保存路径

    wb = openpyxl.load_workbook(path)
    wb2 = openpyxl.Workbook()

    sheetnames = wb.sheetnames
    for sheetname in sheetnames:
        print(sheetname)
        sheet = wb[sheetname]
        sheet2 = wb2.create_sheet(sheetname)

        # tab颜色
        sheet2.sheet_properties.tabColor = sheet.sheet_properties.tabColor

        # 开始处理合并单元格形式为“(<CellRange A1：A4>,)，替换掉(<CellRange 和 >,)' 找到合并单元格
        wm = list(sheet.merged_cells)
        if len(wm) > 0:
            for i in range(0, len(wm)):
                cell2 = str(wm[i]).replace('(<CellRange ', '').replace('>,)', '')
                sheet2.merge_cells(cell2)

        for i, row in enumerate(sheet.iter_rows()):
            sheet2.row_dimensions[i+1].height = sheet.row_dimensions[i+1].height
            for j, cell in enumerate(row):
                sheet2.column_dimensions[get_column_letter(j+1)].width = sheet.column_dimensions[get_column_letter(j+1)].width
                sheet2.cell(row=i + 1, column=j + 1, value=cell.value)

                # 设置单元格格式
                source_cell = sheet.cell(i+1, j+1)
                target_cell = sheet2.cell(i+1, j+1)
                target_cell.fill = copy.copy(source_cell.fill)
                if source_cell.has_style:
                    target_cell._style = copy.copy(source_cell._style)
                    target_cell.font = copy.copy(source_cell.font)
                    target_cell.border = copy.copy(source_cell.border)
                    target_cell.fill = copy.copy(source_cell.fill)
                    target_cell.number_format = copy.copy(source_cell.number_format)
                    target_cell.protection = copy.copy(source_cell.protection)
                    target_cell.alignment = copy.copy(source_cell.alignment)

    if 'Sheet' in wb2.sheetnames:
        del wb2['Sheet']
    wb2.save(save_path)

    wb.close()
    wb2.close()

    print('Done.')

def xhs(dh):
    global b0, b1, b2, b3, b4, b5
    for a0 in range(b0, rows):
        # b的位置跟长度相等时，说明表中数据一匹配完，直接在list1中加1
        if b0 == rows:
            list1.append(0)
            break
        else:
            cell_0 = sheet0.cell(a0, 0).value
            # 前五个表有且最多只能存在一条数据，判断一次是否等于匹配值就可以
            if cell_0 == dh:
                list1.append(a0)
                b0 = a0 + 1
                break



if __name__ == '__main__':
    b0 = 1
    b1 = 1
    b2 = 1
    b3 = 1
    b4 = 1
    b5 = 1
    sheets_list = []  # 存放合并后的sheet名列表
    sheets_list = ['202012 小微店补', '202012 服务费清单原始数据', '202012 活动明细表', '202012 补扣款备注', '202012 门店信息表']
    # # zhenghe_path = zhenghe()

    data = xlrd.open_workbook('D:\\fuwufei2\\xx.xlsx')
    sheet0 = data.sheet_by_name("202012 服务费清单原始数据")
    rows = sheet0.nrows
    dh_list= []
    sheet0.row(5)
    for xhdh in range(1,  rows):
        dh1 = sheet0.cell(xhdh-1, 0).value
        dh2 = sheet0.cell(xhdh, 0).value # 需要制作清单的店号
        print(dh2)
        if dh1 != dh2 :
            # xhs(dh2)
            pass