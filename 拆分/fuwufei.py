# coding=gbk
import datetime
import os
import pandas as pd
import xlrd
from openpyxl import *
import numpy as np

def getmonth():
    # ��ȡ��ǰ�·�-1
    year = datetime.datetime.now().year
    month = datetime.datetime.now().month
    if month ==1 :
        ym = str(year-1) + str(12)
    elif month < 10:
        ym = str(year) + '0' + str(month - 1)
    else:
        ym = str(year) + str(month - 1)
    return ym

# �����ļ����еı�һ��excel��
def zhenghe():
    global sheets_list
    path = 'D:\\fuwufei\\c\\' # �ļ���·��
    save_path = 'D:\\fuwufei\\xx.xlsx'  # �ϲ���excel�Ĵ��·��
    file_list = os.listdir(path)  # ��ȡ�ļ����е������ļ������б�
    writer = pd.ExcelWriter(save_path)  # ʹ��ExcelWriter����������д������ʱ���Ḳ��sheet
    # �����ļ����е������ļ������б�
    for a in file_list:
        file_path = path + a  # ƴ������·��
        data = pd.read_excel(file_path, sheet_name=0)  # ��ȡ�ļ����еı�
        df = pd.DataFrame(data)  # ת����ʽ

        # ѭ��ÿ�����б�ͷ�ĵ�һ���ֶΣ�����������
        for btname in data:
            #�жϱ��Ƿ��Ƿ�����嵥ԭʼ���ݱ��ǵĻ���Ҫ���е�һ���ֶκ�˳���ֶ�����
            # if 'ԭʼ����' in a:
            #     df.sort_values(by=[btname, '˳��'], ascending=True, inplace=True)  # ʹ�õ�һ���ֶ�������ascending����inplace���
            # else:
            #     df.sort_values(by=btname, ascending=True, inplace=True)  # ʹ�õ�һ���ֶ�������ascending����inplace���
            df.sort_values(by=btname, ascending=True, inplace=True)
            df.to_excel(writer, encoding='utf-8', sheet_name=a[:-5], index=None)  # д��sheet�� index������
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
    # '202012 ������嵥ԭʼ����','202012 �ŵ���Ϣ��', '202012 С΢�겹', '202012 ���ۿע',  '202012 ���ϸ��'
    list1 = []  # �洢ƥ�������嵥ԭʼ����
    list2 = []  # �洢ƥ�䲹�ۿע,С΢�겹,�ŵ���Ϣ��
    list3 = []  # �洢ƥ����ϸ��

    # �ҵ�ÿ��������������д���������׷�ӵ�list1,�粻���ڣ���list1�м�0�󵽺�һ����ƥ�������
    # b�����Ǽ�¼ÿ�β�ѯ����λ��
    for a0 in range(b0, row_list[0]):
        mh0 = []
        # b��λ�ø��������ʱ��˵����������һƥ���ֱ꣬����list�м�0
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

    # forѭ�������range�е�����������һ��ʱ,�����ѭ���Ͳ�ִ��,Ϊ����ѭ������ִ������+1
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

    #���ݴ���
    list1  = dataCl1(list1,list2)
    list3 = dataCl3(list3)
    # print(list3)
    # print(len(list3))
    # ����xr����
    xr(list1,list2,list3)

# list1�����ݵĴ���
def dataCl1(list1,list2):
    def take2(elem):
        return elem[1]
    list1.sort(key=take2)
    sum = 0
    for yiwei in range(0, len(list1)):
        # �жϵ�һ��
        if yiwei == 0:
            # �ж�list1�ĵ�һ���Ƿ�곤
            if list1[yiwei][1] == '�곤':
                if list2[5] != '':
                    # list2��Сޱ�겹��Ϊ��,��һ��Сޱ,һ�кϼ�
                    sss = []
                    sss.append(list2[0])
                    sss.append(list2[1])
                    sss.append(list2[4])
                    sss.append('С΢�겹')
                    sss.append(list2[5])
                    list1.insert(1, sss)
                    sss1 = ['���ӻ��ָ����˻����˺ϼ�']
                    nuM = float(list1[0][16]) + float(list2[5])
                    sss1.append(nuM)
                    list1.insert(2, sss1)
                else:
                    # list2��Сޱ�겹Ϊ��,��һ�кϼ�
                    sss1 = ['���ӻ��ָ����˻����˺ϼ�']
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
                    sss1 = ['���ӻ��ָ����˻����˺ϼ�']
                    sss1.append(list2[5])
                    list1.insert(1, sss1)
        else:
            if list1[yiwei][1] == '�ŵ������200Ԫ':
                if yiwei+1 == len(list1):
                    break
                elif list1[yiwei][1] == list1[yiwei + 1][1]:
                    sum += float(list1[yiwei][16])
                    continue
                elif list1[yiwei][1] != list1[yiwei + 1][1]:
                    sum += float(list1[yiwei][16])
                    sss1 = ['���ӻ��ִ����˻����˺ϼ�']
                    sss1.append(sum)
                    list1.insert(yiwei + 1, sss1)
                    break
    for i in range(0,len(list1)):
        list1[i].insert(0, i+1)
    return list1

# list3���ݴ���
def dataCl3(list3):
    if len(list3) > 0:
        list4 = np.array(list3)
        idex = np.lexsort([list4[:, 1], list4[:, 5]])
        list3 = list4[idex, :]
        list3 = list3.tolist()
        sum = 0.0 #��ÿ���˵ĺϼ�
        ss = [] #��ڼ��в���ϼ�
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
            cc = ['�ϼƣ�']
            cc.append(ss[y][1])
            # ��Ϊѭ���˼��ξͲ����˼���,�����仯���Լ���y
            list3.insert(ss[y][0] + y, cc)
        return list3
    else:
        return list3


# �½�ģ�岢д������
def xr(list1, list2,list3):
    global yue,wcnum
    wb = load_workbook('D:\\fuwufei2\\�嵥ģ��.xlsx')
    ws = wb["�嵥"]
    ws1 = wb['�������ϸ']
    a = 'ע������ѷ���˵����\n1�������ȫ��ֱ�ӷ��� + ��������ֱ�ӷ��� + �������񽱣���ϸ������2\n2������ϼ� = �������۽� + ���۷��� + ��� + �쵼�� + ׿Խ��\n3��������ˣ�����������ķ�������ֱ�ӷ��񽱺Ͷ�Ӧ�������۽�����ʵʱ������Ӧ���ӻ�������˻����������½�����\n4���½����� = ����ϼ� - ������� - ���մ��� - �����۳�\n5����ɫ�������룬��ɫ����֧����\n'
    n = 6  #�嵥sheet�����п�ʼд

    #�嵥sheetд������
    for yiwei in range(0, len(list1)):
        if len(list1[yiwei]) >10:
            ss = 1#�嵥sheet�����п�ʼд
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

    #�������ϸsheetд������
    ss1 = 3  # �������ϸsheet�����п�ʼд
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
    wcnum = 0 #��ɵ�����
    row_list = []  # ���ÿ���������
    col_list = []  # ���ÿ���������
    yue = getmonth()
    # sheets_list = []  # ��źϲ����sheet���б�
    sheets_list = ['202012 С΢�겹', '202012 ������嵥ԭʼ����', '202012 ���ϸ��', '202012 ���ۿע', '202012 �ŵ���Ϣ��']
    # zhenghe_path = zhenghe()

    #�����ļ��д�Ÿ�����Ϣ
    dir_path = 'D:\\fuwufei2\\' + getmonth() + '������嵥'
    # os.mkdir(dir_path)

    data = xlrd.open_workbook("D:\\fuwufei2\\xx.xlsx")
    # ��˳��򿪸���sheet����ȡ����
    for num in range(0, len(sheets_list)):
        for sheet_list in sheets_list:
            if ('ԭʼ����' in sheet_list) and num == 0:
                sheet0 = data.sheet_by_name(sheet_list)
                rows0 = sheet0.nrows
                cols0 = sheet0.ncols
                row_list.append(rows0)
                col_list.append(cols0)
                break
            elif ('�ŵ���Ϣ' in sheet_list) and num == 1:
                sheet1 = data.sheet_by_name(sheet_list)
                rows1 = sheet1.nrows
                cols1 = sheet1.ncols
                row_list.append(rows1)
                col_list.append(cols1)
                break
            elif ('С΢�겹' in sheet_list) and num == 2:
                sheet2 = data.sheet_by_name(sheet_list)
                rows2 = sheet2.nrows
                cols2 = sheet2.ncols
                row_list.append(rows2)
                col_list.append(cols2)
                break
            elif ('���ۿע' in sheet_list) and num == 3:
                sheet3 = data.sheet_by_name(sheet_list)
                rows3 = sheet3.nrows
                cols3 = sheet3.ncols
                row_list.append(rows3)
                col_list.append(cols3)
                break
            elif ('���ϸ��' in sheet_list) and num == 4:
                sheet4 = data.sheet_by_name(sheet_list)
                rows4 = sheet4.nrows
                cols4 = sheet4.ncols
                row_list.append(rows4)
                col_list.append(cols4)
                break

    print(row_list)
    for xhdh in range(1, row_list[0] ):
        dh1 = sheet0.cell(xhdh-1, 0).value
        dh2 = sheet0.cell(xhdh, 0).value # ��Ҫ�����嵥�ĵ��
        if dh1 != dh2 :
            xhs(dh2)

