#-*- coding: gbk -*-
import xlrd

data = xlrd.open_workbook('sf42_example.xlsx') # ��xls�ļ�
table = data.sheets()[0] # �򿪵�һ�ű�
nrows = table.nrows # ��ȡ�������

for i in range(nrows): # ѭ�����д�ӡ
    if i == 0: # ������һ��
        continue
    print table.row_values(i)[:13] # ȡǰʮ����