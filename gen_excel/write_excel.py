#-*- coding: gbk -*-

import xlwt
#����workbook��sheet����
#workbook = xlwt.Workbook() #ע��Workbook�Ŀ�ͷWҪ��д
for i in range(1000):
    workbook = xlwt.Workbook(encoding = 'gbk')
    sheet1 = workbook.add_sheet('��˰�����������ά�������',cell_overwrite_ok=True)
    #��sheetҳ��д������
    sheet1.write(0,0,'''�����''')
    sheet1.write(1,0,'wangdian')
    sheet1.write(7,0,'���ݿ�ʼ')
    for j in range(500):
        sheet1.write(j+8,0,str(j+1))
        sheet1.write(j+8,2,'test')
        sheet1.write(j+8,3,str(163000+i+1)+str(100+j+1)+'000')#12λ
        sheet1.write(j+8,4,str(420170000+(j+1)%119))#420170001
        sheet1.write(j+8,5,'1')
    sheet1.write(1008,0,'���ݽ���')
    workbook.save('sf42_result\\500_'+str(i+1)+'.xls')
"""
#-----------ʹ����ʽ-----------------------------------
#��ʼ����ʽ
style = xlwt.XFStyle()
#Ϊ��ʽ��������
font = xlwt.Font()
font.name = 'Times New Roman'
font.bold = True
#������ʽ������
style.font = font
#ʹ����ʽ
sheet.write(0,1,'some bold Times text',style)
"""
#�����excel�ļ�,��ͬ���ļ�ʱֱ�Ӹ���

#print '����excel�ļ���ɣ�'