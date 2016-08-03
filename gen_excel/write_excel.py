#-*- coding: gbk -*-

import xlwt
#创建workbook和sheet对象
#workbook = xlwt.Workbook() #注意Workbook的开头W要大写
for i in range(1000):
    workbook = xlwt.Workbook(encoding = 'gbk')
    sheet1 = workbook.add_sheet('财税网点机构参数维护申请表',cell_overwrite_ok=True)
    #向sheet页中写入数据
    sheet1.write(0,0,'''申请表''')
    sheet1.write(1,0,'wangdian')
    sheet1.write(7,0,'数据开始')
    for j in range(500):
        sheet1.write(j+8,0,str(j+1))
        sheet1.write(j+8,2,'test')
        sheet1.write(j+8,3,str(163000+i+1)+str(100+j+1)+'000')#12位
        sheet1.write(j+8,4,str(420170000+(j+1)%119))#420170001
        sheet1.write(j+8,5,'1')
    sheet1.write(1008,0,'数据结束')
    workbook.save('sf42_result\\500_'+str(i+1)+'.xls')
"""
#-----------使用样式-----------------------------------
#初始化样式
style = xlwt.XFStyle()
#为样式创建字体
font = xlwt.Font()
font.name = 'Times New Roman'
font.bold = True
#设置样式的字体
style.font = font
#使用样式
sheet.write(0,1,'some bold Times text',style)
"""
#保存该excel文件,有同名文件时直接覆盖

#print '创建excel文件完成！'