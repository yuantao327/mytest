#coding=utf-8
#######################################################
#filename:test_xlwt.py
#author:defias
#date:xxxx-xx-xx
#function：新建excel文件并写入数据
#######################################################
import xlwt
import random
#创建workbook和sheet对象
workbook = xlwt.Workbook() #注意Workbook的开头W要大写

#-----------使用样式-----------------------------------
#初始化样式
#为样式创建字体
font = xlwt.Font()
font.name =  'TimesNew Roman'
font.bold = True
font.colour_index = 0x40
font.height = 800
font.size = 30

style = xlwt.XFStyle()
style.font = font

for sheet_cnt in range(1,31):

    sheet_name = 'sheet'+str(sheet_cnt)
    sheet  = workbook.add_sheet(sheet_name,cell_overwrite_ok=True)

    for num in range(0,20):
        sign=random.choice([-1,1])

        if sign == 1 :
            sign_show = '+'
            a=random.randint(10,2001)
            b=random.randint(10,2001)
            c=a+b
            while c>2000 :
               a=random.randint(10,2001)
               b=random.randint(10,2001)
               c=a+b
        else:
            sign_show = '-'
            a=random.randint(10,2001)
            b=random.randint(10,a)
            c=a-b

        sheet.write(num,0,a,style)
        sheet.write(num,1,sign_show,style)
        sheet.write(num,2,b,style)
        sheet.write(num,3,'=',style)
        sheet.write(num,4,c,style)


#保存该excel文件,有同名文件时直接覆盖
workbook.save('D:\\github\\mytest\\auto_calc\\result.xls')
print '创建excel文件完成！'
