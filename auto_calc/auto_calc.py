#coding=utf-8
#######################################################
#filename:test_xlwt.py
#author:defias
#date:xxxx-xx-xx
#function���½�excel�ļ���д������
#######################################################
import xlwt
import random
#����workbook��sheet����
workbook = xlwt.Workbook() #ע��Workbook�Ŀ�ͷWҪ��д

#-----------ʹ����ʽ-----------------------------------
#��ʼ����ʽ
#Ϊ��ʽ��������
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


#�����excel�ļ�,��ͬ���ļ�ʱֱ�Ӹ���
workbook.save('D:\\github\\mytest\\auto_calc\\result.xls')
print '����excel�ļ���ɣ�'
