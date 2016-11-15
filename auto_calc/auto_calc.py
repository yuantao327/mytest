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
font.height = 250

alignment = xlwt.Alignment() # Create Alignment
alignment.horz = xlwt.Alignment.HORZ_CENTER # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER,

borders = xlwt.Borders() # Create Borders
borders.left   = xlwt.Borders.MEDIUM # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK,
borders.right  = xlwt.Borders.MEDIUM # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK,
borders.bottom = xlwt.Borders.MEDIUM # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK,
borders.top    = xlwt.Borders.MEDIUM # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK,

style = xlwt.XFStyle()
style.font = font
style.alignment = alignment
#style.borders = borders


def calc_block_gen(offset_col,offset_row ):
    for num in range(offset_row+0,offset_row+20):
        sign=random.choice([-1,1])

        if sign == 1 :
            sign_show = '+'
            a=random.randint(50,2001)
            b=random.randint(50,2001)
            c=a+b
            while c>2000 :
               a=random.randint(50,2001)
               b=random.randint(50,2001)
               c=a+b
        else:
            sign_show = '-'
            a=random.randint(50,2001)
            b=random.randint(50,a)
            c=a-b
        sheet.col(offset_col+0).width = 2000
        sheet.col(offset_col+1).width = 600
        sheet.col(offset_col+2).width = 2000
        sheet.col(offset_col+3).width = 600
        sheet.col(offset_col+4).width = 2000
        sheet.write(num,offset_col+0,a,style)
        sheet.write(num,offset_col+1,sign_show,style)
        sheet.write(num,offset_col+2,b,style)
        sheet.write(num,offset_col+3,'=',style)
        sheet.write(num,offset_col+4,c,style)

    return;

for sheet_cnt in range(1,31):

    sheet_name = 'sheet'+str(sheet_cnt)
    sheet  = workbook.add_sheet(sheet_name,cell_overwrite_ok=True)
    sheet.write_merge(0, 0, 0, 10, 'Maths excersise '+sheet_name,style) # Merges row 0's columns 0 through 3.
    calc_block_gen(0,1)
    calc_block_gen(6,1)

    calc_block_gen(0,23)
    calc_block_gen(6,23)

#保存该excel文件,有同名文件时直接覆盖
workbook.save('E:\\work\\python\\auto_calc\\result.xls')
print '创建excel文件完成！'
