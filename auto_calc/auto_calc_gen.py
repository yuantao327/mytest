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

#a_rand_2 = random.choice((1,2,3,4))
#a_rand_3 = random.choice((1,2,3))
#a_rand_4 = random.choice((1,2))
#a_rand_5 = 1


# AAAxA=BBBB
def calc_0(offset_col,offset_row ):
    mult_b = random.choice((2,3,4))
    if mult_b==2:
        mult_a_1 = random.choice((1,2,3,4))
        mult_a_2 = random.choice((1,2,3,4))
        mult_a_3 = random.randint(0,9)
    elif mult_b==3:
        mult_a_1 = random.choice((1,2,3))
        mult_a_2 = random.choice((1,2,3))
        mult_a_3 = random.randint(0,9)
    elif mult_b==4:
        mult_a_1 = random.choice((1,2))
        mult_a_2 = random.choice((1,2))
        mult_a_3 = random.randint(0,9)
    elif mult_b==5:
        mult_a_1 = 1
        mult_a_2 = 1
        mult_a_3 = random.randint(0,9)
    else:
        mult_a_1 = 1
        mult_a_2 = 1
        mult_a_3 = 1

    mult_a= int(str(mult_a_3)+str(mult_a_2)+str(mult_a_1))
    mult = mult_a*mult_b
    sheet.col(offset_col+0).width = 2000
    sheet.col(offset_col+1).width = 600
    sheet.col(offset_col+2).width = 2000
    sheet.col(offset_col+3).width = 600
    sheet.col(offset_col+4).width = 2000
    sheet.write(1,offset_col+0,mult_a,style)
    sheet.write(1,offset_col+1,'x',style)
    sheet.write(1,offset_col+2,mult_b,style)
    sheet.write(1,offset_col+3,'=',style)
    sheet.write(1,offset_col+4,mult,style)
    print mult_a
    print mult_b
    print mult

    return;


for sheet_cnt in range(1,6):
    sheet_name = 'sheet'+str(sheet_cnt)
    sheet  = workbook.add_sheet(sheet_name,cell_overwrite_ok=True)
    sheet.write_merge(0, 0, 0, 10, 'Maths excersise '+sheet_name,style) # Merges row 0's columns 0 through 3.
    calc_0(0,1)
    calc_0(6,1)
workbook.save('result.xls')
print '创建excel文件完成！'

#print '请关闭excel文件重试！'




