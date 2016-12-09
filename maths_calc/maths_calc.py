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
font.height = 400

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

file_h = open('./result.txt', 'w')

file_h.write('三年级数学练习答案'+'\n\n')


# AAAxA=BBBB
def calc_gen(offset_col,offset_row ):
    mult_b = random.randint(2,9)
    mult_a = random.randint(10,499)
    mult = mult_a*mult_b
    sheet.col(offset_col+0).width = 2000
    sheet.col(offset_col+1).width = 600
    sheet.col(offset_col+2).width = 2000
    sheet.col(offset_col+3).width = 600
    sheet.col(offset_col+4).width = 2500
    sheet.col(offset_col+5).width = 500
    exchange = random.randint(0,20)
    if exchange <> 0 :
        sheet.write(offset_row,offset_col+0,mult_a,style)
        sheet.write(offset_row,offset_col+1,'x',style)
        sheet.write(offset_row,offset_col+2,mult_b,style)
        sheet.write(offset_row,offset_col+3,'=',style)
        #sheet.write(offset_row,offset_col+4,mult,style)
    else :
        sheet.write(offset_row,offset_col+0,mult_b,style)
        sheet.write(offset_row,offset_col+1,'x',style)
        sheet.write(offset_row,offset_col+2,mult_a,style)
        sheet.write(offset_row,offset_col+3,'=',style)
        #sheet.write(offset_row,offset_col+4,mult,style)
    file_h.write(str(mult)+'  ')

#    print mult_a
#    print mult_b
#    print mult

    return;

def blank(offset_col,offset_row ):
    sheet.col(offset_col+0).width = 2000
    sheet.col(offset_col+1).width = 600
    sheet.col(offset_col+2).width = 2000
    sheet.col(offset_col+3).width = 600
    sheet.col(offset_col+4).width = 2500
    sheet.col(offset_col+5).width = 500
    sheet.write(offset_row,offset_col+0,'',style)
    sheet.write(offset_row,offset_col+1,'',style)
    sheet.write(offset_row,offset_col+2,'',style)
    sheet.write(offset_row,offset_col+3,'',style)
    return;


for sheet_cnt in range(1,8):
    sheet_name = 'sheet'+str(sheet_cnt)
    sheet_name_out = '第'+str(sheet_cnt)+'页'
    sheet_name_print = u'三年级数学练习    姓名:____________ '
    sheet  = workbook.add_sheet(sheet_name,cell_overwrite_ok=True)
    sheet.write_merge(0, 0, 0, 16,sheet_name_print,style) # Merges row 0's columns 0 through 10.
    file_h.write('----------------'+ sheet_name_out + '--------------------\n\n')
    for i in range(1,8):
        calc_gen(0,i*4-2)
        blank(0,i*4-1)
        blank(0,i*4+1)
        blank(0,i*4)
    for i in range(1,8):
        calc_gen(6,i*4-2)
        blank(6,i*4-1)
        blank(6,i*4)
        blank(6,i*4+1)
    for i in range(1,8):
        calc_gen(12,i*4-2)
        blank(12,i*4-1)
        blank(12,i*4)
        blank(12,i*4+1)

    file_h.write(' \n\n ')

    sheet.write_merge(30, 30, 0, 16, u'等级__________ 签字__________  日期____________',style)

#    for i in range(23,44):
#        calc_gen(0,i)
#    for i in range(23,44):
#        calc_gen(6,i)
#    for i in range(23,44):
#        calc_gen(12,i)

workbook.save('result.xls')
print '创建excel文件完成！'

file_h.close()

#print '请关闭excel文件重试！'




