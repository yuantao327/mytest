#coding=utf-8
#######################################################
#filename:test_xlwt.py
#author:defias
#date:xxxx-xx-xx
#function：新建excel文件并写入数据
#######################################################
import xlwt
import random
import sub_add_sub
import sub_mult
import sub_div
import sub_mix
import sub_comp

#创建workbook和sheet对象
#-----------常量定义-----------------------------------
DATA_WIDTH = 1500
SIGN_WIDTH = 1000
BLANK_HEIGHT = 1050
DATA_HEIGHT = 300
CIR_HEIGHT  = 800

workbook = xlwt.Workbook() #注意Workbook的开头W要大写

#-----------使用样式-----------------------------------
#初始化样式
#为样式创建字体

font1 = xlwt.Font()
font1.name =  'TimesNew Roman'
font1.bold = True
font1.colour_index = 0x40
font1.height = DATA_HEIGHT

font2 = xlwt.Font()
font2.name =  'TimesNew Roman'
font2.bold = True
font2.colour_index = 0x40
font2.height = BLANK_HEIGHT

font3 = xlwt.Font()
font3.name =  'TimesNew Roman'
font3.bold = True
font3.colour_index = 0x40
font3.height = DATA_HEIGHT

font4 = xlwt.Font()
font4.name =  'TimesNew Roman'
font4.bold = False
font4.colour_index = 0x40
font4.height = CIR_HEIGHT

alignment = xlwt.Alignment() # Create Alignment
alignment.horz = xlwt.Alignment.HORZ_CENTER # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER,
alignment.vert = xlwt.Alignment.VERT_CENTER # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER,

borders = xlwt.Borders() # Create Borders
borders.left   = xlwt.Borders.NO_LINE # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK,
borders.right  = xlwt.Borders.NO_LINE # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK,
borders.bottom = xlwt.Borders.MEDIUM # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK,
borders.top    = xlwt.Borders.NO_LINE # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK,

style1 = xlwt.XFStyle()
style1.font = font1
style1.alignment = alignment

style2 = xlwt.XFStyle()
style2.font = font2
style2.alignment = alignment

style3 = xlwt.XFStyle()
style3.font = font3
style3.alignment = alignment
style3.borders = borders

style4 = xlwt.XFStyle()
style4.font = font4
style4.alignment = alignment

file_h = open('./result.txt', 'w')

file_h.write('三年级数学练习答案'+'\n\n')


#-------------------------------------------------------------
#-------------------------------------------------------------

def blank(offset_row,offset_col):
    sheet.col(offset_col+0).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,'',style2)
    return;


for sheet_cnt in range(1,8):
    sheet_name = 'sheet'+str(sheet_cnt)
    sheet_name_out = '第'+str(sheet_cnt)+'页'
    sheet_name_print = u'三年级数学练习    姓名:____________ '
    sheet  = workbook.add_sheet(sheet_name,cell_overwrite_ok=True)
    sheet.header_str = ''
    sheet.footer_str = u'第'+str(sheet_cnt)+u'页'
    sheet.write_merge(0, 0, 0, 15,sheet_name_print,style1) # Merges row 0's columns 0 through 10.
    file_h.write('----------------'+ sheet_name_out + '--------------------\n\n')
    for i in range(1,8):
        sub_add_sub.calc_add_sub_s_sel(i*3-1,0,style1,sheet,file_h)
        blank(i*3+1,0)
    for i in range(1,8):
        sub_div.calc_div_sel(i*3-1,5,style1,style3,sheet,file_h)
        blank(i*3+1,5)
    for i in range(1,8):
        sub_comp.calc_comp_sel(i*3-1,10,style1,style3,style4,sheet,file_h)
        blank(i*3+1,10)
#    for i in range(1,8):
#        calc_gen_mult_s(i*3-1,10)
#        blank(i*3+1,10)

#    calc_gen_mix_0(4,1)
#    blank(4,2)
#    calc_gen_mix_1(4,3)
#    blank(4,4)
#    calc_gen_mix_2(4,5)
#    blank(4,6)
#    calc_gen_mix_3(4,7)
#    blank(4,8)
#    calc_gen_mix_4(4,9)
#    blank(4,10)
#    calc_gen_mix_10(4,11)
#    blank(4,12)
#    calc_gen_mix_10(4,13)
#    blank(4,14)

    file_h.write(' \n\n ')
    sheet.write_merge(23, 23, 0, 15, u'签字__________  日期____________',style1)

workbook.save('result.xls')
print '\n----------Excel file have been created successfully!!!----------'

file_h.close()

##test
#for i in range(1,20):
#    tmp = float(i)/10
#    aaa = int(round(float(i)/10)*10)
#    print str(i)+"   "+str(tmp)+"   "+str(aaa)
##print '请关闭excel文件重试！'

