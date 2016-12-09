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
#-----------常量定义-----------------------------------
FORM_WIDTH = 5200
BLANK_HEIGHT = 1200
DATA_HEIGHT = 400

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

alignment = xlwt.Alignment() # Create Alignment
alignment.horz = xlwt.Alignment.HORZ_CENTER # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER,

borders = xlwt.Borders() # Create Borders
borders.left   = xlwt.Borders.MEDIUM # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK,
borders.right  = xlwt.Borders.MEDIUM # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK,
borders.bottom = xlwt.Borders.MEDIUM # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK,
borders.top    = xlwt.Borders.MEDIUM # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK,

style1 = xlwt.XFStyle()
style1.font = font1
style1.alignment = alignment

style2 = xlwt.XFStyle()
style2.font = font2
style2.alignment = alignment

#style.borders = borders

file_h = open('./result.txt', 'w')

file_h.write('三年级数学练习答案'+'\n\n')

#-------------------------------------------------------------
# AAA+BBB=CCCC
def calc_gen_add(offset_col,offset_row ):
    add_a = random.randint(10,499)
    add_b = random.randint(10,499)
    sum_c = add_a + add_b
    sheet.col(offset_col+0).width = FORM_WIDTH
    str_out=str(add_a)+' + '+str(add_b)+' = '
    sheet.write(offset_row,offset_col+0,str_out,style1)
    file_h.write(str(sum_c)+'  ')

    return;

#-------------------------------------------------------------
# CCCC-BBB=AAA
def calc_gen_sub(offset_col,offset_row ):
    add_a = random.randint(10,499)
    add_b = random.randint(10,499)
    sum_c = add_a + add_b
    sheet.col(offset_col+0).width = FORM_WIDTH
    str_out=str(sum_c)+' - '+str(add_b)+' = '
    sheet.write(offset_row,offset_col+0,str_out,style1)
    file_h.write(str(add_a)+'  ')

    return;

#-------------------------------------------------------------
# AAAxB=CCCC
def calc_gen_mult(offset_col,offset_row ):
    mult_b = random.randint(2,9)
    mult_a = random.randint(10,499)
    mult = mult_a*mult_b
    sheet.col(offset_col+0).width = FORM_WIDTH
    exchange = random.randint(0,20)
    if exchange <> 0 :
        str_out=str(mult_a)+u' × '+str(mult_b)+' = '
        sheet.write(offset_row,offset_col+0,str_out,style1)
    else :
        str_out= str(mult_b) + u' × ' + str(mult_a) + ' = '
        sheet.write(offset_row,offset_col+0,str_out,style1)
    file_h.write(str(mult)+'  ')

    return;

#-------------------------------------------------------------
# AAA+BBB+CCC=DDDD
def calc_gen_mix_0(offset_col,offset_row ):
    add_a = random.randint(10,299)
    add_b = random.randint(10,299)
    add_c = random.randint(10,299)
    sum_d = add_a + add_b + add_c
    sheet.col(offset_col+0).width = FORM_WIDTH
    str_out=str(add_a)+'+'+str(add_b)+ '+' + str(add_c) + '='
    sheet.write(offset_row,offset_col+0,str_out,style1)
    file_h.write(str(sum_d)+'  ')

    return;

#-------------------------------------------------------------
# AAA*B + CCC = DDDD or CCC+AAA*B = DDDD
def calc_gen_mix_1(offset_col,offset_row ):
    mult_a = random.randint(10,499)
    mult_b = random.randint(2,9)
    add_c = random.randint(10,499)
    sum_d = mult_a*mult_b+add_c
    sheet.col(offset_col+0).width = FORM_WIDTH
    exchange = random.randint(0,1)
    if exchange <> 0 :
        str_out=str(mult_a)+u'×'+str(mult_b)+ '+' + str(add_c) +'='
        sheet.write(offset_row,offset_col+0,str_out,style1)
    else :
        str_out=str(add_c)+'+'+str(mult_a)+u'×'+str(mult_b)+'='
        sheet.write(offset_row,offset_col+0,str_out,style1)
    file_h.write(str(sum_d)+'  ')

    return;
#-------------------------------------------------------------
# AAA/B + CCC = DDD
def calc_gen_mix_2(offset_col,offset_row ):
    data_a = random.randint(1,9)
    data_b = random.randint(3,9)
    data_t = data_a*data_b
    data_c = random.randint(10,499)
    data_d = data_t/data_a+data_c
    sheet.col(offset_col+0).width = FORM_WIDTH
#    str_out=str(data_d)+'-'+str(data_t)+u'÷'+str(data_a)+'='
#    sheet.write(offset_row,offset_col+0,str_out,style1)
#    file_h.write(str(data_c)+'  ')
    exchange = random.randint(0,1)
    if exchange <> 0 :
        str_out=str(data_c)+'+'+str(data_t)+ u'÷' + str(data_a) +'='
        sheet.write(offset_row,offset_col+0,str_out,style1)
    else :
        str_out=str(data_t)+ u'÷' + str(data_a) + '+'+ str(data_c)+'='
        sheet.write(offset_row,offset_col+0,str_out,style1)
    file_h.write(str(data_d)+'  ')

    return;

#-------------------------------------------------------------
# DDD-AAA-BBB = CCC
def calc_gen_mix_3(offset_col,offset_row ):
    add_a = random.randint(10,299)
    add_b = random.randint(10,299)
    add_c = random.randint(10,299)
    sum_d = add_a + add_b + add_c
    sheet.col(offset_col+0).width = FORM_WIDTH
    str_out=str(sum_d)+'-'+str(add_a)+ '-'+ str(add_b) + '='
    sheet.write(offset_row,offset_col+0,str_out,style1)
    file_h.write(str(add_c)+'  ')

    return;

#-------------------------------------------------------------
# DDD-(AAA+BBB) = CCC
def calc_gen_mix_4(offset_col,offset_row ):
    add_a = random.randint(10,299)
    add_b = random.randint(10,299)
    add_c = random.randint(10,299)
    sum_d = add_a + add_b + add_c
    sheet.col(offset_col+0).width = FORM_WIDTH
    str_out=str(sum_d) + '-' + '(' + str(add_a) + '+' + str(add_b)+')' + '='
    sheet.write(offset_row,offset_col+0,str_out,style1)
    file_h.write(str(add_c)+'  ')

    return;

#-------------------------------------------------------------
# DDD-AAA+BBB = CCC
def calc_gen_mix_5(offset_col,offset_row ):
    data_d = random.randint(10,999)
    data_a = random.randint(10,data_d)
    data_b = random.randint(10,data_a)
    data_c = data_d - data_a + data_b
    sheet.col(offset_col+0).width = FORM_WIDTH
    str_out=str(data_d) + '-' + str(data_a) + '+' + str(data_b) + '='
    sheet.write(offset_row,offset_col+0,str_out,style1)
    file_h.write(str(data_c)+'  ')

    return;

#-------------------------------------------------------------
# DDD-(AAA-BBB) = CCC
def calc_gen_mix_6(offset_col,offset_row ):
    data_d = random.randint(10,999)
    data_a = random.randint(10,data_d)
    data_b = random.randint(10,data_a)
    data_c = data_d - data_a + data_b
    sheet.col(offset_col+0).width = FORM_WIDTH
    str_out=str(data_d) + '-' + '(' +  str(data_a) + '-' + str(data_b) + ')' + '='
    sheet.write(offset_row,offset_col+0,str_out,style1)
    file_h.write(str(data_c)+'  ')

    return;

#-------------------------------------------------------------
# DDDD-AAA*B = CCC
def calc_gen_mix_7(offset_col,offset_row ):
    mult_a = random.randint(10,199)
    mult_b = random.randint(2,9)
    add_c = random.randint(10,499)
    sum_d = mult_a*mult_b+add_c
    sheet.col(offset_col+0).width = FORM_WIDTH
    str_out=str(sum_d)+'-'+str(mult_a)+u'×'+str(mult_b)+'='
    sheet.write(offset_row,offset_col+0,str_out,style1)
    file_h.write(str(add_c)+'  ')

    return;

#-------------------------------------------------------------
# DDDD-AAA/B = CCC
def calc_gen_mix_8(offset_col,offset_row ):
    data_a = random.randint(1,9)
    data_b = random.randint(2,9)
    data_t = data_a*data_b
    data_c = random.randint(10,499)
    data_d = data_t/data_a+data_c
    sheet.col(offset_col+0).width = FORM_WIDTH
    str_out=str(data_d)+'-'+str(data_t)+u'÷'+str(data_a)+'='
    sheet.write(offset_row,offset_col+0,str_out,style1)
    file_h.write(str(data_c)+'  ')

    return;

#-------------------------------------------------------------
#AA/B*CCC = DDD
def calc_gen_mix_9(offset_col,offset_row ):
    data_a = random.randint(1,9)
    data_b = random.randint(3,9)
    data_t = data_a*data_b
    data_c = random.randint(10,499)
    data_d = data_t/data_a*data_c
    sheet.col(offset_col+0).width = FORM_WIDTH
    str_out=str(data_t)+u'÷'+str(data_a)+u'×'+str(data_c)+'='
    sheet.write(offset_row,offset_col+0,str_out,style1)
    file_h.write(str(data_d)+'  ')
    return;

#-------------------------------------------------------------
#CCC*AA/B = DDD
def calc_gen_mix_10(offset_col,offset_row ):
    data_a = random.randint(1,9)
    data_b = random.randint(3,9)
    data_t = data_a*data_b
    data_c = random.randint(10,499)
    data_d = data_t/data_a*data_c
    sheet.col(offset_col+0).width = FORM_WIDTH
    str_out=str(data_c)+u'×'+str(data_t)+u'÷'+str(data_a)+'='
    sheet.write(offset_row,offset_col+0,str_out,style1)
    file_h.write(str(data_d)+'  ')
    return;

#-------------------------------------------------------------
#-------------------------------------------------------------

def blank(offset_col,offset_row ):
    sheet.col(offset_col+0).width = FORM_WIDTH
    sheet.write(offset_row,offset_col,'',style2)
    return;


for sheet_cnt in range(1,8):
    sheet_name = 'sheet'+str(sheet_cnt)
    sheet_name_out = '第'+str(sheet_cnt)+'页'
    sheet_name_print = u'三年级数学练习    姓名:____________ '
    sheet  = workbook.add_sheet(sheet_name,cell_overwrite_ok=True)
    sheet.header_str = ''
    sheet.footer_str = u'第'+str(sheet_cnt)+u'页'
    sheet.write_merge(0, 0, 0, 4,sheet_name_print,style1) # Merges row 0's columns 0 through 10.
    file_h.write('----------------'+ sheet_name_out + '--------------------\n\n')
    for i in range(1,8):
        calc_gen_mult(0,i*2-1)
        blank(0,i*2)
    for i in range(1,8):
        calc_gen_mult(2,i*2-1)
        blank(2,i*2)
    for i in range(1,8):
        select_mix = random.randint(0,10)
        if select_mix == 0 :
            calc_gen_mix_0(4,i*2-1)
        elif select_mix == 1 :
            calc_gen_mix_1(4,i*2-1)
        elif select_mix == 2 :
            calc_gen_mix_2(4,i*2-1)
        elif select_mix == 3 :
            calc_gen_mix_3(4,i*2-1)
        elif select_mix == 4 :
            calc_gen_mix_4(4,i*2-1)
        elif select_mix == 5 :
            calc_gen_mix_5(4,i*2-1)
        elif select_mix == 6 :
            calc_gen_mix_6(4,i*2-1)
        elif select_mix == 7 :
            calc_gen_mix_7(4,i*2-1)
        elif select_mix == 8 :
            calc_gen_mix_8(4,i*2-1)
        elif select_mix == 9 :
            calc_gen_mix_9(4,i*2-1)
        elif select_mix == 10 :
            calc_gen_mix_10(4,i*2-1)
        else :
            calc_gen_mix_10(4,i*2-1)
        blank(4,i*2)

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

    sheet.write_merge(15, 15, 0, 4, u'签字__________  日期____________',style1)

workbook.save('result.xls')
print '创建excel文件完成！'

file_h.close()

#test
#for i in range(1,10):
#    aaa = random.randint(0,9)
#    print aaa
#print '请关闭excel文件重试！'

