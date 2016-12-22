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
DATA_WIDTH = 1500
SIGN_WIDTH = 1000
BLANK_HEIGHT = 500
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
# AA/CC + BB/CC= DD/CC
def calc_gen_add(offset_row,offset_col):
    data_c = random.randint(11,99)
    data_a = random.randint(11,data_c)
    data_tmp = data_c - data_a
    data_b = random.randint(0,data_tmp)
    data_d = data_a + data_b
    if data_d == data_c:
        data_out = '1'
    else:
    	  data_out = str(data_d)+'/'+str(data_c)
    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(data_a),style3)
    sheet.write(offset_row+1,offset_col,str(data_c),style1)

    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+1, offset_col+1,"+",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_b),style3)
    sheet.write(offset_row+1,offset_col+2,str(data_c),style1)

    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+3, offset_col+3,"=",style1) # Merges row 0's columns 0 through 10.

    file_h.write(str(data_out)+'  ')

    return;

# AA/CC - BB/CC= DD/CC
def calc_gen_sub(offset_row,offset_col):
    data_c = random.randint(11,99)
    data_a = random.randint(11,data_c)
    data_tmp = data_c - data_a
    data_b = random.randint(0,data_tmp)
    data_d = data_a + data_b
    if data_b == 0:
        data_out = '0'
    else:
    	  data_out = str(data_b)+'/'+str(data_c)
    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(data_d),style3)
    sheet.write(offset_row+1,offset_col,str(data_c),style1)

    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+1, offset_col+1,"-",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_a),style3)
    sheet.write(offset_row+1,offset_col+2,str(data_c),style1)

    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+3, offset_col+3,"=",style1) # Merges row 0's columns 0 through 10.

    file_h.write(str(data_out)+'  ')

    return;


#-------------------------------------------------------------
# AA/ZZ + BB/ZZ (?) CC/ZZ + DD/ZZ
def calc_gen_comp1(offset_row,offset_col):
    data_z = random.randint(11,99)
    sum_ab = random.randint(0,data_z)
    data_a = random.randint(0,sum_ab)
    data_b = sum_ab - data_a

    rand_tmp = random.randint(-3,3)
    sum_cd = abs(sum_ab + rand_tmp)

    data_c = random.randint(0,sum_cd)
    data_d = sum_cd - data_c

    if sum_ab == sum_cd:
    	  data_out = "="
    elif sum_ab> sum_cd:
    	  data_out = ">"
    else:
    	  data_out = "<"

    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(data_a),style3)
    sheet.write(offset_row+1,offset_col,str(data_z),style1)

    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+1, offset_col+1,"+",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_b),style3)
    sheet.write(offset_row+1,offset_col+2,str(data_z),style1)

    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+3, offset_col+3,u"○",style4) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(data_c),style3)
    sheet.write(offset_row+1,offset_col+4,str(data_z),style1)

    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+5, offset_col+5,"+",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+6).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+6,str(data_d),style3)
    sheet.write(offset_row+1,offset_col+6,str(data_z),style1)

    file_h.write(data_out+'  ')

    return;


#-------------------------------------------------------------
# SUM/ZZ - AA/ZZ (?) SUM/ZZ - CC/ZZ
def calc_gen_comp2(offset_row,offset_col):
    data_z = random.randint(11,99)
    sum_ab = random.randint(0,data_z)
    data_a = random.randint(0,sum_ab)
    data_b = sum_ab - data_a

    rand_tmp = random.randint(-3,3)
    data_d = abs(data_b + rand_tmp)

    data_c = random.randint(0,sum_ab)
    sum_cd = data_c + data_d

    if data_b == data_d:
    	  data_out = "="
    elif data_b> data_d:
    	  data_out = ">"
    else:
    	  data_out = "<"

    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(sum_ab),style3)
    sheet.write(offset_row+1,offset_col,str(data_z),style1)

    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+1, offset_col+1,"-",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_a),style3)
    sheet.write(offset_row+1,offset_col+2,str(data_z),style1)

    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+3, offset_col+3,u"○",style4) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(sum_cd),style3)
    sheet.write(offset_row+1,offset_col+4,str(data_z),style1)

    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+5, offset_col+5,"-",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+6).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+6,str(data_c),style3)
    sheet.write(offset_row+1,offset_col+6,str(data_z),style1)

    file_h.write(data_out+'  ')

    return;

#-------------------------------------------------------------
# SUM/ZZ - AA/ZZ (?) CC/ZZ + DD/ZZ
def calc_gen_comp3(offset_row,offset_col):
    data_z = random.randint(11,99)
    sum_ab = random.randint(0,data_z)
    data_a = random.randint(0,sum_ab)
    data_b = sum_ab - data_a

    rand_tmp = random.randint(-3,3)
    sum_cd = abs(data_b + rand_tmp)

    data_c = random.randint(0,sum_cd)
    data_d = sum_cd - data_c

    if data_b == sum_cd:
    	  data_out = "="
    elif data_b> sum_cd:
    	  data_out = ">"
    else:
    	  data_out = "<"

    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(sum_ab),style3)
    sheet.write(offset_row+1,offset_col,str(data_z),style1)

    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+1, offset_col+1,"-",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_a),style3)
    sheet.write(offset_row+1,offset_col+2,str(data_z),style1)

    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+3, offset_col+3,u"○",style4) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(data_c),style3)
    sheet.write(offset_row+1,offset_col+4,str(data_z),style1)

    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+5, offset_col+5,"+",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+6).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+6,str(data_d),style3)
    sheet.write(offset_row+1,offset_col+6,str(data_z),style1)

    file_h.write(data_out+'  ')

    return;

#-------------------------------------------------------------
#CC/ZZ + DD/ZZ (?) SUM/ZZ - AA/ZZ
def calc_gen_comp4(offset_row,offset_col):
    data_z = random.randint(11,99)
    sum_ab = random.randint(0,data_z)
    data_a = random.randint(0,sum_ab)
    data_b = sum_ab - data_a

    rand_tmp = random.randint(-3,3)
    sum_cd = abs(data_b + rand_tmp)

    data_c = random.randint(0,sum_cd)
    data_d = sum_cd - data_c

    if data_b == sum_cd:
    	  data_out = "="
    elif data_b> sum_cd:
    	  data_out = "<"
    else:
    	  data_out = ">"

    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(data_c),style3)
    sheet.write(offset_row+1,offset_col,str(data_z),style1)

    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+1, offset_col+1,"+",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_d),style3)
    sheet.write(offset_row+1,offset_col+2,str(data_z),style1)

    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+3, offset_col+3,u"○",style4) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(sum_ab),style3)
    sheet.write(offset_row+1,offset_col+4,str(data_z),style1)

    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+5, offset_col+5,"-",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+6).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+6,str(data_a),style3)
    sheet.write(offset_row+1,offset_col+6,str(data_z),style1)

    file_h.write(data_out+'  ')

    return;

#-------------------------------------------------------------
# AA/ZZ + BB/ZZ (?) CC/ZZ + DD/ZZ
def calc_gen_comp5(offset_row,offset_col):
    data_z = random.randint(11,99)
    sum_ab = random.randint(0,data_z)
    data_a = random.randint(0,sum_ab)
    data_b = sum_ab - data_a

    sum_cd = sum_ab

    data_c = random.randint(0,sum_cd)
    data_d = sum_cd - data_c

    list=[-3,-2,-1,1,2,3]
    rand_tmp = random.choice(list)
    data_z_tmp = data_z + rand_tmp;

    if sum_ab == 0:
    	  data_out = "="
    elif data_z > data_z_tmp:
    	  data_out = "<"
    else:
    	  data_out = ">"

    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(data_a),style3)
    sheet.write(offset_row+1,offset_col,str(data_z),style1)

    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+1, offset_col+1,"+",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_b),style3)
    sheet.write(offset_row+1,offset_col+2,str(data_z),style1)

    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+3, offset_col+3,u"○",style4) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(data_c),style3)
    sheet.write(offset_row+1,offset_col+4,str(data_z_tmp),style1)

    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+5, offset_col+5,"+",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+6).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+6,str(data_d),style3)
    sheet.write(offset_row+1,offset_col+6,str(data_z_tmp),style1)

    file_h.write(data_out+'  ')

    return;

#-------------------------------------------------------------
# SUM/ZZ - AA/ZZ (?) SUM/ZZ - CC/ZZ
def calc_gen_comp6(offset_row,offset_col):
    data_z = random.randint(11,99)
    sum_ab = random.randint(0,data_z)
    data_a = random.randint(0,sum_ab)
    data_b = sum_ab - data_a
    data_d = data_b

    data_c = random.randint(0,sum_ab)
    sum_cd = data_c + data_d

    list=[-3,-2,-1,1,2,3]
    rand_tmp = random.choice(list)
    data_z_tmp = data_z + rand_tmp;

    if data_b == 0:
    	  data_out = "="
    elif data_z > data_z_tmp:
    	  data_out = "<"
    else:
    	  data_out = ">"

    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(sum_ab),style3)
    sheet.write(offset_row+1,offset_col,str(data_z),style1)

    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+1, offset_col+1,"-",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_a),style3)
    sheet.write(offset_row+1,offset_col+2,str(data_z),style1)

    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+3, offset_col+3,u"○",style4) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(sum_cd),style3)
    sheet.write(offset_row+1,offset_col+4,str(data_z_tmp),style1)

    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+5, offset_col+5,"-",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+6).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+6,str(data_c),style3)
    sheet.write(offset_row+1,offset_col+6,str(data_z_tmp),style1)

    file_h.write(data_out+'  ')

    return;

#-------------------------------------------------------------
# SUM/ZZ - AA/ZZ (?) CC/ZZ + DD/ZZ
def calc_gen_comp7(offset_row,offset_col):
    data_z = random.randint(11,99)
    sum_ab = random.randint(0,data_z)
    data_a = random.randint(0,sum_ab)
    data_b = sum_ab - data_a

    sum_cd = data_b

    data_c = random.randint(0,sum_cd)
    data_d = sum_cd - data_c

    list=[-3,-2,-1,1,2,3]
    rand_tmp=random.choice(list)
    data_z_tmp = data_z + rand_tmp;

    if data_b == 0:
    	  data_out = "="
    elif data_z > data_z_tmp:
    	  data_out = "<"
    else:
    	  data_out = ">"

    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(sum_ab),style3)
    sheet.write(offset_row+1,offset_col,str(data_z),style1)

    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+1, offset_col+1,"-",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_a),style3)
    sheet.write(offset_row+1,offset_col+2,str(data_z),style1)

    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+3, offset_col+3,u"○",style4) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(data_c),style3)
    sheet.write(offset_row+1,offset_col+4,str(data_z_tmp),style1)

    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+5, offset_col+5,"+",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+6).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+6,str(data_d),style3)
    sheet.write(offset_row+1,offset_col+6,str(data_z_tmp),style1)

    file_h.write(data_out+'  ')

    return;


#-------------------------------------------------------------
#CC/ZZ + DD/ZZ (?) SUM/ZZ - AA/ZZ
def calc_gen_comp8(offset_row,offset_col):
    data_z = random.randint(11,99)
    sum_ab = random.randint(0,data_z)
    data_a = random.randint(0,sum_ab)
    data_b = sum_ab - data_a

    sum_cd = data_b

    data_c = random.randint(0,sum_cd)
    data_d = sum_cd - data_c

    list=[-3,-2,-1,1,2,3]
    rand_tmp = random.choice(list)
    data_z_tmp = data_z + rand_tmp;

    if data_b == 0:
    	  data_out = "="
    elif data_z > data_z_tmp:
    	  data_out = ">"
    else:
    	  data_out = "<"


    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(data_c),style3)
    sheet.write(offset_row+1,offset_col,str(data_z_tmp),style1)

    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+1, offset_col+1,"+",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_d),style3)
    sheet.write(offset_row+1,offset_col+2,str(data_z_tmp),style1)

    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+3, offset_col+3,u"○",style4) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(sum_ab),style3)
    sheet.write(offset_row+1,offset_col+4,str(data_z),style1)

    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+5, offset_col+5,"-",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+6).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+6,str(data_a),style3)
    sheet.write(offset_row+1,offset_col+6,str(data_z),style1)

    file_h.write(data_out+'  ')

    return;


#-------------------------------------------------------------
#-------------------------------------------------------------

def blank(offset_row,offset_col):
    sheet.write(offset_row,offset_col,'',style2)
    return;

##--------------------------------------------------
## select the calc_compare function

def calc_mixed_1(offset_row,offset_col):
    select_mix = random.randint(1,4)
    if select_mix == 1 :
        calc_gen_comp1(offset_row,offset_col)
    elif select_mix == 2 :
        calc_gen_comp2(offset_row,offset_col)
    elif select_mix == 3 :
        calc_gen_comp3(offset_row,offset_col)
    elif select_mix == 4 :
        calc_gen_comp4(offset_row,offset_col)
    else:
        calc_gen_comp4(offset_row,offset_col)

    return;

def calc_mixed_2(offset_row,offset_col):
    select_mix = random.randint(1,4)
    if select_mix == 1 :
        calc_gen_comp5(offset_row,offset_col)
    elif select_mix == 2 :
        calc_gen_comp6(offset_row,offset_col)
    elif select_mix == 3 :
        calc_gen_comp7(offset_row,offset_col)
    elif select_mix == 4 :
        calc_gen_comp8(offset_row,offset_col)
    else:
        calc_gen_comp8(offset_row,offset_col)

    return;

def calc_comp_mixed(offset_row,offset_col):
    select_mix = random.randint(1,5)
    if select_mix == 1 :
        calc_mixed_2(offset_row,offset_col)
    else:
        calc_mixed_1(offset_row,offset_col)

    return;

def calc_div_mixed(offset_row,offset_col):
    select_mix = random.randint(0,1)
    if select_mix == 0 :
        calc_gen_add(offset_row,offset_col)
    else:
        calc_gen_sub(offset_row,offset_col)

    return;



for sheet_cnt in range(1,8):
    sheet_name = 'sheet'+str(sheet_cnt)
    sheet_name_out = '第'+str(sheet_cnt)+'页'
    sheet_name_print = u'三年级数学练习         姓名:_________ '
    sheet  = workbook.add_sheet(sheet_name,cell_overwrite_ok=True)
    sheet.header_str = ''
    sheet.footer_str = u'第'+str(sheet_cnt)+u'页'
    sheet.write_merge(0, 0, 0, 14,sheet_name_print,style1) # Merges row 0's columns 0 through 10.
    file_h.write('----------------'+ sheet_name_out + '--------------------\n\n')
    for i in range(1,5):
        calc_comp_mixed(i*3-1,0)
        blank(i*3+1,0)
    file_h.write('\n')

    for i in range(5,11):
        calc_div_mixed(i*3-1,0)
        blank(i*3+1,0)
    file_h.write('\n')

    for i in range(1,11):
        calc_div_mixed(i*3-1,8)
        blank(i*3+1,8)
    file_h.write('\n')

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

    sheet.write_merge(34, 34, 0, 14, u'签字__________  日期____________',style1)

workbook.save('result.xls')
print '创建excel文件完成！'

file_h.close()

#test
#for i in range(1,10):
#    aaa = random.randint(0,9)
#    print aaa
#print '请关闭excel文件重试！'

#list=[1,2,3,4,5]
#for i in range(0,5):
#    test = random.sample(list,1)
#    print test
#print list

