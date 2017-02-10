#coding=utf-8
#######################################################
#-------------------------------------------------------------
# AAA+BBB=CCCC

DATA_WIDTH = 1500
SIGN_WIDTH = 1000
BLANK_HEIGHT = 1050
DATA_HEIGHT = 300
CIR_HEIGHT  = 800

import random

def calc_gen_add(offset_row,offset_col,style,sheet,file_h):
    add_a = random.randint(100,2000)
    add_b = random.randint(100,2000)
    sum_c = add_a + add_b

    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(add_a),style)
    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+1,"+",style)
    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(add_b),style)
    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+3,"=",style)

    file_h.write(str(sum_c)+'  ')

    return;

#-------------------------------------------------------------
# CCCC-BBB=AAA
def calc_gen_sub(offset_row,offset_col,style,sheet,file_h):
    add_a = random.randint(100,1000)
    add_b = random.randint(100,1000)
    sum_c = add_a + add_b
#    sheet.col(offset_col+0).width = DATA_WIDTH
#    str_out=str(sum_c)+' - '+str(add_b)+' = '
#    sheet.write(offset_row,offset_col+0,str_out,style1)
    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(sum_c),style)
    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+1,"-",style)
    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(add_b),style)
    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+3,"=",style)

    file_h.write(str(add_a)+'  ')

    return;

#-------------------------------------------------------------
# AAA+BBB=CCCC
def calc_gen_add_s(offset_row,offset_col,style,sheet,file_h):
    add_a = random.randint(50,500)
    add_b = random.randint(50,500)
    sum_c = add_a + add_b
    sum_c_s = int(round(float(add_a)/10)*10)+int(round(float(add_b)/10)*10)

    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(add_a),style)
    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+1,"+",style)
    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(add_b),style)
    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+3,u'≈',style)

    file_h.write(str(sum_c_s)+'  ')

    return;

#-------------------------------------------------------------
# CCCC-BBB=AAA
def calc_gen_sub_s(offset_row,offset_col,style,sheet,file_h):
    add_a = random.randint(50,500)
    add_b = random.randint(50,500)
    sum_c = add_a + add_b
    add_a_s = int(round(float(sum_c)/10)*10)-int(round(float(add_b)/10)*10)
#    sheet.col(offset_col+0).width = DATA_WIDTH
#    str_out=str(sum_c)+' - '+str(add_b)+' = '
#    sheet.write(offset_row,offset_col+0,str_out,style1)
    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(sum_c),style)
    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+1,"-",style)
    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(add_b),style)
    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+3,u'≈',style)

    file_h.write(str(add_a_s)+'  ')

    return;

##--------------------------------------------------
## Random select the calc_div function

def calc_add_sub_sel(offset_row,offset_col,style,sheet,file_h):
    select_mix = random.randint(0,1)
    if select_mix == 0 :
        calc_gen_add(offset_row,offset_col,style,sheet,file_h)
    else:
        calc_gen_sub(offset_row,offset_col,style,sheet,file_h)

    return;

##--------------------------------------------------
## Random select the calc_div function

def calc_add_sub_s_sel(offset_row,offset_col,style,sheet,file_h):
    select_mix = random.randint(0,1)
    if select_mix == 0 :
        calc_gen_add_s(offset_row,offset_col,style,sheet,file_h)
    else:
        calc_gen_sub_s(offset_row,offset_col,style,sheet,file_h)

    return;

