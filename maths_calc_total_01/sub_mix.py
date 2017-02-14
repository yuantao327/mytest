#coding=utf-8
#######################################################
#filename:sub_mix.py
#author:ytliang
#date:xxxx-xx-xx
#function��multiplier and divider functions
#######################################################

#-------------------------------------------------------------
# AAA+BBB=CCCC

DATA_WIDTH = 1500
SIGN_WIDTH = 1000
BLANK_HEIGHT = 1050
DATA_HEIGHT = 300
CIR_HEIGHT  = 800

import random

#-------------------------------------------------------------
# AAA+BBB+CCC=DDDD
def calc_gen_mix_0(offset_row,offset_col,style,sheet,file_h):
    add_a = random.randint(100,999)
    add_b = random.randint(100,999)
    add_c = random.randint(100,999)
    sum_d = add_a + add_b + add_c
    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(add_a),style)
    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+1,"+",style)
    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(add_b),style)
    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+3,"+",style)
    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(add_c),style)
    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+5,"=",style)
    file_h.write(str(sum_d)+'  ')

    return;

#-------------------------------------------------------------
# AAA*B + CCC = DDDD or CCC+AAA*B = DDDD
def calc_gen_mix_1(offset_row,offset_col,style,sheet,file_h):
    mult_a = random.randint(100,499)
    mult_b = random.randint(2,9)
    add_c = random.randint(10,499)
    sum_d = mult_a*mult_b+add_c
#    sheet.col(offset_col+0).width = DATA_WIDTH
    exchange = random.randint(0,1)
    if exchange <> 0 :
        sheet.col(offset_col).width = DATA_WIDTH
        sheet.write(offset_row,offset_col,str(mult_a),style)
        sheet.col(offset_col+1).width = SIGN_WIDTH
        sheet.write(offset_row,offset_col+1,u'×',style)
        sheet.col(offset_col+2).width = DATA_WIDTH
        sheet.write(offset_row,offset_col+2,str(mult_b),style)
        sheet.col(offset_col+3).width = SIGN_WIDTH
        sheet.write(offset_row,offset_col+3,"+",style)
        sheet.col(offset_col+4).width = DATA_WIDTH
        sheet.write(offset_row,offset_col+4,str(add_c),style)
        sheet.col(offset_col+5).width = SIGN_WIDTH
        sheet.write(offset_row,offset_col+5,"=",style)
#        str_out=str(mult_a)+u'�'+str(mult_b)+ '+' + str(add_c) +'='
#        sheet.write(offset_row,offset_col+0,str_out,style)
    else :
        sheet.col(offset_col).width = DATA_WIDTH
        sheet.write(offset_row,offset_col,str(add_c),style)
        sheet.col(offset_col+1).width = SIGN_WIDTH
        sheet.write(offset_row,offset_col+1,"+",style)
        sheet.col(offset_col+2).width = DATA_WIDTH
        sheet.write(offset_row,offset_col+2,str(mult_a),style)
        sheet.col(offset_col+3).width = SIGN_WIDTH
        sheet.write(offset_row,offset_col+3,u'×',style)
        sheet.col(offset_col+4).width = DATA_WIDTH
        sheet.write(offset_row,offset_col+4,str(mult_b),style)
        sheet.col(offset_col+5).width = SIGN_WIDTH
        sheet.write(offset_row,offset_col+5,"=",style)
#        str_out=str(add_c)+'+'+str(mult_a)+u'�'+str(mult_b)+'='
#        sheet.write(offset_row,offset_col+0,str_out,style)
    file_h.write(str(sum_d)+'  ')

    return;
#-------------------------------------------------------------
# AAA/B + CCC = DDD
def calc_gen_mix_2(offset_row,offset_col,style,sheet,file_h):
    data_a = random.randint(1,9)
    data_b = random.randint(3,9)
    data_t = data_a*data_b
    data_c = random.randint(10,499)
    data_d = data_t/data_a+data_c
#    sheet.col(offset_col+0).width = DATA_WIDTH
    exchange = random.randint(0,1)
    if exchange <> 0 :
#        str_out=str(data_c)+'+'+str(data_t)+ u'�' + str(data_a) +'='
#        sheet.write(offset_row,offset_col+0,str_out,style)
        sheet.col(offset_col).width = DATA_WIDTH
        sheet.write(offset_row,offset_col,str(data_c),style)
        sheet.col(offset_col+1).width = SIGN_WIDTH
        sheet.write(offset_row,offset_col+1,"+",style)
        sheet.col(offset_col+2).width = DATA_WIDTH
        sheet.write(offset_row,offset_col+2,str(data_t),style)
        sheet.col(offset_col+3).width = SIGN_WIDTH
        sheet.write(offset_row,offset_col+3,u'÷',style)
        sheet.col(offset_col+4).width = DATA_WIDTH
        sheet.write(offset_row,offset_col+4,str(data_a),style)
        sheet.col(offset_col+5).width = SIGN_WIDTH
        sheet.write(offset_row,offset_col+5,"=",style)
    else :
#        str_out=str(data_t)+ u'�' + str(data_a) + '+'+ str(data_c)+'='
#        sheet.write(offset_row,offset_col+0,str_out,style)
        sheet.col(offset_col).width = DATA_WIDTH
        sheet.write(offset_row,offset_col,str(data_t),style)
        sheet.col(offset_col+1).width = SIGN_WIDTH
        sheet.write(offset_row,offset_col+1,u'÷',style)
        sheet.col(offset_col+2).width = DATA_WIDTH
        sheet.write(offset_row,offset_col+2,str(data_a),style)
        sheet.col(offset_col+3).width = SIGN_WIDTH
        sheet.write(offset_row,offset_col+3,"+",style)
        sheet.col(offset_col+4).width = DATA_WIDTH
        sheet.write(offset_row,offset_col+4,str(data_c),style)
        sheet.col(offset_col+5).width = SIGN_WIDTH
        sheet.write(offset_row,offset_col+5,"=",style)
    file_h.write(str(data_d)+'  ')

    return;

#-------------------------------------------------------------
# DDD-AAA-BBB = CCC
def calc_gen_mix_3(offset_row,offset_col,style,sheet,file_h):
    add_a = random.randint(10,499)
    add_b = random.randint(10,499)
    add_c = random.randint(10,499)
    sum_d = add_a + add_b + add_c
#    sheet.col(offset_col+0).width = DATA_WIDTH
#    str_out=str(sum_d)+'-'+str(add_a)+ '-'+ str(add_b) + '='
#    sheet.write(offset_row,offset_col+0,str_out,style)
    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(sum_d),style)
    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+1,"-",style)
    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(add_a),style)
    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+3,"-",style)
    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(add_b),style)
    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+5,"=",style)
    file_h.write(str(add_c)+'  ')

    return;

#-------------------------------------------------------------
# DDD-(AAA+BBB) = CCC
def calc_gen_mix_4(offset_row,offset_col,style,sheet,file_h):
    add_a = random.randint(10,499)
    add_b = random.randint(10,499)
    add_c = random.randint(10,499)
    sum_d = add_a + add_b + add_c
#    sheet.col(offset_col+0).width = DATA_WIDTH
#    str_out=str(sum_d) + '-' + '(' + str(add_a) + '+' + str(add_b)+')' + '='
#    sheet.write(offset_row,offset_col+0,str_out,style)
    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(sum_d),style)
    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+1,"-",style)
    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,"("+str(add_a),style)
    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+3,"+",style)
    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(add_b)+")",style)
    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+5,"=",style)
    file_h.write(str(add_c)+'  ')

    return;

#-------------------------------------------------------------
# DDD-AAA+BBB = CCC
def calc_gen_mix_5(offset_row,offset_col,style,sheet,file_h):
    data_d = random.randint(10,999)
    data_a = random.randint(10,data_d)
    data_b = random.randint(10,data_a)
    data_c = data_d - data_a + data_b
#    sheet.col(offset_col+0).width = DATA_WIDTH
#    str_out=str(data_d) + '-' + str(data_a) + '+' + str(data_b) + '='
#    sheet.write(offset_row,offset_col+0,str_out,style)
    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(data_d),style)
    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+1,"-",style)
    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_a),style)
    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+3,"+",style)
    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(data_b),style)
    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+5,"=",style)
    file_h.write(str(data_c)+'  ')

    return;

#-------------------------------------------------------------
# DDD-(AAA-BBB) = CCC
def calc_gen_mix_6(offset_row,offset_col,style,sheet,file_h):
    data_d = random.randint(10,999)
    data_a = random.randint(10,data_d)
    data_b = random.randint(10,data_a)
    data_c = data_d - data_a + data_b
#    sheet.col(offset_col+0).width = DATA_WIDTH
#    str_out=str(data_d) + '-' + '(' +  str(data_a) + '-' + str(data_b) + ')' + '='
#    sheet.write(offset_row,offset_col+0,str_out,style)
    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(data_d),style)
    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+1,"-",style)
    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,"("+str(data_a),style)
    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+3,"-",style)
    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(data_b)+")",style)
    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+5,"=",style)
    file_h.write(str(data_c)+'  ')

    return;

#-------------------------------------------------------------
# DDDD-AAA*B = CCC
def calc_gen_mix_7(offset_row,offset_col,style,sheet,file_h):
    mult_a = random.randint(10,199)
    mult_b = random.randint(2,9)
    add_c = random.randint(10,499)
    sum_d = mult_a*mult_b+add_c
#    sheet.col(offset_col+0).width = DATA_WIDTH
#    str_out=str(sum_d)+'-'+str(mult_a)+u'�'+str(mult_b)+'='
#    sheet.write(offset_row,offset_col+0,str_out,style)
    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(sum_d),style)
    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+1,"-",style)
    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(mult_a),style)
    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+3,u'×',style)
    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(mult_b),style)
    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+5,"=",style)
    file_h.write(str(add_c)+'  ')

    return;

#-------------------------------------------------------------
# DDDD-AAA*B = CCC
def calc_gen_mix_8(offset_row,offset_col,style,sheet,file_h):
    data_a = random.randint(1,9)
    data_b = random.randint(2,9)
    data_t = data_a*data_b
    data_c = random.randint(10,499)
    data_d = data_t/data_a+data_c
#    sheet.col(offset_col+0).width = DATA_WIDTH
#    str_out=str(data_d)+'-'+str(data_t)+u'�'+str(data_a)+'='
#    sheet.write(offset_row,offset_col+0,str_out,style)
    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(data_d),style)
    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+1,"-",style)
    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_t),style)
    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+3,u'×',style)
    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(data_a),style)
    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+5,"=",style)
    file_h.write(str(data_c)+'  ')

    return;

#-------------------------------------------------------------
#AA/B*CCC = DDD
def calc_gen_mix_9(offset_row,offset_col,style,sheet,file_h):
    data_a = random.randint(1,9)
    data_b = random.randint(3,9)
    data_t = data_a*data_b
    data_c = random.randint(10,499)
    data_d = data_t/data_a*data_c
#    sheet.col(offset_col+0).width = DATA_WIDTH
#    str_out=str(data_t)+u'�'+str(data_a)+u'�'+str(data_c)+'='
#    sheet.write(offset_row,offset_col+0,str_out,style)
    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(data_t),style)
    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+1,u'÷',style)
    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_a),style)
    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+3,u'×',style)
    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(data_c),style)
    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+5,"=",style)
    file_h.write(str(data_d)+'  ')
    return;

#-------------------------------------------------------------
#CCC*AA/B = DDD
def calc_gen_mix_10(offset_row,offset_col,style,sheet,file_h):
    data_a = random.randint(1,9)
    data_b = random.randint(3,9)
    data_t = data_a*data_b
    data_c = random.randint(10,499)
    data_d = data_t/data_a*data_c
#    sheet.col(offset_col+0).width = DATA_WIDTH
#    str_out=str(data_c)+u'�'+str(data_t)+u'�'+str(data_a)+'='
#    sheet.write(offset_row,offset_col+0,str_out,style)
    sheet.col(offset_col).width = DATA_WIDTH
    sheet.write(offset_row,offset_col,str(data_c),style)
    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+1,u'×',style)
    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_t),style)
    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+3,u'÷',style)
    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(data_a),style)
    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write(offset_row,offset_col+5,"=",style)
    file_h.write(str(data_d)+'  ')
    return;

##--------------------------------------------------
## Random select the calc_compare function
def calc_mixed_sel(offset_row,offset_col,style,sheet,file_h):
    select_mix = random.randint(0,10)
    if select_mix == 0 :
        calc_gen_mix_0(offset_row,offset_col,style,sheet,file_h)
    elif select_mix == 1 :
        calc_gen_mix_1(offset_row,offset_col,style,sheet,file_h)
    elif select_mix == 2 :
        calc_gen_mix_2(offset_row,offset_col,style,sheet,file_h)
    elif select_mix == 3 :
        calc_gen_mix_3(offset_row,offset_col,style,sheet,file_h)
    elif select_mix == 4 :
        calc_gen_mix_4(offset_row,offset_col,style,sheet,file_h)
    elif select_mix == 5 :
        calc_gen_mix_5(offset_row,offset_col,style,sheet,file_h)
    elif select_mix == 6 :
        calc_gen_mix_6(offset_row,offset_col,style,sheet,file_h)
    elif select_mix == 7 :
        calc_gen_mix_7(offset_row,offset_col,style,sheet,file_h)
    elif select_mix == 8 :
        calc_gen_mix_8(offset_row,offset_col,style,sheet,file_h)
    elif select_mix == 9 :
        calc_gen_mix_9(offset_row,offset_col,style,sheet,file_h)
    elif select_mix == 10 :
        calc_gen_mix_10(offset_row,offset_col,style,sheet,file_h)
    else :
        calc_gen_mix_10(offset_row,offset_col,style,sheet,file_h)


