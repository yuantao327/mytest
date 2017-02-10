#coding=utf-8
#######################################################
#filename:sub_mult_div.py
#author:ytliang
#date:xxxx-xx-xx
#function£ºmultiplier and divider functions
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
# AAAxB=CCCC
def calc_gen_mult(offset_row,offset_col,style,sheet,file_h):
    mult_b = random.randint(2,9)
    mult_a = random.randint(10,999)
    mult = mult_a*mult_b
#    sheet.col(offset_col+0).width = DATA_WIDTH
    exchange = random.randint(0,20)
    if exchange <> 0 :
        sheet.col(offset_col).width = DATA_WIDTH
        sheet.write(offset_row,offset_col,str(mult_a),style)
        sheet.col(offset_col+1).width = SIGN_WIDTH
        sheet.write(offset_row,offset_col+1, ' ¡Á ',style)
        sheet.col(offset_col+2).width = DATA_WIDTH
        sheet.write(offset_row,offset_col+2,str(mult_b),style)
        sheet.col(offset_col+3).width = SIGN_WIDTH
        sheet.write(offset_row,offset_col+3,"=",style)
#        str_out=str(mult_a)+u' ¡Á '+str(mult_b)+' = '
#        sheet.write(offset_row,offset_col+0,str_out,style)
    else :
        sheet.col(offset_col).width = DATA_WIDTH
        sheet.write(offset_row,offset_col,str(mult_b),style)
        sheet.col(offset_col+1).width = SIGN_WIDTH
        sheet.write(offset_row,offset_col+1, ' ¡Á ',style)
        sheet.col(offset_col+2).width = DATA_WIDTH
        sheet.write(offset_row,offset_col+2,str(mult_a),style)
        sheet.col(offset_col+3).width = SIGN_WIDTH
        sheet.write(offset_row,offset_col+3,"=",style)
#        str_out= str(mult_b) + u' ¡Á ' + str(mult_a) + ' = '
#        sheet.write(offset_row,offset_col+0,str_out,style)


    file_h.write(str(mult)+'  ')

    return;

#-------------------------------------------------------------
# AAAxB=CCCC
def calc_gen_mult_s(offset_row,offset_col,style,sheet,file_h):
    mult_b = random.randint(2,9)
    mult_a = random.randint(10,999)
    mult = mult_a*mult_b
    mult_s = int(round(float(mult_a)/10)*10)*mult_b
#    sheet.col(offset_col+0).width = DATA_WIDTH
    exchange = random.randint(0,2)
    if exchange <> 0 :
        sheet.col(offset_col).width = DATA_WIDTH
        sheet.write(offset_row,offset_col,str(mult_a),style)
        sheet.col(offset_col+1).width = SIGN_WIDTH
        sheet.write(offset_row,offset_col+1, ' ¡Á ',style)
        sheet.col(offset_col+2).width = DATA_WIDTH
        sheet.write(offset_row,offset_col+2,str(mult_b),style)
        sheet.col(offset_col+3).width = SIGN_WIDTH
        sheet.write(offset_row,offset_col+3,'¡Ö',style)
#        str_out=str(mult_a)+u' ¡Á '+str(mult_b)+' = '
#        sheet.write(offset_row,offset_col+0,str_out,style)
    else :
        sheet.col(offset_col).width = DATA_WIDTH
        sheet.write(offset_row,offset_col,str(mult_b),style)
        sheet.col(offset_col+1).width = SIGN_WIDTH
        sheet.write(offset_row,offset_col+1, ' ¡Á ',style)
        sheet.col(offset_col+2).width = DATA_WIDTH
        sheet.write(offset_row,offset_col+2,str(mult_a),style)
        sheet.col(offset_col+3).width = SIGN_WIDTH
        sheet.write(offset_row,offset_col+3,'¡Ö',style)
#        str_out= str(mult_b) + u' ¡Á ' + str(mult_a) + ' = '
#        sheet.write(offset_row,offset_col+0,str_out,style)


    file_h.write(str(mult_s)+'  ')

    return;

