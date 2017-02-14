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
# AA/CC + BB/CC= DD/CC
def calc_div_add(offset_row,offset_col,style1,style2,sheet,file_h):
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
    sheet.write(offset_row,offset_col,str(data_a),style2)
    sheet.write(offset_row+1,offset_col,str(data_c),style1)

    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+1, offset_col+1,"+",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_b),style2)
    sheet.write(offset_row+1,offset_col+2,str(data_c),style1)

    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+3, offset_col+3,"=",style1) # Merges row 0's columns 0 through 10.

    file_h.write(str(data_out)+'  ')

    return;

# AA/CC - BB/CC= DD/CC
def calc_div_sub(offset_row,offset_col,style1,style2,sheet,file_h):
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
    sheet.write(offset_row,offset_col,str(data_d),style2)
    sheet.write(offset_row+1,offset_col,str(data_c),style1)

    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+1, offset_col+1,"-",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_a),style2)
    sheet.write(offset_row+1,offset_col+2,str(data_c),style1)

    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+3, offset_col+3,"=",style1) # Merges row 0's columns 0 through 10.

    file_h.write(str(data_out)+'  ')

    return;


##--------------------------------------------------
## Random select the calc_div function

def calc_div_sel(offset_row,offset_col,style1,style2,sheet,file_h):
    select_mix = random.randint(0,1)
    if select_mix == 0 :
        calc_div_add(offset_row,offset_col,style1,style2,sheet,file_h)
    else:
        calc_div_sub(offset_row,offset_col,style1,style2,sheet,file_h)

    return;

