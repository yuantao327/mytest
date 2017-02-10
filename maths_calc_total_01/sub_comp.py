#coding=utf-8
#######################################################
#filename:sub_mult_div.py
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
# AA/ZZ + BB/ZZ (?) CC/ZZ + DD/ZZ
def calc_gen_comp1(offset_row,offset_col,style1,style2,style3,sheet,file_h):
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
    sheet.write(offset_row,offset_col,str(data_a),style2)
    sheet.write(offset_row+1,offset_col,str(data_z),style1)

    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+1, offset_col+1,"+",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_b),style2)
    sheet.write(offset_row+1,offset_col+2,str(data_z),style1)

    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+3, offset_col+3,u'○',style3) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(data_c),style2)
    sheet.write(offset_row+1,offset_col+4,str(data_z),style1)

    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+5, offset_col+5,"+",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+6).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+6,str(data_d),style2)
    sheet.write(offset_row+1,offset_col+6,str(data_z),style1)

    file_h.write(data_out+'  ')

    return;


#-------------------------------------------------------------
# SUM/ZZ - AA/ZZ (?) SUM/ZZ - CC/ZZ
def calc_gen_comp2(offset_row,offset_col,style1,style2,style3,sheet,file_h):
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
    sheet.write(offset_row,offset_col,str(sum_ab),style2)
    sheet.write(offset_row+1,offset_col,str(data_z),style1)

    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+1, offset_col+1,"-",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_a),style2)
    sheet.write(offset_row+1,offset_col+2,str(data_z),style1)

    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+3, offset_col+3,u'○',style3) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(sum_cd),style2)
    sheet.write(offset_row+1,offset_col+4,str(data_z),style1)

    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+5, offset_col+5,"-",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+6).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+6,str(data_c),style2)
    sheet.write(offset_row+1,offset_col+6,str(data_z),style1)

    file_h.write(data_out+'  ')

    return;

#-------------------------------------------------------------
# SUM/ZZ - AA/ZZ (?) CC/ZZ + DD/ZZ
def calc_gen_comp3(offset_row,offset_col,style1,style2,style3,sheet,file_h):
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
    sheet.write(offset_row,offset_col,str(sum_ab),style2)
    sheet.write(offset_row+1,offset_col,str(data_z),style1)

    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+1, offset_col+1,"-",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_a),style2)
    sheet.write(offset_row+1,offset_col+2,str(data_z),style1)

    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+3, offset_col+3,u'○',style3) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(data_c),style2)
    sheet.write(offset_row+1,offset_col+4,str(data_z),style1)

    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+5, offset_col+5,"+",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+6).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+6,str(data_d),style2)
    sheet.write(offset_row+1,offset_col+6,str(data_z),style1)

    file_h.write(data_out+'  ')

    return;

#-------------------------------------------------------------
#CC/ZZ + DD/ZZ (?) SUM/ZZ - AA/ZZ
def calc_gen_comp4(offset_row,offset_col,style1,style2,style3,sheet,file_h):
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
    sheet.write(offset_row,offset_col,str(data_c),style2)
    sheet.write(offset_row+1,offset_col,str(data_z),style1)

    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+1, offset_col+1,"+",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_d),style2)
    sheet.write(offset_row+1,offset_col+2,str(data_z),style1)

    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+3, offset_col+3,u"○",style3) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(sum_ab),style2)
    sheet.write(offset_row+1,offset_col+4,str(data_z),style1)

    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+5, offset_col+5,"-",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+6).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+6,str(data_a),style2)
    sheet.write(offset_row+1,offset_col+6,str(data_z),style1)

    file_h.write(data_out+'  ')

    return;

#-------------------------------------------------------------
# AA/ZZ + BB/ZZ (?) CC/ZZ + DD/ZZ
def calc_gen_comp5(offset_row,offset_col,style1,style2,style3,sheet,file_h):
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
    sheet.write(offset_row,offset_col,str(data_a),style2)
    sheet.write(offset_row+1,offset_col,str(data_z),style1)

    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+1, offset_col+1,"+",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_b),style2)
    sheet.write(offset_row+1,offset_col+2,str(data_z),style1)

    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+3, offset_col+3,u"○",style3) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(data_c),style2)
    sheet.write(offset_row+1,offset_col+4,str(data_z_tmp),style1)

    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+5, offset_col+5,"+",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+6).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+6,str(data_d),style2)
    sheet.write(offset_row+1,offset_col+6,str(data_z_tmp),style1)

    file_h.write(data_out+'  ')

    return;

#-------------------------------------------------------------
# SUM/ZZ - AA/ZZ (?) SUM/ZZ - CC/ZZ
def calc_gen_comp6(offset_row,offset_col,style1,style2,style3,sheet,file_h):
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
    sheet.write(offset_row,offset_col,str(sum_ab),style2)
    sheet.write(offset_row+1,offset_col,str(data_z),style1)

    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+1, offset_col+1,"-",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_a),style2)
    sheet.write(offset_row+1,offset_col+2,str(data_z),style1)

    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+3, offset_col+3,u"○",style3) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(sum_cd),style2)
    sheet.write(offset_row+1,offset_col+4,str(data_z_tmp),style1)

    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+5, offset_col+5,"-",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+6).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+6,str(data_c),style2)
    sheet.write(offset_row+1,offset_col+6,str(data_z_tmp),style1)

    file_h.write(data_out+'  ')

    return;

#-------------------------------------------------------------
# SUM/ZZ - AA/ZZ (?) CC/ZZ + DD/ZZ
def calc_gen_comp7(offset_row,offset_col,style1,style2,style3,sheet,file_h):
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
    sheet.write(offset_row,offset_col,str(sum_ab),style2)
    sheet.write(offset_row+1,offset_col,str(data_z),style1)

    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+1, offset_col+1,"-",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_a),style2)
    sheet.write(offset_row+1,offset_col+2,str(data_z),style1)

    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+3, offset_col+3,u"○",style3) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(data_c),style2)
    sheet.write(offset_row+1,offset_col+4,str(data_z_tmp),style1)

    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+5, offset_col+5,"+",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+6).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+6,str(data_d),style2)
    sheet.write(offset_row+1,offset_col+6,str(data_z_tmp),style1)

    file_h.write(data_out+'  ')

    return;


#-------------------------------------------------------------
#CC/ZZ + DD/ZZ (?) SUM/ZZ - AA/ZZ
def calc_gen_comp8(offset_row,offset_col,style1,style2,style3,sheet,file_h):
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
    sheet.write(offset_row,offset_col,str(data_c),style2)
    sheet.write(offset_row+1,offset_col,str(data_z_tmp),style1)

    sheet.col(offset_col+1).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+1, offset_col+1,"+",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+2).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+2,str(data_d),style2)
    sheet.write(offset_row+1,offset_col+2,str(data_z_tmp),style1)

    sheet.col(offset_col+3).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+3, offset_col+3,u"○",style3) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+4).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+4,str(sum_ab),style2)
    sheet.write(offset_row+1,offset_col+4,str(data_z),style1)

    sheet.col(offset_col+5).width = SIGN_WIDTH
    sheet.write_merge(offset_row, offset_row+1, offset_col+5, offset_col+5,"-",style1) # Merges row 0's columns 0 through 10.

    sheet.col(offset_col+6).width = DATA_WIDTH
    sheet.write(offset_row,offset_col+6,str(data_a),style2)
    sheet.write(offset_row+1,offset_col+6,str(data_z),style1)

    file_h.write(data_out+'  ')

    return;

##--------------------------------------------------
## Random select the calc_compare function
def calc_comp_sel1(offset_row,offset_col,style1,style2,style3,sheet,file_h):
    select_mix = random.randint(1,4)
    if select_mix == 1 :
        calc_gen_comp1(offset_row,offset_col,style1,style2,style3,sheet,file_h)
    elif select_mix == 2 :
        calc_gen_comp2(offset_row,offset_col,style1,style2,style3,sheet,file_h)
    elif select_mix == 3 :
        calc_gen_comp3(offset_row,offset_col,style1,style2,style3,sheet,file_h)
    elif select_mix == 4 :
        calc_gen_comp4(offset_row,offset_col,style1,style2,style3,sheet,file_h)
    else:
        calc_gen_comp4(offset_row,offset_col,style1,style2,style3,sheet,file_h)

    return;

def calc_comp_sel2(offset_row,offset_col,style1,style2,style3,sheet,file_h):
    select_mix = random.randint(1,4)
    if select_mix == 1 :
        calc_gen_comp5(offset_row,offset_col,style1,style2,style3,sheet,file_h)
    elif select_mix == 2 :
        calc_gen_comp6(offset_row,offset_col,style1,style2,style3,sheet,file_h)
    elif select_mix == 3 :
        calc_gen_comp7(offset_row,offset_col,style1,style2,style3,sheet,file_h)
    elif select_mix == 4 :
        calc_gen_comp8(offset_row,offset_col,style1,style2,style3,sheet,file_h)
    else:
        calc_gen_comp8(offset_row,offset_col,style1,style2,style3,sheet,file_h)

    return;

def calc_comp_sel(offset_row,offset_col,style1,style2,style3,sheet,file_h):
    select_mix = random.randint(1,3)
    if select_mix == 1 :
        calc_comp_sel2(offset_row,offset_col,style1,style2,style3,sheet,file_h)
    else:
        calc_comp_sel1(offset_row,offset_col,style1,style2,style3,sheet,file_h)

    return;