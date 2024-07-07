# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpyinput_cwlstd_to_cow_heifer00_args.py wbN0 sheetN0 ncol0 wbN1 sheetN1 ncol1 fstcol wbN s0N s1N enofetn
# wbN0 : AB_cowslist.xlsx, sheetN0 : cowslistyyyymmdd, 
# ncol0 : 20 (the number of columns of sheet cowslist*'s list), 
# wbN1 : AB_calving.xlsx, sheetN1 : calving, 
# ncol1 : 11 (the number of columns of sheet calving's list), 
# fstcol : 1 first column number to input data 
# wbN: AB_rpdd.xlsx, s0N : yyyymmddHeifer00, s1N : yyyymmddCow00 
# enofetn : 3(DHINo), 2 (eartagNo)

import sys
import rpdd

wbN0 = sys.argv[1]
sheetN0 = sys.argv[2]
ncol0 = int( sys.argv[3] )
wbN1 = sys.argv[4]
sheetN1 = sys.argv[5]
ncol1 = int( sys.argv[6] )

cowslists = rpdd.fpysrt_into_Cow_Heiferlst(wbN0, sheetN0, ncol0, wbN1, sheetN1, ncol1)

fstcol = int(sys.argv[7])
wbN =  sys.argv[8]
s0N= sys.argv[9]
s1N = sys.argv[10]
enofetn = int(sys.argv[11])
rpdd.fpyinput_cwlstd_to_Cow_Heifer00( cowslists, fstcol, wbN, s0N, s1N, enofetn )

print(  "sheet yyyymmddHeifer00, yyyymmddCow00 にcowslists dataを入力しました。"  )
 