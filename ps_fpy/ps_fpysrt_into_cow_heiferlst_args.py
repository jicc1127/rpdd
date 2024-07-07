# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpysrt_into_cow_heiferlst_args.py wbN0 sheetN0 ncol0 wbN1 sheetN1 ncol1
# wbN0 : AB_cowslist.xlsx, sheetN0 : cowslistyyyymmdd, 
# ncol0 : 20 (the number of columns of sheet cowslist*'s list), 
# wbN1 : AB_calving.xlsx, sheetN1 : calving, 
# ncol1 : 11 (the number of columns of sheet calving's list), 
import sys
import rpdd

wbN0 = sys.argv[1]
sheetN0 = sys.argv[2]
ncol0 = int( sys.argv[3] )
wbN1 = sys.argv[4]
sheetN1 = sys.argv[5]
ncol1 = int( sys.argv[6] )

cowslists = rpdd.fpysrt_into_Cow_Heiferlst(wbN0, sheetN0, ncol0, wbN1, sheetN1, ncol1)

print(  "cowslists = [Heifer, Cow] を作成しました。"  )
 