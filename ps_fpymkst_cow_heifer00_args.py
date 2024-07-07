# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpymkst_cow_heifer00.py wbN scolN colnc colnh bdate
# wbN : AB_rpdd.xlsx, scolN : columns, colnc : 1 (row number of cow's sheet list title), 
# colnh : 3 (row number fo heifer's sheet list title), bdate : yyyy/mm/dd
import sys
import rpdd

wbN = sys.argv[1]
scolN = sys.argv[2]
colnc = int( sys.argv[3] )
colnh= int( sys.argv[4] )
bdate = sys.argv[5]


rpdd.fpymkst_Cow_Heifer00( wbN, scolN, colnc, colnh, bdate)

print( wbN+ ".xlsxにsheet yyyymmddHeifer00 と yyyymmddCow00 を作成しました。"  )
 