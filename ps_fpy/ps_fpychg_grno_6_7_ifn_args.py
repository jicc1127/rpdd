# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpychg_grno_6_7_ifn_args.py wbN sheetN srefN
#wbN :  'AB_rpdd.xlsx', sheetN : ''yyyymmddCow00', srefN : yyyymmddCow_uorg
import sys
import rpdd

wbN = sys.argv[1]
sheetN = sys.argv[2]
srefN = sys.argv[3]

rpdd.fpychg_grNo_6_7_ifn(wbN, sheetN, srefN)

print("Group と Stage に '6 乾乳' and '7 繁殖対象外'　の分類を加えました。")
