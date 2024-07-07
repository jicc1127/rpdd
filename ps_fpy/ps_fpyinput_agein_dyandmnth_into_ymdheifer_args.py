# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpyinput_agein_dyandmnth_into_ymdheifer_args.py wbN sheetN
#  wbN : str
#        Excel file's name : AB_rpdd.xlsx
#    sheetN : str
#        sheet name : yyyymmddHeifer00
import sys
import rpdd


wbN = sys.argv[1]
sheetN = sys.argv[2]

rpdd.fpyinput_agein_dyandmnth_into_ymdheifer(wbN, sheetN)
print(" sheet\'yyyymmddHeifer00\' 産次数(=0), 日齢、 月齢を入力しました。")
 