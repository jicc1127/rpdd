# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpyinput_aidt_into_ymdheifer_args.py wb0N s0N wb1N s1N VWPm VWPM
#	  wb0N : str
#        Excelfile's name : 'AB_rpdd.xlsx'
#    s0N : str
#        sheet name : 'yyyymmddCow00'
#    wb1N : str
#        Excelfile's name : 'AB_AI.xlsx'
#        sheet name : 'AB_AI'
#	  s1N : str
#        sheet name : 'AB_AI'
#    VWPm : int
#        age in month to start waiting 授精待機開始月齢
#   VWPM : int
#		 age in month to start AI 授精開始月齢 
# input AIdata into sheet yyyymmddCow00
import sys
import rpdd


wb0N = sys.argv[1]
s0N = sys.argv[2]
wb1N = sys.argv[3]
s1N  = sys.argv[4]
VWPm = int(sys.argv[5])
VWPM = int(sys.argv[6])

rpdd.fpyinput_AIdt_into_ymdheifer(wb0N, s0N, wb1N, s1N, VWPm, VWPM)

print(" sheet\'yyyymmddHeifer00\' AIデータを入力しました。")
 