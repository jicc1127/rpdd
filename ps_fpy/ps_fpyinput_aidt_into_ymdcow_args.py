# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpyinput_aidt_into_ymdcow_args.py wb0N s0N wb1N s1N VWP
#	  wb0N : str
#        Excelfile's name : 'AB_rpdd.xlsx'
#    s0N : str
#        sheet name : 'yyyymmddCow00'
#    wb1N : str
#        Excelfile's name : 'AB_AI.xlsx'
#        sheet name : 'AB_AI'
#	  s1N : str
#        sheet name : 'AB_AI'
#    VWP : int
#        volantary waiting period 
# input AIdata into sheet yyyymmddCow00
import sys
import rpdd


wb0N = sys.argv[1]
s0N = sys.argv[2]
wb1N = sys.argv[3]
s1N  = sys.argv[4]
VWP = int(sys.argv[5])

rpdd.fpyinput_AIdt_into_ymdcow(wb0N, s0N, wb1N, s1N, VWP)
print(" sheet\'yyyymmddCow00\' AIデータを入力しました。")
 