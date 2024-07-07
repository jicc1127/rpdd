# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpyinput_clvdt_into_ymdcow_args.py wb0N s0N wb1N s1N
#  wb0N : AB_rpdd.xlsx, s0N : yyyymmddCow00, wb1N : AB_calving.xlsx  s1N : calving
import sys
import rpdd


wb0N = sys.argv[1]
s0N = sys.argv[2]
wb1N = sys.argv[3]
s1N = sys.argv[4]

rpdd.fpyinput_clvdt_into_ymdcow(wb0N, s0N, wb1N, s1N)
print(" sheet\'yyyymmddCow00\' 分娩日, 産次を入力しました。")
 