# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpyinput_eartagno_into_ymdcow_args.py wb0N s0N cid0n cet0n 
#		wb1N s1N cid1n cet1n
#  wb0N : AB_rpdd.xlsx, s0N : yyyymmddCow00, cid0n : 6, cet0n : 4, wb1N : AB_cowslist.xlsx  s1N :  cowslist, cid1n : 2, cet1n : 3
import sys
import rpdd


wb0N = sys.argv[1]
s0N = sys.argv[2]
cid0n = int(sys.argv[3])
cet0n = int(sys.argv[4])
wb1N = sys.argv[5]
s1N = sys.argv[6]
cid1n = int(sys.argv[7])
cet1n = int(sys.argv[8])

rpdd.fpyinput_eartagno_into_ymdcow(wb0N, s0N, cid0n, cet0n, wb1N, s1N, cid1n, cet1n)
print(" sheet\'yyyymmddCow00\' 耳標を入力しました。")
 