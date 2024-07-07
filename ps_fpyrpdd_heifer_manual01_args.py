# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpyrpdd_heifer_manual01_args.py wb0N s0hN s0h_N wb1N s1N 
# wb0N : AB_rpdd.xlsx, s0hN : yyyymmddHeifer00, 
# s0h_N : yyyymmddHeifer01, 
# wb1N : AB_AI.xlsx, s1N : AB_AI, 

import sys
import rpdd

wb0N = sys.argv[1]
s0hN = sys.argv[2]
s0h_N =  sys.argv[3]
wb1N = sys.argv[4]
s1N = sys.argv[5]

rpdd.fpyrpdd_Heifer_manual01( wb0N, s0hN, s0h_N, wb1N, s1N )

print(  "sheet , yyyymmddHeifer01 に鑑定結果を入力しました。"  )
 