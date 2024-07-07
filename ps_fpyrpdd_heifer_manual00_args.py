# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpyrpdd_heifer_manual00_args.py wb0N s0hN wb1N s1N VWPm VWPM 
# wb0N : AB_rpdd.xlsx, 
# s0hN : yyyymmddHeifer00, 
# wb1N : AB_AI.xlsx, s1N : AB_AI, 
#VWPm : 13
#VWPM : 14

import sys
import rpdd

wb0N = sys.argv[1]
s0hN = sys.argv[2]
wb1N = sys.argv[3]
s1N = sys.argv[4]
VWPm = int(sys.argv[5])
VWPM = int(sys.argv[6])

rpdd.fpyrpdd_Heifer_manual00( wb0N, s0hN, wb1N, s1N, VWPm, VWPM )

print(  "sheet  yyyymmddHeifer00 を完成しました。"  )
 