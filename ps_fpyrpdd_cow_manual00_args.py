# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpyrpdd_cow_manual00_args.py wb0N s0cN s0hN wb1N s1N wb2N s2N VWP 
# wb0N : AB_rpdd.xlsx, s0cN : yyyymmddCow00, 
# s0hN : yyyymmddHeifer00, 
# wb1N : AB_calving.xlsx, s1N : calving, 
# wb2N : AB_AI.xlsx,
#s2Nl : AB_AI
#VWP : 30, 50

import sys
import rpdd

wb0N = sys.argv[1]
s0cN = sys.argv[2]
s0hN =  sys.argv[3]
wb1N = sys.argv[4]
s1N = sys.argv[5]
wb2N =  sys.argv[6]
s2N= sys.argv[7]
VWP = int(sys.argv[8])

rpdd.fpyrpdd_Cow_manual00( wb0N, s0cN, s0hN, wb1N, s1N, wb2N, s2N, VWP )

print(  "sheet  yyyymmddCow00 を完成しました。"  )
 