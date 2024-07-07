# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpyrpdd_mh_tools00_args.py wb0N s0cN ncolc s0uN coluidNo s1cN
# wb0N : AB_rpdd.xlsx, s0cN : yyyymmddCow00, 
# ncolc : 18  (number of columns of sheet s0cN 'yyyymmddCow00' )
# s0uN : yyyymmddCow_uorg
# coluidNo : 5 (idNo's columun number' at sheet s0uN 'yyyymmddCow_uorg')
#s1cN: yyyymmddCow00all


import sys
import rpdd

wb0N = sys.argv[1]
s0cN = sys.argv[2]
ncolc =  int(sys.argv[3])
s0uN = sys.argv[4]
coluidNo = int(sys.argv[5])
s1cN =  sys.argv[6]


rpdd.fpyrpdd_MH_tools00( wb0N, s0cN, ncolc, s0uN, coluidNo, s1cN)

print(  "sheet yyyymmddCowout を分離して、sheet  yyyymmddCow00 を完成しました。"  )
 