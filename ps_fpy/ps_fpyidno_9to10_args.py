# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpyidno_9to10_args.py wbN sheetN col
import sys
import fmstls

wbN = sys.argv[1]
sheetN = sys.argv[2]
col = int(sys.argv[3])

fmstls.fpyidNo_9to10(wbN, sheetN, col)

print("個体識別番号を１０桁文字列にしました。")
