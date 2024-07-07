# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpyopenmnths_args.py wbN yyyymmddHeifer00
import sys
import mh_rpdu

wbN = sys.argv[1]
sheetN = sys.argv[2]

mh_rpdu.fpyopenMnths( wbN, sheetN )
print(" sheet" + sheetN + " の 空胎月数を計算、入力しました。")
 