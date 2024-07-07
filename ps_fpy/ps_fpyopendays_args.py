# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpyopendays_args.py wbN yyyymmddCow00
import sys
import mh_rpdu

wbN = sys.argv[1]
sheetN = sys.argv[2]

mh_rpdu.fpyopenDays( wbN, sheetN )
print(" sheet" + sheetN + " の 空胎日数を計算、入力しました。")
 