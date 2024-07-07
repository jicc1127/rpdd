# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpysheet_copy_args.py wbN sheetN sheetN_
#wbN : ..\AB_rpdd.xlsx, sheetN : yyyymmddCow00, sheetN_ : yyyymmddCow00all
import sys
import fmstls

wbN = sys.argv[1]
sheetN = sys.argv[2]
sheetN_ = sys.argv[3]

fmstls.fpysheet_copy( wbN, sheetN, sheetN_ )

print(sheetN + "をコピーして、" + sheetN_ + "を作成しました。")
