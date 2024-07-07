# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpysep_out_frm_00_args.py wbN sheetN ncol
import sys
import rpdd

wbN = sys.argv[1]
sheetN = sys.argv[2]
ncol = int(sys.argv[3])

rpdd.fpysep_out_frm_00(wbN, sheetN, ncol)

print("繁殖対象外の個体を、分離しました。")
