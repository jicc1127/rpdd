# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpydaysfrmcalving_args.py wbN sheetN col_bd col_clv col_dsfrmclv
#	  wbN : str
#        Excelbook's name  : '.\\AB_rpdd.xlsx'
#    sheetN : str
#        sheet name : 'yyyymmddCow00'
#    col_bd : int
#        column's number of base_date 基準日 : 18
#    col_clv : int
#        column's number of calving_date 分娩日 : 9
#    col_dsfrmclv : int
#        column's number of daysfrmcalving 分娩後日数 : 10
# calculate days from calving at sheet yyyymmddCow00
import sys
import rpdd


wbN = sys.argv[1]
sheetN = sys.argv[2]
col_bd = int(sys.argv[3])
col_clv  = int(sys.argv[4])
col_dsfrmclv = int(sys.argv[5])

rpdd.fpydaysfrmcalving(wbN, sheetN, col_bd, col_clv, col_dsfrmclv)
print(" sheet\'yyyymmddCow00\' 分娩後日数を入力しました。")
 