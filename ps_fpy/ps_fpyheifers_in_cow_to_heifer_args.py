# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpyheifers_in_cow_to_heifer_args.py wbN s0N s1N
# wbN : str
#        Excelbook's name  : '.\\MH_rpdd.xlsx'
#    s0N : str
#        sheet name : 'yyyymmddCow00'
#    s1N : str
#        sheet name : 'yyyymmddHeifer00'
# transfer heifers from sheet yyyymmddCow00 to yyyymmddHeifer00
#    基準日以降に初産分娩した未経産牛を、
#    sheet yyyymmddCow00 から yyyymmddHeifer00 へ移動する
#    移動前のyyyymmddCow00をyyyymmddCow00bckとして残すようにしてある。

import sys
import rpdd

wbN = sys.argv[1]
s0N = sys.argv[2]
s1N  = sys.argv[3]

rpdd.fpyheifers_in_cow_to_heifer(wbN, s0N, s1N)

print("基準日以降に初産分娩した未経産牛がある場合、")
print('sheet yyyymmddCow00 から yyyymmddHeifer00 へ移動しています。')
