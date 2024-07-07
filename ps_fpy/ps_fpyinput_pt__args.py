# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpyinput_pt__args.py wbAIN sheetAIN wbRPDN sRPDN
# wbAIN : AB_AI.xlsx, sheetAIN : AB_AI, wbRPDN : AB_rpdd.xlsx, sRPDN : yyyymmddCow01
# mh_rpdu.py と区別するため、 ps_fpyinput_pth_ _args.py と_ _ アンダーバー2つにしている。
import sys
import rpdd


wbAIN = sys.argv[1]
sheetAIN = sys.argv[2]
wbRPDN = sys.argv[3]
sRPDN = sys.argv[4]

rpdd.fpyinput_PT_(wbAIＮ, sheetAIN, wbRPDN, sRPDN)
print(" sheet\'yyyymmddCow01\' 鑑定結果を入力しました。")
 