import openpyxl 
import glob
import time
from copy import copy

# 定義
max_r = 14 # 最大の行
max_l = 66 # 最大の列

calc_r = [2,3,4,5,6,7,8,9,10,11,12,13] # 合計計算対象行
calc_l_tmp = [3,4,5] # 合計計算対象列
calc_l = [3,4,5]

for i in range(10):# 合計計算対象拡張
	calc_l_tmp = list(map(lambda x: x+6, calc_l_tmp))
	calc_l.extend(calc_l_tmp)

measure_sheet_name = 'ドキュメント指摘計測'
outputfile_name = 'AllReviewAnalysis_Rst.xlsx'

# エクセル作成
new_wb = openpyxl.load_workbook(outputfile_name, data_only=False) # 数式でなく計算結果を取得する

# ドキュメント指摘計測 シート作成
new_measure_sheet = new_wb[measure_sheet_name]

# 合計値代入
for i in calc_r:
	for j in calc_l:
		if (i == 1) or (j % 6 != 5):
			new_measure_sheet.cell(row=i, column=j).value = 9999

# 保存
new_wb.save(outputfile_name)