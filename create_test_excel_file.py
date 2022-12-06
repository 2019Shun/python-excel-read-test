# テスト用のエクセルファイルを作成
# 100x100または1000×1000の乱数データを作成して格納
import openpyxl
import random
random.seed(0)

save_file_name = 'test_file_1.xlsx'     # 軽量ファイル用
iter_num = 100                          # 軽量ファイル用

# save_file_name = 'test_file_2.xlsx'   # 重量ファイル用
# iter_num = 1000                       # 重量ファイル用

sheet_name = 'test_sheet'

wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = sheet_name

for i in range(iter_num):
    for j in range(iter_num):
        cell = sheet.cell(row=i+1, column=j+1)
        cell.value = random.randint(0,1000)

wb.save(save_file_name)
wb.close()