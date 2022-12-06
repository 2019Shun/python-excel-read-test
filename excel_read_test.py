# Excelファイルのread性能を評価するスクリプト
import time
import openpyxl
import xlwings
import pylightxl
import zipfile
from lxml import etree
import pandas as pd
import numpy as np

# テストファイルの切り替え
test_file_name = 'test_file_1.xlsx'     # 軽量ファイル
# test_file_name = 'test_file_2.xlsx'   # 重量ファイル

test_sheet_name = 'test_sheet'

# 実行時間計測用
def print_proc_time(f):
    def print_proc_time_func(*args, **kwargs):
        time_list = []
        for i in range(100):
            start_time = time.perf_counter()
            return_val = f(*args, **kwargs)
            end_time = time.perf_counter()
            elapsed_time = end_time - start_time
            time_list.append(elapsed_time)
        print(f.__name__, max(time_list), min(time_list), sum(time_list)/len(time_list))
        return return_val
    return print_proc_time_func

@print_proc_time
def read_test_openpyxl():
    wb = openpyxl.load_workbook(test_file_name)
    sheet = wb[test_sheet_name]

    tmp_list = []
    for col in sheet.iter_cols():
        for raw_value in col:
            tmp_list.append(raw_value.value)
    wb.close()

@print_proc_time
def read_test_xlwings():
    xlwings.App(visible=False)
    wb = xlwings.Book(test_file_name)
    sheet = wb.sheets[test_sheet_name]
    rng = sheet.used_range

    tmp_list = []
    for c in rng.columns:
        for v in c.value:
            tmp_list.append(v)
    wb.close()

@print_proc_time
def read_test_pylightxl():
    db = pylightxl.readxl(fn=test_file_name)

    tmp_list = []
    for col in db.ws(ws=test_sheet_name).cols:
        for v in col:
            tmp_list.append(v)

@print_proc_time
def read_test_xml():
    with zipfile.ZipFile(test_file_name, 'r') as zip_data:
        file_data = zip_data.read('xl/worksheets/sheet1.xml')
    root = etree.XML(file_data)

    tmp_list = []
    for v in root.findall('./{*}sheetData/{*}row/{*}c/{*}v'):
        tmp_list.append(v.text)

@print_proc_time
def read_test_pandas():
    df = pd.read_excel(test_file_name, header=None, index_col=None)
    itr = np.nditer(df.values)
    tmp_list = []
    for v in itr:
        tmp_list.append(v)

if __name__ == '__main__':
    read_test_openpyxl()
    read_test_xlwings()
    read_test_pylightxl()
    read_test_xml()
    read_test_pandas()