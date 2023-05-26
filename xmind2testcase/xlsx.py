#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import pandas as pd
from xmind2testcase.zentao import xmind_to_zentao_csv_file
def csv_to_xlsx_pd(xmind_file):
    zentao_csv_file = xmind_to_zentao_csv_file(xmind_file)
    save_xlsx_file = zentao_file = xmind_file[:-6] + '.xlsx'
    csv = pd.read_csv(zentao_csv_file, encoding='utf-8', index_col=0)
    csv.to_excel(save_xlsx_file, sheet_name='data')
    return save_xlsx_file

if __name__ == '__main__':
    xmind_file = '../docs/zentao_testcase_template.xmind'
    xlsx_file = csv_to_xlsx_pd(xmind_file)
    print('Conver the xmind file to a xlsx file succssfully: %s', xlsx_file)