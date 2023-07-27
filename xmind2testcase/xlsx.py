#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import pandas as pd
from xmind2testcase.zentao import xmind_to_zentao_csv_file
def csv_to_xlsx_pd(xmind_file):
    zentao_csv_file = xmind_to_zentao_csv_file(xmind_file)
    save_xlsx_file = zentao_file = xmind_file[:-6] + '.xlsx'
    csv = pd.read_csv(zentao_csv_file, encoding='utf-8', index_col=0)
    csv.to_excel(save_xlsx_file, sheet_name='测试用例')

    data_sheet2_columns = ['序列','需求名称(必填)','优先级(必填，可取值：P1、P2、P3、P4）','需求来源（选填）','功能模块（选填）','迭代关联（选填）','描述（选填）']
    data_sheet2 = [['','','','','','','']]
    # 创建一个ExcelWriter对象来写入多个sheet页
    with pd.ExcelWriter(save_xlsx_file, engine='openpyxl', mode='a') as writer:
        # 将数据转换为DataFrame
        df_new_sheet = pd.DataFrame(data_sheet2, columns=data_sheet2_columns)

        # 将DataFrame写入新的sheet页
        df_new_sheet.to_excel(writer, sheet_name='测试需求', index=False)

    data_sheet3_columns = ['序列','功能模块']
    data_sheet3 = [['','']]
    # 创建一个ExcelWriter对象来写入多个sheet页
    with pd.ExcelWriter(save_xlsx_file, engine='openpyxl', mode='a') as writer:
        # 将数据转换为DataFrame
        df_new_sheet = pd.DataFrame(data_sheet3, columns=data_sheet3_columns)
        # 将DataFrame写入新的sheet页
        df_new_sheet.to_excel(writer, sheet_name='功能模块', index=False)
    return save_xlsx_file

if __name__ == '__main__':
    xmind_file = 'E:\LIU\脑图\智能客服1.0_测试产品.xmind'
    xlsx_file = csv_to_xlsx_pd(xmind_file)
    print('Conver the xmind file to a xlsx file succssfully: %s', xlsx_file)