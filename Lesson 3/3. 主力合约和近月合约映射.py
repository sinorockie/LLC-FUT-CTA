import datetime

import numpy as np
import pandas as pd

df1 = pd.read_excel("../output/期货量化实践_主力合约复权价格.xlsx", None)
df2 = pd.read_excel("../output/期货量化实践_合约最后交易日期.xlsx", None)
df3 = pd.read_excel("../期货量化实践_原始数据_日频行情.xlsx")

with pd.ExcelWriter('../output/期货量化实践_主力合约复权价格_近月合约价格_合并.xlsx') as writer:
    for sheet_name in df1:
        print(sheet_name)
        sheet1 = df1[sheet_name]
        if sheet_name not in df2:
            continue
        sheet2 = df2[sheet_name]
        sheet1['合约乘数'] = np.nan
        sheet1['最后交易日期'] = datetime.datetime.strptime('2023-11-13', '%Y-%m-%d')
        sheet1['近月合约'] = ''
        sheet1['近月合约开盘价'] = np.nan
        sheet1['近月合约最高价'] = np.nan
        sheet1['近月合约最低价'] = np.nan
        sheet1['近月合约收盘价'] = np.nan
        sheet1['近月合约结算价'] = np.nan
        sheet1['近月最后交易日期'] = datetime.datetime.strptime('2023-11-13', '%Y-%m-%d')
        for i in range(len(sheet1)):
            # 获取最后交易日期
            row = sheet2[sheet2['合约'] == sheet1.iloc[i, 2]]
            if len(row) == 0:
                continue
            sheet1.iloc[i, 13] = row['合约乘数'].values[0]
            sheet1.iloc[i, 14] = row['最后交易日期'].values[0]
            trade_date = sheet1.iloc[i, 1]
            # 获取合约 trade_date之后的第一个合约
            contract = sheet2[sheet2['最后交易日期'] >= trade_date].sort_values(by=['最后交易日期', '合约'], ascending=False)
            row = df3[(df3['日期'] == trade_date) & (df3['合约'] == contract.iloc[-1, 0])]
            if len(row) == 0:
                continue
            sheet1.iloc[i, 15] = contract.iloc[-1, 0]
            sheet1.iloc[i, 16] = row['开'].values[0]
            sheet1.iloc[i, 17] = row['高'].values[0]
            sheet1.iloc[i, 18] = row['低'].values[0]
            sheet1.iloc[i, 19] = row['收'].values[0]
            sheet1.iloc[i, 20] = row['结'].values[0]
            sheet1.iloc[i, 21] = contract.iloc[-1, 3]
        sheet1['前主力合约'] = sheet1.shift(1)['合约']
        sheet1['前主力合约开盘价'] = np.nan
        sheet1['前主力合约最高价'] = np.nan
        sheet1['前主力合约最低价'] = np.nan
        sheet1['前主力合约收盘价'] = np.nan
        sheet1['前主力合约结算价'] = np.nan
        sheet1['前主力最后交易日期'] = datetime.datetime.strptime('2023-11-13', '%Y-%m-%d')
        for i in range(len(sheet1)):
            row = df3[(df3['日期'] == trade_date) & (df3['合约'] == sheet1.iloc[i, 22])]
            if len(row) == 0:
                continue
            sheet1.iloc[i, 23] = row['开'].values[0]
            sheet1.iloc[i, 24] = row['高'].values[0]
            sheet1.iloc[i, 25] = row['低'].values[0]
            sheet1.iloc[i, 26] = row['收'].values[0]
            sheet1.iloc[i, 27] = row['结'].values[0]
        sheet1.to_excel(writer, sheet_name=sheet_name, index=False)
