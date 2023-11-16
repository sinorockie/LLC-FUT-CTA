import datetime

import numpy as np
import pandas as pd

df1 = pd.read_excel("../output/期货量化实践_主力合约复权价格.xlsx", None)
df2 = pd.read_excel("../output/期货量化实践_合约最后交易日期.xlsx", None)
df3 = pd.read_excel("../期货量化实践_原始数据_日频行情.xlsx")

with pd.ExcelWriter('../output/期货量化实践_主力合约复权价格_次主力合约价格_合并.xlsx') as writer:
    for sheet_name in df1:
        print(sheet_name)
        sheet1 = df1[sheet_name]
        if sheet_name not in df2:
            continue
        sheet2 = df2[sheet_name]
        sheet1['合约乘数'] = np.nan
        sheet1['最后交易日期'] = None
        sheet1['次主力合约'] = None
        sheet1['次主力合约开盘价'] = np.nan
        sheet1['次主力合约最高价'] = np.nan
        sheet1['次主力合约最低价'] = np.nan
        sheet1['次主力合约收盘价'] = np.nan
        sheet1['次主力合约结算价'] = np.nan
        sheet1['次主力最后交易日期'] = None
        for i in range(len(sheet1)):
            # 获取最后交易日期
            row = sheet2[sheet2['合约'] == sheet1.iloc[i, 2]]
            if len(row) == 0:
                continue
            sheet1.iloc[i, 13] = row['合约乘数'].values[0]
            sheet1.iloc[i, 14] = row['最后交易日期'].values[0]
            trade_date = sheet1.iloc[i, 1]
            # 获取该品种该交易日所有合约的行情
            contract = df3[(df3['日期'] == trade_date) & (df3['品种'] == sheet_name)]
            # 合并最后交易日期

            contract = pd.merge(contract, sheet2, on='合约', how='left')
            # 过滤掉最后交易日期(含)之前的合约
            contract = contract[contract['最后交易日期'] > row['最后交易日期'].values[0]]
            if len(contract) == 0:
                continue
            # 取成交量最大的合约
            contract = contract.sort_values(by=['成交量'], ascending=False)
            sheet1.iloc[i, 15] = contract['合约'].values[0]
            sheet1.iloc[i, 16] = contract['开'].values[0]
            sheet1.iloc[i, 17] = contract['高'].values[0]
            sheet1.iloc[i, 18] = contract['低'].values[0]
            sheet1.iloc[i, 19] = contract['收'].values[0]
            sheet1.iloc[i, 20] = contract['结'].values[0]
            sheet1.iloc[i, 21] = contract['最后交易日期'].values[0]
        sheet1['前主力合约'] = sheet1.shift(1)['合约']
        sheet1['前主力合约开盘价'] = np.nan
        sheet1['前主力合约最高价'] = np.nan
        sheet1['前主力合约最低价'] = np.nan
        sheet1['前主力合约收盘价'] = np.nan
        sheet1['前主力合约结算价'] = np.nan
        sheet1['前主力最后交易日期'] = sheet1.shift(1)['最后交易日期']
        for i in range(len(sheet1)):
            row = df3[(df3['日期'] == sheet1.iloc[i, 1]) & (df3['合约'] == sheet1.iloc[i, 22])]
            if len(row) == 0:
                continue
            sheet1.iloc[i, 23] = row['开'].values[0]
            sheet1.iloc[i, 24] = row['高'].values[0]
            sheet1.iloc[i, 25] = row['低'].values[0]
            sheet1.iloc[i, 26] = row['收'].values[0]
            sheet1.iloc[i, 27] = row['结'].values[0]
        sheet1['Carry收益'] = sheet1.apply(lambda x: (x['主力合约收盘价'] / x['次主力合约收盘价'] - 1) * 365 / (x['次主力最后交易日期'] - x['最后交易日期']), axis=1)
        sheet1.to_excel(writer, sheet_name=sheet_name, index=False)
