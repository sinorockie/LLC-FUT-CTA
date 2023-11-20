import datetime
import platform

import numpy as np
import pandas as pd
from matplotlib import pyplot as plt

if platform.system() == 'Windows':
    font = ['Microsoft YaHei']
else:
    font = ['Songti SC']

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
        sheet1['次主力开盘价'] = np.nan
        sheet1['次主力最高价'] = np.nan
        sheet1['次主力最低价'] = np.nan
        sheet1['次主力收盘价'] = np.nan
        sheet1['次主力结算价'] = np.nan
        sheet1['次主力调整比例'] = np.nan
        sheet1['次主力收盘价(调整后)'] = np.nan
        sheet1['次主力最后交易日期'] = None
        for i in range(len(sheet1)):
            # 获取最后交易日期
            row = sheet2[sheet2['合约'] == sheet1.loc[i, '合约']]
            if len(row) == 0:
                continue
            sheet1.loc[i, '合约乘数'] = row['合约乘数'].values[0]
            sheet1.loc[i, '最后交易日期'] = row['最后交易日期'].values[0]
            trade_date = sheet1.loc[i, '日期']
            # 获取该品种该交易日所有合约的行情
            contract = df3[(df3['日期'] == trade_date) & (df3['品种'] == sheet_name)]
            # 合并最后交易日期
            contract = pd.merge(contract, sheet2, on='合约', how='left', validate="many_to_many")
            # 过滤掉最后交易日期(含)之前的合约
            contract = contract[contract['最后交易日期'] > row['最后交易日期'].values[0]]
            if len(contract) == 0:
                continue
            # 取成交量最大的合约
            contract = contract.sort_values(by=['成交量'], ascending=False)
            sheet1.loc[i, '次主力合约'] = contract['合约'].values[0]
            sheet1.loc[i, '次主力开盘价'] = contract['开'].values[0]
            sheet1.loc[i, '次主力最高价'] = contract['高'].values[0]
            sheet1.loc[i, '次主力最低价'] = contract['低'].values[0]
            sheet1.loc[i, '次主力收盘价'] = contract['收'].values[0]
            sheet1.loc[i, '次主力结算价'] = contract['结'].values[0]
            sheet1.loc[i, '次主力最后交易日期'] = contract['最后交易日期'].values[0]
        sheet1['次主力调整比例'] = sheet1['次主力收盘价'].shift() - sheet1['次主力开盘价']
        sheet1['次主力调整比例'] = np.where(sheet1['合约'] != sheet1['合约'].shift(), sheet1['次主力调整比例'], 0)
        sheet1['次主力调整比例'] = sheet1['次主力调整比例'].fillna(0)
        sheet1['次主力收盘价(调整后)'] = sheet1['次主力收盘价'] + sheet1['次主力调整比例'].cumsum()
        sheet1['前主力合约'] = sheet1.shift(1)['合约']
        sheet1['前主力开盘价'] = np.nan
        sheet1['前主力最高价'] = np.nan
        sheet1['前主力最低价'] = np.nan
        sheet1['前主力收盘价'] = np.nan
        sheet1['前主力结算价'] = np.nan
        sheet1['前主力最后交易日期'] = sheet1.shift(1)['最后交易日期']
        for i in range(len(sheet1)):
            row = df3[(df3['日期'] == sheet1.loc[i, '日期']) & (df3['合约'] == sheet1.loc[i, '前主力合约'])]
            if len(row) == 0:
                continue
            sheet1.loc[i, '前主力开盘价'] = row['开'].values[0]
            sheet1.loc[i, '前主力最高价'] = row['高'].values[0]
            sheet1.loc[i, '前主力最低价'] = row['低'].values[0]
            sheet1.loc[i, '前主力收盘价'] = row['收'].values[0]
            sheet1.loc[i, '前主力结算价'] = row['结'].values[0]
        sheet1['成交金额'] = sheet1['收盘价'] * sheet1['成交量'] * sheet1['合约乘数']
        sheet1['成交金额(180日平均)'] = sheet1['成交金额'].rolling(180).mean()
        sheet1['Carry收益'] = sheet1.apply(lambda x: (x['收盘价(调整后)'] / x['次主力收盘价(调整后)'] - 1) * 365 / (x['次主力最后交易日期'] - x['最后交易日期']).astype('timedelta64[D]').astype(int) if x['合约'] is not None and x['次主力合约'] is not None else np.nan, axis=1)
        sheet1['Carry收益(10日平均)'] = sheet1['Carry收益'].rolling(10).mean()
        sheet1['TR'] = 0
        for i in range(1, len(sheet1)):
            sheet1.loc[i, 'TR'] = max(sheet1.loc[i, '最高价'], sheet1.loc[i - 1, '收盘价']) - min(sheet1.loc[i, '最低价'], sheet1.loc[i - 1, '收盘价'])
        sheet1['ATR'] = sheet1['TR'].rolling(10).mean()
        sheet1.to_excel(writer, sheet_name=sheet_name, index=False)
        # 依据日期 画折线图 收盘价 次主力收盘价
        sheet1.plot(x='日期', y=['收盘价', '次主力收盘价'])
        # 保存到图片 需要支持中文
        plt.rcParams['font.sans-serif'] = font
        # x轴标签旋转0度
        plt.xticks(rotation=0)
        # 宽度
        plt.gcf().set_size_inches(20, 10)
        # 保存图片
        plt.savefig(f'../output/image/{sheet_name}.png')
