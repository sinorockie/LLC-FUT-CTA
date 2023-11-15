from iFinDPy import *

import numpy as np
import pandas as pd


THS_iFinDLogin('xhrf085', 'db232a')

default_date = datetime.strptime('2023-11-13', '%Y-%m-%d')

df_change = pd.read_excel('../期货量化实践_原始数据_主连数据.xlsx', sheet_name='对应主力')
change = pd.DataFrame(df_change.iloc[1:70, 3:])
change = change.set_index(df_change.iloc[1:70, 0])
change.columns = df_change.iloc[0, 3:]
change = change.T
exchange_map = {}
for column in change.columns:
    exchange_map[column.split('.')[0][:-1].upper()] = column.split('.')[1].upper()

df_multi = pd.read_excel('../期货量化实践_合约乘数.xlsx', sheet_name='主表')

daily_fut_quote_raw_data = pd.read_excel('../期货量化实践_原始数据_日频行情.xlsx')
df_info = pd.DataFrame()
df_info['合约'] = daily_fut_quote_raw_data['合约'].unique()
df_info['品种'] = ''
df_info['合约乘数'] = np.nan
df_info['最后交易日期'] = default_date
for i in range(len(df_info)):
    print(df_info.iloc[i, 0])
    row = daily_fut_quote_raw_data[daily_fut_quote_raw_data['合约'] == df_info.iloc[i, 0]]
    if row['交易所'].values[0] not in ['CZC', 'DCE', 'SHF', 'unknown', 'INE']:
        continue
    wind = row['wind代码'].values[0]
    if row['交易所'].values[0] == 'unknown':
        exchange = exchange_map.get(row['品种'].values[0])
        if exchange is None:
            continue
        wind = wind.replace('unknown', exchange)
    if wind.endswith('.INE'):
        wind = wind.replace('.INE', '.SHF')
    df_info.iloc[i, 1] = row['品种'].values[0]
    multi = df_multi[df_multi['名字'] == df_info.iloc[i, 1]]
    if len(multi) == 0:
        multi = THS_BD(wind, 'ths_contract_multiplier_product_future', '2023-11-13')
        df_info.iloc[i, 2] = multi.data['ths_contract_multiplier_product_future'].values[0]
        df_multi.iloc[len(df_multi)] = ['', df_info.iloc[i, 2], df_info.iloc[i, 1]]
    else:
        df_info.iloc[i, 2] = multi['合约乘数'].values[0]
    if row['日期'].max() == default_date:
        # 去同花顺查询最后交易日期
        last_trading_date = THS_BD(wind, 'ths_last_td_date_future', '2023-11-13')
        df_info.iloc[i, 3] = datetime.strptime(last_trading_date.data['ths_last_td_date_future'].values[0], '%Y%m%d')
    else:
        df_info.iloc[i, 3] = row['日期'].max()
df_multi.to_excel('../output/期货量化实践_合约乘数(暂存).xlsx', sheet_name='主表', index=False)
df_info.to_excel('../output/期货量化实践_合约最后交易日期(暂存).xlsx', index=False)
