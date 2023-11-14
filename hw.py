import datetime

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

file_name_1 = '期货量化实践_五大期货交易所数据.xlsx'


def filter_data():
    # 品种, 合约, 日期, 高, 开, 低, 收, 结, 昨收, 昨结, 持仓量, 成交量, 收益率, 交易所, 日期数字, wind代码, std代码
    df = pd.read_excel('daily_fut_quote_raw_data.xlsx')
    print(df.size)

    # 仅保留交易所为CZC DCE SHF INF GFE unknown的数据
    df = df[(df['交易所'] == 'CZC') | (df['交易所'] == 'DCE') | (df['交易所'] == 'SHF') | (df['交易所'] == 'INF') | (
            df['交易所'] == 'GFE') | (df['交易所'] == 'unknown')]
    print(df.size)

    # 忽略品种 ['IC','IF','TF','T','IH','BB','FB','RR','WH','PM','RI','LR','JR','RS','WR']
    df = df[(df['品种'] != 'IC') & (df['品种'] != 'IF') & (df['品种'] != 'TF') & (df['品种'] != 'T') & (
            df['品种'] != 'IH') & (df['品种'] != 'BB') & (df['品种'] != 'FB') & (df['品种'] != 'RR') & (
                    df['品种'] != 'WH') & (df['品种'] != 'PM') & (df['品种'] != 'RI') & (df['品种'] != 'LR') & (
                    df['品种'] != 'JR') & (df['品种'] != 'RS') & (df['品种'] != 'WR')]
    # 保存数据文件
    df.to_excel(file_name_1, index=False)


file_name_2 = 'tmp/五大期货交易所起始日期和结束日期.xlsx'


def get_start_end_date():
    # 读取数据文件
    df = pd.read_excel(file_name_1)
    # 根据合约分组 统计每个合约的起始日期和结束日期 格式为yyyy-MM-dd 保存到一个单独的excel中
    start_end_date = pd.DataFrame()
    start_end_date['合约'] = df['合约'].unique()
    start_end_date['品种'] = ''
    start_end_date['起始日期'] = datetime.date.today()
    start_end_date['结束日期'] = datetime.date.today()
    for i in range(len(start_end_date)):
        start_end_date.loc[i, '品种'] = df[df['合约'] == start_end_date.loc[i, '合约']]['品种'].unique()[0]
        start_end_date.loc[i, '起始日期'] = df[df['合约'] == start_end_date.loc[i, '合约']]['日期'].min()
        start_end_date.loc[i, '结束日期'] = df[df['合约'] == start_end_date.loc[i, '合约']]['日期'].max()
    # 根据合约排序
    start_end_date.sort_values(by='合约', inplace=True)
    # 保存数据文件
    start_end_date.to_excel(file_name_2, index=False)


if __name__ == '__main__':
    daily_fut_quote_raw_data = pd.read_excel('daily_fut_quote_raw_data.xlsx')
    df_change = pd.read_excel('期货量化实践_持仓前50主连数据.xlsx', sheet_name='对应主力')
    change = pd.DataFrame(df_change.iloc[1:70, 3:])
    change = change.set_index(df_change.iloc[1:70, 0])
    change.columns = df_change.iloc[0, 3:]
    change = change.T
    change.columns = change.columns.str.split('.').str[0].str[:-1].str.upper()
    change.to_excel('期货量化实践_主力合约.xlsx', sheet_name='主力合约')
    # daily_fut_quote_raw_data 为具体合约的数据
    # change 为具体交易日的主力合约
    # 向前比例法 保持期货品种上市后第一个合约价格水平不变 将最新的主力合约价格按照缺口比例进行调整
    for column in change.columns:
        print(column)
        df = pd.DataFrame()
        df['日期'] = change[column].index
        df['合约'] = change[column].values
        df['开盘价'] = np.nan
        df['收盘价'] = np.nan
        for i in range(len(df)):
            if type(df.loc[i, '合约']) is not str:
                continue
            row = daily_fut_quote_raw_data[(daily_fut_quote_raw_data['日期'] == df.iloc[i, '日期']) & (
                    daily_fut_quote_raw_data['合约'] == df.iloc[i, '合约'].split('.')[0].upper())]
            df.iloc[i, '开盘价'] = row['开'].values[0]
            df.iloc[i, '收盘价'] = row['收'].values[0]
        df['调整比例'] = df['收盘价'].shift() / df['开盘价']
        df['调整比例'] = np.where(df['合约'] != df['合约'].shift(), df['调整比例'], 1)
        df['调整比例'] = df['调整比例'].fillna(1)
        df['开盘价(调整后)'] = df['开盘价'] * df['调整比例'].cumprod()
        df['收盘价(调整后)'] = df['收盘价'] * df['调整比例'].cumprod()
        df.to_excel(f'期货量化实践_向前比例法{column}.xlsx', sheet_name=column)
