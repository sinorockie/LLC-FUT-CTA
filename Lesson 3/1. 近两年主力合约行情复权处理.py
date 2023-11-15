import numpy as np
import pandas as pd

daily_fut_quote_raw_data = pd.read_excel('../期货量化实践_原始数据_日频行情.xlsx')
df_change = pd.read_excel('../期货量化实践_原始数据_主连数据.xlsx', sheet_name='对应主力')
change = pd.DataFrame(df_change.iloc[1:70, 3:])
change = change.set_index(df_change.iloc[1:70, 0])
change.columns = df_change.iloc[0, 3:]
change = change.T
change.columns = change.columns.str.split('.').str[0].str[:-1].str.upper()
change.to_excel('../output/期货量化实践_主力合约(暂存).xlsx', sheet_name='主力合约')
# daily_fut_quote_raw_data 日频期货行情数据 品种 合约 日期 开 收 成交量 持仓量 成交额
# change 交易日主力合约映射
# 向前比例法 保持期货品种上市后第一个合约价格水平不变 将最新的主力合约价格按照缺口比例进行调整
# 参考文章 https://mp.weixin.qq.com/s/9_TKsOJYi7F2rEnAAqcKSw
with pd.ExcelWriter('../output/期货量化实践_主力合约复权价格.xlsx') as writer:
    for column in change.columns:
        print(column)
        df = pd.DataFrame()
        df['日期'] = change[column].index
        df['合约'] = change[column].values
        df['开盘价'] = np.nan
        df['最高价'] = np.nan
        df['最低价'] = np.nan
        df['收盘价'] = np.nan
        df['结算价'] = np.nan
        df['成交量'] = np.nan
        df['持仓量'] = np.nan
        for i in range(len(df)):
            if type(df.loc[i, '合约']) is not str:
                continue
            contract = df.iloc[i, 1]
            if contract.endswith('.CZC'):
                contract = contract.split('.')[0]
                # 删除倒数第四个字符
                contract = contract[:-4] + contract[-3:]
            else:
                contract = contract.split('.')[0]
            row = daily_fut_quote_raw_data[(daily_fut_quote_raw_data['日期'] == df.iloc[i, 0]) & (
                    daily_fut_quote_raw_data['合约'] == contract.upper())]
            if len(row) == 0:
                continue
            df.iloc[i, 1] = row['合约'].values[0]
            df.iloc[i, 2] = row['开'].values[0]
            df.iloc[i, 3] = row['高'].values[0]
            df.iloc[i, 4] = row['低'].values[0]
            df.iloc[i, 5] = row['收'].values[0]
            df.iloc[i, 6] = row['结'].values[0]
            df.iloc[i, 7] = row['成交量'].values[0]
            df.iloc[i, 8] = row['持仓量'].values[0]
        df['调整比例'] = df['收盘价'].shift() / df['开盘价']
        df['调整比例'] = np.where(df['合约'] != df['合约'].shift(), df['调整比例'], 1)
        df['调整比例'] = df['调整比例'].fillna(1)
        df['开盘价(调整后)'] = df['开盘价'] * df['调整比例'].cumprod()
        df['收盘价(调整后)'] = df['收盘价'] * df['调整比例'].cumprod()
        df.to_excel(writer, sheet_name=column)
