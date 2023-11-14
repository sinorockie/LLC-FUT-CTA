import numpy as np
import pandas as pd

daily_fut_quote_raw_data = pd.read_excel('../期货量化实践_原始数据_日频行情.xlsx')
df_info = pd.DataFrame()
df_info['品种'] = ''
df_info['合约'] = daily_fut_quote_raw_data['合约'].unique()
df_info['上市日期'] = np.nan
df_info['最后交易日期'] = np.nan
for i in range(len(df_info)):
    print(df_info.iloc[i, 1])
    row = daily_fut_quote_raw_data[daily_fut_quote_raw_data['合约'] == df_info.iloc[i, 1]]
    df_info.iloc[i, 0] = row['品种'].values[0]
    df_info.iloc[i, 2] = row['日期'].min()
    df_info.iloc[i, 3] = row['日期'].max()
# 根据品种分组
df_info = df_info.groupby('品种')
# 保存到excel
with pd.ExcelWriter('../output/期货量化实践_合约最后交易日期.xlsx') as writer:
    for name, group in df_info:
        group.to_excel(writer, sheet_name=name.__str__(), index=False)
