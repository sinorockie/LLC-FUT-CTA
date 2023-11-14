import numpy as np
import pandas as pd

df1 = pd.read_excel("../output/期货量化实践_主力合约复权价格.xlsx", None)
df2 = pd.read_excel("../output/期货量化实践_合约最后交易日期.xlsx", None)
df3 = pd.read_excel("../期货量化实践_原始数据_日频行情.xlsx")

with pd.ExcelWriter('../output/期货量化实践_主力合约复权价格_近月合约价格_合并.xlsx') as writer:
    for sheet_name in df1:
        print(sheet_name)
        sheet1 = df1[sheet_name]
        sheet2 = df2[sheet_name]
        sheet1['近月合约'] = ''
        sheet1['近月合约开盘价'] = np.nan
        sheet1['近月合约收盘价'] = np.nan
        for i in range(len(sheet1)):
            trade_date = sheet1.iloc[i, 1]
            # 获取合约 trade_date之后的第一个合约
            contract = sheet2[sheet2['最后交易日期'] >= trade_date].sort_values(by=['最后交易日期', '合约'], ascending=False).iloc[-1, 1]
            row = df3[(df3['日期'] == trade_date) & (df3['合约'] == contract)]
            if len(row) == 0:
                continue
            sheet1.iloc[i, 8] = contract
            sheet1.iloc[i, 9] = row['开'].values[0]
            sheet1.iloc[i, 10] = row['收'].values[0]
        sheet1.to_excel(writer, sheet_name=sheet_name, index=False)
