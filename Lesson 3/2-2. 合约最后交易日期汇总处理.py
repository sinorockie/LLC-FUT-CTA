import pandas as pd


df_info = pd.read_excel('../output/期货量化实践_合约最后交易日期(暂存).xlsx')
# 删除品种为空的行
df_info = df_info.dropna(subset=['品种'])
# 根据品种分组
df_info = df_info.groupby('品种')
# 保存到excel
with pd.ExcelWriter('../output/期货量化实践_合约最后交易日期.xlsx') as writer:
    for name, group in df_info:
        group = group.sort_values(by=['最后交易日期', '合约'], ascending=False)
        group.to_excel(writer, sheet_name=name, index=False)
