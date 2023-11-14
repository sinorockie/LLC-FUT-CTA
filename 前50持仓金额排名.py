import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

df_volume = pd.read_excel('期货量化实践_持仓前50主连数据.xlsx', sheet_name='20231023持仓量（单边）')

# 取第2列和第4列并重命名为'合约'和'持仓量' 并忽略第一行的数据
df_volume = df_volume.iloc[1:, [1, 3]]
df_volume.columns = ['主力合约', '持仓量']
# 第一列的合约代码删除'.'之后的部分
df_volume['主力合约'] = df_volume['主力合约'].str.split('.').str[0]
# 删除最后一个字母并转大写
df_volume['主力合约'] = df_volume['主力合约'].str[:-1].str.upper()

df_price = pd.read_excel('期货量化实践_持仓前50主连数据.xlsx', sheet_name='近2年主力连续')
df_close = pd.DataFrame(df_price.iloc[3:, 3::3])
# 使用第二列的日期作为index
df_close.index = df_price.iloc[3:, 1]
df_close.columns = df_price.iloc[0, 3::3]
# column重命名 删除'.'之后的部分 并删除对后一个字母 并转大写
df_close.columns = df_close.columns.str.split('.').str[0].str[:-1].str.upper()

df_multi = pd.read_excel('合约乘数.xlsx', sheet_name='主表')
# 仅保留名字和合约乘数两列 并去重
df_multi = df_multi[['名字', '合约乘数']].drop_duplicates()
# 合并合约乘数
df = pd.merge(df_volume, df_multi, left_on='主力合约', right_on='名字', how='left')[['主力合约', '持仓量', '合约乘数']]

# 取2023-10-23的行情并与持仓量合并
df = pd.merge(df, df_close.loc['2023-10-23'], left_on='主力合约', right_index=True, how='left')
# 第四列重命名为'收盘价'
df.rename(columns={df.columns[3]: '收盘价'}, inplace=True)
# 计算持仓金额
df['持仓金额'] = df['持仓量'] * df['收盘价'] * df['合约乘数']
# 并排序
df.sort_values(by='持仓金额', ascending=False, inplace=True)
# 保留前50
df = df.iloc[:50, :]
# 重置index
df.reset_index(drop=True, inplace=True)
# 保存到excel
df.to_excel('持仓金额前50主力合约.xlsx', sheet_name='持仓金额前50主力合约', index=False)
# 根据x轴为合约 y轴为持仓金额绘制柱状图
df.plot.bar(x='主力合约', y='持仓金额')
# 保存到图片 需要支持中文
plt.rcParams['font.sans-serif'] = ['Songti SC']
# x轴标签旋转45度
plt.xticks(rotation=0)
# 宽度
plt.gcf().set_size_inches(20, 10)
plt.savefig('持仓金额前50主力合约.png')
print()
