import platform

import pandas as pd
from matplotlib import pyplot as plt

if platform.system() == 'Windows':
    font = ['Microsoft YaHei']
else:
    font = ['Songti SC']

# 读取策略收益数据 前两行作为表头
profit_df = pd.read_excel("../output/期货量化实践_策略收益.xlsx", sheet_name='策略收益', header=[0, 1])
# 选取第一列 第三列 第四列 生成新的dataframe
profit_df = profit_df.iloc[:, [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]]
# 删除第一行
profit_df = profit_df.drop(index=0)
# 重命名列名
# profit_df.columns = ['日期', '当日可用资金', '当日收益', '收益率']
profit_df.columns = ['日期', '当日可用资金', '当日收益', '当日可用资金(含当日收益)', '累计收益(365自然日内)', '年化收益率', '最大年化收益率', '年化收益率回撤', '夏普比率', 'Calmar Ratio(卡玛比率)']

# 将日期列转换为日期格式
profit_df['日期'] = profit_df['日期'].apply(lambda x: x.strftime('%Y-%m-%d'))
# 求累计收益
profit_df['累计收益'] = profit_df['当日收益'].cumsum()
# 将日期列设置为索引列
profit_df = profit_df.set_index('日期')

# 画柱状图 显示10个刻度 第一个和最后一个要显示
profit_df['当日收益'].plot.bar()
plt.locator_params(axis='x', nbins=10)
# 小于0的收益柱状图颜色为绿色 大于0的收益柱状图颜色为红色
for i in range(len(profit_df)):
    if profit_df.iloc[i]['当日收益'] < 0:
        plt.bar(i, profit_df.iloc[i]['当日收益'], color='g')
    else:
        plt.bar(i, profit_df.iloc[i]['当日收益'], color='r')
# 保存
plt.rcParams['font.sans-serif'] = font
plt.xticks(rotation=0)
plt.gcf().set_size_inches(20, 10)
plt.savefig("../output/期货量化实践_当日收益.png")

# 重置plt
plt.clf()
# 画折线图 显示10个刻度 第一个和最后一个要显示
profit_df['累计收益'].plot()
plt.locator_params(axis='x', nbins=10)
# 线为红色
plt.plot(profit_df['累计收益'], color='r')
# 保存
plt.rcParams['font.sans-serif'] = font
plt.xticks(rotation=0)
plt.gcf().set_size_inches(20, 10)
plt.savefig("../output/期货量化实践_累计收益.png")

# 数据截取 日期从2022-10-24开始
profit_df = profit_df.loc['2022-10-24':]

# 重置plt
plt.clf()
# 画折线图 显示10个刻度 第一个和最后一个要显示
profit_df['年化收益率'].plot()
plt.locator_params(axis='x', nbins=10)
# 保存
plt.rcParams['font.sans-serif'] = font
plt.xticks(rotation=0)
plt.gcf().set_size_inches(20, 10)
plt.savefig("../output/期货量化实践_年化收益率.png")

# 重置plt
plt.clf()
# 画折现图 显示10个刻度 第一个和最后一个要显示
profit_df['夏普比率'].plot()
profit_df['Calmar Ratio(卡玛比率)'].plot()
plt.gcf().set_size_inches(20, 10)
plt.legend()
# 新增一个图层
ax = plt.twinx()
# 画折线图 截取日期在2021-12-03之后的数据
ax.plot(profit_df['年化收益率'], 'r', label='年化收益率')
# 画图例
ax.legend(loc='upper right')
plt.locator_params(axis='x', nbins=10)
# 保存
plt.rcParams['font.sans-serif'] = font
plt.xticks(rotation=0)
# 画图例
plt.legend(loc='upper left')
plt.savefig("../output/期货量化实践_夏普比率.png")
