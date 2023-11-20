import platform

import pandas as pd
from matplotlib import pyplot as plt

if platform.system() == 'Windows':
    font = ['Microsoft YaHei']
else:
    font = ['Songti SC']

carry_df = pd.read_excel("../output/期货量化实践_Carry收益.xlsx", None)

for image_name in ['TA', 'SS']:
    # 读取sheet名为TA的数据
    sheet = carry_df['SS']
    # 依据日期 画折线图 收盘价 次主力收盘价
    sheet.plot(x='日期', y=['收盘价', '次主力合约收盘价'])
    # 保存到图片 需要支持中文
    plt.rcParams['font.sans-serif'] = font
    # x轴标签旋转0度
    plt.xticks(rotation=0)
    # 宽度
    plt.gcf().set_size_inches(20, 10)
    # 保存图片
    plt.savefig(f'../output/{image_name}.png')

    # 选取2021-11-05至2022-02-07的数据 重新画图
    sheet = sheet[(sheet['日期'] >= '2021-11-05') & (sheet['日期'] <= '2022-02-07')]
    # 依据日期 画折线图 收盘价 次主力收盘价
    sheet.plot(x='日期', y=['收盘价', '次主力合约收盘价'])
    # 2021-12-06至2022-01-04的数据 需要特别标注 且区间背景灰色
    sheet2 = sheet[(sheet['日期'] >= '2021-12-06') & (sheet['日期'] <= '2022-01-04')]
    # 画折线图 收盘价 次主力收盘价
    plt.plot(sheet2['日期'], sheet2['收盘价'], 'r', label='收盘价')
    plt.plot(sheet2['日期'], sheet2['次主力合约收盘价'], 'b', label='次主力合约收盘价')
    # 画区间背景
    plt.axvspan('2021-12-06', '2022-01-04', facecolor='gray', alpha=0.3)
    # 画图例
    plt.legend(loc='upper left')
    # 保存到图片 需要支持中文
    plt.rcParams['font.sans-serif'] = font
    # x轴标签旋转0度
    plt.xticks(rotation=0)
    # 宽度
    plt.gcf().set_size_inches(20, 10)
    # 保存图片
    plt.savefig("../output/SS(区间).png")

# 重置plt
plt.clf()
# 读取策略收益数据 前两行作为表头
profit_df = pd.read_excel("../output/期货量化实践_策略收益.xlsx", sheet_name='策略收益', header=[0, 1])
# 选取第一列 第三列 第四列 生成新的dataframe
profit_df = profit_df.iloc[:, [0, 1, 2, 3]]
# 删除第一行
profit_df = profit_df.drop(index=0)
# 重命名列名
profit_df.columns = ['日期', '当日可用资金', '当日收益', '年化收益率']
# 求累计收益
profit_df['累计收益'] = profit_df['当日收益'].cumsum()
# 将日期列设置为索引列
profit_df = profit_df.set_index('日期')
# 画折线图
plt.plot(profit_df.index, profit_df['当日收益'], 'b', label='当日收益')
plt.plot(profit_df.index, profit_df['累计收益'], 'g', label='累计收益')
plt.xlabel("日期")
plt.legend()
# 画图例
plt.legend(loc='upper left')
# 新增一个图层
ax = plt.twinx()
# 画折线图 截取日期在2021-12-03之后的数据
ax.plot(profit_df.index[profit_df.index >= '2021-12-03'], profit_df['年化收益率'][profit_df.index >= '2021-12-03'], 'r', label='年化收益率')
# 画图例
ax.legend(loc='upper right')
# 保存到图片 需要支持中文
plt.rcParams['font.sans-serif'] = font
# x轴标签旋转0度
plt.xticks(rotation=0)
# 宽度
plt.gcf().set_size_inches(20, 10)
# 保存图片
plt.savefig("../output/策略收益.png")
# 导出excel
profit_df.to_excel("../output/期货量化实践_策略收益(暂存).xlsx")
