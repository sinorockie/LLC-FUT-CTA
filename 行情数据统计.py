# Encoding: UTF-8
import platform

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

excel = pd.read_excel('换月处理后的期货主连合约收盘价数据.xlsx', sheet_name='收盘价数据')

# 第0列是日期 第1列至第50列是合约收盘价
# 依次使用第1列与第n列数据 1<=n<=50
# 计算第n列每日涨跌幅 1<=n<=50
# 将每个合约的最大连续上涨天数和幅度、最大连续下跌天数和幅度一起保存到一个单独的excel中
max_n_min = pd.DataFrame()
max_n_min['合约'] = excel.columns[1:]
max_n_min['最大连续上涨天数'] = np.nan
max_n_min['最大连续上涨幅度'] = np.nan
max_n_min['最大连续下跌天数'] = np.nan
max_n_min['最大连续下跌幅度'] = np.nan
for i in range(1, excel.shape[1]):
    df = pd.DataFrame()
    df['日期'] = excel['date']
    df['收盘价'] = excel.iloc[:, i]
    # 删除收盘价为0的行
    df = df[df['收盘价'] != 0]
    # reset index
    df.reset_index(drop=True, inplace=True)
    df['涨跌幅'] = df['收盘价'].pct_change()
    # 统计收盘价连续上涨天数以及幅度
    df['连续上涨天数'] = np.nan
    df['连续上涨幅度'] = np.nan
    up_days = 0
    up_pct = 0
    for j in range(1, len(df)):
        if df.iloc[j, 2] > 0:
            up_days += 1
            up_pct += df['涨跌幅'].iloc[j]
        else:
            if j > 1:
                df.loc[j-1, '连续上涨天数'] = up_days
                df.loc[j-1, '连续上涨幅度'] = up_pct
            up_days = 0
            up_pct = 0
    df['连续上涨天数'].replace(0, np.nan, inplace=True)
    df['连续上涨幅度'].replace(0, np.nan, inplace=True)
    # 统计收盘价连续下跌天数以及幅度
    df['连续下跌天数'] = np.nan
    df['连续下跌幅度'] = np.nan
    down_days = 0
    down_pct = 0
    for j in range(1, len(df)):
        if df.iloc[j, 2] < 0:
            down_days += 1
            down_pct += df['涨跌幅'].iloc[j]
        else:
            if j > 1:
                df.loc[j-1, '连续下跌天数'] = down_days
                df.loc[j-1, '连续下跌幅度'] = down_pct
            down_days = 0
            down_pct = 0
    df['连续下跌天数'].replace(0, np.nan, inplace=True)
    df['连续下跌幅度'].replace(0, np.nan, inplace=True)
    # 统计收盘价最大连续上涨天数及幅度
    max_up_days = 0
    max_up_pct = 0
    for j in range(1, len(df)):
        if df['连续上涨天数'].iloc[j] > max_up_days:
            max_up_days = df['连续上涨天数'].iloc[j]
            max_up_pct = df['连续上涨幅度'].iloc[j]
    # 统计收盘价最大连续下跌天数及幅度
    max_down_days = 0
    max_down_pct = 0
    for j in range(1, len(df)):
        if df['连续下跌天数'].iloc[j] > max_down_days:
            max_down_days = df['连续下跌天数'].iloc[j]
            max_down_pct = df['连续下跌幅度'].iloc[j]
    # 将最大连续上涨天数及幅度、最大连续下跌天数及幅度保存到max_n_min中
    max_n_min.loc[i-1, '最大连续上涨天数'] = max_up_days
    max_n_min.loc[i-1, '最大连续上涨幅度'] = max_up_pct
    max_n_min.loc[i-1, '最大连续下跌天数'] = max_down_days
    max_n_min.loc[i-1, '最大连续下跌幅度'] = max_down_pct
    # 连续上涨天数及幅度的统计分布数据及图表
    up_days_df = df['连续上涨天数']
    up_pct_df = df['连续上涨幅度']
    up_days_df = up_days_df[up_days_df > 0]
    up_pct_df = up_pct_df[up_pct_df > 0]
    up_days_df = up_days_df.value_counts()
    up_pct_df = up_pct_df.value_counts()
    up_days_df = pd.DataFrame(up_days_df)
    up_pct_df = pd.DataFrame(up_pct_df)
    up_days_df.columns = ['频数']
    up_pct_df.columns = ['频数']
    up_days_df = up_days_df.reset_index()
    up_pct_df = up_pct_df.reset_index()
    up_days_df.columns = ['连续上涨天数', '频数']
    up_pct_df.columns = ['连续上涨幅度', '频数']
    up_days_df = up_days_df.sort_values(by='连续上涨天数')
    up_pct_df = up_pct_df.sort_values(by='连续上涨幅度')
    up_days_df['频率'] = up_days_df['频数'] / up_days_df['频数'].sum()
    up_pct_df['频率'] = up_pct_df['频数'] / up_pct_df['频数'].sum()
    up_days_df['累计频率'] = up_days_df['频率'].cumsum()
    up_pct_df['累计频率'] = up_pct_df['频率'].cumsum()
    up_days_df = up_days_df.reset_index(drop=True)
    up_pct_df = up_pct_df.reset_index(drop=True)
    # 连续下跌天数及幅度的统计分布数据及图表
    down_days_df = df['连续下跌天数']
    down_pct_df = df['连续下跌幅度']
    down_days_df = down_days_df[down_days_df > 0]
    down_pct_df = down_pct_df[down_pct_df < 0]
    down_days_df = down_days_df.value_counts()
    down_pct_df = down_pct_df.value_counts()
    down_days_df = pd.DataFrame(down_days_df)
    down_pct_df = pd.DataFrame(down_pct_df)
    down_days_df.columns = ['频数']
    down_pct_df.columns = ['频数']
    down_days_df = down_days_df.reset_index()
    down_pct_df = down_pct_df.reset_index()
    down_days_df.columns = ['连续下跌天数', '频数']
    down_pct_df.columns = ['连续下跌幅度', '频数']
    down_days_df = down_days_df.sort_values(by='连续下跌天数')
    down_pct_df = down_pct_df.sort_values(by='连续下跌幅度')
    down_days_df['频率'] = down_days_df['频数'] / down_days_df['频数'].sum()
    down_pct_df['频率'] = down_pct_df['频数'] / down_pct_df['频数'].sum()
    down_days_df['累计频率'] = down_days_df['频率'].cumsum()
    down_pct_df['累计频率'] = down_pct_df['频率'].cumsum()
    down_days_df = down_days_df.reset_index(drop=True)
    down_pct_df = down_pct_df.reset_index(drop=True)
    # 5日涨幅=(第5日收盘价-第1日收盘价)/第1日收盘价 5日振幅=(5日内最高价-5日内最低价)/5日内最低价
    df['5日涨跌幅'] = df['收盘价'] / df['收盘价'].shift(4) - 1
    df['5日波幅'] = df['收盘价'].rolling(5).max() / df['收盘价'].rolling(5).min() - 1
    # 10日
    df['10日涨跌幅'] = df['收盘价'] / df['收盘价'].shift(9) - 1
    df['10日波幅'] = df['收盘价'].rolling(10).max() / df['收盘价'].rolling(10).min() - 1
    # 20日
    df['20日涨跌幅'] = df['收盘价'] / df['收盘价'].shift(19) - 1
    df['20日波幅'] = df['收盘价'].rolling(20).max() / df['收盘价'].rolling(20).min() - 1
    # 保存文件
    if True:
        writer = pd.ExcelWriter(f'output/{excel.columns[i]}.xlsx')
        df.to_excel(writer, sheet_name='涨跌幅', index=False)
        up_days_df.to_excel(writer, sheet_name='连续上涨天数统计', index=False)
        down_days_df.to_excel(writer, sheet_name='连续下跌天数统计', index=False)
        writer.close()

if platform.system() == 'Windows':
    font = ['Microsoft YaHei']
else:
    font = ['Songti SC']
# 最大连续上涨天数及幅度 生成柱状图
# x轴合约 y轴最大连续上涨天数
max_n_min.plot.bar(x='合约', y='最大连续上涨天数')
# 保存到图片 需要支持中文
plt.rcParams['font.sans-serif'] = font
# x轴标签旋转0度
plt.xticks(rotation=0)
# 宽度
plt.gcf().set_size_inches(20, 10)
plt.savefig('最大连续上涨天数.png')
# 最大连续上涨幅度 生成柱状图
max_n_min.plot.bar(x='合约', y='最大连续上涨幅度')
# 保存到图片 需要支持中文
plt.rcParams['font.sans-serif'] = font
# x轴标签旋转0度
plt.xticks(rotation=0)
# 宽度
plt.gcf().set_size_inches(20, 10)
plt.savefig('最大连续上涨幅度.png')
# 最大连续下跌天数及幅度 生成柱状图
# x轴合约 y轴最大连续下跌天数
max_n_min.plot.bar(x='合约', y='最大连续下跌天数')
# 保存到图片 需要支持中文
plt.rcParams['font.sans-serif'] = font
# x轴标签旋转0度
plt.xticks(rotation=0)
# 宽度
plt.gcf().set_size_inches(20, 10)
plt.savefig('最大连续下跌天数.png')
# 最大连续下跌幅度 生成柱状图
max_n_min.plot.bar(x='合约', y='最大连续下跌幅度')
# 保存到图片 需要支持中文
plt.rcParams['font.sans-serif'] = font
# x轴标签旋转0度
plt.xticks(rotation=0)
# 宽度
plt.gcf().set_size_inches(20, 10)
plt.savefig('最大连续下跌幅度.png')
# 连续下跌天数及幅度 生成柱状图
max_n_min.to_excel('合约最大连续涨跌天数与幅度汇总.xlsx', index=False)
