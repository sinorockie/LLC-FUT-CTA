import pandas as pd
import platform

from matplotlib import pyplot as plt

average_days = 10

carry_df = pd.read_excel('../output/期货量化实践_主力合约复权价格_次主力合约价格_合并.xlsx', None)

with pd.ExcelWriter('../output/期货量化实践_每日单位收益.xlsx') as writer:
    for sheet_name in carry_df:
        print(f'当前品种:{sheet_name}')
        # 获取当前品种数据
        sheet = carry_df[sheet_name]
        # 计算成交量成交金额(10日平均)
        sheet['成交金额(10日平均)'] = sheet['成交金额'].rolling(average_days).mean()
        # 当成交量以及持仓量均在 10 日均线之下则为缩量状态
        sheet['成交量(10日平均)'] = sheet['成交量'].rolling(average_days).mean()
        sheet['持仓量(10日平均)'] = sheet['持仓量'].rolling(average_days).mean()
        sheet['缩量信号'] = sheet.apply(lambda x: 1 if x['成交量(10日平均)'] > x['成交量'] and x['持仓量(10日平均)'] > x['持仓量'] else 0, axis=1)
        # 计算 ATR
        sheet['TR1'] = sheet['最高价'] - sheet['最低价']
        sheet['TR2'] = abs(sheet['最高价'] - sheet.shift(1)['收盘价'])
        sheet['TR3'] = abs(sheet['最低价'] - sheet.shift(1)['收盘价'])
        sheet['TR'] = sheet[['TR1', 'TR2', 'TR3']].max(axis=1)
        sheet.drop(['TR1', 'TR2', 'TR3'], axis=1, inplace=True)
        sheet['ATR'] = sheet['TR'].rolling(average_days).mean()
        # 计算Carry收益(10日平均)
        sheet['Carry收益(10日平均)'] = sheet['Carry收益'].rolling(average_days).mean()
        # 计算收盘价(10日平均)
        sheet['收盘价(10日平均)'] = sheet['收盘价'].rolling(average_days).mean()
        # 计算开仓信号
        # 当 Carry 为正（负）且主力合约收盘价在均线下（上）方时，满仓开仓。
        sheet['开仓信号'] = sheet.apply(lambda x: 1 if (x['Carry收益(10日平均)'] > 0 and x['收盘价(10日平均)'] > x['收盘价']) or (x['Carry收益(10日平均)'] < 0 and x['收盘价(10日平均)'] < x['收盘价']) else 0, axis=1)
        # 当 Carry 为正（负）且主力合约收盘价在均线上（下）方时，若缩量，半仓开仓
        sheet['开仓信号'] = sheet.apply(lambda x: 0.5 if (x['Carry收益(10日平均)'] > 0 and x['收盘价(10日平均)'] < x['收盘价']) or (x['Carry收益(10日平均)'] < 0 and x['收盘价(10日平均)'] > x['收盘价']) and x['缩量信号'] == 1 else x['开仓信号'], axis=1)
        sheet['最优价格'] = 0.
        sheet['策略开仓价'] = 0.
        sheet['策略结算价'] = 0.
        sheet['策略收益'] = 0.
        for i in range(average_days, len(sheet)):
            if sheet.loc[i - 1, 'Carry收益(10日平均)'] > 0:
                # 卖空
                sheet.loc[i, '最优价格'] = sheet.loc[i, '最低价']
                sheet.loc[i, '策略开仓价'] = sheet.loc[i, '开盘价']
                margin_price = sheet.loc[i, '最优价格'] + 2.5 * sheet.loc[i, 'ATR']
                sheet.loc[i, '策略结算价'] = margin_price if margin_price < sheet.loc[i, '收盘价'] else sheet.loc[i, '收盘价']
                sheet.loc[i, '策略收益'] = sheet.loc[i, '策略开仓价'] - sheet.loc[i, '策略结算价'] - (sheet.loc[i, '策略开仓价'] + sheet.loc[i, '策略结算价']) * 2 / 100 / 100
            else:
                # 买多
                sheet.loc[i, '最优价格'] = sheet.loc[i, '最高价']
                sheet.loc[i, '策略开仓价'] = sheet.loc[i, '开盘价']
                margin_price = sheet.loc[i, '最优价格'] - 2.5 * sheet.loc[i, 'ATR']
                sheet.loc[i, '策略结算价'] = margin_price if margin_price > sheet.loc[i, '收盘价'] else sheet.loc[i, '收盘价']
                sheet.loc[i, '策略收益'] = sheet.loc[i, '策略结算价'] - sheet.loc[i, '策略开仓价'] - (sheet.loc[i, '策略开仓价'] + sheet.loc[i, '策略结算价']) * 2 / 100 / 100
        sheet.to_excel(writer, sheet_name=sheet_name, index=False)

profit_df = pd.read_excel("../output/期货量化实践_每日单位收益.xlsx", None)

trade_notional_map = {}
trade_date_series = []

for sheet_name in profit_df:
    sheet = profit_df[sheet_name]
    # 取最后一行
    trade_notional_map[sheet_name] = sheet.iloc[average_days - 1]['成交金额(10日平均)']
    trade_date_series = list(sheet['日期'][average_days:len(sheet) - 1])

# 将dict数据转换为dataframe 日期为索引列
trade_notional_df = pd.DataFrame.from_dict(trade_notional_map, orient='index', columns=['成交金额(10日平均)'])
trade_notional_df.index.name = '品种'
# 按照成交金额(10日平均)排序从大到小排序
trade_notional_df.fillna(0, inplace=True)
trade_notional_df = trade_notional_df.sort_values(by=['成交金额(10日平均)'], ascending=False)
trade_notional_df.plot.bar()
# 大于等于50亿画一条横线
plt.axhline(y=5000000000, color='r', linestyle='-')
# 小于50亿的品种背景灰色
for i in range(len(trade_notional_df)):
    if trade_notional_df.iloc[i]['成交金额(10日平均)'] < 5000000000:
        plt.axvspan(i - 0.5, i + 0.5, facecolor='gray', alpha=0.3)
# 导出图片
if platform.system() == 'Windows':
    font = ['Microsoft YaHei']
else:
    font = ['Songti SC']
plt.rcParams['font.sans-serif'] = font
# x轴标签旋转0度
plt.xticks(rotation=0)
plt.xlabel('')
# 宽度
plt.gcf().set_size_inches(18, 7)
# 保存图片
plt.savefig("../output/期货量化实践_成交金额(10日平均).png")

available_budget = 5000000

# 创建一个空的 DataFrame 其索引为交易日
multi_index = pd.MultiIndex.from_tuples(
    [
        ('', '当日可用资金'),
        ('', '当日收益'),
        ('', '当日可用资金(含当日收益)'),
        ('', '累计收益(365自然日内)'),
        ('', '年化收益率'),
        ('', '最大年化收益率'),
        ('', '年化收益率回撤'),
        ('', '夏普比率'),
        ('', 'Calmar Ratio(卡玛比率)')
    ]
)
df = pd.DataFrame(index=trade_date_series, columns=multi_index)

for trade_date in trade_date_series:
    trade_notional_map = {}
    carry_ma10_map = {}
    budget_map = {}
    amount_map = {}
    # 获取每个品种的成交金额(10日平均)
    for sheet_name in profit_df:
        sheet = profit_df[sheet_name]
        trade_notional_map[sheet_name] = sheet[sheet['日期'] == trade_date]['成交金额(10日平均)'].values[0]
    # 过滤掉成交金额(10日平均)小于50亿的品种
    trade_notional_map = {k: v for k, v in trade_notional_map.items() if v >= 5000000000}
    for sheet_name in profit_df:
        # ZC后期停止交易 剔除
        if sheet_name not in trade_notional_map or sheet_name == 'ZC':
            continue
        sheet = profit_df[sheet_name]
        row = sheet[sheet['日期'] == trade_date]
        carry_ma10_map[sheet_name] = row.iloc[0]['Carry收益(10日平均)']
    print(f'当前日期: {trade_date.strftime("%Y-%m-%d")}')
    # 从大到小排序的最大前20%品种
    positive_carry_ma10_map = {k: v for k, v in carry_ma10_map.items() if v > 0.}
    positive_carry_ma10_map_list = sorted(positive_carry_ma10_map.items(), key=lambda x: x[1], reverse=True)
    carry_ma10_map_first_20 = positive_carry_ma10_map_list[:int(len(positive_carry_ma10_map_list) * 0.20)]
    print(carry_ma10_map_first_20)
    # 从小到大排序的最小前20%品种
    negative_carry_ma10_map = {k: v for k, v in carry_ma10_map.items() if v < 0.}
    negative_carry_ma10_map_list = sorted(negative_carry_ma10_map.items(), key=lambda x: x[1], reverse=True)
    carry_ma10_map_last_20 = negative_carry_ma10_map_list[-int(len(negative_carry_ma10_map_list) * 0.20):]
    print(carry_ma10_map_last_20)
    # 过滤掉开仓信号为0的品种
    for item in carry_ma10_map_first_20 + carry_ma10_map_last_20:
        sheet = profit_df[item[0]]
        row = sheet[sheet['日期'] == trade_date]
        if len(row) == 0 or row.iloc[0]['开仓信号'] == 0:
            continue
        amount_map[item[0]] = 0
    each_budget = available_budget / len(amount_map)
    # 每个品种的持仓量
    for k in amount_map.keys():
        sheet = profit_df[k]
        row = sheet[sheet['日期'] == trade_date]
        amount = int(each_budget / 4 / 0.15 / row['策略开仓价'].values[0] / row['合约乘数'].values[0])
        limit = min(5, int(row['成交量'].values[0] * 1 / 100000))
        if amount > limit:
            amount = 1 if limit == 0 else limit
        amount_map[k] = amount * row['合约乘数'].values[0]
        if (k, '持仓量') not in df.columns:
            df[(k, '持仓量')] = 0.
            df[(k, '单位收益')] = 0.
            df[(k, '持仓收益')] = 0.
        budget_map[k] = each_budget
        row_index = sheet[sheet['日期'] == trade_date].index[0]
        amount = amount_map[k] * sheet.iloc[row_index]['开仓信号']
        unit_profit = sheet.iloc[row_index + 1]['策略收益']
        total_profit = unit_profit * amount
        next_date = sheet.iloc[row_index + 1]['日期']
        df.loc[next_date, (k, '持仓量')] = amount
        df.loc[next_date, (k, '单位收益')] = unit_profit
        df.loc[next_date, (k, '持仓收益')] = total_profit
        available_budget += total_profit
    print(budget_map)
    print(amount_map)
    print(available_budget)

# 当日收益等于所有持仓收益之和
df[('', '当日收益')] = df.xs('持仓收益', axis=1, level=1).sum(axis=1)
# 第一天初始化可用资金为100000000
df.iloc[0, 0] = 5000000

for i in range(1, len(df)):
    df.iloc[i, 0] = df.iloc[i - 1, 0] + df.iloc[i - 1, 1]
    df.iloc[i, 2] = df.iloc[i, 0] + df.iloc[i, 1]
    min_index = 1
    # 获取时间序列里 小于当前日期-365天的日期里最小的一天
    min_date = df.index[(df.index > df.index[0]) & (df.index > df.index[i] - pd.Timedelta(days=365)) & (df.index < df.index[i])]
    if len(min_date) > 0:
        min_index = df.index.get_loc(min_date[0])
    # 区间数据
    rows = df.iloc[min_index:(i + 1), :]
    # 累计收益(365自然日内)
    df.iloc[i, 3] = rows.iloc[:, 1].sum()
    # 区间年化收益率
    df.iloc[i, 4] = df.iloc[i, 3] / rows.iloc[0, 0] * 252 / (i - min_index + 1)
    # 区间最大年化收益率
    df.iloc[i, 5] = rows.iloc[:, 4].max()
    # 区间最大回撤
    max_index = rows.iloc[:, 4].idxmax()
    df.iloc[i, 6] = df.iloc[i, 5] - rows[rows.index > max_index].iloc[:, 4].min()
    # 夏普比率
    df.iloc[i, 7] = (rows.iloc[:, 4].mean() - 0.015) / rows.iloc[:, 4].std()
    # 卡玛比率
    df.iloc[i, 8] = df.iloc[i, 5] / df.iloc[i, 6]

df.to_excel('../output/期货量化实践_策略收益.xlsx', sheet_name='策略收益')
