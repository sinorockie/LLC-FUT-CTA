import datetime

import pandas as pd
# import dataframe_image as dfi

profit_df = pd.read_excel("../output/期货量化实践_Carry收益.xlsx", None)

trade_notional_map = {}
trade_date_series = []

for sheet_name in profit_df:
    sheet = profit_df[sheet_name]
    # 取最后一行
    trade_notional_map[sheet_name] = sheet.iloc[-1]['成交金额(180日平均)']
    trade_date_series = list(sheet['日期'][9:])

trade_notional_map = {k: v for k, v in trade_notional_map.items() if not pd.isna(v) and v >= 5000000000}

available_budget = 100000000

# 创建一个空的 DataFrame 其索引为交易日
multi_index = pd.MultiIndex.from_tuples(
    [
        ('', '当日可用资金'),
        ('', '当日收益'),
        ('', '年化收益率'),
        ('', '夏普比率'),
        ('', 'Calmar Ratio(卡玛比率)')
    ]
)
df = pd.DataFrame(index=trade_date_series, columns=multi_index)

carry_ma10_map = {}
start_date = trade_date_series[0]
adjust_date = trade_date_series[0] + datetime.timedelta(days=30)
budget_map = {}
amount_map = {}
carry_ma10_df = pd.DataFrame()
for i in range(len(trade_date_series)):
    if trade_date_series[i] == start_date:
        for sheet_name in profit_df:
            if sheet_name not in trade_notional_map:
                continue
            sheet = profit_df[sheet_name]
            row = sheet[sheet['日期'] == start_date]
            if len(row) == 0:
                continue
            carry_ma10_map[sheet_name] = row.iloc[0]['Carry收益(10日平均)']
        # 按照Carry收益(10日平均)排序从大到小排序
        carry_ma10_list = sorted(carry_ma10_map.items(), key=lambda x: x[1], reverse=True)
        time_str = start_date.strftime('%Y-%m-%d')
        new_columns = pd.MultiIndex.from_tuples(
            [
                (time_str, '品种'),
                (time_str, 'Carry收益(10日平均)')
            ]
        )
        for column in new_columns:
            carry_ma10_df[column] = None
        carry_ma10_df[(time_str, '品种')] = [item[0] for item in carry_ma10_list]
        carry_ma10_df[(time_str, 'Carry收益(10日平均)')] = [item[1] for item in carry_ma10_list]
        # 取大于0 且从大到小排序的前20%品种
        carry_ma10_map_positive = {k: v for k, v in carry_ma10_map.items() if not pd.isna(v) and v > 0}
        carry_ma10_map_positive = sorted(carry_ma10_map_positive.items(), key=lambda x: x[1], reverse=True)
        carry_ma10_map_positive_20 = carry_ma10_map_positive[:int(len(carry_ma10_map_positive) * 0.2)]
        print(carry_ma10_map_positive_20)
        # 取小于0 且从小到大排序的前20%品种
        carry_ma10_map_negative = {k: v for k, v in carry_ma10_map.items() if not pd.isna(v) and v < 0}
        carry_ma10_map_negative = sorted(carry_ma10_map_negative.items(), key=lambda x: x[1], reverse=False)
        carry_ma10_map_negative_20 = carry_ma10_map_negative[:int(len(carry_ma10_map_negative) * 0.2)]
        print(carry_ma10_map_negative_20)
        each_budget = available_budget / (len(carry_ma10_map_positive_20) + len(carry_ma10_map_negative_20))
        for item in carry_ma10_map_positive_20 + carry_ma10_map_negative_20:
            budget_map[item[0]] = each_budget
            sheet = profit_df[item[0]]
            row = sheet[sheet['日期'] == start_date]
            amount_map[item[0]] = int(each_budget / 4 / 0.15 / row.iloc[0]['策略开仓价'] / row.iloc[0]['合约乘数']) * row.iloc[0]['合约乘数']
            if (item[0], '持仓量') not in df.columns:
                df[(item[0], '持仓量')] = 0.
                df[(item[0], '单位收益')] = 0.
                df[(item[0], '持仓收益')] = 0.
        for key in amount_map.keys():
            sheet = profit_df[key]
            # 取日期在start_date与adjust_date之间的行
            rows = sheet[(sheet['日期'] >= start_date) & (sheet['日期'] <= adjust_date)]
            for index in rows['日期']:
                df.loc[index, (key, '持仓量')] = amount_map[key] if rows[rows['日期'] == index]['最优价格'].values[0] > 0 else 0.
                df.loc[index, (key, '单位收益')] = rows[rows['日期'] == index]['策略收益'].values[0]
                df.loc[index, (key, '持仓收益')] = rows[rows['日期'] == index]['策略收益'].values[0] * amount_map[key]
            # 累计盈亏求和
            profit = rows['策略收益'].sum()
            available_budget += (profit * amount_map[key])
        print(budget_map)
        print(amount_map)
        print(available_budget)
    else:
        if i == len(trade_date_series) - 1:
            break
        else:
            if trade_date_series[i + 1] > adjust_date:
                start_date = trade_date_series[i + 1]
                adjust_date = start_date + datetime.timedelta(days=30)
                budget_map = {}
                amount_map = {}
carry_ma10_df.columns = pd.MultiIndex.from_tuples(carry_ma10_df.columns)
# 删除第一行
carry_ma10_df = carry_ma10_df.drop(index=0)
# 所有Carry收益(10日平均)转换为百分比 小数点后两位
carry_ma10_df[carry_ma10_df.columns[1]] = carry_ma10_df[carry_ma10_df.columns[1]].apply(lambda x: round(x * 100, 2))
carry_ma10_df.to_excel('../output/期货量化实践_Carry收益(10日平均).xlsx', sheet_name='Carry收益(10日平均)')
# 导出图片
# dfi.export(carry_ma10_df, '../output/期货量化实践_Carry收益(10日平均).png', max_cols=-1)

# 当日收益等于所有持仓收益之和
df[('', '当日收益')] = df.xs('持仓收益', axis=1, level=1).sum(axis=1)
# 第一天初始化可用资金为100000000
df.iloc[0, 0] = 100000000
for i in range(len(df)):
    if i == len(df) - 1:
        print()
    if i > 0:
        # 每一行当日可用资金等于前一日可用资金加上前一日收益
        df.iloc[i, 0] = df.iloc[i - 1, 0] + df.iloc[i - 1, 1]
    # 取前252个交易日的数据 不足252个交易日则取所有交易日的数据
    rows = df.iloc[max(0, i - 252):(i+1), :]
    profit = rows.iloc[:, 1].sum()
    budget = rows.iloc[0, 0]
    df.iloc[i, 2] = profit / budget
    # 计算夏普比率 夏普比率=(年化收益率-无风险利率)/年化收益率的标准差
    df.iloc[i, 3] = df.iloc[i, 2] / rows.iloc[:, 2].std()
    # 可用资金最大值
    max_balance = rows.iloc[:, 0].max()
    # 可用资金最大值对应的日期
    max_balance_date = rows[rows.iloc[:, 0] == max_balance].index[0]
    # 在rows数据范围内可用资金最大值后的数据
    rows = rows[rows.index > max_balance_date]
    min_balance = rows.iloc[:, 0].min()
    # 计算Calmar Ratio(卡玛比率) Calmar Ratio(卡玛比率)=年化收益率/最大回撤
    if i > 0:
        df.iloc[i, 4] = df.iloc[i, 2] / (max_balance - min_balance) * max_balance

df.to_excel('../output/期货量化实践_策略收益.xlsx', sheet_name='策略收益')
