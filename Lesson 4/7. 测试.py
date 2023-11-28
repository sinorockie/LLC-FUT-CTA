import datetime
import platform

import numpy as np
import pandas as pd
from matplotlib import pyplot as plt

result_map = {}
average_days = 10
period_days = 21
def test_loop():
    carry_df = pd.read_excel("../output/期货量化实践_主力合约复权价格_次主力合约价格_合并.xlsx", None)

    with pd.ExcelWriter('../output/期货量化实践_Carry收益.xlsx') as writer:
        for sheet_name in carry_df:
            print(sheet_name)
            sheet = carry_df[sheet_name]
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
            sheet['Carry收益(10日平均)'] = sheet['Carry收益'].rolling(average_days).mean()
            sheet['收盘价(10日平均)'] = sheet['收盘价'].rolling(average_days).mean()
            # 计算开仓信号
            # 当 Carry 为正（负）且主力合约收盘价在均线下（上）方时，满仓开仓。
            sheet['开仓信号'] = sheet.apply(lambda x: 1 if (x['Carry收益(10日平均)'] > 0 and x['收盘价(10日平均)'] > x['收盘价']) or (x['Carry收益(10日平均)'] < 0 and x['收盘价(10日平均)'] < x['收盘价']) else 0, axis=1)
            # 当 Carry 为正（负）且主力合约收盘价在均线上（下）方时，若缩量，半仓开仓
            sheet['开仓信号'] = sheet.apply(lambda x: 0.5 if (x['Carry收益(10日平均)'] > 0 and x['收盘价(10日平均)'] < x['收盘价']) or (x['Carry收益(10日平均)'] < 0 and x['收盘价(10日平均)'] > x['收盘价']) and x['缩量信号'] == 1 else x['开仓信号'], axis=1)
            sheet['最优价格'] = 0.
            sheet['平仓信号'] = 0.
            sheet['策略开仓方向'] = 0.
            sheet['策略开仓价'] = 0.
            sheet['策略昨日结算价'] = 0.
            sheet['策略结算价'] = 0.
            sheet['策略收益'] = 0.
            sheet['移仓开仓价'] = 0.
            sheet['移仓结算价'] = 0.
            sheet['移仓收益'] = 0.
            start_date = sheet.loc[average_days - 1, '日期']
            adjust_date = sheet.loc[average_days - 1 + period_days, '日期']
            for i in range(len(sheet)):
                if sheet.loc[i, '日期'] < start_date:
                    continue
                if sheet.loc[i, '日期'] == start_date:
                    sheet.loc[i, '平仓信号'] = 0
                    if sheet.loc[i, 'Carry收益(10日平均)'] > 0:
                        # 卖空
                        sheet.loc[i, '最优价格'] = sheet.loc[i, '最低价']
                        sheet.loc[i, '策略开仓方向'] = 1
                        sheet.loc[i, '策略开仓价'] = sheet.loc[i, '最低价']
                        sheet.loc[i, '策略昨日结算价'] = 0
                        sheet.loc[i, '策略结算价'] = sheet.loc[i, '结算价']
                        sheet.loc[i, '策略收益'] = sheet.loc[i, '最低价'] - sheet.loc[i, '结算价']
                    else:
                        # 买多
                        sheet.loc[i, '最优价格'] = sheet.loc[i, '最高价']
                        sheet.loc[i, '策略开仓方向'] = -1
                        sheet.loc[i, '策略开仓价'] = sheet.loc[i, '最高价']
                        sheet.loc[i, '策略昨日结算价'] = 0
                        sheet.loc[i, '策略结算价'] = sheet.loc[i, '结算价']
                        sheet.loc[i, '策略收益'] = sheet.loc[i, '结算价'] - sheet.loc[i, '最高价']
                else:
                    # 已经平仓收益为0
                    if sheet.loc[i, '平仓信号'] == 1:
                        sheet.loc[i, '策略开仓方向'] = 0
                        sheet.loc[i, '策略开仓价'] = 0
                        sheet.loc[i, '策略昨日结算价'] = 0
                        sheet.loc[i, '策略结算价'] = 0
                        sheet.loc[i, '策略收益'] = 0
                    else:
                        if sheet.loc[i, '策略开仓方向'] > 0:
                            # 卖空
                            if sheet.loc[i, '最低价'] < sheet.loc[i - 1, '最优价格']:
                                sheet.loc[i, '最优价格'] = sheet.loc[i, '最低价']
                            else:
                                sheet.loc[i, '最优价格'] = sheet.loc[i - 1, '最优价格']
                            if sheet.loc[i, '日期'] == adjust_date:
                                sheet.loc[i, '平仓信号'] = 1
                                sheet.loc[i, '策略结算价'] = sheet.loc[i, '收盘价']
                            elif sheet.loc[i, '最优价格'] + 2.5 * sheet.loc[i, 'ATR'] < sheet.loc[i, '收盘价']:
                                sheet.loc[i, '平仓信号'] = 1 if sheet.loc[i, '平仓信号'] == 0.6667 else 0.6667 if sheet.loc[i, '平仓信号'] == 0.3333 else 0.3333
                                sheet.loc[i, '策略结算价'] = sheet.loc[i, '收盘价']
                            else:
                                sheet.loc[i, '策略结算价'] = sheet.loc[i, '结算价']
                            sheet.loc[i, '策略收益'] = sheet.loc[i, '策略昨日结算价'] - sheet.loc[i, '策略结算价']
                        else:
                            # 买多
                            if sheet.loc[i, '最高价'] > sheet.loc[i - 1, '最优价格']:
                                sheet.loc[i, '最优价格'] = sheet.loc[i, '最高价']
                            else:
                                sheet.loc[i, '最优价格'] = sheet.loc[i - 1, '最优价格']
                            if sheet.loc[i, '日期'] == adjust_date:
                                sheet.loc[i, '平仓信号'] = 1
                                sheet.loc[i, '策略结算价'] = sheet.loc[i, '收盘价']
                            elif sheet.loc[i, '最优价格'] - 2.5 * sheet.loc[i, 'ATR'] > sheet.loc[i, '收盘价']:
                                sheet.loc[i, '平仓信号'] = 1 if sheet.loc[i, '平仓信号'] == 0.6667 else 0.6667 if sheet.loc[i, '平仓信号'] == 0.3333 else 0.3333
                                sheet.loc[i, '策略结算价'] = sheet.loc[i, '收盘价']
                            else:
                                sheet.loc[i, '策略结算价'] = sheet.loc[i, '结算价']
                            sheet.loc[i, '策略收益'] = sheet.loc[i, '策略结算价'] - sheet.loc[i, '策略昨日结算价']
                        if sheet.loc[i, '平仓信号'] != 1 and sheet.loc[i, '合约'] != sheet.loc[i, '前主力合约']:
                            if sheet.loc[i, '策略开仓方向'] > 0:
                                # 卖空
                                sheet.loc[i, '最优价格'] = sheet.loc[i, '最低价']
                                sheet.loc[i, '策略结算价'] = sheet.loc[i, '前主力开盘价']
                                sheet.loc[i, '策略收益'] = sheet.loc[i, '策略昨日结算价'] - sheet.loc[i, '前主力开盘价']
                                sheet.loc[i, '移仓开仓价'] = sheet.loc[i, '开盘价']
                                sheet.loc[i, '移仓昨日结算价'] = 0
                                sheet.loc[i, '移仓结算价'] = sheet.loc[i, '结算价']
                                sheet.loc[i, '移仓收益'] = sheet.loc[i, '开盘价'] - sheet.loc[i, '结算价']
                            else:
                                # 买多
                                sheet.loc[i, '最优价格'] = sheet.loc[i, '最高价']
                                sheet.loc[i, '策略结算价'] = sheet.loc[i, '前主力开盘价']
                                sheet.loc[i, '策略收益'] = sheet.loc[i, '前主力开盘价'] - sheet.loc[i, '策略昨日结算价']
                                sheet.loc[i, '移仓开仓价'] = sheet.loc[i, '开盘价']
                                sheet.loc[i, '移仓昨日结算价'] = 0
                                sheet.loc[i, '移仓结算价'] = sheet.loc[i, '结算价']
                                sheet.loc[i, '移仓收益'] = sheet.loc[i, '结算价'] - sheet.loc[i, '开盘价']
                if i == len(sheet) - 1:
                    break
                else:
                    # 将当前平仓信号默认向前填充
                    sheet.loc[i + 1, '平仓信号'] = sheet.loc[i, '平仓信号']
                    # 将当前策略开仓方向默认向前填充
                    sheet.loc[i + 1, '策略开仓方向'] = sheet.loc[i, '策略开仓方向']
                    # 将当前策略开仓价默认向前填充
                    sheet.loc[i + 1, '策略开仓价'] = sheet.loc[i, '策略开仓价']
                    # 将当前策略结算价默认向前填充
                    sheet.loc[i + 1, '策略昨日结算价'] = sheet.loc[i, '策略结算价']
                    # 如果向前日期大于调整日期 重置开始日期 与 调整日期
                    if sheet.loc[i + 1, '日期'] > adjust_date:
                        start_date = sheet.loc[i + 1, '日期']
                        adjust_index = i + 1 + period_days
                        adjust_date = sheet.loc[adjust_index if adjust_index < len(sheet) else len(sheet) - 1, '日期']
            sheet['盈亏'] = sheet['策略收益'] + sheet['移仓收益']
            # # 删除第一列
            # sheet.drop(sheet.columns[0], axis=1, inplace=True)
            sheet.to_excel(writer, sheet_name=sheet_name, index=False)

    profit_df = pd.read_excel("../output/期货量化实践_Carry收益.xlsx", None)

    trade_notional_map = {}
    trade_date_series = []

    for sheet_name in profit_df:
        sheet = profit_df[sheet_name]
        # 取最后一行
        trade_notional_map[sheet_name] = sheet.iloc[average_days-1]['成交金额(10日平均)']
        trade_date_series = list(sheet['日期'][average_days-1:])
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
    adjust_date = trade_date_series[0 + period_days]
    budget_map = {}
    amount_map = {}
    for i in range(len(trade_date_series)):
        if trade_date_series[i] == start_date:
            for sheet_name in profit_df:
                sheet = profit_df[sheet_name]
                # 取最后一行
                trade_notional_map[sheet_name] = sheet[sheet['日期'] == start_date]['成交金额(10日平均)'].values[0]
            trade_notional_map = {k: v for k, v in trade_notional_map.items() if not pd.isna(v) and v >= 5000000000}
            for sheet_name in profit_df:
                if sheet_name not in trade_notional_map or sheet_name == 'ZC':
                    continue
                sheet = profit_df[sheet_name]
                # 删除第一行
                sheet = sheet.iloc[average_days-1:]
                row = sheet[sheet['日期'] == start_date]
                if len(row) == 0:
                    continue
                carry_ma10_map[sheet_name] = row.iloc[0]['Carry收益(10日平均)']
            # 按照Carry收益(10日平均)排序从大到小排序
            carry_ma10_list = sorted(carry_ma10_map.items(), key=lambda x: x[1], reverse=True)
            time_str = start_date.strftime('%Y-%m-%d')
            print(time_str)
            # 从大到小排序的前20%品种
            positive_carry_ma10_map = {k: v for k, v in carry_ma10_map.items() if v > 0.}
            positive_carry_ma10_map_list = sorted(positive_carry_ma10_map.items(), key=lambda x: x[1], reverse=True)
            carry_ma10_map_first_20 = positive_carry_ma10_map_list[:int(len(positive_carry_ma10_map_list) * 0.20)]
            print(carry_ma10_map_first_20)
            # 从小到大排序的前20%品种
            negative_carry_ma10_map = {k: v for k, v in carry_ma10_map.items() if v < 0.}
            negative_carry_ma10_map_list = sorted(negative_carry_ma10_map.items(), key=lambda x: x[1], reverse=True)
            carry_ma10_map_last_20 = negative_carry_ma10_map_list[-int(len(negative_carry_ma10_map_list) * 0.20):]
            print(carry_ma10_map_last_20)
            for item in carry_ma10_map_first_20 + carry_ma10_map_last_20:
                sheet = profit_df[item[0]]
                row = sheet[sheet['日期'] == start_date]
                if (len(row) == 0) or row.iloc[0]['开仓信号'] == 0:
                    continue
                amount_map[item[0]] = 0
            each_budget = available_budget / len(amount_map.items())
            for item in amount_map.items():
                sheet = profit_df[item[0]]
                row = sheet[sheet['日期'] == start_date]
                amount = int(each_budget / 4 / 0.30 / row.iloc[0]['策略开仓价'] / row.iloc[0]['合约乘数'])
                limit = min(3, int(row.iloc[0]['成交量'] * 1/100000))
                if amount > limit:
                    amount = 1 if limit == 0 else limit
                amount_map[item[0]] = amount * row.iloc[0]['合约乘数']
                if (item[0], '持仓量') not in df.columns:
                    df[(item[0], '持仓量')] = 0.
                    df[(item[0], '单位收益')] = 0.
                    df[(item[0], '持仓收益')] = 0.
                budget_map[item[0]] = each_budget
            for key in amount_map.keys():
                sheet = profit_df[key]
                # 取日期在start_date与adjust_date之间的行
                rows = sheet[(sheet['日期'] >= start_date) & (sheet['日期'] <= adjust_date)]
                profit = 0.
                open_signal = rows.loc[rows['日期'] == start_date, '开仓信号'].values[0]
                close_signal = 0
                for index in rows['日期']:
                    amount = amount_map[key] * open_signal * (1 - close_signal)
                    close_signal = rows[rows['日期'] == index]['平仓信号'].values[0]
                    unit_profit = rows[rows['日期'] == index]['盈亏'].values[0]
                    total_profit = unit_profit * amount
                    profit += total_profit
                    df.loc[index, (key, '持仓量')] = amount
                    df.loc[index, (key, '单位收益')] = unit_profit
                    df.loc[index, (key, '持仓收益')] = total_profit
                available_budget += profit
            print(budget_map)
            print(amount_map)
            print(available_budget)
        else:
            if i == len(trade_date_series) - 1:
                break
            else:
                if trade_date_series[i + 1] > adjust_date:
                    start_date = trade_date_series[i + 1]
                    adjust_index = i + 1 + period_days
                    adjust_date = trade_date_series[adjust_index if adjust_index < len(trade_date_series) else (len(trade_date_series) - 1)]
                    budget_map = {}
                    amount_map = {}

    # 当日收益等于所有持仓收益之和
    df[('', '当日收益')] = df.xs('持仓收益', axis=1, level=1).sum(axis=1)
    # 第一天初始化可用资金为100000000
    df.iloc[0, 0] = 100000000

    total = 0
    for i in range(len(df)):
        if i == len(df) - 1:
            total += df.iloc[i, 1]
        if i > 0:
            # 每一行当日可用资金等于前一日可用资金加上前一日收益
            df.iloc[i, 0] = df.iloc[i - 1, 0] + df.iloc[i - 1, 1]
        # 总收益 / 100000000 * 252 / 交易日数
        rows = df.iloc[0:(i+1), :]
        profit = rows.iloc[:, 1].sum()
        df.iloc[i, 2] = profit / 100000000 * 252 / (i + 1) * 100
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

    result_map[period_days] = total
    df.to_excel('../output/期货量化实践_策略收益.xlsx', sheet_name='策略收益')


test_loop()
#
# for i in range(10, 30):
#     period_days = i
#     test_loop()
#
# print(result_map)