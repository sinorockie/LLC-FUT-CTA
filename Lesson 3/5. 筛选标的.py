import datetime

import pandas as pd

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

carry_ma10_map = {}
start_date = trade_date_series[0]
adjust_date = trade_date_series[0] + datetime.timedelta(days=30)
budget_map = {}
amount_map = {}
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
        print(budget_map)
        print(amount_map)
        for key in amount_map.keys():
            sheet = profit_df[key]
            # 取日期在start_date与adjust_date之间的行
            rows = sheet[(sheet['日期'] >= start_date) & (sheet['日期'] < adjust_date)]
            # 累计盈亏求和
            profit = rows['策略收益'].sum()
            available_budget += (profit * amount_map[key])
        print(available_budget)
    else:
        if i == len(trade_date_series) - 1:
            break
        else:
            if trade_date_series[i+1] > adjust_date:
                start_date = trade_date_series[i+1]
                adjust_date = start_date + datetime.timedelta(days=30)
                budget_map = {}
                amount_map = {}
