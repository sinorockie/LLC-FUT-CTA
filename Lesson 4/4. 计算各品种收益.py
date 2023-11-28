import datetime

import numpy as np
import pandas as pd


average_days = 10
period_days = 44
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
