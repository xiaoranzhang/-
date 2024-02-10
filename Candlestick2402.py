#这段代码是读取.xlsx数据，并且画出蜡烛图

import pandas as pd
import plotly.graph_objects as go
import plotly.offline as pyo

# 设置pandas读取数据后格式
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', None)

# 用pandas打开excel文件
df = pd.read_excel(r'C:\Users\Shawn Zhang\Desktop\data\2015\dc\20150105\l1505_20150105.xlsx')

# 将'时间' 设置为索引
df['时间'] = pd.to_datetime(df['时间'])
df.set_index('时间', inplace=True)

# 将时间resample为1分钟
df_resampled = df['最新'].resample('1min').ohlc()

# 计算每日加权平均数
df_resampled['weighted_avg'] = df_resampled.apply(lambda row: (row['open'] * 0.1 + row['high'] * 0.2 + row['low'] * 0.3 + row['close'] * 0.4), axis=1)

# 计算3分钟级别移动加权平均数（注释掉）
#df_resampled['moving_avg'] = df_resampled['close'].rolling(window=3).mean()

# 画一个蜡烛图
fig = go.Figure(data=[go.Candlestick(x=df_resampled.index,
                                     open=df_resampled['open'],
                                     high=df_resampled['high'],
                                     low=df_resampled['low'],
                                     close=df_resampled['close'],
                                     increasing_line_color='red', increasing_fillcolor='red',
                                     decreasing_line_color='green', decreasing_fillcolor='green')])

# 将加权平均曲线画在蜡烛图
fig.add_trace(go.Scatter(x=df_resampled.index, y=df_resampled['weighted_avg'], mode='lines', name='Weighted Average'))
# update y-axis units to 25
fig.update_yaxes(dtick=15)

# 画在HTML文件，使之变成可互动图表
pyo.plot(fig, filename='candlestick.html')
