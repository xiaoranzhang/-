#这段代码是读取.xlsx数据，并且画出蜡烛图

import pandas as pd
import plotly.graph_objects as go
import plotly.offline as pyo

# Set pandas display options
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', None)

# Read the Excel file
df = pd.read_excel(r'C:\Users\Shawn Zhang\Desktop\data\2015\dc\20150105\l1505_20150105.xlsx')

# Convert '时间' to datetime and set as index
df['时间'] = pd.to_datetime(df['时间'])
df.set_index('时间', inplace=True)

# Resample to one-minute intervals
df_resampled = df['最新'].resample('1min').ohlc()

# Calculate the weighted average for each day
df_resampled['weighted_avg'] = df_resampled.apply(lambda row: (row['open'] * 0.1 + row['high'] * 0.2 + row['low'] * 0.3 + row['close'] * 0.4), axis=1)

# Calculate the 3-minute moving average
#df_resampled['moving_avg'] = df_resampled['close'].rolling(window=3).mean()

# Create a candlestick chart
fig = go.Figure(data=[go.Candlestick(x=df_resampled.index,
                                     open=df_resampled['open'],
                                     high=df_resampled['high'],
                                     low=df_resampled['low'],
                                     close=df_resampled['close'],
                                     increasing_line_color='red', increasing_fillcolor='red',
                                     decreasing_line_color='green', decreasing_fillcolor='green')])

# Add the weighted average line to the plot
fig.add_trace(go.Scatter(x=df_resampled.index, y=df_resampled['weighted_avg'], mode='lines', name='Weighted Average'))
# update y-axis units to 25
fig.update_yaxes(dtick=15)

# Plot the figure in an interactive HTML file
pyo.plot(fig, filename='candlestick.html')
