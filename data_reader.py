# 导入pandas模块
import pandas as pd
import os

#pandas打开csv文件
df = pd.read_csv(r'D:\数据\2019\dc\20190109\a1905_20190109.csv', encoding='gbk')
#将DataFrame转换为列表
data = df.values.tolist()

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', 1000)

#打印表格
print(df)
# 将“时间”列转换为日期类型
df['时间'] = pd.to_datetime(df['时间'])

# 将DataFrame保存到新的Excel文件中
df.to_excel(r'C:\Users\Shawn Zhang\Desktop\test.xlsx', index=False)
