# -*- coding: utf-8 -*-
#author: Shawn Zhang

#这段代码作用是将文件夹中所有.csv文件全部转换为xlsx

#导入模块
import os
import pandas as pd
# 添加进度条
from tqdm import tqdm

# 指定要读取的文件夹路径
folder_path = r'C:\Users\Shawn Zhang\Desktop\data\csv\2015\sc'

# 获取文件夹中所有的.csv文件
csv_files = [f for f in os.listdir(folder_path) if f.endswith('.csv')]

# 创建一个空列表，用于存储出错的文件名
error_files = []

# 遍历所有的.csv文件
for csv_file in tqdm(csv_files, desc='Converting .csv to .xlsx'):
    try:
        # 读取.csv文件
        df = pd.read_csv(os.path.join(folder_path, csv_file), encoding='gbk')
    except UnicodeDecodeError:
        print(f"Error decoding file {csv_file}. Skipping...")
        error_files.append(csv_file)
        continue

    # 获取不包括扩展名的文件名
    file_name = os.path.splitext(csv_file)[0]

    # 指定新的保存路径
    save_path = r'C:\Users\Shawn Zhang\Desktop\data\csv\2015\01\sc'

    try:
        # 将DataFrame保存为Excel文件，文件名与.csv文件名一致
        df.to_excel(os.path.join(save_path, file_name + '.xlsx'), index=False)
    except Exception as e:
        print(f"Error saving file {file_name}.xlsx: {e}")
        error_files.append(csv_file)
        continue
# 转换完成后提示“已经完成”
print("转换完成！")

# 打印出所有出错的文件名
if error_files:
    print("以下文件在处理过程中出现错误：")
    for file in error_files:
        print(file)
else:
    print("所有文件都已成功处理。")




