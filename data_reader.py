# -*- coding: utf-8 -*-
#author: Shawn Zhang

#这段代码作用是将文件夹中所有.csv文件全部转换为xlsx

#导入模块
import os
import pandas as pd
# 添加进度条
from tqdm import tqdm

# 指定要读取的文件夹路径
folder_path = r'C:\Users\Shawn Zhang\Desktop\data\csv\2015\dc'

# 获取文件夹中所有的.csv文件
csv_files = [f for f in os.listdir(folder_path) if f.endswith('.csv')]

# 遍历所有的.csv文件
for csv_file in tqdm(csv_files, desc='Converting .csv to .xlsx'):
    # 读取.csv文件
    df = pd.read_csv(os.path.join(folder_path, csv_file), encoding='gbk')

    # 获取不包括扩展名的文件名
    file_name = os.path.splitext(csv_file)[0]

    # 指定新的保存路径
    save_path = r'C:\Users\Shawn Zhang\Desktop\data\csv\2015\01\dc'

    # 将DataFrame保存为Excel文件，文件名与.csv文件名一致
    df.to_excel(os.path.join(save_path, file_name + '.xlsx'), index=False)

# 转换完成后提示“已经完成”
print("转换完成！")
