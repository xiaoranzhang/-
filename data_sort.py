# -*- coding: utf-8 -*-
#author: Shawn Zhang

from tqdm import tqdm
import os
import pandas as pd

# 指定要读取的文件夹路径
root_folder_path = r'C:\Users\Shawn Zhang\Desktop\data\excel\dc2015 2016 2017'

# 遍历根目录下的所有子目录
for folder_name in tqdm(os.listdir(root_folder_path), desc="数据处理中"):
    folder_path = os.path.join(root_folder_path, folder_name)

    # 检查是否为目录
    if os.path.isdir(folder_path):
        # 创建一个空列表，用于存储数据
        data = []

        # 使用os.listdir()函数遍历子目录中的文件
        for file in os.listdir(folder_path):
            # 检查文件是否是.xlsx文件
            if file.endswith('.xlsx'):
                # 读取Excel文件的B2，C2和E2单元格
                df = pd.read_excel(os.path.join(folder_path, file), usecols="B,C,E", nrows=1)
                # 将读取的数据添加到列表中
                data.append(df.values[0])

        # 将列表转换为DataFrame
        df = pd.DataFrame(data, columns=['B2', 'C2', 'E2'])

        # 指定新的保存路径
        save_path = os.path.join(r'C:\Users\Shawn Zhang\Desktop\data\_temp\dc', f"{folder_name}.xlsx")

        # 将DataFrame写入新的Excel文件
        df.to_excel(save_path, index=False)
