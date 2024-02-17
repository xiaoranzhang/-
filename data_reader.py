# -*- coding: utf-8 -*-
#author: Shawn Zhang

#这段代码作用是将文件夹及其所有子文件夹中所有.csv文件全部转换为xlsx

#导入模块
import os
import pandas as pd
# 添加进度条
from tqdm import tqdm

# 指定要读取的文件夹路径
folder_path = r'C:\Users\Shawn Zhang\Desktop\data\2016\dc'

# 创建一个空列表，用于存储出错的文件名
error_files = []

# 使用os.walk()函数遍历文件夹及其所有子文件夹中的文件
for root, dirs, files in os.walk(folder_path):
    # 创建一个进度条
    with tqdm(total=len(files), desc="转换中：") as pbar:
        # 遍历所有的文件
        for f in files:
            # 检查文件是否是.csv文件
            if f.endswith('.csv'):
                try:
                    # 读取.csv文件
                    df = pd.read_csv(os.path.join(root, f), encoding='gbk')
                except UnicodeDecodeError:
                    print(f"Error decoding file {f}. 跳过...")
                    error_files.append(f)
                    continue

                # 获取不包括扩展名的文件名
                file_name = os.path.splitext(f)[0]

                # 指定新的保存路径
                save_path = r'C:\Users\Shawn Zhang\Desktop\data\excel\dc'

                try:
                    # 将DataFrame保存为Excel文件，文件名与.csv文件名一致
                    df.to_excel(os.path.join(save_path, file_name + '.xlsx'), index=False)
                except Exception as e:
                    print(f"Error saving file {file_name}.xlsx: {e}")
                    error_files.append(f)
                    continue

                # 更新进度条
                pbar.update(1)

# 转换完成后提示“已经完成”
print("转换完成！")
# 打印出所有出错的文件名
if error_files:
    print("以下文件在处理过程中出现错误：")
    for file in error_files:
        print(file)
else:
    print("所有文件都已成功处理。")
    #转换完成后如果没有报错，自动关机
    os.system('shutdown -s -t 0')




