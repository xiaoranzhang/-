# -*- coding: utf-8 -*-
#author: Shawn Zhang

import os
import pandas as pd
from tqdm import tqdm
from pathlib import Path
from multiprocessing import Pool

#定义函数，接收一个文件路径作为输入
def process_file(file_path):
    error_files = []
    try:
        df = pd.read_csv(file_path, encoding='gbk')
        file_name = file_path.stem
        #保存新的路径
        save_path = Path(r'C:\Users\Shawn Zhang\Desktop\data\zc')
        df.to_excel(save_path / f"{file_name}.xlsx", index=False)
    except Exception as e:
        print(f"Error processing file {file_path.name}: {e}")
        error_files.append(file_path.name)
    return error_files

def main():
    #读取文件夹中的所有.csv文件
    folder_path = Path(r'C:\Users\Shawn Zhang\Desktop\data\2016\zc')
    csv_files = list(folder_path.glob('**/*.csv'))

    #使用多进程处理文件
    with Pool() as p:
        results = list(tqdm(p.imap(process_file, csv_files), total=len(csv_files), desc="转换中"))
    #将嵌套列表展开
    error_files = [item for sublist in results for item in sublist]

    print("转换完成！")
    if error_files:
        print("以下文件在处理过程中出现错误：")
        for file in error_files:
            print(file)
    else:
        print("所有文件都已成功处理。")

if __name__ == "__main__":
    main()
