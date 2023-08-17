# !/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2023-08-14 19:29
# @Author  : Jiaxin Liu
# @File    : classify_files.py
# 用来自动化遍历文件夹下所有文件并收集整理对应文件后缀名
import os
import pandas as pd
import zipfile


# 遍历文件夹并分类
def classify_files(directory):
    file_info = []

    for root, _, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            # 先解压缩文件夹下压缩包，再遍历文件，遍历一次即可
            if file_path.endswith('.zip'):
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    unzip_dir = os.path.join(root, 'unzipped_' + file)
                    zip_ref.extractall(unzip_dir)

            name, ext = os.path.splitext(file)
            ext = ext.lower() if ext else "No Extension"  # 处理没有后缀名的文件名
            file_info.append((name, ext))

    return file_info


# 将分类结果写入Excel文件
def write_to_excel(data, output_file):
    df = pd.DataFrame(data, columns=['File Name', 'File Extension'])
    df.to_excel(output_file, index=False)


if __name__ == "__main__":
    data_folder = "data"
    output_excel = "file_info.xlsx"

    classified_data = classify_files(data_folder)
    write_to_excel(classified_data, output_excel)
