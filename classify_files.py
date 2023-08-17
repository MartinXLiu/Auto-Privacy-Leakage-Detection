# !/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2023-08-14 19:29
# @Author  : Jiaxin Liu
# @File    : classify_files.py
# 用来自动化遍历文件夹下所有文件并收集整理对应文件后缀名
import os
import pandas as pd
import zipfile
import rarfile


def decompress(file_name, dir_name):
    """
    对输入文件解压，然后对文件夹内部压缩包解压，同时删除掉文件夹内部原始压缩包
    :param file_name: 输入文件路径
    :param dir_name: 文件解压目录
    :return:
    """
    with rarfile.RarFile(file_name) as rar_ref:
        rar_ref.extractall(dir_name)

    for root, _, files in os.walk(dir_name):
        for file in files:
            file_path = os.path.join(root, file)
            if file_path.endswith('.zip'):
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    unzip_dir = os.path.join(root, file[:-4])
                    zip_ref.extractall(unzip_dir)
                os.remove(file_path)


# 遍历文件夹并分类
def classify_files(directory):
    file_info = []

    for root, _, files in os.walk(directory):
        for file in files:
            name, ext = os.path.splitext(file)
            ext = ext.lower() if ext else "No Extension"  # 处理没有后缀名的文件名
            file_info.append((name, ext))

    return file_info


# 将分类结果写入Excel文件
def write_to_excel(data, output_file):
    df = pd.DataFrame(data, columns=['File Name', 'File Extension'])
    df.to_excel(output_file, index=False)


if __name__ == "__main__":
    data_source = "data.rar"
    data_folder = "data"
    output_excel = "file_info.xlsx"

    # 先解压原始rar压缩包, 然后解压文件夹里面的zip压缩包，并删除掉zip原始压缩包文件
    decompress(data_source, data_folder)

    classified_data = classify_files(data_folder)
    write_to_excel(classified_data, output_excel)