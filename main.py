import re

from classify_files import classify_files, write_to_excel, decompress
from DataPreprocessing import read_file, preprocess_files_and_save


if __name__ == "__main__":
    data_source = "data.rar"
    # 假设解压后的数据在名为data的文件夹中
    data_folder = "data"

    # 调用classify_files.py文件，得到所有文件后缀
    output_excel = "file_info.xlsx"

    # 先解压原始rar压缩包, 然后解压文件夹里面的zip压缩包，并删除掉zip原始压缩包文件
    decompress(data_source, data_folder)

    classified_data = classify_files(data_folder)
    write_to_excel(classified_data, output_excel)

    # 调用DataPreprocessing.py，解析文件，输出文本格式数据
    output_file = "preprocessed_data.txt"
    preprocess_log = "preprocess_log.txt"
    preprocess_files_and_save(data_folder, output_file, preprocess_log)

    preprocessed_data = read_file(output_file)
    # 编写正则表达式模式
    ip_pattern = r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}'
    username_pattern = r'(?i)username|user|name|login'
    password_pattern = r'(?i)password|pass|pwd'
    sensitive_info_pattern = fr'({ip_pattern}|{username_pattern}|{password_pattern})'

    # 查找匹配的敏感信息
    matches = re.findall(sensitive_info_pattern, preprocessed_data)

    # 去除重复项
    unique_matches = set(matches)

    # 输出提取的敏感信息
    print(unique_matches)


