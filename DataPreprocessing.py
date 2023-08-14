# !/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2023-08-03 17:26
# @Author  : Jiaxin Liu
# @File    : DataPreprocessing.py
import os
import re
from pptx import Presentation
from PIL import Image
import easyocr

def read_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()
    return content


# 文本清洗和预处理
def clean_text(text):
    # （可增量更新）需要根据具体情况编写一些正则表达式或规则来清洗文本
    # 这里只是一个简单示例，去除所有非字母和数字的字符
    # cleaned_text = re.sub(r'[^a-zA-Z0-9]', ' ', text)
    cleaned_text = text
    return cleaned_text


def preprocess_files_and_save(directory, output_file):
    with open(output_file, 'w', encoding='utf-8') as output:
        for root, _, files in os.walk(directory):
            for file in files:
                file_path = os.path.join(root, file)
                # 对于文本文件，进行文本清洗和预处理，并将结果写入文件
                if file_path.endswith('.txt') or file_path.endswith('.docx'):
                    with open(file_path, 'r', encoding='utf-8') as f:
                        text = f.read()
                    cleaned_text = clean_text(text)
                    output.write(cleaned_text + '\n')
                # 对于其他文件类型，可以根据需要执行相应的处理操作
                elif file_path.endswith('.jpg') or file_path.endswith('.png') or file_path.endswith('.xml'):
                    print(file_path)
                    img = Image.open(file_path)
                    # img.show()  # 显示图片
                    # pixels = list(img.getdata())  # 获取像素数
                    ocr = easyocr.Reader(['ch_sim', 'en'], gpu=False)
                    # 识别图片文字
                    content = ocr.readtext(img
                    print(content)
                else:
                    print(f"Unsupported file format: {file_path}")


# def preprocess_files(directory):
#     for root, _, files in os.walk(directory):
#         for file in files:
#             file_path = os.path.join(root, file)
#             # 对于txt文件，直接读取文本内容并进行清洗
#             if file_path.endswith('.txt'):
#                 with open(file_path, 'r', encoding='utf-8') as f:
#                     text = f.read()
#                 cleaned_text = clean_text(text)
#                 with open(file_path, 'w', encoding='utf-8') as f:
#                     f.write(cleaned_text)
#             # 对于Word文档（.docx），使用docx2txt库进行解析并清洗
#             elif file_path.endswith('.docx'):
#                 text = docx2txt.process(file_path)
#                 cleaned_text = clean_text(text)
#                 with open(file_path.replace('.docx', '.txt'), 'w', encoding='utf-8') as f:
#                     f.write(cleaned_text)
#             # 对于PowerPoint文档（.pptx），使用python-pptx库进行解析并清洗
#             elif file_path.endswith('.pptx'):
#                 prs = Presentation(file_path)
#                 text = ''
#                 for slide in prs.slides:
#                     for shape in slide.shapes:
#                         if hasattr(shape, "text"):
#                             text += shape.text + '\n'
#                 cleaned_text = clean_text(text)
#                 with open(file_path.replace('.pptx', '.txt'), 'w', encoding='utf-8') as f:
#                     f.write(cleaned_text)
#             # 对于其他文件类型，可以根据需要进行类似处理（待增量更新）
#             elif file_path.endswith('.png') or file_path.endswith('.jpg') or file_path.endswith(
#                     '.eml') or file_path.endswith('.yml') or file_path.endswith('.xml') or file_path.endswith(
#                 '.properties') or file_path.endswith('.zip') or file_path.endswith('.hiv') or file_path.endswith(
#                 '.sh') or file_path.endswith('.py') or file_path.endswith('.md'):
#                 # 对于图片、邮件、配置文件、压缩包等其他类型文件，可以跳过文本处理，或者执行相应的解析操作
#                 pass
#             else:
#                 print(f"Unsupported file format: {file_path}")
