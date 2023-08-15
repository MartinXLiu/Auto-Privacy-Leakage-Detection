# !/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2023-08-03 17:26
# @Author  : Jiaxin Liu
# @File    : DataPreprocessing.py
import os
import re
from pptx import Presentation
from PIL import Image
import docx
# import easyocr
import win32com.client as wc
import time
ROOT = "D:\\Auto-Privacy-Leakage-Detection\\Auto-Privacy-Leakage-Detection\\"


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
                if file_path.endswith('.wps'):
                    file_path = ROOT + file_path
                    file_path.replace('.wps','.docx')
                    try:
                        text = ""
                        with open(file_path, 'r', encoding='utf-8') as f:
                            text += f.read()
                    except UnicodeDecodeError:
                        print("maybe there are figures in .wps")
                    cleaned_text = clean_text(text)
                    print(".wps",cleaned_text)
                    output.write(cleaned_text + '\n')
                    pass
                elif file_path.endswith('.doc'):
                    doc_path = ROOT+file_path #使用绝对路径
                    print(file_path)
                    word = wc.Dispatch("Word.Application")
                    print("49",doc_path)
                    doc = word.Documents.Open(doc_path)
                    docx_path = doc_path.replace('.doc','.docx')
                    doc.SaveAs(docx_path, 12)
                    doc.Close()
                    word.Quit()
                    document = docx.Document(docx_path)
                    try:
                        text = ""
                        for p in document.paragraphs:
                            # print(p.text)
                            text += p.text
                        cleaned_text = clean_text(text)
                        output.write(cleaned_text + '\n')
                    except UnicodeDecodeError:
                        print("maybe there are figures in .doc")
                    os.remove(docx_path)
                    time.sleep(2)  # 设置延时，避免程序太快，导致上一个word没有关闭
                elif file_path.endswith('.txt') or file_path.endswith('.docx'):
                    with open(file_path, 'r', encoding='utf-8') as f:
                        text = f.read()
                    cleaned_text = clean_text(text)
                    output.write(cleaned_text + '\n')
                elif file_path.endswith('.jpg') or file_path.endswith('.png') or file_path.endswith('.xml'):
                    pass
                    # print(file_path)
                    # img = Image.open(file_path)
                    # img.show()  # 显示图片
                    # pixels = list(img.getdata())  # 获取像素数
                    # ocr = easyocr.Reader(['ch_sim', 'en'], gpu=False)
                    # # 识别图片文字
                    # content = ocr.readtext(img)
                    # print(content)
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
