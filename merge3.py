import re
from glob import glob
import os
import docx
from docxcompose.composer import Composer

base_dir = r"D:/Desktop/test2"  # 一级文件夹路径
save_path = r"D:/Desktop/test3"  # 合并文档的保存路径


def sort_key(filename):
    # 使用正则表达式提取文件名前面的数字
    match = re.match(r'(\d+)', os.path.basename(filename))
    # 如果找到数字，则转换为整数，否则使用文件名作为排序键
    return int(match.group(1)) if match else os.path.basename(filename)


def combine_docs_in_subfolder(subfolder_path):
    # 获取二级文件夹中的所有.docx文件并按名称排序
    files_list = sorted(glob(os.path.join(subfolder_path, '*.docx')), key=sort_key)
    master_doc = docx.Document()
    composer = Composer(master_doc)

    for file_path in files_list:
        doc_temp = docx.Document(file_path)
        composer.append(doc_temp)

    return master_doc


def save_combined_docs(base_dir, save_path):
    # 确保保存路径存在
    if not os.path.exists(save_path):
        os.makedirs(save_path)

    # 遍历一级文件夹中的所有二级文件夹
    for subfolder_name in os.listdir(base_dir):
        subfolder_path = os.path.join(base_dir, subfolder_name)
        if os.path.isdir(subfolder_path):  # 确保是文件夹
            # 合并二级文件夹中的文档
            combined_doc = combine_docs_in_subfolder(subfolder_path)
            # 以二级文件夹名称保存合并后的文档
            subfolder_name_without_path = subfolder_name.rstrip('.docx')  # 去除文件扩展名
            combined_doc.save(os.path.join(save_path, subfolder_name_without_path + '.docx'))


# 执行合并操作
save_combined_docs(base_dir, save_path)
