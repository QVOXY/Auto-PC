import os
import sys
import shutil
from win32com import client

# 添加 fun_library 到系统路径
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'fun_library'))
from Thesis_processing import split_doc_by_paragraphs

# 读取 para_library\paper_postion 里的信息作为一个文件的路径
with open(os.path.join(os.path.dirname(__file__), '..', 'para_library', 'paper_position'), 'r', encoding='gbk') as file:
    paper_path = file.read().strip()

# 复制文件到 procedure\origin_file 目录里重命名为 paper.docx
origin_file_dir = os.path.join(os.path.dirname(__file__), '..', 'procedure', 'origin_file')
if not os.path.exists(origin_file_dir):
    os.makedirs(origin_file_dir)
shutil.copy(paper_path, os.path.join(origin_file_dir, 'paper.docx'))

# 设置 original_doc_path 和 save_path
original_doc_path = os.path.join(origin_file_dir, 'paper.docx')
save_path = os.path.join(os.path.dirname(__file__), '..', 'procedure', 'process1')

# 确保 save_path 存在
if not os.path.exists(save_path):
    os.makedirs(save_path)

# 调用函数分段处理文档
split_doc_by_paragraphs(original_doc_path, save_path)
