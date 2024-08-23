import os
import sys

# 添加 fun_library 到系统路径
fun_library_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'fun_library'))
sys.path.append(fun_library_path)
from Thesis_processing import find_keywords_indices

# 初始化一个空列表
keylist = []

# 打开文件，文件路径是相对于当前脚本文件的上一级目录中的para_library文件夹下的transmit1文件
transmit1_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'para_library', 'transmit1'))
try:
    with open(transmit1_path, 'r', encoding='utf-8') as f:
        for line in f:  # 遍历文件中的每一行
            keylist.append(line.strip())  # 使用strip()去除每行的前后空白字符，然后添加到列表中
except FileNotFoundError:
    print(f"文件 {transmit1_path} 未找到。")
    sys.exit(1)
except IOError:
    print(f"读取文件 {transmit1_path} 时发生错误。")
    sys.exit(1)

# 去掉列表的前两行
keylist = keylist[2:]

# 原始文档路径
original_doc_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'procedure', 'origin_file', 'paper.docx'))

# 调用函数，传入原始文档路径和关键词列表，返回关键词在文档中的索引列表
try:
    indices_list = find_keywords_indices(original_doc_path, keylist)
except Exception as e:
    print(f"处理文档时发生错误: {e}")
    sys.exit(1)

# 将列表存入一个文件，用于传参
transmit2_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'para_library', 'transmit2'))
with open(transmit2_path, 'w', encoding='utf-8') as f:
    for item in indices_list:
        f.write("%s\n" % item)
