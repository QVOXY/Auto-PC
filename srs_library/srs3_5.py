import os
import sys

def get_absolute_path(relative_path):
    return os.path.abspath(os.path.join(os.path.dirname(__file__), '..', relative_path))

def read_file_to_list(file_path):
    try:
        with open(file_path, 'r') as f:
            return [line.strip() for line in f.readlines()]
    except FileNotFoundError:
        print(f"文件 {file_path} 不存在。")
        sys.exit(1)
    except Exception as e:
        print(f"读取文件 {file_path} 时发生错误: {e}")
        sys.exit(1)

# 添加 fun_library 到系统路径
fun_library_path = get_absolute_path(r'fun_library')
sys.path.append(fun_library_path)
from Thesis_processing import combine_documents

# 读取索引列表
list_path = get_absolute_path(r'para_library\transmit2')
try:
    num_list = [int(num) for num in read_file_to_list(list_path)]
except ValueError:
    print(f"文件 {list_path} 中的内容不是有效的整数。")
    sys.exit(1)

# 读取关键字列表
keylist_path = get_absolute_path(r'para_library\transmit1')
keylist = read_file_to_list(keylist_path)[2:]

# 定义基础路径和新路径
base_path = get_absolute_path(r'procedure\process1')
new_path = get_absolute_path(r'procedure\process2')

# 确保新路径存在
if not os.path.exists(new_path):
    os.makedirs(new_path)

# 调用 combine_documents 函数
try:
    combine_documents(num_list[2], num_list[3]-1, base_path, os.path.join(new_path, f"{keylist[2]}.docx"))
    print(f"成功生成文件: {keylist[2]}.docx")  # 添加的提示信息
except IndexError:
    print("索引列表或关键字列表的长度不足。")
    sys.exit(1)
except Exception as e:
    print(f"发生错误: {e}")
    sys.exit(1)
