import os
import shutil

# 原始文件路径
original_file_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', r'procedure\process1\paragraph_2.docx'))

# 目标文件夹路径
destination_folder_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', r'procedure\process2'))

# 新文件名
new_file_name = 'Author-.docx'

# 确保目标文件夹存在
if not os.path.exists(destination_folder_path):
    os.makedirs(destination_folder_path)

# 目标文件的完整路径
destination_file_path = os.path.join(destination_folder_path, new_file_name)

# 复制文件并添加错误处理
try:
    shutil.copy(original_file_path, destination_file_path)
    print("成功生成文件: Author-.docx")
except FileNotFoundError:
    print(f"文件 {original_file_path} 不存在。")
except PermissionError:
    print(f"没有权限复制文件到 {destination_file_path}。")
except Exception as e:
    print(f"发生错误: {e}")
