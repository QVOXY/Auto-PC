import os
import shutil

# 定义路径
current_template_path = 'para_library\\current_template'
tem_library_path = 'tem_library'
procedure_tem_file_path = 'procedure\\tem_file\\tem.docx'
procedure_tem_file_path_2 = 'procedure\\process3\\tem.docx'
# 读取期刊信息
with open(current_template_path, 'r', encoding='utf-8') as file:
    journal_info = file.read().strip()

# 构建源文件路径
source_file_path = os.path.join(tem_library_path, journal_info, 'tem.docx')

# 确保源文件存在
if not os.path.exists(source_file_path):
    print(f"源文件 {source_file_path} 不存在")
    exit(1)

# 复制文件到目标路径，如果目标路径有文件则覆盖
shutil.copy2(source_file_path, procedure_tem_file_path)
shutil.copy2(source_file_path, procedure_tem_file_path_2)

