import os
from docx import Document

def remove_empty_paragraphs(file_path):
    # 打开docx文件
    doc = Document(file_path)
    
    # 遍历所有段落，删除空行段落
    for paragraph in doc.paragraphs:
        if paragraph.text.strip() == "":
            p = paragraph._element
            p.getparent().remove(p)
            paragraph._p = paragraph._element = None
    
    # 保存修改后的文件
    doc.save(file_path)

def traverse_directory(directory_path):
    # 遍历目录中的所有文件和子目录
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            if file.endswith(".docx"):
                file_path = os.path.join(root, file)
                print(f"Processing file: {file_path}")
                remove_empty_paragraphs(file_path)

if __name__ == "__main__":
    directory_path = "procedure/process2"
    traverse_directory(directory_path)
