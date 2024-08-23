import os
from win32com import client
import win32com.client as win32  # 这里修改了导入语句



'''
find_keywords_indices(doc_path)：“关键词索引函数”
输入参数doc_path（原文件），根据关键词列表识别关键词，最后返回“段落索引列表"，用于后续分割文件
'''
def find_keywords_indices(doc_path,find_key_list):
    # 参数为 原文地址 与 关键词列表
    # 启动Word应用程序并打开文档
    word = client.DispatchEx("Word.Application")
    word.Visible = False  # 设置Word可见
    doc = word.Documents.Open(doc_path)

    # 初始化段落索引列表
    paragraph_indices = []

    # 创建一个范围对象，用于查找操作
    range = doc.Content

    # 存储段落的起始位置，以便计算索引
    paragraph_start_positions = [0]

    # 用于计算段落索引的变量
    current_position = 0

    # 对每个关键词进行查找
    for key in find_key_list:
        # 重置范围的起始位置为文档的开始
        range.Start = 0
        range.End = 0  # 确保End也重置到文档开始，否则Find.Execute可能不会工作
        found = False

        # 遍历文档中的所有段落
        for paragraph in doc.Paragraphs:
            # 更新段落起始位置
            if paragraph.Range.Start > current_position:
                paragraph_start_positions.append(paragraph.Range.Start)
                current_position = paragraph.Range.Start

            # 检查段落是否包含搜索的文本
            if key in paragraph.Range.Text:
                found = True
                # 计算当前段落的索引
                paragraph_index = paragraph_start_positions.index(paragraph.Range.Start)
                # 添加段落的索引到索引列表
                paragraph_indices.append(paragraph_index + 1)  # 索引从1开始
                #print(f"Found key '{key}' at paragraph index {paragraph_index + 1}")

        if not found:
            print(f"No paragraph containing '{key}' was found.")

    # 关闭文档并退出Word
    doc.Close(False)
    word.Quit()

    # 返回段落索引列表
    return paragraph_indices



'''
split_doc_by_paragraphs(original_doc_path, save_path):“分割段落函数”
将原文件的每一段分割为一个单独的段落文件，段落文件存储在中间文件夹中，以便于后续操作
输入参数为“原始文件”与需要存储的“中间文件夹”，无返回值
'''
def split_doc_by_paragraphs(original_doc_path, save_path):
    # 确保保存路径存在
    if not os.path.exists(save_path):
        os.makedirs(save_path)

    # 启动Word应用程序
    word = client.DispatchEx("Word.Application")
    word.Visible = False  # 设置Word可见

    # 打开原始文档
    doc = word.Documents.Open(original_doc_path)

    # 遍历文档中的每一段落
    for i in range(doc.Paragraphs.Count):
        # 创建一个新的Word文档
        new_doc = word.Documents.Add()

        # 将当前段落文本复制到新文档
        doc.Paragraphs(i + 1).Range.Copy()
        new_doc.Range().Paste()

        # 定义新文档的文件名，使用前缀 "paragraph_" 并附加段落编号
        new_doc_title = f"paragraph_{i + 1}"
        new_doc_path = os.path.join(save_path, f"{new_doc_title}.docx")

        # 保存新文档
        new_doc.SaveAs(new_doc_path)

        # 关闭新文档，不保存更改
        new_doc.Close(False)

    # 关闭原始文档
    doc.Close(False)

    # 清理：退出Word
    word.Quit()



'''
combine_documents(start_num, end_num, base_path, path_new)：“组合段落函数”
输入参数分别为起始段、终止段、原段落文件目录、组合后文件名（注意是文件名，需要逐个设置）
'''
def combine_documents(start_num, end_num, base_path, path_new):
    # 文件扩展名
    file_suffix = ".docx"
    # 假设我们要处理的文档前缀
    file_prefix = "paragraph"

    # 使用列表推导式生成所有文档的路径
    paths = [os.path.join(base_path, f"{file_prefix}_{i}{file_suffix}") for i in range(start_num, end_num + 1)]

    # 启动Word应用程序
    word = client.Dispatch("Word.Application")
    word.Visible = False  # 确保Word不可见

    # 创建并保存新文档
    doc_new = word.Documents.Add()
    doc_new.SaveAs(path_new)
    doc_new.Close()

    for path in paths:
        # 打开文档
        doc = word.Documents.Open(path)
        # 复制文档的全部内容
        doc.Content.Copy()
        # 关闭文档，不保存更改
        doc.Close(False)

        # 打开新文档
        doc_new = word.Documents.Open(path_new)
        s = word.Selection
        # 将光标移动到文末
        s.MoveRight(1, doc_new.Content.End)
        # 粘贴复制的内容
        s.Paste()

        # 保存并关闭新文档
        doc_new.SaveAs(path_new)
        doc_new.Close()

    # 关闭Word应用程序
    word.Quit()



'''
remove_empty_paragraphs(file_path):"删除空行段落函数"
'''
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

'''
traverse_directory(directory_path):“遍历目录函数”
'''
def traverse_directory(directory_path):
    # 遍历目录中的所有文件和子目录
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            if file.endswith(".docx"):
                file_path = os.path.join(root, file)
                print(f"Processing file: {file_path}")
                remove_empty_paragraphs(file_path)