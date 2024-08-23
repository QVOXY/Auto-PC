import os
import sys
import win32com.client

def create_sections_dict(keylist, directory):
    sections_dict = {}
    for key in keylist:
        filename = f"{key}.docx"
        full_path = os.path.join(directory, filename)
        sections_dict[key] = full_path
    return sections_dict

def process_document(sections_dict, target_doc_path):
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False

    try:
        target_doc = word.Documents.Open(target_doc_path)

        for field_to_find, source_doc_path in sections_dict.items():
            find = word.Selection.Find
            find.Text = field_to_find
            find.Forward = True

            while find.Execute():
                source_doc = word.Documents.Open(source_doc_path)
                source_doc.Content.Select()
                source_doc.Content.Copy()
                source_doc.Close(False)
                word.Selection.Paste()
                word.Selection.Collapse(0)

        target_doc.Save()
    except Exception as e:
        print(f"处理文档时发生错误: {e}")
    finally:
        if 'target_doc' in locals():
            target_doc.Close()
        word.Quit()

def main():
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

    # 创建sections_dict
    sections_dict = create_sections_dict(keylist, os.path.abspath(os.path.join('procedure', 'process2')))

    # 处理目标文档
    target_doc_path = os.path.abspath(os.path.join( 'procedure', 'process3', 'tem.docx'))
    process_document(sections_dict, target_doc_path)

if __name__ == "__main__":
    main()
