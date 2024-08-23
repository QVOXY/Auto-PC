import os
import sys
import win32com.client

def set_paragraph_styles(doc, style_dict):
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False

    try:
        doc = word.Documents.Open(doc)
        paragraphs = doc.Paragraphs

        keys = list(style_dict.keys())
        for i in range(len(keys)):
            key = keys[i]
            style = style_dict[key]
            next_key = keys[i + 1] if i + 1 < len(keys) else None

            found = False
            for j, para in enumerate(paragraphs):
                if key in para.Range.Text:
                    found = True
                    start_index = j
                elif found and next_key and next_key in para.Range.Text:
                    end_index = j
                    break
                elif found and not next_key:
                    end_index = len(paragraphs)

            if found:
                if key == 'Author-':
                    paragraphs[start_index].Range.Style = style
                else:
                    for k in range(start_index, end_index):
                        paragraphs[k].Range.Style = style

        doc.Save()
    except Exception as e:
        print(f"处理文档时发生错误: {e}")
    finally:
        if 'doc' in locals():
            doc.Close()
        word.Quit()

def main():
    style_dict_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'para_library', 'style_dictionary'))
    try:
        with open(style_dict_path, 'r', encoding='utf-8') as f:
            style_dict = eval(f.read())
    except FileNotFoundError:
        print(f"文件 {style_dict_path} 未找到。")
        sys.exit(1)
    except IOError:
        print(f"读取文件 {style_dict_path} 时发生错误。")
        sys.exit(1)

    target_doc_path = os.path.abspath(os.path.join('procedure', 'process3', 'tem.docx'))
    set_paragraph_styles(target_doc_path, style_dict)

if __name__ == "__main__":
    main()
