import docx

def read_keywords(file_path):
    """
    从文件中读取关键词，每行一个关键词
    """
    with open(file_path, 'r', encoding='utf-8') as file:
        keywords = [line.strip() for line in file if line.strip()]
    return keywords

def remove_keywords_from_docx(docx_path, keywords, output_path):
    """
    从.docx文件中移除指定的关键词，并保存到新的路径
    """
    doc = docx.Document(docx_path)
    for paragraph in doc.paragraphs:
        for keyword in keywords:
            paragraph.text = paragraph.text.replace(keyword, '')
    doc.save(output_path)

def main():
    keywords_file_path = 'para_library\\transmit1'
    docx_file_path = 'procedure\\process3\\tem.docx'
    output_file_path = 'procedure\\process4\\tem.docx'
    
    keywords = read_keywords(keywords_file_path)
    remove_keywords_from_docx(docx_file_path, keywords, output_file_path)

if __name__ == "__main__":
    main()
