from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

def list_table_indices_with_captions(doc_path):
    doc = Document(doc_path)

    # Build a list of elements (paragraphs and tables) in document order
    elements = []
    for block in doc.element.body:
        if block.tag.endswith('tbl'):
            elements.append(('table', block))
        elif block.tag.endswith('p'):
            elements.append(('paragraph', block))

    table_info = []
    table_index = 0
    last_caption = None

    for i, (el_type, el) in enumerate(elements):
        if el_type == 'paragraph':
            para = Paragraph(el, doc)
            text = para.text.strip()
            if text.startswith("表格 "):  # This is a caption
                last_caption = text
        elif el_type == 'table':
            if last_caption:
                caption = last_caption
            else:
                caption = "(No caption found above this table)"
            table_info.append((table_index, caption))
            table_index += 1
            last_caption = None  # Reset after use

    # Print the results
    for idx, caption in table_info:
        print(f"Table {idx}: {caption}")

# Example usage — change to your actual file path
doc_path = r"D:\user\Desktop\learning\code_ip\Report_Builder_Code\template\template.docx"
list_table_indices_with_captions(doc_path)
