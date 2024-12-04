import pdfplumber
import os
from docx import Document
from docx.shared import Pt
import re  # 导入正则表达式模块

def extract_text_from_pdf(pdf_path):
    extracted_text = ""

    # 使用 pdfplumber 提取文本和表格
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                extracted_text += text + "\n\n"

            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    # 将每个单元格转换为字符串，并用 | 分隔
                    row_text = " | ".join(str(cell) if cell is not None else "" for cell in row)
                    extracted_text += row_text + "\n"
                extracted_text += "\n"

    # 使用正则表达式移除所有形如 (cid:数字) 的模式
    cleaned_text = re.sub(r'\(cid:\d+\)', '', extracted_text)
    return cleaned_text

def save_text_to_docx(text, docx_path):
    document = Document()

    style = document.styles['Normal']
    font = style.font
    font.size = Pt(12)

    # 分段添加文本
    for line in text.split('\n'):
        if line.strip():  # 只添加非空行
            document.add_paragraph(line)

    # 确保目标目录存在
    os.makedirs(os.path.dirname(docx_path), exist_ok=True)
    document.save(docx_path)

def process_pdfs(input_dir, output_dir):
    for root, dirs, files in os.walk(input_dir):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_path = os.path.join(root, file)

                relative_path = os.path.relpath(pdf_path, input_dir)
                # 更改文件扩展名为 .docx
                docx_relative_path = os.path.splitext(relative_path)[0] + '.docx'
                docx_path = os.path.join(output_dir, docx_relative_path)

                print(f"正在处理: {pdf_path}")
                try:
                    extracted_content = extract_text_from_pdf(pdf_path)
                    save_text_to_docx(extracted_content, docx_path)
                    print(f"已保存到: {docx_path}")
                except Exception as e:
                    print(f"处理 {pdf_path} 时出错: {e}")

if __name__ == "__main__":
    # 输入文件夹路径（包含PDF文件）
    input_folder = "D:/Desktop/Balance Accessories"
    # 输出文件夹路径（将保存.docx文件）
    output_folder = "D:/Desktop/1"

    process_pdfs(input_folder, output_folder)
    print("所有PDF文件已处理完毕。")
