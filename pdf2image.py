import fitz
import os
from pathlib import Path


def pdf2img(pdf_dir, img_base_dir, error_log_path):
    error_log = open(error_log_path, 'a')  # 打开错误日志文件，以追加模式
    # 遍历pdf文件夹中的所有文件
    for filename in os.listdir(pdf_dir):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(pdf_dir, filename)
            try:
                # 获取不含扩展名的文件名
                base_name = os.path.splitext(filename)[0].strip()
                # 创建目标保存图片的目录路径
                img_dir = Path(img_base_dir) / base_name
                # 确保图片保存目录存在，如果不存在则创建
                img_dir.mkdir(parents=True, exist_ok=True)

                doc = fitz.open(pdf_path)
                for page_number in range(len(doc)):
                    page = doc[page_number]
                    mat = fitz.Matrix(2, 2)  #将图片放大两倍，便于识别
                    pix = page.get_pixmap(matrix=mat)
                    save_path = img_dir / f"{page_number + 1}.png"
                    pix.save(str(save_path))  # 保存图片，使用str()确保路径是字符串
                doc.close()
            except Exception as e:
                print(f"Error processing file {filename}: {e}")
                error_log.write(f"{filename}\n")  # 将出错的文件名写入错误日志文件
    error_log.close()


if __name__ == '__main__':
    pdf_dir = r"D:/Desktop/Balance Accessories"  #pdf文件文件夹
    img_base_dir = r"D:/Desktop/test1"  #输出图片文件夹
    error_log_path = r"D:/Desktop/error1.txt"  #错误日志文件路径

    pdf2img(pdf_dir, img_base_dir, error_log_path)
