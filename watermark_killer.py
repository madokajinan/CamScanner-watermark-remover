import os
import re
import fitz  # PyMuPDF
from PIL import Image
from tkinter import Tk, filedialog, messagebox
import shutil  # 用于删除文件夹

# 提取PDF中的嵌入图片
def extract_images_from_pdf(input_pdf, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    pdf_document = fitz.open(input_pdf)
    for page_number in range(len(pdf_document)):
        page = pdf_document[page_number]
        images = page.get_images(full=True)
        if images:
            for img_index, img in enumerate(images):
                xref = img[0]
                base_image = pdf_document.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]
                image_filename = f"{output_folder}/page{page_number + 1}_img{img_index + 1}.{image_ext}"
                with open(image_filename, "wb") as image_file:
                    image_file.write(image_bytes)
        else:
            pass
    pdf_document.close()

# 删除水印图片
def remove_watermarked_images(folder_path):
    for filename in os.listdir(folder_path):
        if filename.startswith("page") and "_img2" in filename and filename.endswith(('.png', '.jpeg', '.jpg')):
            os.remove(os.path.join(folder_path, filename))

# 将图片转换为PDF
def images_to_pdf(image_folder, output_pdf, target_width=595.2, target_height=841.8):
    # 提取文件名中的页码，用于排序
    def extract_number(filename):
        match = re.search(r'page(\d+)', filename)
        return int(match.group(1)) if match else float('inf')

    # 获取并排序所有图片文件
    image_files = sorted(
        [os.path.join(image_folder, f) for f in os.listdir(image_folder) if f.endswith(('.png', '.jpeg', '.jpg'))],
        key=lambda x: extract_number(os.path.basename(x))
    )

    # 创建PDF文档
    pdf_document = fitz.open()
    for img_file in image_files:
        img = Image.open(img_file)
        img_width, img_height = img.size
        scale = min(target_width / img_width, target_height / img_height)
        new_width = int(img_width * scale)
        new_height = int(img_height * scale)
        page = pdf_document.new_page(width=target_width, height=target_height)
        rect = fitz.Rect(
            (target_width - new_width) // 2,
            (target_height - new_height) // 2,
            (target_width + new_width) // 2,
            (target_height + new_height) // 2
        )
        page.insert_image(rect, filename=img_file)
    pdf_document.save(output_pdf)
    pdf_document.close()

# GUI 主函数
def main():
    root = Tk()
    root.withdraw()

    input_pdf = filedialog.askopenfilename(
        title="选择需要处理的PDF文件",
        filetypes=[("PDF 文件", "*.pdf")]
    )
    if not input_pdf:
        messagebox.showerror("错误", "未选择任何文件！")
        return

    images_folder = "images"
    output_pdf = "output.pdf"

    try:
        extract_images_from_pdf(input_pdf, images_folder)
        remove_watermarked_images(images_folder)
        images_to_pdf(images_folder, output_pdf)
        shutil.rmtree(images_folder)
        messagebox.showinfo("完成", f"文件已成功处理，生成的PDF为：{output_pdf}")
    except Exception as e:
        messagebox.showerror("错误", f"处理过程中出错：{e}")

if __name__ == "__main__":
    main()
