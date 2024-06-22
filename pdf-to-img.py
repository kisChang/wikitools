import os
import fitz


def pdf_image(pdf_path, img_path='./pdf_imgs/', img_name='page-%i.png', zoom_x=2.0, zoom_y=2.0):
    if not os.path.exists(img_path):
        os.makedirs(img_path)
    mat = fitz.Matrix(zoom_x, zoom_y)
    doc = fitz.open(pdf_path)
    for page in doc:
        pix = page.get_pixmap(matrix=mat)
        if img_name.count('%i') > 0:
            save_path = img_path + img_name % page.number
        else:
            save_path = img_path + img_name
        pix.save(save_path)


def list_pdfs_to_img(folder_path, out_path):
    files = os.listdir(folder_path)
    for file in files:
        file_path = os.path.join(folder_path, file)
        if os.path.isdir(file_path):
            list_pdfs_to_img(file_path, out_path)
        else:
            print(f"文件: {file}")
            if file.endswith(".pdf"):
                pdf_image(file_path, out_path, file + '.png')


list_pdfs_to_img("C:\\Users\\KisChang\\Desktop\\PDF\\", 'C:\\Users\\KisChang\\Desktop\\PDF_img\\')
