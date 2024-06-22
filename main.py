import base64
import hashlib
import os
import re
import shutil
import sys
import fitz
import _thread
import logging
import tkinter as tk
from tkinter import filedialog, simpledialog, ttk
from tkinter.scrolledtext import ScrolledText

import pythoncom
from win32com import client

import mammoth
from lxml import etree
from tkinterdnd2 import DND_FILES, TkinterDnD

import requests
from bs4 import BeautifulSoup

from PIL import Image
from io import BytesIO

def clear_dir(base_path):
    if os.path.isdir(base_path):
        shutil.rmtree(base_path)
    os.makedirs(base_path)


def get_base_dir():
    if getattr(sys, 'frozen', False):
        print('on exe')
        return os.path.dirname(os.path.realpath(sys.executable))
    else:
        print('on test')
        return os.path.dirname(os.path.abspath(__file__))


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
            pdf_image(file_path, out_path, file + '.png')


def pdf_to_image_pdf(pdf_path, zoom_x=2.0, zoom_y=2.0):
    _dpi = int(75 * zoom_x)
    doc = fitz.open(pdf_path)
    _file = None
    merges = []
    for page in doc:
        img = page.get_pixmap(dpi=_dpi)
        img_bytes = img.pil_tobytes(format="png")
        image = Image.open(BytesIO(img_bytes))
        if _file is None:
            _file = image  # 取第一张图片用于创建PDF文档的首页
        pix: Image.Image = image.quantize(colors=256, method=0).convert('RGB')  # 单张图片压缩处理
        merges.append(pix)

    _file.save(f"image_by_{_dpi}dpi.pdf", "pdf", save_all=True, append_images=merges[1:])


def pdf_compress(_pdf, _dpi=150, _type="png", method=0):
    '''
    本方法适用于纯图片型（包含文字型图片）的PDF文档压缩，可复制型的文字类的PDF文档不建议使用本方法
    :param _pdf: 文件名全路径
    :param _dpi: 转化后图片的像素（范围72-600），默认150，想要清晰点，可以设置成高一点，这个参数直接影响PDF文件大小
                 测试：  纯图片PDF文件（即单个页面就是一个图片，内容不可复制）
                        300dpi，压缩率约为30-50%，即原来大小的30-50%，基本无损，看不出来压缩后导致的分辨率差异
                        200dpi，压缩率约为20-30%，轻微有损
                        150dpi，压缩率约为5-10%，有损，但是基本不影响图片形文字的阅读
    :param _type: 保存格式，默认为png，其他：JPEG, PNM, PGM, PPM, PBM, PAM, PSD, PS
    :param method:  int，图像压缩方法，只支持下面3个选项，默认值是0
                0 : `MEDIANCUT` (median cut)
                1 : `MAXCOVERAGE` (maximum coverage)
                2 : `FASTOCTREE` (fast octree)
    :return:
    '''
    merges = []
    _file = None
    with fitz.open(_pdf) as doc:
        for i, page in enumerate(doc.pages(), start=0):
            img = page.get_pixmap(dpi=_dpi)  # 将PDF页面转化为图片
            img_bytes = img.pil_tobytes(format=_type)  # 将图片转为为bytes对象
            image = Image.open(BytesIO(img_bytes))  # 将bytes对象转为PIL格式的图片对象
            if i == 0:
                _file = image  # 取第一张图片用于创建PDF文档的首页
            pix: Image.Image = image.quantize(colors=256, method=method).convert('RGB')  # 单张图片压缩处理
            merges.append(pix)  # 组装pdf
    _file.save(f"{_pdf.rsplit('.')[0]}_by_{_dpi}dpi.pdf",
               "pdf",  # 用PIL自带的功能保存为PDF格式文件
               save_all=True,
               append_images=merges[1:])
    print("All completed！")


def docx2pdf(file, outFile="./out.pdf"):
    word = client.Dispatch("Word.Application")
    doc = word.Documents.Open(file)
    doc.SaveAs(outFile, 17)
    doc.Close()
    word.Quit()


class GUI:
    def __init__(self, master):
        self.root = master
        self.root.title('工具')
        # 创建两个单选按钮
        self.convert_filepath = None
        self.convert_result = None
        location_frame = tk.Frame(master)
        location_frame.pack(side=tk.TOP, padx=5, pady=5)
        self.location_var = tk.StringVar()
        self.location_var.set('MarkDown')
        self.location_var.trace("w", self.toggle_tool_mode)
        tk.Label(location_frame, text="请选择作业内容：") \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(location_frame, text='HTML', value='HTML', variable=self.location_var, font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(location_frame, text='MarkDown', value='MarkDown', variable=self.location_var, font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(location_frame, text='PDF转图片', value='PDF转图片', variable=self.location_var, font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(location_frame, text='PDF转图片PDF', value='PDF转图片PDF', variable=self.location_var, font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)

        # 操作栏
        self.opt_frame = ttk.Frame(master)
        self.opt_frame.pack(side=tk.TOP, padx=5, pady=5)

        self.file_opt_var = tk.StringVar()
        self.file_opt_var.set('Only')
        self.file_opt_frame = ttk.Frame(self.opt_frame)
        #  第一栏配置信息
        opt_radio_frame = ttk.Frame(self.file_opt_frame)
        opt_radio_frame.pack(side=tk.TOP, padx=5, pady=5)
        tk.Label(opt_radio_frame, text="(HTML\MD)操作：") \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(opt_radio_frame, text='原文输出', value='Source', variable=self.file_opt_var, font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(opt_radio_frame, text='仅格式化', value='Only', variable=self.file_opt_var, font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(opt_radio_frame, text='格式化带清理', value='Clear', variable=self.file_opt_var,
                       font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)
        #  第二栏项目信息
        opt_name_frame = ttk.Frame(self.file_opt_frame)
        opt_name_frame.pack(side=tk.TOP, padx=5, pady=5)
        tk.Label(opt_name_frame, text="(HTML\MD)项目标识：") \
            .pack(side=tk.LEFT, padx=5)
        self.mindoc_key = tk.StringVar(value='mindoc')
        tk.Entry(opt_name_frame, font=('楷体', 15), textvariable=self.mindoc_key, width=40) \
            .pack(side=tk.LEFT, padx=5)

        self.pdf_zoom_var = tk.DoubleVar()
        self.pdf_zoom_var.set(2.0)
        self.pdf_zoom_frame = ttk.Frame(self.opt_frame)

        opt_radio_frame = ttk.Frame(self.pdf_zoom_frame)
        opt_radio_frame.pack(side=tk.TOP, padx=5, pady=5)
        tk.Label(opt_radio_frame, text="(PDF)清晰度：") \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(opt_radio_frame, text='1', value=1.0, variable=self.pdf_zoom_var, font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(opt_radio_frame, text='2', value=2.0, variable=self.pdf_zoom_var, font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(opt_radio_frame, text='3', value=3.0, variable=self.pdf_zoom_var, font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(opt_radio_frame, text='5', value=5.0, variable=self.pdf_zoom_var, font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(opt_radio_frame, text='10', value=10.0, variable=self.pdf_zoom_var, font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)

        opt_radio_frame = ttk.Frame(self.pdf_zoom_frame)
        opt_radio_frame.pack(side=tk.TOP, padx=5, pady=5)
        tk.Radiobutton(opt_radio_frame, text='0.3', value=0.3, variable=self.pdf_zoom_var, font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(opt_radio_frame, text='0.5', value=0.5, variable=self.pdf_zoom_var, font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(opt_radio_frame, text='0.8', value=0.8, variable=self.pdf_zoom_var, font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)

        self.toggle_tool_mode()

        # 创建文件选择和确认按钮
        file_frame = ttk.Frame(master)
        file_frame.pack(fill=tk.BOTH, padx=5)
        self.file_label = ttk.Label(file_frame, text='点击选择或拖拽(Docx\Pdf)文件到此处即可打开', font=('楷体', 15),
                                    anchor="center",
                                    background='#c3d7a8')
        self.file_label.pack(side=tk.LEFT, fill=tk.BOTH, padx=5, expand=True)

        # 启用点击功能
        self.file_label.bind("<Button-1>", self.select_file)
        # 启用拖放功能
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind("<<Drop>>", lambda e: self.open_file(e.data))

        # 创建文本框和复制按钮
        text_frame = ttk.Frame(master)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.text = ScrolledText(text_frame, height=20)
        self.text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 创建Frame用于放置Copy按钮
        bottom_frame = tk.Frame(master)
        bottom_frame.pack(side=tk.BOTTOM, padx=5, pady=5)
        copy_button = ttk.Button(bottom_frame, text="复制到粘贴板", width=15)
        copy_button.configure(command=self.copy_to_clipboard)
        copy_button.pack(side=tk.LEFT)

        mulu_button = ttk.Button(bottom_frame, text="生成文档目录MD", width=15)
        mulu_button.configure(command=self.btn_to_mulu)
        mulu_button.pack(side=tk.RIGHT)

        copy_button = ttk.Button(bottom_frame, text="保存到文件", width=15)
        copy_button.configure(command=self.save_to_file)
        copy_button.pack(side=tk.RIGHT)

    def btn_to_mulu(self):
        url = simpledialog.askstring("生成文档目录", "请输入文档的地址：")
        if url is not None:
            self.to_mulu(url)

    def to_mulu(self, url):
        self.text.delete(1.0, tk.END)  # 清空文本框
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        sidebar = soup.find(id="sidebar")
        for li in sidebar.find_all("a"):
            href = li.attrs.get('href').replace('http://wiki.babyceo.cn', '')
            title = li.text
            line = f'1. [{title}]({href} "{title}")'
            self.log(line)

    def toggle_tool_mode(self, *args):
        if self.location_var.get() == "PDF转图片":
            self.pdf_zoom_frame.pack()
            self.file_opt_frame.pack_forget()
        elif self.location_var.get() == "PDF转图片PDF":
            self.pdf_zoom_frame.pack()
            self.file_opt_frame.pack_forget()
        else:
            self.file_opt_frame.pack()
            self.pdf_zoom_frame.pack_forget()

    # 定义文件选择函数
    def select_file(self, event=None):
        filetypes = [('Word files', '*.docx')]
        if self.location_var.get() == "PDF转图片":
            filetypes = [('PDF files', '*.pdf')]
        elif self.location_var.get() == "PDF转图片PDF":
            filetypes = [('PDF files', '*.pdf')]
        filepath = filedialog.askopenfilename(filetypes=filetypes)
        print(f"Selected file: {filepath}")
        self.open_file(filepath)

    def log(self, str):
        if str is not None:
            self.text.insert(tk.END, str)
            self.text.insert(tk.END, "\n")

    def open_file(self, file_name):
        def fun():
            pythoncom.CoInitialize()
            self.open_file_real(file_name)
        _thread.start_new_thread(fun, ())

    def open_file_real(self, file_name):
        try:
            if len(file_name) <= 1:
                return
            if file_name.startswith("{"):
                file_name = file_name[1:len(file_name) - 1]
            print('file_name', file_name)

            if file_name and os.path.isfile(file_name):
                name, extension = os.path.splitext(file_name)
                extension = extension.lower()
                if extension == '.pdf':
                    if self.location_var.get() == "PDF转图片":
                        self.location_var.set('PDF转图片')
                        self.toggle_tool_mode()
                        pdf_image(file_name, zoom_x=self.pdf_zoom_var.get(), zoom_y=self.pdf_zoom_var.get())
                        tk.messagebox.showinfo('操作完成', '请查看文件夹 pdf_imgs')
                    if self.location_var.get() == "PDF转图片PDF":
                        self.location_var.set('PDF转图片PDF')
                        self.toggle_tool_mode()
                        pdf_to_image_pdf(file_name, zoom_x=self.pdf_zoom_var.get(), zoom_y=self.pdf_zoom_var.get())
                        tk.messagebox.showinfo('操作完成', '请查看文件 image.pdf')
                elif extension == '.docx':
                    if self.location_var.get() == "PDF转图片":
                        pass
                        tmp_pdf = os.path.join(get_base_dir(), 'out.pdf')
                        self.log(f"成功加载文件：{file_name}")
                        self.log(f"1. 将文件转换为PDF")
                        # 先转pdf
                        docx2pdf(file_name, tmp_pdf)
                        self.log(f"Succeed!")
                        # 再把pdf转图片
                        self.log(f"2. 将pdf转图片")
                        pdf_image(tmp_pdf, zoom_x=self.pdf_zoom_var.get(), zoom_y=self.pdf_zoom_var.get())
                        self.log(f"Succeed!")
                        # 再删除pdf
                        self.log(f"3. 清理")
                        os.remove(tmp_pdf)
                        self.log(f"Succeed!")
                        tk.messagebox.showinfo('操作完成', '请查看文件夹 pdf_imgs')
                    else:
                        self.convert_filepath = file_name
                        self.file_label.config(text=os.path.basename(file_name))
                        self.convert_run()  # 执行转换
                else:
                    if self.location_var.get() == "PDF转图片":
                        tk.messagebox.showerror('操作失败', '请选择Pdf 文件！')
                    else:
                        tk.messagebox.showerror('操作失败', '请选择Docx 文件！')
            else:
                self.convert_filepath = None
                list_pdfs_to_img(file_name, './pdf_imgs/')
                self.file_label.config(text='您选择的不是文件！请重新选择文件或拖放文件至此')
        except Exception as e:
            self.log("App Error: {}".format(e))
            logging.exception(e)
            tk.messagebox.showerror('操作失败', '文件异常: {}'.format(e))

    def convert_run(self):
        if self.convert_filepath is None:
            tk.messagebox.showerror('操作失败', '请先加载文件。')
            self.select_file()
        else:
            with open(self.convert_filepath, 'rb') as file:
                if self.location_var.get() == "MarkDown":
                    self.convert_result = mammoth.convert_to_markdown(file)
                    self.fmt_md()
                else:
                    style_maps = """p =>
                    table => table.table.table-striped.table-hover
                    """
                    self.convert_result = mammoth.convert_to_html(file, style_map=style_maps)
                    self.fmt_html()

    def fmt_md(self):
        source_file_name = self.file_label.cget("text")
        base_path = os.path.join(get_base_dir(), 'tool_imgs')
        # 先清理
        clear_dir(base_path)
        ind = 0
        data_uri_regex = re.compile(r"data:image/(.*?);base64,([A-Za-z0-9+/]+=*)([\)]?)")
        markdown = self.convert_result.value
        # 清理空的a标签
        markdown = re.sub(r'<a.*?></a>', " ", markdown)
        # 清理问题字符
        markdown = re.sub(r" {2,}__", "", markdown)
        markdown = re.sub(r"\)__", "", markdown)
        # 处理图片
        for match in data_uri_regex.finditer(markdown):
            image_type = match.group(1)
            img_data = match.group(2)
            ind = ind + 1
            filedata = base64.b64decode(img_data)
            file_md5 = md5(filedata)
            # filename = '{}_{}.png'.format(source_file_name, ind)
            filename = '{}.png'.format(file_md5)
            save_to = os.path.join(base_path, filename)
            with open(save_to, 'wb') as f:
                f.write(filedata)
            link = f"/uploads/{self.mindoc_key.get()}/images/m_{file_md5}_r.png)"
            markdown = markdown.replace(match.group(), link)

        self.convert_result.value = markdown
        self.set_text()
        # self.save_to_file()

    def fmt_html(self):
        html = self.convert_result.value
        # 清理空的a标签
        html = re.sub(r'<a.*?></a>', " ", html)
        # 处理格式化和图片
        parser = etree.HTMLParser(remove_blank_text=True)
        tree = etree.fromstring(html, parser)
        file_name = self.file_label.cget("text")
        base_path = os.path.join(get_base_dir(), 'tool_imgs')
        clear_dir(base_path)
        ind = 0
        for img in tree.xpath("//img"):
            img_data = img.get("src")
            img_data = img_data[img_data.index('base64') + 7:]
            ind = ind + 1
            filename = os.path.join(base_path, '{}_{}.png'.format(file_name, ind))
            filedata = base64.b64decode(img_data)
            with open(filename, "wb") as f:
                f.write(filedata)
            file_md5 = md5(filedata)
            img.set('src', f'/uploads/{self.mindoc_key.get()}/images/m_{file_md5}_r.png')

        if self.file_opt_var.get() == "Source":
            result = html
        elif self.file_opt_var.get() == "Only":
            result = etree.tostring(tree, pretty_print=True, encoding='unicode')
        elif self.file_opt_var.get() == "Clear":
            for elem in tree.xpath('//*[not(node())]'):
                if elem.tag != 'img':
                    elem.getparent().remove(elem)
            result = etree.tostring(tree, pretty_print=True, encoding='unicode')
        self.convert_result.value = result
        self.set_text()
        # self.save_to_file()

    def set_text(self):
        self.text.delete(1.0, tk.END)  # 清空文本框
        self.text.insert(tk.END, self.convert_result.value)  # 显示转换结果

    def save_to_file(self):
        file_name = self.file_label.cget("text")
        file_type = 'md' if self.location_var.get() == "MarkDown" else 'html'
        save_path = os.path.join(get_base_dir(), '{}.{}'.format(file_name, file_type))
        with open(save_path, 'w', encoding="utf-8") as f:
            f.write(self.convert_result.value)
            tk.messagebox.showinfo('操作成功', '文件已存储至: {}'.format(save_path))

    def copy_to_clipboard(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(self.convert_result.value)
        tk.messagebox.showinfo("Copy", "已复制到剪贴板。")


# 定义一个函数来计算文件的MD5哈希值
def md5(filedata):
    hash_md5 = hashlib.md5()
    hash_md5.update(filedata)
    return hash_md5.hexdigest()


import ota
# .py to .exe :
# pyinstaller  --onefile --windowed --additional-hooks-dir=. main.py
if __name__ == '__main__':
    root = TkinterDnD.Tk()
    gui = GUI(root)
    nScreenWid = root.winfo_screenwidth()
    nScreenHei = root.winfo_screenheight()
    nCurWid = 600  # 窗体宽
    nCurHeight = 450  # 窗体高
    geometry = "%dx%d+%d+%d" % (nCurWid, nCurHeight, nScreenWid / 2 - nCurWid / 2, nScreenHei / 2 - nCurHeight / 2)
    root.geometry(geometry)
    ota.check_for_updates("23", root)
    root.mainloop()
