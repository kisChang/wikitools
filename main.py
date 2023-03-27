import base64
import hashlib
import os
import re
import shutil
import sys
import fitz
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

import mammoth
from lxml import etree
from tkinterdnd2 import DND_FILES, TkinterDnD


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
        save_path = img_path + img_name % page.number
        pix.save(save_path)


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

        # 操作栏
        self.opt_frame = ttk.Frame(master)
        self.opt_frame.pack(side=tk.TOP, padx=5, pady=5)

        self.file_opt_var = tk.StringVar()
        self.file_opt_var.set('Only')
        self.file_opt_frame = ttk.Frame(self.opt_frame)
        # self.file_opt_frame.pack(side=tk.TOP, padx=5, pady=5)
        tk.Label(self.file_opt_frame, text="(HTML\MD)操作：") \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(self.file_opt_frame, text='原文输出', value='Source', variable=self.file_opt_var, font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(self.file_opt_frame, text='仅格式化', value='Only', variable=self.file_opt_var, font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(self.file_opt_frame, text='格式化带清理', value='Clear', variable=self.file_opt_var,
                       font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)

        self.pdf_zoom_var = tk.DoubleVar()
        self.pdf_zoom_var.set(2.0)
        self.pdf_zoom_frame = ttk.Frame(self.opt_frame)
        # self.pdf_zoom_frame.pack(side=tk.TOP, padx=5, pady=5)
        tk.Label(self.pdf_zoom_frame, text="(PDF)清晰度：") \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(self.pdf_zoom_frame, text='0.8', value=0.8, variable=self.pdf_zoom_var, font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(self.pdf_zoom_frame, text='1', value=1.0, variable=self.pdf_zoom_var, font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(self.pdf_zoom_frame, text='2', value=2.0, variable=self.pdf_zoom_var, font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(self.pdf_zoom_frame, text='5', value=5.0, variable=self.pdf_zoom_var, font=('楷体', 13)) \
            .pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(self.pdf_zoom_frame, text='10', value=10.0, variable=self.pdf_zoom_var, font=('楷体', 13)) \
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

        copy_button = ttk.Button(bottom_frame, text="保存到文件", width=15)
        copy_button.configure(command=self.save_to_file)
        copy_button.pack(side=tk.RIGHT)

    def toggle_tool_mode(self, *args):
        if self.location_var.get() == "PDF转图片":
            self.pdf_zoom_frame.pack()
            self.file_opt_frame.pack_forget()
        else:
            self.file_opt_frame.pack()
            self.pdf_zoom_frame.pack_forget()

    # 定义文件选择函数
    def select_file(self, event=None):
        filepath = filedialog.askopenfilename(
            filetypes=[('Word files', '*.docx')]
        )
        print(f"Selected file: {filepath}")
        self.open_file(filepath)

    def open_file(self, file_name):
        try:
            if file_name and os.path.isfile(file_name):
                name, extension = os.path.splitext(file_name)
                if self.location_var.get() == "PDF转图片":
                    if extension != '.pdf':
                        tk.messagebox.showerror('操作失败', '请选择Pdf 文件！')
                    else:
                        pdf_image(file_name, zoom_x=self.pdf_zoom_var.get(), zoom_y=self.pdf_zoom_var.get())
                        tk.messagebox.showinfo('操作完成', '请查看文件夹 pdf_imgs')
                else:
                    if extension != '.docx':
                        tk.messagebox.showerror('操作失败', '请选择Docx 文件！')
                    else:
                        self.convert_filepath = file_name
                        self.file_label.config(text=os.path.basename(file_name))
                        # 执行转换
                        self.convert_run()
            else:
                self.convert_filepath = None
                self.file_label.config(text='您选择的不是文件！请重新选择文件或拖放文件至此')
        except Exception as e:
            tk.messagebox.showerror('操作失败', '文件异常: {}'.format(e))

    def convert_run(self):
        if self.convert_filepath is None:
            tk.messagebox.showerror('操作失败', '请先加载文件。')
            self.select_file()
        else:
            try:
                with open(self.convert_filepath, 'rb') as file:
                    if self.location_var.get() == "MarkDown":
                        self.convert_result = mammoth.convert_to_markdown(file)
                        self.fmt_md()
                    else:
                        self.convert_result = mammoth.convert_to_html(file)
                        self.fmt_html()
            except Exception as e:
                print(e)
                tk.messagebox.showerror('操作失败', '处理异常: {}'.format(e))

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
            filename = '{}_{}.png'.format(source_file_name, ind)
            save_to = os.path.join(base_path, filename)
            filedata = base64.b64decode(img_data)
            with open(save_to, 'wb') as f:
                f.write(filedata)
            file_md5 = md5(filedata)
            link = f"/uploads/mindoc/images/m_{file_md5}_r.png)"
            # markdown = markdown.replace(match, link)
            markdown = markdown.replace(match.group(), link)
            # markdown = markdown[:match.start()] + link + markdown[:match.end()]

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
            img.set('src', f'/uploads/mindoc/images/m_{file_md5}_r.png')

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


import mimetypes
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
    root.mainloop()
