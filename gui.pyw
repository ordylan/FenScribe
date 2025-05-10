import win32com.client
import os
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import threading
import fitz  # PyMuPDF
from PIL import Image, ImageTk
from docx import Document
import subprocess
import sys
from tkinterdnd2 import TkinterDnD, DND_FILES 


class PDFProcessorApp:
    def __init__(self, master):
        self.master = master
        master.title("FenScribe ~ PDF空白删")
        master.geometry("800x688")

        # 初始化配置参数
        self.config = {
            'threshold': 137,
            'blank_height': 10,
            'min_height': 10,
            'bottom_margin': 0,
            'dpi': 300,
            'template_docx': ""
        }

        # 创建界面组件
        self.create_widgets()
        self.load_templates()
        self.load_macros()

        # 状态变量
        self.is_processing = False
        self.progress = 0

       # 绑定拖放事件
        self.master.drop_target_register(DND_FILES)  
        self.master.dnd_bind('<<Drop>>', self.handle_drop)  
    def handle_drop(self, event):
        """处理文件拖放事件"""

        files = event.data.strip('{}').split('} {')
        if len(files) > 1:
            messagebox.showerror("Error", "Only one PDF file can be dragged and dropped at a time")
            return
        file_path = files[0].strip()
        if file_path.lower().endswith('.pdf'):
            self.pdf_path.set(file_path)
            self.log(f"Selected file: {file_path}")
        else:
            messagebox.showerror("Error", "The file path should end with .pdf")

    def create_widgets(self):
        # 主容器
        main_frame = ttk.Frame(self.master)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="Step1: Select the PDF file")
        file_frame.pack(fill="x", pady=5)

        self.pdf_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.pdf_path, width=60).pack(side="left", padx=5, fill="x", expand=True)
        ttk.Button(file_frame, text="Browse...", command=self.select_pdf).pack(side="right", padx=5)

        # 参数设置区域
        param_frame = ttk.LabelFrame(main_frame, text="Step2. Set the processing parameters")
        param_frame.pack(fill="x", pady=5)

        ttk.Label(param_frame, text="Blank threshold (20-255):").grid(row=0, column=0, padx=5, sticky="w")
        self.threshold = ttk.Spinbox(param_frame, from_=20, to=255, width=8)
        self.threshold.set(137)
        self.threshold.grid(row=0, column=1, padx=5)

        ttk.Label(param_frame, text="DPI (72-600):").grid(row=0, column=2, padx=5, sticky="w")
        self.dpi = ttk.Spinbox(param_frame, from_=72, to=600, width=8)
        self.dpi.set(300)
        self.dpi.grid(row=0, column=3, padx=5)


        ttk.Label(param_frame, text="min_height:").grid(row=0, column=4, padx=5, sticky="w")
        self.min_height = ttk.Spinbox(param_frame, from_=1, to=300, width=8)
        self.min_height.set(10)
        self.min_height.grid(row=0, column=5, padx=5)

        ttk.Label(param_frame, text="blank_height:").grid(row=0, column=6, padx=5, sticky="w")
        self.blank_height = ttk.Spinbox(param_frame, from_=1, to=300, width=8)
        self.blank_height.set(10)
        self.blank_height.grid(row=0, column=7, padx=5)





        # 模板选择区域
        template_frame = ttk.LabelFrame(main_frame, text="Step3. Select an output template")
        template_frame.pack(fill="x", pady=5)

        self.template_var = tk.StringVar()
        self.template_combobox = ttk.Combobox(template_frame, textvariable=self.template_var, state="readonly")
        self.template_combobox.pack(pady=5, fill="x", padx=5)
        macro_frame = ttk.LabelFrame(main_frame, text="Step4. Select a macro (optional)")
        macro_frame.pack(fill="x", pady=5)

        self.macro_var = tk.StringVar()
        self.macro_combobox = ttk.Combobox(macro_frame, textvariable=self.macro_var, state="readonly")
        self.macro_combobox.pack(pady=5, fill="x", padx=5)
        # 参数测试区域
        #test_frame = ttk.LabelFrame(main_frame, text="参数测试(已坏)")
        #test_frame.pack(fill="x", pady=5)

        #ttk.Label(test_frame, text="测试页码:").pack(side="left", padx=5)
        #self.page_entry = ttk.Entry(test_frame, width=10)
        #self.page_entry.pack(side="left", padx=5)
        #self.test_btn = ttk.Button(test_frame, text="测试", command=self.start_test)
        #self.test_btn.pack(side="left", padx=5)

        # 图片显示区域
        #img_display_frame = ttk.Frame(main_frame)
        #img_display_frame.pack(fill="both", expand=True, pady=10)

        #self.left_image_label = ttk.Label(img_display_frame)
        #self.left_image_label.pack(side="left", fill="both", expand=True, padx=5)
        #self.right_image_label = ttk.Label(img_display_frame)
        #self.right_image_label.pack(side="right", fill="both", expand=True, padx=5)




        # 进度条
        self.progress_bar = ttk.Progressbar(main_frame, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", pady=5)

        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="Process logs")
        log_frame.pack(fill="both", expand=True, pady=5)

        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=6)
        self.log_area.pack(fill="both", expand=True)

        # 控制按钮
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(pady=10)

        self.process_btn = ttk.Button(control_frame, text="Start processing", command=self.start_processing)
        self.process_btn.pack(side="left", padx=5)

        ttk.Button(control_frame, text="Open Output Folder", command=self.openf).pack(side="left", padx=5)
        ttk.Button(control_frame, text="DoubleColumnCut", command=self.open_double_column).pack(side="left", padx=5)
        ttk.Button(control_frame, text="Clear the log", command=self.clear_log).pack(side="left", padx=5)

        ttk.Button(control_frame, text="Quit", command=self.master.quit).pack(side="right", padx=5)
        self.log(f"""
注意: 请先手动裁边pdf哦哦! pdf扫描版需要矫正角度; 多栏pdf请分成单栏进行切割哦! 
这是第三代代码 已经支持gui操作!  图标暂时与程序无关 只是放一下

欢迎: 以下是参数说明: 
1. threshold: 判断某一行是否为空白的亮度阈值
   将像素RGB值转为灰度(取平均值), 当整行像素的平均灰度 ≤ 该值时视为空白行

2. dpi: PDF转图片的分辨率  |   PS: 我们不采用智能换行排版有以下原因: 技术有限  格式错乱概率更大. 

3. min_height: 有效内容过滤阈值
   切割后的段落高度 ≥ 该值才会被保留, 否则视为无效内容

4. blank_height: 识别段落分隔的基准线
   当连续空白行数 ≥ 该值时, 认为前后内容属于不同段落
""")

        self.log(f"""
        测试输出默认参数:
        - Threshold: {self.threshold.get()}
        - DPI: {self.dpi.get()}
        - Min Height: {self.min_height.get()}
        - Blank Height: {self.blank_height.get()}
        """)
    def open_double_column(self):
        py_path = os.path.join(os.path.dirname(__file__), "DoubleColumnCut.pyw")
        if os.path.exists(py_path):
            try:
                subprocess.Popen([sys.executable, py_path])
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open: {str(e)}")
        else:
            messagebox.showerror("Error", "DoubleColumnCut.pyw not found!")
    def load_templates(self):
        """加载模板文件"""
        template_dir = "_Templates"
        try:
            if not os.path.exists(template_dir):
                os.makedirs(template_dir)
                self.log(f"A template catalog has been created: {template_dir}")
                
            templates = [f for f in os.listdir(template_dir) if f.endswith('.docx')]
            if templates:
                self.template_combobox['values'] = templates
                self.template_var.set(templates[0])
                self.config['template_docx'] = os.path.join(template_dir, templates[0])
            else:
                self.log("Tip: There are no template files in the template directory, please place .docx files in the _Templates directory")
        except Exception as e:
                self.log(f"Loading Template Error: {str(e)}")
    def load_macros(self):
        """加载宏文件"""
        macro_dir = "_Macros"
        try:
            if not os.path.exists(macro_dir):
                os.makedirs(macro_dir)
                self.log(f"Macro directory created: {macro_dir}")
                
            macros = [f for f in os.listdir(macro_dir) if f.endswith('.vba')]
            macros.insert(0, "Do not run")  # 默认选项
            self.macro_combobox['values'] = macros
            self.macro_var.set("Do not run")
        except Exception as e:
            self.log(f"Error loading macros: {str(e)}")
    def select_pdf(self):
        """选择PDF文件"""
        file_path = filedialog.askopenfilename(
            filetypes=[("PDF file", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            self.pdf_path.set(file_path)
            self.log(f"Selected file: {file_path}")

    def log(self, message, tag=None):
        """记录日志信息"""
        self.log_area.configure(state="normal")
        self.log_area.insert("end", message + "\n", tag)
        self.log_area.configure(state="disabled")
        self.log_area.see("end")
        self.master.update_idletasks()

    def clear_log(self):
        """清空日志"""
        self.log_area.configure(state="normal")
        self.log_area.delete(1.0, "end")
        self.log_area.configure(state="disabled")

    def validate_inputs(self):
        """验证输入参数"""
        try:
            if not self.pdf_path.get().endswith('.pdf'):
                raise ValueError("需要加上.pdf后缀名啊啊啊啊")

            self.config.update({
                'threshold': int(self.threshold.get()),
                'dpi': int(self.dpi.get()),
                'template_docx': os.path.join("_Templates", self.template_var.get())
            })

            if not (20 <= self.config['threshold'] <= 255):
               raise ValueError("空白阈值需在20-255之间")
          #  if not (72 <= self.config['dpi'] <= 600):
          #      raise ValueError("DPI需在72-600之间")
                
            return True
        except Exception as e:
            self.log(f"Input Error: {str(e)}", "error")
            messagebox.showerror("Input Error", str(e))
            return False

    def start_processing(self):
        """启动处理线程"""
        if self.is_processing:
            return

        if not self.validate_inputs():
            return

        self.is_processing = True
        self.process_btn.config(state="disabled")
        self.progress_bar["value"] = 0
        threading.Thread(target=self.process_pdf, daemon=True).start()
    def openf(self):
        main_program_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "__Output")
        try:
            subprocess.Popen(f'explorer "{main_program_dir}"')
        except Exception as e:
            self.log(f"失败啊: {e}")
    def start_test(self):
        """启动测试线程"""
        if self.is_processing:
            return
        
        try:
            page_num = int(self.page_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid page number")
            return
        
        pdf_path = self.pdf_path.get()
        if not pdf_path.endswith('.pdf'):
            messagebox.showerror("Error", "需要加上.pdf后缀名啊啊啊啊")
            return
        
        with fitz.open(pdf_path) as doc:
            if page_num < 1 or page_num > len(doc):
                messagebox.showerror("Error", f"Invalid page number, valid range: 1-{len(doc)}")
                return
        
        self.is_processing = True
        self.test_btn.config(state="disabled")
        threading.Thread(target=self.run_test, args=(page_num,), daemon=True).start()

    def run_test(self, page_num):
        """执行测试逻辑"""
        try:
            temp_dir = "_temp_test"
            os.makedirs(temp_dir, exist_ok=True)
            
            pdf_path = self.pdf_path.get()
            doc = fitz.open(pdf_path)
            page = doc.load_page(page_num - 1)
            pix = page.get_pixmap(dpi=self.config['dpi'])
            img_path = os.path.join(temp_dir, f"test_page_{page_num}.png")
            pix.save(img_path)
            doc.close()
            
            retained, removed = self.process_test_image(img_path)
            
            retained_img = self.concat_images(retained, 400) if retained else None
            removed_img = self.concat_images(removed, 400) if removed else None
            
            self.display_images(retained_img, removed_img)
            
            os.remove(img_path)
            for f in retained + removed:
                if os.path.exists(f):
                    os.remove(f)
            
            self.log(f"测试完成: 第 {page_num} 页分析结果已显示")
        except Exception as e:
            self.log(f"测试失败: {str(e)}", "error")
            messagebox.showerror("测试Error", str(e))
        finally:
            self.is_processing = False
            self.master.after(0, lambda: self.test_btn.config(state="normal"))

    def process_test_image(self, img_path):
        """处理测试图片并返回保留/删除的片段路径"""
        retained = []
        removed = []
        
        try:
            image = Image.open(img_path)
            width, height = image.size
            pixels = list(image.getdata())
            bottom_margin = self.config['bottom_margin']
            cropped_image = image.crop((0, 0, width, max(0, height - bottom_margin)))
            segments = self.find_segments(pixels, width, cropped_image.height)
            
            temp_dir = "_temp_test"
            counter = 1
            
            for seg in segments:
                seg_height = seg['end'] - seg['start']
                segment_img = Image.new('RGB', (width, seg_height))
                segment_pixels = pixels[seg['start']*width : seg['end']*width]
                segment_img.putdata(segment_pixels)
                
                output_path = os.path.join(temp_dir, f"seg_{counter}.png")
                if seg_height >= self.config['min_height']:
                    retained.append(output_path)
                else:
                    removed.append(output_path)
                
                segment_img.save(output_path)
                counter += 1
                
        except Exception as e:
            self.log(f"Image Processing Error: {str(e)}", "error")
        
        return retained, removed

    def concat_images(self, image_paths, max_width=None):
        """垂直拼接图片"""
        if not image_paths:
            return None
        
        images = []
        for path in image_paths:
            img = Image.open(path)
            if max_width and img.width > max_width:
                ratio = max_width / img.width
                new_height = int(img.height * ratio)
                img = img.resize((max_width, new_height), Image.Resampling.LANCZOS)
            images.append(img)
        
        total_height = sum(img.height for img in images)
        max_width = max(img.width for img in images)
        
        combined = Image.new('RGB', (max_width, total_height))
        y_offset = 0
        for img in images:
            combined.paste(img, (0, y_offset))
            y_offset += img.height
        
        return combined

    def display_images(self, retained_img, removed_img):
        """更新界面显示图片"""
        def update_display():
            self.left_image_label.config(image='')
            self.right_image_label.config(image='')
            
            if retained_img:
                tk_image = ImageTk.PhotoImage(retained_img)
                self.left_image_label.config(image=tk_image)
                self.left_image_label.image = tk_image
            
            if removed_img:
                tk_image = ImageTk.PhotoImage(removed_img)
                self.right_image_label.config(image=tk_image)
                self.right_image_label.image = tk_image
        
        self.master.after(0, update_display)

    def is_blank(self, row):
        """判断是否为空白行"""
        return all(sum(pixel) / 3 >= self.config['threshold'] for pixel in row)

    def find_segments(self, pixels, width, height):
        """查找有效段落"""
        segments = []
        start = None
        blanks = 0
        blank_height = self.config['blank_height']

        for y in range(height):
            row = [pixels[y * width + x] for x in range(width)]
            if self.is_blank(row):
                blanks += 1
                if start is not None and blanks > blank_height:
                    segments.append({'start': start, 'end': y - blank_height})
                    start = None
            else:
                if start is None:
                    start = max(0, y - 2)
                blanks = 0

        if start is not None:
            segments.append({'start': start, 'end': height})

        return segments

    def pdf_to_imgs(self, pdf_path, output_dir):
        """将PDF转换为图片"""
        try:
            pdf_doc = fitz.open(pdf_path)
            img_paths = []
            
            for page_num in range(len(pdf_doc)):
                page = pdf_doc.load_page(page_num)
                pix = page.get_pixmap(dpi=self.config['dpi'])
                output_path = os.path.join(output_dir, f"page_{page_num+1}.png")
                pix.save(output_path)
                img_paths.append(output_path)
                self.log(f"Pages converted {page_num+1}/{len(pdf_doc)}")
                self.update_progress((page_num+1)/len(pdf_doc)*50)
                
            return img_paths
        except Exception as e:
            self.log(f"Error while converting PDF: {str(e)}", "error")
            return []

    def process_pdf(self):
        """主处理流程"""
        try:
            pdf_path = self.pdf_path.get()
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            output_dir = os.path.join("_temp", f"{base_name}_segments")
            
            os.makedirs(output_dir, exist_ok=True)
            self.log(f"A temporary directory is created: {output_dir}")

            doc = Document(self.config['template_docx'])
            total_pages = len(fitz.open(pdf_path))
            segment_counter = 1

            self.log("Start converting PDF to image...")
            img_paths = self.pdf_to_imgs(pdf_path, output_dir)
            if not img_paths:
                raise ValueError("PDF converting error")

            self.log("Start splitting the picture paragraph...")
            for idx, img_path in enumerate(img_paths):
                self.log(f"Processing the picture {idx+1}/{len(img_paths)}: {img_path}")
                segments, segment_counter = self.save_segments(img_path, output_dir, segment_counter)
                
                for seg_path in segments:
                    doc.add_picture(seg_path)
                    self.log(f"Documents Inserted: {seg_path}")
                
                os.remove(img_path)
                self.log(f"Deleted Temporary Files: {img_path}")
                self.update_progress(50 + (idx+1)/len(img_paths)*50)

            output_docx = f"__Output/output_{base_name}.docx"
            output_dira = os.path.join(os.path.dirname(os.path.abspath(__file__)), "__Output")
            if not os.path.exists(output_dira):
                os.makedirs(output_dira)  
            doc.save(output_docx)
            self.log(f"Processing is complete! The results have been saved to: {os.path.abspath(output_docx)}")
            selected_macro = self.macro_var.get()
            if selected_macro != "Do not run":
                macro_path = os.path.join("_Macros", selected_macro)
                self.inject_and_run_macro(output_docx, macro_path)

            messagebox.showinfo("Processing is complete!", f"The results have been saved to:\n{output_docx}")

        except Exception as e:
            self.log(f"Processing Failure: {str(e)}", "error")
            messagebox.showerror("error", str(e))
        finally:
            self.is_processing = False
            self.master.after(0, lambda: self.process_btn.config(state="normal"))
            self.update_progress(0)
    def inject_and_run_macro(self, docx_path, macro_path):
        """注入并运行选定的宏"""
        try:
            # 读取宏代码
            with open(macro_path, 'r', encoding='utf-8') as f:
                vba_code = f.read()
            
            # 获取宏名称（假设文件名不带扩展名是宏名称）
            #macro_name = os.path.splitext(os.path.basename(macro_path))[0]
            macro_name = 'FenScribeMacro'
            # 启动Word应用
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False  # 后台运行
            try:
            # 打开文档
                doc = word.Documents.Open(os.path.abspath(docx_path))
            
            # 注入宏到临时模块
                vba_project = doc.VBProject
                new_module = vba_project.VBComponents.Add(1)  # 1 表示标准模块
                new_module.CodeModule.AddFromString(vba_code)
                
                # 运行宏
                try:
                    word.Application.Run(macro_name)
                except Exception as e:
                    self.log(f"Error running macro: {e}")
                
                # 保存并关闭（覆盖原文件，不保留宏）
                doc.SaveAs2(
                    FileName=os.path.abspath(docx_path),
                    FileFormat=12,  # wdFormatXMLDocument
                    AddToRecentFiles=False
                )
            finally:
                doc.Close(SaveChanges=False)
                word.Quit()
            self.log(f"Macro executed and document saved: {docx_path}")
        except Exception as e:
            if -2146822220 in e.args:
                self.log("Please enable VBA project access in Word: File → Options → Trust Center → Trust Center Settings → Macro Settings → Trust access to the VBA project object model")
            else:
                self.log(f"Error injecting macro: {e}")
    def update_progress(self, value):
        """更新进度条"""
        self.progress_bar["value"] = value
        self.master.update_idletasks()

    def save_segments(self, image_path, output_dir, counter):
        """保存分割后的图片片段"""
        try:
            image = Image.open(image_path)
            width, height = image.size
            pixels = list(image.getdata())
            bottom_margin = self.config['bottom_margin']
            
            cropped_image = image.crop((0, 0, width, max(0, height - bottom_margin)))
            segments = self.find_segments(pixels, width, cropped_image.height)

            segment_paths = []
            for seg in segments:
                seg_height = seg['end'] - seg['start']
                if seg_height >= self.config['min_height']:
                    segment_img = Image.new('RGB', (width, seg_height))
                    segment_pixels = pixels[seg['start']*width : seg['end']*width]
                    segment_img.putdata(segment_pixels)
                    
                    output_path = os.path.join(output_dir, f"segment_{counter}.png")
                    segment_img.save(output_path, dpi=(self.config['dpi'], self.config['dpi']))
                    segment_paths.append(output_path)
                    self.log(f"Saved segment:  {output_path}")
                    counter += 1

            return segment_paths, counter
        except Exception as e:
            self.log(f"Image Processing Error: {str(e)}", "error")
            return [], counter

if __name__ == "__main__":
    root = TkinterDnD.Tk()  
    app = PDFProcessorApp(root)
    try:
        root.iconbitmap('icon.ico')
    except Exception as e:
        print("Error loading icon:", e)

    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        if file_path.lower().endswith('.pdf'):
            app.pdf_path.set(file_path)
            app.log(f"File was loaded via the command line: {file_path}")
        else:
            messagebox.showerror("Error", "The file path should end with .pdf")    
    # 设置样式
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("TButton", padding=6)
    style.configure("TEntry", padding=5)
    style.configure("TCombobox", padding=5)
    os.makedirs("__Output", exist_ok=True)
    root.mainloop()
