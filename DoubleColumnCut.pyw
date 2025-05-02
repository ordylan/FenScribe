import fitz
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from PIL import Image, ImageTk

class PDFSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF分栏工具-DoubleColumnCut")
        
        self.input_path = ""
        self.split_ratio = 0.5  # 默认50%
        self.page_width = 0
        self.zoom = 1.0  
        
        self.create_widgets()
        self.dragging = False
        self.line_id = None

    def create_widgets(self):
        file_frame = ttk.Frame(self.root)
        file_frame.pack(pady=10, padx=10, fill="x")
        
        ttk.Button(file_frame, text="选择PDF文件", command=self.select_file).pack(side="left")
        self.file_label = ttk.Label(file_frame, text="未选择文件")
        self.file_label.pack(side="left", padx=10)
        
        self.canvas = tk.Canvas(self.root, width=600, height=400, bg="white")
        self.canvas.pack(pady=10)
        
        self.canvas.bind("<Button-1>", self.start_drag)
        self.canvas.bind("<B1-Motion>", self.do_drag)
        self.canvas.bind("<ButtonRelease-1>", self.stop_drag)
        
        info_frame = ttk.Frame(self.root)
        info_frame.pack(pady=5)
        
        ttk.Label(info_frame, text="分栏比例:").pack(side="left")
        self.ratio_label = ttk.Label(info_frame, text="50.0%")
        self.ratio_label.pack(side="left", padx=5)
        
        ttk.Button(self.root, text="开始处理", command=self.process_pdf).pack(pady=10)

    def select_file(self):
        self.input_path = filedialog.askopenfilename(filetypes=[("PDF文件", "*.pdf")])
        if self.input_path:
            self.file_label.config(text=self.input_path)
            self.show_preview()

    def show_preview(self):
        doc = fitz.open(self.input_path)
        page = doc[0]
        

        self.page_width = page.rect.width
        page_height = page.rect.height
        
        self.zoom = min(600 / self.page_width, 400 / page_height)
        
        pix = page.get_pixmap(matrix=fitz.Matrix(self.zoom, self.zoom))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.tk_img = ImageTk.PhotoImage(img)
        
        self.canvas.create_image(0, 0, anchor="nw", image=self.tk_img)
        initial_x = self.page_width * self.zoom * self.split_ratio
        self.line_id = self.canvas.create_line(initial_x, 0, initial_x, 400, fill="red", width=2)
        doc.close()

    def start_drag(self, event):
        if self.line_id:
            self.dragging = True

    def do_drag(self, event):
        if self.dragging and self.page_width > 0:
            x = max(0, min(event.x, self.page_width * self.zoom))
            
            self.canvas.coords(self.line_id, x, 0, x, 400)
            
            self.split_ratio = x / (self.page_width * self.zoom)
            self.ratio_label.config(text=f"{self.split_ratio*100:.1f}%")

    def stop_drag(self, event):
        self.dragging = False

    def process_pdf(self):
        if not self.input_path:
            return
            
        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF文件", "*.pdf")]
        )
        if not output_path:
            return
        
        doc = fitz.open(self.input_path)
        new_doc = fitz.open()
        
        for page in doc:
            original_rect = page.rect
            width = original_rect.width
            height = original_rect.height
            
            split_x = width * self.split_ratio
            
            # 创建左右页面
            new_page_left = new_doc.new_page(width=split_x, height=height)
            new_page_left.show_pdf_page(
                new_page_left.rect,
                doc,
                page.number,
                clip=fitz.Rect(0, 0, split_x, height)
            )
            
            new_page_right = new_doc.new_page(width=width-split_x, height=height)
            new_page_right.show_pdf_page(
                new_page_right.rect,
                doc,
                page.number,
                clip=fitz.Rect(split_x, 0, width, height)
            )
        
        new_doc.save(output_path)
        doc.close()
        messagebox.showinfo("Yes", f"处理完成, 已保存到: {output_path}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFSplitterApp(root)
    root.mainloop()