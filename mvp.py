import fitz  # PyMuPDF
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import filedialog, messagebox
from pptx import Presentation
from pptx.util import Inches
import os  # 파일 상단에 추가

class PDFCaptureApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 여러 페이지 캡처 → PPT 저장기")

        self.canvas = tk.Canvas(root, cursor="cross")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        control_frame = tk.Frame(root)
        control_frame.pack(fill=tk.X)

        self.load_button = tk.Button(control_frame, text="PDF 열기", command=self.load_pdf)
        self.load_button.pack(side=tk.LEFT)

        self.prev_button = tk.Button(control_frame, text="← 이전", command=self.prev_page, state=tk.DISABLED)
        self.prev_button.pack(side=tk.LEFT)

        self.next_button = tk.Button(control_frame, text="다음 →", command=self.next_page, state=tk.DISABLED)
        self.next_button.pack(side=tk.LEFT)

        self.save_button = tk.Button(control_frame, text="PPT 저장", command=self.save_ppt, state=tk.DISABLED)
        self.save_button.pack(side=tk.RIGHT)

        self.rect_start = None
        self.rect_id = None
        self.image_on_canvas = None
        self.image_path = None
        self.display_scale = 1.0

        self.canvas.bind("<ButtonPress-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)

        self.pdf_path = None
        self.doc = None
        self.current_page_index = 0
        self.capture_data = []  # (page_index, rect)
        self.capture_dir = "capture"  # 캡처 폴더 경로 추가
        if not os.path.exists(self.capture_dir):
            os.makedirs(self.capture_dir)

    def load_pdf(self):
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF 파일", "*.pdf")])
        if not self.pdf_path:
            return

        # PDF 파일명으로 폴더 생성
        pdf_name = os.path.splitext(os.path.basename(self.pdf_path))[0]
        self.current_capture_dir = os.path.join(self.capture_dir, pdf_name)
        if not os.path.exists(self.current_capture_dir):
            os.makedirs(self.current_capture_dir)

        self.doc = fitz.open(self.pdf_path)
        self.current_page_index = 0
        self.capture_data = []
        self.load_page_image()
        self.prev_button.config(state=tk.NORMAL)
        self.next_button.config(state=tk.NORMAL)
        self.save_button.config(state=tk.NORMAL)

    def load_page_image(self):
        page = self.doc.load_page(self.current_page_index)
        pix = page.get_pixmap()
        self.image_path = os.path.join(self.current_capture_dir, "page_preview.png")
        pix.save(self.image_path)

        img = Image.open(self.image_path)
        self.display_scale = 800 / img.width
        img = img.resize((int(img.width * self.display_scale), int(img.height * self.display_scale)))
        self.tk_image_tk = ImageTk.PhotoImage(img)

        self.canvas.config(width=img.width, height=img.height)
        self.canvas.delete("all")
        self.image_on_canvas = self.canvas.create_image(0, 0, anchor="nw", image=self.tk_image_tk)

    def prev_page(self):
        if self.current_page_index > 0:
            self.current_page_index -= 1
            self.load_page_image()

    def next_page(self):
        if self.current_page_index < len(self.doc) - 1:
            self.current_page_index += 1
            self.load_page_image()

    def on_mouse_down(self, event):
        self.rect_start = (event.x, event.y)
        if self.rect_id:
            self.canvas.delete(self.rect_id)
            self.rect_id = None

    def on_mouse_drag(self, event):
        if self.rect_start:
            x0, y0 = self.rect_start
            x1, y1 = event.x, event.y
            if self.rect_id:
                self.canvas.coords(self.rect_id, x0, y0, x1, y1)
            else:
                self.rect_id = self.canvas.create_rectangle(x0, y0, x1, y1, outline="red", width=2)

    def on_mouse_up(self, event):
        if self.rect_start:
            x0, y0 = self.rect_start
            x1, y1 = event.x, event.y
            # 원래 좌표로 환산해서 저장
            scale = 1 / self.display_scale
            rect = fitz.Rect(min(x0, x1) * scale, min(y0, y1) * scale,
                             max(x0, x1) * scale, max(y0, y1) * scale)
            self.capture_data.append((self.current_page_index, rect))
            messagebox.showinfo("캡처 저장됨", f"{self.current_page_index + 1}페이지에서 영역이 저장되었습니다.")
            self.rect_start = None

    def save_ppt(self):
        if not self.capture_data:
            messagebox.showwarning("경고", "선택한 영역이 없습니다.")
            return

        first_page_rect = self.get_first_page_rect()
        if not first_page_rect:
            messagebox.showwarning("경고", "첫 번째 페이지에서 영역을 선택해주세요.")
            return

        prs = Presentation()
        # 모든 페이지에 대해 첫 번째 페이지의 선택 영역 적용
        for page_index in range(len(self.doc)):
            page = self.doc.load_page(page_index)
            pix = page.get_pixmap(clip=first_page_rect)
            img_path = os.path.join(self.current_capture_dir, f"temp_capture_{page_index}.png")
            pix.save(img_path)

            slide = prs.slides.add_slide(prs.slide_layouts[6])
            img = Image.open(img_path)
            width_inch = img.width / 96
            height_inch = img.height / 96
            slide.shapes.add_picture(img_path, Inches(1), Inches(1),
                                     width=Inches(width_inch), height=Inches(height_inch))

        # PPT 파일도 캡처 폴더에 저장
        default_ppt_name = os.path.join(self.current_capture_dir, "presentation.pptx")
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pptx",
            filetypes=[("PowerPoint 파일", "*.pptx")],
            initialfile=os.path.basename(default_ppt_name)
        )
        if save_path:
            prs.save(save_path)
            messagebox.showinfo("저장 완료", f"PPT가 '{save_path}'에 저장되었습니다.")

    def get_first_page_rect(self):
        if not self.capture_data:
            return None
        # 첫 번째 페이지의 선택 영역 반환
        for page_index, rect in self.capture_data:
            if page_index == 0:
                return rect
        return None

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFCaptureApp(root)
    root.mainloop()