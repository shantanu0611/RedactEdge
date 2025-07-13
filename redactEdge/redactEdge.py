import os
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from PIL import Image, ImageTk

from spire.pdf.common import *
from spire.pdf import *

import fitz  # PyMuPDF

class VerticalScrolledFrame(ttk.Frame):
    def __init__(self, parent, *args, **kw):
        super().__init__(parent, *args, **kw)
        vscrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL)
        vscrollbar.pack(fill=tk.Y, side=tk.RIGHT, expand=False)
        canvas = tk.Canvas(self, bd=0, highlightthickness=0, yscrollcommand=vscrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vscrollbar.config(command=canvas.yview)
        canvas.xview_moveto(0)
        canvas.yview_moveto(0)
        self.interior = interior = ttk.Frame(canvas)
        interior_id = canvas.create_window(0, 0, window=interior, anchor=tk.NW)

        def _configure_interior(event):
            size = (interior.winfo_reqwidth(), interior.winfo_reqheight())
            canvas.config(scrollregion="0 0 %s %s" % size)
            if interior.winfo_reqwidth() != canvas.winfo_width():
                canvas.config(width=interior.winfo_reqwidth())
        interior.bind('<Configure>', _configure_interior)

        def _configure_canvas(event):
            if interior.winfo_reqwidth() != canvas.winfo_width():
                canvas.itemconfigure(interior_id, width=canvas.winfo_width())
        canvas.bind('<Configure>', _configure_canvas)

class PDFBatchEditorApp:
    def __init__(self, master):
        self.master = master
        master.title("Batch PDF Text & Image Editor")
        master.geometry("900x900")
        master.resizable(True, True)

        self.scroll_frame = VerticalScrolledFrame(master)
        self.scroll_frame.pack(fill="both", expand=True)
        container = self.scroll_frame.interior

        self.selected_files = []
        self.output_dir = ""
        self.replacement_image_path = ""

        # --- Operation Mode Checkbuttons ---
        self.delete_text_var = tk.BooleanVar()
        self.replace_text_var = tk.BooleanVar()
        self.replace_img_var = tk.BooleanVar()
        self.delete_img_var = tk.BooleanVar()
        self.add_textbox_var = tk.BooleanVar()
        self.delete_table_area_var = tk.BooleanVar()

        self.textbox_position = None
        self.textbox_page_num = None
        self.page_preview_canvas = None
        self.page_preview_hscrollbar = None
        self.page_preview_vscrollbar = None
        self.page_preview_frame = None
        self.tk_img = None
        self.pdf_preview_pix = None
        self.pdf_preview_page_rect = None
        self.current_preview_page = 0
        self.num_preview_pages = 1

        self.rect_start = None
        self.rect_end = None
        self.rect_id = None

        # --- UI Elements ---
        self.upload_btn = tk.Button(container, text="Upload PDF Files", command=self.upload_files, width=20)
        self.upload_btn.pack(pady=10)
        
        self.dir_btn = tk.Button(container, text="Select Output Directory", command=self.select_output_dir, width=20)
        self.dir_btn.pack(pady=10)

        # --- Operation selection ---
        self.mode_frame = tk.LabelFrame(container, text="Operation Mode", padx=10, pady=5)
        self.mode_frame.pack(pady=10)
        tk.Checkbutton(self.mode_frame, text="Delete Text", variable=self.delete_text_var, command=self.toggle_image_btn).pack(anchor="w")
        tk.Checkbutton(self.mode_frame, text="Replace Text", variable=self.replace_text_var, command=self.toggle_image_btn).pack(anchor="w")
        tk.Checkbutton(self.mode_frame, text="Replace Image", variable=self.replace_img_var, command=self.toggle_image_btn).pack(anchor="w")
        tk.Checkbutton(self.mode_frame, text="Delete Image(s)", variable=self.delete_img_var, command=self.toggle_image_btn).pack(anchor="w")
        tk.Checkbutton(self.mode_frame, text="Add Textbox", variable=self.add_textbox_var, command=self.toggle_image_btn).pack(anchor="w")
        tk.Checkbutton(self.mode_frame, text="Delete Table Area", variable=self.delete_table_area_var, command=self.toggle_image_btn).pack(anchor="w")

        self.find_text_var = tk.StringVar()
        self.replace_text_var_str = tk.StringVar()
        self.find_text_label = tk.Label(container, text="Text to find (comma-separated for multiple):")
        self.find_text_label.pack()
        self.find_text_entry = tk.Entry(container, textvariable=self.find_text_var, width=60)
        self.find_text_entry.pack(pady=2)
        self.replace_text_label = tk.Label(container, text="Replacement text (comma-separated for multiple):")
        self.replace_text_label.pack()
        self.replace_text_entry = tk.Entry(container, textvariable=self.replace_text_var_str, width=60)
        self.replace_text_entry.pack(pady=2)

        self.image_index_var = tk.StringVar()
        self.image_index_label = tk.Label(container, text="Image Index to Replace/Delete (per page, 0=first, blank=all):")
        self.image_index_label.pack()
        self.image_index_entry = tk.Entry(container, textvariable=self.image_index_var, width=10)
        self.image_index_entry.pack(pady=2)

        self.image_btn = tk.Button(container, text="Select Replacement Image", command=self.select_image, width=25, state=tk.DISABLED)
        self.image_btn.pack(pady=10)

        # --- Add Textbox widgets (hidden by default) ---
        self.add_text_label = tk.Label(container, text="Text to Add (for Add Textbox mode):")
        self.add_text_box = tk.Text(container, height=5, width=60)
        self.add_text_box.bind("<KeyRelease>", lambda e: self.check_ready())
        self.preview_label = tk.Label(container, text="Click and drag on the preview to select textbox rectangle:")
        self.add_text_label.pack_forget()
        self.add_text_box.pack_forget()
        self.preview_label.pack_forget()

        # --- Delete Table Area widgets (hidden by default) ---
        self.delete_table_label = tk.Label(container, text="Click and drag on the preview to select table area to delete:")
        self.delete_area_btn = tk.Button(container, text="Delete Selected Area", command=self.delete_selected_area)
        self.delete_table_label.pack_forget()
        self.delete_area_btn.pack_forget()

        # --- Page navigation controls ---
        self.page_nav_frame = tk.Frame(container)
        self.prev_page_btn = tk.Button(self.page_nav_frame, text="Previous Page", command=self.prev_preview_page)
        self.next_page_btn = tk.Button(self.page_nav_frame, text="Next Page", command=self.next_preview_page)
        self.goto_page_label = tk.Label(self.page_nav_frame, text="Go to page:")
        self.goto_page_entry = tk.Entry(self.page_nav_frame, width=5)
        self.goto_page_btn = tk.Button(self.page_nav_frame, text="Go", command=self.goto_preview_page)
        self.page_info_label = tk.Label(self.page_nav_frame, text="Page 1/1")
        self.prev_page_btn.pack(side=tk.LEFT, padx=2)
        self.next_page_btn.pack(side=tk.LEFT, padx=2)
        self.goto_page_label.pack(side=tk.LEFT)
        self.goto_page_entry.pack(side=tk.LEFT)
        self.goto_page_btn.pack(side=tk.LEFT, padx=2)
        self.page_info_label.pack(side=tk.LEFT, padx=8)
        self.page_nav_frame.pack_forget()

        self.process_btn = tk.Button(container, text="Process Files", command=self.process_files, width=20, state=tk.DISABLED)
        self.process_btn.pack(pady=10)
        self.save_as_jpeg_var = tk.BooleanVar(value=False)
        self.save_as_jpeg_cb = tk.Checkbutton(container, text="Save processed output as JPEG (in addition to PDF)", variable=self.save_as_jpeg_var)
        self.save_as_jpeg_cb.pack(pady=2)

        self.status_label = tk.Label(container, text="Status:", anchor="w")
        self.status_label.pack(fill="x", padx=10, pady=(20, 0))

        self.status_text = scrolledtext.ScrolledText(container, height=14, width=95, state=tk.DISABLED)
        self.status_text.pack(padx=10, pady=5)

        self.container = container

    # --- UI/Preview Methods ---
    def toggle_image_btn(self):
        # Show/hide widgets based on which checkbuttons are checked
        if self.add_textbox_var.get():
            self.add_text_label.pack()
            self.add_text_box.pack(pady=2)
            self.preview_label.pack()
            self.delete_table_label.pack_forget()
            self.delete_area_btn.pack_forget()
            self.page_nav_frame.pack(pady=3)
            self.current_preview_page = 0
            self.textbox_position = None
            self.textbox_page_num = None
            self.show_pdf_preview()
        elif self.delete_table_area_var.get():
            self.add_text_label.pack_forget()
            self.add_text_box.pack_forget()
            self.preview_label.pack_forget()
            self.delete_table_label.pack()
            self.delete_area_btn.pack(pady=2)
            self.page_nav_frame.pack(pady=3)
            self.current_preview_page = 0
            self.textbox_position = None
            self.textbox_page_num = None
            self.show_pdf_preview()
        else:
            self.add_text_label.pack_forget()
            self.add_text_box.pack_forget()
            self.preview_label.pack_forget()
            self.delete_table_label.pack_forget()
            self.delete_area_btn.pack_forget()
            self.page_nav_frame.pack_forget()
            self.hide_pdf_preview()
        if self.replace_img_var.get():
            self.image_btn.config(state=tk.NORMAL)
        else:
            self.image_btn.config(state=tk.DISABLED)
        self.check_ready()

    def show_pdf_preview(self):
        if not self.selected_files:
            return
        pdf_path = self.selected_files[0]
        doc = fitz.open(pdf_path)
        self.num_preview_pages = len(doc)
        page_num = self.current_preview_page
        if page_num >= self.num_preview_pages:
            page_num = 0
            self.current_preview_page = 0
        page = doc[page_num]
        self.pdf_preview_page_rect = page.rect
        pix = page.get_pixmap(dpi=120)
        self.pdf_preview_pix = pix
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.tk_img = ImageTk.PhotoImage(img)
        doc.close()

        if self.page_preview_frame:
            self.page_preview_frame.destroy()
            self.page_preview_frame = None
        self.page_preview_canvas = None
        self.page_preview_hscrollbar = None
        self.page_preview_vscrollbar = None

        self.page_preview_frame = tk.Frame(self.container)
        self.page_preview_frame.pack(pady=5, fill=tk.BOTH, expand=True)

        self.page_preview_canvas = tk.Canvas(
            self.page_preview_frame,
            width=800, height=1000,
            bg="white",
            xscrollincrement=10,
            yscrollincrement=10
        )
        self.page_preview_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.page_preview_canvas.create_image(0, 0, anchor='nw', image=self.tk_img)
        self.page_preview_canvas.config(scrollregion=self.page_preview_canvas.bbox("all"))

        self.page_preview_vscrollbar = tk.Scrollbar(
            self.page_preview_frame, orient='vertical', command=self.page_preview_canvas.yview
        )
        self.page_preview_vscrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.page_preview_canvas.configure(yscrollcommand=self.page_preview_vscrollbar.set)

        self.page_preview_hscrollbar = tk.Scrollbar(
            self.page_preview_frame, orient='horizontal', command=self.page_preview_canvas.xview
        )
        self.page_preview_hscrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.page_preview_canvas.configure(xscrollcommand=self.page_preview_hscrollbar.set)

        self.page_preview_canvas.bind("<Button-1>", self.on_preview_press)
        self.page_preview_canvas.bind("<B1-Motion>", self.on_preview_drag)
        self.page_preview_canvas.bind("<ButtonRelease-1>", self.on_preview_release)
        self.page_info_label.config(text=f"Page {page_num+1}/{self.num_preview_pages}")
        self.check_ready()

    def hide_pdf_preview(self):
        if self.page_preview_frame:
            self.page_preview_frame.destroy()
            self.page_preview_frame = None
        self.page_preview_canvas = None
        self.page_preview_hscrollbar = None
        self.page_preview_vscrollbar = None

    def prev_preview_page(self):
        if self.current_preview_page > 0:
            self.current_preview_page -= 1
            self.textbox_position = None
            self.textbox_page_num = None
            self.show_pdf_preview()

    def next_preview_page(self):
        if self.current_preview_page < self.num_preview_pages - 1:
            self.current_preview_page += 1
            self.textbox_position = None
            self.textbox_page_num = None
            self.show_pdf_preview()

    def goto_preview_page(self):
        try:
            page = int(self.goto_page_entry.get()) - 1
            if 0 <= page < self.num_preview_pages:
                self.current_preview_page = page
                self.textbox_position = None
                self.textbox_page_num = None
                self.show_pdf_preview()
        except Exception:
            pass

    def on_preview_press(self, event):
        self.rect_start = (self.page_preview_canvas.canvasx(event.x), self.page_preview_canvas.canvasy(event.y))
        self.rect_end = self.rect_start
        if self.rect_id:
            self.page_preview_canvas.delete(self.rect_id)
            self.rect_id = None

    def on_preview_drag(self, event):
        self.rect_end = (self.page_preview_canvas.canvasx(event.x), self.page_preview_canvas.canvasy(event.y))
        if self.rect_id:
            self.page_preview_canvas.delete(self.rect_id)
        x0, y0 = self.rect_start
        x1, y1 = self.rect_end
        self.rect_id = self.page_preview_canvas.create_rectangle(x0, y0, x1, y1, outline="red")

    def on_preview_release(self, event):
        self.rect_end = (self.page_preview_canvas.canvasx(event.x), self.page_preview_canvas.canvasy(event.y))
        if self.rect_id:
            self.page_preview_canvas.delete(self.rect_id)
            self.rect_id = None
        x0, y0 = self.rect_start
        x1, y1 = self.rect_end
        x_min, x_max = sorted([x0, x1])
        y_min, y_max = sorted([y0, y1])
        x_ratio = self.pdf_preview_page_rect.width / self.pdf_preview_pix.width
        y_ratio = self.pdf_preview_page_rect.height / self.pdf_preview_pix.height
        pdf_x1 = x_min * x_ratio
        pdf_y1 = y_min * y_ratio
        pdf_x2 = x_max * x_ratio
        pdf_y2 = y_max * y_ratio
        if self.delete_table_area_var.get():
            self.table_rect_pdf = (pdf_x1, pdf_y1, pdf_x2, pdf_y2)
            self.table_rect_page = self.current_preview_page
            self.log_status(f"Table area rectangle set: ({pdf_x1:.1f}, {pdf_y1:.1f}) to ({pdf_x2:.1f}, {pdf_y2:.1f}) on page {self.current_preview_page+1}")
        else:
            if abs(pdf_x2 - pdf_x1) < 5 or abs(pdf_y2 - pdf_y1) < 5:
                self.textbox_position = None
                self.textbox_page_num = None
                self.log_status("Textbox rectangle is too small. Please drag a larger area.")
            else:
                self.textbox_position = (pdf_x1, pdf_y1, pdf_x2, pdf_y2)
                self.textbox_page_num = self.current_preview_page
                self.log_status(f"Textbox rectangle set: ({pdf_x1:.1f}, {pdf_y1:.1f}) to ({pdf_x2:.1f}, {pdf_y2:.1f}) on page {self.current_preview_page+1}")
        self.check_ready()

    def delete_selected_area(self):
        if not hasattr(self, "table_rect_pdf") or not hasattr(self, "table_rect_page"):
            messagebox.showwarning("Warning", "Please select an area to delete.")
            return
        pdf_path = self.selected_files[0]
        doc = fitz.open(pdf_path)
        page = doc[self.table_rect_page]
        x0, y0, x1, y1 = self.table_rect_pdf
        rect = fitz.Rect(x0, y0, x1, y1)
        page.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1))
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
            temp_path = tmp_file.name
        doc.save(temp_path)
        doc.close()
        self.selected_files[0] = temp_path
        self.show_pdf_preview()
        messagebox.showinfo("Info", "Selected area deleted for preview! Remember to process files to save.")

    def upload_files(self):
        files = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if files:
            self.selected_files = list(files)
            self.log_status(f"Selected {len(self.selected_files)} file(s).")
            self.check_ready()
            if self.add_textbox_var.get() or self.delete_table_area_var.get():
                self.current_preview_page = 0
                self.textbox_position = None
                self.textbox_page_num = None
                self.show_pdf_preview()
        else:
            self.log_status("No files selected.")

    def select_output_dir(self):
        directory = filedialog.askdirectory(
            title="Select Output Directory for Modified PDFs"
        )
        if directory:
            self.output_dir = directory
            self.log_status(f"Output directory set to: {self.output_dir}")
            self.check_ready()
        else:
            self.log_status("No output directory selected.")

    def select_image(self):
        file_path = filedialog.askopenfilename(
            title="Select Replacement Image",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp"), ("All files", "*.*")]
        )
        if file_path:
            self.replacement_image_path = file_path
            self.log_status(f"Replacement image selected: {file_path}")
            self.check_ready()
        else:
            self.log_status("No replacement image selected.")

    def check_ready(self):
        ready = bool(self.selected_files and self.output_dir)
        if self.replace_img_var.get():
            ready = ready and bool(self.replacement_image_path)
        if self.add_textbox_var.get():
            ready = ready and bool(self.add_text_box.get("1.0", "end-1c").strip()) and self.textbox_position is not None and self.textbox_page_num is not None
        if self.replace_text_var.get():
            ready = ready and bool(self.find_text_var.get().strip()) and bool(self.replace_text_var_str.get().strip())
        if self.delete_text_var.get():
            ready = ready and bool(self.find_text_var.get().strip())
        if self.delete_table_area_var.get():
            ready = ready and hasattr(self, "table_rect_pdf") and hasattr(self, "table_rect_page")
        self.process_btn.config(state=tk.NORMAL if ready else tk.DISABLED)

    def hide_spire_watermark(self, pdf_path):
        doc = fitz.open(pdf_path)
        for page in doc:
            page_width = page.rect.width
            rect_height = 25
            rect = fitz.Rect(0, 0, page_width, rect_height)
            page.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1))
        temp_path = pdf_path + ".nowatermark.pdf"
        doc.save(temp_path)
        doc.close()
        os.replace(temp_path, pdf_path)

    def replace_text_in_pdf(self, pdf_path, output_path, pairs):
        """Replace text using Spire.PDF (case-insensitive, all occurrences)."""
        try:
            doc = PdfDocument()
            doc.LoadFromFile(pdf_path)
            for find_text, replace_text in pairs:
                for i in range(doc.Pages.Count):
                    page = doc.Pages[i]
                    replacer = PdfTextReplacer(page)
                    replacer.ReplaceAllText(find_text, replace_text)
            doc.SaveToFile(output_path)
            doc.Close()
            self.hide_spire_watermark(output_path)
            return True
        except Exception as e:
            self.log_status(f"Error replacing text: {e}")
            return False

    def replace_images_in_pdf(self, pdf_path, output_path, replacement_image_path, image_index=None):
        """Replace images using PyMuPDF."""
        try:
            doc = fitz.open(pdf_path)
            for page in doc:
                images = page.get_images(full=True)
                if image_index is not None:
                    if 0 <= image_index < len(images):
                        xref = images[image_index][0]
                        page.replace_image(xref, filename=replacement_image_path)
                else:
                    for img in images:
                        xref = img[0]
                        page.replace_image(xref, filename=replacement_image_path)
            doc.save(output_path)
            if getattr(self, 'save_as_jpeg_var', None) and self.save_as_jpeg_var.get():
                try:
                    processed_doc = fitz.open(output_path)
                    for i, page in enumerate(processed_doc):
                        pix = page.get_pixmap(dpi=200)
                        jpeg_path = os.path.join(self.output_dir, f"{os.path.splitext(os.path.basename(output_path))[0]}_page_{i+1}.jpg")
                        pix.save(jpeg_path, "jpeg")
                    processed_doc.close()
                    self.log_status(f"Saved {i+1} JPEG(s) for {output_path}")
                except Exception as e:
                    self.log_status(f"Error saving JPEG(s) for {output_path}: {e}")
            doc.close()
            return True
        except Exception as e:
            self.log_status(f"Error replacing images: {e}")
            return False

    def delete_images_in_pdf(self, pdf_path, output_path, image_index=None):
        """Delete images using PyMuPDF"""
        try:
            doc = fitz.open(pdf_path)
            for page in doc:
                images = page.get_images(full=True)
                if image_index is not None:
                    if 0 <= image_index < len(images):
                        xref = images[image_index][0]
                        page.delete_image(xref)
                else:
                    for img in images:
                        xref = img[0]
                        page.delete_image(xref)
            doc.save(output_path)
            if getattr(self, 'save_as_jpeg_var', None) and self.save_as_jpeg_var.get():
                try:
                    processed_doc = fitz.open(output_path)
                    for i, page in enumerate(processed_doc):
                        pix = page.get_pixmap(dpi=200)
                        jpeg_path = os.path.join(self.output_dir, f"{os.path.splitext(os.path.basename(output_path))[0]}_page_{i+1}.jpg")
                        pix.save(jpeg_path, "jpeg")
                    processed_doc.close()
                    self.log_status(f"Saved {i+1} JPEG(s) for {output_path}")
                except Exception as e:
                    self.log_status(f"Error saving JPEG(s) for {output_path}: {e}")
            doc.close()
            return True
        except Exception as e:
            self.log_status(f"Error deleting images: {e}")
            return False

    def add_textbox_to_pdf(self, pdf_path, output_path, text, position, page_num):
        """Add static text to a PDF using Spire.PDF for Python."""
        try:
            doc = PdfDocument()
            doc.LoadFromFile(pdf_path)
            if page_num < doc.Pages.Count:
                page = doc.Pages[page_num]
                x1, y1, x2, y2 = position
                font = PdfFont(PdfFontFamily.Helvetica, 12.0)
                brush = PdfBrushes.get_Black()
                page.Canvas.DrawString(text, font, brush, float(x1), float(y1))
            doc.SaveToFile(output_path)
            doc.Close()
            self.hide_spire_watermark(output_path)
            return True
        except Exception as e:
            self.log_status(f"Error adding textbox: {e}")
            return False

    def delete_text_in_pdf(self, pdf_path, output_path, delete_texts):
        """Delete text using PyMuPDF redaction annotations."""
        try:
            doc = fitz.open(pdf_path)
            for page in doc:
                for txt in delete_texts:
                    rects = page.search_for(txt)
                    for rect in rects:
                        page.add_redact_annot(rect)
                page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
            doc.save(output_path, garbage=3, deflate=True)
            doc.close()
            return True
        except Exception as e:
            self.log_status(f"Error deleting text: {e}")
            return False

    def process_files(self):
        processed = 0

        find_texts = [x.strip() for x in self.find_text_var.get().split(",") if x.strip()]
        replace_texts = [x.strip() for x in self.replace_text_var_str.get().split(",") if x.strip()]
        pairs = list(zip(find_texts, replace_texts))
        image_index_str = self.image_index_var.get().strip()
        image_index = int(image_index_str) if image_index_str.isdigit() else None

        for file_path in self.selected_files:
            try:
                base_name = os.path.basename(file_path)
                name, ext = os.path.splitext(base_name)
                input_path = file_path
                temp_files = []

                # 1. Delete Text
                if self.delete_text_var.get():
                    temp_path = input_path + ".deltext.pdf"
                    if self.delete_text_in_pdf(input_path, temp_path, find_texts):
                        input_path = temp_path
                        temp_files.append(temp_path)
                        self.log_status(f"Deleted text(s) {find_texts} in: {input_path}")

                # 2. Replace Text
                if self.replace_text_var.get():
                    temp_path = input_path + ".replacetext.pdf"
                    if self.replace_text_in_pdf(input_path, temp_path, pairs):
                        input_path = temp_path
                        temp_files.append(temp_path)
                        self.log_status(f"Text replacement completed: {input_path}")

                # 3. Replace Image
                if self.replace_img_var.get():
                    temp_path = input_path + ".replaceimg.pdf"
                    if self.replace_images_in_pdf(input_path, temp_path, self.replacement_image_path, image_index):
                        input_path = temp_path
                        temp_files.append(temp_path)
                        self.log_status(f"Image replacement completed: {input_path}")

                # 4. Delete Image
                if self.delete_img_var.get():
                    temp_path = input_path + ".delimg.pdf"
                    if self.delete_images_in_pdf(input_path, temp_path, image_index):
                        input_path = temp_path
                        temp_files.append(temp_path)
                        self.log_status(f"Image deletion completed: {input_path}")

                # 5. Add Textbox
                if self.add_textbox_var.get():
                    temp_path = input_path + ".addtextbox.pdf"
                    textbox_text = self.add_text_box.get("1.0", "end-1c").strip()
                    if self.add_textbox_to_pdf(input_path, temp_path, textbox_text, self.textbox_position, self.textbox_page_num):
                        input_path = temp_path
                        temp_files.append(temp_path)
                        self.log_status(f"Textbox addition completed: {input_path}")

                # 6. Delete Table Area
                if self.delete_table_area_var.get():
                    doc = fitz.open(input_path)
                    page_num = getattr(self, "table_rect_page", 0)
                    x0, y0, x1, y1 = getattr(self, "table_rect_pdf", (0, 0, 0, 0))
                    if page_num >= len(doc):
                        page_num = 0
                    rect = fitz.Rect(x0, y0, x1, y1)
                    page = doc[page_num]
                    page.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1))
                    temp_path = input_path + ".deltable.pdf"
                    doc.save(temp_path)
                    doc.close()
                    input_path = temp_path
                    temp_files.append(temp_path)
                    self.log_status(f"Deleted table area on page {page_num+1} and saved: {input_path}")

                # Final output
                output_path = os.path.join(self.output_dir, f"{name}_modified{ext}")
                os.replace(input_path, output_path)
                self.log_status(f"Processed file saved: {output_path}")
                processed += 1
                if self.save_as_jpeg_var.get():
                    try:
                        doc = fitz.open(output_path)
                        for i, page in enumerate(doc):
                            pix = page.get_pixmap(dpi=200)
                            jpeg_path = os.path.join(
                                self.output_dir,
                                f"{os.path.splitext(os.path.basename(output_path))[0]}_page_{i+1}.jpg"
            )
                            # Use pil_save to ensure JPEG format (PyMuPDF v1.22+)
                            pix.pil_save(jpeg_path, format="JPEG")
                        doc.close()
                        self.log_status(f"Saved {i+1} JPEG(s) for {output_path}")
                    except Exception as e:
                        self.log_status(f"Error saving JPEG(s) for {output_path}: {e}")

                # Clean up temp files (except output)
                for temp_file in temp_files:
                    if os.path.exists(temp_file) and temp_file != output_path:
                        os.remove(temp_file)

            except Exception as e:
                self.log_status(f"Error processing {file_path}: {e}")

        messagebox.showinfo("Batch Processing Complete", f"Processed {processed} file(s).")
        self.log_status("Batch processing complete.")

    def log_status(self, message):
        self.status_text.config(state=tk.NORMAL)
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFBatchEditorApp(root)
    root.mainloop()
