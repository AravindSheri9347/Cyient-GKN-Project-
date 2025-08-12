# Cyient-GKN-Project-
#Image to Text extraction using Easyocr and opencv,for character recognition we are use deep learning model, and save into the excels 
from docx2pdf import convert
import win32com.client
from PIL import Image, ImageTk
from docx import Document
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import docx2txt
import threading
import time
from datetime import datetime
from pathlib import Path
import os
from docx import Document
import io
import fitz  # PyMuPDF
import cv2
import numpy as np
import pandas as pd
import easyocr
from matplotlib import pyplot as plt
from openpyxl import load_workbook
import re
from tkinter import ttk
from collections import defaultdict
from openpyxl import load_workbook, Workbook
from fuzzywuzzy import fuzz
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
import pytesseract
import warnings
from io import StringIO

 
warnings.filterwarnings("ignore")
 
class RedTableExtractorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("GKN PROJECT")
        self.geometry("1300x700")
        self.configure(bg="white")
        self.selected_file = None
        self.zoom_level = 1.0  # Default zoom factor
        self.tree = None
        self.canvas = tk.Canvas(self, bg="white")
        # self.canvas.pack(fill=tk.BOTH, expand=True)
        self.create_titlebar()
        self.create_topbar()
        self.create_main_area()
        self.df = pd.DataFrame()
        self.pdf_doc = None
        self.final_data = []
        # self.cropped_data = []
        # self.final_cropped_data = []
        # self.final_cropped_columns = []
        self.page_index = 0
        self.scale = 1.0
        self.rotation = 0
        self.image_path = None  # store the uploaded PDF path
        self.image_id = None
        self.current_image = None
        self.crop_mode = False
        
        # Crop rectangle vars
        self.start_x = self.start_y = None
        self.rect = None
        # Loading elements
        self.loading_label = tk.Label(self.container, text="", font=("Arial", 16), fg="blue", bg="white")
        self.progress = ttk.Progressbar(self.container, orient="horizontal", length=300, mode="indeterminate")
       
       
    def create_titlebar(self):
        title_frame = tk.Frame(self, bg="#a2c8f7")
        title_frame.pack(fill="x")
        title_frame.configure(height=40)
        title_frame.pack_propagate(False)
        title_label = tk.Label(
            title_frame,
            text="CYIENT",
            font=("Arial", 24, "bold"),
            fg="#2c3e50",
            bg="#a2c8f7",
            borderwidth=0,
            highlightthickness=0
        )
        title_label.pack(pady=5)
        bottom_line = tk.Label(title_frame, bg="#a2c8f7")
        bottom_line.pack(fill="x", padx=20, pady=4)
 
    def create_topbar(self):
        topbar = tk.Frame(self, bg="white")
        topbar.pack(side="top", fill="x")
 
        upload_btn = tk.Button(topbar, text="Upload", command=self.upload_file, bg="#3498db", fg="white", width=10)
        upload_btn.pack(side="left", padx=100, pady=3)
 
        generate_btn = tk.Button(topbar, text="Generate", command=self.generate_action, bg="#2ecc71", fg="white", width=15)
        generate_btn.pack(side="left", padx=100)
 
        zoom_in_btn = tk.Button(topbar, text="➕Zoom In", command=self.zoom_in, bg="#9b59b6", fg="white", width=10)
        zoom_in_btn.pack(side="left", padx=20)
 
        zoom_out_btn = tk.Button(topbar, text="➖ Zoom Out", command=self.zoom_out, bg="#f1c40f", fg="black", width=10)
        zoom_out_btn.pack(side="left", padx=20)

        self.crop_btn = tk.Button(topbar, text="✂Crop Mode: OFF", command=self.toggle_crop_mode, bg="#3498db", fg="white", width=15)
        self.crop_btn.pack(side=tk.LEFT, padx=5)

        submit_btn = tk.Button(topbar, text="➡ Submit", command=self.Cropped_main,bg="#1abc9c", fg="black", width=15)
        submit_btn.pack(side="left", padx=20, pady=3)

        back_btn = tk.Button(topbar, text="Back", command=self.back_action, bg="#e74c3c", fg="white", width=15)
        back_btn.pack(side="left", padx=100)


        self.canvas.bind("<ButtonPress-1>", self.start_crop)
        self.canvas.bind("<B1-Motion>", self.draw_crop_rect)
        self.canvas.bind("<ButtonRelease-1>", self.finish_crop)
 
    def create_main_area(self):
        self.container = tk.Frame(self, bg="white")
        self.container.pack(fill="both", expand=True)
 
        self.container.columnconfigure(0, weight=8)
        self.container.columnconfigure(1, weight=2)
 
        # Left Frame
        self.left_frame = tk.Frame(self.container, bg="white", width=1000, height=600)
        self.left_frame.grid(row=0, column=0, sticky="nsew")
        self.left_frame.grid_propagate(False)
 
        self.canvas = tk.Canvas(self.left_frame, bg="white", width=1000, height=600)
        self.left_frame.grid_rowconfigure(0, weight=1)  # Add this line
        self.left_frame.grid_columnconfigure(0, weight=1)  # Add this line
        self.canvas.grid(row=0, column=0, sticky="nsew")

        self.v_scroll = tk.Scrollbar(self.left_frame, orient="vertical", command=self.canvas.yview, width=20)
        self.v_scroll.grid(row=0, column=1, sticky="ns")
 
        self.h_scroll = tk.Scrollbar(self.left_frame, orient="horizontal", command=self.canvas.xview, width=20)
        self.h_scroll.grid(row=1, column=0, sticky="ew")
 
        self.canvas.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)
        self.canvas.config(scrollregion=(0, 0, 1500, 1500))
        self.scrollable_frame = tk.Frame(self.canvas, bg="white")
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
 
        
        def on_configure(event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            self.canvas.itemconfig(self.canvas_window, width=self.canvas.winfo_width())

        self.scrollable_frame.bind("<Configure>",on_configure)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Shift-MouseWheel>", self._on_shiftwheel)
        self.canvas.bind("<Button-4>", self._on_mousewheel)
        self.canvas.bind("<Button-5>", self._on_mousewheel)
 
        # Right Frame
        self.right_frame = tk.Frame(self.container, bg="white")
        self.right_frame.grid(row=0, column=1, sticky="nsew")
 
        self.df_text = tk.Text(self.right_frame, wrap="none", font=("Courier", 10))
        self.df_text.pack(padx=10, pady=(10, 0), fill="both", expand=True)
 
        self.df_v_scroll = tk.Scrollbar(self.right_frame, orient="vertical", command=self.df_text.yview)
        self.df_h_scroll = tk.Scrollbar(self.right_frame, orient="horizontal", command=self.df_text.xview)
 
        self.df_text.configure(yscrollcommand=self.df_v_scroll.set, xscrollcommand=self.df_h_scroll.set)
 
        self.df_v_scroll.pack(side="right", fill="y")
        self.df_h_scroll.pack(side="bottom", fill="x")
 
        self.download_button = tk.Button(
            self.right_frame, text="Download Output",
            command=self.download_output, bg="#27ae60", fg="white"
        )
        self.download_button.pack(pady=10)
 
        self.update_button = tk.Button(
        self.right_frame, text="Update Output",
        command=self.update_output, bg="#f39c12", fg="white"
        )
        self.update_button.pack(pady=10)
 
 
    def _on_mousewheel(self, event):
        if event.num == 4 or event.delta > 0:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            self.canvas.yview_scroll(1, "units")
 
    def _on_shiftwheel(self, event):
        if event.delta > 0:
            self.canvas.xview_scroll(-1, "units")
        elif event.delta < 0:
            self.canvas.xview_scroll(1, "units")

    def zoom_in(self):
        self.zoom_level += 0.2
        self.display_file()

        # Update scrollregion after zoom
        self.canvas.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def zoom_out(self):
        if self.zoom_level > 0.4:
            self.zoom_level -= 0.2
            self.display_file()

    def start_crop(self, event):
        if self.crop_mode:
            self.start_x = self.canvas.canvasx(event.x)
            self.start_y = self.canvas.canvasy(event.y)
            self.crop_rect = None  # Clear any previous rectangle

    def draw_crop_rect(self, event):
        if self.crop_mode and self.start_x is not None and self.start_y is not None:
            end_x, end_y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
            # Delete the previous rectangle if it exists
            if self.crop_rect:
                self.canvas.delete(self.crop_rect)
            # Draw a new rectangle
            self.crop_rect = self.canvas.create_rectangle(
                self.start_x, self.start_y, end_x, end_y,
                outline='red', dash=(4, 2), width=2
            )

    def finish_crop(self, event):
        if self.crop_mode and self.crop_rect:
            end_x, end_y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
            x1, y1 = min(self.start_x, end_x), min(self.start_y, end_y)
            x2, y2 = max(self.start_x, end_x), max(self.start_y, end_y)

            for i, offset in enumerate(self.page_offsets):
                img = self.page_pil_images[i]
                img_height = img.height

                if offset <= y1 < offset + img_height:
                    local_y1 = y1 - offset
                    local_y2 = y2 - offset
                    cropped_image = img.crop((x1, local_y1, x2, local_y2))

                    cropped_folder = "Cropped_image_folder"
                    os.makedirs(cropped_folder, exist_ok=True)
                    base_name = os.path.basename(self.selected_file)
                    file_name = os.path.splitext(base_name)[0]
                    if not hasattr(self, 'crop_count'):
                        self.crop_count = 1
                    crop_path = os.path.join(cropped_folder, f"{file_name}_cropped_{self.crop_count}.png")
                    cropped_image.save(crop_path)
                    self.crop_count += 1
                    messagebox.showinfo("Saved", f"Cropped image saved to:\n{crop_path}")
                    break
            else:
                messagebox.showwarning("Warning", "Crop area does not match any page.")

        if self.crop_rect:
            self.canvas.delete(self.crop_rect)
            self.crop_rect = None
        self.start_x = None
        self.start_y = None


    def toggle_crop_mode(self):
        self.crop_mode = not self.crop_mode
        if self.crop_mode:
            self.crop_btn.config(text="✂Crop Mode: ON", bg="#e67e22")
        else:
            self.crop_btn.config(text="✂Crop Mode: OFF", bg="#3498db")

 
    def upload_file(self):
        filetypes = [
            ("All Supported Files", "*.pdf *.docx *.doc"),
            ("PDF files", "*.pdf"),
            ("Word files", "*.docx"),
            ("Doc files", "*.doc")
        ]
        self.selected_file = filedialog.askopenfilename(title="Open File", filetypes=filetypes)
        if self.selected_file:
            self.display_file()

 
    def convert_word_to_pdf(self,file_path):
        file_path = os.path.abspath(file_path)
        ext = os.path.splitext(file_path)[1].lower()
        pdf_path = os.path.splitext(file_path)[0] + ".pdf"

        if ext == ".docx":
            try:
                convert(file_path, pdf_path)
                print(f"✅ Converted DOCX to PDF: {pdf_path}")
                return pdf_path
            except Exception as e:
                print(f"❌ Error converting DOCX to PDF: {e}")
                return None

        elif ext == ".doc":
            try:
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
                doc = word.Documents.Open(file_path)
                doc.SaveAs(pdf_path, FileFormat=17)  # PDF format
                doc.Close()
                word.Quit()
                print(f"✅ Converted DOC to PDF: {pdf_path}")
                return pdf_path
            except Exception as e:
                print(f"❌ Error converting DOC to PDF: {e}")
                try:
                    word.Quit()
                except:
                    pass
                return None

        else:
            print("❌ Unsupported file type. Please provide a .doc or .docx file.")
            return None


    def display_file(self):
        # Convert .doc or .docx to .pdf before proceeding
        if self.selected_file.endswith(".doc") or self.selected_file.endswith(".docx"):
            converted_pdf = self.convert_word_to_pdf(self.selected_file)
            if not converted_pdf or not os.path.exists(converted_pdf):
                print("Conversion failed.")
                return
            self.selected_file = converted_pdf  # Now treat it as a PDF

        if self.selected_file.endswith(".pdf"):
            self.pdf_doc = fitz.open(self.selected_file)
            self.page_images = []
            self.page_pil_images = []
            self.page_offsets = []

            y_offset = 0  # Track canvas Y offset
            for page_num in range(len(self.pdf_doc)):
                page = self.pdf_doc.load_page(page_num)
                mat = fitz.Matrix(self.zoom_level, self.zoom_level)
                pix = page.get_pixmap(matrix=mat)
                img_data = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                self.page_pil_images.append(img_data.copy())     # Save the actual PIL image
                self.page_offsets.append(y_offset)               # Track canvas offset

                img_tk = ImageTk.PhotoImage(img_data)
                self.page_images.append(img_tk)
                self.image_id = self.canvas.create_image(0, y_offset, image=img_tk, anchor="nw")
                self.canvas.image = img_tk

                self.canvas.tag_bind(self.image_id, "<ButtonPress-1>", self.start_crop)
                self.canvas.tag_bind(self.image_id, "<B1-Motion>", self.draw_crop_rect)
                self.canvas.tag_bind(self.image_id, "<ButtonRelease-1>", self.finish_crop)

                y_offset += pix.height  # Increase offset for next page
    

    
    def generate_output(self):
        if not self.file_path:
            messagebox.showwarning("No file", "Please upload a PDF or DOCX file first.")
            return

        # Start loading animation
        self.loading_label.config(text="Processing, please wait...")
        self.progress.start()

        # Run the OCR and image processing in a separate thread
        def task():
            try:
                if self.file_path.lower().endswith(".pdf"):
                    # Process cropped and full-page images
                    # pdf_name = os.path.splitext(os.path.basename(self.file_path))[0]
                    cropped_output_df = self.Cropped_main("Cropped_image_folder")
                    self.final_cropped_columns = cropped_output_df

                    full_output_df = self.simulate_generation(self.file_path)
                    self.final_columns = full_output_df

                    # Combine both DataFrames column-wise (align by column names)
                    combined_df = pd.concat([self.final_columns, self.final_cropped_columns], ignore_index=True, sort=False)

                    # Fill NaN with None for clarity
                    combined_df = combined_df.where(pd.notnull(combined_df), None)

                    # Store the final combined DataFrame
                    self.combined_final_df = combined_df

                    # Update the GUI with the combined output
                    self.display_output(combined_df)

                    # Show the second frame
                    self.show_second_frame()

                else:
                    messagebox.showerror("Invalid File", "Only PDF files are supported for processing.")

            finally:
                # Stop loading animation
                self.progress.stop()
                self.loading_label.config(text="")

        threading.Thread(target=task).start()

    def generate_action(self):
        self.df_text.delete("1.0", tk.END)
    
        # Center and show loading
        self.loading_label.place(relx=0.5, rely=0.5, anchor="center")
        self.loading_label.config(text="🔄Generating output..Please wait..5mins")

        self.progress.place(relx=0.5, rely=0.55, anchor="center")
        self.progress.start(10)

        # Background task
        threading.Thread(target=self.simulate_generation).start()
 
    # def count_decimals(self, value):
    #     """Count decimal digits after dot in string form of number."""
    #     match = re.match(r'^-?\d*\.(\d+)$', str(value))
    #     return len(match.group(1)) if match else 0
    
    def pdf_to_images(self, pdf_path, output_dir):
        os.makedirs(output_dir, exist_ok=True)
        doc = fitz.open(pdf_path)
        paths = []
        for i, page in enumerate(doc):
            if i < 2:
                continue
            img_path = os.path.join(output_dir, f"page_{i+1}.jpg")
            if not os.path.exists(img_path):
                pix = page.get_pixmap(dpi=300)
                pix.save(img_path)
            paths.append(img_path)
        return paths

    
    def preprocess_for_ocr(self,img):
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, binary = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY+ cv2.THRESH_OTSU)

        resized = cv2.resize(binary, None, fx=2, fy=2, interpolation=cv2.INTER_LINEAR)
        return resized
    
    def process_pdf_for_tables(self, pdf_path, image_dir, tables_dir):
        image_paths = self.pdf_to_images(pdf_path, image_dir)  # ✅ Only 2 arguments now
        for img in image_paths:
            self.extract_tables_from_image(img, tables_dir)

    def extract_tables_from_image(self,image_path, output_dir):
        os.makedirs(output_dir, exist_ok=True)
        image = cv2.imread(image_path)
        if image is None:
            return

        lower_black = np.array([0, 0, 0])
        upper_black = np.array([95, 95, 95])  # allow a bit of tolerance for anti-aliased or compressed blacks
        thresh = cv2.inRange(image, lower_black, upper_black)

        #h_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (55, 1)), iterations=2)
        v_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, 55)), iterations=2)

        verticals,_=cv2.findContours(v_lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        #cv2.imwrite(f"{image_path}_thr.png",thresh)
        redcont=self.red_contours(image)
        count=0
        checkcount=0
        for redcnt in redcont:
            x,y,w,h=cv2.boundingRect(redcnt)
            redx=x+(w//2)
            redy=y+h
            verts=[cnt for cnt in verticals if abs(cv2.boundingRect(cnt)[1]-redy)<6]
            if verts==[]:
                continue
            verts.sort(key=lambda x:abs(cv2.boundingRect(x)[0]-redx))


            linex,_,_,tblh=cv2.boundingRect(verts[0])

            if linex<x+5 or linex>x+w-6:
                continue

            h+=tblh
            rightx=min(image.shape[1],x+w+4)
            x=max(0,x-4)
            w=rightx-x


            if w < 40 or h < 40:
                continue

            roi = thresh[y:y+h, x:x+w]
            checkcount+=1
            h_count, _ = self.count_lines(roi, 'horizontal')
            v_count, _ = self.count_lines(roi, 'vertical')
            if 4 <= h_count <= 8 and v_count == 3:
                cropped = image[y:y+h, x:x+w]
                '''if is_green_or_purple_present(cropped):
                    cv2.rectangle(image, (x, y), (x + w, y + h), (0, 0, 255), 2)
                    cv2.drawContours(image, [redcnt], -1, (0,0,255), -1)
                    cv2.drawContours(image, [verts[0]], -1, (0,0,255), -1)
                    continue'''
                out_path = os.path.join(output_dir, f"{os.path.splitext(os.path.basename(image_path))[0]}_table_{count+1}.jpg")
                cv2.imwrite(out_path, cropped)
                count += 1
            elif v_count!=3:
                '''cv2.rectangle(image, (x, y), (x + w, y + h), (255, 0, 0), 2)
                cv2.drawContours(image, [redcnt], -1, (255,0,0), -1)
                cv2.drawContours(image, [verts[0]], -1, (255,0,0), -1)
                cv2.putText(image, f"h={h_count},v={v_count}", (x,y), cv2.FONT_HERSHEY_SIMPLEX, 1, (255,0,0), 2, cv2.LINE_AA)'''
            else:
                '''cv2.rectangle(image, (x, y), (x + w, y + h), (127, 0, 127), 2)
                cv2.drawContours(image, [redcnt], -1, (127,0,127), -1)
                cv2.drawContours(image, [verts[0]], -1, (127,0,127), -1)
                cv2.putText(image, f"h={h_count},v={v_count}", (x,y), cv2.FONT_HERSHEY_SIMPLEX, 1, (127,0,127), 2, cv2.LINE_AA)'''
        cv2.imwrite(image_path,image)

    def red_contours(self,image):
        hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)

        # Define the range of red color in HSV space
        # Red can have two ranges, one near 0° (lower range) and another near 180° (upper range)
        lower_red1 = np.array([0, 120, 70])
        upper_red1 = np.array([10, 255, 255])

        lower_red2 = np.array([170, 120, 70])
        upper_red2 = np.array([180, 255, 255])

        # Create masks for both red ranges
        mask1 = cv2.inRange(hsv, lower_red1, upper_red1)
        mask2 = cv2.inRange(hsv, lower_red2, upper_red2)

        # Combine both masks
        mask = cv2.bitwise_or(mask1, mask2)


        lower_black = np.array([0, 0, 0])
        upper_black = np.array([127, 127, 180])  # allow a bit of tolerance for anti-aliased or compressed blacks
        blk = cv2.inRange(image, lower_black, upper_black)

        mask=cv2.subtract(mask,blk)

        h_line_mask=cv2.morphologyEx(mask, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (55, 1)), iterations=2)
    
        # Find contours to identify horizontal lines
        contours, _ = cv2.findContours(h_line_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        return contours
    
    def count_lines(self,binary_img, axis='horizontal'):
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (25, 1)) if axis == 'horizontal' else cv2.getStructuringElement(cv2.MORPH_RECT, (1, 25))
        lines = cv2.morphologyEx(binary_img, cv2.MORPH_OPEN, kernel, iterations=2)
        contours, _ = cv2.findContours(lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        contours=list(contours)
        if axis=='vertical':
            contours=self.filter_close_x_contours(contours)
        return len(contours), lines
    
    def filter_close_x_contours(self,contours, threshold=4):
        contours.sort(key=lambda x:cv2.boundingRect(x)[0])
        x_coords = [cv2.boundingRect(c)[0] for c in contours]
        to_remove = set()
        length=len(contours)

        for i, x1 in enumerate(x_coords):
            x1+=cv2.boundingRect(contours[i])[2]-1
            if i in to_remove:
                continue
            for j in range(i+1,length):
                x2=x_coords[j]
                if abs(x1 - x2) <= threshold:
                    to_remove.add(j)
                else:
                    break

        # Keep only contours not marked for removal
        filtered = [c for i, c in enumerate(contours) if i not in to_remove]
        return filtered
    
    def extract_cleaned_table_dataframe(self, tables_folder):
        reader = easyocr.Reader(
        lang_list=['en'],
        recog_network="best_accuracy",
        download_enabled=False
        )
        data_rows = []
        columns = ["Sketch_No", "Measured(mm)", "Nominal(mm)", "+Tol(mm)", "-Tol(mm)", "Deviation(mm)", "OT(mm)", "Thickness Ratio", "Reduction %"]
        failcount=0

        os.makedirs(f"{tables_folder}/alltbs", exist_ok=True)
        os.makedirs(f"{tables_folder}/traindata", exist_ok=True)
        labelfile=open(f"{tables_folder}/traindata/labels.txt","a")

        for fname in sorted(os.listdir(tables_folder)):
            if not fname.lower().endswith(('.jpg', '.png')):
                continue
            img_path = os.path.join(tables_folder, fname)
            img = cv2.imread(img_path)
            if img is None:
                continue

            ocr_img = self.preprocess_for_ocr(img)
            cv2.imwrite(f"{img_path}_pre.bmp", ocr_img)
            results = reader.readtext(ocr_img, detail=1)

            rows = []
            for bbox, text, _ in results:
                x = sum([pt[0] for pt in bbox]) / 4
                y = sum([pt[1] for pt in bbox]) / 4
                rows.append({'text': text, 'cx': x, 'cy': y})

            sorted_rows, used = [], [False]*len(rows)
            for i, r in enumerate(rows):
                if used[i]: continue
                group = [r]; used[i] = True
                for j in range(i+1, len(rows)):
                    if used[j]: continue
                    if abs(rows[j]['cy'] - r['cy']) < 15:
                        group.append(rows[j])
                        used[j] = True
                sorted_rows.append([t['text'] for t in sorted(group, key=lambda x: x['cx'])])

            values = []

            def clean_value(val):
                replacements = {'S': '5', 's': '5','Q':'0','-Q':'-0','q': '0', 'o': '0', 'O': '0', 'e': '8', 'B': '8', ',':'.',' ,':'.','j': '0','w':'0','U':'0','u':'0','D':'0','€':'6','l':'1','W':'0','p':'0','&':'0','.,':'.',',.':'.',',.':'.','+':'0','f':'6','Z':'2','T':'7','I':'1', 'J':'0',
                            'QQ':'00','~':'-'," ":"",'@':'0'}#_ with . ??
                for char, repl in replacements.items():
                    val = val.replace(char, repl)
                return val

            numrows=len(sorted_rows)
            for i in range(0,numrows):
                row=sorted_rows[i]
                if len(row) < 2:
                    continue
                val = clean_value(".".join(row[1:]))
                val = val.replace(' ,', '.').replace(',.', '.').replace(',', '.')
                val = re.sub(r"\.+", ".", val)
                val = val[0] + re.sub(r"-", '.', val[1:len(val)])
                if not re.search(r'\.', val):
                    if len(val) > 3:
                        val = val[0:len(val)-3] + '.' + val[len(val)-3:]
                    elif len(val) == 3:
                        val = val[0:len(val)-2] + '.' + val[len(val)-2:]
                    else:
                        continue
                try:
                    values.append(round(float(val), 3))
                except:
                    # for j in range (1,len(bboxes[i])):
                    continue
                # for j in range (1,len(bboxes[i])):
                #     bbox=bboxes[i][j]
                #     cv2.imwrite(f"{tables_folder}/alltbs/{os.path.basename(img_path)}_{i+1}_{j+1}.png",ocr_img[bbox[0][1]:bbox[3][1],bbox[0][0]:bbox[1][0]])
    
            # ✅ Prevent crash on invalid cases
            if len(values) == 6:
                vals = values
            elif len(values) == 4:
                vals = [values[0], values[1], 0.000, 0.000, values[2], values[3]]
            else:
                continue  # 🛡️ Skip this image if values are invalid

            sketch = os.path.splitext(fname)[0]
            m, n, pt, mt, d, ot = vals
            d = round(m - n, 3)
            ot = round(d - pt, 3) if d > 0 else round(-d - mt, 3)
            ratio = round(m / n, 3) if n != 0 else 0
            red = round((m / n - 1) * 100, 3) if n != 0 else 0
            data_rows.append([sketch, m, n, pt, mt, d, ot, ratio, red])

        df = pd.DataFrame(data_rows, columns=columns)
        if df.empty:
            print("\n No data rows extracted.")
            return df

        df["Page"] = df["Sketch_No"].str.extract(r'page_(\d+)')[0].fillna(0).astype(int)
        for col in columns[1:]:
            df[col] = df[col].astype(float)

        df = df.groupby('Page', group_keys=False).apply(lambda group: self.patch_full_tolerance_pair(group) if any((group['+Tol(mm)'] == 0.0) | (group['-Tol(mm)'] == 0.0)) else group)
        df[["Deviation(mm)", "OT(mm)", "Thickness Ratio", "Reduction %"]] = df.apply(self.recalculate_fields, axis=1)
        return df

    
    def patch_full_tolerance_pair(self,group):
        valid = group[(group['+Tol(mm)'] > 0) & (group['-Tol(mm)'] > 0)]
        ref_plus = valid['+Tol(mm)'].iloc[0] if not valid.empty else 0.000
        ref_minus = valid['-Tol(mm)'].iloc[0] if not valid.empty else 0.000
        def patch(row):
            if row['+Tol(mm)'] == 0.0 or row['-Tol(mm)'] == 0.0:
                row['+Tol(mm)'] = ref_plus
                row['-Tol(mm)'] = ref_minus
            return row
        return group.apply(patch, axis=1)

    def recalculate_fields(self,row):
        m, n, pt, mt = row["Measured(mm)"], row["Nominal(mm)"], row["+Tol(mm)"], row["-Tol(mm)"]
        dev = round(m - n, 3)
        ot = round(dev - pt, 3) if dev > 0 else round(-dev - mt, 3)
        ratio = round(m / n, 3) if n != 0 else 0.0
        red = round((m / n - 1) * 100, 3) if n != 0 else 0.0
        return pd.Series([dev, ot, ratio, red])
    
    def bytes_to_cv2(self,image_bytes):
        return cv2.imdecode(np.frombuffer(image_bytes, np.uint8), cv2.IMREAD_COLOR)

    def extract_yellow_box_images(self,pdf_path, output_folder):
        doc = fitz.open(pdf_path)
        area_counts = defaultdict(int)
        extracted_images = []
 
        for i, page in enumerate(doc):
            page_number = i 
            # if i in [11,12]:
            for img_index, img in enumerate(page.get_images(full=True)):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image = self.bytes_to_cv2(base_image["image"])

                heading_text, matched, area_code = self.extract_yellow_heading_text(image)
                if matched:
                    area_counts[area_code] += 1
                    filename = f"{area_code}.png" if area_counts[area_code] == 1 else f"{area_code}_{area_counts[area_code]}.png"
                    full_path = os.path.join(output_folder, filename)
                    cv2.imwrite(full_path, image)
                    extracted_images.append((full_path,page_number))
                    # print(f" Saved page no {page_number} yellow image: {filename}")
                    self.df["Area"]=area_code#area
        return extracted_images
    
    def extract_yellow_heading_text(self,image):
        pytesseract.pytesseract.tesseract_cmd = r"Tesseract-OCR\tesseract.exe"
        KEYWORDS = ['POSITION', 'THICKNESS', 'AREA']
        hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
        mask = cv2.inRange(hsv, np.array([20, 100, 100]), np.array([40, 255, 255]))
        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        yellow_boxes = [cv2.boundingRect(c) for c in contours if cv2.contourArea(c) > 500]
 
        for (x, y, w, h) in sorted(yellow_boxes, key=lambda b: b[1]):
            cropped = image[y:y+h, x:x+w]
            gray = cv2.cvtColor(cropped, cv2.COLOR_BGR2GRAY)
            blurred = cv2.GaussianBlur(gray, (3, 3), 0)
            thresh = cv2.adaptiveThreshold(blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                        cv2.THRESH_BINARY_INV, 11, 2)
            inverted = cv2.bitwise_not(thresh)
            text = pytesseract.image_to_string(inverted, config='--oem 3 --psm 6').strip()
 
            if self.fuzzy_match_heading(text, KEYWORDS):
                area_id = self.extract_area_number(text)
                return text, True, area_id
           
        return "", False, None
    
    def fuzzy_match_heading(self,text, keywords, threshold=0.5, fuzz_ratio=80):
        lines = text.upper().splitlines()
        for line in lines:
            match_count = sum(1 for k in keywords if fuzz.partial_ratio(k.strip().upper(), line) >= fuzz_ratio)
            # if (match_count / len(keywords)) >= threshold:
            if match_count>0:
                return True
        return False
    
    def extract_area_number(self,text):
        matches = re.findall(r'AREA\s*(\d+(?:[\s_]+\d+)*)', text, re.IGNORECASE)
        if matches:
            pattern = re.sub(r'\s+', '_', matches[0].strip())
            return f"AREA_{pattern}"
        return "AREA"
    
    # --- Table Detection from PDF ---
    def detect_and_crop_table(self,image):
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        blurred = cv2.GaussianBlur(gray, (5, 5), 0)
        edged = cv2.Canny(blurred, 50, 150)
        contours, _ = cv2.findContours(edged, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        contours = sorted(contours, key=cv2.contourArea, reverse=True)

        for cnt in contours:
            approx = cv2.approxPolyDP(cnt, 0.02 * cv2.arcLength(cnt, True), True)
            x, y, w, h = cv2.boundingRect(approx)
            if w > 300 and h > 200:
                return image[y:y+h, x:x+w]
        return None

    def save_tabular_images_from_pdf(self,pdf_path, output_folder, keywords=None):
        os.makedirs(output_folder, exist_ok=True)
        doc = fitz.open(pdf_path)
        reader = easyocr.Reader(['en'], gpu=False)

        for page_num in range(len(doc)):
            page = doc[page_num]
            images = page.get_images(full=True)

            for img_idx, img in enumerate(images):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]
                image_filename = f"page{page_num+1}_img{img_idx+1}.{image_ext}"
                image_path = os.path.join(output_folder, image_filename)

                with open(image_path, "wb") as f:
                    f.write(image_bytes)

                image = cv2.imread(image_path)
                if image is None:
                    print(f"⚠️ Could not read image: {image_path}")
                    continue

                image_rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
                results = reader.readtext(image_rgb, detail=0)
                joined_text = " ".join(results).upper()

                if keywords and any(keyword.upper() in joined_text for keyword in keywords):
                    cropped_table = self.detect_and_crop_table(image)
                    if cropped_table is not None:
                        cv2.imwrite(image_path, cropped_table)
                        print(f"✅ Saved: {image_filename}")
                    else:
                        print(f"⚠️ No table found in {image_filename}")
                        os.remove(image_path)
                else:
                    os.remove(image_path)

    # --- OCR Extraction Functions ---
    def extract_columnwise_text(self,image_path):
        image = cv2.imread(image_path)
        h, w = image.shape[:2]
        column_percentages = {
            "Dimensional Check Number": (0.00, 0.17),
            "Required Dimension": (0.17, 0.33),
            "Tolerances +": (0.33, 0.47),
            "Tolerances -": (0.47, 0.61),
            "Measured Dimension": (0.61, 0.77),
            "Deviation (*)": (0.77, 1.00),
        }
        columns_ranges = {col: (int(start * w), int(end * w)) for col, (start, end) in column_percentages.items()}
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        blur = cv2.GaussianBlur(gray, (3, 3), 0)
        sharpened = cv2.addWeighted(gray, 1.5, blur, -0.5, 0)
        _, thresh = cv2.threshold(sharpened, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        data = pytesseract.image_to_data(thresh, output_type=pytesseract.Output.DICT)

        column_data = {col: defaultdict(str) for col in columns_ranges}
        for i in range(len(data['text'])):
            text = data['text'][i].strip()
            if not text:
                continue
            x, y = data['left'][i], data['top'][i]
            for col, (x_min, x_max) in columns_ranges.items():
                if x_min <= x < x_max:
                    row_key = y // 10
                    column_data[col][row_key] += text + " "
                    break

        all_row_keys = sorted(set(k for d in column_data.values() for k in d))
        final_data = []
        for key in all_row_keys:
            row = [column_data[col].get(key, "").strip() for col in columns_ranges]
            if any(row):
                final_data.append(row)
        return final_data

    def split_cell_tokens(self,final_data):
        i = 0
        while i < len(final_data):
            row = final_data[i]
            col_idx = 1
            cell = row[col_idx].strip()
            parts = cell.split()
            if len(parts) > 1:
                final_data[i][col_idx] = parts[0]
                for offset, token in enumerate(parts[1:], start=1):
                    new_row = row.copy()
                    new_row[col_idx] = token
                    if i + offset >= len(final_data):
                        final_data.append(new_row)
                    else:
                        final_data.insert(i + offset, new_row)
            i += 1
        return final_data

    def clean_cell(self,text):
        cleaned = re.sub(r"[^A-Za-z0-9.+\- ]", "", text)
        cleaned = re.sub(r"^\.+|\.+$", "", cleaned)
        return cleaned.strip()

    def is_numeric(self,s):
        try:
            float(s)
            return True
        except ValueError:
            return False

    def add_decimal_point(self,text):
        text = text.strip()
        if self.is_numeric(text) and len(text.replace("-", "")) >= 3 and "." not in text:
            is_negative = text.startswith("-")
            num_str = text[1:] if is_negative else text
            num_str = num_str[:-2] + "." + num_str[-2:]
            return "-" + num_str if is_negative else num_str
        return text

    def merge_sparse_rows(self,df):
        df = df.copy()
        i = 0
        while i < len(df) - 1:
            non_empty_cells = df.iloc[i].astype(str).str.strip().replace('', np.nan).dropna()
            if len(non_empty_cells) <= 2:
                for col in non_empty_cells.index:
                    df.at[i + 1, col] = (str(non_empty_cells[col]) + " " + str(df.at[i + 1, col])).strip()
                df = df.drop(df.index[i]).reset_index(drop=True)
            else:
                i += 1
        return df

    def recalc_deviation(self,df):
        for i, row in df.iterrows():
            measured = str(row["Measured Dimension"]).strip()
            required = str(row["Required Dimension"]).strip()
            deviation = ""
            if self.is_numeric(measured) and self.is_numeric(required):
                try:
                    deviation_value = float(measured) - float(required)
                    deviation = f"{deviation_value:.2f}"
                except:
                    deviation = ""
            df.at[i, "Deviation"] = deviation
        return df

    # --- Main Function ---
    def extract_tables_from_pdf_to_excel(self, pdf_path, output_folder=None, output_excel_path=None, calculate_deviation=True, keywords=None):
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]

        if output_folder is None:
            output_folder = f"{base_name}_Full_table_images"

        if output_excel_path is None:
            output_excel_path = f"{base_name}_Full_table_output.xlsx"

        if keywords is None:
            keywords = ["DEVIATION", "DIMENSIONAL", "TOLERANCES", "MEASURED"]

        self.save_tabular_images_from_pdf(pdf_path, output_folder, keywords=keywords)

        # Check if any images were saved (i.e., tables found)
        table_images = [f for f in os.listdir(output_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]

        if not table_images:
            print("⚠️ No table images found in PDF. Skipping Excel generation.")
            return  # Exit or handle as needed
        else:
            print(f"\n📁 Table images saved in: {output_folder}")


        all_dataframes = []
        for filename in sorted(os.listdir(output_folder)):
            if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
                image_path = os.path.join(output_folder, filename)
                try:
                    final_data = self.extract_columnwise_text(image_path)
                    final_data = self.split_cell_tokens(final_data)

                    # Shift columns if needed
                    converted_data = []
                    for sublist in final_data:
                        if len(sublist) < 6 or sublist[4] != '':
                            converted_data.append(sublist)
                            continue
                        new_sublist = sublist[:-1]
                        new_sublist[4] = sublist[5]
                        converted_data.append(new_sublist)
                    final_data = converted_data

                    # Clean and convert
                    for i in range(len(final_data)):
                        final_data[i] = [self.clean_cell(cell) for cell in final_data[i]]
                        if len(final_data[i]) >= 5:
                            for idx in [1, 2, 3, 4]:
                                final_data[i][idx] = self.add_decimal_point(final_data[i][idx])

                    # Deviation
                    for i in range(len(final_data)):
                        measured = final_data[i][4]
                        required = final_data[i][1]
                        deviation = ""
                        if self.is_numeric(measured) and self.is_numeric(required):
                            try:
                                deviation_value = float(measured) - float(required)
                                deviation = f"{deviation_value:.2f}"
                            except:
                                deviation = ""
                        final_data[i].append(deviation if calculate_deviation else "")

                    # Clean non-numeric
                    for i in range(len(final_data)):
                        for j in range(len(final_data[i])):
                            cell = final_data[i][j]
                            if re.search(r'[A-Za-z]', cell) and not self.is_numeric(cell):
                                final_data[i][j] = ""

                    filtered_data = [row for row in final_data if any(self.is_numeric(x) for x in row)]
                    output_columns = [
                        "Dimensional Check Number",
                        "Required Dimension",
                        "Tolerances +",
                        "Tolerances -",
                        "Measured Dimension",
                        "Deviation"
                    ]
                    filtered_data_cleaned = [row[:len(output_columns)] for row in filtered_data if len(row) >= len(output_columns)]
                    df = pd.DataFrame(filtered_data_cleaned, columns=output_columns)
                    df = self.merge_sparse_rows(df)

                    # Normalize Required Dimension values
                    values = df["Required Dimension"].astype(str).to_list()
                    result, carry_over, i = [], [], 0
                    while i < len(values) or carry_over:
                        parts = carry_over + (values[i].strip().split() if i < len(values) else [])
                        carry_over = parts[1:] if len(parts) > 1 else []
                        result.append(parts[0] if parts else "")
                        i += 1
                    df = df.reindex(range(len(result)))
                    df["Required Dimension"] = result

                    if calculate_deviation:
                        df = self.recalc_deviation(df)

                    all_dataframes.append(df)

                except Exception as e:
                    print(f"❌ Error processing {image_path}: {e}")

        if all_dataframes:
            final_df = pd.concat(all_dataframes, ignore_index=True)
            final_df.to_excel(output_excel_path, index=False)
            print(f"✅📄 Final Excel saved to: {output_excel_path}")
            self.df_text.delete(1.0, tk.END)
            self.df_text.insert(tk.END,"Successfully Full table images saved in excel")#final_df
            self.after(0, self.end_generation)
        else:
            print("⚠️ No data extracted from images.")

    def simulate_generation(self):
        global final_data, selected_file, df_string

        pdf_path = self.selected_file
        pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
        base_output_dir = os.path.join(os.getcwd(), pdf_name)
        os.makedirs(base_output_dir, exist_ok=True)

        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_folder = f"{base_name}_Full_table_images"
        output_excel_path = f"{base_name}_Full_table_output.xlsx"

        temp_dir = os.path.join(base_output_dir, "temp_images")
        table_dir = os.path.join(base_output_dir, "Extracted_Tables")
        os.makedirs(temp_dir, exist_ok=True)
        os.makedirs(table_dir, exist_ok=True)

        #  Extract yellow box images
        yellow_img_dir = os.path.join(base_output_dir, "Extracted_Images")
        os.makedirs(yellow_img_dir, exist_ok=True)
        extracted_images = self.extract_yellow_box_images(pdf_path, yellow_img_dir)

        self.process_pdf_for_tables(pdf_path, temp_dir, table_dir)
        df = self.extract_cleaned_table_dataframe(table_dir)

        if not df.empty:
            # ✅ If yellow boxes were found, extract and insert Area and Location at index 1 and 2
            if extracted_images:
                area_mapping = {}
                location_mapping = {}
                for img_path, page_no in extracted_images:
                    filename = os.path.basename(img_path)
                    match_area = re.match(r"(AREA_\d+)", filename)
                    match_loc = re.search(r"_(\d+)", filename)
                    if match_area and match_loc:
                        area_mapping[int(page_no)+1] = (match_area.group(1), match_loc.group(1))

                # Map Area and Location or fill with default AREA_1
                df['Area'] = df['Page'].map(lambda pg: area_mapping.get(pg, (None,))[0])
                df['Location'] = df['Page'].map(lambda pg: area_mapping.get(pg, (None, ""))[1])

                # Fill empty Area per page with sequential AREA_1, AREA_2...
                for pg in df['Page'].unique():
                    mask = (df['Page'] == pg) & (df['Area'].isnull())
                    count = mask.sum()
                    if count > 0:
                        df.loc[mask, 'Area'] = [f"AREA_{i+1}" for i in range(count)]

                # Insert at desired index
                area_col = df.pop("Area")
                location_col = df.pop("Location")
                df.insert(1, "Area", area_col)
                df.insert(2, "Location", location_col)

            else:
                # If no yellow boxes, fill Area and Location with blanks
                df.insert(1, "Area", ["" for _ in range(len(df))])
                df.insert(2, "Location", ["" for _ in range(len(df))])

            # ✅ Fill Location from Sketch_No suffix
            df['Location'] = df['Sketch_No'].apply(lambda x: re.search(r'_table_(\d+)', x).group(1) if re.search(r'_table_(\d+)', x) else "")

            # ✅ Sort DataFrame by Sketch_No numeric part
            df['__sk_order__'] = df['Sketch_No'].apply(lambda x: [int(num) for num in re.findall(r'\d+', x)] if re.findall(r'\d+', x) else [0])
            df = df.sort_values(by='__sk_order__').drop(columns='__sk_order__').reset_index(drop=True)

            # ✅ Format only numeric columns to .3f
            numeric_cols = [
                "Measured(mm)", "Nominal(mm)", "+Tol(mm)", "-Tol(mm)",
                "Deviation(mm)", "OT(mm)", "Thickness Ratio", "Reduction %"
            ]
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = df[col].apply(lambda x: format(float(x), ".3f"))
            # ✅ Drop Page column before final assignment
            if "Page" in df.columns:
                df.drop(columns=["Page"], inplace=True)
            # ✅ Assign for GUI display
            self.final_data = df.values.tolist()
            self.final_df = df
            self.final_columns = df.columns.tolist()
            self.df_to_string()
            self.after(0, self.end_generation)
        else:
            # if os.path.exists(output_excel_path):
            #     print(" Excel already exists. Skipping processing.")

            if os.path.exists(output_folder) and any(
                f.lower().endswith((".png", ".jpg", ".jpeg")) for f in os.listdir(output_folder)):
                print("Table images already available. Running OCR extraction.")
                self.extract_tables_from_pdf_to_excel(pdf_path, output_folder, output_excel_path)

            else:
                # print("No images found. Extracting from PDF and saving to Excel.")
                self.extract_tables_from_pdf_to_excel(pdf_path, output_folder, output_excel_path)

            # final check
            if os.path.exists(output_excel_path):
                try:
                    df = pd.read_excel(output_excel_path)
                    self.final_data = df.values.tolist()
                    self.final_df = df
                    self.final_columns = df.columns.tolist()
                    self.df_to_string()
                except Exception as e:
                    print(f"⚠️ Error loading Excel: {e}")
                    self.df_text.delete(1.0, tk.END)
                    self.df_text.insert(tk.END, "[Error loading the Excel file]")
            else:
                print("[WARNING] No data extracted. Excel not saved.")
                self.df_text.delete(1.0, tk.END)
                self.df_text.insert(tk.END, "[No table data found in the PDF]")
            
            self.after(0, self.end_generation)


    def df_to_string(self):
        if not hasattr(self, 'final_data') or not self.final_data:
            self.df_text.delete(1.0, tk.END)
            self.df_text.insert(tk.END, "[No data to display]")
            return

        output = ""
        output += "\t".join(self.final_columns) + "\n"
        for row in self.final_data:
            output += "\t".join(map(str, row)) + "\n"

        self.df_text.delete(1.0, tk.END)
        self.df_text.insert(tk.END, output)


    def update_output(self):
        if not hasattr(self, 'df_text'):
            messagebox.showerror("UI Error", "Text widget not found.")
            return

        edited_text = self.df_text.get("1.0", "end-1c").strip()
        if not edited_text:
            messagebox.showwarning("Empty", "Text area is empty.")
            return

        try:
            from io import StringIO
            df = pd.read_csv(StringIO(edited_text), sep="\t")

            expected_columns = ["Sketch_No", "Measured(mm)", "Nominal(mm)", "+Tol(mm)", "-Tol(mm)",
                    "Deviation(mm)", "OT(mm)", "Thickness Ratio", "Reduction %"]

            # Check that all expected columns (excluding "Page") are present,
            # and the total number of columns is between 9 and 12
            if not (set(expected_columns).issubset(df.columns) and 9 <= len(df.columns) <= 12):
                raise ValueError("Text format is invalid or column names changed.")


            # 🔁 Recalculate derived fields only for edited rows
            for idx, row in df.iterrows():
                try:
                    # Convert 4 key columns to float safely
                    m = float(row["Measured(mm)"])
                    n = float(row["Nominal(mm)"])
                    pt = float(row["+Tol(mm)"])
                    mt = float(row["-Tol(mm)"])

                    # 🔁 Call existing function
                    new_vals = self.recalculate_fields(row)

                    # Update derived columns in df
                    df.at[idx, "Deviation(mm)"] = new_vals[0]
                    df.at[idx, "OT(mm)"] = new_vals[1]
                    df.at[idx, "Thickness Ratio"] = new_vals[2]
                    df.at[idx, "Reduction %"] = new_vals[3]

                except Exception as recalc_err:
                    print(f"[WARNING] Row {idx} skipped during recalculation: {recalc_err}")
                    continue

            # ✅ Store updated data
            self.final_df = df
            self.final_data = df.values.tolist()
            self.final_columns = df.columns.tolist()

            # ✅ Update GUI display with .3f formatting
            self.df_string = "\t".join(df.columns.tolist()) + "\n"
            for _, row in df.iterrows():
                formatted_row = []
                for val in row.tolist():
                    if isinstance(val, float):
                        formatted_row.append(f"{val:.3f}")
                    else:
                        formatted_row.append(str(val))
                self.df_string += "\t".join(formatted_row) + "\n"

            self.df_text.delete("1.0", "end")
            self.df_text.insert("end", self.df_string)

            messagebox.showinfo("Success", "Changes updated and recalculated successfully.")

        except Exception as e:
            print("[ERROR]", e)
            messagebox.showerror("Update Failed", f"Invalid data format:\n{e}")
    def preprocess_image(self,image_path):
        image = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
        blurred = cv2.GaussianBlur(image, (5, 5), 0)
        sharp = cv2.addWeighted(image, 1.5, blurred, -0.5, 0)
        _, binary = cv2.threshold(sharp, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        kernel = np.ones((2,2), np.uint8)
        processed = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel)
        processed_image_path = image_path.replace(".jpg", "_processed.jpg")
        cv2.imwrite(processed_image_path, processed)
        return processed_image_path

 
    def Cropped_main(self):
        # pdf_name = os.path.splitext(os.path.basename(self.file_path))[0]
        # cropped_output_df = self.cropped_main(f"{pdf_name}_folder")
        pdf_path=self.selected_file
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            # image_output_dir = f"{base_name}_temp_images"
        image_folder = os.path.join(os.getcwd(), f"Cropped_image_folder")
        os.makedirs(image_folder, exist_ok=True)
        if not os.path.isdir(image_folder):
            messagebox.showerror("Error", f"Folder 'Crop_Images_folder' not found at: {image_folder}")
            return

        image_files = [os.path.join(image_folder, f) for f in os.listdir(image_folder)
                            if f.lower().endswith(('.png', '.jpg', '.jpeg'))]


        if not image_files:
            messagebox.showerror("Error", "No image files found in 'Cropped_image_folder'.")
            return

        self.cropped_data = []
       


        for img_path in image_files:
            table_images = self.red_table(img_path)

            for table_img in table_images:
                processed_img_path = self.preprocess_image(table_img[0])
                extracted_text = self.extract_text_from_image(processed_img_path)

                clean_vals = []
                for val in extracted_text:
                    val = str(val).strip().replace(",", ".")
                    if re.match(r'^\d{1,2} \d{3}$', val):
                        val = val.replace(" ", ".")
                    if val.startswith("."):
                        val = "0" + val
                    elif val.isdigit():
                        val = "0." + val
                    clean_vals.append(val)

                self.cropped_data.append(clean_vals)


        self.cropped_data = [row for row in self.cropped_data if row and any(cell not in [None, "", " "] for cell in row)]
        cleaned_cropped_data = [row[:6] + [''] * (6 - len(row)) for row in self.cropped_data if len(row) >= 4]

        df1 = pd.DataFrame(cleaned_cropped_data, columns=[
            "Measured(mm)", "Nominal(mm)",
            "+Tol(mm)", "-Tol(mm)", "Deviation(mm)", "OT(mm)"
        ])

        for col in df1.columns:
            df1[col] = pd.to_numeric(df1[col], errors='coerce').fillna(0)

        # self.final_cropped_data = []
        for idx, row in df1.iterrows():
            row_crop_str = ["","",""]

            # Calculate Deviation
            deviation = row["Measured(mm)"] - row["Nominal(mm)"]

            # Calculate OT
            if deviation > 0:
                ot = deviation - row["+Tol(mm)"]
            else:
                ot = (-deviation) - row["-Tol(mm)"]

            # Calculate Ratio and Reduction
            nominal = row["Nominal(mm)"]
            if nominal == 0:
                ratio = 0
                reduction_percent = 0
            else:
                ratio = row["Measured(mm)"] / nominal
                reduction_percent = (ratio - 1) * 100

            # Format all values to .3f and store
            row_crop_str.append(f"{row['Measured(mm)']:.3f}")
            row_crop_str.append(f"{row['Nominal(mm)']:.3f}")
            row_crop_str.append(f"{row['+Tol(mm)']:.3f}")
            row_crop_str.append(f"{row['-Tol(mm)']:.3f}")
            row_crop_str.append(f"{deviation:.3f}")         # Deviation(mm)
            row_crop_str.append(f"{ot:.3f}")                # OT(mm)
            row_crop_str.append(f"{ratio:.3f}")             # Thickness Ratio
            row_crop_str.append(f"{reduction_percent:.3f}") # Thickness Increase/Reduction(%)

            self.final_data.append(row_crop_str)

        self.df_to_string()

 
    def end_generation(self):
        self.loading_label.place_forget()
        self.progress.stop()
        self.progress.place_forget()
       
 
    def download_output(self):
        import pandas as pd
        from tkinter import filedialog, messagebox

        default_name = "output.xlsx"
        if hasattr(self, 'selected_file') and self.selected_file:
            base_name = os.path.splitext(os.path.basename(self.selected_file))[0]
            default_name = f"{base_name}_output.xlsx"

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=default_name,
            title="Save Output As"
        )

        if not save_path:
            return  # User canceled save

        try:
            dataframes = []
            if hasattr(self, 'final_df'):
                if isinstance(self.final_df, pd.DataFrame):
                    if not self.final_df.empty:
                        df_to_save = self.final_df.copy()

                        if hasattr(self, 'final_columns'):
                            df_to_save = df_to_save[self.final_columns]

                        # ✅ Format float columns to .3f
                        for col in df_to_save.select_dtypes(include=['float', 'float64']).columns:
                            df_to_save[col] = df_to_save[col].apply(lambda x: format(x, ".3f"))

                        dataframes.append(df_to_save)


            # 🔍 Check cropped data
            if hasattr(self, 'final_cropped_data') and isinstance(self.final_cropped_data, list) and self.final_cropped_data:
                print("✅ final_cropped_data found")
                df_cropped = pd.DataFrame(self.final_cropped_data, columns=self.final_cropped_columns)
                if not df_cropped.empty:
                    dataframes.append(df_cropped)

            # 🔄 Optional fallback: from self.final_data if nothing else
            if not dataframes and hasattr(self, 'final_data') and hasattr(self, 'final_columns'):
                if isinstance(self.final_data, list) and self.final_data:
                    df_fallback = pd.DataFrame(self.final_data, columns=self.final_columns)
                    dataframes.append(df_fallback)

            if not dataframes:
                messagebox.showwarning("No Data", "No output data available to download.")
                return

            combined_df = pd.concat(dataframes, ignore_index=True).fillna("")

            # ✅ Save with .3f formatting
            for col in combined_df.select_dtypes(include=['float', 'float64']).columns:
                combined_df[col] = combined_df[col].apply(lambda x: format(float(x), ".3f"))

            combined_df.to_excel(save_path, index=False)
            messagebox.showinfo("Saved", f"Output saved to:\n{save_path}")

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to save file:\n{e}")
 
    def back_action(self):
        self.df_text.delete("1.0", tk.END)
        self.display_file("1.0", tk.END)
        self.selected_file.delete("1.0", tk.END)
        if self.selected_file:
            self.display_file()

    
    def red_table(self,img_file,page_number=None):
    
        input_image_path = img_file
        output_dir = r"OUTPUTS\Extracted_Tables"
        min_area = 50
        resize_width, resize_height = 500, 800
 
        os.makedirs(output_dir, exist_ok=True)
 
        image = cv2.imread(input_image_path)
        # if image is None:
        #     # print(f"Failed to load image: {input_image_path}")
        #     return []
        image_name = os.path.splitext(os.path.basename(img_file))[0]
        h, w, _ = image.shape
        hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
        lower_red1 = np.array([0, 70, 50])
        upper_red1 = np.array([10, 255, 255])
        lower_red2 = np.array([170, 70, 50])
        upper_red2 = np.array([180, 255, 255])
        mask1 = cv2.inRange(hsv, lower_red1, upper_red1)
        mask2 = cv2.inRange(hsv, lower_red2, upper_red2)
        red_mask = cv2.bitwise_or(mask1, mask2)
        contours, _ = cv2.findContours(red_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        contours = sorted(contours, key=lambda cnt: cv2.boundingRect(cnt)[1])
        table_count = 0
        extracted_tables = []
        reader = easyocr.Reader(['en'])
        for cnt in contours:
            x, y, w_box, h_box = cv2.boundingRect(cnt)
            area = cv2.contourArea(cnt)
            # if cv2.contourArea(cnt) > 50:
                # for table_width in [80, 200]:
                #     table_height=190
            if area > min_area: 
                cv2.drawContours(image, [cnt], -1, (0, 255, 0), 2)              
                table_width = 80
                table_height = 110
 
                x_center = x + w_box // 2
                y_end = min(y + table_height, h)
                x_start = max(x_center - table_width // 2, 0)
                x_end = min(x_start + table_width, image.shape[1])
                cropped_table = image[y:y_end, x_start:x_end]
                cropped_table = cv2.resize(cropped_table, (resize_width, resize_height), interpolation=cv2.INTER_CUBIC)
                cropped_table = cv2.cvtColor(cropped_table, cv2.COLOR_BGR2GRAY)
                blurred = cv2.GaussianBlur(cropped_table, (5, 5), 0)
                sharp = cv2.addWeighted(cropped_table, 1.5, blurred, -0.5, 0)
                _, binary = cv2.threshold(sharp, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                kernel = np.ones((2, 2), np.uint8)
                cropped_table = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel)
                cropped_table = cv2.bitwise_not(cropped_table)
                cropped_table = cv2.dilate(cropped_table, np.ones((3, 3), np.uint8), iterations=1)
                cropped_table = cv2.bitwise_not(cropped_table)
 
                if page_number:
                    ocr_results = reader.readtext(cropped_table, detail=1)
                    matched = False
                    location_name = ''
    
                    # Check for MS row
                    for bbox, text, _ in ocr_results:
                        # final_location = []
                        line = text.strip().upper()
                        if line.startswith("MS") or line.startswith("NM") or line.startswith("DV") or line.startswith("OT"):
                            matched = True
                            # break
    
                    if not matched:
                        # print(f"❌ Skipping table {area} (No valid MS or NM row)")
                        cv2.imwrite(f"{page_number}-{area}.jpg", cropped_table) 
                        continue
                    
                    for bbox, text, _ in ocr_results:
                        clean_text = text.strip().upper().replace(" ", "").replace("_", "")
                        if re.match(r"AREA\d+[A-Z]?", clean_text):
                            location_name = clean_text
                            break
    
                    if not location_name and matched:
                        for nearby_bbox, nearby_text, _ in ocr_results:
                            t = nearby_text.strip().upper().replace(" ", "").replace("_", "")
                            if re.match(r"AREA\d+[A-Z]?", t) and abs(bbox[0][1] - nearby_bbox[0][1]) < 20:
                                location_name = t
                                break
    
                    if not location_name:
                        location_name = f"{image_name}_{table_count + 1}"
    
                # resized_table = cv2.resize(cropped_table, (resize_width, resize_height), interpolation=cv2.INTER_CUBIC)
                table_count += 1
                if page_number==None:
                    output_dir = r"OUTPUTS\cropped_Tables"
                    os.makedirs(output_dir, exist_ok=True)
                    location_name=Path(img_file).name.replace('.jpg','_cropped')
                output_path = os.path.join(output_dir, f'{location_name}.jpg')
                cv2.imwrite(output_path, cropped_table)
                extracted_tables.append((output_path, location_name,page_number))
                print(f"✅ Table {table_count} on page {page_number} with name {location_name}")
            
        cv2.imwrite(f"Extracted_imgs{page_number}.jpg", image)
        return extracted_tables
 
    def extract_text_from_image(self,image_path):
        pattern = r'(\d{0,2})\s*([., ]+\s*(\d{3}))+'
        img = cv2.imread(image_path)
        reader = easyocr.Reader(['en'])
        result = reader.readtext(img, detail=0, low_text=0.35)
        extracted_text = " ".join(result).strip()
        matches = [match.group(1)+'.'+match.group(3) for match in re.finditer(pattern, extracted_text)]
        return matches

 
if __name__ == "__main__":
    app = RedTableExtractorApp()
    app.mainloop()
