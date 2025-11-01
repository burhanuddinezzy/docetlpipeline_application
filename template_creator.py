import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import json
import os
import io
import cv2
import numpy as np
import pdfplumber

class TemplateCreator:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Template Creator")
        self.root.geometry("1200x800")
        
        self.pdf_doc = None
        self.pdf_path = None
        self.raw_text = ""
        self.current_page = None
        self.current_page_num = 1
        self.total_pages = 0
        self.canvas = None
        
        # Multi-page template structure
        self.template_data = {
            "template_name": "",
            "pages": {}
        }
        
        # Page-specific data
        self.pages_data = {}  # {page_num: {"raw_text": "", "boxes": []}}
        
        self.drawing = False
        self.start_x = None
        self.start_y = None
        self.current_rect = None
        self.scale_factor = 1.0
        
        # Box type system
        self.current_box_type = tk.StringVar(value="general")
        self.table_sensitivity = tk.DoubleVar(value=0.5)
        self.table_edit_mode = False
        self.detected_lines = {"horizontal": [], "vertical": []}
        self.line_items = []
        
        self.templates_dir = "templates"
        os.makedirs(self.templates_dir, exist_ok=True)

        self.include_unboxed_content = tk.BooleanVar(value=True)
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the user interface."""
        # Control frame
        control_frame = tk.Frame(self.root)
        control_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)
        
        # File operation buttons
        tk.Button(control_frame, text="Load PDF", command=self.load_pdf).pack(side=tk.LEFT, padx=5)
        tk.Button(control_frame, text="Clear Boxes", command=self.clear_boxes).pack(side=tk.LEFT, padx=5)
        tk.Button(control_frame, text="Save Template", command=self.save_template).pack(side=tk.LEFT, padx=5)
        tk.Button(control_frame, text="Load Template", command=self.load_template).pack(side=tk.LEFT, padx=5)
        
        # Page navigation frame
        nav_frame = tk.Frame(self.root)
        nav_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)
        
        # Page navigation controls
        tk.Label(nav_frame, text="Page:").pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(nav_frame, text="◀ Prev", command=self.prev_page, width=8).pack(side=tk.LEFT, padx=2)
        
        # Page input
        self.page_var = tk.StringVar(value="1")
        self.page_entry = tk.Entry(nav_frame, textvariable=self.page_var, width=4, justify='center')
        self.page_entry.pack(side=tk.LEFT, padx=2)
        self.page_entry.bind('<Return>', self.goto_page)
        
        tk.Button(nav_frame, text="Go", command=self.goto_page, width=4).pack(side=tk.LEFT, padx=2)
        tk.Button(nav_frame, text="Next ▶", command=self.next_page, width=8).pack(side=tk.LEFT, padx=2)
        
        # Page info
        self.page_info_var = tk.StringVar(value="Page 1 of 1")
        tk.Label(nav_frame, textvariable=self.page_info_var, font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=(10, 0))
        
        # Boxes on current page info
        self.boxes_info_var = tk.StringVar(value="Boxes: 0")
        tk.Label(nav_frame, textvariable=self.boxes_info_var, font=('Arial', 9)).pack(side=tk.LEFT, padx=(20, 0))
        
        # Box type selection frame
        type_frame = tk.Frame(self.root)
        type_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)
        
        tk.Label(type_frame, text="Box Type:").pack(side=tk.LEFT, padx=(0, 10))
        tk.Radiobutton(type_frame, text="General", variable=self.current_box_type, value="general").pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(type_frame, text="Table", variable=self.current_box_type, value="table").pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(type_frame, text="Paragraph", variable=self.current_box_type, value="paragraph").pack(side=tk.LEFT, padx=5)
        
        # ADD SIMPLE UNBOXED TOGGLE:
        tk.Frame(type_frame, width=2, bg='gray').pack(side=tk.LEFT, fill=tk.Y, padx=15)
        tk.Checkbutton(type_frame, text="Include unboxed content", 
        variable=self.include_unboxed_content).pack(side=tk.LEFT, padx=5)

        # Info frame
        info_frame = tk.Frame(self.root)
        info_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)
        
        self.info_label = tk.Label(info_frame, text="Instructions: Load PDF → Navigate pages → Draw rectangles for data extraction OR save template without boxes for simple reading-order extraction", 
                                  font=('Arial', 10))
        self.info_label.pack(side=tk.LEFT)
        
        # Canvas frame with scrollbars
        canvas_frame = tk.Frame(self.root)
        canvas_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Create canvas with scrollbars
        self.canvas = tk.Canvas(canvas_frame, bg='white', cursor='crosshair')
        v_scrollbar = tk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        h_scrollbar = tk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        
        self.canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Bind mouse events
        self.canvas.bind("<Button-1>", self.start_draw)
        self.canvas.bind("<B1-Motion>", self.draw_rect)
        self.canvas.bind("<ButtonRelease-1>", self.end_draw)
        self.canvas.bind("<Button-3>", self.delete_box)  # Right click to delete
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready - Load a PDF to start creating multi-page template")
        status_bar = tk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def extract_raw_text(self):
        """Extract raw text from the entire PDF and per-page text."""
        if not self.pdf_doc:
            return
        
        # Extract combined text for template matching
        self.raw_text = ""
        
        # Extract per-page text and initialize page data
        self.pages_data = {}
        for page_num in range(len(self.pdf_doc)):
            page = self.pdf_doc[page_num]
            page_text = page.get_text()
            
            # Add to combined text
            self.raw_text += page_text
            
            # Store per-page data
            actual_page_num = page_num + 1  # 1-indexed
            self.pages_data[actual_page_num] = {
                "raw_text": page_text,
                "boxes": []
            }
    
    def load_pdf(self):
        """Load a PDF file for template creation."""
        file_path = filedialog.askopenfilename(
            title="Select PDF for Template Creation",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if not file_path:
            return
        
        try:
            self.pdf_path = file_path
            self.pdf_doc = fitz.open(file_path)
            self.total_pages = len(self.pdf_doc)
            self.current_page_num = 1
            self.current_page = self.pdf_doc[0]  # Use first page
            
            # Extract all text data
            self.extract_raw_text()
            
            # Update UI
            self.update_page_info()
            self.display_page()
            self.status_var.set(f"Loaded: {os.path.basename(file_path)} ({self.total_pages} pages)")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load PDF: {e}")
    
    def update_page_info(self):
        """Update page information display."""
        self.page_info_var.set(f"Page {self.current_page_num} of {self.total_pages}")
        self.page_var.set(str(self.current_page_num))
        
        # Update boxes info
        current_page_boxes = self.get_current_page_boxes()
        self.boxes_info_var.set(f"Boxes: {len(current_page_boxes)}")
    
    def get_current_page_boxes(self):
        """Get boxes for the current page."""
        return self.pages_data.get(self.current_page_num, {}).get("boxes", [])
    
    def prev_page(self):
        """Navigate to previous page."""
        if self.current_page_num > 1:
            self.current_page_num -= 1
            self.current_page = self.pdf_doc[self.current_page_num - 1]
            self.update_page_info()
            self.display_page()
    
    def next_page(self):
        """Navigate to next page."""
        if self.current_page_num < self.total_pages:
            self.current_page_num += 1
            self.current_page = self.pdf_doc[self.current_page_num - 1]
            self.update_page_info()
            self.display_page()
    
    def goto_page(self, event=None):
        """Go to specific page."""
        try:
            page_num = int(self.page_var.get())
            if 1 <= page_num <= self.total_pages:
                self.current_page_num = page_num
                self.current_page = self.pdf_doc[self.current_page_num - 1]
                self.update_page_info()
                self.display_page()
            else:
                messagebox.showwarning("Invalid Page", f"Please enter a page number between 1 and {self.total_pages}")
                self.page_var.set(str(self.current_page_num))
        except ValueError:
            messagebox.showwarning("Invalid Page", "Please enter a valid page number")
            self.page_var.set(str(self.current_page_num))
    
    def display_page(self):
        """Display the current PDF page on canvas."""
        if not self.current_page:
            return
        
        # Render page as image
        mat = fitz.Matrix(2, 2)  # 2x zoom for better quality
        pix = self.current_page.get_pixmap(matrix=mat)
        
        # Convert to PIL Image
        img_data = pix.tobytes("ppm")
        pil_image = Image.open(io.BytesIO(img_data))
        
        # Calculate scale to fit canvas (max 800px width)
        canvas_width = 800
        img_width, img_height = pil_image.size
        
        if img_width > canvas_width:
            self.scale_factor = canvas_width / img_width
            new_width = int(img_width * self.scale_factor)
            new_height = int(img_height * self.scale_factor)
            pil_image = pil_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
        else:
            self.scale_factor = 1.0
        
        # Display on canvas
        self.photo = ImageTk.PhotoImage(pil_image)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)
        
        # Update canvas scroll region
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
        # Redraw boxes for current page
        self.redraw_current_page_boxes()
    
    def redraw_current_page_boxes(self):
        """Redraw boxes for the current page only."""
        current_page_boxes = self.get_current_page_boxes()
        
        for box in current_page_boxes:
            self._draw_box(box)
    
    def start_draw(self, event):
        """Start drawing a rectangle."""
        if not self.current_page:
            return
        
        self.drawing = True
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
    
    def draw_rect(self, event):
        """Draw rectangle as user drags."""
        if not self.drawing:
            return
        
        current_x = self.canvas.canvasx(event.x)
        current_y = self.canvas.canvasy(event.y)
        
        # Delete previous preview rectangle
        if hasattr(self, 'current_rect') and self.current_rect:
            self.canvas.delete(self.current_rect)
        
        # Normal extraction box
        self.current_rect = self.canvas.create_rectangle(
            self.start_x, self.start_y, current_x, current_y,
            outline='red', width=2, fill='', stipple='gray25'
        )
    
    def end_draw(self, event):
        """Finish drawing rectangle and save box."""
        if not self.drawing:
            return
        
        self.drawing = False
        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)
        
        # Ensure we have a valid rectangle
        if abs(end_x - self.start_x) < 10 or abs(end_y - self.start_y) < 10:
            if hasattr(self, 'current_rect') and self.current_rect:
                self.canvas.delete(self.current_rect)
            return
        
        # Convert screen coordinates to PDF coordinates
        pdf_coords = self.screen_to_pdf_coords(self.start_x, self.start_y, end_x, end_y)
        
        # Create normal extraction box
        self._create_extraction_box(pdf_coords, end_x, end_y)
    
    def _create_extraction_box(self, pdf_coords, end_x, end_y):
        """Create an extraction box based on selected type."""
        box_type = self.current_box_type.get()
        
        if box_type == "table":
            self._create_table_box(pdf_coords, end_x, end_y)
        else:
            self._create_standard_box(pdf_coords, end_x, end_y, box_type)
    
    def _create_standard_box(self, pdf_coords, end_x, end_y, box_type):
        """Create a standard (general/paragraph) extraction box."""
        # Get current page boxes for counting
        current_page_boxes = self.get_current_page_boxes()
        
        # Get box label from user
        label = simpledialog.askstring("Box Label", 
                                     f"Enter label for this {box_type} box on page {self.current_page_num}:",
                                     initialvalue=f"{box_type.title()}_{len(current_page_boxes)+1}")
        
        if not label:
            if hasattr(self, 'current_rect') and self.current_rect:
                self.canvas.delete(self.current_rect)
            return
        
        # Create unique box ID with page prefix
        box_id = f"p{self.current_page_num}_box{len(current_page_boxes)+1}"
        
        # Store box data
        box_data = {
            "id": box_id,
            "label": label,
            "page": self.current_page_num,
            "coordinates": pdf_coords,
            "extraction_order": len(current_page_boxes) + 1,
            "screen_coords": [self.start_x, self.start_y, end_x, end_y],
            "box_type": box_type
        }
        
        # Add to current page
        if self.current_page_num not in self.pages_data:
            self.pages_data[self.current_page_num] = {"raw_text": "", "boxes": []}
        
        self.pages_data[self.current_page_num]["boxes"].append(box_data)
        self._draw_box(box_data)
        
        # Update UI
        self.update_page_info()
        self.status_var.set(f"Added {box_type} box: {label} on page {self.current_page_num}")
        
        if hasattr(self, 'current_rect'):
            self.current_rect = None
    
    def _create_table_box(self, pdf_coords, end_x, end_y):
        """Create a table box with line detection."""
        current_page_boxes = self.get_current_page_boxes()
        
        # Get box label first
        label = simpledialog.askstring("Table Box Label", 
                                     f"Enter label for this table box on page {self.current_page_num}:",
                                     initialvalue=f"Table_{len(current_page_boxes)+1}")
        
        if not label:
            if hasattr(self, 'current_rect') and self.current_rect:
                self.canvas.delete(self.current_rect)
            return
        
        # Start table editing mode
        self._enter_table_edit_mode(pdf_coords, end_x, end_y, label)
    
    def _enter_table_edit_mode(self, pdf_coords, end_x, end_y, label):
        """Enter table editing mode with line detection."""
        self.table_edit_mode = True
        
        # Create table editing dialog
        self.table_dialog = tk.Toplevel(self.root)
        self.table_dialog.title(f"Table Editor - {label} (Page {self.current_page_num})")
        self.table_dialog.geometry("350x250")
        self.table_dialog.transient(self.root)
        self.table_dialog.grab_set()
        
        # Status display
        self.line_status = tk.StringVar()
        self.line_status.set("Detecting lines...")
        status_label = tk.Label(self.table_dialog, textvariable=self.line_status, 
                               font=('Arial', 10, 'bold'))
        status_label.pack(padx=10, pady=5)
        
        # Sensitivity controls
        sens_frame = tk.Frame(self.table_dialog)
        sens_frame.pack(padx=10, pady=10, fill=tk.X)
        
        tk.Label(sens_frame, text="Line Detection Sensitivity:").pack()
        sensitivity_scale = tk.Scale(sens_frame, from_=0.1, to=1.0, resolution=0.1, 
                                   orient=tk.HORIZONTAL, variable=self.table_sensitivity,
                                   command=lambda x: self._on_sensitivity_changed(pdf_coords))
        sensitivity_scale.pack(fill=tk.X, pady=5)
        
        # Current sensitivity display
        self.sens_display = tk.StringVar()
        tk.Label(sens_frame, textvariable=self.sens_display).pack()
        
        # Buttons
        btn_frame = tk.Frame(self.table_dialog)
        btn_frame.pack(padx=10, pady=10, fill=tk.X)
        
        tk.Button(btn_frame, text="Re-detect Lines", 
                 command=lambda: self._detect_table_lines(pdf_coords)).pack(pady=2, fill=tk.X)
        tk.Button(btn_frame, text="OK - Save Table", 
                 command=lambda: self._save_table_box(pdf_coords, end_x, end_y, label)).pack(pady=2, fill=tk.X)
        tk.Button(btn_frame, text="Cancel", 
                 command=self._cancel_table_edit).pack(pady=2, fill=tk.X)
        
        # Instructions
        instructions = tk.Text(self.table_dialog, height=4, wrap=tk.WORD)
        instructions.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        instructions.insert("1.0", 
            "Instructions:\n"
            "• Red lines = Horizontal table lines\n"
            "• Blue lines = Vertical table lines\n"
            "• Adjust sensitivity slider to fine-tune detection\n"
            "• Higher sensitivity = more lines detected")
        instructions.config(state=tk.DISABLED)
        
        # Initial line detection
        self._detect_table_lines(pdf_coords)
        self._update_sensitivity_display()
    
    def _on_sensitivity_changed(self, pdf_coords):
        """Handle sensitivity slider change with auto re-detection."""
        self._update_sensitivity_display()
        # Auto re-detect lines when sensitivity changes
        self.root.after(300, lambda: self._detect_table_lines(pdf_coords))  # Debounce 300ms
    
    def _update_sensitivity_display(self):
        """Update sensitivity display in dialog."""
        if hasattr(self, 'sens_display'):
            sens_val = self.table_sensitivity.get()
            self.sens_display.set(f"Current: {sens_val:.1f}")
    
    def _update_line_status(self):
        """Update the line detection status display."""
        if hasattr(self, 'line_status'):
            h_count = len(self.detected_lines.get("horizontal", []))
            v_count = len(self.detected_lines.get("vertical", []))
            expected_cells = (h_count + 1) * (v_count + 1)
            self.line_status.set(f"Lines: {h_count}H + {v_count}V = {expected_cells} cells")
    
    def _detect_table_lines(self, pdf_coords):
        """Detect table lines using visual line detection (OpenCV)."""
        if not self.pdf_path:
            return
        
        try:
            # Clear existing line visualizations
            for item in self.line_items:
                self.canvas.delete(item)
            self.line_items = []
            
            # Get the page as an image
            doc = fitz.open(self.pdf_path)
            page = doc[0]
            
            # Render page as high-resolution image
            mat = fitz.Matrix(3, 3)  # 3x zoom for better line detection
            pix = page.get_pixmap(matrix=mat)
            
            # Convert to numpy array for OpenCV
            img_data = pix.tobytes("ppm")
            pil_image = Image.open(io.BytesIO(img_data))
            img_array = np.array(pil_image)
            

            # Convert to grayscale for line detection
            gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)

            # Crop to the box area (convert PDF coords to image coords)
            x1, y1, x2, y2 = pdf_coords
            img_x1 = int(x1 * 3)
            img_y1 = int(y1 * 3)
            img_x2 = int(x2 * 3)
            img_y2 = int(y2 * 3)

            # Crop the image to the box area (color and gray)
            cropped_gray = gray[img_y1:img_y2, img_x1:img_x2]
            cropped_color = img_array[img_y1:img_y2, img_x1:img_x2]

            if cropped_gray.size == 0 or cropped_color.size == 0:
                print("Warning: Cropped area is empty")
                doc.close()
                return

            # --- Background region detection using direct grayscale thresholding ---
            # Use Otsu's method to find the optimal threshold
            # This will separate background from foreground regions automatically
            blur = cv2.GaussianBlur(cropped_gray, (5, 5), 0)
            _, thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
            
            # Find contours of distinct regions
            region_contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            
            # Convert region contours directly into lines
            color_band_lines = []
            min_width = cropped_gray.shape[1] * 0.5  # Only consider regions that span at least 50% of width
            for contour in region_contours:
                x, y, w, h = cv2.boundingRect(contour)
                if w > min_width:  # If it's wide enough to be a row
                    # Convert region boundaries to PDF coordinates
                    pdf_y1 = (y + img_y1) / 3
                    pdf_y2 = (y + h + img_y1) / 3
                    color_band_lines.extend([
                        [x1, pdf_y1, x2, pdf_y1],  # Top edge
                        [x1, pdf_y2, x2, pdf_y2]   # Bottom edge
                    ])

            # --- Enhanced line detection ---
            # Use adaptive thresholding for better line detection
            sensitivity = self.table_sensitivity.get()
            binary = cv2.adaptiveThreshold(
                cropped_gray,
                255,
                cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                cv2.THRESH_BINARY_INV,
                11,  # Block size
                2    # C constant
            )
            
            # Create kernels scaled to image size
            horizontal_size = max(cropped_gray.shape[1] // 30, 25)  # At least 25 pixels
            vertical_size = max(cropped_gray.shape[0] // 30, 25)    # At least 25 pixels
            
            # Create structure elements for detecting lines
            horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (horizontal_size, 1))
            vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, vertical_size))
            
            # Enhance horizontal lines
            horizontal_lines_img = cv2.erode(binary, horizontal_kernel)
            horizontal_lines_img = cv2.dilate(horizontal_lines_img, horizontal_kernel)
            # Clean up small noise
            horizontal_lines_img = cv2.morphologyEx(horizontal_lines_img, cv2.MORPH_OPEN, 
                                                  cv2.getStructuringElement(cv2.MORPH_RECT, (5, 1)))
            
            # Enhance vertical lines
            vertical_lines_img = cv2.erode(binary, vertical_kernel)
            vertical_lines_img = cv2.dilate(vertical_lines_img, vertical_kernel)
            # Clean up small noise
            vertical_lines_img = cv2.morphologyEx(vertical_lines_img, cv2.MORPH_OPEN,
                                                cv2.getStructuringElement(cv2.MORPH_RECT, (1, 5)))
            
            # Find contours separately for horizontal and vertical lines
            h_contours, _ = cv2.findContours(horizontal_lines_img, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            v_contours, _ = cv2.findContours(vertical_lines_img, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

            # Find and filter horizontal lines
            h_contours, _ = cv2.findContours(horizontal_lines_img, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            horizontal_lines = []
            min_line_length = cropped_gray.shape[1] * 0.3  # At least 30% of width
            for contour in h_contours:
                x, y, w, h = cv2.boundingRect(contour)
                aspect_ratio = w / float(h) if h > 0 else 0
                # Must be long enough and have high aspect ratio
                if w > min_line_length and aspect_ratio > 10:
                    # Convert to PDF coordinates
                    pdf_x1 = (x + img_x1) / 3
                    pdf_y1 = (y + img_y1) / 3
                    pdf_x2 = (x + w + img_x1) / 3
                    pdf_y2 = (y + h + img_y1) / 3
                    horizontal_lines.append([pdf_x1, pdf_y1, pdf_x2, pdf_y2])

            # Find and filter vertical lines
            v_contours, _ = cv2.findContours(vertical_lines_img, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            vertical_lines = []
            min_line_length = cropped_gray.shape[0] * 0.3  # At least 30% of height
            for contour in v_contours:
                x, y, w, h = cv2.boundingRect(contour)
                aspect_ratio = h / float(w) if w > 0 else 0
                # Must be tall enough and have high aspect ratio
                if h > min_line_length and aspect_ratio > 10:
                    # Convert to PDF coordinates
                    pdf_x1 = (x + img_x1) / 3
                    pdf_y1 = (y + img_y1) / 3
                    pdf_x2 = (x + w + img_x1) / 3
                    pdf_y2 = (y + h + img_y1) / 3
                    vertical_lines.append([pdf_x1, pdf_y1, pdf_x2, pdf_y2])

            # Combine horizontal lines and color band boundaries, then deduplicate and sort
            all_horizontal_lines = horizontal_lines + color_band_lines
            # Remove near-duplicates (lines within 2px vertically)
            def dedup_lines(lines):
                lines = sorted(lines, key=lambda l: l[1])
                deduped = []
                for l in lines:
                    if not deduped or abs(l[1] - deduped[-1][1]) > 5:
                        deduped.append(l)
                return deduped
            all_horizontal_lines = dedup_lines(all_horizontal_lines)

            self.detected_lines = {
                "horizontal": all_horizontal_lines,
                "vertical": vertical_lines
            }

            print(f"Detected {len(all_horizontal_lines)} horizontal and {len(vertical_lines)} vertical lines (including color bands)")
            doc.close()

            # Visualize detected lines
            # Also visualize the cell grid
            self._visualize_cell_grid(pdf_coords)
            
        except Exception as e:
            messagebox.showerror("Line Detection Error", f"Failed to detect lines: {e}")
            print(f"Line detection error: {e}")
            import traceback
            traceback.print_exc()
        
    def _visualize_cell_grid(self, box_coords):
        """Visualize the cell grid boundaries."""
        cells = self._generate_table_cells(box_coords)
        
        # Draw cell boundaries in orange
        for cell in cells:
            left, top, right, bottom = cell['coordinates']
            screen_coords = self._pdf_to_screen_coords([left, top, right, bottom])
            
            # Draw cell rectangle
            cell_rect = self.canvas.create_rectangle(
                screen_coords[0], screen_coords[1], 
                screen_coords[2], screen_coords[3],
                outline='orange', width=1, fill='',
                tags="cell_grid"
            )
            self.line_items.append(cell_rect)
            
            # Add cell coordinate label
            label_x = (screen_coords[0] + screen_coords[2]) / 2
            label_y = (screen_coords[1] + screen_coords[3]) / 2
            cell_label = self.canvas.create_text(
                label_x, label_y, 
                text=f"{cell['row']},{cell['col']}", 
                fill='gray', font=('Arial', 8),
                tags="cell_grid"
            )
            self.line_items.append(cell_label)
    
    def _save_table_box(self, pdf_coords, end_x, end_y, label):
        """Save the table box with detected lines as cells."""
        # Generate table cells from detected lines
        cells = self._generate_table_cells(pdf_coords)
        
        # Get current page boxes for ID generation
        current_page_boxes = self.get_current_page_boxes()
        box_id = f"p{self.current_page_num}_box{len(current_page_boxes)+1}"
        
        # Store box data
        box_data = {
            "id": box_id,
            "label": label,
            "page": self.current_page_num,
            "coordinates": pdf_coords,
            "extraction_order": len(current_page_boxes) + 1,
            "screen_coords": [self.start_x, self.start_y, end_x, end_y],
            "box_type": "table",
            "table_cells": cells,
            "detected_lines": self.detected_lines
        }
        
        # Add to current page
        if self.current_page_num not in self.pages_data:
            self.pages_data[self.current_page_num] = {"raw_text": "", "boxes": []}
        
        self.pages_data[self.current_page_num]["boxes"].append(box_data)
        self._draw_box(box_data)
        
        # Update UI
        self.update_page_info()
        self.status_var.set(f"Added table box: {label} with {len(cells)} cells on page {self.current_page_num}")
        self._cancel_table_edit()
    
    def _generate_table_cells(self, box_coords):
        x1, y1, x2, y2 = box_coords
        h_line_positions = set()

        # Collect horizontal lines
        for line in self.detected_lines["horizontal"]:
            y_pos = (line[1] + line[3]) / 2
            h_line_positions.add(y_pos)

        h_positions = sorted(h_line_positions)
        if len(h_positions) < 2:
            return []

        # Map vertical lines to rows
        vertical_segments = {}
        tolerance = 5
        for line in self.detected_lines["vertical"]:
            x_pos = (line[0] + line[2]) / 2
            line_top = min(line[1], line[3])
            line_bottom = max(line[1], line[3])
            for i in range(len(h_positions) - 1):
                seg_top = h_positions[i]
                seg_bottom = h_positions[i + 1]
                if line_top <= seg_top + tolerance and line_bottom >= seg_bottom - tolerance:
                    vertical_segments.setdefault((seg_top, seg_bottom), set()).add(x_pos)

        cells = []
        cell_id = 0

        # Row-major iteration
        for i in range(len(h_positions) - 1):
            top = h_positions[i]
            bottom = h_positions[i + 1]

            v_positions = sorted(vertical_segments.get((top, bottom), set()))

            if len(v_positions) < 2:
                # Create a single cell spanning full width
                cells.append({
                    "cell_id": cell_id,
                    "row": i,
                    "col": 0,
                    "coordinates": [x1, top, x2, bottom]
                })
                cell_id += 1
                continue

            # Create cells between consecutive verticals
            for j in range(len(v_positions) - 1):
                left = v_positions[j]
                right = v_positions[j + 1]
                cells.append({
                    "cell_id": cell_id,
                    "row": i,
                    "col": j,
                    "coordinates": [left, top, right, bottom]
                })
                cell_id += 1

        return cells
    
    def _cancel_table_edit(self):
        """Cancel table editing mode."""
        self.table_edit_mode = False
        
        # Clear line visualizations
        for item in self.line_items:
            self.canvas.delete(item)
        self.line_items = []
        
        # Close dialog
        if hasattr(self, 'table_dialog'):
            self.table_dialog.destroy()
        
        # Clear current rectangle
        if hasattr(self, 'current_rect') and self.current_rect:
            self.canvas.delete(self.current_rect)
            self.current_rect = None
    
    def _draw_box(self, box_data):
        """Draw a box on the canvas with appropriate styling."""
        x1, y1, x2, y2 = box_data['screen_coords']
        box_type = box_data.get('box_type', 'general')
        
        # Different colors for different box types
        colors = {
            'general': ('blue', 'lightblue'),
            'table': ('green', 'lightgreen'),
            'paragraph': ('purple', 'lavender')
        }
        
        outline_color, fill_color = colors.get(box_type, colors['general'])
        
        rect_id = self.canvas.create_rectangle(
            x1, y1, x2, y2,
            outline=outline_color, width=2, fill=fill_color, stipple='gray12'
        )
        
        # Add label with box type and page indicator
        label_x = (x1 + x2) / 2
        label_y = y1 - 10
        display_text = f"{box_data['label']} ({box_type}) p{box_data['page']}"
        self.canvas.create_text(label_x, label_y, text=display_text, 
                              fill=outline_color, font=('Arial', 9, 'bold'))
    
    def delete_box(self, event):
        """Delete a box on right-click."""
        click_x = self.canvas.canvasx(event.x)
        click_y = self.canvas.canvasy(event.y)
        
        # Find box to delete on current page
        current_page_boxes = self.get_current_page_boxes()
        for i, box in enumerate(current_page_boxes):
            x1, y1, x2, y2 = box['screen_coords']
            if x1 <= click_x <= x2 and y1 <= click_y <= y2:
                # Remove box from current page
                del self.pages_data[self.current_page_num]["boxes"][i]
                # Redraw current page
                self.display_page()
                # Update UI
                self.update_page_info()
                self.status_var.set(f"Deleted box: {box['label']} from page {self.current_page_num}")
                return
    
    def screen_to_pdf_coords(self, x1, y1, x2, y2):
        """Convert screen coordinates to PDF coordinates."""
        # Account for scaling
        pdf_x1 = min(x1, x2) / self.scale_factor / 2  # Divide by 2 because of 2x zoom
        pdf_y1 = min(y1, y2) / self.scale_factor / 2
        pdf_x2 = max(x1, x2) / self.scale_factor / 2
        pdf_y2 = max(y1, y2) / self.scale_factor / 2
        
        return [pdf_x1, pdf_y1, pdf_x2, pdf_y2]
    
    def _pdf_to_screen_coords(self, pdf_coords):
        """Convert PDF coordinates to screen coordinates."""
        x1, y1, x2, y2 = pdf_coords
        # Account for scaling (reverse of screen_to_pdf_coords)
        screen_x1 = x1 * self.scale_factor * 2  # Multiply by 2 because of 2x zoom
        screen_y1 = y1 * self.scale_factor * 2
        screen_x2 = x2 * self.scale_factor * 2
        screen_y2 = y2 * self.scale_factor * 2
        
        return [screen_x1, screen_y1, screen_x2, screen_y2]
    
    def clear_boxes(self):
        """Clear boxes from current page or all pages."""
        # Ask user what to clear
        result = messagebox.askyesnocancel(
            "Clear Boxes",
            f"What would you like to clear?\n\n"
            f"• Yes: Clear boxes from current page ({self.current_page_num})\n"
            f"• No: Clear boxes from ALL pages\n"
            f"• Cancel: Don't clear anything"
        )
        
        if result is None:  # Cancel
            return
        elif result is True:  # Clear current page only
            if self.current_page_num in self.pages_data:
                self.pages_data[self.current_page_num]["boxes"] = []
            self.display_page()
            self.update_page_info()
            self.status_var.set(f"Cleared boxes from page {self.current_page_num}")
        else:  # Clear all pages
            for page_num in self.pages_data:
                self.pages_data[page_num]["boxes"] = []
            self.display_page()
            self.update_page_info()
            self.status_var.set("Cleared boxes from all pages")
    
    def get_total_box_count(self):
        """Get total number of boxes across all pages."""
        total = 0
        for page_data in self.pages_data.values():
            total += len(page_data.get("boxes", []))
        return total
    
    def save_template(self):
        """Save the multi-page template to JSON file."""
        if not self.raw_text:
            messagebox.showwarning("No PDF Loaded", "Please load a PDF first!")
            return
        
        # Get template name
        template_name = simpledialog.askstring("Template Name", 
                                             "Enter template name:",
                                             initialvalue="my_multipage_bol_template")
        
        if not template_name:
            return
        
        # Count total boxes
        total_boxes = self.get_total_box_count()
        
        # Check if no boxes are defined and confirm with user
        if total_boxes == 0:
            result = messagebox.askyesnocancel(
                "No Boxes Template", 
                f"No extraction boxes are defined across all {self.total_pages} pages.\n\n"
                f"This will create a 'simple reading-order' template that extracts all text "
                f"based on coordinate positioning without specific zones.\n\n"
                f"Do you want to save this no-box multi-page template?\n"
                f"• Yes: Save template for simple reading-order extraction\n"
                f"• No: Go back to add boxes\n"
                f"• Cancel: Don't save"
            )
            
            if result is None:  # Cancel
                return
            elif result is False:  # No - go back to add boxes
                messagebox.showinfo("Add Boxes", 
                                  "Navigate through pages and draw rectangles to define extraction zones, "
                                  "then save the template.")
                return
        
        # Create multi-page template data
        template_data = {
            "template_name": template_name,
            "template_raw_text": self.raw_text,
            "template_type": "no_boxes" if total_boxes == 0 else "boxed",
            "include_unboxed_content": self.include_unboxed_content.get(),  # Store the checkbox value
            "total_pages": self.total_pages,
            "pages": {}
        }
        
        # Add page data
        for page_num in range(1, self.total_pages + 1):
            page_data = self.pages_data.get(page_num, {"raw_text": "", "boxes": []})
            template_data["pages"][str(page_num)] = {
                "page_raw_text": page_data.get("raw_text", ""),
                "boxes": page_data.get("boxes", [])
            }
        
        # Save to JSON file
        filename = f"{template_name}.json"
        filepath = os.path.join(self.templates_dir, filename)
        abs_filepath = os.path.abspath(filepath)
        
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(template_data, f, indent=2, ensure_ascii=False)
            
            unboxed_status = "with unboxed content" if self.include_unboxed_content.get() else "boxes only"
            if total_boxes > 0:
                message = (f"Multi-page template saved as:\n{abs_filepath}\n\n"
                        f"Pages: {self.total_pages}\n"
                        f"Total extraction boxes: {total_boxes} ({unboxed_status})")
            else:
                message = (f"Multi-page template saved as:\n{abs_filepath}\n\n"
                        f"Type: Simple reading-order extraction (no boxes)\n"
                        f"Pages: {self.total_pages}")
            
            messagebox.showinfo("Template Saved", message)
            
            template_type = "no-box" if total_boxes == 0 else "boxed"
            self.status_var.set(f"Template saved: {filename} ({template_type}, {self.total_pages} pages, {total_boxes} total boxes)")
            
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save template: {e}")
    
    def load_template(self):
        """Load an existing multi-page template."""
        # Make sure templates directory exists
        if not os.path.exists(self.templates_dir):
            messagebox.showwarning("Templates Directory", f"Templates directory '{self.templates_dir}' not found")
            return
            
        # Get absolute path to templates directory
        abs_templates_dir = os.path.abspath(self.templates_dir)
        
        file_path = filedialog.askopenfilename(
            title="Select Template File",
            initialdir=abs_templates_dir,
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                template_data = json.load(f)
            
            # Check if we need a PDF loaded first
            if not self.current_page:
                messagebox.showwarning("No PDF", "Please load a PDF first to display the template boxes")
                return

            # Restore unboxed content setting if it exists
            if 'include_unboxed_content' in template_data:
                self.include_unboxed_content.set(template_data['include_unboxed_content'])

            # Load template pages
            template_pages = template_data.get('pages', {})
            template_type = template_data.get('template_type', 'boxed')
            total_boxes = 0
            
            # Clear existing page data
            self.pages_data = {}
            
            # Load each template page
            for page_str, page_data in template_pages.items():
                page_num = int(page_str)
                boxes = page_data.get('boxes', [])
                total_boxes += len(boxes)
                
                # Convert PDF coordinates back to screen coordinates for display
                for box in boxes:
                    pdf_coords = box['coordinates']
                    x1 = pdf_coords[0] * 2 * self.scale_factor
                    y1 = pdf_coords[1] * 2 * self.scale_factor
                    x2 = pdf_coords[2] * 2 * self.scale_factor
                    y2 = pdf_coords[3] * 2 * self.scale_factor
                    box['screen_coords'] = [x1, y1, x2, y2]
                
                # Store page data
                self.pages_data[page_num] = {
                    "raw_text": page_data.get('page_raw_text', ''),
                    "boxes": boxes
                }
            
            # Update display
            self.display_page()
            self.update_page_info()
            
            # Show appropriate message
            template_name = template_data.get('template_name', 'Unknown')
            template_pages_count = len(template_pages)
            
            if total_boxes > 0:
                message = f"Loaded template: {template_name} ({template_pages_count} pages, {total_boxes} total boxes)"
            else:
                message = f"Loaded template: {template_name} ({template_pages_count} pages, simple reading-order, no boxes)"
            
            self.status_var.set(message)

        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load template: {e}")

    def run(self):
        """Run the multi-page template creator."""
        self.root.mainloop()

if __name__ == "__main__":
    creator = TemplateCreator()
    creator.run()