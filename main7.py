import os, re, sys
import tkinter as tk
from tkinter import filedialog, ttk, messagebox, colorchooser, font, simpledialog
from PIL import Image, ImageDraw, ImageFont, ImageTk
import pandas as pd
from pathlib import Path

def sanitize_filename(s):
    """Clean filename for safe saving"""
    s = str(s)
    s = re.sub(r'[\\/:"*?<>|]+', '', s)
    s = re.sub(r'\s+', '_', s.strip())
    return s if s else "unnamed"

def pick_file(title, filetypes):
    """File picker dialog with fallback options"""
    try:
        # Try standard tkinter file dialog
        root = tk.Tk()
        root.withdraw()
        root.lift()  # Bring to front
        root.attributes('-topmost', True)  # Keep on top
        
        path = filedialog.askopenfilename(
            title=title, 
            filetypes=filetypes,
            parent=root
        )
        root.destroy()
        
        if path:
            return path
        else:
            print("No file selected")
            return None
            
    except Exception as e:
        print(f"File dialog error: {e}")
        
        # Fallback: Manual file path entry
        root = tk.Tk()
        root.title("File Selection")
        root.geometry("500x200")
        root.configure(bg='#f0f0f0')
        
        # Center window
        root.update_idletasks()
        x = (root.winfo_screenwidth() // 2) - (250)
        y = (root.winfo_screenheight() // 2) - (100)
        root.geometry(f"+{x}+{y}")
        
        tk.Label(root, text=f"{title}", font=("Arial", 12, "bold"), bg='#f0f0f0').pack(pady=10)
        tk.Label(root, text="File dialog failed. Please enter the file path manually:", 
                bg='#f0f0f0').pack(pady=5)
        
        path_var = tk.StringVar()
        entry_frame = tk.Frame(root, bg='#f0f0f0')
        entry_frame.pack(pady=10, padx=20, fill='x')
        
        tk.Entry(entry_frame, textvariable=path_var, width=50).pack(side='left', fill='x', expand=True)
        
        def browse_manually():
            try:
                import subprocess
                import platform
                
                if platform.system() == "Windows":
                    subprocess.Popen('explorer')
                elif platform.system() == "Darwin":  # macOS
                    subprocess.Popen(['open', '.'])
                else:  # Linux
                    subprocess.Popen(['xdg-open', '.'])
            except:
                pass
        
        tk.Button(entry_frame, text="Browse", command=browse_manually).pack(side='right', padx=(5,0))
        
        result = {'path': None}
        
        def confirm():
            path = path_var.get().strip()
            if path and os.path.exists(path):
                result['path'] = path
                root.quit()
            else:
                messagebox.showerror("Error", "Please enter a valid file path")
        
        def cancel():
            root.quit()
        
        button_frame = tk.Frame(root, bg='#f0f0f0')
        button_frame.pack(pady=20)
        
        tk.Button(button_frame, text="OK", command=confirm, bg='#4CAF50', fg='white', 
                 font=("Arial", 10, "bold")).pack(side='left', padx=5)
        tk.Button(button_frame, text="Cancel", command=cancel, bg='#f44336', fg='white',
                 font=("Arial", 10, "bold")).pack(side='left', padx=5)
        
        # Add example text
        example_frame = tk.Frame(root, bg='#f0f0f0')
        example_frame.pack(pady=5)
        tk.Label(example_frame, text="Example: C:\\Users\\YourName\\Documents\\template.png", 
                font=("Arial", 9), fg='gray', bg='#f0f0f0').pack()
        
        root.mainloop()
        root.destroy()
        
        return result['path']

def get_system_fonts():
    """Get all available system fonts"""
    fonts = set()
    
    # Common font directories
    font_dirs = []
    
    if sys.platform == "win32":
        font_dirs = [
            "C:/Windows/Fonts/",
            os.path.expanduser("~/AppData/Local/Microsoft/Windows/Fonts/")
        ]
    elif sys.platform == "darwin":  # macOS
        font_dirs = [
            "/System/Library/Fonts/",
            "/Library/Fonts/",
            os.path.expanduser("~/Library/Fonts/")
        ]
    else:  # Linux
        font_dirs = [
            "/usr/share/fonts/",
            "/usr/local/share/fonts/",
            os.path.expanduser("~/.fonts/"),
            os.path.expanduser("~/.local/share/fonts/")
        ]
    
    # Scan font directories
    for font_dir in font_dirs:
        if os.path.exists(font_dir):
            for root, dirs, files in os.walk(font_dir):
                for file in files:
                    if file.lower().endswith(('.ttf', '.otf', '.ttc')):
                        fonts.add(os.path.join(root, file))
    
    # Add common font names for fallback
    common_fonts = [
        "Arial", "Arial Bold", "Arial Italic", "Arial Black",
        "Times New Roman", "Times New Roman Bold", "Times New Roman Italic",
        "Calibri", "Calibri Bold", "Calibri Italic",
        "Helvetica", "Helvetica Bold", "Helvetica Italic",
        "Georgia", "Georgia Bold", "Georgia Italic",
        "Verdana", "Verdana Bold", "Verdana Italic",
        "Trebuchet MS", "Trebuchet MS Bold", "Trebuchet MS Italic",
        "Comic Sans MS", "Comic Sans MS Bold",
        "Impact", "Tahoma", "Tahoma Bold",
        "Courier New", "Courier New Bold", "Courier New Italic",
        "Lucida Console", "Consolas", "Monaco",
        "Palatino", "Palatino Bold", "Palatino Italic",
        "Bookman", "Bookman Bold", "Bookman Italic",
        "Century Gothic", "Century Gothic Bold",
        "Franklin Gothic", "Franklin Gothic Bold",
        "Garamond", "Garamond Bold", "Garamond Italic"
    ]
    
    return sorted(list(fonts)) + common_fonts

class ProfessionalCertificateEditor:
    def __init__(self, template, data, text_areas):
        self.template = template
        self.data = data  # DataFrame containing all data
        self.text_areas = text_areas  # List of dictionaries with 'rect' and 'column'
        self.index = 0
        self.font_size = 50
        self.alignment = "center"
        self.text_color = (0, 0, 0)
        
        # Set initial text position to the center of the first text area if available
        if text_areas:
            rect = text_areas[0]['rect']
            self.text_x, self.text_y = (rect[0] + rect[2]) // 2, (rect[1] + rect[3]) // 2
        else:
            self.text_x, self.text_y = template.width // 2, template.height // 2
            
        self.current_font_path = None
        self.font_family = "Times New Roman"
        self.font_style = "Regular"
        
        # Font management
        self.available_fonts = get_system_fonts()
        self.font_families = self._organize_fonts()
        
        # Load default font
        self._load_font()
        
        # Text editing variables
        self.selected_text_area = None
        self.text_positions = {}  # Store positions for each text area
        
        # Create main window
        self.root = tk.Tk()
        self.root.title("Professional Certificate Editor v2.0")
        self.root.state('zoomed') if sys.platform == "win32" else self.root.attributes('-zoomed', True)
        self.root.configure(bg='#f0f0f0')
        
        # Configure styles
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        self.setup_ui()
        self.update_preview()
        self.root.mainloop()
        
    def _organize_fonts(self):
        """Organize fonts by family"""
        families = {}
        
        # Add common font families
        common_families = {
            "Times New Roman": ["Regular", "Bold", "Italic", "Bold Italic"],
            "Arial": ["Regular", "Bold", "Italic", "Bold Italic", "Black"],
            "Calibri": ["Regular", "Bold", "Italic", "Bold Italic"],
            "Georgia": ["Regular", "Bold", "Italic", "Bold Italic"],
            "Verdana": ["Regular", "Bold", "Italic", "Bold Italic"],
            "Helvetica": ["Regular", "Bold", "Italic", "Bold Italic"],
            "Trebuchet MS": ["Regular", "Bold", "Italic", "Bold Italic"],
            "Comic Sans MS": ["Regular", "Bold"],
            "Impact": ["Regular"],
            "Tahoma": ["Regular", "Bold"],
            "Courier New": ["Regular", "Bold", "Italic", "Bold Italic"],
            "Palatino": ["Regular", "Bold", "Italic", "Bold Italic"],
            "Garamond": ["Regular", "Bold", "Italic", "Bold Italic"],
            "Century Gothic": ["Regular", "Bold"],
            "Franklin Gothic": ["Regular", "Bold"],
            "Bookman": ["Regular", "Bold", "Italic", "Bold Italic"]
        }
        
        for family, styles in common_families.items():
            families[family] = styles
            
        return families
        
    def _load_font(self):
        """Load the current font"""
        try:
            # Try to load system font by name first
            font_name = self.font_family
            if self.font_style != "Regular":
                font_name += f" {self.font_style}"
            
            self.font = ImageFont.truetype(font_name, self.font_size)
        except:
            # Try common font files
            font_files = {
                ("Times New Roman", "Regular"): ["times.ttf", "TimesNewRomanPSMT.ttf"],
                ("Times New Roman", "Bold"): ["timesbd.ttf", "TimesNewRomanPS-BoldMT.ttf"],
                ("Times New Roman", "Italic"): ["timesi.ttf", "TimesNewRomanPS-ItalicMT.ttf"],
                ("Times New Roman", "Bold Italic"): ["timesbi.ttf", "TimesNewRomanPS-BoldItalicMT.ttf"],
                ("Arial", "Regular"): ["arial.ttf", "ArialMT.ttf"],
                ("Arial", "Bold"): ["arialbd.ttf", "Arial-BoldMT.ttf"],
                ("Arial", "Italic"): ["ariali.ttf", "Arial-ItalicMT.ttf"],
                ("Arial", "Bold Italic"): ["arialbi.ttf", "Arial-BoldItalicMT.ttf"],
                ("Calibri", "Regular"): ["calibri.ttf"],
                ("Calibri", "Bold"): ["calibrib.ttf"],
                ("Calibri", "Italic"): ["calibrii.ttf"],
                ("Calibri", "Bold Italic"): ["calibriz.ttf"],
            }
            
            key = (self.font_family, self.font_style)
            if key in font_files:
                for font_file in font_files[key]:
                    try:
                        self.font = ImageFont.truetype(font_file, self.font_size)
                        break
                    except:
                        continue
                else:
                    # If no specific font found, use default
                    self.font = ImageFont.load_default()
            else:
                self.font = ImageFont.load_default()
                
    def setup_ui(self):
        """Setup the user interface"""
        # Main container
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Left panel for controls
        control_panel = ttk.Frame(main_frame, width=350)
        control_panel.pack(side="left", fill="y", padx=(0, 10))
        control_panel.pack_propagate(False)
        
        # Right panel for canvas
        canvas_frame = ttk.Frame(main_frame)
        canvas_frame.pack(side="right", fill="both", expand=True)
        
        # Setup control panels
        self.setup_control_panel(control_panel)
        self.setup_canvas(canvas_frame)
        
    def setup_control_panel(self, parent):
        """Setup the control panel"""
        # Title
        title_frame = ttk.Frame(parent)
        title_frame.pack(fill="x", pady=(0, 20))
        ttk.Label(title_frame, text="Professional Certificate Editor", 
                 font=("Arial", 14, "bold")).pack()
        
        # Save All button at the top
        save_all_frame = ttk.Frame(parent)
        save_all_frame.pack(fill="x", pady=(0, 10))
        ttk.Button(save_all_frame, text="💾 Save All Certificates", 
                  command=self.generate_all_pdf).pack(fill="x", pady=5)
        
        # Name navigation
        nav_frame = ttk.LabelFrame(parent, text="Record Navigation", padding=10)
        nav_frame.pack(fill="x", pady=(0, 10))
        
        name_info_frame = ttk.Frame(nav_frame)
        name_info_frame.pack(fill="x", pady=(0, 10))
        
        self.name_label = ttk.Label(name_info_frame, text="", font=("Arial", 12, "bold"))
        self.name_label.pack()
        
        self.index_label = ttk.Label(name_info_frame, text="")
        self.index_label.pack()
        
        nav_buttons_frame = ttk.Frame(nav_frame)
        nav_buttons_frame.pack(fill="x")
        
        ttk.Button(nav_buttons_frame, text="◀ Previous", 
                  command=self.prev_name).pack(side="left", padx=(0, 5))
        ttk.Button(nav_buttons_frame, text="Next ▶", 
                  command=self.next_name).pack(side="right", padx=(5, 0))
        
        # Font selection
        font_frame = ttk.LabelFrame(parent, text="Font Settings", padding=10)
        font_frame.pack(fill="x", pady=(0, 10))
        
        # Font family
        ttk.Label(font_frame, text="Font Family:").pack(anchor="w")
        self.font_family_var = tk.StringVar(value=self.font_family)
        font_family_combo = ttk.Combobox(font_frame, textvariable=self.font_family_var,
                                       values=list(self.font_families.keys()), width=25)
        font_family_combo.pack(fill="x", pady=(0, 10))
        font_family_combo.bind("<<ComboboxSelected>>", self.on_font_family_change)
        
        # Font style
        ttk.Label(font_frame, text="Font Style:").pack(anchor="w")
        self.font_style_var = tk.StringVar(value=self.font_style)
        self.font_style_combo = ttk.Combobox(font_frame, textvariable=self.font_style_var,
                                           values=self.font_families.get(self.font_family, ["Regular"]), 
                                           width=25)
        self.font_style_combo.pack(fill="x", pady=(0, 10))
        self.font_style_combo.bind("<<ComboboxSelected>>", self.on_font_style_change)
        
        # Font size
        size_frame = ttk.Frame(font_frame)
        size_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(size_frame, text="Font Size:").pack(side="left")
        self.size_var = tk.IntVar(value=self.font_size)
        size_spinbox = tk.Spinbox(size_frame, from_=10, to=200, textvariable=self.size_var, 
                                width=8, command=self.on_size_change)
        size_spinbox.pack(side="left", padx=(10, 5))
        
        ttk.Button(size_frame, text="−", width=3, 
                  command=self.decrease_size).pack(side="left", padx=(0, 2))
        ttk.Button(size_frame, text="+", width=3, 
                  command=self.increase_size).pack(side="left")
        
        # Alignment
        align_frame = ttk.Frame(font_frame)
        align_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(align_frame, text="Alignment:").pack(side="left")
        self.align_var = tk.StringVar(value=self.alignment)
        align_combo = ttk.Combobox(align_frame, textvariable=self.align_var,
                                 values=["left", "center", "right"], width=10)
        align_combo.pack(side="left", padx=(10, 0))
        align_combo.bind("<<ComboboxSelected>>", self.on_alignment_change)
        
        # Color selection
        color_frame = ttk.LabelFrame(parent, text="Text Color", padding=10)
        color_frame.pack(fill="x", pady=(0, 10))
        
        # Predefined colors
        colors_grid = ttk.Frame(color_frame)
        colors_grid.pack(fill="x", pady=(0, 10))
        
        colors = [
            ("Black", (0, 0, 0)), ("Navy", (0, 0, 128)), ("Blue", (0, 0, 255)),
            ("Red", (255, 0, 0)), ("Maroon", (128, 0, 0)), ("Purple", (128, 0, 128)),
            ("Green", (0, 128, 0)), ("Olive", (128, 128, 0)), ("Orange", (255, 165, 0)),
            ("Gray", (128, 128, 128)), ("Silver", (192, 192, 192)), ("Gold", (255, 215, 0))
        ]
        
        for i, (name, rgb) in enumerate(colors):
            row = i // 3
            col = i % 3
            color_hex = f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
            btn = tk.Button(colors_grid, bg=color_hex, width=8, height=2,
                          command=lambda c=rgb: self.set_color(c))
            btn.grid(row=row, column=col, padx=2, pady=2)
        
        ttk.Button(color_frame, text="Custom Color", 
                  command=self.pick_custom_color).pack(pady=(5, 0))
        
        # Position controls
        pos_frame = ttk.LabelFrame(parent, text="Text Position", padding=10)
        pos_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(pos_frame, text="Click and drag on canvas to move text").pack()
        ttk.Label(pos_frame, text="Double-click to edit text content").pack()
        ttk.Label(pos_frame, text="Right-click and drag to resize font").pack()
        
        pos_info_frame = ttk.Frame(pos_frame)
        pos_info_frame.pack(fill="x", pady=(10, 0))
        
        self.pos_label = ttk.Label(pos_info_frame, text=f"Position: ({self.text_x}, {self.text_y})")
        self.pos_label.pack()
        
        # Save Options
        save_frame = ttk.LabelFrame(parent, text="Save Options", padding=10)
        save_frame.pack(fill="x", pady=(0, 10))
        
        # Current certificate options
        ttk.Label(save_frame, text="Current Certificate:", font=("Arial", 10, "bold")).pack(anchor="w", pady=(5, 0))
        ttk.Button(save_frame, text="💾 Save Current (PDF+PNG)", 
                  command=self.save_current).pack(fill="x", pady=(0, 5))
        ttk.Button(save_frame, text="📄 Save Current as PDF Only", 
                  command=self.save_current_pdf).pack(fill="x", pady=(0, 5))
        
        # All certificates options
        ttk.Label(save_frame, text="All Certificates:", font=("Arial", 10, "bold")).pack(anchor="w", pady=(10, 0))
        ttk.Button(save_frame, text="📁 Generate All (PDF+PNG)", 
                  command=self.generate_all).pack(fill="x", pady=(0, 5))
        ttk.Button(save_frame, text="📄 Generate All as PDF Only", 
                  command=self.generate_all_pdf).pack(fill="x", pady=(0, 5))
        
        # Other actions
        action_frame = ttk.LabelFrame(parent, text="Other Actions", padding=10)
        action_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Button(action_frame, text="🔄 Reset Settings", 
                  command=self.reset_settings).pack(fill="x", pady=(0, 5))
        
        # Preview settings
        preview_frame = ttk.LabelFrame(parent, text="Preview Options", padding=10)
        preview_frame.pack(fill="x")
        
        self.show_guides_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(preview_frame, text="Show text boundary", 
                       variable=self.show_guides_var, 
                       command=self.update_preview).pack(anchor="w")
        
    def setup_canvas(self, parent):
        """Setup the canvas area"""
        # Canvas with scrollbars
        canvas_container = ttk.Frame(parent)
        canvas_container.pack(fill="both", expand=True)
        
        self.canvas = tk.Canvas(canvas_container, bg="white", cursor="crosshair")
        
        # Scrollbars
        v_scroll = ttk.Scrollbar(canvas_container, orient="vertical", command=self.canvas.yview)
        h_scroll = ttk.Scrollbar(canvas_container, orient="horizontal", command=self.canvas.xview)
        
        self.canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        # Pack scrollbars and canvas
        v_scroll.pack(side="right", fill="y")
        h_scroll.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)
        
        # Bind events
        self.canvas.bind("<ButtonPress-1>", self.start_move)
        self.canvas.bind("<B1-Motion>", self.do_move)
        self.canvas.bind("<ButtonRelease-1>", self.end_move)
        self.canvas.bind("<Double-Button-1>", self.edit_text)
        self.canvas.bind("<ButtonPress-3>", self.start_resize)
        self.canvas.bind("<B3-Motion>", self.do_resize)
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)
        
        self.drag_data = {"x": 0, "y": 0, "item": None}
        self.resize_data = {"y": 0}
        
    def on_font_family_change(self, event=None):
        """Handle font family change"""
        self.font_family = self.font_family_var.get()
        # Update style options
        styles = self.font_families.get(self.font_family, ["Regular"])
        self.font_style_combo['values'] = styles
        self.font_style = styles[0]
        self.font_style_var.set(self.font_style)
        self._load_font()
        self.update_preview()
        
    def on_font_style_change(self, event=None):
        """Handle font style change"""
        self.font_style = self.font_style_var.get()
        self._load_font()
        self.update_preview()
        
    def on_size_change(self):
        """Handle font size change from spinbox"""
        self.font_size = self.size_var.get()
        self._load_font()
        self.update_preview()
        
    def on_alignment_change(self, event=None):
        """Handle alignment change"""
        self.alignment = self.align_var.get()
        self.update_preview()
        
    def on_mousewheel(self, event):
        """Handle mouse wheel for canvas scrolling"""
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
    def set_color(self, rgb):
        """Set text color"""
        self.text_color = rgb
        self.update_preview()
        
    def pick_custom_color(self):
        """Pick custom color"""
        color = colorchooser.askcolor(title="Choose Text Color")[0]
        if color:
            self.text_color = tuple(int(c) for c in color)
            self.update_preview()
            
    def get_current_name(self):
        """Get current name being edited (from the first column)"""
        if self.data.shape[1] > 0:
            return str(self.data.iloc[self.index, 0])
        return "No data"
        
    def render_image(self, show_guides=True):
        """Render certificate with all text areas for the current row"""
        img = self.template.copy()
        draw = ImageDraw.Draw(img)
        
        # Get the current row from the DataFrame
        current_row = self.data.iloc[self.index]
        
        # For each text area
        for i, area in enumerate(self.text_areas):
            rect = area['rect']
            column = area['column']
            
            # Get the text for this area from the current row
            # If we have custom text for this area, use it
            if i in self.text_positions and 'text' in self.text_positions[i]:
                text = self.text_positions[i]['text']
            else:
                text = str(current_row[column])
            
            # Get text dimensions
            bbox = draw.textbbox((0, 0), text, font=self.font)
            tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
            
            # Calculate position based on alignment (using the current alignment setting for all)
            if self.alignment == "center":
                tx = (rect[0] + rect[2]) // 2 - tw // 2
            elif self.alignment == "left":
                tx = rect[0]
            else:  # right
                tx = rect[2] - tw
            
            ty = (rect[1] + rect[3]) // 2 - th // 2
            
            # Store position for this text area if not already stored
            if i not in self.text_positions:
                self.text_positions[i] = {
                    'x': (rect[0] + rect[2]) // 2,
                    'y': (rect[1] + rect[3]) // 2,
                    'text': text
                }
            
            # Update stored position if it exists
            self.text_positions[i]['x'] = (rect[0] + rect[2]) // 2
            self.text_positions[i]['y'] = (rect[1] + rect[3]) // 2
            
            # Draw text
            draw.text((tx, ty), text, font=self.font, fill=self.text_color)
            
            # Draw guides if enabled
            if show_guides and self.show_guides_var.get():
                # Text boundary
                draw.rectangle([tx, ty, tx + tw, ty + th], outline="red", width=2)
                # Rectangle area
                draw.rectangle([rect[0], rect[1], rect[2], rect[3]], outline="blue", width=1)
        
        return img
        
    def update_preview(self):
        """Update canvas preview"""
        img = self.render_image(show_guides=True)
        
        # Calculate display size
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        
        if canvas_width > 1 and canvas_height > 1:  # Canvas is initialized
            scale = min(canvas_width / img.width, canvas_height / img.height, 1.0)
            disp_w = int(img.width * scale)
            disp_h = int(img.height * scale)
            
            if disp_w > 0 and disp_h > 0:
                disp_img = img.resize((disp_w, disp_h), Image.LANCZOS)
                self.tk_img = ImageTk.PhotoImage(disp_img)
                
                self.canvas.delete("all")
                self.canvas.create_image(disp_w//2, disp_h//2, image=self.tk_img)
                
                # Update scroll region
                self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
        # Update labels
        self.name_label.config(text=self.get_current_name())
        self.index_label.config(text=f"Record {self.index + 1} of {len(self.data)}")
        self.pos_label.config(text=f"Position: ({self.text_x}, {self.text_y})")
        
    def next_name(self):
        """Navigate to next name"""
        self.index = (self.index + 1) % len(self.data)
        self.text_positions = {}  # Reset custom text positions when changing records
        self.update_preview()
        
    def prev_name(self):
        """Navigate to previous name"""
        self.index = (self.index - 1) % len(self.data)
        self.text_positions = {}  # Reset custom text positions when changing records
        self.update_preview()
        
    def increase_size(self):
        """Increase font size"""
        self.font_size += 2
        self.size_var.set(self.font_size)
        self._load_font()
        self.update_preview()
        
    def decrease_size(self):
        """Decrease font size"""
        self.font_size = max(10, self.font_size - 2)
        self.size_var.set(self.font_size)
        self._load_font()
        self.update_preview()
        
    def start_move(self, event):
        """Start text movement"""
        # Calculate movement in original image coordinates
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        
        if canvas_width > 1 and canvas_height > 1:
            scale = min(canvas_width / self.template.width, 
                       canvas_height / self.template.height, 1.0)
            
            # Convert click coordinates to original image coordinates
            img_x = int(event.x / scale)
            img_y = int(event.y / scale)
            
            # Find which text area was clicked
            for i, area in enumerate(self.text_areas):
                rect = area['rect']
                if (rect[0] <= img_x <= rect[2] and rect[1] <= img_y <= rect[3]):
                    self.selected_text_area = i
                    self.canvas.configure(cursor="fleur")
                    self.drag_data["x"] = event.x
                    self.drag_data["y"] = event.y
                    return
            
            # If no text area was clicked, use the default behavior
            self.selected_text_area = None
            self.canvas.configure(cursor="fleur")
            self.drag_data["x"] = event.x
            self.drag_data["y"] = event.y
            
    def do_move(self, event):
        """Handle text movement"""
        # Calculate movement in original image coordinates
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        
        if canvas_width > 1 and canvas_height > 1:
            scale = min(canvas_width / self.template.width, 
                       canvas_height / self.template.height, 1.0)
            
            dx = (event.x - self.drag_data["x"]) / scale
            dy = (event.y - self.drag_data["y"]) / scale
            
            if self.selected_text_area is not None:
                # Move the selected text area
                if self.selected_text_area in self.text_positions:
                    self.text_positions[self.selected_text_area]['x'] += dx
                    self.text_positions[self.selected_text_area]['y'] += dy
                    
                    # Update the rectangle position as well
                    rect = self.text_areas[self.selected_text_area]['rect']
                    width = rect[2] - rect[0]
                    height = rect[3] - rect[1]
                    center_x = self.text_positions[self.selected_text_area]['x']
                    center_y = self.text_positions[self.selected_text_area]['y']
                    
                    self.text_areas[self.selected_text_area]['rect'] = (
                        center_x - width // 2,
                        center_y - height // 2,
                        center_x + width // 2,
                        center_y + height // 2
                    )
            else:
                # Move all text areas (default behavior)
                for i in self.text_positions:
                    self.text_positions[i]['x'] += dx
                    self.text_positions[i]['y'] += dy
                    
                    # Update the rectangle position as well
                    rect = self.text_areas[i]['rect']
                    width = rect[2] - rect[0]
                    height = rect[3] - rect[1]
                    center_x = self.text_positions[i]['x']
                    center_y = self.text_positions[i]['y']
                    
                    self.text_areas[i]['rect'] = (
                        center_x - width // 2,
                        center_y - height // 2,
                        center_x + width // 2,
                        center_y + height // 2
                    )
            
            self.drag_data["x"] = event.x
            self.drag_data["y"] = event.y
            
            self.update_preview()
            
    def end_move(self, event):
        """End text movement"""
        self.selected_text_area = None
        self.canvas.configure(cursor="crosshair")
        
    def edit_text(self, event):
        """Edit text content on double-click"""
        # Calculate movement in original image coordinates
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        
        if canvas_width > 1 and canvas_height > 1:
            scale = min(canvas_width / self.template.width, 
                       canvas_height / self.template.height, 1.0)
            
            # Convert click coordinates to original image coordinates
            img_x = int(event.x / scale)
            img_y = int(event.y / scale)
            
            # Find which text area was clicked
            for i, area in enumerate(self.text_areas):
                rect = area['rect']
                if (rect[0] <= img_x <= rect[2] and rect[1] <= img_y <= rect[3]):
                    # Get current text
                    current_row = self.data.iloc[self.index]
                    column = area['column']
                    
                    # If we have custom text for this area, use it
                    if i in self.text_positions and 'text' in self.text_positions[i]:
                        current_text = self.text_positions[i]['text']
                    else:
                        current_text = str(current_row[column])
                    
                    # Create dialog to edit text
                    dialog = tk.Toplevel(self.root)
                    dialog.title("Edit Text")
                    dialog.geometry("400x150")
                    dialog.transient(self.root)
                    dialog.grab_set()
                    
                    # Center dialog
                    dialog.update_idletasks()
                    x = (dialog.winfo_screenwidth() // 2) - (200)
                    y = (dialog.winfo_screenheight() // 2) - (75)
                    dialog.geometry(f"+{x}+{y}")
                    
                    ttk.Label(dialog, text=f"Edit text for {column}:").pack(pady=10)
                    
                    text_var = tk.StringVar(value=current_text)
                    text_entry = ttk.Entry(dialog, textvariable=text_var, width=40)
                    text_entry.pack(pady=10)
                    text_entry.select_range(0, tk.END)
                    text_entry.focus_set()
                    
                    def save_text():
                        new_text = text_var.get()
                        if i not in self.text_positions:
                            self.text_positions[i] = {}
                        self.text_positions[i]['text'] = new_text
                        dialog.destroy()
                        self.update_preview()
                    
                    button_frame = ttk.Frame(dialog)
                    button_frame.pack(pady=10)
                    
                    ttk.Button(button_frame, text="Save", command=save_text).pack(side="left", padx=5)
                    ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side="left", padx=5)
                    
                    # Bind Enter key to save
                    dialog.bind("<Return>", lambda e: save_text())
                    
                    return
            
    def start_resize(self, event):
        """Start font resizing"""
        self.canvas.configure(cursor="sizing")
        self.resize_data["y"] = event.y
        
    def do_resize(self, event):
        """Handle font resizing"""
        dy = event.y - self.resize_data["y"]
        self.font_size = max(10, self.font_size + dy // 3)
        self.size_var.set(self.font_size)
        self.resize_data["y"] = event.y
        self._load_font()
        self.update_preview()
        
    def save_current(self):
        """Save current certificate as both PDF and PNG"""
        try:
            output_dir = "Certificates"
            os.makedirs(output_dir, exist_ok=True)
            
            # Use the first column value for filename
            name = self.get_current_name()
            img = self.render_image(show_guides=False)
            safe_name = sanitize_filename(name)
            
            # Save as high-quality PDF
            pdf_path = os.path.join(output_dir, f"{safe_name}.pdf")
            img.save(pdf_path, "PDF", resolution=300.0, quality=100)
            
            # Also save as PNG
            png_path = os.path.join(output_dir, f"{safe_name}.png")
            img.save(png_path, "PNG", dpi=(300, 300))
            
            messagebox.showinfo("Certificate Saved", 
                              f"Certificate saved successfully!\n\nPDF: {pdf_path}\nPNG: {png_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save certificate:\n{str(e)}")
            
    def save_current_pdf(self):
        """Save current certificate as PDF only"""
        try:
            output_dir = "Certificates"
            os.makedirs(output_dir, exist_ok=True)
            
            # Use the first column value for filename
            name = self.get_current_name()
            img = self.render_image(show_guides=False)
            safe_name = sanitize_filename(name)
            
            # Save as high-quality PDF only
            pdf_path = os.path.join(output_dir, f"{safe_name}.pdf")
            img.save(pdf_path, "PDF", resolution=300.0, quality=100)
            
            messagebox.showinfo("Certificate Saved", 
                              f"Certificate saved successfully as PDF!\n\nPDF: {pdf_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save certificate:\n{str(e)}")
        
    def reset_settings(self):
        """Reset all settings to defaults"""
        self.font_size = 50
        if self.text_areas:
            rect = self.text_areas[0]['rect']
            self.text_x, self.text_y = (rect[0] + rect[2]) // 2, (rect[1] + rect[3]) // 2
        else:
            self.text_x, self.text_y = self.template.width // 2, template.height // 2
            
        self.alignment = "center"
        self.text_color = (0, 0, 0)
        self.font_family = "Times New Roman"
        self.font_style = "Regular"
        
        # Reset custom text positions
        self.text_positions = {}
        
        # Update UI
        self.align_var.set(self.alignment)
        self.size_var.set(self.font_size)
        self.font_family_var.set(self.font_family)
        self.font_style_var.set(self.font_style)
        
        self._load_font()
        self.update_preview()
        
        messagebox.showinfo("Settings Reset", "All settings have been reset to defaults.")
        
    def generate_all(self):
        """Generate all certificates (PDF and PNG)"""
        try:
            output_dir = "Certificates"
            os.makedirs(output_dir, exist_ok=True)
            
            progress_window = tk.Toplevel(self.root)
            progress_window.title("Generating Certificates")
            progress_window.geometry("400x150")
            progress_window.transient(self.root)
            progress_window.grab_set()
            
            ttk.Label(progress_window, text="Generating certificates (PDF+PNG)...").pack(pady=20)
            
            progress_var = tk.DoubleVar()
            progress_bar = ttk.Progressbar(progress_window, variable=progress_var, 
                                         maximum=len(self.data))
            progress_bar.pack(padx=20, fill="x")
            
            status_label = ttk.Label(progress_window, text="")
            status_label.pack(pady=10)
            
            # Save current text positions to restore later
            saved_text_positions = self.text_positions.copy()
            
            for i in range(len(self.data)):
                self.index = i
                status_label.config(text=f"Processing: {self.get_current_name()}")
                progress_window.update()
                
                # Reset text positions for each certificate
                self.text_positions = {}
                
                img = self.render_image(show_guides=False)
                safe_name = sanitize_filename(self.get_current_name())
                
                # Save PDF
                pdf_path = os.path.join(output_dir, f"{safe_name}.pdf")
                img.save(pdf_path, "PDF", resolution=300.0, quality=100)
                
                # Save PNG
                png_path = os.path.join(output_dir, f"{safe_name}.png")
                img.save(png_path, "PNG", dpi=(300, 300))
                
                progress_var.set(i + 1)
            
            # Restore text positions
            self.text_positions = saved_text_positions
            self.index = 0
            
            progress_window.destroy()
            
            messagebox.showinfo("Generation Complete", 
                              f"All {len(self.data)} certificates have been generated!\n\n"
                              f"Files saved in: {os.path.abspath(output_dir)}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate certificates:\n{str(e)}")
    
    def generate_all_pdf(self):
        """Generate all certificates as PDF only"""
        try:
            output_dir = "Certificates"
            os.makedirs(output_dir, exist_ok=True)
            
            progress_window = tk.Toplevel(self.root)
            progress_window.title("Generating PDF Certificates")
            progress_window.geometry("400x150")
            progress_window.transient(self.root)
            progress_window.grab_set()
            
            ttk.Label(progress_window, text="Generating PDF certificates only...").pack(pady=20)
            
            progress_var = tk.DoubleVar()
            progress_bar = ttk.Progressbar(progress_window, variable=progress_var, 
                                         maximum=len(self.data))
            progress_bar.pack(padx=20, fill="x")
            
            status_label = ttk.Label(progress_window, text="")
            status_label.pack(pady=10)
            
            # Save current text positions to restore later
            saved_text_positions = self.text_positions.copy()
            
            for i in range(len(self.data)):
                self.index = i
                status_label.config(text=f"Processing: {self.get_current_name()}")
                progress_window.update()
                
                # Reset text positions for each certificate
                self.text_positions = {}
                
                img = self.render_image(show_guides=False)
                safe_name = sanitize_filename(self.get_current_name())
                
                # Save PDF only
                pdf_path = os.path.join(output_dir, f"{safe_name}.pdf")
                img.save(pdf_path, "PDF", resolution=300.0, quality=100)
                
                progress_var.set(i + 1)
            
            # Restore text positions
            self.text_positions = saved_text_positions
            self.index = 0
            
            progress_window.destroy()
            
            messagebox.showinfo("PDF Generation Complete", 
                              f"All {len(self.data)} certificates have been generated as PDF!\n\n"
                              f"Files saved in: {os.path.abspath(output_dir)}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate PDF certificates:\n{str(e)}")

def select_text_areas(img, df):
    """Interactive multiple text area selection with column assignment"""
    orig_w, orig_h = img.size
    
    root = tk.Tk()
    root.title("Select Text Areas - Click to add rectangles")
    root.configure(bg='#2c3e50')
    
    # Calculate window size
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    scale = min((screen_w - 200) / orig_w, (screen_h - 200) / orig_h, 1.0)
    
    disp_w, disp_h = int(orig_w * scale), int(orig_h * scale)
    disp_img = img.resize((disp_w, disp_h), Image.LANCZOS)
    tk_img = ImageTk.PhotoImage(disp_img)
    
    # Center window
    root.geometry(f"{disp_w + 40}x{disp_h + 150}+{(screen_w - disp_w) // 2}+{(screen_h - disp_h) // 2}")
    
    # Instructions
    instructions = ttk.Label(root, text="Click and drag to create rectangles for text areas", 
                           font=("Arial", 12), background='#2c3e50', foreground='white')
    instructions.pack(pady=10)
    
    # Canvas
    canvas = tk.Canvas(root, width=disp_w, height=disp_h, bg='white', cursor="crosshair")
    canvas.pack(pady=10)
    canvas.create_image(0, 0, anchor="nw", image=tk_img)
    
    # Selection variables
    rectangles = []  # List to store rectangle objects and their data
    selected_columns = []  # List to track which columns have been selected
    current_rect = None
    start_x, start_y = 0, 0
    
    # Frame for buttons
    button_frame = ttk.Frame(root)
    button_frame.pack(pady=10)
    
    # Result
    result = {"text_areas": []}
    
    def on_press(evt):
        nonlocal start_x, start_y, current_rect
        start_x, start_y = evt.x, evt.y
        current_rect = canvas.create_rectangle(evt.x, evt.y, evt.x, evt.y, 
                                             outline="lime", width=3, dash=(5, 5))
    
    def on_drag(evt):
        nonlocal current_rect
        if current_rect:
            canvas.coords(current_rect, start_x, start_y, evt.x, evt.y)
    
    def on_release(evt):
        nonlocal current_rect
        if current_rect:
            x1, y1, x2, y2 = canvas.coords(current_rect)
            # Convert back to original image coordinates
            rect_coords = (
                int(min(x1, x2) / scale), int(min(y1, y2) / scale),
                int(max(x1, x2) / scale), int(max(y1, y2) / scale)
            )
            
            # Get available columns (excluding already selected ones)
            available_columns = [col for col in df.columns if col not in selected_columns]
            
            if not available_columns:
                messagebox.showinfo("All Columns Used", "All columns have been assigned to text areas.")
                canvas.delete(current_rect)
                current_rect = None
                return
            
            # Ask user to select a column for this rectangle
            column_window = tk.Toplevel(root)
            column_window.title("Select Column")
            column_window.geometry("300x200")
            column_window.transient(root)
            column_window.grab_set()
            
            ttk.Label(column_window, text="Select the column for this area:").pack(pady=10)
            
            column_var = tk.StringVar()
            column_combo = ttk.Combobox(column_window, textvariable=column_var, 
                                      values=available_columns, width=20)
            column_combo.pack(pady=10)
            if available_columns:
                column_combo.current(0)
            
            selected = {"column": None}
            
            def confirm_column():
                selected["column"] = column_var.get()
                column_window.destroy()
            
            ttk.Button(column_window, text="Confirm", command=confirm_column).pack(pady=10)
            
            column_window.wait_window()
            
            if selected["column"]:
                # Store the rectangle and its column
                rectangles.append({
                    "rect": rect_coords,
                    "column": selected["column"],
                    "canvas_id": current_rect
                })
                # Add to selected columns list
                selected_columns.append(selected["column"])
                # Change the rectangle color to indicate it's confirmed
                canvas.itemconfig(current_rect, outline="blue", dash=())
                
                # Show selected columns info
                info_text = "Selected columns: " + ", ".join(selected_columns)
                info_label.config(text=info_text)
            else:
                # Remove the rectangle
                canvas.delete(current_rect)
            
            current_rect = None
    
    canvas.bind("<ButtonPress-1>", on_press)
    canvas.bind("<B1-Motion>", on_drag)
    canvas.bind("<ButtonRelease-1>", on_release)
    
    def finish_selection():
        # Collect all confirmed rectangles
        for rect_data in rectangles:
            result["text_areas"].append({
                "rect": rect_data["rect"],
                "column": rect_data["column"]
            })
        root.quit()
    
    ttk.Button(button_frame, text="Finish", command=finish_selection).pack(side="left", padx=5)
    ttk.Button(button_frame, text="Cancel", command=lambda: [result.clear(), root.quit()]).pack(side="left", padx=5)
    
    # Add info label to show selected columns
    info_label = ttk.Label(root, text="Selected columns: None", 
                          font=("Arial", 10), background='#2c3e50', foreground='white')
    info_label.pack(pady=5)
    
    root.mainloop()
    root.destroy()
    
    return result.get("text_areas", [])

def load_names_from_excel():
    """Load data from Excel file"""
    print("Opening file dialog for Excel file...")
    excel_path = pick_file("Select Excel file with data", 
                          [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
    
    if not excel_path:
        print("No Excel file selected")
        return None
    
    print(f"Selected Excel file: {excel_path}")
    
    try:
        # Check if file exists
        if not os.path.exists(excel_path):
            messagebox.showerror("Error", f"File not found: {excel_path}")
            return None
        
        print("Reading Excel file...")
        # Read Excel file
        df = pd.read_excel(excel_path)
        print(f"Excel file loaded. Shape: {df.shape}")
        print(f"Columns: {list(df.columns)}")
        
        if df.empty:
            messagebox.showerror("Error", "The Excel file is empty!")
            return None
        
        return df
    
    except ImportError as e:
        messagebox.showerror("Missing Library", 
                           f"Required library missing: {str(e)}\n\n"
                           f"Please install required packages:\n"
                           f"pip install pandas openpyxl xlrd")
        return None
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load Excel file:\n{str(e)}")
        print(f"Excel loading error: {e}")
        return None

def main():
    """Main application entry point"""
    try:
        print("=== Professional Certificate Editor Starting ===")
        
        # Check required libraries
        try:
            import pandas as pd
            from PIL import Image
            print("✓ Required libraries found")
        except ImportError as e:
            error_msg = f"Missing required library: {str(e)}\n\nPlease install:\npip install pillow pandas openpyxl xlrd"
            messagebox.showerror("Missing Dependencies", error_msg)
            print(f"✗ {error_msg}")
            return
        
        # Welcome screen
        print("Showing welcome screen...")
        splash = tk.Tk()
        splash.title("Professional Certificate Editor")
        splash.geometry("600x350")
        splash.configure(bg='#2c3e50')
        splash.resizable(False, False)
        
        # Center splash screen
        splash.update_idletasks()
        x = (splash.winfo_screenwidth() // 2) - (300)
        y = (splash.winfo_screenheight() // 2) - (175)
        splash.geometry(f"+{x}+{y}")
        
        # Splash content
        title_label = tk.Label(splash, text="Professional Certificate Editor", 
                              font=("Arial", 20, "bold"), fg='white', bg='#2c3e50')
        title_label.pack(pady=20)
        
        subtitle_label = tk.Label(splash, text="Generate beautiful certificates with ease", 
                                 font=("Arial", 12), fg='#ecf0f1', bg='#2c3e50')
        subtitle_label.pack(pady=5)
        
        version_label = tk.Label(splash, text="Version 2.0 - Professional Edition", 
                                font=("Arial", 10), fg='#95a5a6', bg='#2c3e50')
        version_label.pack(pady=5)
        
        steps_frame = tk.Frame(splash, bg='#2c3e50')
        steps_frame.pack(pady=20)
        
        steps = [
            "1. Select your certificate template image (PNG, JPG)",
            "2. Choose Excel file with data (.xlsx, .xls)",
            "3. Mark the text areas on template and assign columns",
            "4. Customize fonts, colors, and positioning",
            "5. Generate individual or batch certificates"
        ]
        
        for step in steps:
            step_label = tk.Label(steps_frame, text=step, font=("Arial", 10), 
                                 fg='#bdc3c7', bg='#2c3e50')
            step_label.pack(anchor='w', pady=3)
        
        button_frame = tk.Frame(splash, bg='#2c3e50')
        button_frame.pack(pady=20)
        
        start_button = tk.Button(button_frame, text="🚀 Start Creating Certificates", 
                                font=("Arial", 12, "bold"), bg='#3498db', fg='white',
                                command=splash.destroy, cursor="hand2", padx=20, pady=10)
        start_button.pack(side='left', padx=5)
        
        exit_button = tk.Button(button_frame, text="❌ Exit", 
                               font=("Arial", 10), bg='#e74c3c', fg='white',
                               command=lambda: [splash.destroy(), sys.exit()], cursor="hand2", padx=15, pady=8)
        exit_button.pack(side='left', padx=5)
        
        splash.mainloop()
        
        # Step 1: Select template
        print("\n=== Step 1: Select Certificate Template ===")
        print("Opening file dialog for template image...")
        template_path = pick_file("Select Certificate Template Image", 
                                 [("Image files", "*.jpg *.jpeg *.png *.bmp *.tiff *.gif"), 
                                  ("PNG files", "*.png"),
                                  ("JPEG files", "*.jpg *.jpeg"),
                                  ("All files", "*.*")])
        
        if not template_path:
            print("✗ Template selection cancelled")
            messagebox.showinfo("Cancelled", "Template selection cancelled.")
            return
        
        print(f"✓ Selected template: {template_path}")
        
        # Step 2: Load data from Excel
        print("\n=== Step 2: Load Data from Excel ===")
        df = load_names_from_excel()
        
        if df is None:
            print("✗ Data loading cancelled or failed")
            messagebox.showinfo("Cancelled", "Data loading cancelled.")
            return
        
        print(f"✓ Loaded data: {df.shape[0]} rows, {df.shape[1]} columns")
        print(f"   Columns: {list(df.columns)}")
        
        # Step 3: Load template image
        print(f"\n=== Step 3: Loading Template Image ===")
        try:
            if not os.path.exists(template_path):
                raise FileNotFoundError(f"Template file not found: {template_path}")
            
            print("Loading and converting image...")
            img_template = Image.open(template_path).convert("RGB")
            print(f"✓ Template loaded: {img_template.size[0]}x{img_template.size[1]} pixels")
        except Exception as e:
            print(f"✗ Failed to load template: {e}")
            messagebox.showerror("Error", f"Failed to load template image:\n{str(e)}")
            return
        
        # Step 4: Select multiple text areas and assign columns
        print(f"\n=== Step 4: Select Text Areas ===")
        print("Opening text area selection window...")
        text_areas = select_text_areas(img_template, df)
        
        if not text_areas:
            print("✗ Text area selection cancelled")
            messagebox.showinfo("Cancelled", "Text area selection cancelled.")
            return
        
        print(f"✓ Selected {len(text_areas)} text areas")
        for i, area in enumerate(text_areas):
            print(f"   Area {i+1}: {area['column']} at {area['rect']}")
        
        # Step 5: Start the editor
        print(f"\n=== Step 5: Starting Certificate Editor ===")
        print(f"✓ Configuration complete:")
        print(f"   Template: {os.path.basename(template_path)}")
        print(f"   Data: {df.shape[0]} rows, {df.shape[1]} columns")
        print(f"   Text areas: {len(text_areas)}")
        print(f"   Starting editor...")
        
        # Launch the professional editor
        ProfessionalCertificateEditor(img_template, df, text_areas)
        
        print("=== Certificate Editor Session Ended ===")
        
    except KeyboardInterrupt:
        print("\n✗ Application interrupted by user")
    except Exception as e:
        error_msg = f"An unexpected error occurred:\n{str(e)}"
        print(f"✗ {error_msg}")
        messagebox.showerror("Error", error_msg)
        import traceback
        print("Full error trace:")
        traceback.print_exc()

if __name__ == "__main__":
    main()