import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageTk
from barcode import Code128
from barcode.writer import ImageWriter
from io import BytesIO
import os
import sys
import json
import threading
import uuid
import base64
from datetime import datetime

# --- Theme Setup ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

# --- Configuration Loading ---
def load_config():
    """Load configuration from config.json"""
    config_path = os.path.join(os.path.dirname(__file__), 'config.json')
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        # Default configuration if file doesn't exist
        return {
            "business": {
                "name": "Gift Card Generator",
                "default_font": "Arial",
                "default_font_size": 12
            },
            "barcode": {
                "format": "Code128",
                "default_size": "Medium"
            },
            "ui": {
                "window_width": 1200,
                "window_height": 800,
                "canvas_width": 400,
                "canvas_height": 250
            }
        }

# Load configuration
CONFIG = load_config()

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

def show_toast(widget, message, duration=3000, color="#00FF00"):
    """Show a temporary toast message"""
    toast = ctk.CTkLabel(widget, text=message, text_color=color, font=("Arial", 12, "bold"))
    toast.pack(pady=5)
    widget.after(duration, toast.destroy)

def safe_get_input(entry, default="", strip=True, convert_type=None):
    """Safely get input from entry widget with optional type conversion"""
    try:
        value = entry.get()
        if strip:
            value = value.strip()
        if not value:
            return default
        if convert_type:
            return convert_type(value)
        return value
    except Exception:
        return default

class GiftCardGenerator(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"Gift Card Generator - {CONFIG['business']['name']}")
        self.geometry(f"{CONFIG['ui']['window_width']}x{CONFIG['ui']['window_height']}")
        self.configure(fg_color="#121212")
        
        # Initialize canvas dimensions
        self.canvas_width = CONFIG['ui']['canvas_width']
        self.canvas_height = CONFIG['ui']['canvas_height']
        
        # Initialize variables
        self.background_path = None
        self.data_path = None
        self.output_path = None
        
        # Positioning variables
        self.barcode_position_var = tk.StringVar(value="Bottom-Right")
        self.barcode_x = tk.StringVar(value="85")
        self.barcode_y = tk.StringVar(value="85")
        self.barcode_size_var = tk.StringVar(value=CONFIG['barcode']['default_size'])
        
        self.text_position_var = tk.StringVar(value="Bottom-Left")
        self.text_x = tk.StringVar(value="15")
        self.text_y = tk.StringVar(value="85")
        self.text_background_var = tk.StringVar(value="White Box")
        self.text_transparency = tk.DoubleVar(value=50.0)
        self.text_alignment_var = tk.StringVar(value="Left")
        self.text_scale = tk.DoubleVar(value=100.0)
        
        # Preview canvas variables
        self.preview_canvas = None
        self.preview_image = None
        self.canvas_scale = 1.0
        self.dragging_item = None
        self.drag_data = {"x": 0, "y": 0}
        
        # Canvas item IDs
        self.barcode_rect = None
        self.text_rect = None
        
        # Preview variables
        self.preview_bg_photo = None
        self.preview_update_timer = None
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the user interface"""
        # Create main container
        main_container = ctk.CTkFrame(self, fg_color="#121212")
        main_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create left panel for controls
        left_panel = ctk.CTkFrame(main_container, fg_color="#1E1E1E", corner_radius=12, width=500)
        left_panel.pack(side="left", fill="y", padx=(0, 10))
        left_panel.pack_propagate(False)
        
        # Create scrollable frame for controls
        self.scrollable_frame = ctk.CTkScrollableFrame(
            left_panel, 
            fg_color="#1E1E1E",
            scrollbar_button_color="#333333",
            scrollbar_button_hover_color="#444444"
        )
        self.scrollable_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Title
        ctk.CTkLabel(self.scrollable_frame, text="Gift Card Generator", 
                    font=("Segoe UI", 18, "bold")).pack(pady=(10, 15), padx=20, fill="x")
        
        self.setup_file_selection()
        self.setup_column_configuration()
        self.setup_barcode_configuration()
        self.setup_text_configuration()
        self.setup_output_configuration()
        self.setup_generation_controls()
        
        # Right panel for preview
        self.setup_preview_panel(main_container)
    
    def setup_file_selection(self):
        """Setup file selection section"""
        file_frame = ctk.CTkFrame(self.scrollable_frame, fg_color="#2B2B2B", corner_radius=12)
        file_frame.pack(pady=(10, 5), padx=10, fill="x")
        
        ctk.CTkLabel(file_frame, text="📁 File Selection", 
                    font=("Segoe UI", 14, "bold")).pack(pady=(8, 5), padx=20, fill="x")
        
        # Background image selection
        bg_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
        bg_frame.pack(pady=(3, 3), padx=20, fill="x")
        ctk.CTkLabel(bg_frame, text="Background Image:", width=120, anchor="w").pack(side="left", padx=(0, 10))
        self.bg_path_var = tk.StringVar(value="No file selected")
        self.bg_path_label = ctk.CTkLabel(bg_frame, textvariable=self.bg_path_var, anchor="w")
        self.bg_path_label.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(bg_frame, text="Browse", command=self.select_background, width=80).pack(side="right")
        
        # Data file selection
        data_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
        data_frame.pack(pady=(3, 8), padx=20, fill="x")
        ctk.CTkLabel(data_frame, text="Data File (CSV/Excel):", width=120, anchor="w").pack(side="left", padx=(0, 10))
        self.data_path_var = tk.StringVar(value="No file selected")
        self.data_path_label = ctk.CTkLabel(data_frame, textvariable=self.data_path_var, anchor="w")
        self.data_path_label.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(data_frame, text="Browse", command=self.select_data_file, width=80).pack(side="right")
    
    def setup_column_configuration(self):
        """Setup column configuration section"""
        config_frame = ctk.CTkFrame(self.scrollable_frame, fg_color="#2B2B2B", corner_radius=12)
        config_frame.pack(pady=(5, 5), padx=10, fill="x")
        
        ctk.CTkLabel(config_frame, text="⚙️ Column Configuration", 
                    font=("Segoe UI", 14, "bold")).pack(pady=(8, 5), padx=20, fill="x")
        
        # Column name inputs
        col_grid = ctk.CTkFrame(config_frame, fg_color="transparent")
        col_grid.pack(pady=(3, 8), padx=20, fill="x")
        
        # Configure grid weights
        col_grid.columnconfigure(1, weight=1)
        
        # Barcode column
        ctk.CTkLabel(col_grid, text="Barcode Column:", width=120, anchor="w").grid(row=0, column=0, padx=(0, 10), pady=2, sticky="w")
        self.barcode_col = ctk.CTkEntry(col_grid, placeholder_text="e.g., barcode")
        self.barcode_col.grid(row=0, column=1, padx=(0, 20), pady=2, sticky="ew")
        
        # Member number column
        ctk.CTkLabel(col_grid, text="Member Number:", width=120, anchor="w").grid(row=1, column=0, padx=(0, 10), pady=2, sticky="w")
        self.member_col = ctk.CTkEntry(col_grid, placeholder_text="e.g., member_number")
        self.member_col.grid(row=1, column=1, padx=(0, 20), pady=2, sticky="ew")
        
        # Verification code column
        ctk.CTkLabel(col_grid, text="Verification Code:", width=120, anchor="w").grid(row=2, column=0, padx=(0, 10), pady=2, sticky="w")
        self.verification_col = ctk.CTkEntry(col_grid, placeholder_text="e.g., verification_code")
        self.verification_col.grid(row=2, column=1, padx=(0, 20), pady=2, sticky="ew")
        
        # Card serial column
        ctk.CTkLabel(col_grid, text="Card Serial:", width=120, anchor="w").grid(row=3, column=0, padx=(0, 10), pady=2, sticky="w")
        self.serial_col = ctk.CTkEntry(col_grid, placeholder_text="e.g., card_serial")
        self.serial_col.grid(row=3, column=1, padx=(0, 20), pady=2, sticky="ew")
    
    def setup_barcode_configuration(self):
        """Setup barcode configuration section"""
        barcode_frame = ctk.CTkFrame(self.scrollable_frame, fg_color="#2B2B2B", corner_radius=12)
        barcode_frame.pack(pady=(5, 5), padx=10, fill="x")
        
        ctk.CTkLabel(barcode_frame, text="📊 Barcode Configuration", 
                    font=("Segoe UI", 14, "bold")).pack(pady=(8, 5), padx=20, fill="x")
        
        # Barcode position and size
        pos_frame = ctk.CTkFrame(barcode_frame, fg_color="transparent")
        pos_frame.pack(pady=(3, 3), padx=20, fill="x")
        
        ctk.CTkLabel(pos_frame, text="Position:", width=80, anchor="w").pack(side="left")
        position_menu = ctk.CTkOptionMenu(pos_frame, variable=self.barcode_position_var,
                                        values=["Top-Left", "Top-Right", "Bottom-Left", "Bottom-Right", "Center", "Custom"],
                                        command=self.on_barcode_position_change)
        position_menu.pack(side="left", padx=(10, 20))
        
        ctk.CTkLabel(pos_frame, text="Size:", width=60, anchor="w").pack(side="left")
        size_menu = ctk.CTkOptionMenu(pos_frame, variable=self.barcode_size_var,
                                    values=["Small", "Medium", "Large", "XL"],
                                    command=self.update_preview)
        size_menu.pack(side="left", padx=(10, 0))
        
        # Custom position controls
        self.custom_barcode_frame = ctk.CTkFrame(barcode_frame, fg_color="transparent")
        self.custom_barcode_frame.pack(pady=(3, 8), padx=20, fill="x")
        
        ctk.CTkLabel(self.custom_barcode_frame, text="X Position:", width=80, anchor="w").pack(side="left")
        x_entry = ctk.CTkEntry(self.custom_barcode_frame, textvariable=self.barcode_x, width=60)
        x_entry.pack(side="left", padx=(10, 20))
        x_entry.bind('<KeyRelease>', lambda e: self.update_preview())
        
        ctk.CTkLabel(self.custom_barcode_frame, text="Y Position:", width=80, anchor="w").pack(side="left")
        y_entry = ctk.CTkEntry(self.custom_barcode_frame, textvariable=self.barcode_y, width=60)
        y_entry.pack(side="left", padx=(10, 0))
        y_entry.bind('<KeyRelease>', lambda e: self.update_preview())
        
        # Initially hide custom controls
        self.custom_barcode_frame.pack_forget()
    
    def setup_text_configuration(self):
        """Setup text configuration section"""
        text_frame = ctk.CTkFrame(self.scrollable_frame, fg_color="#2B2B2B", corner_radius=12)
        text_frame.pack(pady=(5, 5), padx=10, fill="x")
        
        ctk.CTkLabel(text_frame, text="📝 Text Configuration", 
                    font=("Segoe UI", 14, "bold")).pack(pady=(8, 5), padx=20, fill="x")
        
        # Text position
        text_pos_frame = ctk.CTkFrame(text_frame, fg_color="transparent")
        text_pos_frame.pack(pady=(3, 3), padx=20, fill="x")
        
        ctk.CTkLabel(text_pos_frame, text="Position:", width=80, anchor="w").pack(side="left")
        text_position_menu = ctk.CTkOptionMenu(text_pos_frame, variable=self.text_position_var,
                                             values=["Top-Left", "Top-Right", "Bottom-Left", "Bottom-Right", "Center", "Custom"],
                                             command=self.on_text_position_change)
        text_position_menu.pack(side="left", padx=(10, 20))
        
        ctk.CTkLabel(text_pos_frame, text="Background:", width=80, anchor="w").pack(side="left")
        bg_menu = ctk.CTkOptionMenu(text_pos_frame, variable=self.text_background_var,
                                  values=["None", "White Box", "Black Box", "Semi-Transparent White", "Semi-Transparent Black"],
                                  command=self.update_preview)
        bg_menu.pack(side="left", padx=(10, 0))
        
        # Custom text position controls
        self.custom_text_frame = ctk.CTkFrame(text_frame, fg_color="transparent")
        self.custom_text_frame.pack(pady=(3, 3), padx=20, fill="x")
        
        ctk.CTkLabel(self.custom_text_frame, text="X Position:", width=80, anchor="w").pack(side="left")
        text_x_entry = ctk.CTkEntry(self.custom_text_frame, textvariable=self.text_x, width=60)
        text_x_entry.pack(side="left", padx=(10, 20))
        text_x_entry.bind('<KeyRelease>', lambda e: self.update_preview())
        
        ctk.CTkLabel(self.custom_text_frame, text="Y Position:", width=80, anchor="w").pack(side="left")
        text_y_entry = ctk.CTkEntry(self.custom_text_frame, textvariable=self.text_y, width=60)
        text_y_entry.pack(side="left", padx=(10, 0))
        text_y_entry.bind('<KeyRelease>', lambda e: self.update_preview())
        
        # Initially hide custom controls
        self.custom_text_frame.pack_forget()
        
        # Text alignment and scale
        align_frame = ctk.CTkFrame(text_frame, fg_color="transparent")
        align_frame.pack(pady=(3, 8), padx=20, fill="x")
        
        ctk.CTkLabel(align_frame, text="Alignment:", width=80, anchor="w").pack(side="left")
        align_menu = ctk.CTkOptionMenu(align_frame, variable=self.text_alignment_var,
                                     values=["Left", "Center", "Right"],
                                     command=self.update_preview)
        align_menu.pack(side="left", padx=(10, 20))
        
        ctk.CTkLabel(align_frame, text="Scale:", width=60, anchor="w").pack(side="left")
        scale_slider = ctk.CTkSlider(align_frame, from_=50, to=200, 
                                   variable=self.text_scale, command=self.update_preview)
        scale_slider.pack(side="left", padx=(10, 10), fill="x", expand=True)
        
        scale_label = ctk.CTkLabel(align_frame, text="100%", width=40)
        scale_label.pack(side="left")
        
        # Update scale label
        def update_scale_label(*args):
            scale_label.configure(text=f"{int(self.text_scale.get())}%")
        self.text_scale.trace('w', update_scale_label)
    
    def setup_output_configuration(self):
        """Setup output configuration section"""
        output_frame = ctk.CTkFrame(self.scrollable_frame, fg_color="#2B2B2B", corner_radius=12)
        output_frame.pack(pady=(5, 5), padx=10, fill="x")
        
        ctk.CTkLabel(output_frame, text="💾 Output Configuration", 
                    font=("Segoe UI", 14, "bold")).pack(pady=(8, 5), padx=20, fill="x")
        
        # Output folder selection
        folder_frame = ctk.CTkFrame(output_frame, fg_color="transparent")
        folder_frame.pack(pady=(3, 8), padx=20, fill="x")
        ctk.CTkLabel(folder_frame, text="Output Folder:", width=120, anchor="w").pack(side="left", padx=(0, 10))
        self.output_path_var = tk.StringVar(value="No folder selected")
        self.output_path_label = ctk.CTkLabel(folder_frame, textvariable=self.output_path_var, anchor="w")
        self.output_path_label.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(folder_frame, text="Browse", command=self.select_output_folder, width=80).pack(side="right")
    
    def setup_generation_controls(self):
        """Setup generation control buttons"""
        control_frame = ctk.CTkFrame(self.scrollable_frame, fg_color="#2B2B2B", corner_radius=12)
        control_frame.pack(pady=(5, 5), padx=10, fill="x")
        
        ctk.CTkLabel(control_frame, text="🚀 Generation Controls", 
                    font=("Segoe UI", 14, "bold")).pack(pady=(8, 5), padx=20, fill="x")
        
        button_frame = ctk.CTkFrame(control_frame, fg_color="transparent")
        button_frame.pack(pady=(3, 8), padx=20, fill="x")
        
        # Preview button
        self.preview_button = ctk.CTkButton(button_frame, text="Preview Sample", 
                                          command=self.threaded_preview, height=35)
        self.preview_button.pack(side="left", padx=(0, 10), fill="x", expand=True)
        
        # Generate button
        self.generate_button = ctk.CTkButton(button_frame, text="Generate All Cards", 
                                           command=self.threaded_generate, height=35)
        self.generate_button.pack(side="right", fill="x", expand=True)
        
        # Status label
        self.status_label = ctk.CTkLabel(control_frame, text="Ready to generate gift cards", 
                                       font=("Arial", 11), text_color="#CCCCCC")
        self.status_label.pack(pady=(0, 8))
    
    def setup_preview_panel(self, parent):
        """Setup the preview panel"""
        right_panel = ctk.CTkFrame(parent, fg_color="#1E1E1E", corner_radius=12)
        right_panel.pack(side="right", fill="both", expand=True)
        
        # Preview title
        ctk.CTkLabel(right_panel, text="👁️ Live Preview", 
                    font=("Segoe UI", 16, "bold")).pack(pady=(15, 10))
        
        # Canvas for preview
        canvas_frame = ctk.CTkFrame(right_panel, fg_color="#2B2B2B", corner_radius=8)
        canvas_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        self.preview_canvas = tk.Canvas(canvas_frame, 
                                      width=self.canvas_width, 
                                      height=self.canvas_height,
                                      bg="#FFFFFF", 
                                      highlightthickness=1, 
                                      highlightcolor="#444444")
        self.preview_canvas.pack(pady=20)
        
        # Preview info
        info_text = ("Select a background image and data file to see preview.\n"
                    "Drag elements on the canvas to reposition them.")
        ctk.CTkLabel(right_panel, text=info_text, 
                    font=("Arial", 11), text_color="#888888", 
                    justify="center").pack(pady=(0, 20))
    
    def log(self, message):
        """Log a message to the status label"""
        self.status_label.configure(text=message)
        self.update()
    
    def select_background(self):
        """Select background image file"""
        file_path = filedialog.askopenfilename(
            title="Select Background Image",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.gif *.tiff")]
        )
        if file_path:
            self.background_path = file_path
            self.bg_path_var.set(os.path.basename(file_path))
            self.update_preview()
    
    def select_data_file(self):
        """Select data file (CSV or Excel)"""
        file_path = filedialog.askopenfilename(
            title="Select Data File",
            filetypes=[("Data files", "*.csv *.xlsx *.xls")]
        )
        if file_path:
            self.data_path = file_path
            self.data_path_var.set(os.path.basename(file_path))
            self.update_preview()
    
    def select_output_folder(self):
        """Select output folder"""
        folder_path = filedialog.askdirectory(title="Select Output Folder")
        if folder_path:
            self.output_path = folder_path
            self.output_path_var.set(folder_path)
    
    def on_barcode_position_change(self, value):
        """Handle barcode position change"""
        if value == "Custom":
            self.custom_barcode_frame.pack(pady=(3, 8), padx=20, fill="x")
        else:
            self.custom_barcode_frame.pack_forget()
            # Set preset positions
            positions = {
                "Top-Left": ("10", "10"),
                "Top-Right": ("85", "10"),
                "Bottom-Left": ("10", "85"),
                "Bottom-Right": ("85", "85"),
                "Center": ("50", "50")
            }
            if value in positions:
                x, y = positions[value]
                self.barcode_x.set(x)
                self.barcode_y.set(y)
        self.update_preview()
    
    def on_text_position_change(self, value):
        """Handle text position change"""
        if value == "Custom":
            self.custom_text_frame.pack(pady=(3, 3), padx=20, fill="x")
        else:
            self.custom_text_frame.pack_forget()
            # Set preset positions
            positions = {
                "Top-Left": ("15", "15"),
                "Top-Right": ("85", "15"),
                "Bottom-Left": ("15", "85"),
                "Bottom-Right": ("85", "85"),
                "Center": ("50", "50")
            }
            if value in positions:
                x, y = positions[value]
                self.text_x.set(x)
                self.text_y.set(y)
        self.update_preview()
    
    def threaded_preview(self):
        """Generate preview in a separate thread"""
        threading.Thread(target=self.generate_preview, daemon=True).start()
    
    def threaded_generate(self):
        """Generate gift cards in a separate thread"""
        threading.Thread(target=self.generate_gift_cards, daemon=True).start()
    
    def generate_preview(self):
        """Generate a preview of the gift card"""
        if not self.background_path:
            self.log("❌ Please select a background image")
            return
        
        try:
            # Create sample data for preview
            sample_data = {
                "barcode": "1234567890123",
                "member_number": "12345",
                "verification_code": "ABCD1234",
                "card_serial": "GC001"
            }
            
            self.log("🔄 Generating preview...")
            
            # Generate preview image
            preview_image = self.create_gift_card_image(
                self.background_path,
                sample_data["barcode"],
                sample_data["member_number"],
                sample_data["verification_code"],
                sample_data["card_serial"]
            )
            
            if preview_image:
                # Resize for preview canvas
                preview_image = preview_image.resize((self.canvas_width, self.canvas_height), Image.Resampling.LANCZOS)
                
                # Convert to PhotoImage for tkinter
                self.preview_bg_photo = ImageTk.PhotoImage(preview_image)
                
                # Update canvas
                self.preview_canvas.delete("all")
                self.preview_canvas.create_image(
                    self.canvas_width//2, self.canvas_height//2,
                    image=self.preview_bg_photo
                )
                
                self.log("✅ Preview generated successfully")
            else:
                self.log("❌ Failed to generate preview")
                
        except Exception as e:
            self.log(f"❌ Preview error: {str(e)}")
    
    def update_preview(self, *args):
        """Update the preview when settings change"""
        if self.preview_update_timer:
            self.after_cancel(self.preview_update_timer)
        self.preview_update_timer = self.after(500, self.generate_preview)  # Debounce updates
    
    def generate_barcode(self, barcode_data):
        """Generate barcode image"""
        try:
            # Generate Code128 barcode
            code128 = Code128(str(barcode_data), writer=ImageWriter())
            
            # Create barcode image
            barcode_buffer = BytesIO()
            code128.write(barcode_buffer)
            barcode_buffer.seek(0)
            
            # Open as PIL Image
            barcode_image = Image.open(barcode_buffer)
            
            # Size mapping
            size_mapping = {
                "Small": (120, 30),
                "Medium": (160, 40),
                "Large": (200, 50),
                "XL": (240, 60)
            }
            
            size = size_mapping.get(self.barcode_size_var.get(), (160, 40))
            barcode_image = barcode_image.resize(size, Image.Resampling.LANCZOS)
            
            return barcode_image
            
        except Exception as e:
            self.log(f"❌ Barcode generation error: {str(e)}")
            return None
    
    def create_gift_card_image(self, background_path, barcode_data, member_number, verification_code, card_serial):
        """Create a single gift card image"""
        try:
            # Open background image
            background = Image.open(background_path).convert("RGBA")
            
            # Generate barcode
            barcode_image = self.generate_barcode(barcode_data)
            if not barcode_image:
                return None
            
            # Convert barcode to RGBA
            barcode_image = barcode_image.convert("RGBA")
            
            # Calculate barcode position
            bg_width, bg_height = background.size
            barcode_x_percent = float(self.barcode_x.get())
            barcode_y_percent = float(self.barcode_y.get())
            
            barcode_x = int((barcode_x_percent / 100) * bg_width) - barcode_image.width // 2
            barcode_y = int((barcode_y_percent / 100) * bg_height) - barcode_image.height // 2
            
            # Ensure barcode stays within bounds
            barcode_x = max(0, min(barcode_x, bg_width - barcode_image.width))
            barcode_y = max(0, min(barcode_y, bg_height - barcode_image.height))
            
            # Paste barcode onto background
            background.paste(barcode_image, (barcode_x, barcode_y), barcode_image)
            
            # Add text information
            self.draw_text_block(background, member_number, verification_code, card_serial)
            
            return background.convert("RGB")
            
        except Exception as e:
            self.log(f"❌ Gift card creation error: {str(e)}")
            return None
    
    def draw_text_block(self, image, member_number, verification_code, card_serial):
        """Draw text block on the image"""
        try:
            draw = ImageDraw.Draw(image)
            
            # Prepare text
            text_lines = [
                f"Member: {member_number}",
                f"Code: {verification_code}",
                f"Serial: {card_serial}"
            ]
            
            # Calculate text position
            img_width, img_height = image.size
            text_x_percent = float(self.text_x.get())
            text_y_percent = float(self.text_y.get())
            
            # Try to load font
            try:
                font_size = int(12 * (self.text_scale.get() / 100))
                font = ImageFont.truetype("arial.ttf", font_size)
            except:
                font = ImageFont.load_default()
            
            # Calculate text dimensions
            text_heights = [draw.textbbox((0, 0), line, font=font)[3] for line in text_lines]
            max_text_width = max([draw.textbbox((0, 0), line, font=font)[2] for line in text_lines])
            total_text_height = sum(text_heights) + (len(text_lines) - 1) * 5  # 5px spacing
            
            # Calculate position
            text_x = int((text_x_percent / 100) * img_width)
            text_y = int((text_y_percent / 100) * img_height)
            
            # Adjust for alignment
            if self.text_alignment_var.get() == "Center":
                text_x -= max_text_width // 2
            elif self.text_alignment_var.get() == "Right":
                text_x -= max_text_width
                
            text_y -= total_text_height // 2
            
            # Draw background if specified
            background_type = self.text_background_var.get()
            if background_type != "None":
                padding = 10
                bg_x1 = text_x - padding
                bg_y1 = text_y - padding
                bg_x2 = text_x + max_text_width + padding
                bg_y2 = text_y + total_text_height + padding
                
                if background_type == "White Box":
                    draw.rectangle([bg_x1, bg_y1, bg_x2, bg_y2], fill=(255, 255, 255, 255))
                elif background_type == "Black Box":
                    draw.rectangle([bg_x1, bg_y1, bg_x2, bg_y2], fill=(0, 0, 0, 255))
                elif background_type == "Semi-Transparent White":
                    # Create semi-transparent overlay
                    overlay = Image.new('RGBA', image.size, (255, 255, 255, int(255 * self.text_transparency.get() / 100)))
                    mask = Image.new('RGBA', image.size, (0, 0, 0, 0))
                    mask_draw = ImageDraw.Draw(mask)
                    mask_draw.rectangle([bg_x1, bg_y1, bg_x2, bg_y2], fill=(255, 255, 255, 255))
                    image.paste(overlay, (0, 0), mask)
                elif background_type == "Semi-Transparent Black":
                    # Create semi-transparent overlay
                    overlay = Image.new('RGBA', image.size, (0, 0, 0, int(255 * self.text_transparency.get() / 100)))
                    mask = Image.new('RGBA', image.size, (0, 0, 0, 0))
                    mask_draw = ImageDraw.Draw(mask)
                    mask_draw.rectangle([bg_x1, bg_y1, bg_x2, bg_y2], fill=(255, 255, 255, 255))
                    image.paste(overlay, (0, 0), mask)
            
            # Draw text lines
            current_y = text_y
            text_color = (0, 0, 0) if background_type in ["White Box", "Semi-Transparent White"] else (255, 255, 255)
            
            for line in text_lines:
                if self.text_alignment_var.get() == "Center":
                    line_width = draw.textbbox((0, 0), line, font=font)[2]
                    line_x = text_x + (max_text_width - line_width) // 2
                elif self.text_alignment_var.get() == "Right":
                    line_width = draw.textbbox((0, 0), line, font=font)[2]
                    line_x = text_x + (max_text_width - line_width)
                else:
                    line_x = text_x
                
                draw.text((line_x, current_y), line, font=font, fill=text_color)
                current_y += draw.textbbox((0, 0), line, font=font)[3] + 5
                
        except Exception as e:
            self.log(f"❌ Text drawing error: {str(e)}")
    
    def generate_gift_cards(self):
        """Generate all gift cards from data file"""
        if not self.background_path:
            self.log("❌ Please select a background image")
            return
        
        if not self.data_path:
            self.log("❌ Please select a data file")
            return
        
        if not self.output_path:
            self.log("❌ Please select an output folder")
            return
        
        try:
            self.log("📄 Reading data file...")
            
            # Read data file
            if self.data_path.endswith('.csv'):
                df = pd.read_csv(self.data_path)
            else:
                df = pd.read_excel(self.data_path)
            
            # Get column names
            barcode_col = safe_get_input(self.barcode_col, "barcode")
            member_col = safe_get_input(self.member_col, "member_number")
            verification_col = safe_get_input(self.verification_col, "verification_code")
            serial_col = safe_get_input(self.serial_col, "card_serial")
            
            # Validate columns exist
            missing_cols = []
            for col_name, col_entry in [(barcode_col, "Barcode"), (member_col, "Member Number"), 
                                       (verification_col, "Verification Code"), (serial_col, "Card Serial")]:
                if col_name not in df.columns:
                    missing_cols.append(f"{col_entry} ({col_name})")
            
            if missing_cols:
                self.log(f"❌ Missing columns: {', '.join(missing_cols)}")
                return
            
            self.log(f"🔄 Generating {len(df)} gift cards...")
            
            # Generate cards
            success_count = 0
            for index, row in df.iterrows():
                try:
                    # Extract data
                    barcode_data = row[barcode_col]
                    member_number = row[member_col]
                    verification_code = row[verification_col]
                    card_serial = row[serial_col]
                    
                    # Generate gift card
                    card_image = self.create_gift_card_image(
                        self.background_path,
                        barcode_data,
                        member_number,
                        verification_code,
                        card_serial
                    )
                    
                    if card_image:
                        # Save image
                        filename = f"gift_card_{card_serial}_{index+1:04d}.png"
                        output_file = os.path.join(self.output_path, filename)
                        card_image.save(output_file, "PNG", quality=95)
                        success_count += 1
                        
                        # Update progress
                        progress = f"Progress: {success_count}/{len(df)} cards generated"
                        self.log(progress)
                    
                except Exception as e:
                    self.log(f"❌ Error generating card {index+1}: {str(e)}")
                    continue
            
            self.log(f"✅ Generated {success_count}/{len(df)} gift cards successfully!")
            
            if success_count > 0:
                show_toast(self.scrollable_frame, f"🎉 {success_count} gift cards generated!", 5000, "#00FF00")
            
        except Exception as e:
            self.log(f"❌ Generation error: {str(e)}")

if __name__ == "__main__":
    app = GiftCardGenerator()
    app.mainloop()