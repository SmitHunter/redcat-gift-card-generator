import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
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
        
        # Initialize canvas dimensions first
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
        self.barcode_size_var = tk.StringVar(value="Medium")
        
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
        
        # Custom color variable
        self.custom_bg_color = "#E0E0E0"
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the user interface"""
        # Create scrollable frame for all content
        self.scrollable_frame = ctk.CTkScrollableFrame(
            self, 
            fg_color="#121212",
            scrollbar_button_color="#333333",
            scrollbar_button_hover_color="#444444"
        )
        self.scrollable_frame.pack(fill="both", expand=True)
        
        # Title
        ctk.CTkLabel(self.scrollable_frame, text="Gift Card Generator", font=("Segoe UI", 18, "bold")).pack(pady=(20, 15), padx=20, fill="x")
        
        self.setup_file_selection()
        self.setup_column_configuration()
        self.setup_layout_designer()
        self.setup_output_settings()
        self.setup_generation_controls()
        self.setup_log_section()
        
        # Initialize event bindings
        self.setup_event_bindings()
    
    def setup_file_selection(self):
        """Setup file selection section"""
        file_frame = ctk.CTkFrame(self.scrollable_frame, fg_color="#1E1E1E", corner_radius=12)
        file_frame.pack(pady=(10, 5), padx=20, fill="x")
        
        ctk.CTkLabel(file_frame, text="📁 File Selection", font=("Segoe UI", 14, "bold")).pack(pady=(8, 5), padx=20, fill="x")
        
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
        config_frame = ctk.CTkFrame(self.scrollable_frame, fg_color="#1E1E1E", corner_radius=12)
        config_frame.pack(pady=(5, 5), padx=20, fill="x")
        
        ctk.CTkLabel(config_frame, text="⚙️ Column Configuration", font=("Segoe UI", 14, "bold")).pack(pady=(8, 5), padx=20, fill="x")
        
        # Column name inputs
        col_grid = ctk.CTkFrame(config_frame, fg_color="transparent")
        col_grid.pack(pady=(3, 8), padx=20, fill="x")
        
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
        self.verification_col = ctk.CTkEntry(col_grid, placeholder_text="e.g., pin")
        self.verification_col.grid(row=2, column=1, padx=(0, 20), pady=2, sticky="ew")
        
        col_grid.grid_columnconfigure(1, weight=1)
        
        # Add event listeners for column configuration changes
        self.barcode_col.bind("<KeyRelease>", self.on_column_config_change)
        self.member_col.bind("<KeyRelease>", self.on_column_config_change)
        self.verification_col.bind("<KeyRelease>", self.on_column_config_change)
    
    def setup_layout_designer(self):
        """Setup layout designer section"""
        layout_frame = ctk.CTkFrame(self.scrollable_frame, fg_color="#1E1E1E", corner_radius=12)
        layout_frame.pack(pady=(5, 5), padx=20, fill="x")
        
        ctk.CTkLabel(layout_frame, text="🎨 Layout Designer", font=("Segoe UI", 14, "bold")).pack(pady=(8, 5), padx=20, fill="x")
        
        # Preview and controls container
        layout_container = ctk.CTkFrame(layout_frame, fg_color="transparent")
        layout_container.pack(pady=(3, 8), padx=20, fill="x")
        
        # Left side - Preview Canvas
        preview_frame = ctk.CTkFrame(layout_container, fg_color="#2B2B2B", corner_radius=8)
        preview_frame.pack(side="left", padx=(0, 10), fill="y")
        
        canvas_label = ctk.CTkLabel(preview_frame, text="🎴 Gift Card Preview", font=("Segoe UI", 12, "bold"))
        canvas_label.pack(pady=(5, 3))
        
        # Create live preview canvas
        self.preview_canvas = tk.Canvas(
            preview_frame, 
            width=self.canvas_width, 
            height=self.canvas_height, 
            bg="#2B2B2B", 
            highlightthickness=0
        )
        self.preview_canvas.pack(pady=(0, 5), padx=5)
        
        # Bind canvas events for drag and drop
        self.preview_canvas.bind("<Button-1>", self.on_canvas_click)
        self.preview_canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.preview_canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        
        # Right side - Position Controls
        controls_frame = ctk.CTkFrame(layout_container, fg_color="transparent")
        controls_frame.pack(side="right", fill="both", expand=True)
        
        self.setup_barcode_positioning(controls_frame)
        self.setup_text_positioning(controls_frame)
        
        # Reset button
        reset_btn = ctk.CTkButton(
            controls_frame,
            text="Reset to Default",
            command=self.reset_positions,
            height=25,
            fg_color="#555555"
        )
        reset_btn.pack(pady=(5, 0))
    
    def setup_barcode_positioning(self, parent):
        """Setup barcode positioning controls"""
        barcode_frame = ctk.CTkFrame(parent, fg_color="#333333", corner_radius=8)
        barcode_frame.pack(fill="x", pady=(0, 5))
        
        ctk.CTkLabel(barcode_frame, text="📊 Barcode Positioning", font=("Segoe UI", 12, "bold")).pack(pady=(5, 3), padx=10)
        
        # Barcode position dropdown
        pos_frame1 = ctk.CTkFrame(barcode_frame, fg_color="transparent")
        pos_frame1.pack(fill="x", padx=10, pady=2)
        
        ctk.CTkLabel(pos_frame1, text="Position:", width=80, anchor="w").pack(side="left")
        barcode_pos_combo = ctk.CTkComboBox(
            pos_frame1, 
            variable=self.barcode_position_var,
            values=["Top-Left", "Top-Right", "Bottom-Left", "Bottom-Right", "Center", "Custom"],
            command=self.on_barcode_position_change,
            width=120
        )
        barcode_pos_combo.pack(side="left", padx=(5, 0))
        
        # Custom coordinates (initially hidden)
        self.barcode_custom_frame = ctk.CTkFrame(barcode_frame, fg_color="transparent")
        
        coord_frame1 = ctk.CTkFrame(self.barcode_custom_frame, fg_color="transparent")
        coord_frame1.pack(fill="x", pady=1)
        ctk.CTkLabel(coord_frame1, text="X (%):", width=40, anchor="w").pack(side="left")
        x_entry1 = ctk.CTkEntry(coord_frame1, textvariable=self.barcode_x, width=60)
        x_entry1.pack(side="left", padx=(5, 10))
        x_entry1.bind("<KeyRelease>", self.on_position_change)
        
        ctk.CTkLabel(coord_frame1, text="Y (%):", width=40, anchor="w").pack(side="left")
        y_entry1 = ctk.CTkEntry(coord_frame1, textvariable=self.barcode_y, width=60)
        y_entry1.pack(side="left", padx=(5, 0))
        y_entry1.bind("<KeyRelease>", self.on_position_change)
        
        # Barcode size
        size_frame1 = ctk.CTkFrame(barcode_frame, fg_color="transparent")
        size_frame1.pack(fill="x", padx=10, pady=2)
        
        ctk.CTkLabel(size_frame1, text="Size:", width=80, anchor="w").pack(side="left")
        size_combo1 = ctk.CTkComboBox(
            size_frame1,
            variable=self.barcode_size_var,
            values=["Small", "Medium", "Large", "XL"],
            command=self.on_position_change,
            width=120
        )
        size_combo1.pack(side="left", padx=(5, 0))
    
    def setup_text_positioning(self, parent):
        """Setup text positioning controls"""
        text_frame = ctk.CTkFrame(parent, fg_color="#333333", corner_radius=8)
        text_frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(text_frame, text="📝 Text Block Positioning", font=("Segoe UI", 12, "bold")).pack(pady=(5, 3), padx=10)
        
        # Text position dropdown
        pos_frame2 = ctk.CTkFrame(text_frame, fg_color="transparent")
        pos_frame2.pack(fill="x", padx=10, pady=2)
        
        ctk.CTkLabel(pos_frame2, text="Position:", width=80, anchor="w").pack(side="left")
        text_pos_combo = ctk.CTkComboBox(
            pos_frame2,
            variable=self.text_position_var,
            values=["Top-Left", "Top-Right", "Bottom-Left", "Bottom-Right", "Center", "Custom"],
            command=self.on_text_position_change,
            width=120
        )
        text_pos_combo.pack(side="left", padx=(5, 0))
        
        # Custom coordinates for text
        self.text_custom_frame = ctk.CTkFrame(text_frame, fg_color="transparent")
        
        coord_frame2 = ctk.CTkFrame(self.text_custom_frame, fg_color="transparent")
        coord_frame2.pack(fill="x", pady=1)
        ctk.CTkLabel(coord_frame2, text="X (%):", width=40, anchor="w").pack(side="left")
        x_entry2 = ctk.CTkEntry(coord_frame2, textvariable=self.text_x, width=60)
        x_entry2.pack(side="left", padx=(5, 10))
        x_entry2.bind("<KeyRelease>", self.on_position_change)
        
        ctk.CTkLabel(coord_frame2, text="Y (%):", width=40, anchor="w").pack(side="left")
        y_entry2 = ctk.CTkEntry(coord_frame2, textvariable=self.text_y, width=60)
        y_entry2.pack(side="left", padx=(5, 0))
        y_entry2.bind("<KeyRelease>", self.on_position_change)
        
        # Text background style
        bg_frame = ctk.CTkFrame(text_frame, fg_color="transparent")
        bg_frame.pack(fill="x", padx=10, pady=2)
        
        ctk.CTkLabel(bg_frame, text="Background:", width=80, anchor="w").pack(side="left")
        bg_combo = ctk.CTkComboBox(
            bg_frame,
            variable=self.text_background_var,
            values=["None", "White Box", "Semi-transparent", "Custom Color"],
            command=self.on_background_change,
            width=120
        )
        bg_combo.pack(side="left", padx=(5, 0))
        
        # Transparency control (for semi-transparent backgrounds)
        self.transparency_frame = ctk.CTkFrame(text_frame, fg_color="transparent")
        
        transparency_frame = ctk.CTkFrame(self.transparency_frame, fg_color="transparent")
        transparency_frame.pack(fill="x", pady=2)
        
        ctk.CTkLabel(transparency_frame, text="Transparency (%):", width=120, anchor="w").pack(side="left")
        
        # Transparency entry field
        transparency_entry = ctk.CTkEntry(transparency_frame, textvariable=self.text_transparency, width=50)
        transparency_entry.pack(side="left", padx=(5, 5))
        transparency_entry.bind("<KeyRelease>", self.on_transparency_change)
        
        # Transparency slider
        transparency_slider = ctk.CTkSlider(
            transparency_frame,
            from_=0, to=100,
            variable=self.text_transparency,
            command=self.on_transparency_change,
            width=100
        )
        transparency_slider.pack(side="left", padx=(5, 0))
        
        # Custom color selection (initially hidden)
        self.custom_color_frame = ctk.CTkFrame(text_frame, fg_color="transparent")
        
        color_frame = ctk.CTkFrame(self.custom_color_frame, fg_color="transparent")
        color_frame.pack(fill="x", pady=1)
        ctk.CTkLabel(color_frame, text="Color:", width=80, anchor="w").pack(side="left")
        
        # Color entry field
        self.custom_color_entry = ctk.CTkEntry(color_frame, placeholder_text="#E0E0E0", width=100)
        self.custom_color_entry.pack(side="left", padx=(5, 5))
        self.custom_color_entry.bind("<KeyRelease>", self.on_custom_color_change)
        
        # Color preview button
        self.color_preview_btn = ctk.CTkButton(
            color_frame, text="", width=30, height=24, 
            fg_color="#E0E0E0", hover_color="#D0D0D0",
            command=self.open_color_picker
        )
        self.color_preview_btn.pack(side="left", padx=(0, 5))
        
        # Set initial custom color
        self.custom_color_entry.insert(0, self.custom_bg_color)
        
        # Text alignment
        align_frame = ctk.CTkFrame(text_frame, fg_color="transparent")
        align_frame.pack(fill="x", padx=10, pady=(2, 5))
        
        ctk.CTkLabel(align_frame, text="Alignment:", width=80, anchor="w").pack(side="left")
        align_combo = ctk.CTkComboBox(
            align_frame,
            variable=self.text_alignment_var,
            values=["Left", "Center", "Right"],
            command=self.on_position_change,
            width=120
        )
        align_combo.pack(side="left", padx=(5, 0))
        
        # Text scale
        text_scale_frame = ctk.CTkFrame(text_frame, fg_color="transparent")
        text_scale_frame.pack(fill="x", padx=10, pady=2)
        
        ctk.CTkLabel(text_scale_frame, text="Text Scale (%):", width=80, anchor="w").pack(side="left")
        text_scale_entry = ctk.CTkEntry(text_scale_frame, textvariable=self.text_scale, width=60)
        text_scale_entry.pack(side="left", padx=(5, 5))
        text_scale_entry.bind("<KeyRelease>", self.on_position_change)
        
        text_scale_slider = ctk.CTkSlider(
            text_scale_frame,
            from_=25, to=300,
            variable=self.text_scale,
            command=self.on_text_scale_change,
            width=100
        )
        text_scale_slider.pack(side="left", padx=(5, 0))
    
    def setup_output_settings(self):
        """Setup output settings section"""
        output_frame = ctk.CTkFrame(self.scrollable_frame, fg_color="#1E1E1E", corner_radius=12)
        output_frame.pack(pady=(5, 5), padx=20, fill="x")
        
        ctk.CTkLabel(output_frame, text="💾 Output Settings", font=("Segoe UI", 14, "bold")).pack(pady=(8, 5), padx=20, fill="x")
        
        # Output folder selection
        out_frame = ctk.CTkFrame(output_frame, fg_color="transparent")
        out_frame.pack(pady=(3, 8), padx=20, fill="x")
        ctk.CTkLabel(out_frame, text="Output Folder:", width=120, anchor="w").pack(side="left", padx=(0, 10))
        self.output_path_var = tk.StringVar(value="No folder selected")
        self.output_path_label = ctk.CTkLabel(out_frame, textvariable=self.output_path_var, anchor="w")
        self.output_path_label.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(out_frame, text="Browse", command=self.select_output_folder, width=80).pack(side="right")
    
    def setup_generation_controls(self):
        """Setup generation control buttons"""
        generate_frame = ctk.CTkFrame(self.scrollable_frame, fg_color="#1E1E1E", corner_radius=12)
        generate_frame.pack(pady=(5, 5), padx=20, fill="x")
        
        # Button grid for Generate
        btn_grid = ctk.CTkFrame(generate_frame, fg_color="transparent")
        btn_grid.pack(pady=(8, 8), padx=20, fill="x")
        
        self.generate_btn = ctk.CTkButton(btn_grid, text="Generate Gift Cards", command=self.threaded_generate, height=40)
        self.generate_btn.pack(fill="x", expand=True)
        
        self.generate_loading = ctk.CTkLabel(generate_frame, text="", font=("Arial", 12), text_color="#00BFFF")
        self.generate_loading.pack(pady=(3, 8), padx=20, fill="x")
    
    def setup_log_section(self):
        """Setup log section"""
        log_frame = ctk.CTkFrame(self.scrollable_frame, fg_color="#1E1E1E", corner_radius=12)
        log_frame.pack(pady=(5, 10), padx=20, fill="both", expand=True)
        
        log_header = ctk.CTkFrame(log_frame, fg_color="transparent")
        log_header.pack(fill="x", padx=20, pady=(8, 5))
        
        ctk.CTkLabel(log_header, text="📝 Generation Log", font=("Segoe UI", 14, "bold")).pack(side="left")
        
        # Add toggle button for log visibility
        self.log_visible = True
        self.toggle_log_btn = ctk.CTkButton(log_header, text="Hide Log", command=self.toggle_log_visibility, width=80, height=25)
        self.toggle_log_btn.pack(side="right", padx=(5, 0))
        
        clear_log_btn = ctk.CTkButton(log_header, text="Clear Log", command=self.clear_log, width=80, height=25)
        clear_log_btn.pack(side="right")
        
        # Log container with minimum height
        self.log_container = ctk.CTkFrame(log_frame, fg_color="transparent")
        self.log_container.pack(fill="both", expand=True, padx=20, pady=(3, 8))
        
        self.log_box = ctk.CTkTextbox(self.log_container, state="disabled", fg_color="#1E1E1E", text_color="#CCCCCC", wrap="word", height=250)
        self.log_box.pack(fill="both", expand=True)
        
        # Add some bottom padding to ensure content is always visible
        bottom_spacer = ctk.CTkFrame(self.scrollable_frame, height=20, fg_color="transparent")
        bottom_spacer.pack(fill="x")
    
    def setup_event_bindings(self):
        """Setup event bindings for live preview updates"""
        # Bind events for live preview updates
        self.barcode_position_var.trace_add('write', self.update_live_preview)
        self.barcode_size_var.trace_add('write', self.update_live_preview)
        self.barcode_x.trace_add('write', self.update_live_preview)
        self.barcode_y.trace_add('write', self.update_live_preview)
        self.text_position_var.trace_add('write', self.update_live_preview)
        self.text_background_var.trace_add('write', self.update_live_preview)
        self.text_transparency.trace_add('write', self.update_live_preview)
        self.text_alignment_var.trace_add('write', self.update_live_preview)
        self.text_scale.trace_add('write', self.update_live_preview)
        self.text_x.trace_add('write', self.update_live_preview)
        self.text_y.trace_add('write', self.update_live_preview)
        
        # Initialize transparency control visibility
        self.on_position_change()
    
    def log(self, msg):
        """Log a message"""
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
    
    def select_background(self):
        """Select background image file"""
        file_path = filedialog.askopenfilename(
            title="Select Background Image",
            filetypes=[("Image files", "*.png *.jpg *.jpeg"), ("All files", "*.*")]
        )
        if file_path:
            self.background_path = file_path
            self.bg_path_var.set(os.path.basename(file_path))
            self.log(f"✅ Background image selected: {os.path.basename(file_path)}")
            # Update live preview with new background
            self.refresh_preview_canvas()
    
    def select_data_file(self):
        """Select data file (CSV or Excel)"""
        file_path = filedialog.askopenfilename(
            title="Select Data File",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.data_path = file_path
            self.data_path_var.set(os.path.basename(file_path))
            self.log(f"✅ Data file selected: {os.path.basename(file_path)}")
            # Update live preview with new data
            self.refresh_preview_canvas()
    
    def select_output_folder(self):
        """Select output folder"""
        folder_path = filedialog.askdirectory(title="Select Output Folder")
        if folder_path:
            self.output_path = folder_path
            self.output_path_var.set(os.path.basename(folder_path))
            self.log(f"✅ Output folder selected: {os.path.basename(folder_path)}")
    
    def threaded_generate(self):
        """Generate gift cards in a separate thread"""
        threading.Thread(target=self.generate_gift_cards, daemon=True).start()
    
    def on_column_config_change(self, event=None):
        """Handle column configuration changes"""
        self.refresh_preview_canvas()
    
    def on_barcode_position_change(self, value=None):
        """Handle barcode position change"""
        if self.barcode_position_var.get() == "Custom":
            self.barcode_custom_frame.pack(fill="x", padx=10, pady=2)
        else:
            self.barcode_custom_frame.pack_forget()
            # Set preset positions
            positions = {
                "Top-Left": ("10", "10"),
                "Top-Right": ("90", "10"),
                "Bottom-Left": ("10", "90"),
                "Bottom-Right": ("90", "90"),
                "Center": ("50", "50")
            }
            pos = self.barcode_position_var.get()
            if pos in positions:
                x, y = positions[pos]
                self.barcode_x.set(x)
                self.barcode_y.set(y)
        self.update_live_preview()
    
    def on_text_position_change(self, value=None):
        """Handle text position change"""
        if self.text_position_var.get() == "Custom":
            self.text_custom_frame.pack(fill="x", padx=10, pady=2)
        else:
            self.text_custom_frame.pack_forget()
            # Set preset positions
            positions = {
                "Top-Left": ("10", "10"),
                "Top-Right": ("90", "10"),
                "Bottom-Left": ("10", "90"),
                "Bottom-Right": ("90", "90"),
                "Center": ("50", "50")
            }
            pos = self.text_position_var.get()
            if pos in positions:
                x, y = positions[pos]
                self.text_x.set(x)
                self.text_y.set(y)
        self.update_live_preview()
    
    def on_background_change(self, value=None):
        """Handle background change"""
        if self.text_background_var.get() == "Semi-transparent":
            self.transparency_frame.pack(fill="x", padx=10, pady=2)
        else:
            self.transparency_frame.pack_forget()
            
        if self.text_background_var.get() == "Custom Color":
            self.custom_color_frame.pack(fill="x", padx=10, pady=2)
        else:
            self.custom_color_frame.pack_forget()
        
        self.update_live_preview()
    
    def on_transparency_change(self, value=None):
        """Handle transparency change"""
        self.update_live_preview()
    
    def on_custom_color_change(self, event=None):
        """Handle custom color change"""
        color = self.custom_color_entry.get().strip()
        if color.startswith('#') and len(color) == 7:
            try:
                # Validate color
                self.color_preview_btn.configure(fg_color=color)
                self.custom_bg_color = color
                self.update_live_preview()
            except:
                pass
    
    def open_color_picker(self):
        """Open color picker dialog"""
        color = colorchooser.askcolor(color=self.custom_bg_color)
        if color[1]:  # If user didn't cancel
            self.custom_bg_color = color[1]
            self.custom_color_entry.delete(0, tk.END)
            self.custom_color_entry.insert(0, self.custom_bg_color)
            self.color_preview_btn.configure(fg_color=self.custom_bg_color)
            self.update_live_preview()
    
    def on_position_change(self, event=None):
        """Handle position changes"""
        self.update_live_preview()
    
    def on_text_scale_change(self, value=None):
        """Handle text scale changes"""
        self.update_live_preview()
    
    def reset_positions(self):
        """Reset all positions to default"""
        self.barcode_position_var.set("Bottom-Right")
        self.barcode_x.set("85")
        self.barcode_y.set("85")
        self.barcode_size_var.set("Medium")
        
        self.text_position_var.set("Bottom-Left")
        self.text_x.set("15")
        self.text_y.set("85")
        self.text_background_var.set("White Box")
        self.text_transparency.set(50.0)
        self.text_alignment_var.set("Left")
        self.text_scale.set(100.0)
        
        self.update_live_preview()
    
    def update_live_preview(self, *args):
        """Update live preview"""
        self.refresh_preview_canvas()
    
    def refresh_preview_canvas(self):
        """Refresh the preview canvas with current settings"""
        if not self.background_path:
            return
        
        try:
            # Create sample data for preview
            sample_data = {
                "barcode": "1234567890123",
                "member_number": "12345",
                "verification_code": "ABCD1234"
            }
            
            # Generate preview image
            preview_image = self.create_gift_card_image(
                self.background_path,
                sample_data["barcode"],
                sample_data["member_number"],
                sample_data["verification_code"]
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
                
        except Exception as e:
            self.log(f"❌ Preview error: {str(e)}")
    
    def on_canvas_click(self, event):
        """Handle canvas click for drag and drop"""
        # Placeholder for drag and drop functionality
        pass
    
    def on_canvas_drag(self, event):
        """Handle canvas drag"""
        # Placeholder for drag and drop functionality
        pass
    
    def on_canvas_release(self, event):
        """Handle canvas release"""
        # Placeholder for drag and drop functionality
        pass
    
    def toggle_log_visibility(self):
        """Toggle log visibility"""
        if self.log_visible:
            self.log_container.pack_forget()
            self.toggle_log_btn.configure(text="Show Log")
            self.log_visible = False
        else:
            self.log_container.pack(fill="both", expand=True, padx=20, pady=(3, 8))
            self.toggle_log_btn.configure(text="Hide Log")
            self.log_visible = True
    
    def clear_log(self):
        """Clear the log"""
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", tk.END)
        self.log_box.configure(state="disabled")
    
    def get_actual_barcode_size(self):
        """Get actual barcode size based on selection"""
        size_mapping = {
            "Small": {"width": 120, "height": 30},
            "Medium": {"width": 160, "height": 40},
            "Large": {"width": 200, "height": 50},
            "XL": {"width": 240, "height": 60}
        }
        return size_mapping.get(self.barcode_size_var.get(), {"width": 160, "height": 40})
    
    def generate_barcode(self, barcode_data):
        """Generate POS scanner-compatible Code128 barcode with optimal settings"""
        try:
            # Compact POS-compatible settings for label-style layout
            writer = ImageWriter()
            writer.set_options({
                'module_width': 0.5,     # Wider bars for easier scanning
                'module_height': 25.0,   # Taller bars for better accuracy
                'quiet_zone': 6.5,       # Reduced quiet zone for compact label
                'font_size': 0,          # Hide barcode string
                'write_text': False,     # Do not display raw string under barcode
                'dpi': 300               # High resolution for print quality
            })
            
            # Generate Code128 barcode - ensure barcode_data includes proper format
            # If barcode_data doesn't have semicolon/question mark, format it properly
            formatted_data = str(barcode_data)
            if not formatted_data.startswith(';'):
                formatted_data = f';{formatted_data}?'
            elif not formatted_data.endswith('?'):
                formatted_data = f'{formatted_data}?'
                
            code = Code128(formatted_data, writer=writer)
            buffer = BytesIO()
            code.write(buffer, text='')  # Explicitly pass empty text to ensure no string displays
            buffer.seek(0)
            barcode_img = Image.open(buffer)
            
            # Scale to user-selected size from dropdown
            width, height = barcode_img.size
            
            # Get target dimensions from user's size selection
            barcode_size = self.get_actual_barcode_size()
            target_width = barcode_size['width']
            target_height = barcode_size['height']
            
            # Calculate scale to fit target dimensions
            scale_x = target_width / width
            scale_y = target_height / height
            scale = min(scale_x, scale_y)  # Maintain aspect ratio
            
            new_width = int(width * scale)
            new_height = int(height * scale)
            
            barcode_img = barcode_img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            return barcode_img
            
        except Exception as e:
            self.log(f"❌ Barcode generation error: {str(e)}")
            return None
    
    def create_gift_card_image(self, background_path, barcode_data, member_number, verification_code):
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
            self.draw_text_block_full(background, member_number, verification_code)
            
            return background.convert("RGB")
            
        except Exception as e:
            self.log(f"❌ Gift card creation error: {str(e)}")
            return None
    
    def draw_text_block_full(self, image, member_number, verification_code):
        """Draw text block on the image with full functionality"""
        try:
            draw = ImageDraw.Draw(image)
            
            # Prepare text
            text_lines = [
                f"Member: {member_number}",
                f"PIN: {verification_code}"
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
                try:
                    font = ImageFont.load_default()
                except:
                    font = None
            
            if not font:
                return
            
            # Calculate text dimensions
            text_heights = []
            text_widths = []
            for line in text_lines:
                bbox = draw.textbbox((0, 0), line, font=font)
                text_widths.append(bbox[2] - bbox[0])
                text_heights.append(bbox[3] - bbox[1])
            
            max_text_width = max(text_widths)
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
                elif background_type == "Semi-transparent":
                    # Create semi-transparent overlay
                    overlay = Image.new('RGBA', image.size, (255, 255, 255, int(255 * (100 - self.text_transparency.get()) / 100)))
                    mask = Image.new('RGBA', image.size, (0, 0, 0, 0))
                    mask_draw = ImageDraw.Draw(mask)
                    mask_draw.rectangle([bg_x1, bg_y1, bg_x2, bg_y2], fill=(255, 255, 255, 255))
                    image.paste(overlay, (0, 0), mask)
                elif background_type == "Custom Color":
                    try:
                        # Parse hex color
                        color = self.custom_bg_color
                        if color.startswith('#'):
                            color = color[1:]
                        r = int(color[0:2], 16)
                        g = int(color[2:4], 16)
                        b = int(color[4:6], 16)
                        draw.rectangle([bg_x1, bg_y1, bg_x2, bg_y2], fill=(r, g, b, 255))
                    except:
                        draw.rectangle([bg_x1, bg_y1, bg_x2, bg_y2], fill=(224, 224, 224, 255))
            
            # Draw text lines
            current_y = text_y
            text_color = (0, 0, 0) if background_type == "White Box" else (255, 255, 255)
            
            for i, line in enumerate(text_lines):
                if self.text_alignment_var.get() == "Center":
                    line_x = text_x + (max_text_width - text_widths[i]) // 2
                elif self.text_alignment_var.get() == "Right":
                    line_x = text_x + (max_text_width - text_widths[i])
                else:
                    line_x = text_x
                
                draw.text((line_x, current_y), line, font=font, fill=text_color)
                current_y += text_heights[i] + 5
                
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
            self.generate_loading.configure(text="🔄 Reading data file...")
            
            # Read data file
            if self.data_path.endswith('.csv'):
                df = pd.read_csv(self.data_path)
            else:
                df = pd.read_excel(self.data_path)
            
            # Get column names
            barcode_col = safe_get_input(self.barcode_col, "barcode")
            member_col = safe_get_input(self.member_col, "member_number")
            verification_col = safe_get_input(self.verification_col, "pin")
            
            # Validate columns exist
            missing_cols = []
            for col_name, col_entry in [(barcode_col, "Barcode"), (member_col, "Member Number"), (verification_col, "Verification Code")]:
                if col_name not in df.columns:
                    missing_cols.append(f"{col_entry} ({col_name})")
            
            if missing_cols:
                self.log(f"❌ Missing columns: {', '.join(missing_cols)}")
                self.generate_loading.configure(text="")
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
                    
                    # Generate gift card
                    card_image = self.create_gift_card_image(
                        self.background_path,
                        barcode_data,
                        member_number,
                        verification_code
                    )
                    
                    if card_image:
                        # Save image
                        filename = f"gift_card_{member_number}_{index+1:04d}.png"
                        output_file = os.path.join(self.output_path, filename)
                        card_image.save(output_file, "PNG", quality=95)
                        success_count += 1
                        
                        # Update progress
                        progress = f"Progress: {success_count}/{len(df)} cards generated"
                        self.generate_loading.configure(text=progress)
                    
                except Exception as e:
                    self.log(f"❌ Error generating card {index+1}: {str(e)}")
                    continue
            
            self.generate_loading.configure(text="")
            self.log(f"✅ Generated {success_count}/{len(df)} gift cards successfully!")
            
            if success_count > 0:
                show_toast(self.scrollable_frame, f"🎉 {success_count} gift cards generated!", 5000, "#00FF00")
            
        except Exception as e:
            self.generate_loading.configure(text="")
            self.log(f"❌ Generation error: {str(e)}")

if __name__ == "__main__":
    app = GiftCardGenerator()
    app.mainloop()