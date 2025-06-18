#!/usr/bin/env python3
"""
Work Order PDF Extractor
A GUI application to process work order PDFs, extract information using OpenAI API,
and organize files based on CSV reference data.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import shutil
import csv
import json
from pathlib import Path
import threading
from typing import List, Dict, Tuple, Optional
import logging
from datetime import datetime
import asyncio
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor, as_completed
import time

# PDF and image processing
from pdf2image import convert_from_path
from PIL import Image, ImageTk

# PyMuPDF (optional fallback)
try:
    import fitz
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False
    print("PyMuPDF not available - using pdf2image only")

# OpenAI API
import openai

class WorkOrderExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("Work Order PDF Extractor")
        # Modern responsive sizing - don't fill whole screen
        width = 1200
        height = 850
        self.root.geometry(f"{width}x{height}")
        self.root.minsize(1000, 700)
        
        # Center window on screen
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
        # Configure root grid weights for better responsiveness
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        # Configuration
        self.config = {
            'openai_api_key': '',
            'selected_model': 'gpt-4.1-nano',  # Default model
            'crop_x1': 0,
            'crop_y1': 0,
            'crop_x2': 0.25,  # 1/4 of width
            'crop_y2': 0.25,  # 1/4 of height
            'pdf_folder': 'workOrderPDF',
            'ref_csv_file': 'workOrderRef/MCAN_work_inprogress.csv',
            'not_match_folder': 'not_match'
        }
        
        # Model pricing (USD per million tokens) - 2025 rates
        self.model_pricing = {
            'gpt-4.1-nano': {
                'input': 0.10,    # $0.10 per 1M input tokens
                'output': 0.40,   # $0.40 per 1M output tokens
                'description': 'Fastest and cheapest model with 1M context'
            },
            'gpt-4.1-mini': {
                'input': 0.40,    # $0.40 per 1M input tokens
                'output': 1.60,   # $1.60 per 1M output tokens
                'description': 'Balanced performance and cost with 1M context'
            },
            'gpt-4.1': {
                'input': 3.00,    # $3.00 per 1M input tokens (estimated)
                'output': 12.00,  # $12.00 per 1M output tokens (estimated)
                'description': 'Most capable model with 1M context'
            }
        }
        
        # Currency conversion (33 THB = 1 USD)
        self.usd_to_thb = 33.0
        
        # Performance settings
        self.max_concurrent_workers = 5  # Limit concurrent API calls
        self.pdf_processing_workers = 8  # PDF conversion workers
        self.batch_size = 10  # Files per batch for progress updates
        
        # Token and cost tracking
        self.session_stats = {
            'total_input_tokens': 0,
            'total_output_tokens': 0,
            'total_cost_usd': 0.0,
            'total_cost_thb': 0.0,
            'files_processed': 0,
            'api_calls': 0,
            'successful_files': 0,
            'failed_files': 0
        }
        
        # Create directories if they don't exist
        self.ensure_directories()
        
        # Setup logging
        self.setup_logging()
        
        # Load reference data
        self.reference_orders = self.load_reference_data()
        
        # Create GUI
        self.create_widgets()
        
        # Load saved settings
        self.load_settings()
        
        # Initialize manual crop file list after settings are loaded
        if hasattr(self, 'manual_crop_file_var'):
            self.root.after(100, self.refresh_pdf_list)  # Delay to ensure UI is ready
        
    def ensure_directories(self):
        """Create necessary directories if they don't exist"""
        directories = [
            self.config['pdf_folder'],
            self.config['not_match_folder'],
            os.path.dirname(self.config['ref_csv_file'])
        ]
        for directory in directories:
            os.makedirs(directory, exist_ok=True)
    
    def setup_logging(self):
        """Setup logging configuration"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('work_order_extractor.log'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def create_widgets(self):
        """Create modern UI following Refactoring UI principles"""
        # Modern design system
        style = ttk.Style()
        style.theme_use('clam')
        
        # Systematic color palette
        self.colors = {
            'primary': '#6366F1',        # Modern indigo
            'success': '#10B981',        # Modern green
            'error': '#EF4444',          # Modern red
            'warning': '#F59E0B',        # Modern amber
            'info': '#3B82F6',           # Modern blue
            'neutral_900': '#111827',    # Almost black
            'neutral_700': '#374151',    # Dark grey
            'neutral_600': '#4B5563',    # Medium grey
            'neutral_500': '#6B7280',    # Light grey
            'neutral_300': '#D1D5DB',    # Very light
            'neutral_100': '#F3F4F6',    # Background
            'surface': '#FFFFFF',        # Pure white
            'surface_secondary': '#F8FAFC' # Slight tint
        }
        
        # Typography system
        self.fonts = {
            'display': ('Inter', 24, 'bold'),
            'title': ('Inter', 18, 'bold'),
            'heading': ('Inter', 14, 'bold'),
            'body': ('Inter', 13, 'normal'),
            'small': ('Inter', 11, 'normal'),
            'mono': ('JetBrains Mono', 12, 'normal')
        }
        
        # Spacing system (8px base)
        self.spacing = {'xs': 4, 'sm': 8, 'md': 16, 'lg': 24, 'xl': 32}
        
        # Configure systematic styles
        style.configure('Display.TLabel', font=self.fonts['display'], foreground=self.colors['neutral_900'])
        style.configure('Title.TLabel', font=self.fonts['title'], foreground=self.colors['neutral_700'])
        style.configure('Heading.TLabel', font=self.fonts['heading'], foreground=self.colors['neutral_600'])
        style.configure('Body.TLabel', font=self.fonts['body'], foreground=self.colors['neutral_600'])
        style.configure('Small.TLabel', font=self.fonts['small'], foreground=self.colors['neutral_500'])
        
        style.configure('Success.TLabel', font=self.fonts['body'], foreground=self.colors['success'])
        style.configure('Error.TLabel', font=self.fonts['body'], foreground=self.colors['error'])
        style.configure('Warning.TLabel', font=self.fonts['body'], foreground=self.colors['warning'])
        style.configure('Primary.TLabel', font=self.fonts['body'], foreground=self.colors['primary'])
        
        # Card and surface styles
        style.configure('Card.TFrame', background=self.colors['surface'], relief='flat')
        style.configure('Surface.TFrame', background=self.colors['surface_secondary'])
        style.configure('Accent.TFrame', background=self.colors['primary'])
        
        # Main container with breathing room
        main_container = ttk.Frame(self.root, style='Surface.TFrame')
        main_container.pack(fill=tk.BOTH, expand=True, padx=self.spacing['xl'], pady=self.spacing['xl'])
        
        # Modern app header with visual hierarchy
        self.create_app_header(main_container)
        
        # Modern notebook
        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Manual crop variables - initialize before creating tabs
        self.crop_window = None
        self.crop_canvas = None
        self.crop_image = None
        self.selection_rect = None
        self.start_x = 0
        self.start_y = 0
        self.current_x = 0
        self.current_y = 0
        self.crop_preview_image = None
        self.crop_preview_label = None
        self.crop_coords_text = None
        self.pdf_canvas = None
        self.manual_crop_file_var = None
        
        # Create modern tabs
        self.create_settings_tab()
        self.create_manual_crop_tab()
        self.create_processing_tab()
        self.create_log_tab()
    
    def create_app_header(self, parent):
        """Create modern app header with proper hierarchy"""
        header_card = ttk.Frame(parent, style='Card.TFrame')
        header_card.pack(fill=tk.X, pady=(0, self.spacing['lg']))
        
        # Accent border at top
        accent_border = ttk.Frame(header_card, style='Accent.TFrame', height=4)
        accent_border.pack(fill=tk.X)
        
        header_content = ttk.Frame(header_card)
        header_content.pack(fill=tk.X, padx=self.spacing['lg'], pady=self.spacing['lg'])
        
        # Title with visual hierarchy
        ttk.Label(header_content, text="Work Order PDF Extractor", style='Display.TLabel').pack(anchor=tk.W)
        ttk.Label(header_content, text="AI-powered document processing with OpenAI", 
                 style='Body.TLabel').pack(anchor=tk.W, pady=(self.spacing['xs'], 0))
    
    def create_settings_tab(self):
        """Create modern settings tab with card-based design"""
        settings_container = ttk.Frame(self.notebook, style='Surface.TFrame')
        self.notebook.add(settings_container, text="⚙️ Settings")
        
        # Scrollable content container
        canvas = tk.Canvas(settings_container, bg=self.colors['surface_secondary'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(settings_container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style='Surface.TFrame')
        
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True, padx=self.spacing['lg'], pady=self.spacing['lg'])
        scrollbar.pack(side="right", fill="y")
        
        # Content with max width constraint
        settings_frame = ttk.Frame(scrollable_frame, style='Surface.TFrame')
        settings_frame.pack(fill=tk.X, padx=(0, self.spacing['lg']))
        
        # API Configuration Card
        api_card = ttk.Frame(settings_frame, style='Card.TFrame')
        api_card.pack(fill=tk.X, pady=(0, self.spacing['lg']))
        
        # Accent border
        accent = ttk.Frame(api_card, style='Accent.TFrame', height=3)
        accent.pack(fill=tk.X)
        
        api_content = ttk.Frame(api_card)
        api_content.pack(fill=tk.X, padx=self.spacing['lg'], pady=self.spacing['lg'])
        
        ttk.Label(api_content, text="🔐 OpenAI API Configuration", style='Title.TLabel').pack(anchor=tk.W)
        ttk.Label(api_content, text="Configure your OpenAI API credentials and model preferences", 
                 style='Body.TLabel').pack(anchor=tk.W, pady=(self.spacing['xs'], self.spacing['md']))
        
        # API Key input
        key_frame = ttk.Frame(api_content)
        key_frame.pack(fill=tk.X, pady=(0, self.spacing['md']))
        
        ttk.Label(key_frame, text="API Key", style='Heading.TLabel').pack(anchor=tk.W, pady=(0, self.spacing['xs']))
        self.api_key_var = tk.StringVar()
        api_entry = ttk.Entry(key_frame, textvariable=self.api_key_var, show="*", 
                             width=60, font=self.fonts['mono'])
        api_entry.pack(fill=tk.X, ipady=self.spacing['xs'])
        
        # Model selection
        model_frame = ttk.Frame(api_content)
        model_frame.pack(fill=tk.X)
        
        ttk.Label(model_frame, text="Model Selection", style='Heading.TLabel').pack(anchor=tk.W, pady=(0, self.spacing['xs']))
        
        self.model_var = tk.StringVar(value=self.config['selected_model'])
        self.model_combo = ttk.Combobox(model_frame, textvariable=self.model_var,
                                       values=list(self.model_pricing.keys()),
                                       state="readonly", width=30, font=self.fonts['body'])
        self.model_combo.pack(anchor=tk.W, ipady=self.spacing['xs'])
        self.model_combo.bind('<<ComboboxSelected>>', self.on_model_changed)
        
        # Model description
        self.model_desc_var = tk.StringVar()
        self.update_model_description()
        ttk.Label(model_frame, textvariable=self.model_desc_var, style='Small.TLabel').pack(anchor=tk.W, pady=(self.spacing['xs'], 0))
        
        # Cost Tracking Card
        cost_card = ttk.Frame(settings_frame, style='Card.TFrame')
        cost_card.pack(fill=tk.X, pady=(0, self.spacing['lg']))
        
        cost_content = ttk.Frame(cost_card)
        cost_content.pack(fill=tk.X, padx=self.spacing['lg'], pady=self.spacing['lg'])
        
        # Header with reset button
        header_frame = ttk.Frame(cost_content)
        header_frame.pack(fill=tk.X, pady=(0, self.spacing['md']))
        
        ttk.Label(header_frame, text="📊 Session Cost Tracking", style='Title.TLabel').pack(side=tk.LEFT)
        reset_btn = ttk.Button(header_frame, text="🔄 Reset", command=self.reset_session_stats, width=12)
        reset_btn.pack(side=tk.RIGHT)
        
        # Modern metrics grid
        metrics_grid = ttk.Frame(cost_content)
        metrics_grid.pack(fill=tk.X)
        
        for i in range(3):
            metrics_grid.grid_columnconfigure(i, weight=1)
        
        # Create metric cards
        self.create_metric_card(metrics_grid, "API Calls", "api_calls_var", "0", self.colors['info'], 0, 0)
        self.create_metric_card(metrics_grid, "Input Tokens", "input_tokens_var", "0", self.colors['success'], 0, 1)
        self.create_metric_card(metrics_grid, "Output Tokens", "output_tokens_var", "0", self.colors['warning'], 0, 2)
        
        self.create_metric_card(metrics_grid, "Cost (USD)", "cost_usd_var", "$0.00", self.colors['error'], 1, 0)
        self.create_metric_card(metrics_grid, "Cost (THB)", "cost_thb_var", "฿0.00", self.colors['primary'], 1, 1)
        
        # Crop Coordinates Card
        crop_card = ttk.Frame(settings_frame, style='Card.TFrame')
        crop_card.pack(fill=tk.X, pady=(0, self.spacing['lg']))
        
        crop_content = ttk.Frame(crop_card)
        crop_content.pack(fill=tk.X, padx=self.spacing['lg'], pady=self.spacing['lg'])
        
        ttk.Label(crop_content, text="✂️ Crop Coordinates", style='Title.TLabel').pack(anchor=tk.W)
        ttk.Label(crop_content, text="Define the area to extract from PDFs (as fraction of page size)", 
                 style='Body.TLabel').pack(anchor=tk.W, pady=(self.spacing['xs'], self.spacing['md']))
        
        # Coordinates grid
        coords_grid = ttk.Frame(crop_content)
        coords_grid.pack(fill=tk.X, pady=(0, self.spacing['md']))
        
        for i in range(4):
            coords_grid.grid_columnconfigure(i, weight=1)
        
        # Coordinate inputs
        self.create_coord_input(coords_grid, "X1 (Left)", "x1_var", self.config['crop_x1'], 0, 0)
        self.create_coord_input(coords_grid, "Y1 (Top)", "y1_var", self.config['crop_y1'], 0, 1)
        self.create_coord_input(coords_grid, "X2 (Right)", "x2_var", self.config['crop_x2'], 0, 2)
        self.create_coord_input(coords_grid, "Y2 (Bottom)", "y2_var", self.config['crop_y2'], 0, 3)
        
        # Reset button
        ttk.Button(crop_content, text="🔄 Reset to Default", command=self.reset_crop_default, width=20).pack()
        
        # File Paths Card
        paths_card = ttk.Frame(settings_frame, style='Card.TFrame')
        paths_card.pack(fill=tk.X, pady=(0, self.spacing['lg']))
        
        paths_content = ttk.Frame(paths_card)
        paths_content.pack(fill=tk.X, padx=self.spacing['lg'], pady=self.spacing['lg'])
        
        ttk.Label(paths_content, text="📁 File Paths", style='Title.TLabel').pack(anchor=tk.W)
        ttk.Label(paths_content, text="Configure input and output directories", 
                 style='Body.TLabel').pack(anchor=tk.W, pady=(self.spacing['xs'], self.spacing['md']))
        
        # PDF folder
        self.create_path_input(paths_content, "PDF Folder", "pdf_folder_var", 
                              self.config['pdf_folder'], self.browse_pdf_folder)
        
        # CSV file
        self.create_path_input(paths_content, "Reference CSV File", "csv_file_var", 
                              self.config['ref_csv_file'], self.browse_csv_file)
        
        # Save button
        save_frame = ttk.Frame(paths_content)
        save_frame.pack(fill=tk.X, pady=(self.spacing['md'], 0))
        
        ttk.Button(save_frame, text="💾 Save Settings", command=self.save_settings, width=20).pack()
        
    
    def create_metric_card(self, parent, label, var_name, initial_value, color, row, col):
        """Create individual metric card"""
        metric_frame = ttk.Frame(parent, style='Surface.TFrame')
        metric_frame.grid(row=row, column=col, padx=self.spacing['xs'], pady=self.spacing['xs'], sticky="ew")
        
        metric_content = ttk.Frame(metric_frame)
        metric_content.pack(padx=self.spacing['sm'], pady=self.spacing['sm'])
        
        ttk.Label(metric_content, text=label, style='Small.TLabel').pack()
        
        setattr(self, var_name, tk.StringVar(value=initial_value))
        value_label = ttk.Label(metric_content, textvariable=getattr(self, var_name), style='Heading.TLabel')
        value_label.configure(foreground=color)
        value_label.pack(pady=(self.spacing['xs'], 0))
    
    def create_coord_input(self, parent, label, var_name, initial_value, row, col):
        """Create coordinate input field"""
        coord_frame = ttk.Frame(parent)
        coord_frame.grid(row=row, column=col, padx=self.spacing['xs'], sticky="ew")
        
        ttk.Label(coord_frame, text=label, style='Small.TLabel').pack(anchor=tk.W)
        
        setattr(self, var_name, tk.DoubleVar(value=initial_value))
        entry = ttk.Entry(coord_frame, textvariable=getattr(self, var_name), 
                         width=12, font=self.fonts['mono'])
        entry.pack(fill=tk.X, ipady=self.spacing['xs'], pady=(self.spacing['xs'], 0))
    
    def create_path_input(self, parent, label, var_name, initial_value, browse_command):
        """Create path input with browse button"""
        path_frame = ttk.Frame(parent)
        path_frame.pack(fill=tk.X, pady=(0, self.spacing['md']))
        
        ttk.Label(path_frame, text=label, style='Heading.TLabel').pack(anchor=tk.W, pady=(0, self.spacing['xs']))
        
        input_frame = ttk.Frame(path_frame)
        input_frame.pack(fill=tk.X)
        
        setattr(self, var_name, tk.StringVar(value=initial_value))
        entry = ttk.Entry(input_frame, textvariable=getattr(self, var_name), 
                         font=self.fonts['mono'])
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=self.spacing['xs'])
        
        ttk.Button(input_frame, text="📂 Browse", command=browse_command, width=12).pack(side=tk.RIGHT, padx=(self.spacing['xs'], 0))
    
    def create_processing_tab(self):
        """Create the main processing tab"""
        process_frame = ttk.Frame(self.notebook)
        self.notebook.add(process_frame, text="Process PDFs")
        
        # Enhanced status section
        status_frame = ttk.LabelFrame(process_frame, text="📊 Processing Status", padding=15)
        status_frame.pack(fill=tk.X, padx=15, pady=10)
        
        self.status_var = tk.StringVar(value="Ready to process")
        ttk.Label(status_frame, textvariable=self.status_var).pack(anchor=tk.W)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var, 
                                          maximum=100, length=300)
        self.progress_bar.pack(fill=tk.X, pady=5)
        
        # Current session cost display with better contrast
        cost_summary_frame = ttk.Frame(status_frame)
        cost_summary_frame.pack(fill=tk.X, pady=5)
        
        self.session_cost_var = tk.StringVar(value="Session Cost: $0.00 USD (฿0.00 THB)")
        ttk.Label(cost_summary_frame, textvariable=self.session_cost_var, 
                 font=("Arial", 11, "bold"), foreground="#FFFFFF", background="#2C3E50").pack(anchor=tk.W, pady=2)
        
        # Add background frame for better contrast
        session_cost_bg = ttk.Frame(cost_summary_frame, style='Dark.TFrame')
        session_cost_bg.pack(fill=tk.X, pady=2)
        self.session_cost_label = ttk.Label(session_cost_bg, textvariable=self.session_cost_var, 
                 font=("Arial", 11, "bold"), foreground="#FFFFFF")
        self.session_cost_label.pack(padx=10, pady=5)
        
        # Add Success/Fail results display
        results_frame = ttk.Frame(status_frame)
        results_frame.pack(fill=tk.X, pady=5)
        
        # Results display with better styling
        results_grid = ttk.Frame(results_frame)
        results_grid.pack(fill=tk.X)
        
        ttk.Label(results_grid, text="Results:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky=tk.W, padx=5)
        
        ttk.Label(results_grid, text="Success:", font=("Arial", 9)).grid(row=0, column=1, sticky=tk.W, padx=10)
        self.success_count_var = tk.StringVar(value="0")
        ttk.Label(results_grid, textvariable=self.success_count_var, 
                 foreground="#27AE60", font=("Arial", 9, "bold")).grid(row=0, column=2, sticky=tk.W, padx=5)
        
        ttk.Label(results_grid, text="Failed:", font=("Arial", 9)).grid(row=0, column=3, sticky=tk.W, padx=10)
        self.fail_count_var = tk.StringVar(value="0")
        ttk.Label(results_grid, textvariable=self.fail_count_var, 
                 foreground="#E74C3C", font=("Arial", 9, "bold")).grid(row=0, column=4, sticky=tk.W, padx=5)
        
        # Enhanced control buttons with improved styling
        button_frame = ttk.Frame(process_frame)
        button_frame.pack(fill=tk.X, padx=15, pady=20)
        
        # Create button grid for better layout on large screens
        button_grid = ttk.Frame(button_frame)
        button_grid.pack(fill=tk.X)
        
        for i in range(4):
            button_grid.grid_columnconfigure(i, weight=1)
        
        self.process_button = ttk.Button(button_grid, text="▶️ Start Processing", 
                                       command=self.start_processing, width=20)
        self.process_button.grid(row=0, column=0, padx=10, pady=8, sticky=tk.EW)
        
        self.stop_button = ttk.Button(button_grid, text="⏹️ Stop Processing", 
                                    command=self.stop_processing, state=tk.DISABLED, width=20)
        self.stop_button.grid(row=0, column=1, padx=10, pady=8, sticky=tk.EW)
        
        # Open PDF Folder button
        self.open_folder_button = ttk.Button(button_grid, text="📂 Open PDF Folder", 
                                           command=self.open_pdf_folder, width=20)
        self.open_folder_button.grid(row=0, column=2, padx=10, pady=8, sticky=tk.EW)
        
        # Preview section with better styling
        preview_frame = ttk.LabelFrame(process_frame, text="Preview & Testing", padding=15)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Enhanced test buttons layout
        test_buttons_frame = ttk.Frame(preview_frame)
        test_buttons_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Test buttons grid for better responsive layout
        test_grid = ttk.Frame(test_buttons_frame)
        test_grid.pack(fill=tk.X)
        
        for i in range(3):
            test_grid.grid_columnconfigure(i, weight=1)
        
        # Test crop button with improved styling
        test_crop_btn = ttk.Button(test_grid, text="🔍 Test Crop on First PDF", 
                                  command=self.test_crop, width=25)
        test_crop_btn.grid(row=0, column=0, padx=8, pady=8, sticky=tk.EW)
        
        # Test API button with improved styling
        test_api_btn = ttk.Button(test_grid, text="🔧 Test OpenAI API", 
                                 command=self.test_api, width=25)
        test_api_btn.grid(row=0, column=1, padx=8, pady=8, sticky=tk.EW)
        
        # Preview image
        self.preview_label = ttk.Label(preview_frame, text="No preview available")
        self.preview_label.pack(expand=True)
    
    def create_manual_crop_tab(self):
        """Create manual crop tab for interactive region selection"""
        tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(tab_frame, text="Manual Crop")
        
        # Manual crop controls
        controls_frame = ttk.Frame(tab_frame)
        controls_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Instructions
        instructions = ttk.Label(
            controls_frame,
            text="Select a PDF file to preview and drag to select crop region. Changes apply to all PDFs.",
            font=("Segoe UI", 10)
        )
        instructions.pack(anchor=tk.W, pady=(0, 10))
        
        # File selection frame
        file_frame = ttk.Frame(controls_frame)
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(file_frame, text="PDF File:", font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)
        
        self.manual_crop_file_var = tk.StringVar()
        file_combo = ttk.Combobox(
            file_frame,
            textvariable=self.manual_crop_file_var,
            state="readonly",
            width=40
        )
        file_combo.pack(side=tk.LEFT, padx=(10, 10))
        
        # Store reference for refresh_pdf_list
        self._manual_crop_combobox = file_combo
        
        ttk.Button(
            file_frame,
            text="Load PDF",
            command=self.load_pdf_for_crop
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(
            file_frame,
            text="Refresh List",
            command=self.refresh_pdf_list
        ).pack(side=tk.LEFT)
        
        # Crop action buttons
        action_frame = ttk.Frame(controls_frame)
        action_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(
            action_frame,
            text="Reset Selection",
            command=self.reset_crop_selection
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(
            action_frame,
            text="Apply Crop Settings",
            command=self.apply_manual_crop
        ).pack(side=tk.LEFT)
        
        # Preview area with two columns
        preview_main = ttk.Frame(tab_frame)
        preview_main.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Left column: Full PDF preview with crop selection
        left_frame = ttk.Frame(preview_main)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        ttk.Label(left_frame, text="PDF Preview (Drag to Select Region)", font=("Segoe UI", 11, "bold")).pack(anchor=tk.W, pady=(0, 5))
        
        # Canvas for PDF preview with scrollbars
        canvas_frame = ttk.Frame(left_frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        self.pdf_canvas = tk.Canvas(
            canvas_frame,
            bg='white',
            highlightthickness=1,
            highlightbackground='#cccccc'
        )
        
        v_scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.pdf_canvas.yview)
        h_scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.pdf_canvas.xview)
        
        self.pdf_canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.pdf_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Bind mouse events for crop selection
        self.pdf_canvas.bind("<Button-1>", self.start_crop_selection)
        self.pdf_canvas.bind("<B1-Motion>", self.update_crop_selection)
        self.pdf_canvas.bind("<ButtonRelease-1>", self.end_crop_selection)
        
        # Right column: Cropped preview
        right_frame = ttk.Frame(preview_main)
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))
        
        ttk.Label(right_frame, text="Crop Preview", font=("Segoe UI", 11, "bold")).pack(anchor=tk.W, pady=(0, 5))
        
        # Crop preview area
        self.crop_preview_frame = ttk.Frame(right_frame, width=300, height=250, relief="sunken", borderwidth=2)
        self.crop_preview_frame.pack(fill=tk.BOTH, expand=True)
        self.crop_preview_frame.pack_propagate(False)
        
        self.crop_preview_label = ttk.Label(
            self.crop_preview_frame,
            text="No crop region selected",
            anchor=tk.CENTER,
            font=("Segoe UI", 10)
        )
        self.crop_preview_label.pack(fill=tk.BOTH, expand=True)
        
        # Current crop coordinates display
        coords_frame = ttk.Frame(right_frame)
        coords_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Label(coords_frame, text="Crop Coordinates:", font=("Segoe UI", 10, "bold")).pack(anchor=tk.W)
        
        self.crop_coords_text = tk.Text(
            coords_frame,
            height=4,
            width=30,
            font=('Consolas', 9),
            bg='#f8f9fa',
            fg='#2c3e50',
            relief='flat',
            borderwidth=1,
            state='disabled'
        )
        self.crop_coords_text.pack(fill=tk.X, pady=(5, 0))

    def create_log_tab(self):
        """Create the log viewing tab"""
        log_frame = ttk.Frame(self.notebook)
        self.notebook.add(log_frame, text="Logs")
        
        # Enhanced log display with better styling for large screens
        self.log_text = scrolledtext.ScrolledText(log_frame, height=35, font=("Consolas", 10), wrap=tk.WORD,
                                                 background="#F8F9FA", foreground="#2C3E50")
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Enhanced log control buttons
        log_button_frame = ttk.Frame(log_frame)
        log_button_frame.pack(pady=10)
        
        log_btn_grid = ttk.Frame(log_button_frame)
        log_btn_grid.pack()
        
        for i in range(3):
            log_btn_grid.grid_columnconfigure(i, weight=1)
        
        ttk.Button(log_btn_grid, text="🗑️ Clear Log", command=self.clear_log, width=15).grid(row=0, column=0, padx=8, pady=5)
        ttk.Button(log_btn_grid, text="📋 Copy Log", command=self.copy_log, width=15).grid(row=0, column=1, padx=8, pady=5)
    
    def reset_crop_default(self):
        """Reset crop coordinates to default (1/4 of PDF size)"""
        self.x1_var.set(0)
        self.y1_var.set(0)
        self.x2_var.set(0.25)
        self.y2_var.set(0.25)
    
    def on_model_changed(self, event=None):
        """Handle model selection change"""
        self.update_model_description()
        self.log_message(f"Model changed to: {self.model_var.get()}")
    
    def update_model_description(self):
        """Update model description text"""
        model = self.model_var.get()
        if model in self.model_pricing:
            pricing = self.model_pricing[model]
            desc = f"{pricing['description']} | Input: ${pricing['input']}/1M tokens | Output: ${pricing['output']}/1M tokens"
            self.model_desc_var.set(desc)
    
    def reset_session_stats(self):
        """Reset session statistics"""
        self.session_stats = {
            'total_input_tokens': 0,
            'total_output_tokens': 0,
            'total_cost_usd': 0.0,
            'total_cost_thb': 0.0,
            'files_processed': 0,
            'api_calls': 0,
            'successful_files': 0,
            'failed_files': 0
        }
        self.update_cost_display()
        self.update_results_display()
        self.log_message("Session statistics reset")
    
    def update_cost_display(self):
        """Update cost tracking display"""
        stats = self.session_stats
        self.api_calls_var.set(str(stats['api_calls']))
        self.input_tokens_var.set(f"{stats['total_input_tokens']:,}")
        self.output_tokens_var.set(f"{stats['total_output_tokens']:,}")
        self.cost_usd_var.set(f"${stats['total_cost_usd']:.4f}")
        self.cost_thb_var.set(f"฿{stats['total_cost_thb']:.2f}")
        
        # Update session cost summary
        if hasattr(self, 'session_cost_var'):
            self.session_cost_var.set(f"Session Cost: ${stats['total_cost_usd']:.4f} USD (฿{stats['total_cost_thb']:.2f} THB) | {stats['api_calls']} API calls")
        
        # Update results display
        self.update_results_display()
    
    def calculate_cost(self, input_tokens: int, output_tokens: int, model: str) -> dict:
        """Calculate cost for token usage"""
        if model not in self.model_pricing:
            return {'usd': 0.0, 'thb': 0.0}
        
        pricing = self.model_pricing[model]
        
        # Calculate cost in USD
        input_cost = (input_tokens / 1_000_000) * pricing['input']
        output_cost = (output_tokens / 1_000_000) * pricing['output']
        total_usd = input_cost + output_cost
        
        # Convert to THB
        total_thb = total_usd * self.usd_to_thb
        
        return {
            'usd': total_usd,
            'thb': total_thb,
            'input_cost': input_cost,
            'output_cost': output_cost
        }
    
    def track_api_usage(self, input_tokens: int, output_tokens: int, model: str):
        """Track API usage and update costs"""
        cost_info = self.calculate_cost(input_tokens, output_tokens, model)
        
        # Update session stats
        self.session_stats['total_input_tokens'] += input_tokens
        self.session_stats['total_output_tokens'] += output_tokens
        self.session_stats['total_cost_usd'] += cost_info['usd']
        self.session_stats['total_cost_thb'] += cost_info['thb']
        self.session_stats['api_calls'] += 1
        
        # Update display
        self.update_cost_display()
        
        # Log the usage
        self.log_message(f"Token usage - Input: {input_tokens:,}, Output: {output_tokens:,}")
        self.log_message(f"Cost - ${cost_info['usd']:.4f} USD (฿{cost_info['thb']:.2f} THB)")
        
        return cost_info
    
    def update_results_display(self):
        """Update success/fail results display"""
        if hasattr(self, 'success_count_var'):
            self.success_count_var.set(str(self.session_stats['successful_files']))
        if hasattr(self, 'fail_count_var'):
            self.fail_count_var.set(str(self.session_stats['failed_files']))
        if hasattr(self, 'total_count_var'):
            total = self.session_stats['successful_files'] + self.session_stats['failed_files']
            self.total_count_var.set(str(total))
    
    def browse_pdf_folder(self):
        """Browse for PDF folder"""
        folder = filedialog.askdirectory(title="Select PDF Folder")
        if folder:
            self.pdf_folder_var.set(folder)
    
    def browse_csv_file(self):
        """Browse for CSV reference file"""
        file = filedialog.askopenfilename(
            title="Select CSV Reference File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file:
            self.csv_file_var.set(file)
    
    def open_pdf_folder(self):
        """Open the PDF folder in the system file explorer"""
        import subprocess
        import platform
        
        folder_path = self.pdf_folder_var.get() if hasattr(self, 'pdf_folder_var') else self.config['pdf_folder']
        
        # Ensure the folder exists
        if not os.path.exists(folder_path):
            messagebox.showwarning("Folder Not Found", 
                                 f"The PDF folder '{folder_path}' does not exist.\n\n"
                                 f"Please check the folder path in Settings or create the folder.")
            self.log_message(f"PDF folder not found: {folder_path}")
            return
        
        try:
            # Get the absolute path
            folder_path = os.path.abspath(folder_path)
            
            # Open folder based on operating system
            system = platform.system()
            if system == "Darwin":  # macOS
                subprocess.run(["open", folder_path])
            elif system == "Windows":  # Windows
                subprocess.run(["explorer", folder_path])
            elif system == "Linux":  # Linux
                subprocess.run(["xdg-open", folder_path])
            else:
                messagebox.showinfo("Unsupported OS", 
                                  f"Cannot open folder automatically on {system}.\n"
                                  f"Please navigate to: {folder_path}")
                return
            
            self.log_message(f"Opened PDF folder: {folder_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open folder:\n{e}")
            self.log_message(f"Error opening PDF folder: {e}")
    
    def load_settings(self):
        """Load settings from config file"""
        try:
            if os.path.exists('config.json'):
                with open('config.json', 'r') as f:
                    saved_config = json.load(f)
                    self.config.update(saved_config)
                    
                # Update GUI variables
                self.api_key_var.set(self.config.get('openai_api_key', ''))
                self.model_var.set(self.config.get('selected_model', 'gpt-4.1-nano'))
                self.x1_var.set(self.config.get('crop_x1', 0))
                self.y1_var.set(self.config.get('crop_y1', 0))
                self.x2_var.set(self.config.get('crop_x2', 0.25))
                self.y2_var.set(self.config.get('crop_y2', 0.25))
                self.pdf_folder_var.set(self.config.get('pdf_folder', 'workOrderPDF'))
                self.csv_file_var.set(self.config.get('ref_csv_file', 'workOrderRef/MCAN_work_inprogress.csv'))
                self.update_model_description()
        except Exception as e:
            self.logger.error(f"Error loading settings: {e}")
    
    def save_settings(self):
        """Save current settings to config file"""
        try:
            self.config.update({
                'openai_api_key': self.api_key_var.get(),
                'selected_model': self.model_var.get(),
                'crop_x1': self.x1_var.get(),
                'crop_y1': self.y1_var.get(),
                'crop_x2': self.x2_var.get(),
                'crop_y2': self.y2_var.get(),
                'pdf_folder': self.pdf_folder_var.get(),
                'ref_csv_file': self.csv_file_var.get()
            })
            
            with open('config.json', 'w') as f:
                json.dump(self.config, f, indent=2)
            
            messagebox.showinfo("Settings", "Settings saved successfully!")
            self.log_message("Settings saved successfully")
            
            # Reload reference data if CSV file changed
            self.reference_orders = self.load_reference_data()
            
        except Exception as e:
            self.logger.error(f"Error saving settings: {e}")
            messagebox.showerror("Error", f"Failed to save settings: {e}")
    
    def load_reference_data(self) -> set:
        """Load work order numbers from CSV reference file"""
        reference_orders = set()
        try:
            csv_file = self.csv_file_var.get() if hasattr(self, 'csv_file_var') else self.config['ref_csv_file']
            if os.path.exists(csv_file):
                with open(csv_file, 'r', encoding='utf-8') as f:
                    # Skip the header and read all order numbers
                    lines = f.readlines()[1:]  # Skip header
                    for line in lines:
                        line = line.strip()
                        if line:
                            reference_orders.add(line)
                self.log_message(f"Loaded {len(reference_orders)} reference work orders")
            else:
                self.log_message(f"Reference CSV file not found: {csv_file}")
        except Exception as e:
            self.logger.error(f"Error loading reference data: {e}")
            self.log_message(f"Error loading reference data: {e}")
        
        return reference_orders
    
    def log_message(self, message: str):
        """Add message to log display"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        if hasattr(self, 'log_text'):
            self.log_text.insert(tk.END, log_entry)
            self.log_text.see(tk.END)
            self.root.update_idletasks()
        
        self.logger.info(message)
    
    def clear_log(self):
        """Clear the log display"""
        if hasattr(self, 'log_text'):
            self.log_text.delete(1.0, tk.END)
    
    def copy_log(self):
        """Copy log contents to clipboard"""
        if hasattr(self, 'log_text'):
            log_content = self.log_text.get(1.0, tk.END)
            self.root.clipboard_clear()
            self.root.clipboard_append(log_content)
            messagebox.showinfo("Success", "Log copied to clipboard!")
    
    # Manual Crop Methods
    def refresh_pdf_list(self):
        """Refresh the PDF file list in the combobox"""
        try:
            if not hasattr(self, 'manual_crop_file_var') or self.manual_crop_file_var is None:
                return
                
            # Get PDF folder from various possible sources
            pdf_folder = self.config['pdf_folder']
            if hasattr(self, 'pdf_folder_var') and self.pdf_folder_var:
                try:
                    pdf_folder = self.pdf_folder_var.get()
                except:
                    pass  # Use config default
            if os.path.exists(pdf_folder):
                pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]
                
                # Store the combobox reference during tab creation
                if hasattr(self, '_manual_crop_combobox') and self._manual_crop_combobox:
                    self._manual_crop_combobox['values'] = pdf_files
                    if pdf_files and not self.manual_crop_file_var.get():
                        self.manual_crop_file_var.set(pdf_files[0])
                
                self.log_message(f"Found {len(pdf_files)} PDF files")
            else:
                self.log_message(f"PDF folder not found: {pdf_folder}")
        except Exception as e:
            self.log_message(f"Error refreshing PDF list: {e}")
    
    def load_pdf_for_crop(self):
        """Load selected PDF for crop preview"""
        try:
            if not hasattr(self, 'manual_crop_file_var') or not self.manual_crop_file_var.get():
                messagebox.showwarning("Warning", "Please select a PDF file first")
                return
            
            # Get PDF folder from various possible sources
            pdf_folder = self.config['pdf_folder']
            if hasattr(self, 'pdf_folder_var') and self.pdf_folder_var:
                try:
                    pdf_folder = self.pdf_folder_var.get()
                except:
                    pass  # Use config default
            pdf_path = os.path.join(pdf_folder, self.manual_crop_file_var.get())
            
            if not os.path.exists(pdf_path):
                messagebox.showerror("Error", f"PDF file not found: {pdf_path}")
                return
            
            self.log_message(f"Loading PDF for crop preview: {self.manual_crop_file_var.get()}")
            
            # Convert PDF to image (full page, no cropping)
            image = self.pdf_to_image_full(pdf_path)
            if image is None:
                messagebox.showerror("Error", "Failed to load PDF")
                return
            
            # Store original image for cropping
            self.crop_original_image = image
            
            # Resize image to fit canvas while maintaining aspect ratio
            canvas_width = 600
            canvas_height = 800
            
            # Calculate scaling factor
            scale_x = canvas_width / image.width
            scale_y = canvas_height / image.height
            scale = min(scale_x, scale_y, 1.0)  # Don't scale up
            
            new_width = int(image.width * scale)
            new_height = int(image.height * scale)
            
            # Resize image for display
            display_image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
            self.crop_display_image = display_image
            self.crop_scale_factor = scale
            
            # Convert to PhotoImage
            self.crop_photo = ImageTk.PhotoImage(display_image)
            
            # Clear canvas and display image
            self.pdf_canvas.delete("all")
            self.pdf_canvas.create_image(0, 0, anchor=tk.NW, image=self.crop_photo)
            
            # Update canvas scroll region
            self.pdf_canvas.configure(scrollregion=self.pdf_canvas.bbox("all"))
            
            # Reset selection
            self.selection_rect = None
            self.update_crop_preview()
            
            self.log_message(f"PDF loaded successfully: {new_width}x{new_height} (scale: {scale:.2f})")
            
        except Exception as e:
            self.logger.error(f"Error loading PDF for crop: {e}")
            self.log_message(f"Error loading PDF: {e}")
            messagebox.showerror("Error", f"Failed to load PDF: {e}")
    
    def pdf_to_image_full(self, pdf_path: str) -> Optional[Image.Image]:
        """Convert PDF to image without any cropping (for manual crop preview)"""
        try:
            # Try with pdf2image first
            pages = convert_from_path(pdf_path, first_page=1, last_page=1, dpi=150)
            if pages:
                return pages[0]
        except Exception as e:
            self.logger.warning(f"pdf2image failed for {pdf_path}: {e}")
            
        if HAS_PYMUPDF:
            try:
                # Fallback to PyMuPDF
                doc = fitz.open(pdf_path)
                page = doc[0]
                mat = fitz.Matrix(1.5, 1.5)
                pix = page.get_pixmap(matrix=mat)
                img_data = pix.tobytes("ppm")
                doc.close()
                
                from io import BytesIO
                return Image.open(BytesIO(img_data))
            except Exception as e:
                self.logger.error(f"PyMuPDF fallback failed for {pdf_path}: {e}")
        
        return None
    
    def start_crop_selection(self, event):
        """Start crop region selection"""
        if not hasattr(self, 'crop_photo'):
            return
        
        # Get canvas coordinates
        self.start_x = self.pdf_canvas.canvasx(event.x)
        self.start_y = self.pdf_canvas.canvasy(event.y)
        
        # Remove existing selection rectangle
        if self.selection_rect:
            self.pdf_canvas.delete(self.selection_rect)
        
        # Create new selection rectangle
        self.selection_rect = self.pdf_canvas.create_rectangle(
            self.start_x, self.start_y, self.start_x, self.start_y,
            outline='red', width=2, dash=(5, 5)
        )
    
    def update_crop_selection(self, event):
        """Update crop region selection during drag"""
        if not self.selection_rect or not hasattr(self, 'crop_photo'):
            return
        
        # Get current canvas coordinates
        self.current_x = self.pdf_canvas.canvasx(event.x)
        self.current_y = self.pdf_canvas.canvasy(event.y)
        
        # Update rectangle coordinates
        self.pdf_canvas.coords(
            self.selection_rect,
            self.start_x, self.start_y,
            self.current_x, self.current_y
        )
        
        # Update real-time preview
        self.update_crop_preview()
    
    def end_crop_selection(self, event):
        """End crop region selection"""
        if not self.selection_rect or not hasattr(self, 'crop_photo'):
            return
        
        # Final update
        self.current_x = self.pdf_canvas.canvasx(event.x)
        self.current_y = self.pdf_canvas.canvasy(event.y)
        
        # Update rectangle coordinates
        self.pdf_canvas.coords(
            self.selection_rect,
            self.start_x, self.start_y,
            self.current_x, self.current_y
        )
        
        # Update preview and coordinates display
        self.update_crop_preview()
        self.update_crop_coordinates()
    
    def update_crop_preview(self):
        """Update the real-time crop preview"""
        try:
            if not hasattr(self, 'crop_preview_label') or self.crop_preview_label is None:
                return  # Preview widget not available
                
            if not hasattr(self, 'crop_original_image') or not self.selection_rect:
                # No selection, show placeholder
                self.crop_preview_label.configure(image='', text="No crop region selected")
                return
            
            # Get selection coordinates (ensure proper order)
            x1 = min(self.start_x, self.current_x) if hasattr(self, 'current_x') else self.start_x
            y1 = min(self.start_y, self.current_y) if hasattr(self, 'current_y') else self.start_y
            x2 = max(self.start_x, self.current_x) if hasattr(self, 'current_x') else self.start_x
            y2 = max(self.start_y, self.current_y) if hasattr(self, 'current_y') else self.start_y
            
            # Check if selection is valid
            if abs(x2 - x1) < 10 or abs(y2 - y1) < 10:
                self.crop_preview_label.configure(image='', text="Selection too small")
                return
            
            # Convert display coordinates to original image coordinates
            orig_x1 = int(x1 / self.crop_scale_factor)
            orig_y1 = int(y1 / self.crop_scale_factor)
            orig_x2 = int(x2 / self.crop_scale_factor)
            orig_y2 = int(y2 / self.crop_scale_factor)
            
            # Ensure coordinates are within image bounds
            orig_x1 = max(0, min(orig_x1, self.crop_original_image.width))
            orig_y1 = max(0, min(orig_y1, self.crop_original_image.height))
            orig_x2 = max(0, min(orig_x2, self.crop_original_image.width))
            orig_y2 = max(0, min(orig_y2, self.crop_original_image.height))
            
            # Crop the original image
            cropped_image = self.crop_original_image.crop((orig_x1, orig_y1, orig_x2, orig_y2))
            
            # Resize for preview (maintain aspect ratio)
            preview_width = 280
            preview_height = 200
            
            # Calculate scaling
            scale_x = preview_width / cropped_image.width
            scale_y = preview_height / cropped_image.height
            scale = min(scale_x, scale_y)
            
            new_width = int(cropped_image.width * scale)
            new_height = int(cropped_image.height * scale)
            
            preview_image = cropped_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            # Convert to PhotoImage and display
            self.crop_preview_photo = ImageTk.PhotoImage(preview_image)
            self.crop_preview_label.configure(image=self.crop_preview_photo, text="")
            
        except Exception as e:
            self.logger.error(f"Error updating crop preview: {e}")
            if hasattr(self, 'crop_preview_label') and self.crop_preview_label is not None:
                self.crop_preview_label.configure(image='', text="Preview error")
    
    def update_crop_coordinates(self):
        """Update the crop coordinates display"""
        try:
            if not hasattr(self, 'crop_coords_text') or self.crop_coords_text is None:
                return  # Coordinates widget not available
                
            if not hasattr(self, 'crop_original_image') or not self.selection_rect:
                coords_text = "No selection"
            else:
                # Get selection coordinates (ensure proper order)
                x1 = min(self.start_x, self.current_x)
                y1 = min(self.start_y, self.current_y)
                x2 = max(self.start_x, self.current_x)
                y2 = max(self.start_y, self.current_y)
                
                # Convert to original image coordinates
                orig_x1 = int(x1 / self.crop_scale_factor)
                orig_y1 = int(y1 / self.crop_scale_factor)
                orig_x2 = int(x2 / self.crop_scale_factor)
                orig_y2 = int(y2 / self.crop_scale_factor)
                
                # Calculate relative coordinates (0.0 to 1.0)
                rel_x1 = orig_x1 / self.crop_original_image.width
                rel_y1 = orig_y1 / self.crop_original_image.height
                rel_x2 = orig_x2 / self.crop_original_image.width
                rel_y2 = orig_y2 / self.crop_original_image.height
                
                coords_text = f"Absolute:\\n({orig_x1}, {orig_y1}) to ({orig_x2}, {orig_y2})\\n\\nRelative:\\n({rel_x1:.3f}, {rel_y1:.3f}) to ({rel_x2:.3f}, {rel_y2:.3f})"
            
            # Update coordinates display
            self.crop_coords_text.configure(state='normal')
            self.crop_coords_text.delete(1.0, tk.END)
            self.crop_coords_text.insert(1.0, coords_text)
            self.crop_coords_text.configure(state='disabled')
            
        except Exception as e:
            self.logger.error(f"Error updating crop coordinates: {e}")
    
    def reset_crop_selection(self):
        """Reset the crop selection"""
        if hasattr(self, 'pdf_canvas') and self.selection_rect:
            self.pdf_canvas.delete(self.selection_rect)
            self.selection_rect = None
        
        # Reset preview
        if hasattr(self, 'crop_preview_label') and self.crop_preview_label is not None:
            self.crop_preview_label.configure(image='', text="No crop region selected")
        
        # Reset coordinates
        self.update_crop_coordinates()
        
        self.log_message("Crop selection reset")
    
    def apply_manual_crop(self):
        """Apply the manually selected crop region to all PDFs"""
        try:
            if not hasattr(self, 'crop_original_image') or not self.selection_rect:
                messagebox.showwarning("Warning", "Please select a crop region first")
                return
            
            # Get selection coordinates
            x1 = min(self.start_x, self.current_x)
            y1 = min(self.start_y, self.current_y)
            x2 = max(self.start_x, self.current_x)
            y2 = max(self.start_y, self.current_y)
            
            # Convert to original image coordinates
            orig_x1 = int(x1 / self.crop_scale_factor)
            orig_y1 = int(y1 / self.crop_scale_factor)
            orig_x2 = int(x2 / self.crop_scale_factor)
            orig_y2 = int(y2 / self.crop_scale_factor)
            
            # Calculate relative coordinates (0.0 to 1.0)
            rel_x1 = orig_x1 / self.crop_original_image.width
            rel_y1 = orig_y1 / self.crop_original_image.height
            rel_x2 = orig_x2 / self.crop_original_image.width
            rel_y2 = orig_y2 / self.crop_original_image.height
            
            # Update crop configuration
            self.config['crop_x1'] = rel_x1
            self.config['crop_y1'] = rel_y1
            self.config['crop_x2'] = rel_x2
            self.config['crop_y2'] = rel_y2
            
            # Update UI variables if they exist
            if hasattr(self, 'x1_var'):
                self.x1_var.set(rel_x1)
                self.y1_var.set(rel_y1)
                self.x2_var.set(rel_x2)
                self.y2_var.set(rel_y2)
            
            # Save to config file
            try:
                config_copy = self.config.copy()
                # Remove non-serializable items
                config_copy.pop('openai_api_key', None)
                with open('config.json', 'w') as f:
                    json.dump(config_copy, f, indent=2)
                self.log_message("Crop settings saved to config.json")
            except Exception as e:
                self.log_message(f"Warning: Could not save to config.json: {e}")
            
            self.log_message(f"Applied crop region: ({rel_x1:.3f}, {rel_y1:.3f}) to ({rel_x2:.3f}, {rel_y2:.3f})")
            messagebox.showinfo("Success", f"Crop settings applied!\\n\\nRegion: ({rel_x1:.3f}, {rel_y1:.3f}) to ({rel_x2:.3f}, {rel_y2:.3f})\\n\\nThis will be used for all PDF processing.")
            
        except Exception as e:
            self.logger.error(f"Error applying manual crop: {e}")
            self.log_message(f"Error applying crop settings: {e}")
            messagebox.showerror("Error", f"Failed to apply crop settings: {e}")
    
    def pdf_to_image(self, pdf_path: str) -> Optional[Image.Image]:
        """Convert first page of PDF to PIL Image with performance optimization and cropping"""
        try:
            # Try with pdf2image first with optimized settings
            pages = convert_from_path(
                pdf_path, 
                first_page=1, 
                last_page=1, 
                dpi=150,  # Reduced DPI for faster processing
                thread_count=2  # Use multiple threads for conversion
            )
            if pages:
                image = pages[0]
                
                # Apply crop if configured
                if hasattr(self, 'config') and 'crop_x1' in self.config:
                    width, height = image.size
                    x1 = int(width * self.config['crop_x1'])
                    y1 = int(height * self.config['crop_y1'])
                    x2 = int(width * self.config['crop_x2'])
                    y2 = int(height * self.config['crop_y2'])
                    image = image.crop((x1, y1, x2, y2))
                
                return image
        except Exception as e:
            self.logger.warning(f"pdf2image failed for {pdf_path}: {e}")
            
        if HAS_PYMUPDF:
            try:
                # Fallback to PyMuPDF with optimized settings
                doc = fitz.open(pdf_path)
                page = doc[0]
                mat = fitz.Matrix(1.5, 1.5)  # Reduced zoom for faster processing
                pix = page.get_pixmap(matrix=mat)
                img_data = pix.tobytes("ppm")
                doc.close()
                
                from io import BytesIO
                image = Image.open(BytesIO(img_data))
                
                # Apply crop if configured
                if hasattr(self, 'config') and 'crop_x1' in self.config:
                    width, height = image.size
                    x1 = int(width * self.config['crop_x1'])
                    y1 = int(height * self.config['crop_y1'])
                    x2 = int(width * self.config['crop_x2'])
                    y2 = int(height * self.config['crop_y2'])
                    image = image.crop((x1, y1, x2, y2))
                
                return image
            except Exception as e:
                self.logger.error(f"PyMuPDF fallback failed for {pdf_path}: {e}")
        
        self.logger.error(f"Failed to convert PDF to image: {pdf_path}")
        return None
    
    def crop_image(self, image: Image.Image, x1: float, y1: float, x2: float, y2: float) -> Image.Image:
        """Crop image using relative coordinates (0-1)"""
        width, height = image.size
        
        # Convert relative coordinates to absolute
        left = int(width * x1)
        top = int(height * y1)
        right = int(width * x2)
        bottom = int(height * y2)
        
        return image.crop((left, top, right, bottom))
    
    def test_crop(self):
        """Test crop settings on the first PDF"""
        try:
            pdf_folder = self.pdf_folder_var.get()
            self.log_message(f"Looking for PDFs in: {pdf_folder}")
            self.log_message(f"Folder exists: {os.path.exists(pdf_folder)}")
            
            if not os.path.exists(pdf_folder):
                messagebox.showerror("Error", f"PDF folder not found: {pdf_folder}")
                return
            
            # List all files for debugging
            all_files = os.listdir(pdf_folder)
            self.log_message(f"All files in folder: {all_files}")
            
            pdf_files = [f for f in all_files if f.lower().endswith('.pdf')]
            self.log_message(f"PDF files found: {pdf_files}")
            
            if not pdf_files:
                messagebox.showwarning("Warning", "No PDF files found in the specified folder")
                return
            
            # Use first PDF file
            first_pdf = os.path.join(pdf_folder, pdf_files[0])
            self.log_message(f"Testing crop on: {pdf_files[0]} (full path: {first_pdf})")
            
            # Convert to image
            image = self.pdf_to_image(first_pdf)
            if image is None:
                messagebox.showerror("Error", "Failed to convert PDF to image")
                return
            
            self.log_message(f"Original image size: {image.size}")
            
            # Crop image
            cropped = self.crop_image(image, 
                                    self.x1_var.get(), self.y1_var.get(),
                                    self.x2_var.get(), self.y2_var.get())
            
            self.log_message(f"Cropped image size: {cropped.size}")
            self.log_message(f"Crop coordinates: x1={self.x1_var.get()}, y1={self.y1_var.get()}, x2={self.x2_var.get()}, y2={self.y2_var.get()}")
            
            # Save cropped image for debugging
            debug_path = "debug_cropped.png"
            cropped.save(debug_path)
            self.log_message(f"Saved debug cropped image to: {debug_path}")
            
            # Display preview
            self.show_preview(cropped)
            self.log_message("Crop test completed successfully")
            
        except Exception as e:
            self.logger.error(f"Error testing crop: {e}")
            self.log_message(f"Error testing crop: {str(e)}")
            messagebox.showerror("Error", f"Failed to test crop: {e}")
    
    def test_api(self):
        """Test OpenAI API connection"""
        try:
            api_key = self.api_key_var.get().strip()
            if not api_key:
                messagebox.showerror("Error", "Please configure OpenAI API key first")
                return
            
            self.log_message("Testing OpenAI API connection...")
            
            # Test with a simple text prompt
            client = openai.OpenAI(api_key=api_key)
            selected_model = self.model_var.get()
            self.log_message(f"Testing API with model: {selected_model}")
            
            response = client.chat.completions.create(
                model=selected_model,
                messages=[{"role": "user", "content": "Say 'API connection successful'"}],
                max_tokens=50
            )
            
            # Track token usage for test
            if hasattr(response, 'usage'):
                usage = response.usage
                self.track_api_usage(usage.prompt_tokens, usage.completion_tokens, selected_model)
            
            result = response.choices[0].message.content
            self.log_message(f"API test successful! Response: {result}")
            messagebox.showinfo("Success", f"API connection successful!\nModel: {selected_model}\nResponse: {result}")
            
        except Exception as e:
            self.logger.error(f"API test failed: {e}")
            self.log_message(f"API test failed: {str(e)}")
            messagebox.showerror("Error", f"API test failed: {e}")
    
    def show_preview(self, image: Image.Image):
        """Show cropped image preview"""
        try:
            # Resize for display
            display_size = (300, 200)
            image.thumbnail(display_size, Image.Resampling.LANCZOS)
            
            # Convert to PhotoImage
            photo = ImageTk.PhotoImage(image)
            
            # Update preview label
            self.preview_label.configure(image=photo, text="")
            self.preview_label.image = photo  # Keep a reference
            
        except Exception as e:
            self.logger.error(f"Error showing preview: {e}")
    
    def extract_text_with_openai(self, image: Image.Image) -> Dict[str, str]:
        """Extract work order number and equipment number using OpenAI API with retry logic"""
        max_retries = 3
        retry_delay = 1  # seconds
        
        for attempt in range(max_retries):
            try:
                api_key = self.api_key_var.get().strip()
                if not api_key:
                    raise Exception("OpenAI API key not configured")
                
                # Set up OpenAI client with timeout
                client = openai.OpenAI(
                    api_key=api_key,
                    timeout=30.0  # 30 second timeout
                )
                
                # Optimize image encoding
                from io import BytesIO
                import base64
                
                buffered = BytesIO()
                # Optimize image quality for API
                image.save(buffered, format="JPEG", quality=85, optimize=True)
                img_str = base64.b64encode(buffered.getvalue()).decode()
                
                # Optimized prompt for faster processing
                prompt = """Extract from this work order document:
1. Work order number (8 digits after "Work Order No.")
2. Equipment number

Return JSON: {"work_order_number": "12345678", "equipment_number": "ABC123"}
Use null if not found."""
                
                selected_model = self.model_var.get()
                
                # Call OpenAI API with optimized settings
                response = client.chat.completions.create(
                    model=selected_model,
                    messages=[
                        {
                            "role": "user",
                            "content": [
                                {"type": "text", "text": prompt},
                                {
                                    "type": "image_url",
                                    "image_url": {
                                        "url": f"data:image/jpeg;base64,{img_str}",
                                        "detail": "low"  # Use low detail for faster processing
                                    }
                                }
                            ]
                        }
                    ],
                    max_tokens=150,  # Reduced for faster response
                    temperature=0.1  # Lower temperature for consistent results
                )
                
                # Extract token usage
                usage = response.usage
                self.track_api_usage(usage.prompt_tokens, usage.completion_tokens, selected_model)
                
                # Parse response
                result_text = response.choices[0].message.content
                
                # Parse JSON response
                try:
                    import json
                    # Clean the response text
                    clean_text = result_text.strip()
                    if clean_text.startswith('```json'):
                        clean_text = clean_text[7:]
                    if clean_text.endswith('```'):
                        clean_text = clean_text[:-3]
                    clean_text = clean_text.strip()
                    
                    result = json.loads(clean_text)
                    return {
                        'work_order_number': result.get('work_order_number'),
                        'equipment_number': result.get('equipment_number')
                    }
                except json.JSONDecodeError:
                    # Fallback: return empty result
                    return {'work_order_number': None, 'equipment_number': None}
                    
            except Exception as e:
                if attempt < max_retries - 1:
                    self.log_message(f"API call failed (attempt {attempt + 1}/{max_retries}), retrying in {retry_delay}s: {str(e)}")
                    time.sleep(retry_delay)
                    retry_delay *= 2  # Exponential backoff
                else:
                    self.logger.error(f"Error extracting text with OpenAI after {max_retries} attempts: {e}")
                    return {'work_order_number': None, 'equipment_number': None}
        
        return {'work_order_number': None, 'equipment_number': None}
    
    def process_single_pdf(self, pdf_path: str, filename: str) -> bool:
        """Process a single PDF file"""
        try:
            self.log_message(f"Processing: {filename}")
            
            # Convert PDF to image
            image = self.pdf_to_image(pdf_path)
            if image is None:
                self.log_message(f"Failed to convert {filename} to image")
                return False
            
            # Crop image
            cropped = self.crop_image(image,
                                    self.config['crop_x1'], self.config['crop_y1'],
                                    self.config['crop_x2'], self.config['crop_y2'])
            
            # Extract text using OpenAI
            extracted_data = self.extract_text_with_openai(cropped)
            work_order_num = extracted_data.get('work_order_number')
            equipment_num = extracted_data.get('equipment_number')
            
            self.log_message(f"Extracted - Work Order: {work_order_num}, Equipment: {equipment_num}")
            
            # Debug: Check CSV matching
            if work_order_num:
                self.log_message(f"Checking if work order {work_order_num} exists in {len(self.reference_orders)} reference orders")
                self.log_message(f"Work order type: {type(work_order_num)}")
                # Sample a few reference orders for debugging
                sample_orders = list(self.reference_orders)[:5] if self.reference_orders else []
                self.log_message(f"Sample reference orders: {sample_orders}")
                
                # Check both string and int comparison
                is_in_ref_str = work_order_num in self.reference_orders
                is_in_ref_int = str(work_order_num) in self.reference_orders if isinstance(work_order_num, int) else False
                is_in_ref_int_direct = int(work_order_num) in self.reference_orders if isinstance(work_order_num, str) and work_order_num.isdigit() else False
                
                self.log_message(f"Match check - as string: {is_in_ref_str}, as int->str: {is_in_ref_int}, as str->int: {is_in_ref_int_direct}")
            
            # Check if work order exists in reference data (handle both string and int types)
            work_order_match = False
            if work_order_num:
                # Try string comparison first
                work_order_str = str(work_order_num)
                work_order_match = work_order_str in self.reference_orders
                
                if not work_order_match:
                    # Try int comparison if string fails
                    try:
                        work_order_int = int(work_order_num)
                        work_order_match = str(work_order_int) in self.reference_orders
                    except:
                        pass
                
                self.log_message(f"Final match result for {work_order_str}: {work_order_match}")
            
            if work_order_num and work_order_match:
                # Match found - rename file
                if equipment_num:
                    new_filename = f"CS-{work_order_num}-{equipment_num}.pdf"
                else:
                    new_filename = f"CS-{work_order_num}-NoEquip.pdf"
                
                new_path = os.path.join(self.config['pdf_folder'], new_filename)
                
                # Rename file
                try:
                    os.rename(pdf_path, new_path)
                    self.log_message(f"Renamed to: {new_filename}")
                    return True
                except Exception as e:
                    self.log_message(f"Failed to rename {filename}: {e}")
                    return False
            else:
                # No match - move to not_match folder
                # Ensure not_match folder exists
                os.makedirs(self.config['not_match_folder'], exist_ok=True)
                not_match_path = os.path.join(self.config['not_match_folder'], filename)
                try:
                    shutil.move(pdf_path, not_match_path)
                    self.log_message(f"Moved to not_match folder: {filename}")
                    return True
                except Exception as e:
                    self.log_message(f"Failed to move {filename} to not_match: {e}")
                    return False
                    
        except Exception as e:
            self.logger.error(f"Error processing {filename}: {e}")
            self.log_message(f"Error processing {filename}: {e}")
            return False
    
    def start_processing(self):
        """Start processing all PDF files"""
        # Validate settings
        if not self.api_key_var.get().strip():
            messagebox.showerror("Error", "Please configure OpenAI API key")
            return
        
        pdf_folder = self.pdf_folder_var.get()
        if not os.path.exists(pdf_folder):
            messagebox.showerror("Error", f"PDF folder not found: {pdf_folder}")
            return
        
        # Update config
        self.config.update({
            'openai_api_key': self.api_key_var.get(),
            'selected_model': self.model_var.get(),
            'crop_x1': self.x1_var.get(),
            'crop_y1': self.y1_var.get(),
            'crop_x2': self.x2_var.get(),
            'crop_y2': self.y2_var.get(),
            'pdf_folder': self.pdf_folder_var.get(),
            'ref_csv_file': self.csv_file_var.get()
        })
        
        # Reload reference data
        self.reference_orders = self.load_reference_data()
        
        # Start processing in separate thread
        self.processing_thread = threading.Thread(target=self.process_all_pdfs)
        self.processing_thread.daemon = True
        self.processing_thread.start()
        
        # Update UI
        self.process_button.configure(state=tk.DISABLED)
        self.stop_button.configure(state=tk.NORMAL)
        self.stop_processing_flag = False
    
    def process_all_pdfs(self):
        """Process all PDF files in the folder with concurrent processing"""
        try:
            pdf_folder = self.config['pdf_folder']
            pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]
            
            if not pdf_files:
                self.log_message("No PDF files found to process")
                self.processing_complete()
                return
            
            self.log_message(f"Found {len(pdf_files)} PDF files to process")
            self.log_message(f"Using {self.max_concurrent_workers} concurrent workers for API calls")
            
            start_time = time.time()
            
            # Use ThreadPoolExecutor for concurrent processing
            with ThreadPoolExecutor(max_workers=self.max_concurrent_workers) as executor:
                # Submit all tasks
                future_to_file = {
                    executor.submit(self.process_single_pdf, 
                                   os.path.join(pdf_folder, filename), 
                                   filename): filename 
                    for filename in pdf_files
                }
                
                processed = 0
                successful = 0
                failed = 0
                
                # Process completed tasks as they finish
                for future in as_completed(future_to_file):
                    if self.stop_processing_flag:
                        self.log_message("Processing stopped by user - cancelling remaining tasks")
                        # Cancel remaining futures
                        for f in future_to_file:
                            if not f.done():
                                f.cancel()
                        break
                    
                    filename = future_to_file[future]
                    processed += 1
                    
                    try:
                        result = future.result()
                        if result:
                            successful += 1
                            self.session_stats['successful_files'] += 1
                        else:
                            failed += 1
                            self.session_stats['failed_files'] += 1
                    except Exception as e:
                        self.logger.error(f"Error processing {filename}: {e}")
                        self.log_message(f"Error processing {filename}: {e}")
                        failed += 1
                        self.session_stats['failed_files'] += 1
                    
                    self.session_stats['files_processed'] += 1
                    
                    # Update progress
                    progress = (processed / len(pdf_files)) * 100
                    self.progress_var.set(progress)
                    self.status_var.set(f"Processed {processed}/{len(pdf_files)} files (Success: {successful}, Failed: {failed})")
                    self.root.update_idletasks()
                    
                    # Update results display
                    self.update_results_display()
                    
                    # Log progress every batch
                    if processed % self.batch_size == 0 or processed == len(pdf_files):
                        elapsed = time.time() - start_time
                        rate = processed / elapsed if elapsed > 0 else 0
                        remaining = len(pdf_files) - processed
                        eta = remaining / rate if rate > 0 else 0
                        self.log_message(f"Progress: {processed}/{len(pdf_files)} files ({rate:.1f} files/sec, ETA: {eta:.0f}s)")
            
            # Complete
            total_time = time.time() - start_time
            self.progress_var.set(100)
            self.status_var.set(f"Completed: {successful} Success, {failed} Failed (Total: {processed} in {total_time:.1f}s)")
            self.log_message(f"Processing complete: {successful} successful, {failed} failed")
            self.log_message(f"Total time: {total_time:.1f}s, Average: {total_time/processed:.1f}s per file")
            
        except Exception as e:
            self.logger.error(f"Error in process_all_pdfs: {e}")
            self.log_message(f"Processing error: {e}")
        finally:
            self.processing_complete()
    
    def stop_processing(self):
        """Stop the processing"""
        self.stop_processing_flag = True
        self.log_message("Stopping processing...")
    
    def processing_complete(self):
        """Called when processing is complete"""
        self.process_button.configure(state=tk.NORMAL)
        self.stop_button.configure(state=tk.DISABLED)
        self.stop_processing_flag = False

def main():
    root = tk.Tk()
    app = WorkOrderExtractor(root)
    root.mainloop()

if __name__ == "__main__":
    main()