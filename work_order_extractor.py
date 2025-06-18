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
        # Enhanced responsive sizing for large monitors and laptops
        self.root.geometry("1400x900")
        self.root.minsize(1200, 800)
        
        # Center window on screen
        self.root.update_idletasks()
        width = 1400
        height = 900
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
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
        """Create the main GUI widgets"""
        # Configure enhanced style for better appearance
        style = ttk.Style()
        style.theme_use('clam')  # Use modern theme
        
        # Configure custom colors and styles
        style.configure('Title.TLabel', font=('Arial', 12, 'bold'), foreground='#2C3E50')
        style.configure('Header.TLabel', font=('Arial', 10, 'bold'), foreground='#34495E')
        style.configure('Info.TLabel', font=('Arial', 9), foreground='#5D6D7E')
        style.configure('Success.TLabel', font=('Arial', 9, 'bold'), foreground='#27AE60')
        style.configure('Error.TLabel', font=('Arial', 9, 'bold'), foreground='#E74C3C')
        style.configure('Warning.TLabel', font=('Arial', 9, 'bold'), foreground='#F39C12')
        style.configure('Primary.TLabel', font=('Arial', 9, 'bold'), foreground='#3498DB')
        
        # Main notebook for tabs with better padding
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Settings tab
        self.create_settings_tab()
        
        # Processing tab
        self.create_processing_tab()
        
        # Log tab
        self.create_log_tab()
    
    def create_settings_tab(self):
        """Create the settings configuration tab"""
        settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(settings_frame, text="Settings")
        
        # API Key section
        api_frame = ttk.LabelFrame(settings_frame, text="OpenAI API Configuration", padding=10)
        api_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(api_frame, text="API Key:").pack(anchor=tk.W)
        self.api_key_var = tk.StringVar()
        api_entry = ttk.Entry(api_frame, textvariable=self.api_key_var, show="*", width=50)
        api_entry.pack(fill=tk.X, pady=2)
        
        # Model selection section
        model_frame = ttk.LabelFrame(settings_frame, text="Model Selection", padding=10)
        model_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(model_frame, text="OpenAI Model:").pack(anchor=tk.W)
        self.model_var = tk.StringVar(value=self.config['selected_model'])
        
        model_combo_frame = ttk.Frame(model_frame)
        model_combo_frame.pack(fill=tk.X, pady=2)
        
        self.model_combo = ttk.Combobox(model_combo_frame, textvariable=self.model_var, 
                                       values=list(self.model_pricing.keys()), 
                                       state="readonly", width=20)
        self.model_combo.pack(side=tk.LEFT)
        self.model_combo.bind('<<ComboboxSelected>>', self.on_model_changed)
        
        # Model description with enhanced styling
        self.model_desc_var = tk.StringVar()
        self.update_model_description()
        desc_label = ttk.Label(model_frame, textvariable=self.model_desc_var, style='Info.TLabel')
        desc_label.pack(anchor=tk.W, pady=(5,0))
        
        # Cost tracking section
        cost_frame = ttk.LabelFrame(settings_frame, text="Cost Tracking (Session)", padding=15)
        cost_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Create enhanced cost display grid with better spacing and responsive design
        cost_grid = ttk.Frame(cost_frame)
        cost_grid.pack(fill=tk.X, padx=5, pady=5)
        
        # Configure grid columns to be responsive for large screens
        for i in range(6):
            cost_grid.grid_columnconfigure(i, weight=1)
        
        # First row - API calls and tokens
        ttk.Label(cost_grid, text="API Calls:", style='Header.TLabel').grid(row=0, column=0, sticky=tk.W, padx=8, pady=3)
        self.api_calls_var = tk.StringVar(value="0")
        ttk.Label(cost_grid, textvariable=self.api_calls_var, style='Primary.TLabel').grid(row=0, column=1, sticky=tk.W, padx=8, pady=3)
        
        ttk.Label(cost_grid, text="Input Tokens:", style='Header.TLabel').grid(row=0, column=2, sticky=tk.W, padx=8, pady=3)
        self.input_tokens_var = tk.StringVar(value="0")
        ttk.Label(cost_grid, textvariable=self.input_tokens_var, style='Success.TLabel').grid(row=0, column=3, sticky=tk.W, padx=8, pady=3)
        
        ttk.Label(cost_grid, text="Output Tokens:", style='Header.TLabel').grid(row=0, column=4, sticky=tk.W, padx=8, pady=3)
        self.output_tokens_var = tk.StringVar(value="0")
        ttk.Label(cost_grid, textvariable=self.output_tokens_var, style='Warning.TLabel').grid(row=0, column=5, sticky=tk.W, padx=8, pady=3)
        
        # Second row - costs
        ttk.Label(cost_grid, text="Cost (USD):", style='Header.TLabel').grid(row=1, column=0, sticky=tk.W, padx=8, pady=3)
        self.cost_usd_var = tk.StringVar(value="$0.00")
        ttk.Label(cost_grid, textvariable=self.cost_usd_var, style='Error.TLabel').grid(row=1, column=1, sticky=tk.W, padx=8, pady=3)
        
        ttk.Label(cost_grid, text="Cost (THB):", style='Header.TLabel').grid(row=1, column=2, sticky=tk.W, padx=8, pady=3)
        self.cost_thb_var = tk.StringVar(value="฿0.00")
        ttk.Label(cost_grid, textvariable=self.cost_thb_var, font=("Arial", 9, "bold"), foreground="#8E44AD").grid(row=1, column=3, sticky=tk.W, padx=8, pady=3)
        
        # Reset stats button with improved styling
        reset_stats_btn = ttk.Button(cost_frame, text="🔄 Reset Session Stats", 
                                    command=self.reset_session_stats)
        reset_stats_btn.pack(pady=8)
        
        # Crop coordinates section with improved layout
        crop_frame = ttk.LabelFrame(settings_frame, text="Crop Coordinates (as fraction of PDF size)", padding=15)
        crop_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Enhanced coordinate inputs with validation
        coord_frame = ttk.Frame(crop_frame)
        coord_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Configure grid for better responsiveness on large screens
        for i in range(8):
            coord_frame.grid_columnconfigure(i, weight=1)
        
        # Enhanced coordinate inputs with better layout for large screens
        ttk.Label(coord_frame, text="X1 (left):", style='Header.TLabel').grid(row=0, column=0, sticky=tk.W, padx=8, pady=5)
        self.x1_var = tk.DoubleVar(value=self.config['crop_x1'])
        ttk.Entry(coord_frame, textvariable=self.x1_var, width=15, font=("Arial", 10)).grid(row=0, column=1, padx=8, pady=5, sticky=tk.EW)
        
        ttk.Label(coord_frame, text="Y1 (top):", style='Header.TLabel').grid(row=0, column=2, sticky=tk.W, padx=8, pady=5)
        self.y1_var = tk.DoubleVar(value=self.config['crop_y1'])
        ttk.Entry(coord_frame, textvariable=self.y1_var, width=15, font=("Arial", 10)).grid(row=0, column=3, padx=8, pady=5, sticky=tk.EW)
        
        ttk.Label(coord_frame, text="X2 (right):", style='Header.TLabel').grid(row=0, column=4, sticky=tk.W, padx=8, pady=5)
        self.x2_var = tk.DoubleVar(value=self.config['crop_x2'])
        ttk.Entry(coord_frame, textvariable=self.x2_var, width=15, font=("Arial", 10)).grid(row=0, column=5, padx=8, pady=5, sticky=tk.EW)
        
        ttk.Label(coord_frame, text="Y2 (bottom):", style='Header.TLabel').grid(row=0, column=6, sticky=tk.W, padx=8, pady=5)
        self.y2_var = tk.DoubleVar(value=self.config['crop_y2'])
        ttk.Entry(coord_frame, textvariable=self.y2_var, width=15, font=("Arial", 10)).grid(row=0, column=7, padx=8, pady=5, sticky=tk.EW)
        
        # Reset to default button
        reset_btn = ttk.Button(crop_frame, text="🔄 Reset to Default (1/4 size)", 
                              command=self.reset_crop_default)
        reset_btn.pack(pady=8)
        
        # File paths section with better spacing
        paths_frame = ttk.LabelFrame(settings_frame, text="File Paths", padding=15)
        paths_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Enhanced PDF folder section
        ttk.Label(paths_frame, text="📁 PDF Folder:", style='Header.TLabel').pack(anchor=tk.W, pady=(5,2))
        self.pdf_folder_var = tk.StringVar(value=self.config['pdf_folder'])
        pdf_frame = ttk.Frame(paths_frame)
        pdf_frame.pack(fill=tk.X, pady=5)
        ttk.Entry(pdf_frame, textvariable=self.pdf_folder_var, width=60, font=("Arial", 10)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        ttk.Button(pdf_frame, text="📂 Browse", command=self.browse_pdf_folder, width=12).pack(side=tk.RIGHT)
        
        # Enhanced CSV reference file section
        ttk.Label(paths_frame, text="📄 Reference CSV File:", style='Header.TLabel').pack(anchor=tk.W, pady=(15,2))
        self.csv_file_var = tk.StringVar(value=self.config['ref_csv_file'])
        csv_frame = ttk.Frame(paths_frame)
        csv_frame.pack(fill=tk.X, pady=5)
        ttk.Entry(csv_frame, textvariable=self.csv_file_var, width=60, font=("Arial", 10)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        ttk.Button(csv_frame, text="📂 Browse", command=self.browse_csv_file, width=12).pack(side=tk.RIGHT)
        
        # Enhanced save settings button
        save_btn = ttk.Button(settings_frame, text="💾 Save Settings", 
                             command=self.save_settings, width=20)
        save_btn.pack(pady=20)
    
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
    
    def pdf_to_image(self, pdf_path: str) -> Optional[Image.Image]:
        """Convert first page of PDF to PIL Image"""
        try:
            # Try with pdf2image first
            pages = convert_from_path(pdf_path, first_page=1, last_page=1, dpi=200)
            if pages:
                return pages[0]
        except Exception as e:
            self.logger.warning(f"pdf2image failed for {pdf_path}: {e}")
            
        if HAS_PYMUPDF:
            try:
                # Fallback to PyMuPDF
                doc = fitz.open(pdf_path)
                page = doc[0]
                mat = fitz.Matrix(2.0, 2.0)  # 2x zoom
                pix = page.get_pixmap(matrix=mat)
                img_data = pix.tobytes("ppm")
                doc.close()
                
                from io import BytesIO
                return Image.open(BytesIO(img_data))
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
        """Extract work order number and equipment number using OpenAI API"""
        try:
            api_key = self.api_key_var.get().strip()
            if not api_key:
                self.log_message("OpenAI API key not configured")
                raise Exception("OpenAI API key not configured")
            
            self.log_message("Starting OpenAI API call...")
            
            # Set up OpenAI client
            client = openai.OpenAI(api_key=api_key)
            
            # Save image to bytes
            from io import BytesIO
            import base64
            
            buffered = BytesIO()
            image.save(buffered, format="PNG")
            img_str = base64.b64encode(buffered.getvalue()).decode()
            self.log_message(f"Image converted to base64, size: {len(img_str)} characters")
            
            # Prepare the prompt
            prompt = """Extract work order number (8 digits after "Work Order No.") and extract Equipment No. from this work order document. 
            Return the response in JSON format with keys "work_order_number" and "equipment_number". 
            If you cannot find either value, set it to null."""
            
            # Get selected model
            selected_model = self.model_var.get()
            self.log_message(f"Using model: {selected_model}")
            
            # Call OpenAI API
            self.log_message("Calling OpenAI API...")
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
                                    "url": f"data:image/png;base64,{img_str}"
                                }
                            }
                        ]
                    }
                ],
                max_tokens=300
            )
            
            # Extract token usage
            usage = response.usage
            input_tokens = usage.prompt_tokens
            output_tokens = usage.completion_tokens
            total_tokens = usage.total_tokens
            
            # Track API usage and costs
            self.track_api_usage(input_tokens, output_tokens, selected_model)
            
            # Parse response
            result_text = response.choices[0].message.content
            self.log_message(f"OpenAI response: {result_text}")
            self.log_message(f"Total tokens used: {total_tokens:,} (Input: {input_tokens:,}, Output: {output_tokens:,})")
            
            # Try to parse as JSON
            try:
                import json
                # Clean the response text - remove markdown code blocks
                clean_text = result_text.strip()
                if clean_text.startswith('```json'):
                    clean_text = clean_text[7:]  # Remove ```json
                if clean_text.endswith('```'):
                    clean_text = clean_text[:-3]  # Remove ```
                clean_text = clean_text.strip()
                
                self.log_message(f"Cleaned JSON text: {clean_text}")
                
                result = json.loads(clean_text)
                extracted_data = {
                    'work_order_number': result.get('work_order_number'),
                    'equipment_number': result.get('equipment_number')
                }
                self.log_message(f"Successfully parsed JSON: {extracted_data}")
                return extracted_data
            except json.JSONDecodeError as e:
                # Fallback: try to extract manually from text
                self.log_message(f"Failed to parse JSON response, error: {e}, raw text: {result_text}")
                return {'work_order_number': None, 'equipment_number': None}
                
        except Exception as e:
            self.logger.error(f"Error extracting text with OpenAI: {e}")
            self.log_message(f"OpenAI API error: {str(e)}")
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
        """Process all PDF files in the folder"""
        try:
            pdf_folder = self.config['pdf_folder']
            pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]
            
            if not pdf_files:
                self.log_message("No PDF files found to process")
                self.processing_complete()
                return
            
            self.log_message(f"Found {len(pdf_files)} PDF files to process")
            
            processed = 0
            successful = 0
            
            for i, filename in enumerate(pdf_files):
                if self.stop_processing_flag:
                    self.log_message("Processing stopped by user")
                    break
                
                pdf_path = os.path.join(pdf_folder, filename)
                
                # Update progress
                progress = (i / len(pdf_files)) * 100
                self.progress_var.set(progress)
                self.status_var.set(f"Processing {i+1}/{len(pdf_files)}: {filename}")
                self.root.update_idletasks()
                
                # Process file
                if self.process_single_pdf(pdf_path, filename):
                    successful += 1
                    self.session_stats['successful_files'] += 1
                else:
                    self.session_stats['failed_files'] += 1
                
                processed += 1
                self.session_stats['files_processed'] += 1
                
                # Update results display in real-time
                self.update_results_display()
            
            # Complete
            self.progress_var.set(100)
            success_count = self.session_stats['successful_files']
            fail_count = self.session_stats['failed_files']
            self.status_var.set(f"Completed: {success_count} Success, {fail_count} Failed (Total: {processed})")
            self.log_message(f"Processing complete: {success_count} successful, {fail_count} failed (Total: {processed} files)")
            
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