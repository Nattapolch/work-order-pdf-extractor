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
        self.root.geometry("800x700")
        
        # Configuration
        self.config = {
            'openai_api_key': '',
            'crop_x1': 0,
            'crop_y1': 0,
            'crop_x2': 0.25,  # 1/4 of width
            'crop_y2': 0.25,  # 1/4 of height
            'pdf_folder': 'workOrderPDF',
            'ref_csv_file': 'workOrderRef/MCAN_work_inprogress.csv',
            'not_match_folder': 'not_match'
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
        # Main notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
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
        
        # Crop coordinates section
        crop_frame = ttk.LabelFrame(settings_frame, text="Crop Coordinates (as fraction of PDF size)", padding=10)
        crop_frame.pack(fill=tk.X, padx=10, pady=5)
        
        coord_frame = ttk.Frame(crop_frame)
        coord_frame.pack(fill=tk.X)
        
        # X1, Y1 (top-left)
        ttk.Label(coord_frame, text="X1 (left):").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.x1_var = tk.DoubleVar(value=self.config['crop_x1'])
        ttk.Entry(coord_frame, textvariable=self.x1_var, width=10).grid(row=0, column=1, padx=5)
        
        ttk.Label(coord_frame, text="Y1 (top):").grid(row=0, column=2, sticky=tk.W, padx=5)
        self.y1_var = tk.DoubleVar(value=self.config['crop_y1'])
        ttk.Entry(coord_frame, textvariable=self.y1_var, width=10).grid(row=0, column=3, padx=5)
        
        # X2, Y2 (bottom-right)
        ttk.Label(coord_frame, text="X2 (right):").grid(row=1, column=0, sticky=tk.W, padx=5)
        self.x2_var = tk.DoubleVar(value=self.config['crop_x2'])
        ttk.Entry(coord_frame, textvariable=self.x2_var, width=10).grid(row=1, column=1, padx=5)
        
        ttk.Label(coord_frame, text="Y2 (bottom):").grid(row=1, column=2, sticky=tk.W, padx=5)
        self.y2_var = tk.DoubleVar(value=self.config['crop_y2'])
        ttk.Entry(coord_frame, textvariable=self.y2_var, width=10).grid(row=1, column=3, padx=5)
        
        # Reset to default button
        ttk.Button(crop_frame, text="Reset to Default (1/4 size)", 
                  command=self.reset_crop_default).pack(pady=5)
        
        # File paths section
        paths_frame = ttk.LabelFrame(settings_frame, text="File Paths", padding=10)
        paths_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # PDF folder
        ttk.Label(paths_frame, text="PDF Folder:").pack(anchor=tk.W)
        self.pdf_folder_var = tk.StringVar(value=self.config['pdf_folder'])
        pdf_frame = ttk.Frame(paths_frame)
        pdf_frame.pack(fill=tk.X, pady=2)
        ttk.Entry(pdf_frame, textvariable=self.pdf_folder_var, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(pdf_frame, text="Browse", command=self.browse_pdf_folder).pack(side=tk.RIGHT, padx=(5,0))
        
        # CSV reference file
        ttk.Label(paths_frame, text="Reference CSV File:").pack(anchor=tk.W, pady=(10,0))
        self.csv_file_var = tk.StringVar(value=self.config['ref_csv_file'])
        csv_frame = ttk.Frame(paths_frame)
        csv_frame.pack(fill=tk.X, pady=2)
        ttk.Entry(csv_frame, textvariable=self.csv_file_var, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(csv_frame, text="Browse", command=self.browse_csv_file).pack(side=tk.RIGHT, padx=(5,0))
        
        # Save settings button
        ttk.Button(settings_frame, text="Save Settings", 
                  command=self.save_settings).pack(pady=10)
    
    def create_processing_tab(self):
        """Create the main processing tab"""
        process_frame = ttk.Frame(self.notebook)
        self.notebook.add(process_frame, text="Process PDFs")
        
        # Status section
        status_frame = ttk.LabelFrame(process_frame, text="Status", padding=10)
        status_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.status_var = tk.StringVar(value="Ready to process")
        ttk.Label(status_frame, textvariable=self.status_var).pack(anchor=tk.W)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var, 
                                          maximum=100, length=300)
        self.progress_bar.pack(fill=tk.X, pady=5)
        
        # Control buttons
        button_frame = ttk.Frame(process_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.process_button = ttk.Button(button_frame, text="Start Processing", 
                                       command=self.start_processing)
        self.process_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(button_frame, text="Stop", 
                                    command=self.stop_processing, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        # Preview section
        preview_frame = ttk.LabelFrame(process_frame, text="Preview", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Test crop button
        ttk.Button(preview_frame, text="Test Crop on First PDF", 
                  command=self.test_crop).pack(pady=5)
        
        # Test API button
        ttk.Button(preview_frame, text="Test OpenAI API", 
                  command=self.test_api).pack(pady=5)
        
        # Preview image
        self.preview_label = ttk.Label(preview_frame, text="No preview available")
        self.preview_label.pack(expand=True)
    
    def create_log_tab(self):
        """Create the log viewing tab"""
        log_frame = ttk.Frame(self.notebook)
        self.notebook.add(log_frame, text="Logs")
        
        # Log display
        self.log_text = scrolledtext.ScrolledText(log_frame, height=30)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Log control buttons
        log_button_frame = ttk.Frame(log_frame)
        log_button_frame.pack(pady=5)
        
        ttk.Button(log_button_frame, text="Clear Log", command=self.clear_log).pack(side=tk.LEFT, padx=5)
        ttk.Button(log_button_frame, text="Copy Log", command=self.copy_log).pack(side=tk.LEFT, padx=5)
    
    def reset_crop_default(self):
        """Reset crop coordinates to default (1/4 of PDF size)"""
        self.x1_var.set(0)
        self.y1_var.set(0)
        self.x2_var.set(0.25)
        self.y2_var.set(0.25)
    
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
                self.x1_var.set(self.config.get('crop_x1', 0))
                self.y1_var.set(self.config.get('crop_y1', 0))
                self.x2_var.set(self.config.get('crop_x2', 0.25))
                self.y2_var.set(self.config.get('crop_y2', 0.25))
                self.pdf_folder_var.set(self.config.get('pdf_folder', 'workOrderPDF'))
                self.csv_file_var.set(self.config.get('ref_csv_file', 'workOrderRef/MCAN_work_inprogress.csv'))
        except Exception as e:
            self.logger.error(f"Error loading settings: {e}")
    
    def save_settings(self):
        """Save current settings to config file"""
        try:
            self.config.update({
                'openai_api_key': self.api_key_var.get(),
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
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": "Say 'API connection successful'"}],
                max_tokens=50
            )
            
            result = response.choices[0].message.content
            self.log_message(f"API test successful! Response: {result}")
            messagebox.showinfo("Success", f"API connection successful!\nResponse: {result}")
            
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
            
            # Call OpenAI API
            self.log_message("Calling OpenAI API...")
            response = client.chat.completions.create(
                model="gpt-4o-mini",  # Using the smaller, faster model
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
            
            # Parse response
            result_text = response.choices[0].message.content
            self.log_message(f"OpenAI response: {result_text}")
            
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
                
                processed += 1
            
            # Complete
            self.progress_var.set(100)
            self.status_var.set(f"Completed: {successful}/{processed} files processed successfully")
            self.log_message(f"Processing complete: {successful}/{processed} files processed successfully")
            
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