#!/usr/bin/env python3
"""
Performance Test Script for Work Order Extractor
Tests the performance improvements with concurrent processing.
"""

import os
import time
import shutil
from pathlib import Path
import json

def create_test_pdfs(test_folder: str, num_files: int = 20):
    """Create test PDF files by copying existing ones"""
    source_folder = "workOrderPDF"
    
    if not os.path.exists(source_folder):
        print(f"Source folder {source_folder} not found")
        return False
    
    # Get existing PDF files
    pdf_files = [f for f in os.listdir(source_folder) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print(f"No PDF files found in {source_folder}")
        return False
    
    # Create test folder
    os.makedirs(test_folder, exist_ok=True)
    
    # Copy files to create test set
    source_pdf = os.path.join(source_folder, pdf_files[0])
    
    for i in range(num_files):
        test_pdf = os.path.join(test_folder, f"test_pdf_{i:03d}.pdf")
        shutil.copy2(source_pdf, test_pdf)
    
    print(f"Created {num_files} test PDF files in {test_folder}")
    return True

def run_performance_test():
    """Run performance test comparing original vs optimized processing"""
    
    # Test configurations
    test_configs = [
        {"name": "Small Batch", "files": 10, "workers": 3},
        {"name": "Medium Batch", "files": 25, "workers": 5},
        {"name": "Large Batch", "files": 50, "workers": 8},
    ]
    
    results = []
    
    for config in test_configs:
        print(f"\n{'='*50}")
        print(f"Running test: {config['name']}")
        print(f"Files: {config['files']}, Workers: {config['workers']}")
        print(f"{'='*50}")
        
        test_folder = f"test_performance_{config['files']}_files"
        
        # Create test files
        if not create_test_pdfs(test_folder, config['files']):
            continue
        
        # Update config for test
        with open('config.json', 'r') as f:
            app_config = json.load(f)
        
        original_folder = app_config.get('pdf_folder', 'workOrderPDF')
        app_config['pdf_folder'] = test_folder
        
        with open('config.json', 'w') as f:
            json.dump(app_config, f, indent=2)
        
        print(f"Test setup complete. Please run the application and process {config['files']} files.")
        print(f"Recommended settings: {config['workers']} concurrent workers")
        print("Press Enter to continue to next test...")
        input()
        
        # Restore original config
        app_config['pdf_folder'] = original_folder
        with open('config.json', 'w') as f:
            json.dump(app_config, f, indent=2)
        
        # Clean up test folder
        if os.path.exists(test_folder):
            shutil.rmtree(test_folder)
    
    print(f"\n{'='*50}")
    print("Performance testing complete!")
    print("Compare processing times and throughput between different configurations.")
    print(f"{'='*50}")

if __name__ == "__main__":
    run_performance_test()