#!/usr/bin/env python3
"""
Setup script for Work Order PDF Extractor
"""

import os
import sys
import subprocess

def check_python_version():
    """Check if Python version is compatible"""
    if sys.version_info < (3, 8):
        print("❌ Python 3.8 or higher is required")
        sys.exit(1)
    print(f"✅ Python {sys.version.split()[0]} detected")

def install_system_dependencies():
    """Install system dependencies"""
    print("📦 Installing system dependencies...")
    
    system = os.uname().sysname.lower()
    
    if system == "darwin":  # macOS
        print("🍎 Detected macOS - installing poppler via Homebrew...")
        try:
            subprocess.run(["brew", "install", "poppler"], check=True)
            print("✅ Poppler installed successfully")
        except subprocess.CalledProcessError:
            print("⚠️  Failed to install poppler. Please install Homebrew and try again.")
            print("   Visit: https://brew.sh/")
        except FileNotFoundError:
            print("⚠️  Homebrew not found. Please install Homebrew first:")
            print("   /bin/bash -c \"$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)\"")
    
    elif system == "linux":
        print("🐧 Detected Linux - installing poppler...")
        try:
            # Try apt first (Ubuntu/Debian)
            subprocess.run(["sudo", "apt-get", "update"], check=True)
            subprocess.run(["sudo", "apt-get", "install", "-y", "poppler-utils"], check=True)
            print("✅ Poppler installed successfully")
        except subprocess.CalledProcessError:
            try:
                # Try yum (CentOS/RHEL)
                subprocess.run(["sudo", "yum", "install", "-y", "poppler-utils"], check=True)
                print("✅ Poppler installed successfully")
            except subprocess.CalledProcessError:
                print("⚠️  Failed to install poppler. Please install manually:")
                print("   Ubuntu/Debian: sudo apt-get install poppler-utils")
                print("   CentOS/RHEL: sudo yum install poppler-utils")
    
    else:
        print("⚠️  Unsupported system. Please install poppler manually:")
        print("   Windows: Download from https://blog.alivate.com.au/poppler-windows/")

def install_python_dependencies():
    """Install Python dependencies"""
    print("🐍 Installing Python dependencies...")
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"], check=True)
        print("✅ Python dependencies installed successfully")
    except subprocess.CalledProcessError:
        print("❌ Failed to install Python dependencies")
        sys.exit(1)

def create_directories():
    """Create necessary directories"""
    print("📁 Creating necessary directories...")
    directories = ["workOrderPDF", "not_match"]
    for directory in directories:
        os.makedirs(directory, exist_ok=True)
        print(f"✅ Created directory: {directory}")

def main():
    """Main setup function"""
    print("🚀 Setting up Work Order PDF Extractor...\n")
    
    check_python_version()
    install_system_dependencies()
    install_python_dependencies()
    create_directories()
    
    print("\n🎉 Setup complete!")
    print("\n📋 Next steps:")
    print("1. Add your PDF files to the 'workOrderPDF' folder")
    print("2. Configure your OpenAI API key in the application")
    print("3. Run: python3 work_order_extractor.py")
    print("\n📖 For more information, see README.md")

if __name__ == "__main__":
    main()