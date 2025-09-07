#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Simple launcher script for the MD2DOCX & MMD2DRAWIO UI application
"""

import sys
import os

# Add the current directory to the Python path so we can import our modules
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from ui_app import main
    
    if __name__ == "__main__":
        main()
        
except ImportError as e:
    print(f"Error importing modules: {e}")
    print("Make sure all required dependencies are installed:")
    print("pip install python-docx markdown beautifulsoup4 playwright flask")
    sys.exit(1)
except Exception as e:
    print(f"Error starting application: {e}")
    sys.exit(1)