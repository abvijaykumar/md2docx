# -*- coding: utf-8 -*-
# MIT License
#
# Copyright (c) 2025 A B Vijay Kumar
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
import sys
from pathlib import Path

# Import the converter classes
try:
    from md2docx import MarkdownToWordConverter
    from mmd2drawio import MermaidToDrawioConverter
except ImportError as e:
    messagebox.showerror("Import Error", f"Failed to import converters: {e}")
    sys.exit(1)

class MD2DocxUI:
    def __init__(self, root):
        self.root = root
        self.root.title("MD2DOCX & MMD2DRAWIO Converter Suite")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # Initialize converters
        self.md_converter = MarkdownToWordConverter()
        self.mmd_converter = MermaidToDrawioConverter()
        
        # Variables for file/folder paths
        self.selected_md_files = []
        self.selected_mmd_files = []
        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        
        # Control variables
        self.combine_md_files = tk.BooleanVar(value=False)
        self.combine_mmd_files = tk.BooleanVar(value=False)
        
        # Create the UI
        self.create_widgets()
        
        # Center the window
        self.center_window()
    
    def center_window(self):
        """Center the window on the screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_widgets(self):
        """Create all UI widgets"""
        # Create main notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create tabs
        self.create_md2docx_tab()
        self.create_mmd2drawio_tab()
        self.create_batch_tab()
        self.create_log_tab()
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief='sunken', anchor='w')
        status_bar.pack(side='bottom', fill='x', padx=5, pady=2)
    
    def create_md2docx_tab(self):
        """Create the MD2DOCX converter tab"""
        md_frame = ttk.Frame(self.notebook)
        self.notebook.add(md_frame, text="MD to DOCX")
        
        # Title
        title_label = ttk.Label(md_frame, text="Markdown to Word Document Converter", 
                               font=('TkDefaultFont', 12, 'bold'))
        title_label.pack(pady=(10, 20))
        
        # Input selection frame
        input_frame = ttk.LabelFrame(md_frame, text="Input Selection", padding=10)
        input_frame.pack(fill='x', padx=10, pady=5)
        
        # File selection buttons
        btn_frame1 = ttk.Frame(input_frame)
        btn_frame1.pack(fill='x', pady=5)
        
        ttk.Button(btn_frame1, text="Select Markdown Files", 
                  command=self.select_md_files).pack(side='left', padx=(0, 10))
        ttk.Button(btn_frame1, text="Select Folder", 
                  command=self.select_md_folder).pack(side='left', padx=10)
        ttk.Button(btn_frame1, text="Clear Selection", 
                  command=self.clear_md_selection).pack(side='left', padx=10)
        
        # Selected files display
        self.md_files_listbox = tk.Listbox(input_frame, height=6)
        md_scrollbar = ttk.Scrollbar(input_frame, orient="vertical")
        self.md_files_listbox.config(yscrollcommand=md_scrollbar.set)
        md_scrollbar.config(command=self.md_files_listbox.yview)
        
        self.md_files_listbox.pack(side="left", fill="both", expand=True, pady=5)
        md_scrollbar.pack(side="right", fill="y")
        
        # Output settings frame
        output_frame1 = ttk.LabelFrame(md_frame, text="Output Settings", padding=10)
        output_frame1.pack(fill='x', padx=10, pady=5)
        
        # Output folder selection
        folder_frame1 = ttk.Frame(output_frame1)
        folder_frame1.pack(fill='x', pady=5)
        
        ttk.Label(folder_frame1, text="Output Folder:").pack(side='left')
        self.md_output_entry = ttk.Entry(folder_frame1, textvariable=self.output_folder, width=60)
        self.md_output_entry.pack(side='left', fill='x', expand=True, padx=(10, 5))
        ttk.Button(folder_frame1, text="Browse", 
                  command=self.select_output_folder).pack(side='right')
        
        # Options
        options_frame1 = ttk.Frame(output_frame1)
        options_frame1.pack(fill='x', pady=5)
        
        ttk.Checkbutton(options_frame1, text="Combine all files into single document", 
                       variable=self.combine_md_files).pack(anchor='w')
        
        # Convert button
        ttk.Button(md_frame, text="Convert to DOCX", command=self.convert_md_to_docx,
                  style='Accent.TButton').pack(pady=20)
    
    def create_mmd2drawio_tab(self):
        """Create the MMD2DRAWIO converter tab"""
        mmd_frame = ttk.Frame(self.notebook)
        self.notebook.add(mmd_frame, text="MMD to Draw.io")
        
        # Title
        title_label = ttk.Label(mmd_frame, text="Mermaid to Draw.io Diagram Converter", 
                               font=('TkDefaultFont', 12, 'bold'))
        title_label.pack(pady=(10, 20))
        
        # Input selection frame
        input_frame2 = ttk.LabelFrame(mmd_frame, text="Input Selection", padding=10)
        input_frame2.pack(fill='x', padx=10, pady=5)
        
        # File selection buttons
        btn_frame2 = ttk.Frame(input_frame2)
        btn_frame2.pack(fill='x', pady=5)
        
        ttk.Button(btn_frame2, text="Select Mermaid Files (.mmd)", 
                  command=self.select_mmd_files).pack(side='left', padx=(0, 10))
        ttk.Button(btn_frame2, text="Select Folder", 
                  command=self.select_mmd_folder).pack(side='left', padx=10)
        ttk.Button(btn_frame2, text="Clear Selection", 
                  command=self.clear_mmd_selection).pack(side='left', padx=10)
        
        # Selected files display
        self.mmd_files_listbox = tk.Listbox(input_frame2, height=6)
        mmd_scrollbar = ttk.Scrollbar(input_frame2, orient="vertical")
        self.mmd_files_listbox.config(yscrollcommand=mmd_scrollbar.set)
        mmd_scrollbar.config(command=self.mmd_files_listbox.yview)
        
        self.mmd_files_listbox.pack(side="left", fill="both", expand=True, pady=5)
        mmd_scrollbar.pack(side="right", fill="y")
        
        # Output settings frame
        output_frame2 = ttk.LabelFrame(mmd_frame, text="Output Settings", padding=10)
        output_frame2.pack(fill='x', padx=10, pady=5)
        
        # Output folder selection (shared with MD converter)
        folder_frame2 = ttk.Frame(output_frame2)
        folder_frame2.pack(fill='x', pady=5)
        
        ttk.Label(folder_frame2, text="Output Folder:").pack(side='left')
        self.mmd_output_entry = ttk.Entry(folder_frame2, textvariable=self.output_folder, width=60)
        self.mmd_output_entry.pack(side='left', fill='x', expand=True, padx=(10, 5))
        ttk.Button(folder_frame2, text="Browse", 
                  command=self.select_output_folder).pack(side='right')
        
        # Options
        options_frame2 = ttk.Frame(output_frame2)
        options_frame2.pack(fill='x', pady=5)
        
        ttk.Checkbutton(options_frame2, text="Combine all files into single Draw.io file", 
                       variable=self.combine_mmd_files).pack(anchor='w')
        
        # Convert button
        ttk.Button(mmd_frame, text="Convert to Draw.io", command=self.convert_mmd_to_drawio,
                  style='Accent.TButton').pack(pady=20)
    
    def create_batch_tab(self):
        """Create the batch processing tab"""
        batch_frame = ttk.Frame(self.notebook)
        self.notebook.add(batch_frame, text="Batch Processing")
        
        # Title
        title_label = ttk.Label(batch_frame, text="Batch Processing", 
                               font=('TkDefaultFont', 12, 'bold'))
        title_label.pack(pady=(10, 20))
        
        # Input folder frame
        input_batch_frame = ttk.LabelFrame(batch_frame, text="Batch Input Folder", padding=10)
        input_batch_frame.pack(fill='x', padx=10, pady=5)
        
        folder_batch_frame = ttk.Frame(input_batch_frame)
        folder_batch_frame.pack(fill='x', pady=5)
        
        ttk.Label(folder_batch_frame, text="Input Folder:").pack(side='left')
        self.batch_input_entry = ttk.Entry(folder_batch_frame, textvariable=self.input_folder, width=60)
        self.batch_input_entry.pack(side='left', fill='x', expand=True, padx=(10, 5))
        ttk.Button(folder_batch_frame, text="Browse", 
                  command=self.select_batch_input_folder).pack(side='right')
        
        # Output folder frame
        output_batch_frame = ttk.LabelFrame(batch_frame, text="Batch Output Folder", padding=10)
        output_batch_frame.pack(fill='x', padx=10, pady=5)
        
        folder_batch_out_frame = ttk.Frame(output_batch_frame)
        folder_batch_out_frame.pack(fill='x', pady=5)
        
        ttk.Label(folder_batch_out_frame, text="Output Folder:").pack(side='left')
        self.batch_output_entry = ttk.Entry(folder_batch_out_frame, textvariable=self.output_folder, width=60)
        self.batch_output_entry.pack(side='left', fill='x', expand=True, padx=(10, 5))
        ttk.Button(folder_batch_out_frame, text="Browse", 
                  command=self.select_output_folder).pack(side='right')
        
        # Batch options frame
        batch_options_frame = ttk.LabelFrame(batch_frame, text="Batch Options", padding=10)
        batch_options_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Checkbutton(batch_options_frame, text="Process Markdown files (.md → .docx)", 
                       variable=self.combine_md_files).pack(anchor='w', pady=2)
        ttk.Checkbutton(batch_options_frame, text="Process Mermaid files (.mmd → .drawio)", 
                       variable=self.combine_mmd_files).pack(anchor='w', pady=2)
        
        # Batch convert buttons
        button_batch_frame = ttk.Frame(batch_frame)
        button_batch_frame.pack(pady=20)
        
        ttk.Button(button_batch_frame, text="Process All Files", 
                  command=self.batch_process_all,
                  style='Accent.TButton').pack(side='left', padx=10)
        
        ttk.Button(button_batch_frame, text="Analyze Folder", 
                  command=self.analyze_batch_folder).pack(side='left', padx=10)
    
    def create_log_tab(self):
        """Create the log/output tab"""
        log_frame = ttk.Frame(self.notebook)
        self.notebook.add(log_frame, text="Processing Log")
        
        # Title
        title_label = ttk.Label(log_frame, text="Processing Log & Output", 
                               font=('TkDefaultFont', 12, 'bold'))
        title_label.pack(pady=(10, 10))
        
        # Log text area
        self.log_text = scrolledtext.ScrolledText(log_frame, height=25, width=100)
        self.log_text.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Log control buttons
        log_btn_frame = ttk.Frame(log_frame)
        log_btn_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(log_btn_frame, text="Clear Log", 
                  command=self.clear_log).pack(side='left')
        ttk.Button(log_btn_frame, text="Save Log", 
                  command=self.save_log).pack(side='left', padx=10)
    
    # File/folder selection methods
    def select_md_files(self):
        """Select markdown files"""
        files = filedialog.askopenfilenames(
            title="Select Markdown Files",
            filetypes=[("Markdown files", "*.md"), ("All files", "*.*")]
        )
        if files:
            self.selected_md_files = list(files)
            self.update_md_listbox()
            self.log_message(f"Selected {len(files)} markdown files")
    
    def select_md_folder(self):
        """Select folder containing markdown files"""
        folder = filedialog.askdirectory(title="Select Folder with Markdown Files")
        if folder:
            md_files = []
            for ext in ['*.md']:
                md_files.extend(Path(folder).glob(ext))
            self.selected_md_files = [str(f) for f in sorted(md_files)]
            self.update_md_listbox()
            self.log_message(f"Found {len(self.selected_md_files)} markdown files in {folder}")
    
    def clear_md_selection(self):
        """Clear markdown file selection"""
        self.selected_md_files = []
        self.update_md_listbox()
        self.log_message("Cleared markdown file selection")
    
    def select_mmd_files(self):
        """Select mermaid files"""
        files = filedialog.askopenfilenames(
            title="Select Mermaid Files",
            filetypes=[("Mermaid files", "*.mmd"), ("All files", "*.*")]
        )
        if files:
            self.selected_mmd_files = list(files)
            self.update_mmd_listbox()
            self.log_message(f"Selected {len(files)} mermaid files")
    
    def select_mmd_folder(self):
        """Select folder containing mermaid files"""
        folder = filedialog.askdirectory(title="Select Folder with Mermaid Files")
        if folder:
            mmd_files = []
            for ext in ['*.mmd']:
                mmd_files.extend(Path(folder).glob(ext))
            self.selected_mmd_files = [str(f) for f in sorted(mmd_files)]
            self.update_mmd_listbox()
            self.log_message(f"Found {len(self.selected_mmd_files)} mermaid files in {folder}")
    
    def clear_mmd_selection(self):
        """Clear mermaid file selection"""
        self.selected_mmd_files = []
        self.update_mmd_listbox()
        self.log_message("Cleared mermaid file selection")
    
    def select_output_folder(self):
        """Select output folder"""
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder.set(folder)
            self.log_message(f"Set output folder to: {folder}")
    
    def select_batch_input_folder(self):
        """Select batch input folder"""
        folder = filedialog.askdirectory(title="Select Batch Input Folder")
        if folder:
            self.input_folder.set(folder)
            self.log_message(f"Set batch input folder to: {folder}")
    
    # UI update methods
    def update_md_listbox(self):
        """Update markdown files listbox"""
        self.md_files_listbox.delete(0, tk.END)
        for file in self.selected_md_files:
            self.md_files_listbox.insert(tk.END, os.path.basename(file))
    
    def update_mmd_listbox(self):
        """Update mermaid files listbox"""
        self.mmd_files_listbox.delete(0, tk.END)
        for file in self.selected_mmd_files:
            self.mmd_files_listbox.insert(tk.END, os.path.basename(file))
    
    # Logging methods
    def log_message(self, message):
        """Add message to log"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def clear_log(self):
        """Clear the log"""
        self.log_text.delete(1.0, tk.END)
        self.status_var.set("Log cleared")
    
    def save_log(self):
        """Save log to file"""
        filename = filedialog.asksaveasfilename(
            title="Save Log",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.get(1.0, tk.END))
                self.log_message(f"Log saved to: {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save log: {str(e)}")
    
    # Conversion methods
    def convert_md_to_docx(self):
        """Convert markdown files to docx"""
        if not self.selected_md_files:
            messagebox.showwarning("Warning", "Please select markdown files first")
            return
        
        if not self.output_folder.get():
            messagebox.showwarning("Warning", "Please select an output folder")
            return
        
        # Create a custom converter with logging
        self.md_converter.log_callback = self.log_message
        
        # Run conversion in separate thread to prevent UI freezing
        thread = threading.Thread(target=self._convert_md_to_docx_thread)
        thread.daemon = True
        thread.start()
    
    def _convert_md_to_docx_thread(self):
        """Thread function for MD to DOCX conversion"""
        try:
            output_dir = self.output_folder.get()
            os.makedirs(output_dir, exist_ok=True)
            
            self.log_message("Starting markdown to DOCX conversion...")
            
            if self.combine_md_files.get():
                # Combine all files into single document
                output_file = os.path.join(output_dir, "combined.docx")
                self.log_message("Converting all files to single document...")
                self.md_converter.convert_combined(self.selected_md_files, output_file)
                self.log_message(f"✓ Created combined document: {output_file}")
            else:
                # Convert each file separately
                for i, md_file in enumerate(self.selected_md_files, 1):
                    filename = os.path.splitext(os.path.basename(md_file))[0]
                    output_file = os.path.join(output_dir, f"{filename}.docx")
                    
                    self.log_message(f"Converting {i}/{len(self.selected_md_files)}: {os.path.basename(md_file)}")
                    self.md_converter.convert_file(md_file, output_file)
                    self.log_message(f"✓ Created: {output_file}")
            
            self.log_message("✓ Markdown to DOCX conversion completed successfully!")
            
        except Exception as e:
            self.log_message(f"✗ Error during conversion: {str(e)}")
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")
    
    def convert_mmd_to_drawio(self):
        """Convert mermaid files to draw.io"""
        if not self.selected_mmd_files:
            messagebox.showwarning("Warning", "Please select mermaid files first")
            return
        
        if not self.output_folder.get():
            messagebox.showwarning("Warning", "Please select an output folder")
            return
        
        # Create a custom converter with logging
        self.mmd_converter.log_callback = self.log_message
        
        # Run conversion in separate thread to prevent UI freezing
        thread = threading.Thread(target=self._convert_mmd_to_drawio_thread)
        thread.daemon = True
        thread.start()
    
    def _convert_mmd_to_drawio_thread(self):
        """Thread function for MMD to Draw.io conversion"""
        try:
            output_dir = self.output_folder.get()
            os.makedirs(output_dir, exist_ok=True)
            
            self.log_message("Starting mermaid to Draw.io conversion...")
            
            if self.combine_mmd_files.get():
                # Combine all files into single Draw.io file
                output_file = os.path.join(output_dir, "combined.drawio")
                self.log_message("Converting all files to single Draw.io file...")
                self.mmd_converter.convert_multiple_files(self.selected_mmd_files, output_file)
                self.log_message(f"✓ Created combined diagram: {output_file}")
            else:
                # Convert each file separately
                for i, mmd_file in enumerate(self.selected_mmd_files, 1):
                    filename = os.path.splitext(os.path.basename(mmd_file))[0]
                    output_file = os.path.join(output_dir, f"{filename}.drawio")
                    
                    self.log_message(f"Converting {i}/{len(self.selected_mmd_files)}: {os.path.basename(mmd_file)}")
                    self.mmd_converter.convert_single_file(mmd_file, output_file)
                    self.log_message(f"✓ Created: {output_file}")
            
            self.log_message("✓ Mermaid to Draw.io conversion completed successfully!")
            
        except Exception as e:
            self.log_message(f"✗ Error during conversion: {str(e)}")
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")
    
    def analyze_batch_folder(self):
        """Analyze batch folder for available files"""
        if not self.input_folder.get():
            messagebox.showwarning("Warning", "Please select a batch input folder")
            return
        
        folder = self.input_folder.get()
        if not os.path.exists(folder):
            messagebox.showerror("Error", "Selected folder does not exist")
            return
        
        self.log_message(f"Analyzing folder: {folder}")
        
        # Find markdown files
        md_files = list(Path(folder).glob("*.md"))
        self.log_message(f"Found {len(md_files)} markdown files:")
        for f in md_files:
            self.log_message(f"  - {f.name}")
        
        # Find mermaid files
        mmd_files = list(Path(folder).glob("*.mmd"))
        self.log_message(f"Found {len(mmd_files)} mermaid files:")
        for f in mmd_files:
            self.log_message(f"  - {f.name}")
        
        # Update selections
        self.selected_md_files = [str(f) for f in sorted(md_files)]
        self.selected_mmd_files = [str(f) for f in sorted(mmd_files)]
        self.update_md_listbox()
        self.update_mmd_listbox()
        
        if not md_files and not mmd_files:
            self.log_message("No markdown or mermaid files found in the selected folder")
        else:
            self.log_message("✓ Folder analysis complete. Files loaded in respective tabs.")
    
    def batch_process_all(self):
        """Process all files in batch mode"""
        if not self.input_folder.get():
            messagebox.showwarning("Warning", "Please select a batch input folder")
            return
        
        if not self.output_folder.get():
            messagebox.showwarning("Warning", "Please select an output folder")
            return
        
        # Analyze folder first
        self.analyze_batch_folder()
        
        # Run batch processing in separate thread
        thread = threading.Thread(target=self._batch_process_thread)
        thread.daemon = True
        thread.start()
    
    def _batch_process_thread(self):
        """Thread function for batch processing"""
        try:
            self.log_message("Starting batch processing...")
            
            # Process markdown files if any exist
            if self.selected_md_files:
                self.log_message("Processing markdown files...")
                self._convert_md_to_docx_thread()
            
            # Process mermaid files if any exist
            if self.selected_mmd_files:
                self.log_message("Processing mermaid files...")
                self._convert_mmd_to_drawio_thread()
            
            if not self.selected_md_files and not self.selected_mmd_files:
                self.log_message("No files to process")
            else:
                self.log_message("✓ Batch processing completed successfully!")
                
        except Exception as e:
            self.log_message(f"✗ Error during batch processing: {str(e)}")
            messagebox.showerror("Error", f"Batch processing failed: {str(e)}")

def main():
    """Main application entry point"""
    # Create the main window
    root = tk.Tk()
    
    # Apply a modern theme if available
    try:
        style = ttk.Style()
        style.theme_use('clam')  # Use clam theme for better appearance
    except:
        pass
    
    # Create the application
    app = MD2DocxUI(root)
    
    # Add some welcome message to log
    app.log_message("Welcome to MD2DOCX & MMD2DRAWIO Converter Suite!")
    app.log_message("Select your files or folders and choose your conversion options.")
    app.log_message("Check the tabs above for different conversion modes.")
    app.log_message("-" * 50)
    
    # Start the GUI event loop
    root.mainloop()

if __name__ == "__main__":
    main()