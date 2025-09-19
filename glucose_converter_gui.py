#!/usr/bin/env python3
"""
Glucose Data Converter GUI - Drag and Drop Interface
Author: Your Assistant
Purpose: GUI version with drag-and-drop support for easy CSV conversion
"""

import sys
import os
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
from typing import Optional

# Import the converter from main script
from glucose_converter import GlucoseConverter, find_latest_csv


class GlucoseConverterGUI:
    """GUI application for glucose data conversion"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Glucose Data Converter")
        self.root.geometry("700x600")
        
        # Set icon if available
        try:
            self.root.iconbitmap('icon.ico')
        except:
            pass
        
        # Initialize converter
        self.converter = GlucoseConverter()
        self.current_file = None
        
        # Setup UI
        self.setup_ui()
        
        # Enable drag and drop
        self.setup_drag_drop()
        
    def setup_ui(self):
        """Create the user interface"""
        
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        # Title
        title_label = ttk.Label(
            main_frame, 
            text="Glucose Data Converter", 
            font=('Arial', 16, 'bold')
        )
        title_label.grid(row=0, column=0, pady=10)
        
        # Drop zone frame
        self.drop_frame = tk.Frame(
            main_frame, 
            bg='#f0f0f0', 
            relief=tk.SUNKEN, 
            bd=2,
            height=150
        )
        self.drop_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=20, pady=10)
        self.drop_frame.grid_propagate(False)
        
        # Drop zone label
        self.drop_label = tk.Label(
            self.drop_frame,
            text="📁 Drag and Drop CSV File Here\nor Click to Browse",
            font=('Arial', 12),
            bg='#f0f0f0',
            fg='#666',
            cursor="hand2"
        )
        self.drop_label.pack(expand=True, fill=tk.BOTH)
        
        # Bind click event to browse
        self.drop_label.bind("<Button-1>", lambda e: self.browse_file())
        
        # File info frame
        info_frame = ttk.LabelFrame(main_frame, text="File Information", padding="10")
        info_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), padx=20, pady=10)
        info_frame.columnconfigure(1, weight=1)
        
        # File path display
        ttk.Label(info_frame, text="Selected File:").grid(row=0, column=0, sticky=tk.W)
        self.file_label = ttk.Label(info_frame, text="No file selected", foreground="gray")
        self.file_label.grid(row=0, column=1, sticky=tk.W, padx=10)
        
        # Settings frame
        settings_frame = ttk.LabelFrame(main_frame, text="Settings", padding="10")
        settings_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), padx=20, pady=10)
        settings_frame.columnconfigure(1, weight=1)
        
        # Output folder
        ttk.Label(settings_frame, text="Output Folder:").grid(row=0, column=0, sticky=tk.W)
        
        output_frame = ttk.Frame(settings_frame)
        output_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=10)
        output_frame.columnconfigure(0, weight=1)
        
        self.output_var = tk.StringVar(value="Same as input file")
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_var, state="readonly")
        self.output_entry.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        ttk.Button(
            output_frame, 
            text="Browse", 
            command=self.select_output_folder,
            width=10
        ).grid(row=0, column=1, padx=(5, 0))
        
        # Threshold settings
        ttk.Label(settings_frame, text="Low Threshold (mmol/L):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.low_threshold_var = tk.DoubleVar(value=4.0)
        low_spinbox = ttk.Spinbox(
            settings_frame, 
            from_=1.0, 
            to=10.0, 
            increment=0.1,
            textvariable=self.low_threshold_var,
            width=10
        )
        low_spinbox.grid(row=1, column=1, sticky=tk.W, padx=10, pady=5)
        
        ttk.Label(settings_frame, text="High Threshold (mmol/L):").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.high_threshold_var = tk.DoubleVar(value=11.9)
        high_spinbox = ttk.Spinbox(
            settings_frame, 
            from_=8.0, 
            to=20.0, 
            increment=0.1,
            textvariable=self.high_threshold_var,
            width=10
        )
        high_spinbox.grid(row=2, column=1, sticky=tk.W, padx=10, pady=5)
        
        ttk.Label(settings_frame, text="Very High Threshold (mmol/L):").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.very_high_threshold_var = tk.DoubleVar(value=17.9)
        very_high_spinbox = ttk.Spinbox(
            settings_frame, 
            from_=15.0, 
            to=30.0, 
            increment=0.1,
            textvariable=self.very_high_threshold_var,
            width=10
        )
        very_high_spinbox.grid(row=3, column=1, sticky=tk.W, padx=10, pady=5)
        
        # Auto-open checkbox
        self.auto_open_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            settings_frame, 
            text="Open file after conversion", 
            variable=self.auto_open_var
        ).grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, pady=20)
        
        # Convert button
        self.convert_btn = ttk.Button(
            button_frame,
            text="Convert to XLSX",
            command=self.convert_file,
            state="disabled"
        )
        self.convert_btn.pack(side=tk.LEFT, padx=5)
        
        # Auto-detect button
        ttk.Button(
            button_frame,
            text="Auto-detect Latest CSV",
            command=self.auto_detect_csv
        ).pack(side=tk.LEFT, padx=5)
        
        # Status bar
        self.status_frame = ttk.Frame(main_frame)
        self.status_frame.grid(row=5, column=0, sticky=(tk.W, tk.E), padx=20)
        
        self.status_label = ttk.Label(self.status_frame, text="Ready", relief=tk.SUNKEN)
        self.status_label.pack(fill=tk.X)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(
            main_frame,
            mode='indeterminate',
            length=200
        )
        
    def setup_drag_drop(self):
        """Setup drag and drop functionality"""
        # Note: Full drag-drop requires tkinterdnd2, but we'll use file dialog as fallback
        try:
            # Try to import tkinterdnd2 for drag-drop support
            from tkinterdnd2 import DND_FILES, TkinterDnD
            
            # Recreate root with DnD support
            self.root.destroy()
            self.root = TkinterDnD.Tk()
            self.__init__(self.root)
            
            # Register drop target
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind('<<Drop>>', self.drop_file)
            
        except ImportError:
            # tkinterdnd2 not available, use file dialog only
            pass
    
    def drop_file(self, event):
        """Handle dropped file"""
        files = self.root.tk.splitlist(event.data)
        if files:
            file_path = files[0]
            if file_path.lower().endswith('.csv'):
                self.load_file(file_path)
            else:
                messagebox.showerror("Error", "Please drop a CSV file")
    
    def browse_file(self):
        """Open file browser to select CSV"""
        file_path = filedialog.askopenfilename(
            title="Select Contour CSV File",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
            initialdir=str(Path.home() / "Downloads")
        )
        
        if file_path:
            self.load_file(file_path)
    
    def load_file(self, file_path):
        """Load selected CSV file"""
        self.current_file = file_path
        self.file_label.config(text=Path(file_path).name, foreground="black")
        self.convert_btn.config(state="normal")
        self.update_status(f"Loaded: {Path(file_path).name}")
        
        # Update drop zone appearance
        self.drop_label.config(
            text=f"✅ {Path(file_path).name}\n\nDrop another file to replace",
            bg='#e6ffe6'
        )
    
    def select_output_folder(self):
        """Select output folder"""
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_var.set(folder)
            self.converter.config['output_folder'] = folder
    
    def auto_detect_csv(self):
        """Auto-detect latest CSV in Downloads"""
        self.update_status("Searching for latest CSV...")
        
        downloads = Path.home() / "Downloads"
        latest_csv = find_latest_csv(str(downloads))
        
        if latest_csv:
            self.load_file(latest_csv)
            messagebox.showinfo("Success", f"Found: {Path(latest_csv).name}")
        else:
            messagebox.showwarning("Not Found", "No Contour CSV files found in Downloads folder")
            self.update_status("No CSV files found")
    
    def convert_file(self):
        """Convert the loaded CSV file"""
        if not self.current_file:
            messagebox.showerror("Error", "No file selected")
            return
        
        # Update converter settings
        self.converter.config['low_threshold'] = self.low_threshold_var.get()
        self.converter.config['high_threshold'] = self.high_threshold_var.get()
        self.converter.config['very_high_threshold'] = self.very_high_threshold_var.get()
        self.converter.config['auto_open'] = self.auto_open_var.get()
        
        # Disable button during conversion
        self.convert_btn.config(state="disabled")
        
        # Show progress
        self.progress_bar.grid(row=6, column=0, pady=10)
        self.progress_bar.start(10)
        self.update_status("Converting...")
        
        # Run conversion in thread to avoid freezing UI
        thread = threading.Thread(target=self.run_conversion)
        thread.start()
    
    def run_conversion(self):
        """Run the conversion process"""
        try:
            output_file = self.converter.convert(self.current_file)
            
            # Update UI in main thread
            self.root.after(0, self.conversion_success, output_file)
            
        except Exception as e:
            # Update UI in main thread
            self.root.after(0, self.conversion_error, str(e))
    
    def conversion_success(self, output_file):
        """Handle successful conversion"""
        self.progress_bar.stop()
        self.progress_bar.grid_remove()
        self.convert_btn.config(state="normal")
        
        self.update_status(f"Success! Saved to: {Path(output_file).name}")
        
        messagebox.showinfo(
            "Success", 
            f"Conversion complete!\n\nFile saved to:\n{output_file}"
        )
    
    def conversion_error(self, error_msg):
        """Handle conversion error"""
        self.progress_bar.stop()
        self.progress_bar.grid_remove()
        self.convert_btn.config(state="normal")
        
        self.update_status("Conversion failed")
        messagebox.showerror("Error", f"Conversion failed:\n{error_msg}")
    
    def update_status(self, message):
        """Update status bar"""
        self.status_label.config(text=message)


def main():
    """Main entry point for GUI application"""
    
    # Check if tkinterdnd2 is available for better drag-drop
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except ImportError:
        # Fall back to standard tkinter
        root = tk.Tk()
        print("Note: Install tkinterdnd2 for drag-and-drop support:")
        print("  pip install tkinterdnd2")
    
    app = GlucoseConverterGUI(root)
    
    # Center window on screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()


if __name__ == '__main__':
    main()