#!/usr/bin/env python3
"""
Simplified Glucose Converter GUI with Date Filtering
Author: Your Assistant
Purpose: Clean GUI for glucose data conversion
"""

import sys
import os
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from datetime import datetime, timedelta
from typing import Optional
import platform

# Import simplified converter
from glucose_converter_simplified import (
    SimplifiedGlucoseConverter, 
    find_latest_csv,
    get_downloads_folder
)


class SimplifiedGlucoseGUI:
    """Simplified GUI with date filtering support"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Glucose Data Converter")
        self.root.geometry("650x600")
        
        # Initialize converter
        self.converter = SimplifiedGlucoseConverter()
        self.current_file = None
        
        # Setup UI
        self.setup_ui()
        
        # Enable drag and drop (if tkinterdnd2 available)
        self.setup_drag_drop()
        
    def setup_ui(self):
        """Create the simplified user interface"""
        
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(
            main_frame, 
            text="Glucose Data Converter", 
            font=('Arial', 16, 'bold')
        )
        title_label.pack(pady=10)
        
        # Drop zone
        self.drop_frame = tk.Frame(
            main_frame,
            bg='#f0f0f0',
            relief=tk.SUNKEN,
            bd=2,
            height=100
        )
        self.drop_frame.pack(fill=tk.X, pady=10)
        self.drop_frame.pack_propagate(False)
        
        self.drop_label = tk.Label(
            self.drop_frame,
            text="📁 Drag and Drop CSV File Here\nor Click to Browse",
            font=('Arial', 11),
            bg='#f0f0f0',
            fg='#666',
            cursor="hand2"
        )
        self.drop_label.pack(expand=True, fill=tk.BOTH)
        self.drop_label.bind("<Button-1>", lambda e: self.browse_file())
        
        # File info
        info_frame = ttk.LabelFrame(main_frame, text="File Information", padding="10")
        info_frame.pack(fill=tk.X, pady=10)
        
        self.file_label = ttk.Label(info_frame, text="No file selected", foreground="gray")
        self.file_label.pack(anchor=tk.W)
        
        self.last_export_label = ttk.Label(info_frame, text="", foreground="blue")
        self.last_export_label.pack(anchor=tk.W)
        
        # Date filtering
        date_frame = ttk.LabelFrame(main_frame, text="Date Filter Options", padding="10")
        date_frame.pack(fill=tk.X, pady=10)
        
        # Filter mode selection
        self.filter_mode = tk.StringVar(value="all")
        
        ttk.Radiobutton(
            date_frame, 
            text="Export all data", 
            variable=self.filter_mode, 
            value="all",
            command=self.update_filter_ui
        ).grid(row=0, column=0, sticky=tk.W, pady=2)
        
        ttk.Radiobutton(
            date_frame, 
            text="Incremental export (only new data since last export)", 
            variable=self.filter_mode, 
            value="incremental",
            command=self.update_filter_ui
        ).grid(row=1, column=0, sticky=tk.W, pady=2)
        
        # Last N days option
        days_frame = ttk.Frame(date_frame)
        days_frame.grid(row=2, column=0, sticky=tk.W, pady=2)
        
        ttk.Radiobutton(
            days_frame, 
            text="Export last", 
            variable=self.filter_mode, 
            value="days",
            command=self.update_filter_ui
        ).pack(side=tk.LEFT)
        
        self.days_var = tk.IntVar(value=30)
        self.days_spin = ttk.Spinbox(
            days_frame, 
            from_=1, 
            to=365, 
            width=5,
            textvariable=self.days_var
        )
        self.days_spin.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(days_frame, text="days").pack(side=tk.LEFT)
        
        # Custom range option
        ttk.Radiobutton(
            date_frame, 
            text="Custom date range", 
            variable=self.filter_mode, 
            value="custom",
            command=self.update_filter_ui
        ).grid(row=3, column=0, sticky=tk.W, pady=2)
        
        # Custom date inputs
        self.custom_frame = ttk.Frame(date_frame)
        self.custom_frame.grid(row=4, column=0, sticky=tk.W, padx=20, pady=5)
        
        ttk.Label(self.custom_frame, text="From:").grid(row=0, column=0)
        self.start_date_var = tk.StringVar(value=(datetime.now() - timedelta(days=30)).strftime('%d.%m.%Y'))
        self.start_date_entry = ttk.Entry(self.custom_frame, textvariable=self.start_date_var, width=12)
        self.start_date_entry.grid(row=0, column=1, padx=5)
        
        ttk.Label(self.custom_frame, text="To:").grid(row=0, column=2, padx=(10, 0))
        self.end_date_var = tk.StringVar(value=datetime.now().strftime('%d.%m.%Y'))
        self.end_date_entry = ttk.Entry(self.custom_frame, textvariable=self.end_date_var, width=12)
        self.end_date_entry.grid(row=0, column=3, padx=5)
        
        # Settings
        settings_frame = ttk.LabelFrame(main_frame, text="Settings", padding="10")
        settings_frame.pack(fill=tk.X, pady=10)
        
        # Thresholds in one row
        threshold_frame = ttk.Frame(settings_frame)
        threshold_frame.pack(fill=tk.X)
        
        ttk.Label(threshold_frame, text="Thresholds (mmol/L):").pack(side=tk.LEFT)
        
        ttk.Label(threshold_frame, text="Low <").pack(side=tk.LEFT, padx=(10, 2))
        self.low_threshold_var = tk.DoubleVar(value=4.0)
        ttk.Spinbox(
            threshold_frame, 
            from_=1.0, 
            to=10.0, 
            increment=0.1,
            textvariable=self.low_threshold_var,
            width=5
        ).pack(side=tk.LEFT)
        
        ttk.Label(threshold_frame, text="  High >").pack(side=tk.LEFT, padx=(10, 2))
        self.high_threshold_var = tk.DoubleVar(value=11.9)
        ttk.Spinbox(
            threshold_frame, 
            from_=8.0, 
            to=20.0, 
            increment=0.1,
            textvariable=self.high_threshold_var,
            width=5
        ).pack(side=tk.LEFT)
        
        ttk.Label(threshold_frame, text="  Very High >").pack(side=tk.LEFT, padx=(10, 2))
        self.very_high_threshold_var = tk.DoubleVar(value=17.9)
        ttk.Spinbox(
            threshold_frame, 
            from_=15.0, 
            to=30.0, 
            increment=0.1,
            textvariable=self.very_high_threshold_var,
            width=5
        ).pack(side=tk.LEFT)
        
        # Auto-open checkbox
        self.auto_open_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            settings_frame,
            text="Open file after conversion",
            variable=self.auto_open_var
        ).pack(anchor=tk.W, pady=5)
        
        # Action buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        self.convert_btn = ttk.Button(
            button_frame,
            text="Convert to XLSX",
            command=self.convert_file,
            state="disabled"
        )
        self.convert_btn.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame,
            text="Auto-detect Latest CSV",
            command=self.auto_detect_csv
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame,
            text="Upload Template",
            command=self.upload_template
        ).pack(side=tk.LEFT, padx=5)
        
        # Status bar
        self.status_frame = ttk.Frame(main_frame)
        self.status_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.status_label = ttk.Label(self.status_frame, text="Ready", relief=tk.SUNKEN)
        self.status_label.pack(fill=tk.X)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(
            self.status_frame,
            mode='indeterminate',
            length=200
        )
        
        # Initial UI state
        self.update_filter_ui()
    
    def update_filter_ui(self):
        """Enable/disable UI elements based on filter mode"""
        mode = self.filter_mode.get()
        
        # Disable/enable custom date inputs
        if mode == "custom":
            self.start_date_entry.config(state="normal")
            self.end_date_entry.config(state="normal")
        else:
            self.start_date_entry.config(state="disabled")
            self.end_date_entry.config(state="disabled")
        
        # Disable/enable days spinner
        if mode == "days":
            self.days_spin.config(state="normal")
        else:
            self.days_spin.config(state="disabled")
    
    def setup_drag_drop(self):
        """Setup drag and drop functionality"""
        try:
            from tkinterdnd2 import DND_FILES
            
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind('<<Drop>>', self.drop_file)
        except ImportError:
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
        """Open file browser"""
        file_path = filedialog.askopenfilename(
            title="Select Contour CSV File",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
            initialdir=str(get_downloads_folder())
        )
        
        if file_path:
            self.load_file(file_path)
    
    def load_file(self, file_path):
        """Load selected CSV file"""
        self.current_file = file_path
        self.file_label.config(text=f"File: {Path(file_path).name}", foreground="black")
        self.convert_btn.config(state="normal")
        self.update_status(f"Loaded: {Path(file_path).name}")
        
        self.drop_label.config(
            text=f"✅ {Path(file_path).name}\n\nDrop another file to replace",
            bg='#e6ffe6'
        )
        
        # Check if incremental export available
        last_export = self.converter.export_tracker.get_last_export_date(file_path)
        if last_export:
            self.last_export_label.config(
                text=f"Last export: {last_export.strftime('%d.%m.%Y %H:%M')}"
            )
        else:
            self.last_export_label.config(text="No previous export found")
    
    def auto_detect_csv(self):
        """Auto-detect latest CSV"""
        self.update_status("Searching for latest CSV...")
        
        downloads = get_downloads_folder()
        latest_csv = find_latest_csv(str(downloads))
        
        if latest_csv:
            self.load_file(latest_csv)
            self.update_status(f"Found: {Path(latest_csv).name}")
        else:
            messagebox.showwarning("Not Found", "No Contour CSV files found in Downloads folder")
            self.update_status("No CSV files found")
    
    def upload_template(self):
        """Upload a template file"""
        file_path = filedialog.askopenfilename(
            title="Select Template XLSX File",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        
        if file_path:
            if self.converter.save_template(file_path):
                messagebox.showinfo("Success", "Template uploaded successfully!\n\nIt will be used for all future conversions.")
            else:
                messagebox.showerror("Error", "Failed to save template")
    
    def convert_file(self):
        """Convert the loaded CSV file"""
        if not self.current_file:
            messagebox.showerror("Error", "No file selected")
            return
        
        # Get date filter settings
        start_date = None
        end_date = None
        incremental = False
        
        filter_mode = self.filter_mode.get()
        
        if filter_mode == "incremental":
            incremental = True
        elif filter_mode == "days":
            days = self.days_var.get()
            end_date = datetime.now()
            start_date = end_date - timedelta(days=days)
        elif filter_mode == "custom":
            try:
                start_date = datetime.strptime(self.start_date_var.get(), '%d.%m.%Y')
                end_date = datetime.strptime(self.end_date_var.get(), '%d.%m.%Y')
            except ValueError:
                messagebox.showerror("Error", "Invalid date format. Use DD.MM.YYYY")
                return
        
        # Update converter settings
        self.converter.config['low_threshold'] = self.low_threshold_var.get()
        self.converter.config['high_threshold'] = self.high_threshold_var.get()
        self.converter.config['very_high_threshold'] = self.very_high_threshold_var.get()
        self.converter.config['auto_open'] = self.auto_open_var.get()
        
        # Disable button and show progress
        self.convert_btn.config(state="disabled")
        self.progress_bar.pack(fill=tk.X, pady=5)
        self.progress_bar.start(10)
        self.update_status("Converting...")
        
        # Run in thread
        thread = threading.Thread(
            target=self.run_conversion,
            args=(start_date, end_date, incremental)
        )
        thread.start()
    
    def run_conversion(self, start_date, end_date, incremental):
        """Run the conversion process"""
        try:
            output_file = self.converter.convert(
                self.current_file,
                start_date=start_date,
                end_date=end_date,
                incremental=incremental
            )
            
            if output_file:
                self.root.after(0, self.conversion_success, output_file)
            else:
                self.root.after(0, self.conversion_error, "No data in specified range")
            
        except Exception as e:
            self.root.after(0, self.conversion_error, str(e))
    
    def conversion_success(self, output_file):
        """Handle successful conversion"""
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        self.convert_btn.config(state="normal")
        
        self.update_status(f"Success! Saved to: {Path(output_file).name}")
        
        # Update last export display
        if self.current_file:
            last_export = self.converter.export_tracker.get_last_export_date(self.current_file)
            if last_export:
                self.last_export_label.config(
                    text=f"Last export: {last_export.strftime('%d.%m.%Y %H:%M')}"
                )
        
        messagebox.showinfo(
            "Success",
            f"Conversion complete!\n\nFile saved to:\n{output_file}"
        )
    
    def conversion_error(self, error_msg):
        """Handle conversion error"""
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        self.convert_btn.config(state="normal")
        
        self.update_status("Conversion failed")
        messagebox.showerror("Error", f"Conversion failed:\n{error_msg}")
    
    def update_status(self, message):
        """Update status bar"""
        self.status_label.config(text=message)


def main():
    """Main entry point"""
    
    # Check for tkinterdnd2
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except ImportError:
        root = tk.Tk()
    
    # Check platform
    system = platform.system()
    print(f"Running on: {system}")
    
    if system == "Linux":
        print("\nNote: If you see a tkinter error, install it with:")
        print("  sudo apt-get install python3-tk python3.12-tk")
    
    app = SimplifiedGlucoseGUI(root)
    
    # Center window
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()


if __name__ == '__main__':
    main()