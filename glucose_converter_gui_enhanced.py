#!/usr/bin/env python3
"""
Enhanced Glucose Converter GUI with Template and Date Filtering Support
Author: Your Assistant
Purpose: GUI with advanced features for glucose data conversion
"""

import sys
import os
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from tkinter import font as tkfont
import threading
from datetime import datetime, timedelta
from typing import Optional
import platform

# Import enhanced converter
from glucose_converter_enhanced import (
    EnhancedGlucoseConverter, 
    find_latest_csv,
    get_downloads_folder,
    TemplateManager
)


class EnhancedGlucoseGUI:
    """Enhanced GUI with template and date filtering support"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Glucose Data Converter - Enhanced")
        self.root.geometry("800x750")
        
        # Initialize converter and managers
        self.converter = EnhancedGlucoseConverter()
        self.template_manager = self.converter.template_manager
        self.current_file = None
        
        # Setup UI
        self.setup_ui()
        
        # Enable drag and drop (if tkinterdnd2 available)
        self.setup_drag_drop()
        
    def setup_ui(self):
        """Create the enhanced user interface"""
        
        # Create notebook for tabs
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Main conversion tab
        self.main_tab = ttk.Frame(notebook)
        notebook.add(self.main_tab, text="Convert")
        self.setup_main_tab()
        
        # Templates tab
        self.template_tab = ttk.Frame(notebook)
        notebook.add(self.template_tab, text="Templates")
        self.setup_template_tab()
        
        # Settings tab
        self.settings_tab = ttk.Frame(notebook)
        notebook.add(self.settings_tab, text="Settings")
        self.setup_settings_tab()
        
        # Export History tab
        self.history_tab = ttk.Frame(notebook)
        notebook.add(self.history_tab, text="Export History")
        self.setup_history_tab()
        
        # Status bar at bottom
        self.status_frame = ttk.Frame(self.root)
        self.status_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(0, 10))
        
        self.status_label = ttk.Label(self.status_frame, text="Ready", relief=tk.SUNKEN)
        self.status_label.pack(fill=tk.X)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(
            self.status_frame,
            mode='indeterminate',
            length=200
        )
    
    def setup_main_tab(self):
        """Setup the main conversion tab"""
        main_frame = ttk.Frame(self.main_tab, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Drop zone
        self.drop_frame = tk.Frame(
            main_frame,
            bg='#f0f0f0',
            relief=tk.SUNKEN,
            bd=2,
            height=120
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
        
        ttk.Label(info_frame, text="Selected File:").pack(anchor=tk.W)
        self.file_label = ttk.Label(info_frame, text="No file selected", foreground="gray")
        self.file_label.pack(anchor=tk.W, padx=20)
        
        # Date filtering
        date_frame = ttk.LabelFrame(main_frame, text="Date Filter", padding="10")
        date_frame.pack(fill=tk.X, pady=10)
        
        # Filter mode selection
        self.filter_mode = tk.StringVar(value="all")
        
        ttk.Radiobutton(
            date_frame, 
            text="All data", 
            variable=self.filter_mode, 
            value="all"
        ).grid(row=0, column=0, sticky=tk.W)
        
        ttk.Radiobutton(
            date_frame, 
            text="Incremental (since last export)", 
            variable=self.filter_mode, 
            value="incremental"
        ).grid(row=1, column=0, sticky=tk.W)
        
        ttk.Radiobutton(
            date_frame, 
            text="Last N days:", 
            variable=self.filter_mode, 
            value="days"
        ).grid(row=2, column=0, sticky=tk.W)
        
        self.days_var = tk.IntVar(value=30)
        days_spin = ttk.Spinbox(
            date_frame, 
            from_=1, 
            to=365, 
            width=10,
            textvariable=self.days_var
        )
        days_spin.grid(row=2, column=1, sticky=tk.W, padx=5)
        
        ttk.Radiobutton(
            date_frame, 
            text="Custom range", 
            variable=self.filter_mode, 
            value="custom"
        ).grid(row=3, column=0, sticky=tk.W)
        
        # Custom date range
        custom_frame = ttk.Frame(date_frame)
        custom_frame.grid(row=4, column=0, columnspan=2, sticky=tk.W, padx=20)
        
        ttk.Label(custom_frame, text="From:").grid(row=0, column=0)
        self.start_date_var = tk.StringVar(value=datetime.now().strftime('%d.%m.%Y'))
        ttk.Entry(custom_frame, textvariable=self.start_date_var, width=12).grid(row=0, column=1, padx=5)
        
        ttk.Label(custom_frame, text="To:").grid(row=0, column=2, padx=(10, 0))
        self.end_date_var = tk.StringVar(value=datetime.now().strftime('%d.%m.%Y'))
        ttk.Entry(custom_frame, textvariable=self.end_date_var, width=12).grid(row=0, column=3, padx=5)
        
        # Template selection
        template_frame = ttk.LabelFrame(main_frame, text="Template", padding="10")
        template_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(template_frame, text="Use template:").pack(anchor=tk.W)
        
        self.template_var = tk.StringVar(value="<default>")
        self.template_combo = ttk.Combobox(
            template_frame,
            textvariable=self.template_var,
            state="readonly"
        )
        self.template_combo.pack(fill=tk.X, pady=5)
        self.refresh_templates()
        
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
    
    def setup_template_tab(self):
        """Setup the templates management tab"""
        template_frame = ttk.Frame(self.template_tab, padding="20")
        template_frame.pack(fill=tk.BOTH, expand=True)
        
        # Template list
        ttk.Label(template_frame, text="Available Templates:", font=('Arial', 11, 'bold')).pack(anchor=tk.W)
        
        list_frame = ttk.Frame(template_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Listbox with scrollbar
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.template_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, height=10)
        self.template_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.template_listbox.yview)
        
        # Template actions
        action_frame = ttk.Frame(template_frame)
        action_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(
            action_frame,
            text="Upload Template",
            command=self.upload_template
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            action_frame,
            text="Delete Template",
            command=self.delete_template
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            action_frame,
            text="Set as Default",
            command=self.set_default_template
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            action_frame,
            text="Refresh List",
            command=self.refresh_template_list
        ).pack(side=tk.LEFT, padx=5)
        
        # Template info
        info_frame = ttk.LabelFrame(template_frame, text="Template Info", padding="10")
        info_frame.pack(fill=tk.X, pady=10)
        
        self.template_info_label = ttk.Label(info_frame, text="Select a template to view details")
        self.template_info_label.pack(anchor=tk.W)
        
        # Initial refresh
        self.refresh_template_list()
    
    def setup_settings_tab(self):
        """Setup the settings tab"""
        settings_frame = ttk.Frame(self.settings_tab, padding="20")
        settings_frame.pack(fill=tk.BOTH, expand=True)
        
        # Thresholds
        threshold_frame = ttk.LabelFrame(settings_frame, text="Glucose Thresholds (mmol/L)", padding="10")
        threshold_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(threshold_frame, text="Low (<):").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.low_threshold_var = tk.DoubleVar(value=self.converter.config['low_threshold'])
        ttk.Spinbox(
            threshold_frame, 
            from_=1.0, 
            to=10.0, 
            increment=0.1,
            textvariable=self.low_threshold_var,
            width=10
        ).grid(row=0, column=1, pady=5)
        
        ttk.Label(threshold_frame, text="High (>):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.high_threshold_var = tk.DoubleVar(value=self.converter.config['high_threshold'])
        ttk.Spinbox(
            threshold_frame, 
            from_=8.0, 
            to=20.0, 
            increment=0.1,
            textvariable=self.high_threshold_var,
            width=10
        ).grid(row=1, column=1, pady=5)
        
        ttk.Label(threshold_frame, text="Very High (>):").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.very_high_threshold_var = tk.DoubleVar(value=self.converter.config['very_high_threshold'])
        ttk.Spinbox(
            threshold_frame, 
            from_=15.0, 
            to=30.0, 
            increment=0.1,
            textvariable=self.very_high_threshold_var,
            width=10
        ).grid(row=2, column=1, pady=5)
        
        # Output settings
        output_frame = ttk.LabelFrame(settings_frame, text="Output Settings", padding="10")
        output_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(output_frame, text="Output Folder:").pack(anchor=tk.W)
        
        folder_frame = ttk.Frame(output_frame)
        folder_frame.pack(fill=tk.X, pady=5)
        
        self.output_var = tk.StringVar(value=self.converter.config.get('output_folder', 'Same as input'))
        self.output_entry = ttk.Entry(folder_frame, textvariable=self.output_var, state="readonly")
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(
            folder_frame,
            text="Browse",
            command=self.select_output_folder
        ).pack(side=tk.RIGHT, padx=(5, 0))
        
        self.auto_open_var = tk.BooleanVar(value=self.converter.config.get('auto_open', False))
        ttk.Checkbutton(
            output_frame,
            text="Open file after conversion",
            variable=self.auto_open_var
        ).pack(anchor=tk.W, pady=5)
        
        # Save settings button
        ttk.Button(
            settings_frame,
            text="Save Settings",
            command=self.save_settings
        ).pack(pady=20)
    
    def setup_history_tab(self):
        """Setup the export history tab"""
        history_frame = ttk.Frame(self.history_tab, padding="20")
        history_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(history_frame, text="Export History:", font=('Arial', 11, 'bold')).pack(anchor=tk.W)
        
        # History text widget
        self.history_text = scrolledtext.ScrolledText(
            history_frame,
            width=60,
            height=15,
            wrap=tk.WORD
        )
        self.history_text.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # History actions
        action_frame = ttk.Frame(history_frame)
        action_frame.pack(fill=tk.X)
        
        ttk.Button(
            action_frame,
            text="Refresh History",
            command=self.refresh_history
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            action_frame,
            text="Clear History",
            command=self.clear_history
        ).pack(side=tk.LEFT, padx=5)
        
        # Initial refresh
        self.refresh_history()
    
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
        self.file_label.config(text=Path(file_path).name, foreground="black")
        self.convert_btn.config(state="normal")
        self.update_status(f"Loaded: {Path(file_path).name}")
        
        self.drop_label.config(
            text=f"✅ {Path(file_path).name}\n\nDrop another file to replace",
            bg='#e6ffe6'
        )
        
        # Check if incremental export available
        last_export = self.converter.export_tracker.get_last_export_date(file_path)
        if last_export:
            self.update_status(f"Last export: {last_export.strftime('%d.%m.%Y %H:%M')}")
    
    def auto_detect_csv(self):
        """Auto-detect latest CSV"""
        self.update_status("Searching for latest CSV...")
        
        downloads = get_downloads_folder()
        latest_csv = find_latest_csv(str(downloads))
        
        if latest_csv:
            self.load_file(latest_csv)
            messagebox.showinfo("Success", f"Found: {Path(latest_csv).name}")
        else:
            messagebox.showwarning("Not Found", "No Contour CSV files found")
    
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
        
        # Get template
        template_name = None
        if self.template_var.get() != "<default>":
            template_name = self.template_var.get()
        
        # Update converter settings
        self.converter.config['low_threshold'] = self.low_threshold_var.get()
        self.converter.config['high_threshold'] = self.high_threshold_var.get()
        self.converter.config['very_high_threshold'] = self.very_high_threshold_var.get()
        self.converter.config['auto_open'] = self.auto_open_var.get()
        
        # Disable button and show progress
        self.convert_btn.config(state="disabled")
        self.progress_bar.pack(side=tk.LEFT, padx=10)
        self.progress_bar.start(10)
        self.update_status("Converting...")
        
        # Run in thread
        thread = threading.Thread(
            target=self.run_conversion,
            args=(start_date, end_date, incremental, template_name)
        )
        thread.start()
    
    def run_conversion(self, start_date, end_date, incremental, template_name):
        """Run the conversion process"""
        try:
            output_file = self.converter.convert(
                self.current_file,
                template_name=template_name,
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
        self.refresh_history()
        
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
    
    def refresh_templates(self):
        """Refresh template dropdown"""
        templates = ['<default>'] + self.template_manager.list_templates()
        self.template_combo['values'] = templates
        
        if self.converter.config.get('default_template') in templates:
            self.template_var.set(self.converter.config['default_template'])
    
    def refresh_template_list(self):
        """Refresh template listbox"""
        self.template_listbox.delete(0, tk.END)
        
        templates = self.template_manager.list_templates()
        for template in templates:
            display_name = template
            if template == self.converter.config.get('default_template'):
                display_name += " (default)"
            self.template_listbox.insert(tk.END, display_name)
        
        self.refresh_templates()
    
    def upload_template(self):
        """Upload a new template"""
        file_path = filedialog.askopenfilename(
            title="Select Template XLSX File",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        
        if file_path:
            name = tk.simpledialog.askstring("Template Name", "Enter template name:")
            if name:
                self.template_manager.save_template(file_path, name)
                self.refresh_template_list()
                messagebox.showinfo("Success", f"Template '{name}' uploaded successfully")
    
    def delete_template(self):
        """Delete selected template"""
        selection = self.template_listbox.curselection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a template to delete")
            return
        
        template_name = self.template_listbox.get(selection[0]).replace(" (default)", "")
        
        if messagebox.askyesno("Confirm", f"Delete template '{template_name}'?"):
            template_path = self.template_manager.get_template_path(template_name)
            if template_path:
                template_path.unlink()
                self.refresh_template_list()
                messagebox.showinfo("Success", f"Template '{template_name}' deleted")
    
    def set_default_template(self):
        """Set selected template as default"""
        selection = self.template_listbox.curselection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a template")
            return
        
        template_name = self.template_listbox.get(selection[0]).replace(" (default)", "")
        self.converter.config['default_template'] = template_name
        self.refresh_template_list()
        self.update_status(f"Default template: {template_name}")
    
    def select_output_folder(self):
        """Select output folder"""
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_var.set(folder)
            self.converter.config['output_folder'] = folder
    
    def save_settings(self):
        """Save current settings"""
        self.converter.config['low_threshold'] = self.low_threshold_var.get()
        self.converter.config['high_threshold'] = self.high_threshold_var.get()
        self.converter.config['very_high_threshold'] = self.very_high_threshold_var.get()
        self.converter.config['auto_open'] = self.auto_open_var.get()
        
        messagebox.showinfo("Success", "Settings saved")
        self.update_status("Settings saved")
    
    def refresh_history(self):
        """Refresh export history display"""
        self.history_text.delete(1.0, tk.END)
        
        history = self.converter.export_tracker.history
        if history:
            for file_path, info in history.items():
                file_name = Path(file_path).name
                last_export = info['last_export']
                updated = info['updated_at']
                
                self.history_text.insert(tk.END, f"File: {file_name}\n")
                self.history_text.insert(tk.END, f"  Last Export: {last_export}\n")
                self.history_text.insert(tk.END, f"  Updated: {updated}\n\n")
        else:
            self.history_text.insert(tk.END, "No export history available")
    
    def clear_history(self):
        """Clear export history"""
        if messagebox.askyesno("Confirm", "Clear all export history?"):
            self.converter.export_tracker.history = {}
            self.converter.export_tracker.save_history()
            self.refresh_history()
            self.update_status("History cleared")
    
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
        print("Note: Install tkinterdnd2 for drag-and-drop support")
    
    # Check platform
    system = platform.system()
    print(f"Running on: {system}")
    
    app = EnhancedGlucoseGUI(root)
    
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