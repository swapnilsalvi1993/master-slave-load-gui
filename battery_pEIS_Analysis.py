# -*- coding: utf-8 -*-
"""
Created on Sun Feb  1 17:30:31 2026

@author: ssalvi
"""

import os
from os import system
from IPython import get_ipython
import csv
import pandas as pd
import numpy as np
import scipy
import matplotlib.pyplot as plt
from matplotlib import rc
from matplotlib import cm
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import re
import glob
rc('mathtext', default='regular')

system('cls')
get_ipython().magic('reset -sf')


#%% GUI for File Selection and Parameter Input

class pEISConfigGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("pEIS Analysis Configuration - Pre/Post Test Comparison")
        self.root.geometry("650x750")
        self.root.resizable(False, False)
        
        # Variables to store user inputs
        self.csv_file_path_pre = None
        self.csv_file_path_post = None
        self.comparison_mode = tk.BooleanVar(value=False)
        self.cell_name = tk.StringVar(value="Unknown Cell")
        self.cell_id = tk.StringVar(value="Auto-detect")
        self.nominal_capacity = tk.StringVar(value="55")
        self.v_min = tk.StringVar(value="2.6")
        self.v_max = tk.StringVar(value="4.3")
        self.energy_capacity = tk.StringVar(value="200.75")
        
        # Step identification variables
        self.auto_detect_steps = tk.BooleanVar(value=True)
        self.char_step = tk.StringVar(value="7")
        self.disc_step = tk.StringVar(value="11")
        self.c_pulse_step = tk.StringVar(value="19")
        self.c_rest_step = tk.StringVar(value="20")
        self.d_pulse_step = tk.StringVar(value="22")
        
        # Plot options
        self.plot_full_test = tk.BooleanVar(value=True)
        self.plot_ocv = tk.BooleanVar(value=True)
        self.plot_impedance = tk.BooleanVar(value=True)
        self.generate_report = tk.BooleanVar(value=True)
        
        self.result = None
        
        self.create_widgets()
        
    def create_widgets(self):
        # Create main container frame
        container = ttk.Frame(self.root)
        container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create canvas
        canvas = tk.Canvas(container, highlightthickness=0)
        
        # Create scrollbar
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        
        # Create scrollable frame
        scrollable_frame = ttk.Frame(canvas)
        
        # Configure scrollable frame
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        # Create window in canvas
        canvas_frame = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        # Configure canvas scrolling
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Bind canvas width to frame width
        def on_canvas_configure(event):
            canvas.itemconfig(canvas_frame, width=event.width)
        canvas.bind('<Configure>', on_canvas_configure)
        
        # Pack scrollbar and canvas
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        # Enable mousewheel scrolling
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def bind_mousewheel(event):
            canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        def unbind_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
        
        canvas.bind('<Enter>', bind_mousewheel)
        canvas.bind('<Leave>', unbind_mousewheel)
        
        # Now use scrollable_frame for all widgets
        main_frame = scrollable_frame
        
        row = 0
        
        # Title
        title_label = ttk.Label(main_frame, text="pEIS Analysis Configuration", 
                                font=('Arial', 14, 'bold'))
        title_label.grid(row=row, column=0, columnspan=3, pady=10, sticky='w')
        row += 1
        
        # Comparison Mode Checkbox
        ttk.Checkbutton(main_frame, text="Enable Pre/Post Test Comparison Mode", 
                        variable=self.comparison_mode, 
                        command=self.toggle_comparison_mode).grid(row=row, column=0, columnspan=3, sticky='w', padx=5, pady=5)
        row += 1
        
        # File Selection Section
        ttk.Separator(main_frame, orient='horizontal').grid(row=row, column=0, columnspan=3, sticky='ew', pady=5)
        row += 1
        
        ttk.Label(main_frame, text="1. DATA FILE(S)", font=('Arial', 10, 'bold')).grid(row=row, column=0, columnspan=3, sticky='w', pady=5)
        row += 1
        
        # Pre-test file
        ttk.Label(main_frame, text="Pre-Test CSV:").grid(row=row, column=0, sticky='w', padx=5)
        self.file_label_pre = ttk.Label(main_frame, text="No file selected", foreground="gray", wraplength=250)
        self.file_label_pre.grid(row=row, column=1, sticky='w', padx=5)
        ttk.Button(main_frame, text="Browse...", command=lambda: self.browse_file('pre')).grid(row=row, column=2, padx=5, sticky='e')
        row += 1
        
        # Store the row where post-test widgets will be inserted
        self.post_test_row = row
        
        # Create post-test widgets but don't grid them yet
        self.post_label = ttk.Label(main_frame, text="Post-Test CSV:")
        self.file_label_post = ttk.Label(main_frame, text="No file selected", foreground="gray", wraplength=250)
        self.post_browse_btn = ttk.Button(main_frame, text="Browse...", command=lambda: self.browse_file('post'))
        self.auto_find_btn = ttk.Button(main_frame, text="Auto-Find Pair", command=self.auto_find_post)
        
        # Increment row for the remaining content
        # (These rows will be adjusted dynamically in toggle_comparison_mode)
        
        # Cell Information Section
        self.cell_info_separator = ttk.Separator(main_frame, orient='horizontal')
        self.cell_info_label = ttk.Label(main_frame, text="2. CELL INFORMATION", font=('Arial', 10, 'bold'))
        
        # Cell Name
        self.cell_name_label = ttk.Label(main_frame, text="Cell Name/Type:")
        self.cell_name_entry = ttk.Entry(main_frame, textvariable=self.cell_name, width=35)
        
        # Cell ID
        self.cell_id_label = ttk.Label(main_frame, text="Cell ID:")
        self.cell_id_entry = ttk.Entry(main_frame, textvariable=self.cell_id, width=35)
        self.cell_id_hint = ttk.Label(main_frame, text="(Leave as 'Auto-detect' to extract from filename)", 
                  font=('Arial', 8), foreground='gray')
        
        # Cell Specifications Section
        self.specs_separator = ttk.Separator(main_frame, orient='horizontal')
        self.specs_label = ttk.Label(main_frame, text="3. CELL SPECIFICATIONS", font=('Arial', 10, 'bold'))
        
        self.cap_label = ttk.Label(main_frame, text="Nominal Capacity (Ah):")
        self.cap_entry = ttk.Entry(main_frame, textvariable=self.nominal_capacity, width=15)
        
        self.vmin_label = ttk.Label(main_frame, text="Min Voltage (V):")
        self.vmin_entry = ttk.Entry(main_frame, textvariable=self.v_min, width=15)
        
        self.vmax_label = ttk.Label(main_frame, text="Max Voltage (V):")
        self.vmax_entry = ttk.Entry(main_frame, textvariable=self.v_max, width=15)
        
        self.energy_label = ttk.Label(main_frame, text="Energy Capacity (Wh):")
        self.energy_entry = ttk.Entry(main_frame, textvariable=self.energy_capacity, width=15)
        
        # Step Identification Section
        self.step_separator = ttk.Separator(main_frame, orient='horizontal')
        self.step_label = ttk.Label(main_frame, text="4. STEP IDENTIFICATION", font=('Arial', 10, 'bold'))
        self.step_auto_check = ttk.Checkbutton(main_frame, text="Auto-detect step numbers from data", 
                        variable=self.auto_detect_steps, 
                        command=self.toggle_step_entries)
        
        # Manual step entry frame
        self.step_frame = ttk.Frame(main_frame)
        
        step_row = 0
        ttk.Label(self.step_frame, text="Charge Step:").grid(row=step_row, column=0, sticky='w', padx=5, pady=2)
        self.char_step_entry = ttk.Entry(self.step_frame, textvariable=self.char_step, width=10)
        self.char_step_entry.grid(row=step_row, column=1, sticky='w', padx=5)
        ttk.Label(self.step_frame, text="(Baseline charge)", font=('Arial', 8), foreground='gray').grid(row=step_row, column=2, sticky='w', padx=5)
        step_row += 1
        
        ttk.Label(self.step_frame, text="Discharge Step:").grid(row=step_row, column=0, sticky='w', padx=5, pady=2)
        self.disc_step_entry = ttk.Entry(self.step_frame, textvariable=self.disc_step, width=10)
        self.disc_step_entry.grid(row=step_row, column=1, sticky='w', padx=5)
        ttk.Label(self.step_frame, text="(Baseline discharge)", font=('Arial', 8), foreground='gray').grid(row=step_row, column=2, sticky='w', padx=5)
        step_row += 1
        
        ttk.Label(self.step_frame, text="C_Pulse Step:").grid(row=step_row, column=0, sticky='w', padx=5, pady=2)
        self.c_pulse_entry = ttk.Entry(self.step_frame, textvariable=self.c_pulse_step, width=10)
        self.c_pulse_entry.grid(row=step_row, column=1, sticky='w', padx=5)
        ttk.Label(self.step_frame, text="(Charge pulse)", font=('Arial', 8), foreground='gray').grid(row=step_row, column=2, sticky='w', padx=5)
        step_row += 1
        
        ttk.Label(self.step_frame, text="C_Rest Step:").grid(row=step_row, column=0, sticky='w', padx=5, pady=2)
        self.c_rest_entry = ttk.Entry(self.step_frame, textvariable=self.c_rest_step, width=10)
        self.c_rest_entry.grid(row=step_row, column=1, sticky='w', padx=5)
        ttk.Label(self.step_frame, text="(Rest after charge)", font=('Arial', 8), foreground='gray').grid(row=step_row, column=2, sticky='w', padx=5)
        step_row += 1
        
        ttk.Label(self.step_frame, text="D_Pulse Step:").grid(row=step_row, column=0, sticky='w', padx=5, pady=2)
        self.d_pulse_entry = ttk.Entry(self.step_frame, textvariable=self.d_pulse_step, width=10)
        self.d_pulse_entry.grid(row=step_row, column=1, sticky='w', padx=5)
        ttk.Label(self.step_frame, text="(Discharge pulse)", font=('Arial', 8), foreground='gray').grid(row=step_row, column=2, sticky='w', padx=5)
        
        # Plot Options Section
        self.plot_separator = ttk.Separator(main_frame, orient='horizontal')
        self.plot_label = ttk.Label(main_frame, text="5. OUTPUT OPTIONS", font=('Arial', 10, 'bold'))
        
        self.plot_full_check = ttk.Checkbutton(main_frame, text="Generate Full Test Plot (Voltage & Current vs Time)", 
                        variable=self.plot_full_test)
        self.plot_ocv_check = ttk.Checkbutton(main_frame, text="Generate OCV vs SOC Plot", 
                        variable=self.plot_ocv)
        self.plot_imp_check = ttk.Checkbutton(main_frame, text="Generate Impedance vs SOC Plot", 
                        variable=self.plot_impedance)
        self.plot_report_check = ttk.Checkbutton(main_frame, text="Generate CSV Report", 
                        variable=self.generate_report)
        
        # Buttons
        self.button_separator = ttk.Separator(main_frame, orient='horizontal')
        self.button_frame = ttk.Frame(main_frame)
        
        ttk.Button(self.button_frame, text="Start Analysis", command=self.start_analysis, 
                   width=15).grid(row=0, column=0, padx=10)
        ttk.Button(self.button_frame, text="Cancel", command=self.cancel, 
                   width=15).grid(row=0, column=1, padx=10)
        
        # Initially layout all widgets
        self.layout_widgets()
        
        # Initially disable step entries if auto-detect is on
        self.toggle_step_entries()
    
    def layout_widgets(self):
        """Layout all widgets based on comparison mode"""
        # Start from after the pre-test file selection
        row = self.post_test_row
        
        # Show/hide post-test widgets based on comparison mode
        if self.comparison_mode.get():
            self.post_label.grid(row=row, column=0, sticky='w', padx=5, pady=3)
            self.file_label_post.grid(row=row, column=1, sticky='w', padx=5)
            self.post_browse_btn.grid(row=row, column=2, padx=5, sticky='e')
            row += 1
            
            self.auto_find_btn.grid(row=row, column=1, columnspan=2, pady=5, sticky='w', padx=5)
            row += 1
        else:
            self.post_label.grid_forget()
            self.file_label_post.grid_forget()
            self.post_browse_btn.grid_forget()
            self.auto_find_btn.grid_forget()
        
        # Cell Information Section
        self.cell_info_separator.grid(row=row, column=0, columnspan=3, sticky='ew', pady=5)
        row += 1
        
        self.cell_info_label.grid(row=row, column=0, columnspan=3, sticky='w', pady=5)
        row += 1
        
        self.cell_name_label.grid(row=row, column=0, sticky='w', padx=5, pady=3)
        self.cell_name_entry.grid(row=row, column=1, columnspan=2, sticky='w', padx=5)
        row += 1
        
        self.cell_id_label.grid(row=row, column=0, sticky='w', padx=5, pady=3)
        self.cell_id_entry.grid(row=row, column=1, columnspan=2, sticky='w', padx=5)
        row += 1
        
        self.cell_id_hint.grid(row=row, column=1, columnspan=2, sticky='w', padx=5)
        row += 1
        
        # Cell Specifications Section
        self.specs_separator.grid(row=row, column=0, columnspan=3, sticky='ew', pady=5)
        row += 1
        
        self.specs_label.grid(row=row, column=0, columnspan=3, sticky='w', pady=5)
        row += 1
        
        self.cap_label.grid(row=row, column=0, sticky='w', padx=5, pady=3)
        self.cap_entry.grid(row=row, column=1, sticky='w', padx=5)
        row += 1
        
        self.vmin_label.grid(row=row, column=0, sticky='w', padx=5, pady=3)
        self.vmin_entry.grid(row=row, column=1, sticky='w', padx=5)
        row += 1
        
        self.vmax_label.grid(row=row, column=0, sticky='w', padx=5, pady=3)
        self.vmax_entry.grid(row=row, column=1, sticky='w', padx=5)
        row += 1
        
        self.energy_label.grid(row=row, column=0, sticky='w', padx=5, pady=3)
        self.energy_entry.grid(row=row, column=1, sticky='w', padx=5)
        row += 1
        
        # Step Identification Section
        self.step_separator.grid(row=row, column=0, columnspan=3, sticky='ew', pady=5)
        row += 1
        
        self.step_label.grid(row=row, column=0, columnspan=3, sticky='w', pady=5)
        row += 1
        
        self.step_auto_check.grid(row=row, column=0, columnspan=3, sticky='w', padx=5, pady=3)
        row += 1
        
        self.step_frame.grid(row=row, column=0, columnspan=3, sticky='w', padx=20, pady=5)
        row += 1
        
        # Plot Options Section
        self.plot_separator.grid(row=row, column=0, columnspan=3, sticky='ew', pady=5)
        row += 1
        
        self.plot_label.grid(row=row, column=0, columnspan=3, sticky='w', pady=5)
        row += 1
        
        self.plot_full_check.grid(row=row, column=0, columnspan=3, sticky='w', padx=5, pady=2)
        row += 1
        
        self.plot_ocv_check.grid(row=row, column=0, columnspan=3, sticky='w', padx=5, pady=2)
        row += 1
        
        self.plot_imp_check.grid(row=row, column=0, columnspan=3, sticky='w', padx=5, pady=2)
        row += 1
        
        self.plot_report_check.grid(row=row, column=0, columnspan=3, sticky='w', padx=5, pady=2)
        row += 1
        
        # Add spacing
        ttk.Label(self.button_frame.master, text="").grid(row=row, column=0, pady=10)
        row += 1
        
        # Buttons
        self.button_separator.grid(row=row, column=0, columnspan=3, sticky='ew', pady=5)
        row += 1
        
        self.button_frame.grid(row=row, column=0, columnspan=3, pady=15)
        row += 1
        
        # Add padding at bottom
        ttk.Label(self.button_frame.master, text="").grid(row=row, column=0, pady=10)
    
    def toggle_comparison_mode(self):
        """Show/hide post-test file selection based on comparison mode"""
        self.layout_widgets()
            
        def toggle_comparison_mode(self):
            """Show/hide post-test file selection based on comparison mode"""
            if self.comparison_mode.get():
                # Show post-test widgets
                self.post_label.grid(row=self.post_row, column=0, sticky='w', padx=5)
                self.file_label_post.grid(row=self.post_row, column=1, sticky='w', padx=5)
                self.post_browse_btn.grid(row=self.post_row, column=2, padx=5, sticky='e')
                self.auto_find_btn.grid(row=self.post_row+1, column=1, columnspan=2, pady=5, sticky='w', padx=5)
            else:
                # Hide post-test widgets
                self.post_label.grid_forget()
                self.file_label_post.grid_forget()
                self.post_browse_btn.grid_forget()
                self.auto_find_btn.grid_forget()
    
    def browse_file(self, file_type):
        file_path = filedialog.askopenfilename(
            title=f"Select {'Pre-Test' if file_type == 'pre' else 'Post-Test'} pEIS CSV Data File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            if file_type == 'pre':
                self.csv_file_path_pre = file_path
                file_name = os.path.basename(file_path)
                self.file_label_pre.config(text=file_name, foreground="black")
                
                # Auto-populate cell info from pre-test file
                if self.cell_name.get() in ["Unknown Cell", ""]:
                    name_without_ext = os.path.splitext(file_name)[0]
                    clean_name = name_without_ext.replace('pEIS0', '').replace('pEIS_', '').replace('_', ' ').strip()
                    if clean_name:
                        self.cell_name.set(clean_name)
                
                if self.cell_id.get() == "Auto-detect":
                    extracted_id = self.extract_cell_id(file_name)
                    if extracted_id != "Unknown":
                        self.cell_id.set(extracted_id)
            else:
                self.csv_file_path_post = file_path
                file_name = os.path.basename(file_path)
                self.file_label_post.config(text=file_name, foreground="black")
    
    def auto_find_post(self):
        """Automatically find post-test file based on pre-test file"""
        if not self.csv_file_path_pre:
            messagebox.showwarning("Warning", "Please select a pre-test file first")
            return
        
        pre_dir = os.path.dirname(self.csv_file_path_pre)
        pre_file = os.path.basename(self.csv_file_path_pre)
        
        # Extract cell ID
        cell_id = self.extract_cell_id(pre_file)
        if cell_id == "Unknown":
            messagebox.showwarning("Warning", "Could not extract Cell ID from pre-test filename")
            return
        
        # Search for post-test files with same cell ID
        post_patterns = [
            f"*pEIS1*{cell_id}*.csv",
            f"*pEIS2*{cell_id}*.csv",
            f"*pEIS3*{cell_id}*.csv",
            f"*{cell_id}*pEIS1*.csv",
            f"*{cell_id}*pEIS2*.csv",
            f"*{cell_id}*pEIS3*.csv",
        ]
        
        found_files = []
        for pattern in post_patterns:
            found_files.extend(glob.glob(os.path.join(pre_dir, pattern)))
        
        # Remove duplicates and the pre-test file itself
        found_files = list(set(found_files))
        found_files = [f for f in found_files if f != self.csv_file_path_pre]
        
        if len(found_files) == 0:
            messagebox.showinfo("Not Found", f"No post-test file found for Cell ID: {cell_id}")
        elif len(found_files) == 1:
            self.csv_file_path_post = found_files[0]
            file_name = os.path.basename(found_files[0])
            self.file_label_post.config(text=file_name, foreground="black")
            messagebox.showinfo("Success", f"Found post-test file:\n{file_name}")
        else:
            # Multiple files found, let user choose
            choice_win = tk.Toplevel(self.root)
            choice_win.title("Select Post-Test File")
            choice_win.geometry("500x300")
            
            ttk.Label(choice_win, text=f"Multiple post-test files found for Cell ID {cell_id}:", 
                     font=('Arial', 10, 'bold')).pack(pady=10)
            
            listbox = tk.Listbox(choice_win, width=70, height=10)
            listbox.pack(padx=10, pady=10)
            
            for f in found_files:
                listbox.insert(tk.END, os.path.basename(f))
            
            def select_file():
                selection = listbox.curselection()
                if selection:
                    selected_file = found_files[selection[0]]
                    self.csv_file_path_post = selected_file
                    file_name = os.path.basename(selected_file)
                    self.file_label_post.config(text=file_name, foreground="black")
                    choice_win.destroy()
            
            ttk.Button(choice_win, text="Select", command=select_file).pack(pady=5)
    
    def extract_cell_id(self, filename):
        """Extract 10-digit Cell ID from filename"""
        name_without_ext = os.path.splitext(filename)[0]
        
        # Pattern 1: Look for 10-11 digit numbers
        pattern1 = re.search(r'\b0?(\d{10})\b', name_without_ext)
        if pattern1:
            return pattern1.group(1)
        
        # Pattern 2: Any sequence of 10+ digits
        pattern2 = re.search(r'0?(\d{10,})', name_without_ext)
        if pattern2:
            return pattern2.group(1)[:10]
        
        return "Unknown"
    
    def toggle_step_entries(self):
        """Enable/disable step entry fields based on auto-detect checkbox"""
        if self.auto_detect_steps.get():
            state = 'disabled'
        else:
            state = 'normal'
        
        self.char_step_entry.config(state=state)
        self.disc_step_entry.config(state=state)
        self.c_pulse_entry.config(state=state)
        self.c_rest_entry.config(state=state)
        self.d_pulse_entry.config(state=state)
    
    def validate_inputs(self):
        """Validate all user inputs"""
        if not self.csv_file_path_pre:
            messagebox.showerror("Error", "Please select a pre-test CSV file")
            return False
        
        if not os.path.exists(self.csv_file_path_pre):
            messagebox.showerror("Error", "Pre-test file does not exist")
            return False
        
        if self.comparison_mode.get():
            if not self.csv_file_path_post:
                messagebox.showerror("Error", "Please select a post-test CSV file for comparison mode")
                return False
            
            if not os.path.exists(self.csv_file_path_post):
                messagebox.showerror("Error", "Post-test file does not exist")
                return False
        
        # Validate numeric inputs
        try:
            float(self.nominal_capacity.get())
            float(self.v_min.get())
            float(self.v_max.get())
            float(self.energy_capacity.get())
        except ValueError:
            messagebox.showerror("Error", "Cell specifications must be valid numbers")
            return False
        
        # Validate step numbers if not auto-detecting
        if not self.auto_detect_steps.get():
            try:
                int(self.char_step.get())
                int(self.disc_step.get())
                int(self.c_pulse_step.get())
                int(self.c_rest_step.get())
                int(self.d_pulse_step.get())
            except ValueError:
                messagebox.showerror("Error", "Step numbers must be valid integers")
                return False
        
        return True
    
    def start_analysis(self):
        """Collect all inputs and close the GUI"""
        if not self.validate_inputs():
            return
        
        self.result = {
            'comparison_mode': self.comparison_mode.get(),
            'csv_file_path_pre': self.csv_file_path_pre,
            'csv_file_path_post': self.csv_file_path_post if self.comparison_mode.get() else None,
            'cell_name': self.cell_name.get(),
            'cell_id': self.cell_id.get(),
            'nominal_capacity': float(self.nominal_capacity.get()),
            'v_min': float(self.v_min.get()),
            'v_max': float(self.v_max.get()),
            'energy_capacity': float(self.energy_capacity.get()),
            'auto_detect_steps': self.auto_detect_steps.get(),
            'char_step': int(self.char_step.get()) if not self.auto_detect_steps.get() else None,
            'disc_step': int(self.disc_step.get()) if not self.auto_detect_steps.get() else None,
            'c_pulse_step': int(self.c_pulse_step.get()) if not self.auto_detect_steps.get() else None,
            'c_rest_step': int(self.c_rest_step.get()) if not self.auto_detect_steps.get() else None,
            'd_pulse_step': int(self.d_pulse_step.get()) if not self.auto_detect_steps.get() else None,
            'plot_full_test': self.plot_full_test.get(),
            'plot_ocv': self.plot_ocv.get(),
            'plot_impedance': self.plot_impedance.get(),
            'generate_report': self.generate_report.get()
        }
        
        self.root.quit()
        self.root.destroy()
    
    def cancel(self):
        """Cancel and exit"""
        self.result = None
        self.root.quit()
        self.root.destroy()


# Create and run the GUI
root = tk.Tk()
root.attributes('-topmost', True)
gui = pEISConfigGUI(root)
root.mainloop()

# Check if user cancelled
if gui.result is None:
    print("Analysis cancelled by user. Exiting...")
    exit()

# Extract configuration
config = gui.result
comparison_mode = config['comparison_mode']
csv_file_path_pre = config['csv_file_path_pre']
csv_file_path_post = config['csv_file_path_post']
Cell_ID = config['cell_id']
file_label = config['cell_name']
Cap0 = config['nominal_capacity']
Vmin = config['v_min']
Vmax = config['v_max']
E0 = config['energy_capacity']

print("\n" + "="*80)
print("pEIS ANALYSIS CONFIGURATION")
print("="*80)
print(f"Comparison Mode: {'ENABLED' if comparison_mode else 'DISABLED'}")
print(f"Pre-Test File: {csv_file_path_pre}")
if comparison_mode:
    print(f"Post-Test File: {csv_file_path_post}")
print(f"Cell Name: {file_label}")
print(f"Cell ID: {Cell_ID}")
print(f"Nominal Capacity: {Cap0} Ah")
print(f"Voltage Range: {Vmin} - {Vmax} V")
print(f"Energy Capacity: {E0} Wh")
print(f"Auto-detect Steps: {config['auto_detect_steps']}")
print("="*80 + "\n")


#%% Function to process a single pEIS file

def process_peis_file(csv_file_path, config, test_label=""):
    """Process a single pEIS CSV file and return results"""
    
    print(f"\n{'='*80}")
    print(f"Processing {test_label} File: {os.path.basename(csv_file_path)}")
    print(f"{'='*80}")
    
    # Load data
    Data0 = pd.read_csv(csv_file_path, header=None, nrows=28, names=range(3))
    print("Loading data header... Done")
    
    Data = pd.read_csv(csv_file_path, header=29-1)
    print("Loading main data... Done")
    
    Step = Data['Step'].values
    Cycle = Data['Cycle'].values
    TT = Data['Total Time (Seconds)'].values
    ST = Data['Step Time (Seconds)'].values
    Voltage = Data['Voltage (V)'].values
    Current = Data['Current (A)'].values
    CCap = Data['Charge Capacity (mAh)'].values
    DCap = Data['Discharge Capacity (mAh)'].values
    CEne = Data['Charge Energy (mWh)'].values
    DEne = Data['Discharge Energy (mWh)'].values
    TPos = Data['K1 (°C)'].values
    Res = Data['DC Internal Resistance (mOhm)'].values
    Temp_Time = TT[~np.isnan(Data['K1 (°C)'])][:]
    
    Power = Voltage*Current
    
    del Data
    
    print(f"Data points loaded: {len(TT)}")
    print(f"Test duration: {TT[-1]/3600:.2f} hours")
    
    # Initialize step variables
    char_step = None
    disc_step = None
    C_Pulse = None
    C_Rest_D = None
    D_Pulse = None
    
    # Auto-detect or use manual step numbers
    if config['auto_detect_steps']:
        print("\nIdentifying capacity calculation steps...")
        
        unique_steps = np.unique(Step[Step > 0])
        max_dcap_per_step = []
        max_ccap_per_step = []
        
        for step_num in unique_steps:
            step_mask = Step == step_num
            dcap_in_step = DCap[step_mask]
            ccap_in_step = CCap[step_mask]
            
            max_dcap = np.max(dcap_in_step[~np.isnan(dcap_in_step)]) if len(dcap_in_step[~np.isnan(dcap_in_step)]) > 0 else 0
            max_ccap = np.max(ccap_in_step[~np.isnan(ccap_in_step)]) if len(ccap_in_step[~np.isnan(ccap_in_step)]) > 0 else 0
            
            max_dcap_per_step.append((step_num, max_dcap))
            max_ccap_per_step.append((step_num, max_ccap))
        
        max_dcap_per_step.sort(key=lambda x: x[1], reverse=True)
        max_ccap_per_step.sort(key=lambda x: x[1], reverse=True)
        
        if len(max_dcap_per_step) > 0 and max_dcap_per_step[0][1] > 1000:
            disc_step = int(max_dcap_per_step[0][0])
            print(f"Found disc_step: Step {disc_step} ({max_dcap_per_step[0][1]/1000:.2f} Ah)")
        else:
            disc_step = 11
            print(f"Using default disc_step: Step {disc_step}")
        
        if len(max_ccap_per_step) > 0 and max_ccap_per_step[0][1] > 1000:
            char_step = int(max_ccap_per_step[0][0])
            print(f"Found char_step: Step {char_step} ({max_ccap_per_step[0][1]/1000:.2f} Ah)")
        else:
            char_step = 7
            print(f"Using default char_step: Step {char_step}")
        
        # Auto-detect pulse steps
        print("\nIdentifying pEIS step numbers...")
        
        Cyc_min = int(min(Cycle[Cycle>0]))
        Cyc_max = int(max(Cycle[Cycle>0]))
        
        pulse_start = Cyc_min + 1
        pulse_end = Cyc_max
        
        pulse_cycles = Cycle[(Cycle >= pulse_start) & (Cycle <= pulse_end)]
        pulse_steps = Step[(Cycle >= pulse_start) & (Cycle <= pulse_end)]
        pulse_current = Current[(Cycle >= pulse_start) & (Cycle <= pulse_end)]
        pulse_ST = ST[(Cycle >= pulse_start) & (Cycle <= pulse_end)]
        pulse_CCap = CCap[(Cycle >= pulse_start) & (Cycle <= pulse_end)]
        pulse_DCap = DCap[(Cycle >= pulse_start) & (Cycle <= pulse_end)]
        
        unique_pulse_steps = np.unique(pulse_steps)
        
        # Identify C_Pulse
        for step_num in unique_pulse_steps:
            step_mask = pulse_steps == step_num
            step_current = pulse_current[step_mask]
            step_ccap = pulse_CCap[step_mask]
            
            if len(step_current) > 0 and len(step_ccap) > 0:
                valid_current = step_current[~np.isnan(step_current)]
                valid_ccap = step_ccap[~np.isnan(step_ccap)]
                
                if len(valid_current) > 0 and len(valid_ccap) > 0:
                    avg_current = np.mean(valid_current)
                    ccap_range = np.max(valid_ccap) - np.min(valid_ccap)
                    
                    if avg_current > 5 and ccap_range > 50 and ccap_range < 2000:
                        C_Pulse = int(step_num)
                        print(f"Found C_Pulse: Step {C_Pulse}")
                        break
        
        if C_Pulse is None:
            charge_candidates = []
            for step_num in unique_pulse_steps:
                step_mask = pulse_steps == step_num
                step_current = pulse_current[step_mask]
                step_st = pulse_ST[step_mask]
                
                if len(step_current) > 0 and len(step_st) > 0:
                    valid_current = step_current[~np.isnan(step_current)]
                    valid_st = step_st[~np.isnan(step_st)]
                    
                    if len(valid_current) > 0 and len(valid_st) > 0:
                        avg_current = np.mean(valid_current)
                        max_step_time = np.max(valid_st)
                        
                        if avg_current > 5 and max_step_time > 10:
                            charge_candidates.append((step_num, avg_current))
            
            if charge_candidates:
                charge_candidates.sort(key=lambda x: x[1], reverse=True)
                C_Pulse = int(charge_candidates[0][0])
                print(f"Found C_Pulse: Step {C_Pulse}")
        
        if C_Pulse is None:
            C_Pulse = 19
            print(f"Using default C_Pulse: Step {C_Pulse}")
        
        # Identify C_Rest_D
        for step_num in unique_pulse_steps:
            if step_num <= C_Pulse:
                continue
            
            step_mask = pulse_steps == step_num
            step_current = pulse_current[step_mask]
            step_st = pulse_ST[step_mask]
            
            if len(step_current) > 0 and len(step_st) > 0:
                valid_current = step_current[~np.isnan(step_current)]
                valid_st = step_st[~np.isnan(step_st)]
                
                if len(valid_current) > 0 and len(valid_st) > 0:
                    avg_current = np.mean(np.abs(valid_current))
                    max_step_time = np.max(valid_st)
                    
                    if avg_current < 5 and max_step_time > 2 and max_step_time < 10:
                        C_Rest_D = int(step_num)
                        print(f"Found C_Rest_D: Step {C_Rest_D}")
                        break
        
        if C_Rest_D is None:
            C_Rest_D = 20
            print(f"Using default C_Rest_D: Step {C_Rest_D}")
        
        # Identify D_Pulse
        for step_num in unique_pulse_steps:
            if step_num <= C_Rest_D:
                continue
            
            step_mask = pulse_steps == step_num
            step_current = pulse_current[step_mask]
            step_dcap = pulse_DCap[step_mask]
            step_st = pulse_ST[step_mask]
            
            if len(step_current) > 0 and len(step_dcap) > 0 and len(step_st) > 0:
                valid_current = step_current[~np.isnan(step_current)]
                valid_dcap = step_dcap[~np.isnan(step_dcap)]
                valid_st = step_st[~np.isnan(step_st)]
                
                if len(valid_current) > 0 and len(valid_dcap) > 0 and len(valid_st) > 0:
                    avg_current = np.mean(valid_current)
                    dcap_range = np.max(valid_dcap) - np.min(valid_dcap)
                    max_step_time = np.max(valid_st)
                    
                    if avg_current < -10 and max_step_time > 2 and max_step_time < 10 and dcap_range > 10:
                        D_Pulse = int(step_num)
                        print(f"Found D_Pulse: Step {D_Pulse}")
                        break
        
        if D_Pulse is None:
            D_Pulse = 22
            print(f"Using default D_Pulse: Step {D_Pulse}")
        
    else:
        # Use manual step numbers
        char_step = config['char_step']
        disc_step = config['disc_step']
        C_Pulse = config['c_pulse_step']
        C_Rest_D = config['c_rest_step']
        D_Pulse = config['d_pulse_step']
        print(f"\nUsing manual step numbers: char={char_step}, disc={disc_step}, C_Pulse={C_Pulse}, C_Rest={C_Rest_D}, D_Pulse={D_Pulse}")
    
    # Print final step numbers
    print(f"\nFinal step numbers used:")
    print(f"  char_step = {char_step}")
    print(f"  disc_step = {disc_step}")
    print(f"  C_Pulse = {C_Pulse}")
    print(f"  C_Rest_D = {C_Rest_D}")
    print(f"  D_Pulse = {D_Pulse}")
    
    # Calculate Cell Capacity
    Cyc_min = int(min(Cycle[Cycle>0]))
    Cyc_max = int(max(Cycle[Cycle>0]))
    n_pulses = Cyc_max - Cyc_min + 1
    
    try:
        discharge_caps = DCap[Step == disc_step]
        if len(discharge_caps) > 0:
            valid_discharge_caps = discharge_caps[~np.isnan(discharge_caps)]
            if len(valid_discharge_caps) > 0:
                Cell_Cap = float(max(valid_discharge_caps))/1000
            else:
                raise ValueError("No valid discharge capacity data")
        else:
            raise ValueError("No data for discharge step")
    except:
        try:
            charge_caps = CCap[Step == char_step]
            if len(charge_caps) > 0:
                valid_charge_caps = charge_caps[~np.isnan(charge_caps)]
                if len(valid_charge_caps) > 0:
                    Cell_Cap = float(max(valid_charge_caps))/1000
                else:
                    raise ValueError("No valid charge capacity data")
            else:
                raise ValueError("No data for charge step")
        except:
            Cell_Cap = config['nominal_capacity']
    
    print(f'Cell Capacity = {Cell_Cap:.2f} Ah')
    print(f'Processing {n_pulses-2} pulses...')
    
    # Initialize arrays
    OCV = np.zeros(n_pulses)
    R_Discharge = np.zeros(n_pulses)
    R_Charge = np.zeros(n_pulses)
    tP_Discharge = np.zeros(n_pulses)
    tP_Charge = np.zeros(n_pulses)
    SOC_C = np.zeros(n_pulses)
    SOC_D = np.zeros(n_pulses)
    SOC_A = np.zeros(n_pulses)
    Cap_C = np.zeros(n_pulses)
    Cap_D = np.zeros(n_pulses)
    Cap_A = np.zeros(n_pulses)
    
    aux = 0
    
    for ip in range(1, n_pulses-2):
        Cyc_ind = Cyc_min + ip + 1
        
        try:
            ccap_data = CCap[(Cycle == Cyc_ind) & (Step == C_Pulse)]
            dcap_data = DCap[(Cycle == Cyc_ind) & (Step == D_Pulse)]
            
            if len(ccap_data) > 0:
                Cap_C[ip] = int(max(ccap_data[~np.isnan(ccap_data)]))/1000 if len(ccap_data[~np.isnan(ccap_data)]) > 0 else 0
            else:
                Cap_C[ip] = 0
            
            if len(dcap_data) > 0:
                Cap_D[ip] = int(max(dcap_data[~np.isnan(dcap_data)]))/1000 if len(dcap_data[~np.isnan(dcap_data)]) > 0 else 0
            else:
                Cap_D[ip] = 0
            
            aux = Cap_D[ip-1] + aux
            Cap_A[ip] = Cap_C[ip] - aux
            
            SOC_C[ip] = 100*Cap_C[ip]/Cell_Cap
            SOC_D[ip] = 100*Cap_D[ip]/Cell_Cap
            SOC_A[ip] = 100*Cap_A[ip]/Cell_Cap
            
            ocv_data = Voltage[(Step == C_Rest_D) & (Cycle == Cyc_ind)]
            if len(ocv_data) > 0:
                OCV[ip] = float(ocv_data[-1])
            
            V_0 = float(Voltage[(Step == C_Rest_D) & (Cycle == Cyc_ind)][-1]) if len(Voltage[(Step == C_Rest_D) & (Cycle == Cyc_ind)]) > 0 else 0
            I_0 = float(Current[(Step == C_Rest_D) & (Cycle == Cyc_ind)][-1]) if len(Current[(Step == C_Rest_D) & (Cycle == Cyc_ind)]) > 0 else 0
            t_0 = float(TT[(Step == C_Rest_D) & (Cycle == Cyc_ind)][-1]) if len(TT[(Step == C_Rest_D) & (Cycle == Cyc_ind)]) > 0 else 0
            
            V_1 = float(Voltage[(Step == D_Pulse) & (Cycle == Cyc_ind)][-1]) if len(Voltage[(Step == D_Pulse) & (Cycle == Cyc_ind)]) > 0 else 0
            I_1 = float(Current[(Step == D_Pulse) & (Cycle == Cyc_ind)][-1]) if len(Current[(Step == D_Pulse) & (Cycle == Cyc_ind)]) > 0 else 0
            t_1 = float(TT[(Step == D_Pulse) & (Cycle == Cyc_ind)][-1]) if len(TT[(Step == D_Pulse) & (Cycle == Cyc_ind)]) > 0 else 0
            
            R_Discharge[ip] = 1000*np.abs((V_1-V_0)/(np.abs(I_1)-np.abs(I_0))) if (np.abs(I_1)-np.abs(I_0)) != 0 else 0
            tP_Discharge[ip] = np.abs(t_1 - t_0)
            
            V_2 = float(Voltage[(Step == C_Rest_D) & (Cycle == Cyc_ind)][-1]) if len(Voltage[(Step == C_Rest_D) & (Cycle == Cyc_ind)]) > 0 else 0
            I_2 = float(Current[(Step == C_Rest_D) & (Cycle == Cyc_ind)][-1]) if len(Current[(Step == C_Rest_D) & (Cycle == Cyc_ind)]) > 0 else 0
            t_2 = float(TT[(Step == C_Rest_D) & (Cycle == Cyc_ind)][-1]) if len(TT[(Step == C_Rest_D) & (Cycle == Cyc_ind)]) > 0 else 0
            
            V_3 = float(Voltage[(Step == C_Pulse) & (Cycle == Cyc_ind)][-1]) if len(Voltage[(Step == C_Pulse) & (Cycle == Cyc_ind)]) > 0 else 0
            I_3 = float(Current[(Step == C_Pulse) & (Cycle == Cyc_ind)][-1]) if len(Current[(Step == C_Pulse) & (Cycle == Cyc_ind)]) > 0 else 0
            t_3 = float(TT[(Step == C_Pulse) & (Cycle == Cyc_ind)][-1]) if len(TT[(Step == C_Pulse) & (Cycle == Cyc_ind)]) > 0 else 0
            
            R_Charge[ip] = 1000*np.abs((V_3-V_2)/(np.abs(I_3)-np.abs(I_2))) if (np.abs(I_3)-np.abs(I_2)) != 0 else 0
            tP_Charge[ip] = np.abs(t_3 - t_2)
            
        except Exception as e:
            continue
    
    print(f"Processing complete!")
    
    return {
        'TT': TT,
        'Voltage': Voltage,
        'Current': Current,
        'Cell_Cap': Cell_Cap,
        'SOC_A': SOC_A,
        'SOC_C': SOC_C,
        'SOC_D': SOC_D,
        'OCV': OCV,
        'R_Charge': R_Charge,
        'R_Discharge': R_Discharge,
        'tP_Charge': tP_Charge,
        'tP_Discharge': tP_Discharge,
        'Data0': Data0,
        # Add detected step numbers - CRITICAL: These must be set
        'char_step': char_step,
        'disc_step': disc_step,
        'C_Pulse': C_Pulse,
        'C_Rest_D': C_Rest_D,
        'D_Pulse': D_Pulse
    }

#%% Process files

results_pre = process_peis_file(csv_file_path_pre, config, "PRE-TEST")

if comparison_mode:
    results_post = process_peis_file(csv_file_path_post, config, "POST-TEST")

# Setup output directory
f_dir = os.path.dirname(csv_file_path_pre) + '\\'
file_name_pre = os.path.splitext(os.path.basename(csv_file_path_pre))[0]

if comparison_mode:
    file_name_post = os.path.splitext(os.path.basename(csv_file_path_post))[0]
    output_folder_name = f"{Cell_ID}_PrePost_Comparison"
else:
    output_folder_name = file_name_pre + '_plots'

f_plots = f_dir + output_folder_name + '\\'

if not os.path.exists(f_plots):
    os.makedirs(f_plots)
    print(f"\nCreated output directory: {f_plots}")


#%% Generate Plots

print("\n" + "="*80)
print("GENERATING PLOTS")
print("="*80)

# Plot 1: Full Test Voltage/Current (Pre-test only or separate for each)
if config['plot_full_test']:
    print("\nGenerating Full Test Plots...")
    
    # Pre-test plot
    f1 = plt.figure(figsize=(13,9))
    title_font = 20
    ticks_font = 16
    
    ax1 = f1.add_subplot(1,1,1)
    lns1 = ax1.plot(results_pre['TT']/60, results_pre['Voltage'], color='blue', linewidth=2, label='Voltage', zorder=2)
    
    ax1.set_ylabel('Voltage [V]', fontsize=title_font, fontweight='bold', labelpad=15)
    ax1.grid(color='gray', linestyle='--', linewidth=0.5)
    ax1.yaxis.set_tick_params(labelsize=ticks_font)
    ax1.set_ylim([Vmin-0.5, Vmax+0.6])
    
    ax2 = ax1.twinx()
    lns2 = ax2.plot(results_pre['TT']/60, results_pre['Current'], color='red', linewidth=2, label='Current', zorder=1)
    ax2.set_ylabel('Current [A]', fontsize=title_font, fontweight='bold', labelpad=15)
    ax2.yaxis.set_tick_params(labelsize=ticks_font)
    
    current_max = np.ceil(np.max(results_pre['Current'])/10)*10
    current_min = np.floor(np.min(results_pre['Current'])/10)*10
    ax2.set_ylim([current_min, current_max])
    
    ax1.set_xlabel('Time [min]', fontsize=title_font, fontweight='bold', labelpad=15)
    ax1.xaxis.set_tick_params(labelsize=ticks_font)
    
    time_max = np.ceil(results_pre['TT'][-1]/60/100)*100
    ax1.set_xlim([0, time_max])
    
    lns = lns1 + lns2
    labs = [l.get_label() for l in lns]
    ax1.legend(lns, labs, fontsize=title_font, title=f"Cell ID: {Cell_ID} - Pre-Test", title_fontsize=title_font, loc='best')
    
    plt.tight_layout()
    f1.savefig(f_plots + f'{Cell_ID}_PreTest_Voltage_Current_vs_Time.png', bbox_inches='tight', dpi=200)
    plt.close()
    print(f"Saved: {Cell_ID}_PreTest_Voltage_Current_vs_Time.png")
    
    # Post-test plot if in comparison mode
    if comparison_mode:
        f1 = plt.figure(figsize=(13,9))
        
        ax1 = f1.add_subplot(1,1,1)
        lns1 = ax1.plot(results_post['TT']/60, results_post['Voltage'], color='blue', linewidth=2, label='Voltage', zorder=2)
        
        ax1.set_ylabel('Voltage [V]', fontsize=title_font, fontweight='bold', labelpad=15)
        ax1.grid(color='gray', linestyle='--', linewidth=0.5)
        ax1.yaxis.set_tick_params(labelsize=ticks_font)
        ax1.set_ylim([Vmin-0.5, Vmax+0.6])
        
        ax2 = ax1.twinx()
        lns2 = ax2.plot(results_post['TT']/60, results_post['Current'], color='red', linewidth=2, label='Current', zorder=1)
        ax2.set_ylabel('Current [A]', fontsize=title_font, fontweight='bold', labelpad=15)
        ax2.yaxis.set_tick_params(labelsize=ticks_font)
        
        current_max = np.ceil(np.max(results_post['Current'])/10)*10
        current_min = np.floor(np.min(results_post['Current'])/10)*10
        ax2.set_ylim([current_min, current_max])
        
        ax1.set_xlabel('Time [min]', fontsize=title_font, fontweight='bold', labelpad=15)
        ax1.xaxis.set_tick_params(labelsize=ticks_font)
        
        time_max = np.ceil(results_post['TT'][-1]/60/100)*100
        ax1.set_xlim([0, time_max])
        
        lns = lns1 + lns2
        labs = [l.get_label() for l in lns]
        ax1.legend(lns, labs, fontsize=title_font, title=f"Cell ID: {Cell_ID} - Post-Test", title_fontsize=title_font, loc='best')
        
        plt.tight_layout()
        f1.savefig(f_plots + f'{Cell_ID}_PostTest_Voltage_Current_vs_Time.png', bbox_inches='tight', dpi=200)
        plt.close()
        print(f"Saved: {Cell_ID}_PostTest_Voltage_Current_vs_Time.png")


# Plot 2: OCV vs SOC Comparison
if config['plot_ocv']:
    print("\nGenerating OCV vs SOC Plot...")
    
    f1 = plt.figure(figsize=(12,8))
    title_font = 16
    ticks_font = 14
    legend_font = 14
    thickness = 1.5
    
    ax1 = f1.add_subplot(1,1,1)
    
    # Pre-test data
    valid_mask_pre = results_pre['OCV'] > 0
    lns1 = ax1.plot(results_pre['SOC_A'][valid_mask_pre], results_pre['OCV'][valid_mask_pre], 
                    color='blue', marker='o', markersize=6, linewidth=thickness, label='Pre-Test', zorder=2)
    
    # Post-test data if in comparison mode
    if comparison_mode:
        valid_mask_post = results_post['OCV'] > 0
        lns2 = ax1.plot(results_post['SOC_A'][valid_mask_post], results_post['OCV'][valid_mask_post], 
                        color='red', marker='s', markersize=6, linewidth=thickness, label='Post-Test', zorder=1)
    
    ax1.set_ylabel('Voltage [V]', fontsize=title_font, fontweight='bold', labelpad=15)
    ax1.grid(color='gray', linestyle='--', linewidth=0.5)
    ax1.yaxis.set_tick_params(labelsize=ticks_font)
    ax1.set_ylim([Vmin-0.5, Vmax+0.6])
    
    ax1.set_xlabel('SOC [%]', fontsize=title_font, fontweight='bold', labelpad=15)
    ax1.xaxis.set_tick_params(labelsize=ticks_font)
    ax1.set_xlim([-1, 120])
    
    if comparison_mode:
        ax1.legend(fontsize=legend_font, title=f"Cell ID: {Cell_ID}\nOCV Comparison", 
                  title_fontsize=title_font, loc='best')
    else:
        ax1.legend(fontsize=legend_font, title=f"Cell ID: {Cell_ID}", 
                  title_fontsize=title_font, loc='best')
    
    plt.tight_layout()
    f1.savefig(f_plots + f'{Cell_ID}_OCV_Comparison.png', bbox_inches='tight', dpi=200)
    plt.close()
    print(f"Saved: {Cell_ID}_OCV_Comparison.png")


# Plot 3: Impedance vs SOC Comparison
if config['plot_impedance']:
    print("\nGenerating Impedance vs SOC Plot...")
    
    f1 = plt.figure(figsize=(12,8))
    title_font = 16
    ticks_font = 14
    legend_font = 14
    thickness = 1.5
    
    ax1 = f1.add_subplot(1,1,1)
    
    # Pre-test data (remove last point)
    valid_mask_pre = results_pre['R_Charge'] > 0
    SOC_pre = results_pre['SOC_A'][valid_mask_pre]
    R_pre = results_pre['R_Charge'][valid_mask_pre]
    
    if len(SOC_pre) > 1:
        SOC_pre = SOC_pre[:-1]
        R_pre = R_pre[:-1]
    
    lns1 = ax1.plot(SOC_pre, R_pre, color='blue', marker='^', markersize=8, 
                    linewidth=thickness, label='Pre-Test', zorder=2)
    
    # Post-test data if in comparison mode
    if comparison_mode:
        valid_mask_post = results_post['R_Charge'] > 0
        SOC_post = results_post['SOC_A'][valid_mask_post]
        R_post = results_post['R_Charge'][valid_mask_post]
        
        if len(SOC_post) > 1:
            SOC_post = SOC_post[:-1]
            R_post = R_post[:-1]
        
        lns2 = ax1.plot(SOC_post, R_post, color='red', marker='v', markersize=8, 
                        linewidth=thickness, label='Post-Test', zorder=1)
    
    ax1.set_ylabel('Transition Impedance (ZTR) [m\u03a9]', fontsize=title_font, fontweight='bold', labelpad=15)
    ax1.grid(color='gray', linestyle='--', linewidth=0.5)
    ax1.yaxis.set_tick_params(labelsize=ticks_font)
    
    # Auto-scale y-axis
    all_R = np.concatenate([R_pre, R_post]) if comparison_mode else R_pre
    if len(all_R) > 0:
        r_max = np.ceil(np.max(all_R)*1.1/0.4)*0.4
        ax1.set_ylim([0, r_max])
    
    ax1.set_xlabel('SOC [%]', fontsize=title_font, fontweight='bold', labelpad=15)
    ax1.xaxis.set_tick_params(labelsize=ticks_font)
    ax1.set_xlim([-1, 110])
    
    if comparison_mode:
        ax1.legend(fontsize=legend_font, title=f"Cell ID: {Cell_ID}\nImpedance Comparison", 
                  title_fontsize=title_font, loc='best')
    else:
        ax1.legend(fontsize=legend_font, title=f"Cell ID: {Cell_ID}", 
                  title_fontsize=title_font, loc='best')
    
    plt.tight_layout()
    f1.savefig(f_plots + f'{Cell_ID}_Impedance_Comparison.png', bbox_inches='tight', dpi=200)
    plt.close()
    print(f"Saved: {Cell_ID}_Impedance_Comparison.png")


#%% Generate Report

if config['generate_report']:
    print("\nGenerating CSV Report...")
    
    if comparison_mode:
        report_file = f_plots + f'{Cell_ID}_PrePost_Comparison_Report.csv'
        
        with open(report_file, 'w', newline='') as file:
            writer = csv.writer(file)
            
            # Header
            writer.writerow(['pEIS Pre/Post Test Comparison Report'])
            writer.writerow(['Cell ID', Cell_ID])
            writer.writerow(['Cell Name', file_label])
            writer.writerow(['Analysis Date', pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')])
            writer.writerow([])
            
            # File information
            writer.writerow(['FILE INFORMATION'])
            writer.writerow(['Pre-Test File', os.path.basename(csv_file_path_pre)])
            writer.writerow(['Post-Test File', os.path.basename(csv_file_path_post)])
            writer.writerow([])
            
            # Step detection information - PRE-TEST
            writer.writerow(['PRE-TEST STEP DETECTION'])
            if config['auto_detect_steps']:
                writer.writerow(['Step Detection Method', 'Auto-Detected'])
                writer.writerow(['char_step (Baseline Charge)', results_pre.get('char_step', 'N/A')])
                writer.writerow(['disc_step (Baseline Discharge)', results_pre.get('disc_step', 'N/A')])
                writer.writerow(['C_Pulse (Charge Pulse)', results_pre.get('C_Pulse', 'N/A')])
                writer.writerow(['C_Rest_D (Rest After Charge)', results_pre.get('C_Rest_D', 'N/A')])
                writer.writerow(['D_Pulse (Discharge Pulse)', results_pre.get('D_Pulse', 'N/A')])
            else:
                writer.writerow(['Step Detection Method', 'Manual Entry'])
                writer.writerow(['char_step (Baseline Charge)', config['char_step']])
                writer.writerow(['disc_step (Baseline Discharge)', config['disc_step']])
                writer.writerow(['C_Pulse (Charge Pulse)', config['c_pulse_step']])
                writer.writerow(['C_Rest_D (Rest After Charge)', config['c_rest_step']])
                writer.writerow(['D_Pulse (Discharge Pulse)', config['d_pulse_step']])
            writer.writerow([])
            
            # Step detection information - POST-TEST
            writer.writerow(['POST-TEST STEP DETECTION'])
            if config['auto_detect_steps']:
                writer.writerow(['Step Detection Method', 'Auto-Detected'])
                writer.writerow(['char_step (Baseline Charge)', results_post.get('char_step', 'N/A')])
                writer.writerow(['disc_step (Baseline Discharge)', results_post.get('disc_step', 'N/A')])
                writer.writerow(['C_Pulse (Charge Pulse)', results_post.get('C_Pulse', 'N/A')])
                writer.writerow(['C_Rest_D (Rest After Charge)', results_post.get('C_Rest_D', 'N/A')])
                writer.writerow(['D_Pulse (Discharge Pulse)', results_post.get('D_Pulse', 'N/A')])
            else:
                writer.writerow(['Step Detection Method', 'Manual Entry'])
                writer.writerow(['char_step (Baseline Charge)', config['char_step']])
                writer.writerow(['disc_step (Baseline Discharge)', config['disc_step']])
                writer.writerow(['C_Pulse (Charge Pulse)', config['c_pulse_step']])
                writer.writerow(['C_Rest_D (Rest After Charge)', config['c_rest_step']])
                writer.writerow(['D_Pulse (Discharge Pulse)', config['d_pulse_step']])
            writer.writerow([])
            
            # Capacity comparison
            writer.writerow(['CAPACITY COMPARISON'])
            writer.writerow(['Pre-Test Capacity (Ah)', f"{results_pre['Cell_Cap']:.2f}"])
            writer.writerow(['Post-Test Capacity (Ah)', f"{results_post['Cell_Cap']:.2f}"])
            writer.writerow(['Capacity Fade (Ah)', f"{results_pre['Cell_Cap'] - results_post['Cell_Cap']:.2f}"])
            writer.writerow(['Capacity Fade (%)', f"{100*(results_pre['Cell_Cap'] - results_post['Cell_Cap'])/results_pre['Cell_Cap']:.2f}"])
            writer.writerow([])
            
            # Impedance comparison
            pre_imp_valid = results_pre['R_Charge'][results_pre['R_Charge'] > 0]
            post_imp_valid = results_post['R_Charge'][results_post['R_Charge'] > 0]
            
            if len(pre_imp_valid) > 0 and len(post_imp_valid) > 0:
                pre_imp_avg = np.mean(pre_imp_valid)
                post_imp_avg = np.mean(post_imp_valid)
                writer.writerow(['IMPEDANCE COMPARISON'])
                writer.writerow(['Pre-Test Average Impedance (mOhms)', f"{pre_imp_avg:.3f}"])
                writer.writerow(['Post-Test Average Impedance (mOhms)', f"{post_imp_avg:.3f}"])
                writer.writerow(['Impedance Increase (mOhms)', f"{post_imp_avg - pre_imp_avg:.3f}"])
                writer.writerow(['Impedance Increase (%)', f"{100*(post_imp_avg - pre_imp_avg)/pre_imp_avg:.2f}"])
                writer.writerow([])
            
            # Pre-test data
            writer.writerow(['PRE-TEST DATA'])
            writer.writerow(['SOC [%]', 'OCV [V]', 'Charge Impedance [mOhms]', 'Discharge Resistance [mOhms]'])
            
            for i in range(len(results_pre['SOC_A'])-1):
                if results_pre['SOC_A'][i+1] != 0 or results_pre['OCV'][i+1] != 0:
                    writer.writerow([
                        f"{results_pre['SOC_A'][i+1]:.2f}",
                        f"{results_pre['OCV'][i+1]:.4f}",
                        f"{results_pre['R_Charge'][i]:.4f}",
                        f"{results_pre['R_Discharge'][i]:.4f}"
                    ])
            
            writer.writerow([])
            
            # Post-test data
            writer.writerow(['POST-TEST DATA'])
            writer.writerow(['SOC [%]', 'OCV [V]', 'Charge Impedance [mOhms]', 'Discharge Resistance [mOhms]'])
            
            for i in range(len(results_post['SOC_A'])-1):
                if results_post['SOC_A'][i+1] != 0 or results_post['OCV'][i+1] != 0:
                    writer.writerow([
                        f"{results_post['SOC_A'][i+1]:.2f}",
                        f"{results_post['OCV'][i+1]:.4f}",
                        f"{results_post['R_Charge'][i]:.4f}",
                        f"{results_post['R_Discharge'][i]:.4f}"
                    ])
        
        print(f"Saved: {report_file}")
    
    else:
        report_file = f_plots + f'{Cell_ID}_Report.csv'
        
        with open(report_file, 'w', newline='') as file:
            writer = csv.writer(file)
            
            # Header
            writer.writerow(['pEIS Analysis Report'])
            writer.writerow(['Cell ID', Cell_ID])
            writer.writerow(['Cell Name', file_label])
            writer.writerow(['Data File', os.path.basename(csv_file_path_pre)])
            writer.writerow(['Analysis Date', pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')])
            writer.writerow([])
            
            # Step detection information
            writer.writerow(['STEP DETECTION INFORMATION'])
            if config['auto_detect_steps']:
                writer.writerow(['Step Detection Method', 'Auto-Detected'])
                writer.writerow(['char_step (Baseline Charge)', results_pre.get('char_step', 'N/A')])
                writer.writerow(['disc_step (Baseline Discharge)', results_pre.get('disc_step', 'N/A')])
                writer.writerow(['C_Pulse (Charge Pulse)', results_pre.get('C_Pulse', 'N/A')])
                writer.writerow(['C_Rest_D (Rest After Charge)', results_pre.get('C_Rest_D', 'N/A')])
                writer.writerow(['D_Pulse (Discharge Pulse)', results_pre.get('D_Pulse', 'N/A')])
            else:
                writer.writerow(['Step Detection Method', 'Manual Entry'])
                writer.writerow(['char_step (Baseline Charge)', config['char_step']])
                writer.writerow(['disc_step (Baseline Discharge)', config['disc_step']])
                writer.writerow(['C_Pulse (Charge Pulse)', config['c_pulse_step']])
                writer.writerow(['C_Rest_D (Rest After Charge)', config['c_rest_step']])
                writer.writerow(['D_Pulse (Discharge Pulse)', config['d_pulse_step']])
            writer.writerow([])
            
            # Cell capacity
            writer.writerow(['CELL CAPACITY'])
            writer.writerow(['Cell Capacity (Ah)', f"{results_pre['Cell_Cap']:.2f}"])
            writer.writerow([])
            
            # Average impedance
            imp_valid = results_pre['R_Charge'][results_pre['R_Charge'] > 0]
            if len(imp_valid) > 0:
                writer.writerow(['IMPEDANCE STATISTICS'])
                writer.writerow(['Average Impedance (mOhms)', f"{np.mean(imp_valid):.3f}"])
                writer.writerow(['Min Impedance (mOhms)', f"{np.min(imp_valid):.3f}"])
                writer.writerow(['Max Impedance (mOhms)', f"{np.max(imp_valid):.3f}"])
                writer.writerow([])
            
            # Data table
            writer.writerow(['MEASUREMENT DATA'])
            writer.writerow(['SOC [%]', 'OCV [V]', 'Charge Impedance [mOhms]', 'Discharge Resistance [mOhms]'])
            
            for i in range(len(results_pre['SOC_A'])-1):
                if results_pre['SOC_A'][i+1] != 0 or results_pre['OCV'][i+1] != 0:
                    writer.writerow([
                        f"{results_pre['SOC_A'][i+1]:.2f}",
                        f"{results_pre['OCV'][i+1]:.4f}",
                        f"{results_pre['R_Charge'][i]:.4f}",
                        f"{results_pre['R_Discharge'][i]:.4f}"
                    ])
        
        print(f"Saved: {report_file}")


#%% Summary

print("\n" + "="*80)
print("ANALYSIS COMPLETE!")
print("="*80)
print(f"Cell ID: {Cell_ID}")
print(f"Cell Name: {file_label}")

if comparison_mode:
    print(f"\nPre-Test Capacity: {results_pre['Cell_Cap']:.2f} Ah")
    print(f"Post-Test Capacity: {results_post['Cell_Cap']:.2f} Ah")
    print(f"Capacity Fade: {results_pre['Cell_Cap'] - results_post['Cell_Cap']:.2f} Ah ({100*(results_pre['Cell_Cap'] - results_post['Cell_Cap'])/results_pre['Cell_Cap']:.2f}%)")
    
    # Impedance comparison
    pre_imp_avg = np.mean(results_pre['R_Charge'][results_pre['R_Charge'] > 0])
    post_imp_avg = np.mean(results_post['R_Charge'][results_post['R_Charge'] > 0])
    print(f"\nAverage Impedance (Pre-Test): {pre_imp_avg:.3f} mΩ")
    print(f"Average Impedance (Post-Test): {post_imp_avg:.3f} mΩ")
    print(f"Impedance Increase: {post_imp_avg - pre_imp_avg:.3f} mΩ ({100*(post_imp_avg - pre_imp_avg)/pre_imp_avg:.2f}%)")
else:
    print(f"\nCell Capacity: {results_pre['Cell_Cap']:.2f} Ah")
    print(f"Average Impedance: {np.mean(results_pre['R_Charge'][results_pre['R_Charge'] > 0]):.3f} mΩ")

print(f"\nAll files saved to: {f_plots}")
print("="*80)
