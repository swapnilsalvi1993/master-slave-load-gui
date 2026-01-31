import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
import sys
import tkinter as tk
from tkinter import Tk, filedialog, ttk, messagebox
from scipy import stats

print("Starting script...")

# Try to import pyxlsb, provide helpful error if not available
try:
    import pyxlsb
    has_pyxlsb = True
    print("✓ pyxlsb module loaded")
except ImportError:
    has_pyxlsb = False
    print("⚠ Warning: pyxlsb not installed.")
    print("To install, run: pip install pyxlsb")

# All available plot parameters grouped by category
# Format: (y_param_name, y_unit, x_param_name, x_unit, short_name_for_file, category_for_ylim, display_name)
ALL_PLOT_PARAMS = {
    'Temperature Plots': [
        ('Maximum Cell Avg Temperature', '°C', 'Usable Capacity', '%', 'Max_Cell_Temperature_vs_UsableCapacity', 'temperature', 'Maximum Cell Avg Temperature vs Usable Capacity'),
        ('Cell Avg Temperature at t=60 sec', '°C', 'Usable Capacity', '%', 'Cell_Temperature_60sec_vs_UsableCapacity', 'temperature', 'Cell Avg Temperature at t=60 sec vs Usable Capacity'),
    ],
    
    'Pressure Plots': [
        ('Maximum Pressure', 'PSIG', 'Usable Capacity', '%', 'Max_Pressure_vs_UsableCapacity', 'max_pressure', 'Maximum Pressure vs Usable Capacity'),
        ('Pressure at t=60 sec', 'PSIG', 'Usable Capacity', '%', 'Pressure_60sec_vs_UsableCapacity', 'pressure_60sec', 'Pressure at t=60 sec vs Usable Capacity'),
    ],
    
    'Energy Plots': [
        ('Energy Released by Cell to Max T_Cell_Avg', 'Wh', 'Usable Capacity', '%', 'Energy_Released_Cell_vs_UsableCapacity', 'energy', 'Energy Released by Cell vs Usable Capacity'),
        ('Total Energy Released to Max T_Cell_Avg', 'Wh', 'Usable Capacity', '%', 'Total_Energy_Released_vs_UsableCapacity', 'total_energy', 'Total Energy Released vs Usable Capacity'),
        ('Total Energy Released to Max T_Cell_Avg', 'Wh', 'Cell specific Mass Loss', 'g', 'Total_Energy_Released_vs_Cell_Mass_Loss', 'total_energy', 'Total Energy Released vs Cell Mass Loss'),
        ('Total Energy Released to Max T_Cell_Avg', 'Wh', 'Maximum Pressure', 'PSIG', 'Total_Energy_Released_vs_Max_Pressure', 'total_energy', 'Total Energy Released vs Maximum Pressure'),
        ('Total Energy Released to Max T_Cell_Avg', 'Wh', 'Cell Carcass Mass', 'g', 'Total_Energy_Released_vs_Cell_Carcass_Mass', 'total_energy', 'Total Energy Released vs Cell Carcass Mass'),
        ('Total Energy Released to Max T_Cell_Avg', 'Wh', 'Particulate Mass', 'g', 'Total_Energy_Released_vs_Particulate_Mass', 'total_energy', 'Total Energy Released vs Particulate Mass'),
    ],
    
    'Mass & Gas Generation': [
        ('Cell specific Mass Loss', 'g', 'Usable Capacity', '%', 'Cell_Mass_Loss_vs_UsableCapacity', 'mass', 'Cell specific Mass Loss vs Usable Capacity'),
        ('Cell Carcass Mass', 'g', 'Usable Capacity', '%', 'Cell_Carcass_Mass_vs_UsableCapacity', 'mass', 'Cell Carcass Mass vs Usable Capacity'),
        ('Particulate Mass', 'g', 'Usable Capacity', '%', 'Particulate_Mass_vs_UsableCapacity', 'mass', 'Particulate Mass vs Usable Capacity'),
        ('Gas generation', 'g', 'Usable Capacity', '%', 'Gas_Generation_vs_UsableCapacity', 'gas', 'Gas Generation vs Usable Capacity'),
        ('Gas rate', 'mmol/Wh', 'Usable Capacity', '%', 'Gas_Rate_mmol_Wh_vs_UsableCapacity', 'gas_rate', 'Gas Rate (mmol/Wh) vs Usable Capacity'),
        ('Gas rate', 'sL/kWh', 'Usable Capacity', '%', 'Gas_Rate_sL_kWh_vs_UsableCapacity', 'gas_rate', 'Gas Rate (sL/kWh) vs Usable Capacity'),
        ('Particle Size Distribution (50th Percentile)', 'μm', 'Usable Capacity', '%', 'Particle_Size_Distribution_50th_Percentile_vs_UsableCapacity', 'particle_size', 'Particle Size Distribution (50th Percentile) vs Usable Capacity'),
    ],
    
    'Major Gas Emissions': [
        ('Carbon Dioxide', 'vol %', 'Usable Capacity', '%', 'Carbon_Dioxide_vs_UsableCapacity', 'emissions', 'Carbon Dioxide vs Usable Capacity'),
        ('CH4', 'vol %', 'Usable Capacity', '%', 'CH4_vs_UsableCapacity', 'emissions', 'CH4 vs Usable Capacity'),
        ('Carbon Monoxide', 'vol %', 'Usable Capacity', '%', 'Carbon_Monoxide_vs_UsableCapacity', 'emissions', 'Carbon Monoxide vs Usable Capacity'),
        ('Oxygen', 'vol %', 'Usable Capacity', '%', 'Oxygen_vs_UsableCapacity', 'emissions', 'Oxygen vs Usable Capacity'),
        ('Nitrogen', 'vol %', 'Usable Capacity', '%', 'Nitrogen_vs_UsableCapacity', 'emissions', 'Nitrogen vs Usable Capacity'),
        ('Hydrogen', 'vol %', 'Usable Capacity', '%', 'Hydrogen_vs_UsableCapacity', 'emissions', 'Hydrogen vs Usable Capacity'),
    ],
    
    'Electrolytes': [
        ('Identified Electrolytes', 'vol %', 'Usable Capacity', '%', 'Identified_Electrolytes_vs_UsableCapacity', 'electrolytes', 'Identified Electrolytes vs Usable Capacity'),
        ('Unknown C2-C4*', 'vol %', 'Usable Capacity', '%', 'Unknown_C2_C4_vs_UsableCapacity', 'electrolytes', 'Unknown C2-C4* vs Usable Capacity'),
        ('C5-C12**', 'vol %', 'Usable Capacity', '%', 'C5_C12_vs_UsableCapacity', 'electrolytes', 'C5-C12** vs Usable Capacity'),
    ],
    
    'C2 Hydrocarbons': [
        ('ETHANE', 'vol %', 'Usable Capacity', '%', 'ETHANE_vs_UsableCapacity', 'hydrocarbons', 'ETHANE vs Usable Capacity'),
        ('ETHYLENE', 'vol %', 'Usable Capacity', '%', 'ETHYLENE_vs_UsableCapacity', 'hydrocarbons', 'ETHYLENE vs Usable Capacity'),
        ('ACETYLENE', 'vol %', 'Usable Capacity', '%', 'ACETYLENE_vs_UsableCapacity', 'hydrocarbons', 'ACETYLENE vs Usable Capacity'),
    ],
    
    'C3 Hydrocarbons': [
        ('PROPANE', 'vol %', 'Usable Capacity', '%', 'PROPANE_vs_UsableCapacity', 'hydrocarbons', 'PROPANE vs Usable Capacity'),
        ('PROPYLENE', 'vol %', 'Usable Capacity', '%', 'PROPYLENE_vs_UsableCapacity', 'hydrocarbons', 'PROPYLENE vs Usable Capacity'),
    ],
    
    'C4 Hydrocarbons': [
        ('BUTANE', 'vol %', 'Usable Capacity', '%', 'BUTANE_vs_UsableCapacity', 'hydrocarbons', 'BUTANE vs Usable Capacity'),
        ('TRANS-2-BUTENE', 'vol %', 'Usable Capacity', '%', 'TRANS_2_BUTENE_vs_UsableCapacity', 'hydrocarbons', 'TRANS-2-BUTENE vs Usable Capacity'),
        ('1-BUTENE', 'vol %', 'Usable Capacity', '%', '1_BUTENE_vs_UsableCapacity', 'hydrocarbons', '1-BUTENE vs Usable Capacity'),
        ('2-METHYLPROPENE (ISOBUTYLENE)', 'vol %', 'Usable Capacity', '%', '2_METHYLPROPENE_ISOBUTYLENE_vs_UsableCapacity', 'hydrocarbons', '2-METHYLPROPENE (ISOBUTYLENE) vs Usable Capacity'),
        ('1,3-BUTADIENE', 'vol %', 'Usable Capacity', '%', '1_3_BUTADIENE_vs_UsableCapacity', 'hydrocarbons', '1,3-BUTADIENE vs Usable Capacity'),
        ('2-METHYLPROPANE (ISOBUTANE)', 'vol %', 'Usable Capacity', '%', '2_METHYLPROPANE_ISOBUTANE_vs_UsableCapacity', 'hydrocarbons', '2-METHYLPROPANE (ISOBUTANE) vs Usable Capacity'),
        ('CIS-2-BUTENE', 'vol %', 'Usable Capacity', '%', 'CIS_2_BUTENE_vs_UsableCapacity', 'hydrocarbons', 'CIS-2-BUTENE vs Usable Capacity'),
    ],
    
    'C5 Hydrocarbons': [
        ('2-METHYLBUTANE (ISOPENTANE)', 'vol %', 'Usable Capacity', '%', '2_METHYLBUTANE_ISOPENTANE_vs_UsableCapacity', 'hydrocarbons', '2-METHYLBUTANE (ISOPENTANE) vs Usable Capacity'),
    ],
}

print(f"Total available plots: {sum(len(plots) for plots in ALL_PLOT_PARAMS.values())}")

# Create plot selection dialog
class PlotSelectionDialog:
    def __init__(self, parent):
        print("Creating plot selection dialog...")
        self.parent = parent
        self.result = []
        self.checkboxes = {}
        self.category_vars = {}
        self.fitting_line_12_var = None
        self.fitting_line_134_var = None
        self.show_cov_var = None
        
        # Create dialog window
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Select Plots to Generate")
        self.dialog.geometry("700x900")
        
        # Make sure dialog appears on top
        self.dialog.lift()
        self.dialog.focus_force()
        
        # Main frame with scrollbar
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.dialog.columnconfigure(0, weight=1)
        self.dialog.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Select Scatter Plots to Generate", 
                               font=('Arial', 14, 'bold'))
        title_label.grid(row=0, column=0, pady=10, sticky=tk.W)
        
        # Create canvas with scrollbar
        canvas = tk.Canvas(main_frame, borderwidth=0, background="#f0f0f0")
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack canvas and scrollbar
        canvas.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        
        # Add checkboxes grouped by category
        row = 0
        for category, plots in ALL_PLOT_PARAMS.items():
            # Category frame
            category_frame = ttk.LabelFrame(scrollable_frame, text=category, padding="10")
            category_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)
            category_frame.columnconfigure(0, weight=1)
            
            # Category select all checkbox
            cat_var = tk.BooleanVar(value=False)
            self.category_vars[category] = cat_var
            cat_checkbox = ttk.Checkbutton(
                category_frame, 
                text=f"Select All {category}",
                variable=cat_var,
                command=lambda c=category: self.toggle_category(c)
            )
            cat_checkbox.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
            
            # Individual plot checkboxes
            self.checkboxes[category] = []
            for idx, plot in enumerate(plots, start=1):
                var = tk.BooleanVar(value=False)
                cb = ttk.Checkbutton(
                    category_frame, 
                    text=plot[6],  # display_name
                    variable=var
                )
                cb.grid(row=idx, column=0, sticky=tk.W, padx=20)
                self.checkboxes[category].append((var, plot))
            
            row += 1
        
        # Fitting Line Options Frame
        fitting_frame = ttk.LabelFrame(scrollable_frame, text="Fitting Line Options", padding="10")
        fitting_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)
        fitting_frame.columnconfigure(0, weight=1)
        
        # Fitting line checkboxes
        self.fitting_line_12_var = tk.BooleanVar(value=False)
        fitting_12_cb = ttk.Checkbutton(
            fitting_frame,
            text="Add Fitting Line: Run1-2 (Linear fit through Run1 and Run2 data points)",
            variable=self.fitting_line_12_var
        )
        fitting_12_cb.grid(row=0, column=0, sticky=tk.W, pady=2)
        
        self.fitting_line_134_var = tk.BooleanVar(value=False)
        fitting_134_cb = ttk.Checkbutton(
            fitting_frame,
            text="Add Fitting Line: Run1-3-4 (Linear fit through Run1, Run3, and Run4 data points)",
            variable=self.fitting_line_134_var
        )
        fitting_134_cb.grid(row=1, column=0, sticky=tk.W, pady=2)
        
        row += 1
        
        # Statistical Analysis Frame
        stats_frame = ttk.LabelFrame(scrollable_frame, text="Statistical Analysis Options", padding="10")
        stats_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)
        stats_frame.columnconfigure(0, weight=1)
        
        # CoV checkbox
        self.show_cov_var = tk.BooleanVar(value=False)
        cov_cb = ttk.Checkbutton(
            stats_frame,
            text="Show Coefficient of Variation (CoV) for each Run (calculated from y-axis values only)",
            variable=self.show_cov_var
        )
        cov_cb.grid(row=0, column=0, sticky=tk.W, pady=2)
        
        row += 1
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        # Select All button
        select_all_btn = ttk.Button(button_frame, text="Select All", 
                                    command=self.select_all, width=15)
        select_all_btn.grid(row=0, column=0, padx=5)
        
        # Deselect All button
        deselect_all_btn = ttk.Button(button_frame, text="Deselect All", 
                                      command=self.deselect_all, width=15)
        deselect_all_btn.grid(row=0, column=1, padx=5)
        
        # OK button
        ok_btn = ttk.Button(button_frame, text="Generate Plots", 
                           command=self.ok, width=15)
        ok_btn.grid(row=0, column=2, padx=5)
        
        # Cancel button
        cancel_btn = ttk.Button(button_frame, text="Cancel", 
                               command=self.cancel, width=15)
        cancel_btn.grid(row=0, column=3, padx=5)
        
        # Bind mouse wheel for scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Center the dialog
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (self.dialog.winfo_width() // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f'+{x}+{y}')
        
        print("Dialog created and displayed")
    
    def toggle_category(self, category):
        """Toggle all checkboxes in a category"""
        state = self.category_vars[category].get()
        for var, plot in self.checkboxes[category]:
            var.set(state)
    
    def select_all(self):
        """Select all plots"""
        print("Selecting all plots...")
        for category in self.checkboxes:
            self.category_vars[category].set(True)
            for var, plot in self.checkboxes[category]:
                var.set(True)
    
    def deselect_all(self):
        """Deselect all plots"""
        print("Deselecting all plots...")
        for category in self.checkboxes:
            self.category_vars[category].set(False)
            for var, plot in self.checkboxes[category]:
                var.set(False)
    
    def ok(self):
        """Collect selected plots and close dialog"""
        print("Collecting selected plots...")
        self.result = []
        for category in self.checkboxes:
            for var, plot in self.checkboxes[category]:
                if var.get():
                    self.result.append(plot)
        
        if not self.result:
            messagebox.showwarning("No Selection", "Please select at least one plot to generate.")
            return
        
        print(f"Selected {len(self.result)} plots")
        print(f"Fitting Line Run1-2: {self.fitting_line_12_var.get()}")
        print(f"Fitting Line Run1-3-4: {self.fitting_line_134_var.get()}")
        print(f"Show CoV: {self.show_cov_var.get()}")
        self.dialog.destroy()
    
    def cancel(self):
        """Cancel and close dialog"""
        print("Dialog cancelled")
        self.result = []
        self.dialog.destroy()
    
    def show(self):
        """Show dialog and return selected plots and options"""
        print("Waiting for user selection...")
        self.parent.wait_window(self.dialog)
        return self.result, self.fitting_line_12_var.get(), self.fitting_line_134_var.get(), self.show_cov_var.get()

# Create root window for dialogs
print("Initializing GUI...")
root = Tk()
root.withdraw()

# Show plot selection dialog
try:
    dialog = PlotSelectionDialog(root)
    selected_plots, add_fitting_12, add_fitting_134, show_cov = dialog.show()
except Exception as e:
    print(f"Error creating dialog: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

if not selected_plots:
    print("No plots selected. Exiting...")
    root.destroy()
    sys.exit(0)

print(f"\n{len(selected_plots)} plots selected for generation.")
print(f"Fitting lines enabled: Run1-2={add_fitting_12}, Run1-3-4={add_fitting_134}")
print(f"Show CoV: {show_cov}\n")

# Ask user to select output folder
print("Opening folder selection dialog...")
root.deiconify()
root.lift()
root.focus_force()
root.withdraw()

output_dir = filedialog.askdirectory(
    title="Select Output Folder for Plots",
    parent=root
)

if not output_dir:
    print("No folder selected. Exiting...")
    root.destroy()
    sys.exit(0)

print(f"Plots will be saved to: {output_dir}\n")

# File path
file_path = r'P:\Programs\EVESE\03-29378 EVESE-II\Working Files\C4 - Cell Abuse Repeatability\Data\C4_Data-Summary.xlsb'
sheet_name = 'Comparison_analysis'

print("Reading data file...")

# Try reading with appropriate engine
try:
    if has_pyxlsb and file_path.endswith('.xlsb') and os.path.exists(file_path):
        print(f"Reading {file_path}...")
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='pyxlsb', header=None)
        print(f"Successfully loaded file with shape: {df.shape}")
    else:
        # Ask for alternative file if xlsb can't be read
        print("Please select the Excel file (.xlsx or .xlsb)...")
        root.deiconify()
        root.lift()
        root.focus_force()
        root.withdraw()
        
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xlsb"), ("All files", "*.*")],
            parent=root
        )
        if not file_path:
            print("No file selected. Exiting...")
            root.destroy()
            sys.exit(0)
        
        print(f"Reading {file_path}...")
        if file_path.endswith('.xlsb'):
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='pyxlsb', header=None)
        else:
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl', header=None)
        print(f"Successfully loaded file with shape: {df.shape}")
        
except Exception as e:
    print(f"Error reading file: {e}")
    print("\nTrying to open file selection dialog...")
    
    root.deiconify()
    root.lift()
    root.focus_force()
    root.withdraw()
    
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xlsb"), ("All files", "*.*")],
        parent=root
    )
    if not file_path:
        print("No file selected. Exiting...")
        root.destroy()
        sys.exit(0)
    
    print(f"Reading {file_path}...")
    if file_path.endswith('.xlsb'):
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='pyxlsb', header=None)
    else:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl', header=None)
    print(f"Successfully loaded file with shape: {df.shape}")

print(f"Successfully loaded: {os.path.basename(file_path)}")
print(f"Data shape: {df.shape}\n")

# Extract headers
print("Processing headers...")
header_row1 = df.iloc[0].fillna('')  # Test1, Test2, Test3, etc.
header_row2 = df.iloc[1].fillna('')  # (Run1), (Run2), etc.

# Create column mapping
column_mapping = {}
for i in range(len(header_row1)):
    if i == 0:
        column_mapping[i] = 'Parameter'
    elif i == 1:
        column_mapping[i] = 'Unit'
    else:
        test = str(header_row1[i]).strip()
        run = str(header_row2[i]).strip().replace('(', '').replace(')', '')
        column_mapping[i] = f'{test}_{run}'

# Start data from row 2 (index 2)
df_data = df.iloc[2:].reset_index(drop=True)
df_data.columns = [column_mapping[i] for i in range(len(df_data.columns))]

print(f"Processed data shape: {df_data.shape}")
print(f"Columns: {list(df_data.columns)[:5]}... (showing first 5)\n")

# Define symbols for Test 1, 2, 3
test_markers = {
    'Test1': 'o',  # circle
    'Test2': 's',  # square
    'Test3': '^'   # triangle
}

# Define colors for Run 1, 2, 3, 4
run_colors = {
    'Run1': '#1f77b4',  # blue
    'Run2': '#ff7f0e',  # orange
    'Run3': '#2ca02c',  # green
    'Run4': '#d62728'   # red
}

# Extract data structure
runs = ['Run1', 'Run2', 'Run3', 'Run4']
tests = ['Test1', 'Test2', 'Test3']

# Use only selected plots
plot_params = selected_plots

# Function to find row index for a parameter with specific unit
def find_parameter_row(df_data, param_name, unit):
    for idx, row in df_data.iterrows():
        param = str(row['Parameter']).strip()
        unit_col = str(row['Unit']).strip()
        
        # Exact match first
        if param == param_name and (unit in unit_col or unit_col == unit):
            return idx
        
        # Partial match for parameters that might have slight variations
        if param_name in param and (unit in unit_col or unit_col == unit):
            return idx
    return None

# Function to calculate coefficient of variation
def calculate_cov(values):
    """Calculate coefficient of variation (CoV) as percentage"""
    if len(values) < 2:
        return None
    mean_val = np.mean(values)
    if mean_val == 0:
        return None
    std_val = np.std(values, ddof=1)  # Sample standard deviation
    cov = (std_val / mean_val) * 100
    return cov

# Function to collect data points for a parameter pair
def collect_data_points(df_data, y_param_name, y_unit, x_param_name, x_unit):
    y_param_idx = find_parameter_row(df_data, y_param_name, y_unit)
    x_param_idx = find_parameter_row(df_data, x_param_name, x_unit)
    
    if y_param_idx is None or x_param_idx is None:
        return []
    
    data_points = []
    for run in runs:
        for test in tests:
            try:
                col_name = f'{test}_{run}'
                if col_name not in df_data.columns:
                    continue
                
                x_val = df_data.iloc[x_param_idx][col_name]
                y_val = df_data.iloc[y_param_idx][col_name]
                
                try:
                    x_val = float(x_val)
                    y_val = float(y_val)
                except (ValueError, TypeError):
                    continue
                
                if pd.notna(x_val) and pd.notna(y_val):
                    data_points.append((x_val, y_val, test, run))
            except Exception as e:
                continue
    
    return data_points

print("Collecting data and determining y-axis limits...")

# Collect data for temperature plots
temp_data = []
for plot in plot_params:
    y_param, y_unit, x_param, x_unit, file_name, category, display_name = plot
    if category == 'temperature':
        points = collect_data_points(df_data, y_param, y_unit, x_param, x_unit)
        temp_data.extend([y for x, y, t, r in points])

# Collect data for max pressure plots
max_pressure_data = []
for plot in plot_params:
    y_param, y_unit, x_param, x_unit, file_name, category, display_name = plot
    if category == 'max_pressure':
        points = collect_data_points(df_data, y_param, y_unit, x_param, x_unit)
        max_pressure_data.extend([y for x, y, t, r in points])

# Collect data for pressure at 60sec plots
pressure_60sec_data = []
for plot in plot_params:
    y_param, y_unit, x_param, x_unit, file_name, category, display_name = plot
    if category == 'pressure_60sec':
        points = collect_data_points(df_data, y_param, y_unit, x_param, x_unit)
        pressure_60sec_data.extend([y for x, y, t, r in points])

# Calculate y-limits with some padding
def calculate_ylim(data):
    if not data:
        return None, None
    min_val = min(data)
    max_val = max(data)
    range_val = max_val - min_val
    padding = range_val * 0.1  # 10% padding
    return min_val - padding, max_val + padding

temp_ylim = calculate_ylim(temp_data)
max_pressure_ylim = calculate_ylim(max_pressure_data)
pressure_60sec_ylim = calculate_ylim(pressure_60sec_data)

if temp_ylim[0] is not None:
    print(f"Temperature y-limits: {temp_ylim}")
if max_pressure_ylim[0] is not None:
    print(f"Maximum Pressure y-limits: {max_pressure_ylim}")
if pressure_60sec_ylim[0] is not None:
    print(f"Pressure at 60sec y-limits: {pressure_60sec_ylim}")
print()

print("Generating scatter plots...")
print("=" * 60)

# Close matplotlib interactive mode to prevent hanging
plt.ioff()

# Create a figure for each parameter
plot_count = 0
skipped_count = 0
for plot_idx, plot in enumerate(plot_params, start=1):
    y_param_name, y_unit, x_param_name, x_unit, file_name, category, display_name = plot
    
    print(f"\nProcessing plot {plot_idx}/{len(plot_params)}: {display_name}")
    
    # Create figure with 12x8 inch size
    fig, ax = plt.subplots(figsize=(12, 8))
    
    # Find the row indices
    y_param_idx = find_parameter_row(df_data, y_param_name, y_unit)
    x_param_idx = find_parameter_row(df_data, x_param_name, x_unit)
    
    if y_param_idx is None:
        print(f"  ⚠ Warning: Could not find parameter '{y_param_name}' with unit '{y_unit}'")
        skipped_count += 1
        plt.close(fig)
        continue
    
    if x_param_idx is None:
        print(f"  ⚠ Warning: Could not find parameter '{x_param_name}' with unit '{x_unit}'")
        skipped_count += 1
        plt.close(fig)
        continue
    
    # Collect all data points for this parameter
    points_plotted = 0
    
    # Store data points by run for fitting lines and CoV
    run_data = {'Run1': {'x': [], 'y': []}, 
                'Run2': {'x': [], 'y': []}, 
                'Run3': {'x': [], 'y': []}, 
                'Run4': {'x': [], 'y': []}}
    
    for run in runs:
        for test in tests:
            try:
                col_name = f'{test}_{run}'
                
                if col_name not in df_data.columns:
                    continue
                
                x_val = df_data.iloc[x_param_idx][col_name]
                y_val = df_data.iloc[y_param_idx][col_name]
                
                # Convert to numeric, skip if invalid
                try:
                    x_val = float(x_val)
                    y_val = float(y_val)
                except (ValueError, TypeError):
                    continue
                
                # Skip if values are NaN
                if pd.notna(x_val) and pd.notna(y_val):
                    ax.scatter(x_val, y_val, 
                             marker=test_markers[test], 
                             color=run_colors[run],
                             s=150, 
                             alpha=0.7,
                             edgecolors='black',
                             linewidth=1.5,
                             zorder=3)
                    points_plotted += 1
                    
                    # Store data for fitting lines and CoV
                    run_data[run]['x'].append(x_val)
                    run_data[run]['y'].append(y_val)
            except Exception as e:
                continue
    
    if points_plotted == 0:
        print(f"  ⚠ No data points found for {y_param_name} vs {x_param_name}")
        skipped_count += 1
        plt.close(fig)
        continue
    
    print(f"  Plotting {points_plotted} data points...")
    
    # Store fitting line information for text display
    fitting_equations = []
    
    # Add fitting lines if requested
    if add_fitting_12:
        # Fitting line for Run1 and Run2
        x_12 = run_data['Run1']['x'] + run_data['Run2']['x']
        y_12 = run_data['Run1']['y'] + run_data['Run2']['y']
        
        if len(x_12) >= 2:
            try:
                slope, intercept, r_value, p_value, std_err = stats.linregress(x_12, y_12)
                x_fit = np.array([min(x_12), max(x_12)])
                y_fit = slope * x_fit + intercept
                
                ax.plot(x_fit, y_fit, 
                       color='purple', 
                       linestyle='--', 
                       linewidth=2.5, 
                       alpha=0.8,
                       label='Fit: Run1-2',
                       zorder=2)
                
                # Store equation for text display
                sign = '+' if intercept >= 0 else ''
                fitting_equations.append({
                    'label': 'Run1-2',
                    'equation': f'y = {slope:.4f}x {sign}{intercept:.4f}',
                    'r_squared': f'R² = {r_value**2:.4f}',
                    'color': 'purple'
                })
                print(f"  ✓ Added fitting line Run1-2: R²={r_value**2:.4f}")
            except Exception as e:
                print(f"  ⚠ Could not create fitting line Run1-2: {e}")
    
    if add_fitting_134:
        # Fitting line for Run1, Run3, and Run4
        x_134 = run_data['Run1']['x'] + run_data['Run3']['x'] + run_data['Run4']['x']
        y_134 = run_data['Run1']['y'] + run_data['Run3']['y'] + run_data['Run4']['y']
        
        if len(x_134) >= 2:
            try:
                slope, intercept, r_value, p_value, std_err = stats.linregress(x_134, y_134)
                x_fit = np.array([min(x_134), max(x_134)])
                y_fit = slope * x_fit + intercept
                
                ax.plot(x_fit, y_fit, 
                       color='brown', 
                       linestyle='-.', 
                       linewidth=2.5, 
                       alpha=0.8,
                       label='Fit: Run1-3-4',
                       zorder=2)
                
                # Store equation for text display
                sign = '+' if intercept >= 0 else ''
                fitting_equations.append({
                    'label': 'Run1-3-4',
                    'equation': f'y = {slope:.4f}x {sign}{intercept:.4f}',
                    'r_squared': f'R² = {r_value**2:.4f}',
                    'color': 'brown'
                })
                print(f"  ✓ Added fitting line Run1-3-4: R²={r_value**2:.4f}")
            except Exception as e:
                print(f"  ⚠ Could not create fitting line Run1-3-4: {e}")
    
    # Calculate CoV for each run if requested
    cov_values = []
    if show_cov:
        for run in runs:
            if len(run_data[run]['y']) >= 2:
                cov = calculate_cov(run_data[run]['y'])
                if cov is not None:
                    cov_values.append({
                        'run': run,
                        'cov': cov,
                        'color': run_colors[run]
                    })
                    print(f"  ✓ {run} CoV: {cov:.2f}%")
    
    # Customize plot
    ax.set_xlabel(f'{x_param_name} ({x_unit})', fontsize=14, fontweight='bold')
    ax.set_ylabel(f'{y_param_name} ({y_unit})', fontsize=14, fontweight='bold')
    
    # Adjust title for long parameter names
    title_text = f'{y_param_name} vs {x_param_name}'
    if len(title_text) > 60:
        # Split long titles into two lines
        ax.set_title(f'{y_param_name}\nvs {x_param_name}\nEVESE C4 - Thermal Runaway Analysis', 
                     fontsize=14, fontweight='bold', pad=20)
    else:
        ax.set_title(f'{title_text}\nEVESE C4 - Thermal Runaway Analysis', 
                     fontsize=16, fontweight='bold', pad=20)
    
    ax.grid(True, alpha=0.3, linestyle='--', linewidth=0.8, zorder=1)
    
    # Set y-limits based on category
    if category == 'temperature' and temp_ylim[0] is not None:
        ax.set_ylim(temp_ylim)
    elif category == 'max_pressure' and max_pressure_ylim[0] is not None:
        ax.set_ylim(max_pressure_ylim)
    elif category == 'pressure_60sec' and pressure_60sec_ylim[0] is not None:
        ax.set_ylim(pressure_60sec_ylim)
    
    # Add minor gridlines
    ax.minorticks_on()
    ax.grid(which='minor', alpha=0.15, linestyle=':', linewidth=0.5, zorder=1)
    
    # Create custom legend
    from matplotlib.lines import Line2D
    
    # Legend for runs (colors)
    run_legend_elements = [Line2D([0], [0], marker='o', color='w', 
                                   markerfacecolor=run_colors[run], 
                                   markersize=11, label=run,
                                   markeredgecolor='black', markeredgewidth=1.5)
                          for run in runs]
    
    # Legend for tests (markers)
    test_legend_elements = [Line2D([0], [0], marker=test_markers[test], color='w', 
                                    markerfacecolor='gray', 
                                    markersize=11, label=test,
                                    markeredgecolor='black', markeredgewidth=1.5)
                           for test in tests]
    
    # Combine both legends side by side
    combined_legend_elements = run_legend_elements + test_legend_elements
    combined_labels = [run for run in runs] + [test for test in tests]
    
    # Add fitting line legend entries if they exist
    if add_fitting_12 and len(x_12) >= 2:
        combined_legend_elements.append(Line2D([0], [0], color='purple', linestyle='--', linewidth=2.5))
        combined_labels.append('Fit: Run1-2')
    
    if add_fitting_134 and len(x_134) >= 2:
        combined_legend_elements.append(Line2D([0], [0], color='brown', linestyle='-.', linewidth=2.5))
        combined_labels.append('Fit: Run1-3-4')
    
    # Calculate number of columns for legend
    ncol = 7 if not (add_fitting_12 or add_fitting_134) else 9
    
    # Create single legend below the plot, centered
    legend = ax.legend(handles=combined_legend_elements,
                      labels=combined_labels,
                      loc='upper center',
                      bbox_to_anchor=(0.5, -0.08),
                      ncol=ncol,
                      frameon=True,
                      framealpha=0.95,
                      edgecolor='black',
                      fontsize=10,
                      columnspacing=1.5,
                      handletextpad=0.5)
    
    # Add fitting equations below the legend with minimal spacing
    if fitting_equations:
        # Get legend bounding box in figure coordinates
        fig.canvas.draw()
        legend_bbox = legend.get_window_extent().transformed(fig.transFigure.inverted())
        
        # Start equations just below the legend with very minimal gap
        equation_y = legend_bbox.y0 - 0.015
        
        for eq_info in fitting_equations:
            equation_text = f"{eq_info['label']}: {eq_info['equation']},  {eq_info['r_squared']}"
            fig.text(0.5, equation_y, equation_text,
                    ha='center', va='top',
                    fontsize=11, color=eq_info['color'],
                    fontweight='bold',
                    transform=fig.transFigure)
            equation_y -= 0.022
    
    # Add CoV values below fitting equations if requested (centralized like equations)
    if cov_values:
        # If no fitting equations, get legend bbox
        if not fitting_equations:
            fig.canvas.draw()
            legend_bbox = legend.get_window_extent().transformed(fig.transFigure.inverted())
            cov_y = legend_bbox.y0 - 0.015
        else:
            cov_y = equation_y - 0.005  # Small gap after fitting equations
        
        # Build CoV text line with colored segments using matplotlib's rich text capability
        from matplotlib import patches
        
        # Create the full text string first
        cov_text_parts = []
        for cov_info in cov_values:
            cov_text_parts.append(f"{cov_info['run']}: {cov_info['cov']:.2f}%")
        
        full_cov_text = "CoV: " + "   ".join(cov_text_parts)
        
        # Use a single centered text with manual color coding using matplotlib Text objects
        # Create label part
        label_text = "CoV:  "
        
        # Calculate approximate character widths for positioning
        char_width = 0.0065  # Approximate character width in figure coordinates
        
        # Start position (centered, then adjust)
        total_text_width = len(full_cov_text) * char_width
        start_x = 0.5 - (total_text_width / 2)
        current_x = start_x
        
        # Add label in black
        fig.text(current_x, cov_y, label_text,
                ha='left', va='top',
                fontsize=11, color='black',
                fontweight='bold',
                transform=fig.transFigure)
        current_x += len(label_text) * char_width
        
        # Add each run's CoV in its respective color
        for idx, cov_info in enumerate(cov_values):
            if idx > 0:
                separator = "    "
                fig.text(current_x, cov_y, separator,
                        ha='left', va='top',
                        fontsize=11, color='black',
                        fontweight='bold',
                        transform=fig.transFigure)
                current_x += len(separator) * char_width
            
            run_text = f"{cov_info['run']}: {cov_info['cov']:.2f}%"
            fig.text(current_x, cov_y, run_text,
                    ha='left', va='top',
                    fontsize=11, color=cov_info['color'],
                    fontweight='bold',
                    transform=fig.transFigure)
            current_x += len(run_text) * char_width
    
    # Adjust layout to make room for legend, equations, and CoV below
    extra_lines = len(fitting_equations) + (1 if cov_values else 0)
    bottom_margin = 0.08 if extra_lines == 0 else 0.08 + (extra_lines * 0.022)
    plt.subplots_adjust(bottom=bottom_margin)
    
    # Save figure with descriptive name
    if add_fitting_12 or add_fitting_134:
        output_filename = f'Fitted_{file_name}.png'
    else:
        output_filename = f'{file_name}.png'
    
    output_path = os.path.join(output_dir, output_filename)
    
    print(f"  Saving to: {output_filename}")
    plt.savefig(output_path, dpi=300, bbox_inches='tight', facecolor='white')
    
    plot_count += 1
    print(f"  ✓ Saved successfully!")
    
    # Close figure to free memory
    plt.close(fig)

print(f"\n{'='*60}")
print(f"SUCCESS! {plot_count} scatter plots generated successfully!")
if skipped_count > 0:
    print(f"⚠ Skipped {skipped_count} plots due to missing data or parameters")
print(f"Location: {output_dir}")
print(f"{'='*60}")

# Close the root window
root.destroy()
print("\nScript completed successfully!")
