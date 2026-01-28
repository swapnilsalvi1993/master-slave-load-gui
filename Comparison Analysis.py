import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
from tkinter import Tk, filedialog

# Try to import pyxlsb, provide helpful error if not available
try:
    import pyxlsb
    has_pyxlsb = True
except ImportError:
    has_pyxlsb = False
    print("⚠ Warning: pyxlsb not installed.")
    print("To install, run: pip install pyxlsb")
    print("\nAlternative: Save your .xlsb file as .xlsx and update the file path in the script.")
    response = input("\nDo you want to continue and select an .xlsx file instead? (y/n): ")
    if response.lower() != 'y':
        print("Exiting. Please install pyxlsb using: pip install pyxlsb")
        exit()

# Hide the main tkinter window
root = Tk()
root.withdraw()
root.attributes('-topmost', True)

# Ask user to select output folder
print("\nPlease select the folder where you want to save the plots...")
output_dir = filedialog.askdirectory(title="Select Output Folder for Plots")

if not output_dir:
    print("No folder selected. Exiting...")
    exit()

print(f"Plots will be saved to: {output_dir}\n")

# File path
file_path = r'P:\Programs\EVESE\03-29378 EVESE-II\Working Files\C4 - Cell Abuse Repeatability\Data\C4_Data-Summary.xlsb'
sheet_name = 'Comparison_analysis'

print("Reading data file...")

# Try reading with appropriate engine
try:
    if has_pyxlsb and file_path.endswith('.xlsb'):
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='pyxlsb', header=None)
    else:
        # Ask for alternative file if xlsb can't be read
        print("Please select the Excel file (.xlsx or .xlsb)...")
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xlsb"), ("All files", "*.*")]
        )
        if not file_path:
            print("No file selected. Exiting...")
            exit()
        
        if file_path.endswith('.xlsb'):
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='pyxlsb', header=None)
        else:
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl', header=None)
except Exception as e:
    print(f"Error reading file: {e}")
    print("\nTrying to open file selection dialog...")
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xlsb"), ("All files", "*.*")]
    )
    if not file_path:
        print("No file selected. Exiting...")
        exit()
    
    if file_path.endswith('.xlsb'):
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='pyxlsb', header=None)
    else:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl', header=None)

print(f"Successfully loaded: {os.path.basename(file_path)}\n")

# Extract headers
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

# Parameters to plot
# Format: (y_param_name, y_unit, x_param_name, x_unit, short_name_for_file, category_for_ylim)
plot_params = [
    # Original plots (y-axis vs Usable Capacity)
    ('Energy Released by Cell to Max T_Cell_Avg', 'Wh', 'Usable Capacity', '%', 'Energy_Released_Cell_vs_UsableCapacity', 'energy'),
    ('Maximum Cell Avg Temperature', '°C', 'Usable Capacity', '%', 'Max_Cell_Temperature_vs_UsableCapacity', 'temperature'),
    ('Cell Avg Temperature at t=60 sec', '°C', 'Usable Capacity', '%', 'Cell_Temperature_60sec_vs_UsableCapacity', 'temperature'),
    ('Maximum Pressure', 'PSIG', 'Usable Capacity', '%', 'Max_Pressure_vs_UsableCapacity', 'pressure'),
    ('Pressure at t=60 sec', 'PSIG', 'Usable Capacity', '%', 'Pressure_60sec_vs_UsableCapacity', 'pressure'),
    ('Cell specific Mass Loss', 'g', 'Usable Capacity', '%', 'Cell_Mass_Loss_vs_UsableCapacity', 'mass'),
    ('Gas generation', 'g', 'Usable Capacity', '%', 'Gas_Generation_vs_UsableCapacity', 'gas'),
    
    # New plots (Total Energy Released vs various parameters)
    ('Total Energy Released to Max T_Cell_Avg', 'Wh', 'Usable Capacity', '%', 'Total_Energy_Released_vs_UsableCapacity', 'total_energy'),
    ('Total Energy Released to Max T_Cell_Avg', 'Wh', 'Cell specific Mass Loss', 'g', 'Total_Energy_Released_vs_Cell_Mass_Loss', 'total_energy'),
    ('Total Energy Released to Max T_Cell_Avg', 'Wh', 'Maximum Pressure', 'PSIG', 'Total_Energy_Released_vs_Max_Pressure', 'total_energy'),
]

# Function to find row index for a parameter with specific unit
def find_parameter_row(df_data, param_name, unit):
    for idx, row in df_data.iterrows():
        param = str(row['Parameter']).strip()
        unit_col = str(row['Unit']).strip()
        if param_name in param and (unit in unit_col or unit_col == unit):
            return idx
    return None

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

print("Collecting data and determining y-axis limits...\n")

# Collect data for temperature plots
temp_data = []
for y_param, y_unit, x_param, x_unit, file_name, category in plot_params:
    if category == 'temperature':
        points = collect_data_points(df_data, y_param, y_unit, x_param, x_unit)
        temp_data.extend([y for x, y, t, r in points])

# Collect data for pressure plots
pressure_data = []
for y_param, y_unit, x_param, x_unit, file_name, category in plot_params:
    if category == 'pressure':
        points = collect_data_points(df_data, y_param, y_unit, x_param, x_unit)
        pressure_data.extend([y for x, y, t, r in points])

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
pressure_ylim = calculate_ylim(pressure_data)

print(f"Temperature y-limits: {temp_ylim}")
print(f"Pressure y-limits: {pressure_ylim}\n")

print("Generating scatter plots...\n")

# Create a figure for each parameter
plot_count = 0
for y_param_name, y_unit, x_param_name, x_unit, file_name, category in plot_params:
    # Create figure with 12x8 inch size
    fig, ax = plt.subplots(figsize=(12, 8))
    
    # Find the row indices
    y_param_idx = find_parameter_row(df_data, y_param_name, y_unit)
    x_param_idx = find_parameter_row(df_data, x_param_name, x_unit)
    
    if y_param_idx is None:
        print(f"⚠ Warning: Could not find parameter '{y_param_name}' with unit '{y_unit}'")
        continue
    
    if x_param_idx is None:
        print(f"⚠ Warning: Could not find parameter '{x_param_name}' with unit '{x_unit}'")
        continue
    
    # Collect all data points for this parameter
    points_plotted = 0
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
                             linewidth=1.5)
                    points_plotted += 1
            except Exception as e:
                print(f"⚠ Error processing {run} - {test} for {y_param_name}: {e}")
                continue
    
    if points_plotted == 0:
        print(f"⚠ No data points found for {y_param_name} vs {x_param_name}")
        plt.close()
        continue
    
    # Customize plot
    ax.set_xlabel(f'{x_param_name} ({x_unit})', fontsize=14, fontweight='bold')
    ax.set_ylabel(f'{y_param_name} ({y_unit})', fontsize=14, fontweight='bold')
    ax.set_title(f'{y_param_name} vs {x_param_name}\nEVESE C4 - Thermal Runaway Analysis', 
                 fontsize=16, fontweight='bold', pad=20)
    ax.grid(True, alpha=0.3, linestyle='--', linewidth=0.8)
    
    # Set y-limits based on category
    if category == 'temperature' and temp_ylim[0] is not None:
        ax.set_ylim(temp_ylim)
    elif category == 'pressure' and pressure_ylim[0] is not None:
        ax.set_ylim(pressure_ylim)
    
    # Add minor gridlines
    ax.minorticks_on()
    ax.grid(which='minor', alpha=0.15, linestyle=':', linewidth=0.5)
    
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
    
    # Create single legend below the plot, centered, with two columns
    legend = ax.legend(handles=combined_legend_elements,
                      labels=combined_labels,
                      loc='upper center',
                      bbox_to_anchor=(0.5, -0.08),
                      ncol=7,  # All items in one row (4 runs + 3 tests)
                      frameon=True,
                      framealpha=0.95,
                      edgecolor='black',
                      fontsize=11,
                      columnspacing=1.5,
                      handletextpad=0.5)
    
    # # Add separator line in legend title to distinguish runs from tests
    # legend.set_title('Runs: Run1, Run2, Run3, Run4  |  Tests: Test1, Test2, Test3', 
    #                 prop={'size': 10, 'weight': 'normal'})
    
    # Adjust layout to make room for legend below
    plt.subplots_adjust(bottom=0.15)
    
    # Save figure with descriptive name
    output_filename = f'{file_name}.png'
    output_path = os.path.join(output_dir, output_filename)
    plt.savefig(output_path, dpi=300, bbox_inches='tight', facecolor='white')
    
    plot_count += 1
    print(f"✓ Saved ({plot_count}/{len(plot_params)}): {output_filename} ({points_plotted} data points)")
    
    # Close figure to free memory
    plt.close()

print(f"\n{'='*60}")
print(f"SUCCESS! All {plot_count} scatter plots generated successfully!")
print(f"Location: {output_dir}")
print(f"{'='*60}")