# -*- coding: utf-8 -*-
"""
Created on Fri Aug 15 17:31:44 2025

@author: ssalvi
"""
import pandas as pd
import matplotlib.pyplot as plt
import glob
import os
from tkinter import filedialog
from tkinter import Tk
from tkinter import messagebox
import tkinter as tk
from tkinter import ttk

def process_csv_data(file_path):
    """
    Processes a single CSV file by finding the start of the test based on
    a synchronization signal and adding relative time columns.
    """
    print(f"Reading data from: {file_path}")
    
    # try:
    #     data = pd.read_csv(file_path, low_memory=False)
    #     sync_column_name = 'Actuator Sync'
    #     required_columns = ['TIME 20 Hz', sync_column_name]

    #     if not all(col in data.columns for col in required_columns):
    #         print(f"!!ERROR!! - Missing required columns in '{os.path.basename(file_path)}'. Skipping.")
    #         print(f"Expected columns: {required_columns}")
    #         print(f"Found columns: {list(data.columns)}")
    #         return

    #     change_indices = data[(data[sync_column_name] > 8) & (data[sync_column_name].shift(1) <= 8)].index
        
    #     if len(change_indices) == 0:
    #         print(f"!!WARNING!! - Synchronization signal change not found in '{os.path.basename(file_path)}'. Skipping.")
    #         return
            
    #     change_index = change_indices[0]
    #     datum_time_sec = data.loc[change_index, 'TIME 20 Hz']
        
    #     data['Time (sec)'] = (data['TIME 20 Hz'] - datum_time_sec).round(3)
    #     data['Time (min)'] = (data['Time (sec)'] / 60).round(3)

    #     data.to_csv(file_path, index=False)
        
    #     print(f"\nSuccessfully processed '{os.path.basename(file_path)}'. Relative time columns added.")
        
    # except Exception as e:
    #     print(f"An error occurred while processing '{os.path.basename(file_path)}': {e}")

def get_temperature_columns(data):
    """
    Extracts all temperature-related columns from the dataframe.
    Looks for columns containing 'TC' or 'Temperature' or 'Temp'.
    """
    temp_columns = []
    for col in data.columns:
        # Check if column name contains temperature indicators
        if any(keyword in col for keyword in ['TC', 'Temperature', 'Temp', 'temperature', 'temp']):
            temp_columns.append(col)
    return temp_columns

def select_channels_dialog(available_channels, dialog_title, instruction_text):
    """
    Opens a dialog box for selecting channels from a list of available channels.
    Returns a list of selected channel names.
    """
    if not available_channels:
        messagebox.showwarning("No Channels", "No temperature channels found in the data!")
        return []
    
    selected_channels = []
    
    # Create dialog window
    dialog = tk.Toplevel()
    dialog.title(dialog_title)
    
    # Set window size based on number of channels
    window_height = min(500, 200 + len(available_channels) * 25)
    dialog.geometry(f"600x{window_height}")
    dialog.resizable(False, True)
    
    # Make it modal
    dialog.transient()
    dialog.grab_set()
    
    # Instruction label
    instruction = tk.Label(dialog, text=instruction_text, 
                          font=('Arial', 10, 'bold'), pady=10, wraplength=550)
    instruction.pack()
    
    # Create a frame with scrollbar for checkboxes
    canvas_frame = tk.Frame(dialog)
    canvas_frame.pack(fill='both', expand=True, padx=10, pady=5)
    
    canvas = tk.Canvas(canvas_frame)
    scrollbar = tk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)
    
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    # Dictionary to store checkbox variables
    checkbox_vars = {}
    
    # Create checkboxes for each channel
    for channel in available_channels:
        var = tk.BooleanVar(value=False)  # Default: none selected
        checkbox_vars[channel] = var
        
        # Create a more readable display name
        display_name = channel.replace('/RTAC Data/', '').strip()
        
        cb = tk.Checkbutton(scrollable_frame, text=display_name, variable=var, 
                           font=('Arial', 9), anchor='w', wraplength=500, justify='left')
        cb.pack(fill='x', padx=10, pady=2)
    
    # Frame for buttons
    button_frame = tk.Frame(dialog)
    button_frame.pack(pady=10)
    
    def select_all():
        for var in checkbox_vars.values():
            var.set(True)
    
    def deselect_all():
        for var in checkbox_vars.values():
            var.set(False)
    
    def confirm_selection():
        selected = [col for col, var in checkbox_vars.items() if var.get()]
        if len(selected) == 0:
            messagebox.showwarning("No Selection", "Please select at least one channel!")
            return
        selected_channels.extend(selected)
        dialog.destroy()
    
    def cancel_selection():
        dialog.destroy()
    
    # Buttons
    select_all_btn = tk.Button(button_frame, text="Select All", command=select_all, width=12)
    select_all_btn.grid(row=0, column=0, padx=5)
    
    deselect_all_btn = tk.Button(button_frame, text="Deselect All", command=deselect_all, width=12)
    deselect_all_btn.grid(row=0, column=1, padx=5)
    
    cancel_btn = tk.Button(button_frame, text="Cancel", command=cancel_selection, 
                           width=12, bg='#f44336', fg='white', font=('Arial', 9, 'bold'))
    cancel_btn.grid(row=1, column=0, pady=10, padx=5)
    
    confirm_btn = tk.Button(button_frame, text="Confirm", command=confirm_selection, 
                           width=12, bg='#4CAF50', fg='white', font=('Arial', 9, 'bold'))
    confirm_btn.grid(row=1, column=1, pady=10, padx=5)
    
    # Wait for dialog to close
    dialog.wait_window()
    
    return selected_channels

def create_qgen_analysis(csv_file, main_data):
    """
    Creates or updates the 'Q_gen Analysis' sheet in the CSV file.
    Copies time columns and calculates T_Cell_Avg and T_Gas_Avg from selected channels.
    Also calculates temperature derivatives dT/dt_Cell and dT/dt_Gas.
    Calculates heat generation rate Q_dot_cell and cumulative energy E_cell.
    Creates a summary report sheet.
    Saves everything back to the original file location as Excel format.
    """
    print(f"\nCreating Q_gen Analysis for: {csv_file}")
    
    # Cell properties
    m_cell = 0.9  # kg
    Cp_cell = 1100  # J/(kg·K)
    
    try:
        # Check if required columns exist
        required_time_cols = ['Time (sec)', 'Time (min)', 'Time (hour)']
        missing_cols = [col for col in required_time_cols if col not in main_data.columns]
        
        if missing_cols:
            print(f"!!ERROR!! - Missing time columns: {missing_cols}")
            return
        
        # Get all temperature columns
        temp_columns = get_temperature_columns(main_data)
        
        if not temp_columns:
            print(f"!!ERROR!! - No temperature columns found in the data!")
            return
        
        print(f"\nFound {len(temp_columns)} temperature channels in the data.")
        
        # Select channels for T_Cell_Avg
        print("\n--- Selecting channels for T_Cell_Avg calculation ---")
        selected_cell_tcs = select_channels_dialog(
            temp_columns,
            "Select Channels for T_Cell_Avg",
            "Select temperature channels to include in T_Cell_Avg calculation:"
        )
        
        if not selected_cell_tcs:
            print("!!WARNING!! - No channels selected for T_Cell_Avg. Skipping analysis.")
            return
        
        # Select channels for T_Gas_Avg
        print("\n--- Selecting channels for T_Gas_Avg calculation ---")
        selected_gas_tcs = select_channels_dialog(
            temp_columns,
            "Select Channels for T_Gas_Avg",
            "Select temperature channels to include in T_Gas_Avg calculation:"
        )
        
        if not selected_gas_tcs:
            print("!!WARNING!! - No channels selected for T_Gas_Avg. Skipping analysis.")
            return
        
        # Create Q_gen Analysis DataFrame
        qgen_df = pd.DataFrame()
        
        # Copy time columns
        qgen_df['Time (sec)'] = main_data['Time (sec)']
        qgen_df['Time (min)'] = main_data['Time (min)']
        qgen_df['Time (hour)'] = main_data['Time (hour)']
        
        # Calculate T_Cell_Avg from selected thermocouples
        tc_data = main_data[selected_cell_tcs]
        qgen_df['T_Cell_Avg'] = tc_data.mean(axis=1).round(3)
        
        # Apply bounds checking: 0 < T_Cell_Avg < 1000
        # Replace out-of-bounds values with previous valid value
        valid_mask = (qgen_df['T_Cell_Avg'] > 0) & (qgen_df['T_Cell_Avg'] < 1000)
        
        # Forward fill: repeat previous valid value for out-of-bounds entries
        qgen_df.loc[~valid_mask, 'T_Cell_Avg'] = None
        qgen_df['T_Cell_Avg'] = qgen_df['T_Cell_Avg'].fillna(method='ffill')
        
        # Count out-of-bounds values
        out_of_bounds_count = (~valid_mask).sum()
        if out_of_bounds_count > 0:
            print(f"⚠ Warning: {out_of_bounds_count} out-of-bounds T_Cell_Avg values detected (not in range 0-1000°C)")
            print(f"  These values were replaced with the previous valid value.")
        
        # Calculate T_Gas_Avg from gas thermocouples
        gas_tc_data = main_data[selected_gas_tcs]
        qgen_df['T_Gas_Avg'] = gas_tc_data.mean(axis=1).round(3)
        
        # Apply bounds checking: 0 < T_Gas_Avg < 1000
        # Replace out-of-bounds values with previous valid value
        valid_gas_mask = (qgen_df['T_Gas_Avg'] > 0) & (qgen_df['T_Gas_Avg'] < 1000)
        
        # Forward fill: repeat previous valid value for out-of-bounds entries
        qgen_df.loc[~valid_gas_mask, 'T_Gas_Avg'] = None
        qgen_df['T_Gas_Avg'] = qgen_df['T_Gas_Avg'].fillna(method='ffill')
        
        # Count out-of-bounds values
        out_of_bounds_gas_count = (~valid_gas_mask).sum()
        if out_of_bounds_gas_count > 0:
            print(f"⚠ Warning: {out_of_bounds_gas_count} out-of-bounds T_Gas_Avg values detected (not in range 0-1000°C)")
            print(f"  These values were replaced with the previous valid value.")
        
        # Calculate temperature derivatives (dT/dt in K/sec)
        # Using central difference method for interior points
        # dT/dt = (T[i+1] - T[i-1]) / (t[i+1] - t[i-1])
        
        # For T_Cell_Avg derivative
        qgen_df['dT/dt_Cell'] = 0.0
        time_sec = qgen_df['Time (sec)'].values
        temp_cell = qgen_df['T_Cell_Avg'].values
        temp_gas = qgen_df['T_Gas_Avg'].values
        
        # Calculate dT/dt_Cell
        dT_dt_cell = []
        for i in range(len(qgen_df)):
            if i == 0:
                # Forward difference for first point
                if len(qgen_df) > 1:
                    dt = time_sec[i+1] - time_sec[i]
                    if dt != 0:
                        derivative = (temp_cell[i+1] - temp_cell[i]) / dt
                    else:
                        derivative = 0.0
                else:
                    derivative = 0.0
            elif i == len(qgen_df) - 1:
                # Backward difference for last point
                dt = time_sec[i] - time_sec[i-1]
                if dt != 0:
                    derivative = (temp_cell[i] - temp_cell[i-1]) / dt
                else:
                    derivative = 0.0
            else:
                # Central difference for interior points
                dt = time_sec[i+1] - time_sec[i-1]
                if dt != 0:
                    derivative = (temp_cell[i+1] - temp_cell[i-1]) / dt
                else:
                    derivative = 0.0
            
            dT_dt_cell.append(round(derivative, 3))
        
        qgen_df['dT/dt_Cell'] = dT_dt_cell
        
        # Calculate dT/dt_Gas
        dT_dt_gas = []
        for i in range(len(qgen_df)):
            if i == 0:
                # Forward difference for first point
                if len(qgen_df) > 1:
                    dt = time_sec[i+1] - time_sec[i]
                    if dt != 0:
                        derivative = (temp_gas[i+1] - temp_gas[i]) / dt
                    else:
                        derivative = 0.0
                else:
                    derivative = 0.0
            elif i == len(qgen_df) - 1:
                # Backward difference for last point
                dt = time_sec[i] - time_sec[i-1]
                if dt != 0:
                    derivative = (temp_gas[i] - temp_gas[i-1]) / dt
                else:
                    derivative = 0.0
            else:
                # Central difference for interior points
                dt = time_sec[i+1] - time_sec[i-1]
                if dt != 0:
                    derivative = (temp_gas[i+1] - temp_gas[i-1]) / dt
                else:
                    derivative = 0.0
            
            dT_dt_gas.append(round(derivative, 3))
        
        qgen_df['dT/dt_Gas'] = dT_dt_gas
        
        print(f"\n✓ Temperature derivatives calculated:")
        print(f"  - dT/dt_Cell range: {qgen_df['dT/dt_Cell'].min():.3f} to {qgen_df['dT/dt_Cell'].max():.3f} K/sec")
        print(f"  - dT/dt_Gas range: {qgen_df['dT/dt_Gas'].min():.3f} to {qgen_df['dT/dt_Gas'].max():.3f} K/sec")
        
        # Calculate heat generation rate Q_dot_cell (W)
        # Q_dot_cell = m_cell × Cp_cell × dT/dt_Cell
        qgen_df['Q_dot_cell'] = (m_cell * Cp_cell * qgen_df['dT/dt_Cell']).round(3)
        
        print(f"\n✓ Heat generation rate calculated:")
        print(f"  - Cell mass (m_cell): {m_cell} kg")
        print(f"  - Cell specific heat (Cp_cell): {Cp_cell} J/(kg·K)")
        print(f"  - Q_dot_cell range: {qgen_df['Q_dot_cell'].min():.3f} to {qgen_df['Q_dot_cell'].max():.3f} W")
        
        # Calculate cumulative energy E_cell (J)
        # E_cell = ∫ Q_dot_cell dt = Σ (Q_dot_cell × Δt)
        E_cell = []
        cumulative_energy = 0.0
        
        for i in range(len(qgen_df)):
            if i == 0:
                # First point, no energy accumulated yet
                E_cell.append(0.0)
            else:
                # Calculate time step
                dt = time_sec[i] - time_sec[i-1]
                # Trapezoidal integration: average power over time step
                avg_power = (qgen_df['Q_dot_cell'].iloc[i] + qgen_df['Q_dot_cell'].iloc[i-1]) / 2.0
                energy_increment = avg_power * dt
                cumulative_energy += energy_increment
                E_cell.append(round(cumulative_energy, 3))
        
        qgen_df['E_cell'] = E_cell
        
        # Convert cumulative energy to kJ for reporting
        total_energy_kJ = cumulative_energy / 1000.0
        
        print(f"\n✓ Cumulative energy calculated:")
        print(f"  - Total energy released: {total_energy_kJ:.3f} kJ ({cumulative_energy:.3f} J)")
        
        # Peak heat rate detection
        peak_idx = qgen_df['Q_dot_cell'].idxmax()
        peak_power = qgen_df['Q_dot_cell'].iloc[peak_idx]
        peak_time_sec = qgen_df['Time (sec)'].iloc[peak_idx]
        peak_time_min = qgen_df['Time (min)'].iloc[peak_idx]
        peak_temp = qgen_df['T_Cell_Avg'].iloc[peak_idx]
        peak_dT_dt = qgen_df['dT/dt_Cell'].iloc[peak_idx]
        
        print(f"\n✓ Peak heat generation detected:")
        print(f"  - Peak power (Q_dot_cell_max): {peak_power:.3f} W ({peak_power/1000:.3f} kW)")
        print(f"  - Time of peak: {peak_time_sec:.3f} sec ({peak_time_min:.3f} min)")
        print(f"  - Cell temperature at peak: {peak_temp:.3f} °C")
        print(f"  - Heating rate at peak: {peak_dT_dt:.3f} K/sec")
        
        # Find maximum T_Cell_Avg and time
        max_cell_temp_idx = qgen_df['T_Cell_Avg'].idxmax()
        max_cell_temp = qgen_df['T_Cell_Avg'].iloc[max_cell_temp_idx]
        max_cell_temp_time = qgen_df['Time (sec)'].iloc[max_cell_temp_idx]
        
        # Calculate energy released from t=0 to time of maximum T_Cell_Avg
        energy_to_max_temp = qgen_df['E_cell'].iloc[max_cell_temp_idx]
        energy_to_max_temp_kJ = energy_to_max_temp / 1000.0
        energy_to_max_temp_Wh = energy_to_max_temp / 3600.0  # NEW: Convert J to Wh
        
        # Find maximum T_Gas_Avg and time
        max_gas_temp_idx = qgen_df['T_Gas_Avg'].idxmax()
        max_gas_temp = qgen_df['T_Gas_Avg'].iloc[max_gas_temp_idx]
        max_gas_temp_time = qgen_df['Time (sec)'].iloc[max_gas_temp_idx]
        
        # Find maximum pressure from main data
        pressure_column = '/RTAC Data/1000 Pressure'
        if pressure_column in main_data.columns:
            max_pressure_idx = main_data[pressure_column].idxmax()
            max_pressure = main_data[pressure_column].iloc[max_pressure_idx]
            max_pressure_time = main_data['Time (sec)'].iloc[max_pressure_idx]
            pressure_available = True
        else:
            print(f"⚠ Warning: Pressure column '{pressure_column}' not found in data.")
            max_pressure = 'N/A'
            max_pressure_time = 'N/A'
            pressure_available = False
        
        # Find values at 60 seconds
        time_60_idx = (qgen_df['Time (sec)'] - 60).abs().idxmin()
        actual_time_60 = qgen_df['Time (sec)'].iloc[time_60_idx]
        cell_temp_at_60 = qgen_df['T_Cell_Avg'].iloc[time_60_idx]
        gas_temp_at_60 = qgen_df['T_Gas_Avg'].iloc[time_60_idx]
        energy_at_60 = qgen_df['E_cell'].iloc[time_60_idx]
        energy_at_60_kJ = energy_at_60 / 1000.0
        
        if pressure_available:
            # Find closest time in main_data to 60 sec
            time_60_main_idx = (main_data['Time (sec)'] - 60).abs().idxmin()
            pressure_at_60 = main_data[pressure_column].iloc[time_60_main_idx]
        else:
            pressure_at_60 = 'N/A'
        
        print(f"\n✓ Additional metrics calculated:")
        print(f"  - Max Cell Avg Temperature: {max_cell_temp:.3f} °C at {max_cell_temp_time:.3f} sec")
        print(f"  - Energy released to max T_Cell_Avg: {energy_to_max_temp_kJ:.3f} kJ ({energy_to_max_temp:.3f} J)")
        print(f"  - Max Gas Avg Temperature: {max_gas_temp:.3f} °C at {max_gas_temp_time:.3f} sec")
        if pressure_available:
            print(f"  - Max Pressure: {max_pressure:.3f} PSIG at {max_pressure_time:.3f} sec")
        print(f"  - At t=60 sec (actual: {actual_time_60:.3f} sec):")
        print(f"    • Cell Avg Temperature: {cell_temp_at_60:.3f} °C")
        print(f"    • Gas Avg Temperature: {gas_temp_at_60:.3f} °C")
        print(f"    • Energy released: {energy_at_60_kJ:.3f} kJ ({energy_at_60:.3f} J)")
        if pressure_available:
            print(f"    • Pressure: {pressure_at_60:.3f} PSIG")
        
        # Create Report DataFrame
        report_data = {
            'Parameter': [
                'Cell Mass (m_cell)',
                'Cell Specific Heat (Cp_cell)',
                '',
                'Energy Released to Max T_Cell_Avg',
                'Energy Released to Max T_Cell_Avg',
                'Energy Released to Max T_Cell_Avg',
                '',
                'Maximum Cell Avg Temperature',
                'Time of Maximum Cell Avg Temperature',
                '',
                'Maximum Gas Avg Temperature',
                'Time of Maximum Gas Avg Temperature',
                '',
                'Maximum Pressure',
                'Time of Maximum Pressure',
                '',
                'Cell Avg Temperature at t=60 sec',
                'Actual Time',
                'Gas Avg Temperature at t=60 sec',
                'Pressure at t=60 sec',
                '',
                'T_Cell_Avg Channels Used',
            ],
            'Value': [
                m_cell,
                Cp_cell,
                '',
                round(energy_to_max_temp, 3),
                round(energy_to_max_temp_kJ, 3),
                round(energy_to_max_temp_Wh, 3),
                '',
                round(max_cell_temp, 3),
                round(max_cell_temp_time, 3),
                '',
                round(max_gas_temp, 3),
                round(max_gas_temp_time, 3),
                '',
                round(max_pressure, 3) if pressure_available else max_pressure,
                round(max_pressure_time, 3) if pressure_available else max_pressure_time,
                '',
                round(cell_temp_at_60, 3),
                round(actual_time_60, 3),
                round(gas_temp_at_60, 3),
                round(pressure_at_60, 3) if pressure_available else pressure_at_60,
                '',
                len(selected_cell_tcs),
            ],
            'Unit': [
                'kg',
                'J/(kg·K)',
                '',
                'J',
                'kJ',
                'Wh',
                '',
                '°C',
                'sec',
                '',
                '°C',
                'sec',
                '',
                'PSIG',
                'sec',
                '',
                '°C',
                'sec',
                '°C',
                'PSIG',
                '',
                'channels',
            ]
        }
        
        report_df = pd.DataFrame(report_data)
        
        # Add channel names to report
        for idx, tc in enumerate(selected_cell_tcs, start=1):
            tc_name = tc.replace('/RTAC Data/', '').strip()
            new_row = pd.DataFrame({
                'Parameter': [f'  Channel {idx}'],
                'Value': [tc_name],
                'Unit': ['']
            })
            report_df = pd.concat([report_df, new_row], ignore_index=True)
        
        # Add blank row
        report_df = pd.concat([report_df, pd.DataFrame({'Parameter': [''], 'Value': [''], 'Unit': ['']})], ignore_index=True)
        
        # Add T_Gas_Avg channels
        gas_row = pd.DataFrame({
            'Parameter': ['T_Gas_Avg Channels Used'],
            'Value': [len(selected_gas_tcs)],
            'Unit': ['channels']
        })
        report_df = pd.concat([report_df, gas_row], ignore_index=True)
        
        for idx, tc in enumerate(selected_gas_tcs, start=1):
            tc_name = tc.replace('/RTAC Data/', '').strip()
            new_row = pd.DataFrame({
                'Parameter': [f'  Channel {idx}'],
                'Value': [tc_name],
                'Unit': ['']
            })
            report_df = pd.concat([report_df, new_row], ignore_index=True)
        
        # Save to Excel file - REPLACE the original CSV with Excel format
        excel_file = csv_file.replace('.csv', '.xlsx')
        
        with pd.ExcelWriter(excel_file, engine='openpyxl', mode='w') as writer:
            main_data.to_excel(writer, sheet_name='Main Data', index=False)
            qgen_df.to_excel(writer, sheet_name='Q_gen Analysis', index=False)
            report_df.to_excel(writer, sheet_name='Report', index=False)
        
        print(f"\n✓ Q_gen Analysis sheet created successfully!")
        print(f"✓ Report sheet created successfully!")
        print(f"✓ File saved as: {excel_file}")
        print(f"✓ Original CSV preserved: {csv_file}")
        print(f"\n✓ T_Cell_Avg calculated using {len(selected_cell_tcs)} channels:")
        for tc in selected_cell_tcs:
            tc_name = tc.replace('/RTAC Data/', '').strip()
            print(f"  - {tc_name}")
        
        print(f"\n✓ T_Gas_Avg calculated using {len(selected_gas_tcs)} channels:")
        for tc in selected_gas_tcs:
            tc_name = tc.replace('/RTAC Data/', '').strip()
            print(f"  - {tc_name}")
        
        return qgen_df
        
    except Exception as e:
        print(f"!!ERROR!! - Failed to create Q_gen Analysis: {e}")
        import traceback
        traceback.print_exc()
        return None


# --- Main Script ---

if __name__ == '__main__':
    # Initialize tkinter and hide the root window
    root = Tk()
    root.withdraw()
    
    # Ask user if they want to generate plots
    generate_plots = messagebox.askyesno("Generate Plots", "Do you want to generate plots?")
    
    if generate_plots:
        print("\n✓ Plots will be generated and saved.")
    else:
        print("\n✗ Plot generation skipped. Only data processing will be performed.")
    
    # Ask user if they want to create Q_gen Analysis
    create_qgen = messagebox.askyesno("Q_gen Analysis", "Do you want to create Q_gen Analysis sheet?")
    
    # Prompt the user to select the directory
    directory = filedialog.askdirectory(title="Select the Directory with CSV Files")
    
    if not directory:
        print("!!ERROR!! - No directory selected. Exiting.")
    else:
        # Change the current working directory to the selected folder
        os.chdir(directory)

        # Now, you can simply use the file names without the full path
        csv_files = glob.glob("*.csv")

        if not csv_files:
            print("!!ERROR!! - No CSV files found in the specified directory.")
        else:            
            d = None  # Initialize d outside the loop
            for csv_file in csv_files:
                process_csv_data(csv_file)
                
                # Read the data for this file
                main_data = pd.read_csv(csv_file, low_memory=False)
                
                # Store first file's data for plotting
                if d is None:
                    d = main_data.copy()
                
                # Create Q_gen Analysis if requested
                if create_qgen:
                    create_qgen_analysis(csv_file, main_data)
            
            if d is not None:
                print("\n\nFirst processed file loaded into DataFrame 'd' for further analysis.")
                print(d.head())
                
                


#%% #######Producing Temperature (all) plots in Minutes#####################################################################

if generate_plots:
    # Creating plot with dataset
    fig, ax1 = plt.subplots()

    color1 = 'tab:red'
    color2 = 'tab:pink'
    color3 = 'tab:purple'
    color4 = 'tab:green'
    color5 = 'tab:blue' 

    ax1.set_xlabel('time (min)')
    ax1.set_ylabel('Temperature (degC)')
    ax1.plot(d['Time (min)'],d['/RTAC Data/TC1  Cell Positive'], label = 'Cell Positive', color = color2)
    ax1.plot(d['Time (min)'],d['/RTAC Data/TC2  Cell Negative'], label = 'Cell Negative', color = color3)     
    ax1.plot(d['Time (min)'],d['/RTAC Data/TC3  Cell Front Center'], label = 'Cell Front Center', color = color4)
    ax1.plot(d['Time (min)'],d['/RTAC Data/TC4  Cell Back Center'],'--', label = 'Cell Back Center', color = color4)
    ax1.plot(d['Time (min)'],d['/RTAC Data/TC5  Cell Vent'], label = 'Cell Vent', color = color1)
    ax1.plot(d['Time (min)'],d['/RTAC Data/TC6  Enclosure Ambient'], label = 'Enclosure Ambient', color = color5)
    ax1.plot(d['Time (min)'],d['/RTAC Data/TC7  Cell Side Positive'],'--', label = 'Cell Side Positive', color = color2)
    ax1.plot(d['Time (min)'],d['/RTAC Data/TC8  Enclosure Ambient Wire'],'--', label = 'Enclosure Ambient Wire', color = color3)   
    ax1.legend(bbox_to_anchor=(0., 1.12, 1., .102), loc='lower left',
               ncol=2, mode="expand", borderaxespad=0.)

    plt.xlim([-1, 65])

    ax1.xaxis.set_major_locator(plt.MultipleLocator(5))


    # Adding Twin Axes to plot using dataset_2

    ax2 = ax1.twinx()
     
    color6 = 'tab:gray'
    color7 = 'black'

    ax2.set_ylabel('Cell Voltage & Actuator Signal (V)')
    ax2.plot(d['Time (min)'],d['/RTAC Data/Actuator Sync'], ':', label = 'Nail Penetration', color = color6)
    ax2.plot(d['Time (min)'],d['/RTAC Data/Cell_V'], label = 'Cell Voltage', color = color7)
    ax2.tick_params(axis ='y')
    ax2.legend(bbox_to_anchor=(0., 1.02, 1., .102), loc='lower left',
                ncol=2, mode="expand", borderaxespad=0.)
    plt.ylim([0, 12])

    ax1.set_ylim(0, 815)
    plt.savefig("EVESE_C4 Temperatures - all.png", dpi=500, bbox_inches= 'tight')   
    # ax1.set_xlim(-0.1, 2)
    # ax1.xaxis.set_major_locator(plt.MultipleLocator(1))
    # ax1.set_ylim(0, 815)
    # plt.savefig("EVESE_C4 Temperatures - all_zoomed.png", dpi=500, bbox_inches= 'tight')    


#%% #######Producing Temperature (2 miutes) plots in seconds#####################################################################

if generate_plots:
    # Creating plot with dataset
    fig, ax1 = plt.subplots()
     
    color1 = 'tab:red'
    color2 = 'tab:pink'
    color3 = 'tab:purple'
    color4 = 'tab:green'
    color5 = 'tab:blue' 

    ax1.set_xlabel('time (sec)')
    ax1.set_ylabel('Temperature (degC)')
    ax1.plot(d['Time (sec)'],d['/RTAC Data/TC1  Cell Positive'], label = 'Cell Positive', color = color2)
    ax1.plot(d['Time (sec)'],d['/RTAC Data/TC2  Cell Negative'], label = 'Cell Negative', color = color3)     
    ax1.plot(d['Time (sec)'],d['/RTAC Data/TC3  Cell Front Center'], label = 'Cell Front Center', color = color4)
    ax1.plot(d['Time (sec)'],d['/RTAC Data/TC4  Cell Back Center'],'--', label = 'Cell Back Center', color = color4)
    ax1.plot(d['Time (sec)'],d['/RTAC Data/TC5  Cell Vent'], label = 'Cell Vent', color = color1)
    ax1.plot(d['Time (sec)'],d['/RTAC Data/TC6  Enclosure Ambient'], label = 'Enclosure Ambient', color = color5)
    ax1.plot(d['Time (sec)'],d['/RTAC Data/TC7  Cell Side Positive'],'--', label = 'Cell Side Positive', color = color2)
    ax1.plot(d['Time (sec)'],d['/RTAC Data/TC8  Enclosure Ambient Wire'],'--', label = 'Enclosure Ambient Wire', color = color3) 
    ax1.legend(bbox_to_anchor=(0., 1.12, 1., .102), loc='lower left',
               ncol=2, mode="expand", borderaxespad=0.)

    plt.ylim([0, 815])
    ax1.xaxis.set_major_locator(plt.MultipleLocator(30))

    # Adding Twin Axes to plot using dataset_2
    ax2 = ax1.twinx()
    color6 = 'tab:gray'
    color7 = 'black'
    ax2.set_ylabel('Cell Voltage & Actuator Signal (V)')
    ax2.plot(d['Time (sec)'],d['/RTAC Data/Actuator Sync'], ':', label = 'Nail Penetration', color = color6)
    ax2.plot(d['Time (sec)'],d['/RTAC Data/Cell_V'], label = 'Cell Voltage', color = color7)
    ax2.tick_params(axis ='y')
    ax2.legend(bbox_to_anchor=(0., 1.02, 1., .102), loc='lower left',
                ncol=2, mode="expand", borderaxespad=0.)
    plt.ylim([0, 12])


    plt.xlim([-6, 120])
    plt.savefig("EVESE_C4 Temperatures - 2 min.png", dpi=500, bbox_inches= 'tight')    

    plt.xlim([-6, 300])
    plt.savefig("EVESE_C4 Temperatures - 5 min.png", dpi=500, bbox_inches= 'tight')    

#%% Plotting alll Individual Parameters

if generate_plots:
    fig, ax1 = plt.subplots()
    ax1.set_xlabel('time (sec)')
    ax1.set_ylabel('Pressure (PSIG)')
    ax1.plot(d['Time (sec)'],d['/RTAC Data/1000 Pressure'], 'r', label = 'Enclosure Pressure')
    ax1.plot(d['Time (sec)'],d['/RTAC Data/Air Pressure'], 'b', label = 'Actuator Pressure')
    plt.ylim([0, 250])
    ax1.legend(bbox_to_anchor=(0., 1.12, 1., .102), loc='lower left',
                ncol=2, mode="expand", borderaxespad=0.)
    ax2 = ax1.twinx()
    color6 = 'tab:gray'
    color7 = 'black'
    ax2.set_ylabel('Cell Voltage & Actuator Signal (V)')
    ax2.plot(d['Time (sec)'],d['/RTAC Data/Actuator Sync'], ':', label = 'Nail Penetration', color = color6)
    ax2.plot(d['Time (sec)'],d['/RTAC Data/Cell_V'], label = 'Cell Voltage', color = color7)
    ax2.tick_params(axis ='y')
    ax2.legend(bbox_to_anchor=(0., 1.02, 1., .102), loc='lower left',
                ncol=2, mode="expand", borderaxespad=0.)
    plt.ylim([0, 12])
     
    plt.xlim([-6, 60])    
    plt.savefig('EVESE_C4 pressures - 1min.png', dpi=300, bbox_inches= 'tight')
    plt.xlim([-6, 120])    
    plt.savefig('EVESE_C4 pressures - 2min.png', dpi=300, bbox_inches= 'tight')
    plt.xlim([-6, 300])    
    plt.savefig('EVESE_C4 pressures - 5min.png', dpi=300, bbox_inches= 'tight')



    fig, ax1 = plt.subplots()
    ax1.set_xlabel('time (sec)')
    ax1.set_ylabel('Nail Travel (mm)')
    ax1.plot(d['Time (sec)'],d['/RTAC Data/Dispacement_LVIT'], 'g', label = 'Nail Displacement')
    plt.ylim([-5, 80])
    ax1.legend(bbox_to_anchor=(0., 1.12, 1., .102), loc='lower left',
                ncol=2, mode="expand", borderaxespad=0.)
    ax2 = ax1.twinx()
    color6 = 'tab:gray'
    color7 = 'black'
    ax2.set_ylabel('Cell Voltage & Actuator Signal (V)')
    ax2.plot(d['Time (sec)'],d['/RTAC Data/Actuator Sync'], ':', label = 'Nail Penetration', color = color6)
    ax2.plot(d['Time (sec)'],d['/RTAC Data/Cell_V'], label = 'Cell Voltage', color = color7)
    ax2.tick_params(axis ='y')
    ax2.legend(bbox_to_anchor=(0., 1.02, 1., .102), loc='lower left',
                ncol=2, mode="expand", borderaxespad=0.)
    plt.ylim([0, 12])
     
    plt.xlim([-2, 2])    
    plt.savefig('EVESE_C4 displacement - 2sec.png', dpi=300, bbox_inches= 'tight')
    plt.xlim([-2, 3])    
    plt.savefig('EVESE_C4 displacement - 3sec.png', dpi=300, bbox_inches= 'tight')
    plt.xlim([-2, 5])    
    plt.savefig('EVESE_C4 displacement - 5sec.png', dpi=300, bbox_inches= 'tight')

    print("\n✓ All plots generated successfully!")
else:
    print("\n✗ Plot generation was skipped as per user selection.")
