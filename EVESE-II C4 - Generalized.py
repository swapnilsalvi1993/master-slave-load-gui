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
    Includes weight input boxes for weighted average calculation.
    Returns a tuple: (list of selected channel names, dictionary of weights)
    """
    if not available_channels:
        messagebox.showwarning("No Channels", "No temperature channels found in the data!")
        return [], {}
    
    selected_channels = []
    channel_weights = {}
    
    # Create dialog window
    dialog = tk.Toplevel()
    dialog.title(dialog_title)
    
    # Set window size based on number of channels
    window_height = min(600, 250 + len(available_channels) * 30)
    dialog.geometry(f"700x{window_height}")
    dialog.resizable(False, True)
    
    # Make it modal
    dialog.transient()
    dialog.grab_set()
    
    # Instruction label
    instruction = tk.Label(dialog, text=instruction_text, 
                          font=('Arial', 10, 'bold'), pady=5, wraplength=650)
    instruction.pack()
    
    # Weight instruction label
    weight_instruction = tk.Label(dialog, 
                                  text="Enter weights for selected channels (must sum to 100%)", 
                                  font=('Arial', 9), pady=5, fg='blue')
    weight_instruction.pack()
    
    # Sum label
    sum_label_var = tk.StringVar(value="Total Weight: 0.0%")
    sum_label = tk.Label(dialog, textvariable=sum_label_var, 
                        font=('Arial', 9, 'bold'), pady=5, fg='red')
    sum_label.pack()
    
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
    
    # Dictionary to store checkbox variables and weight entries
    checkbox_vars = {}
    weight_entries = {}
    weight_vars = {}
    
    def update_weights():
        """Automatically distribute weights equally among selected channels"""
        selected = [col for col, var in checkbox_vars.items() if var.get()]
        
        if len(selected) > 0:
            equal_weight = 100.0 / len(selected)
            
            for channel in available_channels:
                if channel in selected:
                    weight_vars[channel].set(f"{equal_weight:.2f}")
                else:
                    weight_vars[channel].set("0.00")
        else:
            for channel in available_channels:
                weight_vars[channel].set("0.00")
        
        update_sum()
    
    def update_sum(*args):
        """Update the sum of weights and enable/disable confirm button"""
        total = 0.0
        try:
            for channel, var in weight_vars.items():
                if checkbox_vars[channel].get():  # Only sum selected channels
                    val = var.get().strip()
                    if val:
                        total += float(val)
        except ValueError:
            total = 0.0
        
        sum_label_var.set(f"Total Weight: {total:.2f}%")
        
        # Enable/disable confirm button based on sum
        if abs(total - 100.0) < 0.01:  # Allow small floating point error
            sum_label.config(fg='green')
            confirm_btn.config(state='normal')
        else:
            sum_label.config(fg='red')
            confirm_btn.config(state='disabled')
    
    def on_checkbox_change(channel):
        """Handle checkbox state change"""
        update_weights()
    
    # Create checkboxes and weight entries for each channel
    for channel in available_channels:
        # Create frame for each row
        row_frame = tk.Frame(scrollable_frame)
        row_frame.pack(fill='x', padx=5, pady=2)
        
        # Checkbox variable
        var = tk.BooleanVar(value=False)
        checkbox_vars[channel] = var
        
        # Weight variable
        weight_var = tk.StringVar(value="0.00")
        weight_vars[channel] = weight_var
        weight_var.trace_add('write', update_sum)
        
        # Create a more readable display name - remove common prefixes
        display_name = channel
        prefixes_to_remove = ['/RTAC Data/', 'RTAC Data/', '/']
        for prefix in prefixes_to_remove:
            if display_name.startswith(prefix):
                display_name = display_name[len(prefix):]
                break
        display_name = display_name.strip()
        
        # Weight entry (left side)
        weight_entry = tk.Entry(row_frame, textvariable=weight_var, width=8, justify='center')
        weight_entry.pack(side='left', padx=(0, 10))
        weight_entries[channel] = weight_entry
        
        # Checkbox with channel name
        cb = tk.Checkbutton(row_frame, text=display_name, variable=var, 
                           font=('Arial', 9), anchor='w', wraplength=500, justify='left',
                           command=lambda ch=channel: on_checkbox_change(ch))
        cb.pack(side='left', fill='x', expand=True)
    
    # Frame for buttons
    button_frame = tk.Frame(dialog)
    button_frame.pack(pady=10)
    
    def select_all():
        for var in checkbox_vars.values():
            var.set(True)
        update_weights()
    
    def deselect_all():
        for var in checkbox_vars.values():
            var.set(False)
        update_weights()
    
    def confirm_selection():
        selected = [col for col, var in checkbox_vars.items() if var.get()]
        if len(selected) == 0:
            messagebox.showwarning("No Selection", "Please select at least one channel!")
            return
        
        # Collect weights
        weights = {}
        total = 0.0
        try:
            for channel in selected:
                weight = float(weight_vars[channel].get())
                weights[channel] = weight
                total += weight
        except ValueError:
            messagebox.showerror("Invalid Weight", "Please enter valid numbers for weights!")
            return
        
        if abs(total - 100.0) > 0.01:
            messagebox.showerror("Weight Sum Error", f"Total weight must equal 100%!\nCurrent total: {total:.2f}%")
            return
        
        selected_channels.extend(selected)
        channel_weights.update(weights)
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
                           width=12, bg='#4CAF50', fg='white', font=('Arial', 9, 'bold'),
                           state='disabled')  # Initially disabled
    confirm_btn.grid(row=1, column=1, pady=10, padx=5)
    
    # Wait for dialog to close
    dialog.wait_window()
    
    return selected_channels, channel_weights

def create_qgen_analysis(csv_file, main_data, output_directory):
    """
    Creates or updates the 'Q_gen Analysis' sheet in the CSV file.
    Copies time columns and calculates T_Cell_Avg and T_Gas_Avg from selected channels.
    Also calculates temperature derivatives dT/dt_Cell and dT/dt_Gas.
    Calculates heat generation rate Q_dot_cell and cumulative energy E_cell.
    Calculates gas heat generation Q_dot_gas and cumulative energy E_gas.
    Creates a summary report sheet.
    Saves to the specified output directory as Excel format.
    """
    print(f"\nCreating Q_gen Analysis for: {csv_file}")
    
    # Cell properties
    m_cell = 0.9  # kg
    Cp_cell = 1100  # J/(kg·K)
    
    # Gas properties
    R = 8.314463  # J/(mol·K)
    V = 0.025  # m^3
    Cp_gas = 34  # J/(mol·K)
    
    try:
        # Check if Time (sec) column exists (REQUIRED)
        if 'Time (sec)' not in main_data.columns:
            print(f"!!ERROR!! - 'Time (sec)' column is missing. Cannot proceed with analysis.")
            return
        
        # Check for optional time columns
        has_time_min = 'Time (min)' in main_data.columns
        has_time_hour = 'Time (hour)' in main_data.columns
        
        if not has_time_min:
            print(f"⚠ Warning: 'Time (min)' column not found. It will be skipped.")
        if not has_time_hour:
            print(f"⚠ Warning: 'Time (hour)' column not found. It will be skipped.")
        
        # Get all temperature columns
        temp_columns = get_temperature_columns(main_data)
        
        if not temp_columns:
            print(f"!!ERROR!! - No temperature columns found in the data!")
            return
        
        print(f"\nFound {len(temp_columns)} temperature channels in the data.")
        
        # Select channels for T_Cell_Avg
        print("\n--- Selecting channels for T_Cell_Avg calculation ---")
        selected_cell_tcs, cell_weights = select_channels_dialog(
            temp_columns,
            "Select Channels for T_Cell_Avg (Weighted Average)",
            "Select temperature channels and assign weights for T_Cell_Avg calculation:"
        )
        
        if not selected_cell_tcs:
            print("!!WARNING!! - No channels selected for T_Cell_Avg. Skipping analysis.")
            return
        
        # Select channels for T_Gas_Avg
        print("\n--- Selecting channels for T_Gas_Avg calculation ---")
        selected_gas_tcs, gas_weights = select_channels_dialog(
            temp_columns,
            "Select Channels for T_Gas_Avg (Weighted Average)",
            "Select temperature channels and assign weights for T_Gas_Avg calculation:"
        )
        
        if not selected_gas_tcs:
            print("!!WARNING!! - No channels selected for T_Gas_Avg. Skipping analysis.")
            return
        
        # Check if pressure column exists - try multiple possible names
        pressure_column = None
        possible_pressure_names = [
            '/RTAC Data/1000 Pressure',
            '1000 Pressure',
            'Pressure',
            '/RTAC Data/Pressure'
        ]
        
        for col_name in possible_pressure_names:
            if col_name in main_data.columns:
                pressure_column = col_name
                print(f"✓ Found pressure column: '{pressure_column}'")
                break
        
        if pressure_column is None:
            # Try fuzzy search - look for any column containing "Pressure"
            for col in main_data.columns:
                if 'Pressure' in col or 'pressure' in col:
                    pressure_column = col
                    print(f"✓ Found pressure column: '{pressure_column}'")
                    break
        
        if pressure_column is None:
            print(f"!!ERROR!! - No pressure column found. Cannot calculate gas properties.")
            print(f"  Looked for: {possible_pressure_names}")
            print(f"  Available columns: {list(main_data.columns)}")
            return
        
        # Create Q_gen Analysis DataFrame
        qgen_df = pd.DataFrame()
        
        # Copy time columns (only those that exist)
        qgen_df['Time (sec)'] = main_data['Time (sec)']
        if has_time_min:
            qgen_df['Time (min)'] = main_data['Time (min)']
        if has_time_hour:
            qgen_df['Time (hour)'] = main_data['Time (hour)']
        
        # Calculate T_Cell_Avg from selected thermocouples using WEIGHTED AVERAGE
        tc_data = main_data[selected_cell_tcs]
        
        # Apply weights (convert percentages to fractions)
        weighted_sum = pd.Series(0.0, index=tc_data.index)
        for col in selected_cell_tcs:
            weighted_sum += tc_data[col] * (cell_weights[col] / 100.0)
        
        qgen_df['T_Cell_Avg'] = weighted_sum.round(3)
        
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
        
        # Calculate T_Gas_Avg from gas thermocouples using WEIGHTED AVERAGE
        gas_tc_data = main_data[selected_gas_tcs]
        
        # Apply weights (convert percentages to fractions)
        weighted_gas_sum = pd.Series(0.0, index=gas_tc_data.index)
        for col in selected_gas_tcs:
            weighted_gas_sum += gas_tc_data[col] * (gas_weights[col] / 100.0)
        
        qgen_df['T_Gas_Avg'] = weighted_gas_sum.round(3)
        
        # Apply bounds checking: 0 < T_Gas_Avg < 1000
        # Replace out-of-bounds values with previous valid value
        valid_gas_mask = (qgen_df['T_Gas_Avg'] > 0) &   (qgen_df['T_Gas_Avg'] < 1000)
        
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
        
        # Calculate gas heat generation rate Q_dot_gas (W)
        # First, get pressure from main data and convert from PSIG to Pascal
        pressure_psig = main_data[pressure_column].values
        # Convert PSIG to Pascal: 1 psi = 6894.76 Pa, add atmospheric pressure (14.7 psi)
        pressure_pascal = (pressure_psig + 14.7) * 6894.76
        
        # Convert T_Gas_Avg from Celsius to Kelvin
        temp_gas_kelvin = temp_gas + 273.15
        
        # Calculate number of moles: n = (P*V)/(R*T)
        n_moles = []
        for i in range(len(qgen_df)):
            if temp_gas_kelvin[i] > 0:
                n = (pressure_pascal[i] * V) / (R * temp_gas_kelvin[i])
            else:
                n = 0.0
            n_moles.append(n)
        
        n_moles = pd.Series(n_moles)
        
        # Calculate Q_dot_gas = n × Cp_gas × dT/dt_Gas
        qgen_df['Q_dot_gas'] = (n_moles * Cp_gas * qgen_df['dT/dt_Gas']).round(3)
        
        print(f"\n✓ Gas heat generation rate calculated:")
        print(f"  - Enclosure volume (V): {V} m³")
        print(f"  - Gas constant (R): {R} J/(mol·K)")
        print(f"  - Gas specific heat (Cp_gas): {Cp_gas} J/(mol·K)")
        print(f"  - Number of moles range: {n_moles.min():.3f} to {n_moles.max():.3f} mol")
        print(f"  - Q_dot_gas range: {qgen_df['Q_dot_gas'].min():.3f} to {qgen_df['Q_dot_gas'].max():.3f} W")
        
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
        
        print(f"\n✓ Cumulative cell energy calculated:")
        print(f"  - Total cell energy released: {total_energy_kJ:.3f} kJ ({cumulative_energy:.3f} J)")
        
        # Calculate cumulative energy E_gas (J)
        E_gas = []
        cumulative_energy_gas = 0.0
        
        for i in range(len(qgen_df)):
            if i == 0:
                # First point, no energy accumulated yet
                E_gas.append(0.0)
            else:
                # Calculate time step
                dt = time_sec[i] - time_sec[i-1]
                # Trapezoidal integration: average power over time step
                avg_power = (qgen_df['Q_dot_gas'].iloc[i] + qgen_df['Q_dot_gas'].iloc[i-1]) / 2.0
                energy_increment = avg_power * dt
                cumulative_energy_gas += energy_increment
                E_gas.append(round(cumulative_energy_gas, 3))
        
        qgen_df['E_gas'] = E_gas
        
        # Convert cumulative gas energy to kJ for reporting
        total_energy_gas_kJ = cumulative_energy_gas / 1000.0
        
        print(f"\n✓ Cumulative gas energy calculated:")
        print(f"  - Total gas energy released: {total_energy_gas_kJ:.3f} kJ ({cumulative_energy_gas:.3f} J)")
        
        # Calculate E_total and E_ratio
        qgen_df['E_total'] = (qgen_df['E_cell'] + qgen_df['E_gas']).round(3)
        
        # Calculate E_ratio = E_gas / E_cell (avoid division by zero)
        E_ratio = []
        for i in range(len(qgen_df)):
            if qgen_df['E_cell'].iloc[i] != 0:
                ratio = qgen_df['E_gas'].iloc[i] / qgen_df['E_cell'].iloc[i]
            else:
                ratio = 0.0
            E_ratio.append(round(ratio, 3))
        
        qgen_df['E_ratio'] = E_ratio
        
        total_energy_total = qgen_df['E_total'].iloc[-1]
        total_energy_total_kJ = total_energy_total / 1000.0
        
        print(f"\n✓ Total energy calculated:")
        print(f"  - Total energy (E_cell + E_gas): {total_energy_total_kJ:.3f} kJ ({total_energy_total:.3f} J)")
        
        # Peak heat rate detection
        peak_idx = qgen_df['Q_dot_cell'].idxmax()
        peak_power = qgen_df['Q_dot_cell'].iloc[peak_idx]
        peak_time_sec = qgen_df['Time (sec)'].iloc[peak_idx]
        peak_temp = qgen_df['T_Cell_Avg'].iloc[peak_idx]
        peak_dT_dt = qgen_df['dT/dt_Cell'].iloc[peak_idx]
        
        # Only calculate time in minutes if column exists
        if has_time_min:
            peak_time_min = qgen_df['Time (min)'].iloc[peak_idx]
            print(f"\n✓ Peak heat generation detected:")
            print(f"  - Peak power (Q_dot_cell_max): {peak_power:.3f} W ({peak_power/1000:.3f} kW)")
            print(f"  - Time of peak: {peak_time_sec:.3f} sec ({peak_time_min:.3f} min)")
            print(f"  - Cell temperature at peak: {peak_temp:.3f} °C")
            print(f"  - Heating rate at peak: {peak_dT_dt:.3f} K/sec")
        else:
            print(f"\n✓ Peak heat generation detected:")
            print(f"  - Peak power (Q_dot_cell_max): {peak_power:.3f} W ({peak_power/1000:.3f} kW)")
            print(f"  - Time of peak: {peak_time_sec:.3f} sec")
            print(f"  - Cell temperature at peak: {peak_temp:.3f} °C")
            print(f"  - Heating rate at peak: {peak_dT_dt:.3f} K/sec")
        
        # Find maximum T_Cell_Avg and time
        max_cell_temp_idx = qgen_df['T_Cell_Avg'].idxmax()
        max_cell_temp = qgen_df['T_Cell_Avg'].iloc[max_cell_temp_idx]
        max_cell_temp_time = qgen_df['Time (sec)'].iloc[max_cell_temp_idx]
        
        # Calculate energy released from t=0 to time of maximum T_Cell_Avg
        energy_to_max_temp = qgen_df['E_cell'].iloc[max_cell_temp_idx]
        energy_to_max_temp_kJ = energy_to_max_temp / 1000.0
        energy_to_max_temp_Wh = energy_to_max_temp / 3600.0
        
        # Calculate gas energy released from t=0 to time of maximum T_Cell_Avg
        energy_gas_to_max_temp = qgen_df['E_gas'].iloc[max_cell_temp_idx]
        energy_gas_to_max_temp_kJ = energy_gas_to_max_temp / 1000.0
        energy_gas_to_max_temp_Wh = energy_gas_to_max_temp / 3600.0
        
        # Calculate total energy released to max T_Cell_Avg
        energy_total_to_max_temp = qgen_df['E_total'].iloc[max_cell_temp_idx]
        energy_total_to_max_temp_kJ = energy_total_to_max_temp / 1000.0
        energy_total_to_max_temp_Wh = energy_total_to_max_temp / 3600.0
        
        # Calculate energy ratio at max T_Cell_Avg
        energy_ratio_to_max_temp = qgen_df['E_ratio'].iloc[max_cell_temp_idx]
        
        # Find maximum T_Gas_Avg and time
        max_gas_temp_idx = qgen_df['T_Gas_Avg'].idxmax()
        max_gas_temp = qgen_df['T_Gas_Avg'].iloc[max_gas_temp_idx]
        max_gas_temp_time = qgen_df['Time (sec)'].iloc[max_gas_temp_idx]
        
        # Find maximum pressure from main data
        max_pressure_idx = main_data[pressure_column].idxmax()
        max_pressure = main_data[pressure_column].iloc[max_pressure_idx]
        max_pressure_time = main_data['Time (sec)'].iloc[max_pressure_idx]
        
        # Find values at 60 seconds
        time_60_idx = (qgen_df['Time (sec)'] - 60).abs().idxmin()
        actual_time_60 = qgen_df['Time (sec)'].iloc[time_60_idx]
        cell_temp_at_60 = qgen_df['T_Cell_Avg'].iloc[time_60_idx]
        gas_temp_at_60 = qgen_df['T_Gas_Avg'].iloc[time_60_idx]
        energy_at_60 = qgen_df['E_cell'].iloc[time_60_idx]
        energy_at_60_kJ = energy_at_60 / 1000.0
        
        # Find closest time in main_data to 60 sec
        time_60_main_idx = (main_data['Time (sec)'] - 60).abs().idxmin()
        pressure_at_60 = main_data[pressure_column].iloc[time_60_main_idx]
        
        # Find values at 5, 10, and 15 seconds
        # At 5 seconds
        time_5_idx = (qgen_df['Time (sec)'] - 5).abs().idxmin()
        actual_time_5 = qgen_df['Time (sec)'].iloc[time_5_idx]
        cell_temp_at_5 = qgen_df['T_Cell_Avg'].iloc[time_5_idx]
        gas_temp_at_5 = qgen_df['T_Gas_Avg'].iloc[time_5_idx]
        energy_cell_at_5 = qgen_df['E_cell'].iloc[time_5_idx]
        energy_cell_at_5_kJ = energy_cell_at_5 / 1000.0
        energy_cell_at_5_Wh = energy_cell_at_5 / 3600.0
        energy_gas_at_5 = qgen_df['E_gas'].iloc[time_5_idx]
        energy_gas_at_5_kJ = energy_gas_at_5 / 1000.0
        energy_gas_at_5_Wh = energy_gas_at_5 / 3600.0
        energy_total_at_5 = qgen_df['E_total'].iloc[time_5_idx]
        energy_total_at_5_kJ = energy_total_at_5 / 1000.0
        energy_total_at_5_Wh = energy_total_at_5 / 3600.0
        energy_ratio_at_5 = qgen_df['E_ratio'].iloc[time_5_idx]
        time_5_main_idx = (main_data['Time (sec)'] - 5).abs().idxmin()
        pressure_at_5 = main_data[pressure_column].iloc[time_5_main_idx]
        
        # At 10 seconds
        time_10_idx = (qgen_df['Time (sec)'] - 10).abs().idxmin()
        actual_time_10 = qgen_df['Time (sec)'].iloc[time_10_idx]
        cell_temp_at_10 = qgen_df['T_Cell_Avg'].iloc[time_10_idx]
        gas_temp_at_10 = qgen_df['T_Gas_Avg'].iloc[time_10_idx]
        energy_cell_at_10 = qgen_df['E_cell'].iloc[time_10_idx]
        energy_cell_at_10_kJ = energy_cell_at_10 / 1000.0
        energy_cell_at_10_Wh = energy_cell_at_10 / 3600.0
        energy_gas_at_10 = qgen_df['E_gas'].iloc[time_10_idx]
        energy_gas_at_10_kJ = energy_gas_at_10 / 1000.0
        energy_gas_at_10_Wh = energy_gas_at_10 / 3600.0
        energy_total_at_10 = qgen_df['E_total'].iloc[time_10_idx]
        energy_total_at_10_kJ = energy_total_at_10 / 1000.0
        energy_total_at_10_Wh = energy_total_at_10 / 3600.0
        energy_ratio_at_10 = qgen_df['E_ratio'].iloc[time_10_idx]
        time_10_main_idx = (main_data['Time (sec)'] - 10).abs().idxmin()
        pressure_at_10 = main_data[pressure_column].iloc[time_10_main_idx]
        
        # At 15 seconds
        time_15_idx = (qgen_df['Time (sec)'] - 15).abs().idxmin()
        actual_time_15 = qgen_df['Time (sec)'].iloc[time_15_idx]
        cell_temp_at_15 = qgen_df['T_Cell_Avg'].iloc[time_15_idx]
        gas_temp_at_15 = qgen_df['T_Gas_Avg'].iloc[time_15_idx]
        energy_cell_at_15 = qgen_df['E_cell'].iloc[time_15_idx]
        energy_cell_at_15_kJ = energy_cell_at_15 / 1000.0
        energy_cell_at_15_Wh = energy_cell_at_15 / 3600.0
        energy_gas_at_15 = qgen_df['E_gas'].iloc[time_15_idx]
        energy_gas_at_15_kJ = energy_gas_at_15 / 1000.0
        energy_gas_at_15_Wh = energy_gas_at_15 / 3600.0
        energy_total_at_15 = qgen_df['E_total'].iloc[time_15_idx]
        energy_total_at_15_kJ = energy_total_at_15 / 1000.0
        energy_total_at_15_Wh = energy_total_at_15 / 3600.0
        energy_ratio_at_15 = qgen_df['E_ratio'].iloc[time_15_idx]
        time_15_main_idx = (main_data['Time (sec)'] - 15).abs().idxmin()
        pressure_at_15 = main_data[pressure_column].iloc[time_15_main_idx]
        
        print(f"\n✓ Additional metrics calculated:")
        print(f"  - Max Cell Avg Temperature: {max_cell_temp:.3f} °C at {max_cell_temp_time:.3f} sec")
        print(f"  - Energy released by cell to max T_Cell_Avg: {energy_to_max_temp_kJ:.3f} kJ ({energy_to_max_temp:.3f} J, {energy_to_max_temp_Wh:.3f} Wh)")
        print(f"  - Energy released by gas to max T_Cell_Avg: {energy_gas_to_max_temp_kJ:.3f} kJ ({energy_gas_to_max_temp:.3f} J, {energy_gas_to_max_temp_Wh:.3f} Wh)")
        print(f"  - Total energy released to max T_Cell_Avg: {energy_total_to_max_temp_kJ:.3f} kJ ({energy_total_to_max_temp:.3f} J, {energy_total_to_max_temp_Wh:.3f} Wh)")
        print(f"  - Energy ratio to max T_Cell_Avg: {energy_ratio_to_max_temp:.3f}")
        print(f"  - Max Gas Avg Temperature: {max_gas_temp:.3f} °C at {max_gas_temp_time:.3f} sec")
        print(f"  - Max Pressure: {max_pressure:.3f} PSIG at {max_pressure_time:.3f} sec")
        print(f"  - At t=5 sec (actual: {actual_time_5:.3f} sec):")
        print(f"    • Cell Avg Temperature: {cell_temp_at_5:.3f} °C")
        print(f"    • Gas Avg Temperature: {gas_temp_at_5:.3f} °C")
        print(f"    • Energy released by cell: {energy_cell_at_5_kJ:.3f} kJ")
        print(f"    • Energy released by gas: {energy_gas_at_5_kJ:.3f} kJ")
        print(f"    • Total energy: {energy_total_at_5_kJ:.3f} kJ")
        print(f"    • Energy ratio: {energy_ratio_at_5:.3f}")
        print(f"    • Pressure: {pressure_at_5:.3f} PSIG")
        print(f"  - At t=10 sec (actual: {actual_time_10:.3f} sec):")
        print(f"    • Cell Avg Temperature: {cell_temp_at_10:.3f} °C")
        print(f"    • Gas Avg Temperature: {gas_temp_at_10:.3f} °C")
        print(f"    • Energy released by cell: {energy_cell_at_10_kJ:.3f} kJ")
        print(f"    • Energy released by gas: {energy_gas_at_10_kJ:.3f} kJ")
        print(f"    • Total energy: {energy_total_at_10_kJ:.3f} kJ")
        print(f"    • Energy ratio: {energy_ratio_at_10:.3f}")
        print(f"    • Pressure: {pressure_at_10:.3f} PSIG")
        print(f"  - At t=15 sec (actual: {actual_time_15:.3f} sec):")
        print(f"    • Cell Avg Temperature: {cell_temp_at_15:.3f} °C")
        print(f"    • Gas Avg Temperature: {gas_temp_at_15:.3f} °C")
        print(f"    • Energy released by cell: {energy_cell_at_15_kJ:.3f} kJ")
        print(f"    • Energy released by gas: {energy_gas_at_15_kJ:.3f} kJ")
        print(f"    • Total energy: {energy_total_at_15_kJ:.3f} kJ")
        print(f"    • Energy ratio: {energy_ratio_at_15:.3f}")
        print(f"    • Pressure: {pressure_at_15:.3f} PSIG")
        print(f"  - At t=60 sec (actual: {actual_time_60:.3f} sec):")
        print(f"    • Cell Avg Temperature: {cell_temp_at_60:.3f} °C")
        print(f"    • Gas Avg Temperature: {gas_temp_at_60:.3f} °C")
        print(f"    • Energy released: {energy_at_60_kJ:.3f} kJ ({energy_at_60:.3f} J)")
        print(f"    • Pressure: {pressure_at_60:.3f} PSIG")
        
        # Create Report DataFrame
        report_data = {
            'Parameter': [
                'Cell Mass (m_cell)',
                'Cell Specific Heat (Cp_cell)',
                '',
                'Enclosure Volume (V)',
                'Gas Constant (R)',
                'Gas Specific Heat (Cp_gas)',
                '',
                'Energy Released by Cell to Max T_Cell_Avg',
                'Energy Released by Cell to Max T_Cell_Avg',
                'Energy Released by Cell to Max T_Cell_Avg',
                '',
                'Energy Released by Gas to Max T_Cell_Avg',
                'Energy Released by Gas to Max T_Cell_Avg',
                'Energy Released by Gas to Max T_Cell_Avg',
                '',
                'Total Energy Released to Max T_Cell_Avg',
                'Total Energy Released to Max T_Cell_Avg',
                'Total Energy Released to Max T_Cell_Avg',
                '',
                'Energy Ratio to Max T_Cell_Avg',
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
                '--- Values at t = 5 sec ---',
                'Actual Time',
                'Energy Released by Cell at t=5 sec',
                'Energy Released by Cell at t=5 sec',
                'Energy Released by Cell at t=5 sec',
                'Energy Released by Gas at t=5 sec',
                'Energy Released by Gas at t=5 sec',
                'Energy Released by Gas at t=5 sec',
                'Total Energy Released at t=5 sec',
                'Total Energy Released at t=5 sec',
                'Total Energy Released at t=5 sec',
                'Energy Ratio at t=5 sec',
                'Cell Avg Temperature at t=5 sec',
                'Gas Avg Temperature at t=5 sec',
                'Pressure at t=5 sec',
                '',
                '--- Values at t = 10 sec ---',
                'Actual Time',
                'Energy Released by Cell at t=10 sec',
                'Energy Released by Cell at t=10 sec',
                'Energy Released by Cell at t=10 sec',
                'Energy Released by Gas at t=10 sec',
                'Energy Released by Gas at t=10 sec',
                'Energy Released by Gas at t=10 sec',
                'Total Energy Released at t=10 sec',
                'Total Energy Released at t=10 sec',
                'Total Energy Released at t=10 sec',
                'Energy Ratio at t=10 sec',
                'Cell Avg Temperature at t=10 sec',
                'Gas Avg Temperature at t=10 sec',
                'Pressure at t=10 sec',
                '',
                '--- Values at t = 15 sec ---',
                'Actual Time',
                'Energy Released by Cell at t=15 sec',
                'Energy Released by Cell at t=15 sec',
                'Energy Released by Cell at t=15 sec',
                'Energy Released by Gas at t=15 sec',
                'Energy Released by Gas at t=15 sec',
                'Energy Released by Gas at t=15 sec',
                'Total Energy Released at t=15 sec',
                'Total Energy Released at t=15 sec',
                'Total Energy Released at t=15 sec',
                'Energy Ratio at t=15 sec',
                'Cell Avg Temperature at t=15 sec',
                'Gas Avg Temperature at t=15 sec',
                'Pressure at t=15 sec',
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
                V,
                R,
                Cp_gas,
                '',
                round(energy_to_max_temp, 3),
                round(energy_to_max_temp_kJ, 3),
                round(energy_to_max_temp_Wh, 3),
                '',
                round(energy_gas_to_max_temp, 3),
                round(energy_gas_to_max_temp_kJ, 3),
                round(energy_gas_to_max_temp_Wh, 3),
                '',
                round(energy_total_to_max_temp, 3),
                round(energy_total_to_max_temp_kJ, 3),
                round(energy_total_to_max_temp_Wh, 3),
                '',
                round(energy_ratio_to_max_temp, 3),
                '',
                round(max_cell_temp, 3),
                round(max_cell_temp_time, 3),
                '',
                round(max_gas_temp, 3),
                round(max_gas_temp_time, 3),
                '',
                round(max_pressure, 3),
                round(max_pressure_time, 3),
                '',
                '',
                round(actual_time_5, 3),
                round(energy_cell_at_5, 3),
                round(energy_cell_at_5_kJ, 3),
                round(energy_cell_at_5_Wh, 3),
                round(energy_gas_at_5, 3),
                round(energy_gas_at_5_kJ, 3),
                round(energy_gas_at_5_Wh, 3),
                round(energy_total_at_5, 3),
                round(energy_total_at_5_kJ, 3),
                round(energy_total_at_5_Wh, 3),
                round(energy_ratio_at_5, 3),
                round(cell_temp_at_5, 3),
                round(gas_temp_at_5, 3),
                round(pressure_at_5, 3),
                '',
                '',
                round(actual_time_10, 3),
                round(energy_cell_at_10, 3),
                round(energy_cell_at_10_kJ, 3),
                round(energy_cell_at_10_Wh, 3),
                round(energy_gas_at_10, 3),
                round(energy_gas_at_10_kJ, 3),
                round(energy_gas_at_10_Wh, 3),
                round(energy_total_at_10, 3),
                round(energy_total_at_10_kJ, 3),
                round(energy_total_at_10_Wh, 3),
                round(energy_ratio_at_10, 3),
                round(cell_temp_at_10, 3),
                round(gas_temp_at_10, 3),
                round(pressure_at_10, 3),
                '',
                '',
                round(actual_time_15, 3),
                round(energy_cell_at_15, 3),
                round(energy_cell_at_15_kJ, 3),
                round(energy_cell_at_15_Wh, 3),
                round(energy_gas_at_15, 3),
                round(energy_gas_at_15_kJ, 3),
                round(energy_gas_at_15_Wh, 3),
                round(energy_total_at_15, 3),
                round(energy_total_at_15_kJ, 3),
                round(energy_total_at_15_Wh, 3),
                round(energy_ratio_at_15, 3),
                round(cell_temp_at_15, 3),
                round(gas_temp_at_15, 3),
                round(pressure_at_15, 3),
                '',
                round(cell_temp_at_60, 3),
                round(actual_time_60, 3),
                round(gas_temp_at_60, 3),
                round(pressure_at_60, 3),
                '',
                len(selected_cell_tcs),
            ],
            'Unit': [
                'kg',
                'J/(kg·K)',
                '',
                'm³',
                'J/(mol·K)',
                'J/(mol·K)',
                '',
                'J',
                'kJ',
                'Wh',
                '',
                'J',
                'kJ',
                'Wh',
                '',
                'J',
                'kJ',
                'Wh',
                '',
                'E_gas/E_cell',
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
                '',
                'sec',
                'J',
                'kJ',
                'Wh',
                'J',
                'kJ',
                'Wh',
                'J',
                'kJ',
                'Wh',
                'E_gas/E_cell',
                '°C',
                '°C',
                'PSIG',
                '',
                '',
                'sec',
                'J',
                'kJ',
                'Wh',
                'J',
                'kJ',
                'Wh',
                'J',
                'kJ',
                'Wh',
                'E_gas/E_cell',
                '°C',
                '°C',
                'PSIG',
                '',
                '',
                'sec',
                'J',
                'kJ',
                'Wh',
                'J',
                'kJ',
                'Wh',
                'J',
                'kJ',
                'Wh',
                'E_gas/E_cell',
                '°C',
                '°C',
                'PSIG',
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
        
        # Add channel names and weights to report
        for idx, tc in enumerate(selected_cell_tcs, start=1):
            # Clean up display name
            tc_name = tc
            prefixes_to_remove = ['/RTAC Data/', 'RTAC Data/', '/']
            for prefix in prefixes_to_remove:
                if tc_name.startswith(prefix):
                    tc_name = tc_name[len(prefix):]
                    break
            tc_name = tc_name.strip()
            
            new_row = pd.DataFrame({
                'Parameter': [f'  Channel {idx}'],
                'Value': [f"{tc_name} (Weight: {cell_weights[tc]:.2f}%)"],
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
            # Clean up display name
            tc_name = tc
            prefixes_to_remove = ['/RTAC Data/', 'RTAC Data/', '/']
            for prefix in prefixes_to_remove:
                if tc_name.startswith(prefix):
                    tc_name = tc_name[len(prefix):]
                    break
            tc_name = tc_name.strip()
            
            new_row = pd.DataFrame({
                'Parameter': [f'  Channel {idx}'],
                'Value': [f"{tc_name} (Weight: {gas_weights[tc]:.2f}%)"],
                'Unit': ['']
            })
            report_df = pd.concat([report_df, new_row], ignore_index=True)
        
        # Save to Excel file in the specified output directory
        base_filename = os.path.basename(csv_file)
        excel_filename = base_filename.replace('.csv', '.xlsx')
        excel_file = os.path.join(output_directory, excel_filename)
        
        with pd.ExcelWriter(excel_file, engine='openpyxl', mode='w') as writer:
            main_data.to_excel(writer, sheet_name='Main Data', index=False)
            qgen_df.to_excel(writer, sheet_name='Q_gen Analysis', index=False)
            report_df.to_excel(writer, sheet_name='Report', index=False)
        
        print(f"\n✓ Q_gen Analysis sheet created successfully!")
        print(f"✓ Report sheet created successfully!")
        print(f"✓ File saved as: {excel_file}")
        print(f"✓ Original CSV location: {csv_file}")
        print(f"\n✓ T_Cell_Avg calculated using {len(selected_cell_tcs)} channels (weighted average):")
        for tc in selected_cell_tcs:
            # Clean up display name
            tc_name = tc
            prefixes_to_remove = ['/RTAC Data/', 'RTAC Data/', '/']
            for prefix in prefixes_to_remove:
                if tc_name.startswith(prefix):
                    tc_name = tc_name[len(prefix):]
                    break
            tc_name = tc_name.strip()
            print(f"  - {tc_name}: {cell_weights[tc]:.2f}%")
        
        print(f"\n✓ T_Gas_Avg calculated using {len(selected_gas_tcs)} channels (weighted average):")
        for tc in selected_gas_tcs:
            # Clean up display name
            tc_name = tc
            prefixes_to_remove = ['/RTAC Data/', 'RTAC Data/', '/']
            for prefix in prefixes_to_remove:
                if tc_name.startswith(prefix):
                    tc_name = tc_name[len(prefix):]
                    break
            tc_name = tc_name.strip()
            print(f"  - {tc_name}: {gas_weights[tc]:.2f}%")
        
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
    
    # Prompt the user to select the INPUT directory (CSV files location)
    input_directory = filedialog.askdirectory(title="Select the Directory with CSV Files")
    
    if not input_directory:
        print("!!ERROR!! - No input directory selected. Exiting.")
    else:
        # Prompt the user to select the OUTPUT directory (where Excel files will be saved)
        output_directory = None
        if create_qgen:
            output_directory = filedialog.askdirectory(title="Select the Output Directory for Excel Files")
            
            if not output_directory:
                print("!!ERROR!! - No output directory selected. Exiting.")
                create_qgen = False
            else:
                print(f"\n✓ Output files will be saved to: {output_directory}")
        
        # Change the current working directory to the input folder
        os.chdir(input_directory)

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
                if create_qgen and output_directory:
                    create_qgen_analysis(csv_file, main_data, output_directory)
            
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
