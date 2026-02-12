import os
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from nptdms import TdmsFile
import glob
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

class TDMSMatcher:
    def __init__(self, csv_file_path, tdms_folder_path):
        """
        Initialize the TDMS Matcher
        
        Args:
            csv_file_path: Path to the CSV summary file
            tdms_folder_path: Path to folder containing TDMS files
        """
        self.csv_file_path = csv_file_path
        self.tdms_folder_path = tdms_folder_path
        self.summary_df = None
        self.tdms_data_combined = None
        self.group_name = None  # Store the actual group name
        
    def read_summary_csv(self):
        """Read the summary CSV file"""
        self.summary_df = pd.read_csv(self.csv_file_path)
        
        # Try multiple datetime formats
        datetime_formats = [
            '%m/%d/%Y %H:%M:%S',  # With seconds: 7/16/2025 23:21:24
            '%m/%d/%Y %H:%M',     # Without seconds: 7/16/2025 23:21
            '%Y-%m-%d %H:%M:%S',  # ISO format with seconds
            '%Y-%m-%d %H:%M',     # ISO format without seconds
            '%d/%m/%Y %H:%M:%S',  # DD/MM/YYYY with seconds
            '%d/%m/%Y %H:%M',     # DD/MM/YYYY without seconds
        ]
        
        parsed = False
        for fmt in datetime_formats:
            try:
                self.summary_df['DateTime'] = pd.to_datetime(
                    self.summary_df['DateTime'], 
                    format=fmt
                )
                print(f"✓ DateTime parsed successfully using format: {fmt}")
                parsed = True
                break
            except (ValueError, TypeError) as e:
                continue
        
        if not parsed:
            # If all formats fail, let pandas infer the format
            try:
                self.summary_df['DateTime'] = pd.to_datetime(
                    self.summary_df['DateTime'], 
                    infer_datetime_format=True
                )
                print(f"✓ DateTime parsed successfully using auto-detection")
            except Exception as e:
                raise ValueError(f"Could not parse DateTime column. Error: {e}")
        
        print(f"Summary CSV loaded: {len(self.summary_df)} rows")
        print(f"DateTime Range: {self.summary_df['DateTime'].min()} to {self.summary_df['DateTime'].max()}")
        return self.summary_df
    
    def parse_tdms_filename(self, filename):
        """
        Parse TDMS filename to extract datetime
        
        Args:
            filename: TDMS filename (e.g., '20250715_095042_BambiData_0001.tdms')
            
        Returns:
            datetime object or None if parsing fails
        """
        try:
            base_name = os.path.basename(filename)
            datetime_str = '_'.join(base_name.split('_')[:2])
            file_datetime = datetime.strptime(datetime_str, '%Y%m%d_%H%M%S')
            return file_datetime
        except Exception as e:
            print(f"Error parsing filename {filename}: {e}")
            return None
    
    def is_tdms_file_in_range(self, tdms_datetime, start_time, end_time):
        """
        Check if TDMS file is within 12 hours of the datetime range
        
        Args:
            tdms_datetime: datetime object from TDMS filename
            start_time: Start of datetime range
            end_time: End of datetime range
            
        Returns:
            Boolean indicating if file should be processed
        """
        if tdms_datetime is None:
            return False
        
        # Expand range by 12 hours on both sides
        range_start = start_time - timedelta(hours=12)
        range_end = end_time + timedelta(hours=12)
        
        return range_start <= tdms_datetime <= range_end
    
    def find_data_group(self, tdms_file):
        """
        Find the group that contains the actual data
        
        Args:
            tdms_file: TdmsFile object
            
        Returns:
            Group name or None
        """
        groups = tdms_file.groups()
        
        # Try to find a group with channels (not Root)
        for group in groups:
            group_name = group.name
            if group_name and group_name.lower() != 'root':
                # Check if this group has channels
                if len(group.channels()) > 0:
                    print(f"  Found data group: '{group_name}' with {len(group.channels())} channels")
                    return group_name
        
        return None
    
    def read_tdms_file(self, tdms_file_path, channels_needed):
        """
        Read TDMS file and extract specific channels
        
        Args:
            tdms_file_path: Path to TDMS file
            channels_needed: List of channel names to extract
            
        Returns:
            DataFrame with selected TDMS data
        """
        try:
            tdms_file = TdmsFile.read(tdms_file_path)
            
            # Find the data group if not already set
            if self.group_name is None:
                self.group_name = self.find_data_group(tdms_file)
                if self.group_name is None:
                    print(f"  Error: No data group found in {os.path.basename(tdms_file_path)}")
                    return None
            
            group = tdms_file[self.group_name]
            
            data_dict = {}
            
            # Extract only the channels we need
            for channel_name in channels_needed:
                try:
                    channel = group[channel_name]
                    data_dict[channel_name] = channel[:]
                except KeyError:
                    print(f"  Warning: Channel '{channel_name}' not found in {os.path.basename(tdms_file_path)}")
            
            if not data_dict:
                return None
            
            df = pd.DataFrame(data_dict)
            
            # Convert Excel format datetime to readable format
            if 'Date/Time (Excel Format)' in df.columns:
                # Excel datetime: days since 1899-12-30
                df['DateTime'] = pd.to_datetime(
                    df['Date/Time (Excel Format)'], 
                    unit='D', 
                    origin='1899-12-30'
                )
            
            print(f"  Successfully read {os.path.basename(tdms_file_path)}: {len(df)} rows")
            return df
            
        except Exception as e:
            print(f"  Error reading {tdms_file_path}: {e}")
            return None
    
    def load_all_relevant_tdms_data(self, channels_needed):
        """
        Load and combine all relevant TDMS files
        
        Args:
            channels_needed: List of channel names to extract
            
        Returns:
            Combined DataFrame with all TDMS data
        """
        csv_start = self.summary_df['DateTime'].min()
        csv_end = self.summary_df['DateTime'].max()
        
        # Get all TDMS files
        tdms_files = glob.glob(os.path.join(self.tdms_folder_path, '*.tdms'))
        print(f"\nFound {len(tdms_files)} TDMS files")
        
        all_dataframes = []
        processed_count = 0
        
        for tdms_file in tdms_files:
            tdms_datetime = self.parse_tdms_filename(tdms_file)
            
            if tdms_datetime:
                print(f"\nChecking {os.path.basename(tdms_file)} (Start: {tdms_datetime})")
            
            if self.is_tdms_file_in_range(tdms_datetime, csv_start, csv_end):
                print(f"  ✓ Within range - Reading...")
                df = self.read_tdms_file(tdms_file, channels_needed)
                
                if df is not None and 'DateTime' in df.columns:
                    all_dataframes.append(df)
                    processed_count += 1
            else:
                print(f"  ✗ Outside range - Skipping")
        
        print(f"\n{'='*60}")
        print(f"Processed {processed_count} TDMS files")
        print(f"{'='*60}\n")
        
        if all_dataframes:
            combined_df = pd.concat(all_dataframes, ignore_index=True)
            # Sort by DateTime for efficient matching
            combined_df = combined_df.sort_values('DateTime').reset_index(drop=True)
            print(f"Combined TDMS data: {len(combined_df)} total rows")
            print(f"TDMS DateTime Range: {combined_df['DateTime'].min()} to {combined_df['DateTime'].max()}")
            return combined_df
        else:
            print("No TDMS data loaded!")
            return None
    
    def find_nearest_match(self, target_datetime, tdms_df, tolerance_seconds=60):
        """
        Find the nearest matching TDMS data point for a given datetime
        
        Args:
            target_datetime: DateTime to match
            tdms_df: DataFrame with TDMS data
            tolerance_seconds: Maximum time difference in seconds for a valid match
            
        Returns:
            Matched row data or None if no match within tolerance
        """
        # Calculate time differences
        time_diffs = abs((tdms_df['DateTime'] - target_datetime).dt.total_seconds())
        
        # Find the index of minimum difference
        min_idx = time_diffs.idxmin()
        min_diff = time_diffs[min_idx]
        
        # Check if within tolerance
        if min_diff <= tolerance_seconds:
            return tdms_df.loc[min_idx], min_diff
        else:
            return None, min_diff
    
    def match_and_add_tdms_data(self, tdms_channel_name, new_column_name=None, tolerance_seconds=60):
        """
        Match TDMS data to summary CSV and add as new column
        
        Args:
            tdms_channel_name: Name of the TDMS channel to extract (e.g., '9211_6 TC2 T19 (C)')
            new_column_name: Name for the new column in summary CSV (default: same as channel name)
            tolerance_seconds: Maximum time difference in seconds for a valid match
        """
        if new_column_name is None:
            new_column_name = tdms_channel_name
        
        # Load summary CSV
        self.read_summary_csv()
        
        # Load all relevant TDMS data
        channels_needed = ['Date/Time (Excel Format)', tdms_channel_name]
        self.tdms_data_combined = self.load_all_relevant_tdms_data(channels_needed)
        
        if self.tdms_data_combined is None or tdms_channel_name not in self.tdms_data_combined.columns:
            print(f"Error: Could not load TDMS data for channel '{tdms_channel_name}'")
            return
        
        # Match each row in summary CSV
        print(f"\n{'='*60}")
        print(f"Matching TDMS data to Summary CSV...")
        print(f"Tolerance: ±{tolerance_seconds} seconds")
        print(f"{'='*60}\n")
        
        matched_values = []
        match_time_diffs = []
        match_count = 0
        no_match_count = 0
        
        for idx, row in self.summary_df.iterrows():
            target_dt = row['DateTime']
            
            matched_row, time_diff = self.find_nearest_match(
                target_dt, 
                self.tdms_data_combined, 
                tolerance_seconds
            )
            
            if matched_row is not None:
                value = matched_row[tdms_channel_name]
                matched_values.append(value)
                match_time_diffs.append(time_diff)
                match_count += 1
                print(f"Row {idx+1}: {target_dt} → Matched (Δ={time_diff:.1f}s, Value={value:.2f}°C)")
            else:
                matched_values.append(np.nan)
                match_time_diffs.append(time_diff)
                no_match_count += 1
                print(f"Row {idx+1}: {target_dt} → No match (closest: {time_diff:.1f}s away)")
        
        # Add new columns to summary DataFrame
        self.summary_df[new_column_name] = matched_values
        self.summary_df[f'{new_column_name}_TimeDiff_s'] = match_time_diffs
        
        print(f"\n{'='*60}")
        print(f"Matching Summary:")
        print(f"  Total rows: {len(self.summary_df)}")
        print(f"  Matched: {match_count}")
        print(f"  No match: {no_match_count}")
        print(f"  Match rate: {match_count/len(self.summary_df)*100:.1f}%")
        if match_count > 0:
            print(f"  Avg time difference: {np.mean([d for d in match_time_diffs if not np.isnan(d)]):.2f}s")
            print(f"  Max time difference: {max([d for d in match_time_diffs if not np.isnan(d)]):.2f}s")
        print(f"{'='*60}\n")
        
        return self.summary_df
    
    def save_updated_csv(self, output_path=None):
        """
        Save the updated summary CSV with new TDMS data
        
        Args:
            output_path: Path for output file (default: overwrite original)
        """
        if self.summary_df is None:
            print("No data to save!")
            return
        
        if output_path is None:
            output_path = self.csv_file_path
        
        self.summary_df.to_csv(output_path, index=False)
        print(f"✓ Updated CSV saved to: {output_path}")
        print(f"  Total columns: {len(self.summary_df.columns)}")
        print(f"  Column names: {list(self.summary_df.columns)}")


class TDMSMatcherGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("TDMS Data Matcher")
        self.root.geometry("700x650")
        self.root.resizable(False, False)
        
        # Variables
        self.csv_path = tk.StringVar()
        self.tdms_folder = tk.StringVar()
        self.tdms_channel = tk.StringVar()
        self.new_column_name = tk.StringVar(value="T19_Cooling (°C)")
        self.tolerance = tk.IntVar(value=60)
        self.output_path = tk.StringVar()
        
        self.available_channels = []
        self.group_name = None
        self.matcher = None
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the user interface"""
        # Title
        title_frame = tk.Frame(self.root, bg="#2c3e50", height=60)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(
            title_frame, 
            text="TDMS Data Matcher", 
            font=("Arial", 18, "bold"),
            bg="#2c3e50",
            fg="white"
        )
        title_label.pack(pady=15)
        
        # Main content frame
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # CSV File Selection
        csv_frame = tk.LabelFrame(main_frame, text="1. Select Summary CSV File", font=("Arial", 10, "bold"), padx=10, pady=10)
        csv_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Entry(csv_frame, textvariable=self.csv_path, width=60, state='readonly').pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(csv_frame, text="Browse...", command=self.browse_csv, width=12).pack(side=tk.LEFT)
        
        # TDMS Folder Selection
        tdms_frame = tk.LabelFrame(main_frame, text="2. Select TDMS Folder", font=("Arial", 10, "bold"), padx=10, pady=10)
        tdms_frame.pack(fill=tk.X, pady=(0, 10))
        
        folder_entry_frame = tk.Frame(tdms_frame)
        folder_entry_frame.pack(fill=tk.X)
        
        tk.Entry(folder_entry_frame, textvariable=self.tdms_folder, width=60, state='readonly').pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(folder_entry_frame, text="Browse...", command=self.browse_folder, width=12).pack(side=tk.LEFT)
        
        # Load channels button (appears after folder selection)
        self.load_channels_button = tk.Button(
            tdms_frame, 
            text="⟳ Load Available Channels", 
            command=self.load_channels,
            state=tk.DISABLED,
            bg="#3498db",
            fg="white",
            font=("Arial", 9, "bold")
        )
        self.load_channels_button.pack(pady=(10, 0))
        
        # Configuration Frame
        config_frame = tk.LabelFrame(main_frame, text="3. Configuration", font=("Arial", 10, "bold"), padx=10, pady=10)
        config_frame.pack(fill=tk.X, pady=(0, 10))
        
        # TDMS Channel Dropdown
        tk.Label(config_frame, text="TDMS Channel:", anchor='w').grid(row=0, column=0, sticky='w', pady=5)
        channel_frame = tk.Frame(config_frame)
        channel_frame.grid(row=0, column=1, sticky='w', padx=10, pady=5)
        
        self.channel_dropdown = ttk.Combobox(
            channel_frame, 
            textvariable=self.tdms_channel, 
            width=45,
            state='readonly'
        )
        self.channel_dropdown.pack(side=tk.LEFT)
        self.channel_dropdown['values'] = ["Select TDMS folder first..."]
        self.channel_dropdown.current(0)
        
        # New Column Name
        tk.Label(config_frame, text="New Column Name:", anchor='w').grid(row=1, column=0, sticky='w', pady=5)
        tk.Entry(config_frame, textvariable=self.new_column_name, width=47).grid(row=1, column=1, sticky='w', padx=10, pady=5)
        
        # Tolerance
        tk.Label(config_frame, text="Tolerance (seconds):", anchor='w').grid(row=2, column=0, sticky='w', pady=5)
        tolerance_frame = tk.Frame(config_frame)
        tolerance_frame.grid(row=2, column=1, sticky='w', padx=10, pady=5)
        tk.Spinbox(tolerance_frame, from_=1, to=300, textvariable=self.tolerance, width=10).pack(side=tk.LEFT)
        tk.Label(tolerance_frame, text="(±seconds for matching)").pack(side=tk.LEFT, padx=5)
        
        # Output File Selection
        output_frame = tk.LabelFrame(main_frame, text="4. Output Location (Optional - leave blank to overwrite)", font=("Arial", 10, "bold"), padx=10, pady=10)
        output_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Entry(output_frame, textvariable=self.output_path, width=60).pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(output_frame, text="Browse...", command=self.browse_output, width=12).pack(side=tk.LEFT)
        
        # Progress Frame
        progress_frame = tk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.progress_label = tk.Label(progress_frame, text="Ready", fg="blue")
        self.progress_label.pack()
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='indeterminate')
        self.progress_bar.pack(fill=tk.X)
        
        # Process Button
        button_frame = tk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        self.process_button = tk.Button(
            button_frame, 
            text="Process Data", 
            command=self.process_data,
            bg="#27ae60",
            fg="white",
            font=("Arial", 12, "bold"),
            width=20,
            height=2,
            cursor="hand2"
        )
        self.process_button.pack()
        
        # Status text
        status_frame = tk.LabelFrame(main_frame, text="Status", font=("Arial", 10, "bold"))
        status_frame.pack(fill=tk.BOTH, expand=True)
        
        self.status_text = tk.Text(status_frame, height=6, wrap=tk.WORD)
        self.status_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        scrollbar = tk.Scrollbar(self.status_text)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.status_text.yview)
        
    def browse_csv(self):
        """Open file dialog to select CSV file"""
        filename = filedialog.askopenfilename(
            title="Select Summary CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.csv_path.set(filename)
            self.log_status(f"CSV selected: {os.path.basename(filename)}")
            
    def browse_folder(self):
        """Open folder dialog to select TDMS folder"""
        folder = filedialog.askdirectory(
            title="Select TDMS Folder"
        )
        if folder:
            self.tdms_folder.set(folder)
            # Count TDMS files
            tdms_count = len(glob.glob(os.path.join(folder, "*.tdms")))
            self.log_status(f"TDMS folder selected: {tdms_count} TDMS files found")
            
            # Enable the load channels button
            self.load_channels_button.config(state=tk.NORMAL)
            self.log_status("Click 'Load Available Channels' to read channel list")
            
    def load_channels(self):
        """Load channel names from a TDMS file"""
        tdms_files = glob.glob(os.path.join(self.tdms_folder.get(), "*.tdms"))
        
        if not tdms_files:
            messagebox.showerror("Error", "No TDMS files found in the selected folder!")
            return
        
        self.log_status(f"Reading channels from {os.path.basename(tdms_files[0])}...")
        
        try:
            # Read the first TDMS file to get channel names
            tdms_file = TdmsFile.read(tdms_files[0])
            
            # List all groups
            groups = tdms_file.groups()
            self.log_status(f"Found {len(groups)} group(s) in TDMS file")
            
            # Find a group with channels (not Root)
            data_group = None
            for group in groups:
                group_name = group.name
                self.log_status(f"  Group: '{group_name}' ({len(group.channels())} channels)")
                
                if group_name and group_name.lower() != 'root' and len(group.channels()) > 0:
                    data_group = group
                    self.group_name = group_name
                    break
            
            if data_group is None:
                messagebox.showerror("Error", "No data group with channels found in TDMS file!")
                return
            
            # Get all channel names
            self.available_channels = [channel.name for channel in data_group.channels()]
            
            # Filter out the datetime channel for cleaner list
            display_channels = [ch for ch in self.available_channels if 'Date/Time' not in ch]
            
            if display_channels:
                self.channel_dropdown['values'] = display_channels
                
                # Try to set default to T19 channel if available
                default_channel = None
                for ch in display_channels:
                    if 'T19' in ch or '9211_6 TC2' in ch:
                        default_channel = ch
                        break
                
                if default_channel:
                    self.channel_dropdown.set(default_channel)
                    self.new_column_name.set("T19_Cooling (°C)")
                else:
                    self.channel_dropdown.current(0)
                
                self.log_status(f"✓ Found {len(display_channels)} data channels in group '{self.group_name}'")
                messagebox.showinfo(
                    "Success", 
                    f"Loaded {len(display_channels)} channels from TDMS file!\n\nGroup: '{self.group_name}'\n\nSelect a channel from the dropdown."
                )
            else:
                messagebox.showerror("Error", "No data channels found in TDMS file!")
                
        except Exception as e:
            self.log_status(f"✗ Error loading channels: {str(e)}")
            messagebox.showerror("Error", f"Could not read TDMS file:\n\n{str(e)}")
    
    def browse_output(self):
        """Open file dialog to select output location"""
        filename = filedialog.asksaveasfilename(
            title="Save Output CSV As",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.output_path.set(filename)
            self.log_status(f"Output location: {os.path.basename(filename)}")
    
    def log_status(self, message):
        """Add message to status text box"""
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)
        self.root.update()
    
    def validate_inputs(self):
        """Validate user inputs"""
        if not self.csv_path.get():
            messagebox.showerror("Error", "Please select a CSV file!")
            return False
        
        if not self.tdms_folder.get():
            messagebox.showerror("Error", "Please select a TDMS folder!")
            return False
        
        if not os.path.exists(self.csv_path.get()):
            messagebox.showerror("Error", "CSV file does not exist!")
            return False
        
        if not os.path.exists(self.tdms_folder.get()):
            messagebox.showerror("Error", "TDMS folder does not exist!")
            return False
        
        if not self.tdms_channel.get() or self.tdms_channel.get() == "Select TDMS folder first...":
            messagebox.showerror("Error", "Please select a TDMS channel!\n\nClick 'Load Available Channels' first.")
            return False
        
        if not self.new_column_name.get():
            messagebox.showerror("Error", "Please enter a new column name!")
            return False
        
        return True
    
    def process_data_thread(self):
        """Process data in separate thread to prevent GUI freezing"""
        try:
            self.progress_bar.start(10)
            self.process_button.config(state=tk.DISABLED)
            self.progress_label.config(text="Processing...", fg="orange")
            
            # Create matcher
            self.log_status("\n" + "="*60)
            self.log_status("Starting data processing...")
            self.log_status("="*60)
            
            self.matcher = TDMSMatcher(
                self.csv_path.get(),
                self.tdms_folder.get()
            )
            
            # Set the group name from GUI
            self.matcher.group_name = self.group_name
            
            # Match and add data
            self.log_status(f"Channel: {self.tdms_channel.get()}")
            self.log_status(f"Tolerance: ±{self.tolerance.get()} seconds")
            
            updated_df = self.matcher.match_and_add_tdms_data(
                tdms_channel_name=self.tdms_channel.get(),
                new_column_name=self.new_column_name.get(),
                tolerance_seconds=self.tolerance.get()
            )
            
            # Determine output path
            output = self.output_path.get() if self.output_path.get() else None
            
            # Save
            self.matcher.save_updated_csv(output_path=output)
            
            self.log_status("="*60)
            self.log_status("✓ Processing completed successfully!")
            self.log_status("="*60)
            
            self.progress_label.config(text="Completed!", fg="green")
            
            # Show success message
            final_path = output if output else self.csv_path.get()
            messagebox.showinfo(
                "Success", 
                f"Data processed successfully!\n\nOutput saved to:\n{final_path}"
            )
            
        except Exception as e:
            self.log_status(f"\n✗ Error: {str(e)}")
            self.progress_label.config(text="Error occurred!", fg="red")
            messagebox.showerror("Error", f"An error occurred:\n\n{str(e)}")
        
        finally:
            self.progress_bar.stop()
            self.process_button.config(state=tk.NORMAL)
    
    def process_data(self):
        """Validate and start processing"""
        if not self.validate_inputs():
            return
        
        # Clear status
        self.status_text.delete(1.0, tk.END)
        
        # Run in separate thread
        thread = threading.Thread(target=self.process_data_thread)
        thread.daemon = True
        thread.start()
    
    def run(self):
        """Start the GUI"""
        self.root.mainloop()


# Run the GUI application
if __name__ == "__main__":
    app = TDMSMatcherGUI()
    app.run()
