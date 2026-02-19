# Compatibility fix for Python 3.8 and older
import sys
if sys.version_info < (3, 9):
    import typing
    if not hasattr(typing, 'get_origin'):
        # Monkey patch for older Python versions
        import collections.abc
        typing.get_origin = lambda t: getattr(t, '__origin__', None)
        typing.get_args = lambda t: getattr(t, '__args__', ())

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import nidaqmx
from nidaqmx.constants import ThermocoupleType, TemperatureUnits
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
import numpy as np
from datetime import datetime, timedelta
import threading
import time
from nptdms import TdmsWriter, RootObject, GroupObject, ChannelObject
import csv
import os
import json


class ThermocopleDAQGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("NI cDAQ Thermocouple Data Acquisition")
        self.root.geometry("1600x900")
        
        # Data storage
        self.acquisition_running = False
        self.data_lock = threading.Lock()
        self.timestamps = []
        self.temperature_data = []
        
        # File handling
        self.current_file_start_time = None
        self.tdms_writer = None
        self.tdms_file = None
        self.csv_file = None
        self.csv_writer = None
        
        # Display update throttling
        self.last_display_update = time.time()
        self.display_update_interval = 0.5  # Update display every 0.5 seconds
        
        # TDMS buffering
        self.tdms_buffer = []
        self.tdms_buffer_size = 10  # Write every 10 samples
        
        # Acquisition parameters
        self.acquisition_rate = 1.0  # Hz
        self.file_rotation_interval = 12 * 3600  # seconds
        
        # Plot styling
        self.colors = ['red', 'blue', 'green', 'orange', 'purple', 'brown', 'pink', 'gray', 'cyan', 'magenta']
        self.markers = ['o', 's', '^', 'D', 'v', '<', '>', 'p', '*', 'h']
        
        # Channel configuration lists (will be populated based on num_channels)
        self.channel_selection_vars = []
        self.channel_label_vars = []
        self.channel_style_vars = []
        self.channel_color_vars = []
        self.channel_frames = []
        self.temp_display_labels = []  # NEW: Store temp display label widgets
        
        # Create GUI
        self.create_widgets()
        
    def create_widgets(self):
        # Main container with notebook (tabs)
        main_container = ttk.Frame(self.root, padding="10")
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(0, weight=1)
        
        # Create notebook (tabs)
        self.notebook = ttk.Notebook(main_container)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Create Tab 1: Setup
        self.setup_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.setup_tab, text="Setup")
        
        # Create Tab 2: Run
        self.run_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.run_tab, text="Run")
        
        # Build Setup Tab
        self.create_setup_tab()
        
        # Build Run Tab
        self.create_run_tab()
        
    def create_setup_tab(self):
        """Create the Setup tab with all controls and displays"""
        # Configure grid weights
        self.setup_tab.columnconfigure(0, weight=0, minsize=350)
        self.setup_tab.columnconfigure(1, weight=0, minsize=350)
        self.setup_tab.columnconfigure(2, weight=0, minsize=550)
        self.setup_tab.columnconfigure(3, weight=1)
        self.setup_tab.rowconfigure(0, weight=1)
        
        # Store widgets to disable during acquisition
        self.config_widgets = []
        self.channel_name_entries = []
        
        # ===== COLUMN 1: DAQ and Experiment Configuration =====
        col1_frame = ttk.Frame(self.setup_tab)
        col1_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        
        # Store col1_frame to disable all its children
        self.col1_frame = col1_frame
        
        # Make column 1 scrollable
        canvas_col1 = tk.Canvas(col1_frame, width=340)
        scrollbar_col1 = ttk.Scrollbar(col1_frame, orient="vertical", command=canvas_col1.yview)
        scrollable_col1 = ttk.Frame(canvas_col1)
        
        scrollable_col1.bind(
            "<Configure>",
            lambda e: canvas_col1.configure(scrollregion=canvas_col1.bbox("all"))
        )
        
        canvas_col1.create_window((0, 0), window=scrollable_col1, anchor="nw")
        canvas_col1.configure(yscrollcommand=scrollbar_col1.set)
        
        canvas_col1.pack(side="left", fill="both", expand=True)
        scrollbar_col1.pack(side="right", fill="y")
        
        row = 0
        
        # ===== Configuration Save/Load =====
        config_frame = ttk.LabelFrame(scrollable_col1, text="Configuration Management", padding="10")
        config_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        row += 1
        
        # Config folder path
        ttk.Label(config_frame, text="Config Folder:").grid(row=0, column=0, sticky=tk.W, pady=5)
        
        path_frame = ttk.Frame(config_frame)
        path_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        self.config_path_var = tk.StringVar(value=os.getcwd())
        config_path_entry = ttk.Entry(path_frame, textvariable=self.config_path_var, width=20)
        config_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        ttk.Button(path_frame, text="Browse", command=self.browse_config_folder, width=8).pack(side=tk.LEFT)
        
        # Save/Load buttons
        button_frame = ttk.Frame(config_frame)
        button_frame.grid(row=2, column=0, pady=5)
        
        ttk.Button(button_frame, text="Save Configuration", 
                  command=self.save_configuration, width=15).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Load Configuration", 
                  command=self.load_configuration, width=15).pack(side=tk.LEFT, padx=2)
        
        # ===== DAQ Hardware Configuration =====
        daq_config_frame = ttk.LabelFrame(scrollable_col1, text="DAQ Hardware Configuration", padding="10")
        daq_config_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        row += 1
        
        # Device Name
        ttk.Label(daq_config_frame, text="Device Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.device_name_var = tk.StringVar(value="cDAQ2Mod1")
        ttk.Entry(daq_config_frame, textvariable=self.device_name_var, width=25).grid(row=0, column=1, pady=5, padx=5)
        ttk.Label(daq_config_frame, text="(e.g., cDAQ1Mod1, Dev1)", font=("Arial", 8), foreground="gray").grid(row=1, column=1, sticky=tk.W, padx=5)
        
        # Number of Channels
        ttk.Label(daq_config_frame, text="Number of Channels:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.num_channels_var = tk.IntVar(value=4)
        num_channels_spinbox = ttk.Spinbox(daq_config_frame, from_=1, to=32, textvariable=self.num_channels_var, width=23)
        num_channels_spinbox.grid(row=2, column=1, pady=5, padx=5)
        
        # Starting Channel
        ttk.Label(daq_config_frame, text="Starting Channel:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.start_channel_var = tk.IntVar(value=0)
        start_channel_spinbox = ttk.Spinbox(daq_config_frame, from_=0, to=31, textvariable=self.start_channel_var, width=23)
        start_channel_spinbox.grid(row=3, column=1, pady=5, padx=5)
        ttk.Label(daq_config_frame, text="(ai0, ai1, ai2, ...)", font=("Arial", 8), foreground="gray").grid(row=4, column=1, sticky=tk.W, padx=5)
        
        # Thermocouple Type
        ttk.Label(daq_config_frame, text="Thermocouple Type:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.tc_type_var = tk.StringVar(value="K")
        tc_type_combo = ttk.Combobox(daq_config_frame, textvariable=self.tc_type_var,
                                     values=['B', 'E', 'J', 'K', 'N', 'R', 'S', 'T'],
                                     state="readonly", width=22)
        tc_type_combo.grid(row=5, column=1, pady=5, padx=5)
        
        # Temperature Units
        ttk.Label(daq_config_frame, text="Temperature Units:").grid(row=6, column=0, sticky=tk.W, pady=5)
        self.temp_units_var = tk.StringVar(value="Celsius")
        temp_units_combo = ttk.Combobox(daq_config_frame, textvariable=self.temp_units_var,
                                       values=['Celsius', 'Fahrenheit', 'Kelvin'],
                                       state="readonly", width=22)
        temp_units_combo.grid(row=6, column=1, pady=5, padx=5)
        
        # Apply Configuration Button
        apply_button = ttk.Button(daq_config_frame, text="Apply DAQ Configuration", 
                                 command=self.apply_daq_config, width=28)
        apply_button.grid(row=7, column=0, columnspan=2, pady=10)
        
        # ===== Experiment Configuration =====
        exp_config_frame = ttk.LabelFrame(scrollable_col1, text="Experiment Configuration", padding="10")
        exp_config_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        row += 1
        
        # Experiment Name
        ttk.Label(exp_config_frame, text="Experiment Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.experiment_name_var = tk.StringVar(value="TC_Experiment")
        ttk.Entry(exp_config_frame, textvariable=self.experiment_name_var, width=25).grid(row=0, column=1, pady=5, padx=5)
        
        # Data Save Folder
        ttk.Label(exp_config_frame, text="Data Save Folder:").grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        data_path_frame = ttk.Frame(exp_config_frame)
        data_path_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.data_path_var = tk.StringVar(value=os.getcwd())
        data_path_entry = ttk.Entry(data_path_frame, textvariable=self.data_path_var, width=20)
        data_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        ttk.Button(data_path_frame, text="Browse", command=self.browse_data_folder, width=8).pack(side=tk.LEFT)
        
        # Acquisition Rate
        ttk.Label(exp_config_frame, text="Acquisition Rate:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.acq_rate_var = tk.StringVar(value="1 Hz")
        acq_rate_dropdown = ttk.Combobox(exp_config_frame, textvariable=self.acq_rate_var, 
                                          values=["0.1 Hz", "1 Hz", "2 Hz", "5 Hz", "10 Hz"],
                                          state="readonly", width=22)
        acq_rate_dropdown.grid(row=3, column=1, pady=5, padx=5)
        acq_rate_dropdown.bind("<<ComboboxSelected>>", self.update_acquisition_rate)
        
        # X-axis Range
        ttk.Label(exp_config_frame, text="Plot Time Window:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.time_window_var = tk.StringVar(value="1 hour")
        time_window_dropdown = ttk.Combobox(exp_config_frame, textvariable=self.time_window_var,
                                             values=["30 minutes", "1 hour", "6 hours", "1 day", 
                                                    "7 days", "30 days"],
                                             state="readonly", width=22)
        time_window_dropdown.grid(row=4, column=1, pady=5, padx=5)
        
        # CSV Logging Option
        self.csv_logging_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(exp_config_frame, text="Enable CSV Logging", 
                       variable=self.csv_logging_var).grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # ===== COLUMN 2: Channel Configuration =====
        col2_frame = ttk.LabelFrame(self.setup_tab, text="Channel Configuration", padding="10")
        col2_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5)
        
        # Make column 2 scrollable
        canvas_col2 = tk.Canvas(col2_frame, width=340)
        scrollbar_col2 = ttk.Scrollbar(col2_frame, orient="vertical", command=canvas_col2.yview)
        self.channels_container = ttk.Frame(canvas_col2)
        
        self.channels_container.bind(
            "<Configure>",
            lambda e: canvas_col2.configure(scrollregion=canvas_col2.bbox("all"))
        )
        
        canvas_col2.create_window((0, 0), window=self.channels_container, anchor="nw")
        canvas_col2.configure(yscrollcommand=scrollbar_col2.set)
        
        canvas_col2.pack(side="left", fill="both", expand=True)
        scrollbar_col2.pack(side="right", fill="y")
        
        # Placeholder label
        self.channel_placeholder = ttk.Label(self.channels_container, 
                                            text="Apply DAQ Configuration\nto see channel settings",
                                            font=("Arial", 10), 
                                            foreground="gray")
        self.channel_placeholder.pack(pady=50)
        
        # ===== COLUMN 3: Control Buttons and Status Log =====
        col3_frame = ttk.Frame(self.setup_tab)
        col3_frame.grid(row=0, column=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
        
        # Start/Stop Buttons
        button_frame = ttk.LabelFrame(col3_frame, text="Acquisition Control", padding="10")
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.start_button = ttk.Button(button_frame, text="Start Acquisition", 
                                       command=self.start_acquisition, width=28)
        self.start_button.pack(pady=5)
        
        self.stop_button = ttk.Button(button_frame, text="Stop Acquisition", 
                                      command=self.stop_acquisition, state="disabled", width=28)
        self.stop_button.pack(pady=5)
        
        # Status Log
        status_frame = ttk.LabelFrame(col3_frame, text="Status Log", padding="10")
        status_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create scrollable text widget for status log
        log_scroll = ttk.Scrollbar(status_frame, orient="vertical")
        self.status_log = tk.Text(status_frame, height=20, width=55, wrap=tk.WORD, 
                                  yscrollcommand=log_scroll.set, font=("Courier", 9))
        log_scroll.config(command=self.status_log.yview)
        
        self.status_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Make log read-only
        self.status_log.config(state=tk.DISABLED)
        
        # Add initial message
        self.log_status("System initialized")
        self.log_status("Ready - Please apply DAQ configuration")
        
        # ===== COLUMN 4: Temperature Display and Statistics =====
        col4_frame = ttk.Frame(self.setup_tab)
        col4_frame.grid(row=0, column=3, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(10, 0))
        col4_frame.rowconfigure(1, weight=1)
        col4_frame.columnconfigure(0, weight=1)
        
        # Temperature Display
        display_frame = ttk.LabelFrame(col4_frame, text="Current Temperature Readings", padding="10")
        display_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N), pady=(0, 10))
        
        # Scrollable temperature display
        temp_canvas = tk.Canvas(display_frame, height=200)
        temp_scrollbar = ttk.Scrollbar(display_frame, orient="vertical", command=temp_canvas.yview)
        self.temp_display_frame = ttk.Frame(temp_canvas)
        
        self.temp_display_frame.bind(
            "<Configure>",
            lambda e: temp_canvas.configure(scrollregion=temp_canvas.bbox("all"))
        )
        
        temp_canvas.create_window((0, 0), window=self.temp_display_frame, anchor="nw")
        temp_canvas.configure(yscrollcommand=temp_scrollbar.set)
        
        temp_canvas.pack(side="left", fill="both", expand=True)
        temp_scrollbar.pack(side="right", fill="y")
        
        self.temp_labels = []
        
        # Statistics Display
        stats_frame = ttk.LabelFrame(col4_frame, text="Statistics", padding="10")
        stats_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollable statistics display
        stats_canvas = tk.Canvas(stats_frame)
        stats_scrollbar = ttk.Scrollbar(stats_frame, orient="vertical", command=stats_canvas.yview)
        self.stats_display_frame = ttk.Frame(stats_canvas)
        
        self.stats_display_frame.bind(
            "<Configure>",
            lambda e: stats_canvas.configure(scrollregion=stats_canvas.bbox("all"))
        )
        
        stats_canvas.create_window((0, 0), window=self.stats_display_frame, anchor="nw")
        stats_canvas.configure(yscrollcommand=stats_scrollbar.set)
        
        stats_canvas.pack(side="left", fill="both", expand=True)
        stats_scrollbar.pack(side="right", fill="y")
        
        self.stats_labels = []
    
    def log_status(self, message):
        """Add a message to the status log with timestamp"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"
        
        self.status_log.config(state=tk.NORMAL)
        self.status_log.insert(tk.END, log_message)
        self.status_log.see(tk.END)  # Auto-scroll to bottom
        self.status_log.config(state=tk.DISABLED)
    
    def browse_config_folder(self):
        """Browse for configuration folder"""
        folder = filedialog.askdirectory(initialdir=self.config_path_var.get())
        if folder:
            self.config_path_var.set(folder)
            self.log_status(f"Config folder set to: {folder}")
    
    def browse_data_folder(self):
        """Browse for data save folder"""
        folder = filedialog.askdirectory(initialdir=self.data_path_var.get())
        if folder:
            self.data_path_var.set(folder)
            self.log_status(f"Data folder set to: {folder}")
    
    def save_configuration(self):
        """Save current GUI configuration to JSON file"""
        try:
            # Prepare configuration dictionary
            config = {
                'device_name': self.device_name_var.get(),
                'num_channels': self.num_channels_var.get(),
                'start_channel': self.start_channel_var.get(),
                'tc_type': self.tc_type_var.get(),
                'temp_units': self.temp_units_var.get(),
                'experiment_name': self.experiment_name_var.get(),
                'data_path': self.data_path_var.get(),
                'acquisition_rate': self.acq_rate_var.get(),
                'time_window': self.time_window_var.get(),
                'csv_logging': self.csv_logging_var.get(),
                'plot_settings': {
                    'plot_title': self.plot_title_var.get(),
                    'left_yaxis_title': self.left_yaxis_title_var.get(),
                    'right_yaxis_title': self.right_yaxis_title_var.get(),
                    'left_y_auto': self.left_y_auto_var.get(),
                    'left_y_min': self.left_y_min_var.get(),
                    'left_y_max': self.left_y_max_var.get(),
                    'right_y_auto': self.right_y_auto_var.get(),
                    'right_y_min': self.right_y_min_var.get(),
                    'right_y_max': self.right_y_max_var.get(),
                    'x_axis_auto': self.x_axis_auto_var.get()
                },
                'channels': []
            }
            
            # Save channel configurations if they exist
            if hasattr(self, 'channels') and len(self.channel_label_vars) > 0:
                for i in range(len(self.channel_label_vars)):
                    channel_config = {
                        'enabled': self.channel_selection_vars[i].get(),
                        'label': self.channel_label_vars[i].get(),
                        'color': self.channel_color_vars[i].get(),
                        'style': self.channel_style_vars[i].get(),
                        'yaxis': self.channel_yaxis_vars[i].get()
                    }
                    config['channels'].append(channel_config)
            
            # Ask for filename
            filename = filedialog.asksaveasfilename(
                initialdir=self.config_path_var.get(),
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
                title="Save Configuration"
            )
            
            if filename:
                with open(filename, 'w') as f:
                    json.dump(config, f, indent=4)
                self.log_status(f"Configuration saved to: {os.path.basename(filename)}")
                messagebox.showinfo("Success", "Configuration saved successfully!")
        
        except Exception as e:
            self.log_status(f"Error saving configuration: {str(e)}")
            messagebox.showerror("Error", f"Failed to save configuration:\n{str(e)}")
    
    def load_configuration(self):
        """Load GUI configuration from JSON file"""
        try:
            filename = filedialog.askopenfilename(
                initialdir=self.config_path_var.get(),
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
                title="Load Configuration"
            )
            
            if not filename:
                return
            
            with open(filename, 'r') as f:
                config = json.load(f)
            
            # Load basic settings
            self.device_name_var.set(config.get('device_name', 'cDAQ2Mod1'))
            self.num_channels_var.set(config.get('num_channels', 4))
            self.start_channel_var.set(config.get('start_channel', 0))
            self.tc_type_var.set(config.get('tc_type', 'K'))
            self.temp_units_var.set(config.get('temp_units', 'Celsius'))
            self.experiment_name_var.set(config.get('experiment_name', 'TC_Experiment'))
            self.data_path_var.set(config.get('data_path', os.getcwd()))
            self.acq_rate_var.set(config.get('acquisition_rate', '1 Hz'))
            self.time_window_var.set(config.get('time_window', '1 hour'))
            self.csv_logging_var.set(config.get('csv_logging', False))
            
            # Load plot settings
            if 'plot_settings' in config:
                ps = config['plot_settings']
                self.plot_title_var.set(ps.get('plot_title', 'Thermocouple Temperature vs Time'))
                self.left_yaxis_title_var.set(ps.get('left_yaxis_title', 'Temperature (°C)'))
                self.right_yaxis_title_var.set(ps.get('right_yaxis_title', 'Temperature (°C)'))
                self.left_y_auto_var.set(ps.get('left_y_auto', True))
                self.left_y_min_var.set(ps.get('left_y_min', '0'))
                self.left_y_max_var.set(ps.get('left_y_max', '100'))
                self.right_y_auto_var.set(ps.get('right_y_auto', True))
                self.right_y_min_var.set(ps.get('right_y_min', '0'))
                self.right_y_max_var.set(ps.get('right_y_max', '100'))
                self.x_axis_auto_var.set(ps.get('x_axis_auto', True))
            
            # Apply DAQ configuration
            self.apply_daq_config()
            
            # Load channel configurations if they exist
            if 'channels' in config and len(config['channels']) > 0:
                for i, channel_config in enumerate(config['channels']):
                    if i < len(self.channel_label_vars):
                        self.channel_selection_vars[i].set(channel_config.get('enabled', True))
                        self.channel_label_vars[i].set(channel_config.get('label', f'Channel {i}'))
                        self.channel_color_vars[i].set(channel_config.get('color', self.colors[i % len(self.colors)]))
                        self.channel_style_vars[i].set(channel_config.get('style', 'Line'))
                        self.channel_yaxis_vars[i].set(channel_config.get('yaxis', 'Left'))
            
            self.log_status(f"Configuration loaded from: {os.path.basename(filename)}")
            messagebox.showinfo("Success", "Configuration loaded successfully!")
        
        except Exception as e:
            self.log_status(f"Error loading configuration: {str(e)}")
            messagebox.showerror("Error", f"Failed to load configuration:\n{str(e)}")
    
    def apply_daq_config(self):
        """Apply DAQ configuration and rebuild channel configuration UI"""
        try:
            # Get configuration values
            device_name = self.device_name_var.get().strip()
            num_channels = self.num_channels_var.get()
            start_channel = self.start_channel_var.get()
            
            if not device_name:
                messagebox.showerror("Error", "Device name cannot be empty!")
                return
            
            if num_channels < 1 or num_channels > 32:
                messagebox.showerror("Error", "Number of channels must be between 1 and 32!")
                return
            
            # Store configuration
            self.device_name = device_name
            self.num_channels = num_channels
            self.start_channel = start_channel
            
            # Build channel list
            self.channels = [f"{self.device_name}/ai{start_channel + i}" for i in range(num_channels)]
            
            # Clear existing channel configuration UI
            for widget in self.channels_container.winfo_children():
                widget.destroy()
            
            # Clear existing temp and stats displays
            for widget in self.temp_display_frame.winfo_children():
                widget.destroy()
            for widget in self.stats_display_frame.winfo_children():
                widget.destroy()
            
            # Reset lists
            self.channel_selection_vars = []
            self.channel_label_vars = []
            self.channel_style_vars = []
            self.channel_color_vars = []
            self.channel_yaxis_vars = []  # NEW: Y-axis selection
            self.temp_labels = []
            self.stats_labels = []
            self.temperature_data = [[] for _ in range(num_channels)]
            self.channel_name_entries = []
            self.temp_display_labels = []  # NEW: Reset temp display labels
            
            # Create channel configuration UI
            for i in range(num_channels):
                ch_frame = ttk.LabelFrame(self.channels_container, 
                                         text=f"Channel {i} (ai{start_channel + i})", 
                                         padding="5")
                ch_frame.pack(fill=tk.X, pady=5, padx=5)
                
                # Enable checkbox
                var = tk.BooleanVar(value=True)
                self.channel_selection_vars.append(var)
                ttk.Checkbutton(ch_frame, text="Enable", variable=var).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=2)
                
                # Custom Label
                ttk.Label(ch_frame, text="Label:").grid(row=1, column=0, sticky=tk.W, pady=2)
                label_var = tk.StringVar(value=f"{self.device_name}_ai{start_channel + i}")
                self.channel_label_vars.append(label_var)
                label_entry = ttk.Entry(ch_frame, textvariable=label_var, width=18)
                label_entry.grid(row=1, column=1, pady=2, padx=2)
                self.channel_name_entries.append(label_entry)
                
                # Add trace to update Run tab when label changes
                label_var.trace_add('write', lambda *args, idx=i: self.on_channel_label_change(idx))
                
                # Color selection
                ttk.Label(ch_frame, text="Color:").grid(row=2, column=0, sticky=tk.W, pady=2)
                color_var = tk.StringVar(value=self.colors[i % len(self.colors)])
                self.channel_color_vars.append(color_var)
                color_combo = ttk.Combobox(ch_frame, textvariable=color_var,
                                          values=self.colors,
                                          state="readonly", width=15)
                color_combo.grid(row=2, column=1, pady=2, padx=2)
                
                # Style selection
                ttk.Label(ch_frame, text="Style:").grid(row=3, column=0, sticky=tk.W, pady=2)
                style_var = tk.StringVar(value="Line")
                self.channel_style_vars.append(style_var)
                style_combo = ttk.Combobox(ch_frame, textvariable=style_var,
                                          values=['Line', 'Dashed', 'Dotted', 'Dash-Dot', 
                                                 'Scatter', 'Line+Scatter'],
                                          state="readonly", width=15)
                style_combo.grid(row=3, column=1, pady=2, padx=2)
                
                # Y-Axis selection (NEW)
                ttk.Label(ch_frame, text="Y-Axis:").grid(row=4, column=0, sticky=tk.W, pady=2)
                yaxis_var = tk.StringVar(value="Left")
                self.channel_yaxis_vars.append(yaxis_var)
                yaxis_combo = ttk.Combobox(ch_frame, textvariable=yaxis_var,
                                           values=['Left', 'Right'],
                                           state="readonly", width=15)
                yaxis_combo.grid(row=4, column=1, pady=2, padx=2)
                
                # Create temperature display
                temp_frame = ttk.Frame(self.temp_display_frame)
                temp_frame.pack(fill=tk.X, pady=3)
                
                # Use custom label for display
                channel_display_label = ttk.Label(temp_frame, text=f"{self.device_name}_ai{start_channel + i}:", 
                                                 font=("Arial", 10, "bold"), width=20)  # Increased width
                channel_display_label.pack(side=tk.LEFT, padx=5)
                
                temp_label = ttk.Label(temp_frame, text="-- °C", font=("Arial", 11), foreground="darkgreen")
                temp_label.pack(side=tk.LEFT, padx=5)
                self.temp_labels.append(temp_label)
                
                # Store reference to update label dynamically
                self.temp_display_labels = getattr(self, 'temp_display_labels', [])
                self.temp_display_labels.append(channel_display_label)
                
                ttk.Label(temp_frame, text=f"Ch{i} (ai{start_channel + i}):", 
                         font=("Arial", 10, "bold"), width=12).pack(side=tk.LEFT, padx=5)
                temp_label = ttk.Label(temp_frame, text="-- °C", font=("Arial", 11), foreground="darkgreen")
                temp_label.pack(side=tk.LEFT, padx=5)
                self.temp_labels.append(temp_label)
                
                # Create statistics display
                stats_frame_inner = ttk.Frame(self.stats_display_frame)
                stats_frame_inner.pack(fill=tk.X, pady=2)
                
                ttk.Label(stats_frame_inner, text=f"Ch{i}:", width=5).pack(side=tk.LEFT)
                label = ttk.Label(stats_frame_inner, text="Min: -- | Max: -- | Avg: --", font=("Arial", 9))
                label.pack(side=tk.LEFT, padx=5)
                self.stats_labels.append(label)
            
            # Rebuild Run tab channel controls
            self.rebuild_run_tab_controls()
            
            self.log_status(f"DAQ Configuration Applied: {num_channels} channels on {device_name}")
            messagebox.showinfo("Success", f"DAQ Configuration Applied!\n\nDevice: {device_name}\nChannels: ai{start_channel} to ai{start_channel + num_channels - 1}")
            
        except Exception as e:
            self.log_status(f"Error applying DAQ config: {str(e)}")
            messagebox.showerror("Error", f"Failed to apply configuration:\n{str(e)}")
    
    def create_run_tab(self):
        """Create the Run tab with full-size plot and controls"""
        self.run_tab.columnconfigure(0, weight=1)
        self.run_tab.rowconfigure(2, weight=1)
        
        # ===== Top Control Bar - Row 1 =====
        self.control_bar = ttk.Frame(self.run_tab, padding="5")
        self.control_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # Plot Time Window Control
        ttk.Label(self.control_bar, text="Time Window:", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        
        time_window_run = ttk.Combobox(self.control_bar, textvariable=self.time_window_var,
                                        values=["30 minutes", "1 hour", "6 hours", "1 day", 
                                               "7 days", "30 days"],
                                        state="readonly", width=15)
        time_window_run.pack(side=tk.LEFT, padx=5)
        
        ttk.Separator(self.control_bar, orient='vertical').pack(side=tk.LEFT, fill=tk.Y, padx=10)
        
        # Channel Selection label
        ttk.Label(self.control_bar, text="Channels:", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        
        # Channel checkboxes will be added dynamically
        self.run_tab_channel_frame = ttk.Frame(self.control_bar)
        self.run_tab_channel_frame.pack(side=tk.LEFT)
        
        # ===== Plot Customization Bar - Row 2 =====
        customization_frame = ttk.LabelFrame(self.run_tab, text="Plot Customization", padding="5")
        customization_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # Create two rows for controls
        row1_frame = ttk.Frame(customization_frame)
        row1_frame.pack(fill=tk.X, pady=2)
        
        row2_frame = ttk.Frame(customization_frame)
        row2_frame.pack(fill=tk.X, pady=2)
        
        # Row 1: Titles
        # Plot Title
        ttk.Label(row1_frame, text="Plot Title:", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=5)
        self.plot_title_var = tk.StringVar(value="Thermocouple Temperature vs Time")
        ttk.Entry(row1_frame, textvariable=self.plot_title_var, width=30).pack(side=tk.LEFT, padx=5)
        
        ttk.Separator(row1_frame, orient='vertical').pack(side=tk.LEFT, fill=tk.Y, padx=10)
        
        # Left Y-Axis Title
        ttk.Label(row1_frame, text="Left Y-Axis:", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=5)
        self.left_yaxis_title_var = tk.StringVar(value="Temperature (°C)")
        ttk.Entry(row1_frame, textvariable=self.left_yaxis_title_var, width=20).pack(side=tk.LEFT, padx=5)
        
        ttk.Separator(row1_frame, orient='vertical').pack(side=tk.LEFT, fill=tk.Y, padx=10)
        
        # Right Y-Axis Title
        ttk.Label(row1_frame, text="Right Y-Axis:", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=5)
        self.right_yaxis_title_var = tk.StringVar(value="Temperature (°C)")
        ttk.Entry(row1_frame, textvariable=self.right_yaxis_title_var, width=20).pack(side=tk.LEFT, padx=5)
        
        # Row 2: Axis Ranges
        # X-Axis Range (Time)
        ttk.Label(row2_frame, text="X-Axis (Auto):", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=5)
        self.x_axis_auto_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(row2_frame, text="Auto", variable=self.x_axis_auto_var).pack(side=tk.LEFT, padx=2)
        
        ttk.Separator(row2_frame, orient='vertical').pack(side=tk.LEFT, fill=tk.Y, padx=10)
        
        # Left Y-Axis Range
        ttk.Label(row2_frame, text="Left Y-Axis:", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=5)
        self.left_y_auto_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(row2_frame, text="Auto", variable=self.left_y_auto_var).pack(side=tk.LEFT, padx=2)
        
        ttk.Label(row2_frame, text="Min:").pack(side=tk.LEFT, padx=(10, 2))
        self.left_y_min_var = tk.StringVar(value="0")
        ttk.Entry(row2_frame, textvariable=self.left_y_min_var, width=8).pack(side=tk.LEFT, padx=2)
        
        ttk.Label(row2_frame, text="Max:").pack(side=tk.LEFT, padx=(5, 2))
        self.left_y_max_var = tk.StringVar(value="100")
        ttk.Entry(row2_frame, textvariable=self.left_y_max_var, width=8).pack(side=tk.LEFT, padx=2)
        
        ttk.Separator(row2_frame, orient='vertical').pack(side=tk.LEFT, fill=tk.Y, padx=10)
        
        # Right Y-Axis Range
        ttk.Label(row2_frame, text="Right Y-Axis:", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=5)
        self.right_y_auto_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(row2_frame, text="Auto", variable=self.right_y_auto_var).pack(side=tk.LEFT, padx=2)
        
        ttk.Label(row2_frame, text="Min:").pack(side=tk.LEFT, padx=(10, 2))
        self.right_y_min_var = tk.StringVar(value="0")
        ttk.Entry(row2_frame, textvariable=self.right_y_min_var, width=8).pack(side=tk.LEFT, padx=2)
        
        ttk.Label(row2_frame, text="Max:").pack(side=tk.LEFT, padx=(5, 2))
        self.right_y_max_var = tk.StringVar(value="100")
        ttk.Entry(row2_frame, textvariable=self.right_y_max_var, width=8).pack(side=tk.LEFT, padx=2)
        
        ttk.Separator(row2_frame, orient='vertical').pack(side=tk.LEFT, fill=tk.Y, padx=10)
        
        # Apply button
        ttk.Button(row2_frame, text="Apply to Plot", command=self.apply_plot_settings, 
                   width=15).pack(side=tk.LEFT, padx=10)
        
        # ===== Main Plot =====
        plot_frame = ttk.Frame(self.run_tab)
        plot_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        plot_frame.columnconfigure(0, weight=1)
        plot_frame.rowconfigure(0, weight=1)
        
        self.figure = Figure(figsize=(12, 7), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.ax.set_xlabel("Time", fontsize=12)
        self.ax.set_ylabel("Temperature (°C)", fontsize=12)
        self.ax.set_title("Thermocouple Temperature vs Time", fontsize=14, fontweight='bold')
        self.ax.grid(True, alpha=0.3)
        
        self.canvas = FigureCanvasTkAgg(self.figure, master=plot_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Add matplotlib toolbar for zoom/pan
        toolbar = NavigationToolbar2Tk(self.canvas, plot_frame)
        toolbar.update()
        
        self.plot_lines = []
    
    def rebuild_run_tab_controls(self):
        """Rebuild Run tab channel selection controls with custom labels"""
        # Clear existing checkboxes
        for widget in self.run_tab_channel_frame.winfo_children():
            widget.destroy()
        
        # Create new checkboxes with channel labels
        for i in range(len(self.channel_selection_vars)):
            channel_label = self.channel_label_vars[i].get()
            cb = ttk.Checkbutton(self.run_tab_channel_frame, 
                               text=channel_label,  # Use custom label instead of "Ch{i}"
                               variable=self.channel_selection_vars[i],
                               command=self.update_run_tab_labels)  # Update when changed
            cb.pack(side=tk.LEFT, padx=2)
    
    def update_run_tab_labels(self):
        """Update Run tab channel checkbox labels when channel names change"""
        # This will be called when checkboxes need to update their labels
        self.rebuild_run_tab_controls()
    
    def get_plot_style(self, channel_index):
        """Get linestyle and marker for a channel"""
        style_str = self.channel_style_vars[channel_index].get()
        
        style_map = {
            'Line': ('-', ''),
            'Dashed': ('--', ''),
            'Dotted': (':', ''),
            'Dash-Dot': ('-.', ''),
            'Scatter': ('', self.markers[channel_index % len(self.markers)]),
            'Line+Scatter': ('-', self.markers[channel_index % len(self.markers)])
        }
        
        return style_map.get(style_str, ('-', ''))
    
    def get_tc_type(self):
        """Convert thermocouple type string to nidaqmx constant"""
        tc_map = {
            'B': ThermocoupleType.B,
            'E': ThermocoupleType.E,
            'J': ThermocoupleType.J,
            'K': ThermocoupleType.K,
            'N': ThermocoupleType.N,
            'R': ThermocoupleType.R,
            'S': ThermocoupleType.S,
            'T': ThermocoupleType.T
        }
        return tc_map.get(self.tc_type_var.get(), ThermocoupleType.K)
    
    def get_temp_units(self):
        """Convert temperature units string to nidaqmx constant"""
        units_map = {
            'Celsius': TemperatureUnits.DEG_C,
            'Fahrenheit': TemperatureUnits.DEG_F,
            'Kelvin': TemperatureUnits.DEG_C  # Use Celsius internally, we'll convert if needed
        }
        return units_map.get(self.temp_units_var.get(), TemperatureUnits.DEG_C)
    
    def get_temp_unit_symbol(self):
        """Get temperature unit symbol for display"""
        symbols = {
            'Celsius': '°C',
            'Fahrenheit': '°F',
            'Kelvin': 'K'
        }
        return symbols.get(self.temp_units_var.get(), '°C')
        
    def update_acquisition_rate(self, event=None):
        """Update acquisition rate and file rotation interval"""
        rate_str = self.acq_rate_var.get()
        rate_value = float(rate_str.split()[0])
        self.acquisition_rate = rate_value
        
        # Update file rotation interval
        if rate_value <= 1.0:
            self.file_rotation_interval = 12 * 3600  # 12 hours
        else:
            self.file_rotation_interval = 3 * 3600   # 3 hours
    
    def get_time_window_seconds(self):
        """Convert time window string to seconds"""
        window_str = self.time_window_var.get()
        if "minute" in window_str:
            return int(window_str.split()[0]) * 60
        elif "hour" in window_str:
            return int(window_str.split()[0]) * 3600
        elif "day" in window_str:
            return int(window_str.split()[0]) * 86400
        return 3600  # default 1 hour
    
    def generate_filename(self):
        """Generate filename with timestamp and experiment name"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        exp_name = self.experiment_name_var.get().replace(" ", "_")
        base_path = self.data_path_var.get()
        
        # Create full path
        full_path = os.path.join(base_path, f"{exp_name}_{timestamp}")
        return full_path
    
    def create_new_files(self):
        """Create new TDMS and optionally CSV files"""
        # Close existing files
        self.close_files()
        
        # Generate filename
        base_filename = self.generate_filename()
        
        # Ensure directory exists
        os.makedirs(os.path.dirname(base_filename) if os.path.dirname(base_filename) else '.', exist_ok=True)
        
        # Create TDMS file
        self.tdms_filepath = f"{base_filename}.tdms"
        # Open TDMS file (will be written to in write_tdms_data)
        self.tdms_file = open(self.tdms_filepath, 'wb')
        self.tdms_writer = TdmsWriter(self.tdms_file)
        
        # Create CSV file if enabled
        if self.csv_logging_var.get():
            self.csv_filepath = f"{base_filename}.csv"
            self.csv_file = open(self.csv_filepath, 'w', newline='')
            self.csv_writer = csv.writer(self.csv_file)
            # Write header with custom labels
            header = ["Timestamp"] + [self.channel_label_vars[i].get() for i in range(self.num_channels)]
            self.csv_writer.writerow(header)
        
        self.current_file_start_time = time.time()
        self.log_status(f"New file created: {os.path.basename(base_filename)}")
    
    def close_files(self):
        """Close open data files"""
        if self.tdms_writer is not None:
            try:
                # Flush any remaining buffered data
                if hasattr(self, 'tdms_buffer') and len(self.tdms_buffer) > 0:
                    self.log_status(f"Flushing {len(self.tdms_buffer)} remaining samples to TDMS")
                    
                    root_object = RootObject()
                    group_object = GroupObject("Thermocouple_Data")
                    
                    # Initialize data collectors
                    channels_data = {i: [] for i in range(self.num_channels)}
                    timestamps_data = []
                    
                    # Collect buffered data
                    for ts, temps in self.tdms_buffer:
                        for i, temp in enumerate(temps):
                            if i < self.num_channels:
                                channels_data[i].append(float(temp))  # Ensure float
                        
                        # Convert timestamp to Excel date format
                        excel_epoch = datetime(1900, 1, 1)
                        delta = ts - excel_epoch
                        excel_timestamp = delta.total_seconds() / 86400 + 2
                        timestamps_data.append(excel_timestamp)
                    
                    # Create channel objects
                    channels = []
                    for i in range(self.num_channels):
                        channel_name = self.channel_label_vars[i].get()
                        channel_array = np.array(channels_data[i], dtype=np.float64)
                        
                        channel = ChannelObject("Thermocouple_Data", channel_name, 
                                               channel_array, 
                                               properties={
                                                   "Unit": self.temp_units_var.get(),
                                                   "Type": f"{self.tc_type_var.get()}-Type"
                                               })
                        channels.append(channel)
                    
                    # Add timestamp channel
                    timestamp_array = np.array(timestamps_data, dtype=np.float64)
                    time_channel = ChannelObject("Thermocouple_Data", "Timestamp", 
                                                timestamp_array,
                                                properties={"Format": "Excel Date/Time"})
                    channels.append(time_channel)
                    
                    # Write final segment
                    self.tdms_writer.write_segment([root_object, group_object] + channels)
                    
                    # Clear buffer
                    self.tdms_buffer = []
                
                self.tdms_writer.close()
            except Exception as e:
                print(f"Error closing TDMS: {e}")
                self.log_status(f"Error closing TDMS: {e}")
            
            self.tdms_writer = None
            if hasattr(self, 'tdms_file') and self.tdms_file is not None:
                try:
                    self.tdms_file.close()
                except:
                    pass
                self.tdms_file = None
            self.log_status("TDMS file closed")
        
        if self.csv_file is not None:
            try:
                self.csv_file.close()
            except:
                pass
            self.csv_file = None
            self.csv_writer = None
            self.log_status("CSV file closed")
    
    def write_tdms_data(self, timestamp, data):
        """Write data to TDMS file with buffering"""
        if self.tdms_writer is None:
            return
        
        try:
            # Ensure data is in the correct format (list/array of values, one per channel)
            if isinstance(data, np.ndarray):
                data = data.flatten().tolist()
            elif not isinstance(data, (list, tuple)):
                data = [data]
            else:
                data = list(data)  # Ensure it's a regular list, not tuple
            
            # Convert all values to float
            data = [float(val) for val in data]
            
            # Add to buffer
            self.tdms_buffer.append((timestamp, data))
            
            # Write when buffer is full
            if len(self.tdms_buffer) >= self.tdms_buffer_size:
                root_object = RootObject()
                group_object = GroupObject("Thermocouple_Data")
                
                # Initialize data collectors
                channels_data = {i: [] for i in range(self.num_channels)}
                timestamps_data = []
                
                # Collect buffered data
                for ts, temps in self.tdms_buffer:
                    for i, temp in enumerate(temps):
                        if i < self.num_channels:
                            channels_data[i].append(float(temp))  # Ensure float
                    
                    # Convert timestamp to Excel date format
                    excel_epoch = datetime(1900, 1, 1)
                    delta = ts - excel_epoch
                    excel_timestamp = delta.total_seconds() / 86400 + 2
                    timestamps_data.append(excel_timestamp)
                
                # Create channel objects with 1D numpy arrays
                channels = []
                for i in range(self.num_channels):
                    channel_name = self.channel_label_vars[i].get()
                    # Convert to 1D numpy array of float64
                    channel_array = np.array(channels_data[i], dtype=np.float64)
                    
                    channel = ChannelObject("Thermocouple_Data", channel_name, 
                                           channel_array, 
                                           properties={
                                               "Unit": self.temp_units_var.get(),
                                               "Type": f"{self.tc_type_var.get()}-Type"
                                           })
                    channels.append(channel)
                
                # Add timestamp channel as 1D numpy array
                timestamp_array = np.array(timestamps_data, dtype=np.float64)
                time_channel = ChannelObject("Thermocouple_Data", "Timestamp", 
                                            timestamp_array,
                                            properties={"Format": "Excel Date/Time"})
                channels.append(time_channel)
                
                # Write segment
                self.tdms_writer.write_segment([root_object, group_object] + channels)
                
                # Clear buffer
                self.tdms_buffer = []
            
        except Exception as e:
            self.log_status(f"TDMS Write Error: {e}")
            print(f"TDMS Write Error: {e}")
            import traceback
            traceback.print_exc()
    
    def check_file_rotation(self):
        """Check if it's time to rotate to a new file"""
        if self.current_file_start_time is None:
            return False
        
        elapsed = time.time() - self.current_file_start_time
        return elapsed >= self.file_rotation_interval
    
    def start_acquisition(self):
        """Start data acquisition"""
        try:
            # Check if DAQ configuration has been applied
            if not hasattr(self, 'channels') or len(self.channels) == 0:
                messagebox.showerror("Error", "Please apply DAQ configuration before starting acquisition!")
                return
            
            # Update acquisition rate
            self.update_acquisition_rate()
            
            # Clear previous data
            with self.data_lock:
                self.timestamps = []
                self.temperature_data = [[] for _ in range(self.num_channels)]
            
            # Create initial files
            self.create_new_files()
            
            # Update UI
            self.start_button.config(state="disabled")
            self.stop_button.config(state="normal")
            self.acquisition_running = True
            
            # Disable configuration widgets
            self.disable_config_widgets()
            
            self.log_status("=== Acquisition Started ===")
            self.log_status(f"Rate: {self.acq_rate_var.get()}")
            self.log_status(f"Device: {self.device_name}")
            
            # Start acquisition thread
            self.acq_thread = threading.Thread(target=self.acquisition_loop, daemon=True)
            self.acq_thread.start()
            
            # Start plot update timer
            self.update_plot()
            
        except Exception as e:
            self.log_status(f"ERROR starting acquisition: {str(e)}")
            messagebox.showerror("Error", f"Failed to start acquisition:\n{str(e)}")
            self.stop_acquisition()
    
    def stop_acquisition(self):
        """Stop data acquisition"""
        self.acquisition_running = False
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        
        # Re-enable configuration widgets
        self.enable_config_widgets()
        
        self.log_status("=== Acquisition Stopped ===")
        
        # Close files
        self.close_files()
        
        # Close files
        self.close_files()
    
    def acquisition_loop(self):
        """Main acquisition loop running in separate thread with hardware timing"""
        try:
            with nidaqmx.Task() as task:
                # Add thermocouple channels
                tc_type = self.get_tc_type()
                temp_units = self.get_temp_units()
                
                for channel in self.channels:
                    task.ai_channels.add_ai_thrmcpl_chan(
                        channel,
                        thermocouple_type=tc_type,
                        units=temp_units
                    )
                
                # Configure timing for continuous acquisition
                # Note: Thermocouple modules have hardware limitations
                # NI 9211 max aggregate rate is ~14 S/s
                actual_rate = min(self.acquisition_rate, 10)  # Cap at 10 Hz for safety
                
                if self.acquisition_rate > actual_rate:
                    self.root.after(0, lambda: self.log_status(
                        f"WARNING: Rate limited to {actual_rate} Hz due to hardware constraints"
                    ))
                
                try:
                    # Try hardware-timed acquisition
                    task.timing.cfg_samp_clk_timing(
                        rate=actual_rate,
                        sample_mode=nidaqmx.constants.AcquisitionType.CONTINUOUS,
                        samps_per_chan=1000
                    )
                    hardware_timed = True
                    self.log_status(f"Using hardware-timed acquisition at {actual_rate} Hz")
                except:
                    # Fall back to software timing
                    hardware_timed = False
                    self.log_status(f"Using software-timed acquisition at {actual_rate} Hz")
                
                sample_count = 0
                start_time = time.time()
                last_log_time = start_time
                
                if hardware_timed:
                    # Hardware-timed acquisition
                    task.start()
                    
                    while self.acquisition_running:
                        try:
                            # Read available samples (non-blocking with timeout)
                            data = task.read(number_of_samples_per_channel=1, timeout=2.0)
                            current_time = datetime.now()
                            
                            # Handle data format - hardware-timed read returns different structure
                            # For multiple channels: [[ch0_val], [ch1_val], [ch2_val], [ch3_val]]
                            # We need: [ch0_val, ch1_val, ch2_val, ch3_val]
                            if isinstance(data, list) and len(data) > 0:
                                if isinstance(data[0], list):
                                    # Multi-channel: extract first element from each channel
                                    data = [ch[0] if isinstance(ch, list) and len(ch) > 0 else ch for ch in data]
                                elif self.num_channels == 1:
                                    # Single channel with single sample
                                    data = [data[0]] if isinstance(data, list) else [data]
                            
                            # Check for file rotation
                            if self.check_file_rotation():
                                self.create_new_files()
                                self.root.after(0, lambda: self.log_status("File rotated - new file created"))
                            
                            # Store data
                            with self.data_lock:
                                self.timestamps.append(current_time)
                                for i, temp in enumerate(data):
                                    if i < self.num_channels:
                                        self.temperature_data[i].append(temp)
                            
                            # Write to TDMS
                            self.write_tdms_data(current_time, data)
                            
                            # Write to CSV if enabled
                            if self.csv_logging_var.get() and self.csv_writer is not None:
                                # Ensure data is properly formatted as individual float values
                                if isinstance(data, (list, tuple)):
                                    temp_values = [float(val) for val in data]
                                elif isinstance(data, np.ndarray):
                                    temp_values = data.flatten().tolist()
                                else:
                                    temp_values = [float(data)]
                                
                                # Create row with timestamp and individual temperature values
                                row = [current_time.strftime("%Y-%m-%d %H:%M:%S.%f")] + temp_values
                                self.csv_writer.writerow(row)
                                if sample_count % 10 == 0:
                                    self.csv_file.flush()
                            
                            # Update display periodically
                            current_time_check = time.time()
                            if current_time_check - self.last_display_update >= self.display_update_interval:
                                self.root.after(0, self.update_display, data)
                                self.last_display_update = current_time_check
                            
                            sample_count += 1
                            
                            # Log actual rate every 10 seconds
                            if current_time_check - last_log_time >= 10.0:
                                elapsed = current_time_check - start_time
                                actual_rate_measured = sample_count / elapsed
                                self.root.after(0, lambda r=actual_rate_measured: self.log_status(
                                    f"Actual rate: {r:.3f} Hz | Total samples: {sample_count}"
                                ))
                                last_log_time = current_time_check
                            
                        except nidaqmx.DaqError as e:
                            if "timeout" not in str(e).lower():
                                self.root.after(0, lambda e=e: self.log_status(f"DAQ Error: {str(e)}"))
                        except Exception as e:
                            self.root.after(0, lambda e=e: self.log_status(f"Error in acquisition: {str(e)}"))
                            import traceback
                            traceback.print_exc()
                            
                else:
                    # Software-timed acquisition (original method)
                    interval = 1.0 / actual_rate
                    
                    while self.acquisition_running:
                        loop_start = time.time()
                        
                        try:
                            # Check for file rotation
                            if self.check_file_rotation():
                                self.create_new_files()
                                self.root.after(0, lambda: self.log_status("File rotated - new file created"))
                            
                            # Read data
                            data = task.read()
                            current_time = datetime.now()
                            
                            # Store data
                            with self.data_lock:
                                self.timestamps.append(current_time)
                                for i, temp in enumerate(data):
                                    self.temperature_data[i].append(temp)
                            
                            # Write to TDMS
                            self.write_tdms_data(current_time, data)
                            
                            # Write to CSV if enabled
                            if self.csv_logging_var.get() and self.csv_writer is not None:
                                row = [current_time.strftime("%Y-%m-%d %H:%M:%S.%f")] + list(data)
                                self.csv_writer.writerow(row)
                                if sample_count % 10 == 0:
                                    self.csv_file.flush()
                            
                            # Update display periodically
                            current_time_check = time.time()
                            if current_time_check - self.last_display_update >= self.display_update_interval:
                                self.root.after(0, self.update_display, data)
                                self.last_display_update = current_time_check
                            
                            sample_count += 1
                            
                            # Log actual rate every 10 seconds
                            if current_time_check - last_log_time >= 10.0:
                                elapsed = current_time_check - start_time
                                actual_rate_measured = sample_count / elapsed
                                self.root.after(0, lambda r=actual_rate_measured: self.log_status(
                                    f"Actual rate: {r:.3f} Hz | Total samples: {sample_count}"
                                ))
                                last_log_time = current_time_check
                            
                            # Sleep until next sample
                            loop_duration = time.time() - loop_start
                            sleep_time = interval - loop_duration
                            
                            if sleep_time > 0.001:
                                time.sleep(sleep_time)
                                
                        except nidaqmx.DaqError as e:
                            print(f"DAQ Error: {e}")
                            self.root.after(0, lambda e=e: self.log_status(f"DAQ Error: {str(e)}"))
                            time.sleep(interval)
                        
        except Exception as e:
            self.root.after(0, lambda e=e: self.log_status(f"Acquisition Error: {str(e)}"))
            self.root.after(0, lambda e=e: messagebox.showerror("Acquisition Error", str(e)))
            self.root.after(0, self.stop_acquisition)
    
    def update_display(self, data):
        """Update temperature display labels"""
        unit_symbol = self.get_temp_unit_symbol()
        
        for i, temp in enumerate(data):
            self.temp_labels[i].config(text=f"{temp:.2f} {unit_symbol}")
        
        # Update statistics
        with self.data_lock:
            for i in range(self.num_channels):
                if len(self.temperature_data[i]) > 0:
                    temps = self.temperature_data[i]
                    min_temp = min(temps)
                    max_temp = max(temps)
                    avg_temp = np.mean(temps)
                    self.stats_labels[i].config(
                        text=f"Min: {min_temp:.2f} | Max: {max_temp:.2f} | Avg: {avg_temp:.2f}"
                    )
    
    def update_plot(self):
        """Update plot periodically with dual Y-axis support and custom settings"""
        if not self.acquisition_running:
            return
        
        try:
            with self.data_lock:
                if len(self.timestamps) == 0:
                    self.root.after(1000, self.update_plot)
                    return
                
                # Get time window in seconds
                window_seconds = self.get_time_window_seconds()
                current_time = datetime.now()
                start_time = current_time - timedelta(seconds=window_seconds)
                
                # Filter data within time window
                indices = [i for i, t in enumerate(self.timestamps) if t >= start_time]
                
                if len(indices) == 0:
                    self.root.after(1000, self.update_plot)
                    return
                
                plot_times = [self.timestamps[i] for i in indices]
                plot_data = [[self.temperature_data[ch][i] for i in indices] 
                            for ch in range(self.num_channels)]
                
                # Decimate data for plotting (max 1000 points)
                if len(plot_times) > 1000:
                    step = len(plot_times) // 1000
                    plot_times = plot_times[::step]
                    plot_data = [data[::step] for data in plot_data]
            
            # Clear the figure
            self.figure.clear()
            
            # Create primary axis
            self.ax = self.figure.add_subplot(111)
            
            # Check if we need a secondary Y-axis
            has_right_axis = any(self.channel_yaxis_vars[i].get() == "Right" 
                                and self.channel_selection_vars[i].get() 
                                for i in range(self.num_channels))
            
            # Create secondary axis if needed
            if has_right_axis:
                ax_right = self.ax.twinx()
            else:
                ax_right = None
            
            # Plot channels on appropriate axes
            left_plotted = False
            right_plotted = False
            
            for i in range(self.num_channels):
                if self.channel_selection_vars[i].get():
                    linestyle, marker = self.get_plot_style(i)
                    color = self.channel_color_vars[i].get()
                    label = self.channel_label_vars[i].get()
                    yaxis_side = self.channel_yaxis_vars[i].get()
                    
                    # Choose which axis to plot on
                    if yaxis_side == "Right" and ax_right is not None:
                        current_ax = ax_right
                        right_plotted = True
                    else:
                        current_ax = self.ax
                        left_plotted = True
                    
                    # Plot the data
                    if marker:
                        markersize = 4 if linestyle else 6
                        current_ax.plot(plot_times, plot_data[i], 
                                   color=color,
                                   linestyle=linestyle if linestyle else 'None',
                                   marker=marker,
                                   markersize=markersize,
                                   label=label, 
                                   linewidth=2)
                    else:
                        current_ax.plot(plot_times, plot_data[i], 
                                   color=color,
                                   linestyle=linestyle,
                                   label=label, 
                                   linewidth=2)
            
            # Apply custom plot title
            plot_title = self.plot_title_var.get()
            self.ax.set_title(plot_title, fontsize=14, fontweight='bold')
            
            # Configure X-axis
            self.ax.set_xlabel("Time", fontsize=12)
            if not self.x_axis_auto_var.get():
                # X-axis is time-based, auto is typically desired
                pass
            
            # Configure left Y-axis
            left_ylabel = self.left_yaxis_title_var.get()
            self.ax.set_ylabel(left_ylabel, fontsize=12, color='black')
            self.ax.tick_params(axis='y', labelcolor='black')
            self.ax.grid(True, alpha=0.3)
            
            # Apply left Y-axis limits if not auto
            if not self.left_y_auto_var.get():
                try:
                    left_min = float(self.left_y_min_var.get())
                    left_max = float(self.left_y_max_var.get())
                    self.ax.set_ylim(left_min, left_max)
                except ValueError:
                    pass  # Use auto if invalid values
            
            # Configure right Y-axis if it exists
            if ax_right is not None and right_plotted:
                right_ylabel = self.right_yaxis_title_var.get()
                ax_right.set_ylabel(right_ylabel, fontsize=12, color='black')
                ax_right.tick_params(axis='y', labelcolor='black')
                
                # Apply right Y-axis limits if not auto
                if not self.right_y_auto_var.get():
                    try:
                        right_min = float(self.right_y_min_var.get())
                        right_max = float(self.right_y_max_var.get())
                        ax_right.set_ylim(right_min, right_max)
                    except ValueError:
                        pass  # Use auto if invalid values
            
            # Combine legends from both axes
            if left_plotted or right_plotted:
                lines_1, labels_1 = self.ax.get_legend_handles_labels()
                if ax_right is not None and right_plotted:
                    lines_2, labels_2 = ax_right.get_legend_handles_labels()
                    lines = lines_1 + lines_2
                    labels = labels_1 + labels_2
                else:
                    lines = lines_1
                    labels = labels_1
                
                self.ax.legend(lines, labels, loc='best', fontsize=10, framealpha=0.9)
            
            self.figure.autofmt_xdate()
            self.canvas.draw()
            
        except Exception as e:
            print(f"Plot update error: {e}")
            import traceback
            traceback.print_exc()
        
        # Schedule next update
        self.root.after(1000, self.update_plot)
        
    def on_closing(self):
        """Handle window closing"""
        if self.acquisition_running:
            if messagebox.askokcancel("Quit", "Acquisition is running. Stop and quit?"):
                self.stop_acquisition()
                self.root.destroy()
        else:
            self.root.destroy()
            
    def disable_config_widgets(self):
        """Disable configuration widgets during acquisition"""
        # Disable all widgets in column 1 (recursively)
        if hasattr(self, 'col1_frame'):
            self._disable_widget_recursive(self.col1_frame)
        
        # Disable channel name entry boxes
        for entry in self.channel_name_entries:
            entry.config(state='disabled')
        
        self.log_status("Configuration locked during acquisition")
    
    def enable_config_widgets(self):
        """Enable configuration widgets after acquisition stops"""
        # Enable all widgets in column 1 (recursively)
        if hasattr(self, 'col1_frame'):
            self._enable_widget_recursive(self.col1_frame)
        
        # Enable channel name entry boxes
        for entry in self.channel_name_entries:
            entry.config(state='normal')
        
        self.log_status("Configuration unlocked")
    
    def _disable_widget_recursive(self, widget):
        """Recursively disable a widget and all its children"""
        try:
            # Try to disable the widget
            if isinstance(widget, (ttk.Entry, ttk.Combobox, ttk.Spinbox)):
                widget.config(state='disabled')
            elif isinstance(widget, (ttk.Button, ttk.Checkbutton)):
                widget.config(state='disabled')
            elif isinstance(widget, tk.Text):
                widget.config(state='disabled')
        except:
            pass
        
        # Recursively disable children
        try:
            for child in widget.winfo_children():
                self._disable_widget_recursive(child)
        except:
            pass
    
    def _enable_widget_recursive(self, widget):
        """Recursively enable a widget and all its children"""
        try:
            # Try to enable the widget
            if isinstance(widget, (ttk.Entry, ttk.Spinbox)):
                widget.config(state='normal')
            elif isinstance(widget, ttk.Combobox):
                widget.config(state='readonly')
            elif isinstance(widget, (ttk.Button, ttk.Checkbutton)):
                widget.config(state='normal')
            elif isinstance(widget, tk.Text):
                widget.config(state='normal')
        except:
            pass
        
        # Recursively enable children
        try:
            for child in widget.winfo_children():
                self._enable_widget_recursive(child)
        except:
            pass

    def apply_plot_settings(self):
        """Apply user-defined plot settings"""
        try:
            # Validate and store the settings
            if not self.left_y_auto_var.get():
                try:
                    left_min = float(self.left_y_min_var.get())
                    left_max = float(self.left_y_max_var.get())
                    if left_min >= left_max:
                        messagebox.showerror("Error", "Left Y-Axis: Min must be less than Max")
                        return
                except ValueError:
                    messagebox.showerror("Error", "Left Y-Axis: Invalid numeric values")
                    return
            
            if not self.right_y_auto_var.get():
                try:
                    right_min = float(self.right_y_min_var.get())
                    right_max = float(self.right_y_max_var.get())
                    if right_min >= right_max:
                        messagebox.showerror("Error", "Right Y-Axis: Min must be less than Max")
                        return
                except ValueError:
                    messagebox.showerror("Error", "Right Y-Axis: Invalid numeric values")
                    return
            
            self.log_status("Plot settings applied")
            
            # Force immediate plot update if acquisition is running
            if self.acquisition_running:
                self.root.after(100, self.force_plot_update)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to apply plot settings:\n{str(e)}")
            self.log_status(f"Error applying plot settings: {str(e)}")
    
    def force_plot_update(self):
        """Force an immediate plot update with current data"""
        if not self.acquisition_running:
            return
        
        # Just trigger update_plot immediately
        self.update_plot()

    def on_channel_label_change(self, channel_index):
        """Called when a channel label is changed"""
        # Rebuild Run tab controls to reflect new label
        if hasattr(self, 'run_tab_channel_frame'):
            self.root.after(100, self.rebuild_run_tab_controls)
        
        # Update temperature display labels
        self.root.after(100, self.update_temp_display_labels)

    def update_temp_display_labels(self):
        """Update temperature display labels to match channel names"""
        if hasattr(self, 'temp_display_labels') and hasattr(self, 'channel_label_vars'):
            for i in range(min(len(self.temp_display_labels), len(self.channel_label_vars))):
                label_text = self.channel_label_vars[i].get()
                self.temp_display_labels[i].config(text=f"{label_text}:")

def main():
    root = tk.Tk()
    app = ThermocopleDAQGUI(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()
