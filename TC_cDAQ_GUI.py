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
from dataclasses import dataclass
import numpy as np

@dataclass
class RingBufferSnapshot:
    """A consistent view of ring-buffer contents (already time-ordered)."""
    times: np.ndarray          # shape: (n,)
    data: np.ndarray           # shape: (channels, n)
    count: int                 # n
    capacity: int              # total capacity


class MultiChannelRingBuffer:
    """
    Fixed-size ring buffer for time series:
    - times stored as float seconds since epoch (time.time()) for efficiency
    - data stored as (channels, capacity) float32/float64
    """
    def __init__(self, channels: int, capacity: int, dtype=np.float32):
        if channels <= 0:
            raise ValueError("channels must be > 0")
        if capacity <= 1:
            raise ValueError("capacity must be > 1")

        self.channels = int(channels)
        self.capacity = int(capacity)
        self.dtype = dtype

        self._times = np.empty((self.capacity,), dtype=np.float64)
        self._data = np.empty((self.channels, self.capacity), dtype=self.dtype)

        self._write = 0          # next write index
        self._count = 0          # number of valid samples (<= capacity)

    def clear(self):
        self._write = 0
        self._count = 0

    @property
    def count(self) -> int:
        return self._count

    def append(self, t_sec: float, values):
        """
        Append one multi-channel sample.
        values must be length == channels.
        """
        # Convert once; keep this lightweight (called at acquisition rate)
        v = np.asarray(values, dtype=self.dtype)
        if v.shape[0] != self.channels:
            raise ValueError(f"Expected {self.channels} values, got {v.shape[0]}")

        idx = self._write
        self._times[idx] = float(t_sec)
        self._data[:, idx] = v

        self._write = (self._write + 1) % self.capacity
        self._count = min(self._count + 1, self.capacity)

    def snapshot_last(self, n: int) -> RingBufferSnapshot:
        """
        Return the last n samples (time-ordered).
        If n > count, returns all available.
        """
        if self._count == 0:
            return RingBufferSnapshot(
                times=np.empty((0,), dtype=np.float64),
                data=np.empty((self.channels, 0), dtype=self.dtype),
                count=0,
                capacity=self.capacity,
            )

        n = int(max(0, n))
        n = min(n, self._count)

        end = self._write
        start = (end - n) % self.capacity

        if start < end:
            times = self._times[start:end]
            data = self._data[:, start:end]
        else:
            # wrapped
            times = np.concatenate((self._times[start:], self._times[:end]), axis=0)
            data = np.concatenate((self._data[:, start:], self._data[:, :end]), axis=1)

        # return copies to prevent mutation issues if producer continues writing
        return RingBufferSnapshot(times=times.copy(), data=data.copy(), count=n, capacity=self.capacity)

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
        
        # Acquisition parameters
        self.acquisition_rate = 1.0  # Hz
        self.file_rotation_interval = 12 * 3600  # seconds
        
        # Plot styling
        self.colors = ['red', 'blue', 'green', 'orange', 'purple', 'brown', 'pink', 'gray', 'cyan', 'magenta']
        self.markers = ['o', 's', '^', 'D', 'v', '<', '>', 'p', '*', 'h']
        
        # Multiple module support
        self.modules = []  # List of module configurations
        self.total_channels = 0
        
        # Channel configuration lists (will be populated based on total channels across all modules)
        self.channel_selection_vars = []
        self.channel_label_vars = []
        self.channel_style_vars = []
        self.channel_color_vars = []
        self.channel_yaxis_vars = []
        self.channel_frames = []
        self.temp_display_labels = []
        self.channel_name_entries = []
        
        # TDMS buffering
        self.tdms_buffer = []
        self.tdms_buffer_size = 10
        
        # Display update throttling
        self.last_display_update = time.time()
        self.display_update_interval = 0.5
        
        # Plot zoom state tracking - NEW
        self.plot_xlim = None
        self.plot_ylim_left = None
        self.plot_ylim_right = None
        self.user_zoomed = False
        self._updating_plot = False
        
        # Initialize GUI variables (MUST be before create_widgets)
        self.time_window_var = tk.StringVar(value="1 hour")
        self.tc_type_var = tk.StringVar(value="K")
        self.temp_units_var = tk.StringVar(value="Celsius")
        self.experiment_name_var = tk.StringVar(value="TC_Experiment")
        self.data_path_var = tk.StringVar(value=os.getcwd())
        self.config_path_var = tk.StringVar(value=os.getcwd())
        self.acq_rate_var = tk.StringVar(value="1 Hz")
        self.csv_logging_var = tk.BooleanVar(value=False)
        
        # Plot customization variables
        self.plot_title_var = tk.StringVar(value="Thermocouple Temperature vs Time")
        self.left_yaxis_title_var = tk.StringVar(value="Temperature (°C)")
        self.right_yaxis_title_var = tk.StringVar(value="Temperature (°C)")
        self.x_axis_auto_var = tk.BooleanVar(value=True)
        self.left_y_auto_var = tk.BooleanVar(value=True)
        self.left_y_min_var = tk.StringVar(value="0")
        self.left_y_max_var = tk.StringVar(value="100")
        self.right_y_auto_var = tk.BooleanVar(value=True)
        self.right_y_min_var = tk.StringVar(value="0")
        self.right_y_max_var = tk.StringVar(value="100")
        
        # --- Run-tab plot refresh controls (NEW) ---
        self.plot_refresh_interval_var = tk.StringVar(value="2 sec")  # default
        self.plot_info_var = tk.StringVar(value="Plot: --/-- pts | stride=-- | eff.rate=-- Hz")
        
        self.plot_needs_rebuild = False
        
        # Create GUI
        self.create_widgets()
        
        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # --- Plot/window performance settings (NEW) ---
        self.max_plot_window_seconds = 7 * 24 * 3600  # 7 days
        self.max_plot_points_per_channel = 1000
        
        # ring buffer will be created after modules are applied (because channel count is not known yet)
        self.ring = None  # type: MultiChannelRingBuffer | None
        
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
        
        # === Module Configuration (NEW - Scrollable) ===
        module_config_frame = ttk.LabelFrame(col1_frame, text="Module Configuration", padding="10")
        module_config_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Add/Apply buttons at top
        module_button_frame = ttk.Frame(module_config_frame)
        module_button_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(module_button_frame, text="➕ Add Module", 
                  command=lambda: self.add_module_config(log=True), width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(module_button_frame, text="✓ Apply Configuration", 
                  command=self.apply_all_modules, width=20).pack(side=tk.LEFT, padx=5)
        
        # Scrollable frame for modules
        module_canvas = tk.Canvas(module_config_frame, height=250)
        module_scrollbar = ttk.Scrollbar(module_config_frame, orient="vertical", command=module_canvas.yview)
        self.modules_container = ttk.Frame(module_canvas)
        
        self.modules_container.bind(
            "<Configure>",
            lambda e: module_canvas.configure(scrollregion=module_canvas.bbox("all"))
        )
        
        module_canvas.create_window((0, 0), window=self.modules_container, anchor="nw")
        module_canvas.configure(yscrollcommand=module_scrollbar.set)
        
        module_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        module_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Add first module by default (without logging)
        self.add_module_config(log=False)
        
        # === DAQ Settings (Thermocouple Type & Temperature Units) ===
        daq_settings_frame = ttk.LabelFrame(col1_frame, text="DAQ Settings", padding="10")
        daq_settings_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Thermocouple Type
        ttk.Label(daq_settings_frame, text="Thermocouple Type:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.tc_type_var = tk.StringVar(value="K")
        tc_type_combo = ttk.Combobox(daq_settings_frame, textvariable=self.tc_type_var,
                                     values=['J', 'K', 'N', 'R', 'S', 'T', 'B', 'E'],
                                     state="readonly", width=22)
        tc_type_combo.grid(row=0, column=1, pady=5, padx=5)
        
        # Temperature Units
        ttk.Label(daq_settings_frame, text="Temperature Units:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.temp_units_var = tk.StringVar(value="Celsius")
        temp_units_combo = ttk.Combobox(daq_settings_frame, textvariable=self.temp_units_var,
                                        values=['Celsius', 'Fahrenheit'],
                                        state="readonly", width=22)
        temp_units_combo.grid(row=1, column=1, pady=5, padx=5)
        
        # === Experiment Configuration ===
        exp_config_frame = ttk.LabelFrame(col1_frame, text="Experiment Configuration", padding="10")
        exp_config_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Experiment Name
        ttk.Label(exp_config_frame, text="Experiment Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.experiment_name_var = tk.StringVar(value="TC_Experiment")
        ttk.Entry(exp_config_frame, textvariable=self.experiment_name_var, width=24).grid(row=0, column=1, pady=5, padx=5)
        
        # Data Save Path
        ttk.Label(exp_config_frame, text="Data Path:").grid(row=1, column=0, sticky=tk.W, pady=5)
        path_frame = ttk.Frame(exp_config_frame)
        path_frame.grid(row=1, column=1, pady=5, padx=5)
        
        self.data_path_var = tk.StringVar(value=os.getcwd())
        ttk.Entry(path_frame, textvariable=self.data_path_var, width=16).pack(side=tk.LEFT)
        ttk.Button(path_frame, text="Browse", command=self.browse_data_path, width=7).pack(side=tk.LEFT, padx=(5, 0))
        
        # Acquisition Rate
        ttk.Label(exp_config_frame, text="Acquisition Rate:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.acq_rate_var = tk.StringVar(value="1 Hz")
        acq_rate_combo = ttk.Combobox(exp_config_frame, textvariable=self.acq_rate_var,
                                      values=["0.1 Hz", "0.5 Hz", "1 Hz", "2 Hz", "5 Hz", "10 Hz"],
                                      state="readonly", width=22)
        acq_rate_combo.grid(row=2, column=1, pady=5, padx=5)
        acq_rate_combo.bind("<<ComboboxSelected>>", self.update_acquisition_rate)
        
        # CSV Logging Option
        self.csv_logging_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(exp_config_frame, text="Enable CSV Logging", 
                       variable=self.csv_logging_var).grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # === File Management ===
        file_mgmt_frame = ttk.LabelFrame(col1_frame, text="Configuration File Management", padding="10")
        file_mgmt_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Config Path
        ttk.Label(file_mgmt_frame, text="Config Path:").grid(row=0, column=0, sticky=tk.W, pady=5)
        config_path_frame = ttk.Frame(file_mgmt_frame)
        config_path_frame.grid(row=0, column=1, pady=5, padx=5)
        
        self.config_path_var = tk.StringVar(value=os.getcwd())
        ttk.Entry(config_path_frame, textvariable=self.config_path_var, width=16).pack(side=tk.LEFT)
        ttk.Button(config_path_frame, text="Browse", command=self.browse_config_path, width=7).pack(side=tk.LEFT, padx=(5, 0))
        
        # Save/Load Buttons
        button_frame = ttk.Frame(file_mgmt_frame)
        button_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Button(button_frame, text="Save Configuration", 
                  command=self.save_configuration, width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Load Configuration", 
                  command=self.load_configuration, width=20).pack(side=tk.LEFT, padx=5)
        
        # ===== COLUMN 2: Channel Configuration =====
        col2_frame = ttk.Frame(self.setup_tab)
        col2_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5)
        
        channels_frame = ttk.LabelFrame(col2_frame, text="Channel Configuration", padding="10")
        channels_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create scrollable frame for channels
        canvas = tk.Canvas(channels_frame)
        scrollbar = ttk.Scrollbar(channels_frame, orient="vertical", command=canvas.yview)
        self.channels_container = ttk.Frame(canvas)
        
        self.channels_container.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.channels_container, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ===== COLUMN 3: Control Buttons and Status Log =====
        col3_frame = ttk.Frame(self.setup_tab)
        col3_frame.grid(row=0, column=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5)
        
        # Control Buttons
        control_frame = ttk.LabelFrame(col3_frame, text="Control", padding="10")
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.start_button = ttk.Button(control_frame, text="▶ Start Acquisition", 
                                       command=self.start_acquisition, width=25)
        self.start_button.pack(pady=5)
        
        self.stop_button = ttk.Button(control_frame, text="⏹ Stop Acquisition", 
                                      command=self.stop_acquisition, width=25, state="disabled")
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
        
        # ===== COLUMN 4: Temperature Display and Statistics =====
        col4_frame = ttk.Frame(self.setup_tab)
        col4_frame.grid(row=0, column=3, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
        
        # Current Temperature Display
        temp_frame = ttk.LabelFrame(col4_frame, text="Current Temperature", padding="10")
        temp_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Scrollable temperature display
        temp_canvas = tk.Canvas(temp_frame)
        temp_scrollbar = ttk.Scrollbar(temp_frame, orient="vertical", command=temp_canvas.yview)
        self.temp_display_frame = ttk.Frame(temp_canvas)
        
        self.temp_display_frame.bind(
            "<Configure>",
            lambda e: temp_canvas.configure(scrollregion=temp_canvas.bbox("all"))
        )
        
        temp_canvas.create_window((0, 0), window=self.temp_display_frame, anchor="nw")
        temp_canvas.configure(yscrollcommand=temp_scrollbar.set)
        
        temp_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        temp_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.temp_labels = []
        
        # Statistics Display
        stats_frame = ttk.LabelFrame(col4_frame, text="Statistics", padding="10")
        stats_frame.pack(fill=tk.BOTH, expand=True)
        
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
        
        stats_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        stats_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.stats_labels = []
        
        # Initial status message
        self.log_status("Application started. Configure modules and click 'Apply Configuration'.")

    def log_status(self, message):
        """Add message to status log with timestamp"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"
        
        # Check if status_log widget exists yet
        if not hasattr(self, 'status_log'):
            print(log_message.strip())  # Print to console if widget doesn't exist yet
            return
        
        try:
            self.status_log.config(state=tk.NORMAL)
            self.status_log.insert(tk.END, log_message)
            self.status_log.see(tk.END)  # Auto-scroll to bottom
            self.status_log.config(state=tk.DISABLED)
        except Exception as e:
            print(f"Error logging to status: {e}")
            print(log_message.strip())
    
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
                'modules': [],  # NEW
                'channels': []
            }
            
            # Save module configurations - NEW
            for module in self.modules:
                module_config = {
                    'device_name': module['device_name_var'].get(),
                    'start_channel': module['start_channel_var'].get(),
                    'num_channels': module['num_channels_var'].get()
                }
                config['modules'].append(module_config)
            
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
            
            # Load modules - NEW
            if 'modules' in config and len(config['modules']) > 0:
                # Clear existing modules
                for module in self.modules:
                    module['frame'].destroy()
                self.modules = []
                
                # Load modules from config
                for module_config in config['modules']:
                    self.add_module_config()
                    # Set values for the newly added module
                    last_module = self.modules[-1]
                    last_module['device_name_var'].set(module_config.get('device_name', 'cDAQ2Mod1'))
                    last_module['start_channel_var'].set(module_config.get('start_channel', 0))
                    last_module['num_channels_var'].set(module_config.get('num_channels', 4))
            
            # Apply module configuration
            self.apply_all_modules()
            
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
            import traceback
            traceback.print_exc()
    
    def apply_daq_config(self):
        """Apply DAQ configuration and rebuild channel configuration UI"""
        try:
            # Get configuration values
            device_name = self.device_name_var.get().strip()
            num_channels = self.total_channels_var.get()
            start_channel = self.start_channel_var.get()
            
            if not device_name:
                messagebox.showerror("Error", "Device name cannot be empty!")
                return
            
            if num_channels < 1 or num_channels > 32:
                messagebox.showerror("Error", "Number of channels must be between 1 and 32!")
                return
            
            # Store configuration
            self.device_name = device_name
            self.total_channels = num_channels
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
                                        values=["30 minutes", "1 hour", "6 hours", "1 day", "7 days"],
                                        state="readonly", width=15)
        time_window_run.pack(side=tk.LEFT, padx=5)
        
        ttk.Separator(self.control_bar, orient='vertical').pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Label(self.control_bar, text="Plot Refresh:", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        plot_refresh_combo = ttk.Combobox(
            self.control_bar,
            textvariable=self.plot_refresh_interval_var,
            values=["1 sec", "2 sec", "5 sec", "10 sec"],
            state="readonly",
            width=8
        )
        plot_refresh_combo.pack(side=tk.LEFT, padx=5)
        
        ttk.Separator(self.control_bar, orient='vertical').pack(side=tk.LEFT, fill=tk.Y, padx=10)

        plot_info_label = ttk.Label(self.control_bar, textvariable=self.plot_info_var, font=("Arial", 9))
        plot_info_label.pack(side=tk.LEFT, padx=5)
        
        # Channel Selection label
        ttk.Label(self.control_bar, text="Channels:", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        
        # Channel checkboxes will be added dynamically
        self.run_tab_channel_frame = ttk.Frame(self.control_bar)
        self.run_tab_channel_frame.pack(side=tk.LEFT)
        
        # Add Reset Zoom button - NEW
        ttk.Separator(self.control_bar, orient='vertical').pack(side=tk.LEFT, fill=tk.Y, padx=10)
        ttk.Button(self.control_bar, text="Reset Zoom", command=self.reset_zoom, width=12).pack(side=tk.LEFT, padx=5)
        
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
        ttk.Checkbutton(row2_frame, text="Auto", variable=self.x_axis_auto_var, 
                        command=self.on_axis_auto_change).pack(side=tk.LEFT, padx=2)  # Added command
        
        ttk.Separator(row2_frame, orient='vertical').pack(side=tk.LEFT, fill=tk.Y, padx=10)
        
        # Left Y-Axis Range
        ttk.Label(row2_frame, text="Left Y-Axis:", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=5)
        self.left_y_auto_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(row2_frame, text="Auto", variable=self.left_y_auto_var,
                        command=self.on_axis_auto_change).pack(side=tk.LEFT, padx=2)  # Added command
        
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
        ttk.Checkbutton(row2_frame, text="Auto", variable=self.right_y_auto_var,
                        command=self.on_axis_auto_change).pack(side=tk.LEFT, padx=2)  # Added command
        
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
        
        # Connect zoom callbacks - NEW
        self.ax.callbacks.connect('xlim_changed', self.on_xlims_change)
        self.ax.callbacks.connect('ylim_changed', self.on_ylims_change)
        
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
        """Convert time window string to seconds (max 7 days)"""
        window_str = self.time_window_var.get().strip().lower()
    
        if "minute" in window_str:
            return int(window_str.split()[0]) * 60
        if "hour" in window_str:
            return int(window_str.split()[0]) * 3600
        if "day" in window_str:
            days = int(window_str.split()[0])
            return min(days, 7) * 86400
    
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
            header = ["Timestamp"] + [self.channel_label_vars[i].get() for i in range(self.total_channels)]
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
                    channels_data = {i: [] for i in range(self.total_channels)}
                    timestamps_data = []
                    
                    # Collect buffered data
                    for ts, temps in self.tdms_buffer:
                        for i, temp in enumerate(temps):
                            if i < self.total_channels:
                                channels_data[i].append(float(temp))  # Ensure float
                        
                        # Convert timestamp to Excel date format
                        excel_epoch = datetime(1900, 1, 1)
                        delta = ts - excel_epoch
                        excel_timestamp = delta.total_seconds() / 86400 + 2
                        timestamps_data.append(excel_timestamp)
                    
                    # Create channel objects
                    channels = []
                    for i in range(self.total_channels):
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
                channels_data = {i: [] for i in range(self.total_channels)}
                timestamps_data = []
                
                # Collect buffered data
                for ts, temps in self.tdms_buffer:
                    for i, temp in enumerate(temps):
                        if i < self.total_channels:
                            channels_data[i].append(float(temp))  # Ensure float
                    
                    # Convert timestamp to Excel date format
                    excel_epoch = datetime(1900, 1, 1)
                    delta = ts - excel_epoch
                    excel_timestamp = delta.total_seconds() / 86400 + 2
                    timestamps_data.append(excel_timestamp)
                
                # Create channel objects with 1D numpy arrays
                channels = []
                for i in range(self.total_channels):
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
            
            # (NEW) clear ring buffer for a fresh run
            if self.ring is not None:
                self.ring.clear()
            
            # Update acquisition rate
            self.update_acquisition_rate()
            
            # Ensure ring capacity matches acquisition rate (in case rate changed)
            if self.ring is not None:
                self.rebuild_ring_buffer()
            
            # Clear previous data
            with self.data_lock:
                # Keep these for compatibility with any remaining code paths,
                # but DO NOT use them for long-run storage anymore.
                # self.timestamps = []
                # self.temperature_data = [[] for _ in range(self.total_channels)]
                self.timestamps = None
                self.temperature_data = None
            
            # Create initial files
            self.create_new_files()
            
            # Update UI
            self.start_button.config(state="disabled")
            self.stop_button.config(state="normal")
            self.acquisition_running = True
            
            # (NEW) initialize plot line objects once
            self.init_plot_lines()
            
            # Disable configuration widgets
            self.disable_config_widgets()
            
            self.log_status("=== Acquisition Started ===")
            self.log_status(f"Rate: {self.acq_rate_var.get()}")
            self.log_status(f"Total Channels: {self.total_channels}")
            
            # Start acquisition thread
            self.acq_thread = threading.Thread(target=self.acquisition_loop, daemon=True)
            self.acq_thread.start()
            
            # Start stats update loop (NEW)
            self.stats_timer_loop()
            
            # Start plot update timer
            self.update_plot_from_ring()
            
        except Exception as e:
            self.log_status(f"ERROR starting acquisition: {str(e)}")
            messagebox.showerror("Error", f"Failed to start acquisition:\n{str(e)}")
            import traceback
            traceback.print_exc()
            self.stop_acquisition()
    
    def stop_acquisition(self):
        """Stop data acquisition with confirmation dialog"""
        # Create custom confirmation dialog
        response = self.confirm_stop_dialog()
        
        if not response:
            self.log_status("Stop acquisition cancelled by user")
            return
        
        # User confirmed - proceed with stopping
        self.acquisition_running = False
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        
        # Re-enable configuration widgets
        self.enable_config_widgets()
        
        self.log_status("=== Acquisition Stopped ===")
        
        # Close files
        self.close_files()
    
    def confirm_stop_dialog(self):
        """Custom confirmation dialog for stopping acquisition"""
        # Create a custom dialog window
        dialog = tk.Toplevel(self.root)
        dialog.title("Stop Acquisition Confirmation")
        dialog.geometry("500x200")  # Increased size
        dialog.resizable(False, False)
        
        # Make it modal
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog on the parent window
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (250)  # Half of width
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (100)  # Half of height
        dialog.geometry(f"500x200+{x}+{y}")
        
        # Variable to store result
        result = [False]
        
        # Main container
        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Warning icon and message
        message_frame = ttk.Frame(main_frame)
        message_frame.pack(fill=tk.BOTH, expand=True)
        
        warning_label = ttk.Label(message_frame, 
                 text="⚠️  Stop Data Acquisition?", 
                 font=("Arial", 14, "bold"),
                 foreground="red")
        warning_label.pack(pady=(0, 15))
        
        message_label = ttk.Label(message_frame, 
                 text="Are you sure you want to STOP data acquisition?\n\nThis will close the current data files.",
                 font=("Arial", 10),
                 justify=tk.CENTER)
        message_label.pack(pady=(0, 20))
        
        # Button frame at the bottom
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
        
        # Configure button frame columns
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        
        def on_no():
            result[0] = False
            dialog.destroy()
        
        def on_yes():
            result[0] = True
            dialog.destroy()
        
        # Create buttons side by side
        no_button = ttk.Button(button_frame, text="No, Continue Recording", 
                              command=on_no)
        no_button.grid(row=0, column=0, padx=10, pady=5, sticky=(tk.W, tk.E))
        
        yes_button = ttk.Button(button_frame, text="Yes, Stop Recording", 
                               command=on_yes)
        yes_button.grid(row=0, column=1, padx=10, pady=5, sticky=(tk.W, tk.E))
        
        # Set focus to No button and make it default
        no_button.focus_set()
        dialog.bind('<Return>', lambda e: on_no())
        dialog.bind('<Escape>', lambda e: on_no())
        
        # Handle window close button (X)
        dialog.protocol("WM_DELETE_WINDOW", on_no)
        
        # Force update to ensure everything is drawn
        dialog.update_idletasks()
        
        # Wait for dialog to close
        dialog.wait_window()
        
        return result[0]
    
    def acquisition_loop(self):
        """Main acquisition loop running in separate thread with hardware timing"""
        try:
            with nidaqmx.Task() as task:
                # Add thermocouple channels
                tc_type = self.get_tc_type()
                temp_units = self.get_temp_units()
                
                for channel in self.channels:  # This now includes channels from all modules
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
                                elif self.total_channels == 1:
                                    # Single channel with single sample
                                    data = [data[0]] if isinstance(data, list) else [data]
                            
                            # Check for file rotation
                            if self.check_file_rotation():
                                self.create_new_files()
                                self.root.after(0, lambda: self.log_status("File rotated - new file created"))
                            
                            # Store data
                            # NEW: send sample to ring buffer (bounded memory) for plotting/stats
                            if self.ring is not None:
                                t_sec = time.time()
                                with self.data_lock:
                                    try:
                                        self.ring.append(t_sec, data)
                                    except Exception:
                                        pass
                            
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
                            
                            # Store data to ring buffer (bounded memory) for plotting/stats
                            if self.ring is not None:
                                t_sec = time.time()
                                with self.data_lock:
                                    try:
                                        self.ring.append(t_sec, data)
                                    except Exception:
                                        pass
                            
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
            for i in range(self.total_channels):
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
        
        # Set flag to indicate we're programmatically updating
        self._updating_plot = True
        
        try:
            with self.data_lock:
                if len(self.timestamps) == 0:
                    self._updating_plot = False
                    self.root.after(1000, self.update_plot)
                    return
                
                # Get time window in seconds
                window_seconds = self.get_time_window_seconds()
                current_time = datetime.now()
                start_time = current_time - timedelta(seconds=window_seconds)
                
                # Filter data within time window (only if X-axis auto is ON)
                if self.x_axis_auto_var.get() and not self.user_zoomed:
                    # Auto mode: use time window
                    indices = [i for i, t in enumerate(self.timestamps) if t >= start_time]
                else:
                    # Manual zoom mode: use all data
                    indices = list(range(len(self.timestamps)))
                
                if len(indices) == 0:
                    self._updating_plot = False
                    self.root.after(1000, self.update_plot)
                    return
                
                plot_times = [self.timestamps[i] for i in indices]
                plot_data = [[self.temperature_data[ch][i] for i in indices] 
                            for ch in range(self.total_channels)]
                
                # Decimate data for plotting (max 1000 points)
                if len(plot_times) > 1000:
                    step = len(plot_times) // 1000
                    plot_times = plot_times[::step]
                    plot_data = [data[::step] for data in plot_data]
            
            # Calculate data ranges for auto-scaling
            left_axis_data = []
            right_axis_data = []
            
            for i in range(self.total_channels):
                if self.channel_selection_vars[i].get() and len(plot_data[i]) > 0:
                    if self.channel_yaxis_vars[i].get() == "Right":
                        right_axis_data.extend(plot_data[i])
                    else:
                        left_axis_data.extend(plot_data[i])
            
            # Calculate auto ranges with 5% margin
            if left_axis_data:
                left_min_data = min(left_axis_data)
                left_max_data = max(left_axis_data)
                left_range = left_max_data - left_min_data
                if left_range == 0:
                    left_range = 1.0
                left_margin = left_range * 0.05
                left_auto_min = left_min_data - left_margin
                left_auto_max = left_max_data + left_margin
            else:
                left_auto_min, left_auto_max = 0, 100
            
            if right_axis_data:
                right_min_data = min(right_axis_data)
                right_max_data = max(right_axis_data)
                right_range = right_max_data - right_min_data
                if right_range == 0:
                    right_range = 1.0
                right_margin = right_range * 0.05
                right_auto_min = right_min_data - right_margin
                right_auto_max = right_max_data + right_margin
            else:
                right_auto_min, right_auto_max = 0, 100
            
            # Store current zoom state ONLY if user has manually zoomed AND respective auto is OFF
            current_xlim = None
            current_ylim_left = None
            current_ylim_right = None
            
            if hasattr(self, 'ax'):
                # Only preserve X zoom if auto is OFF and user zoomed
                if self.user_zoomed and not self.x_axis_auto_var.get():
                    current_xlim = self.ax.get_xlim()
                
                # Only preserve Y zoom if auto is OFF and user zoomed
                if self.user_zoomed and not self.left_y_auto_var.get():
                    current_ylim_left = self.ax.get_ylim()
            
            # Check if we have a right axis
            had_right_axis = hasattr(self, 'ax_right') and self.ax_right is not None
            if had_right_axis:
                if self.user_zoomed and not self.right_y_auto_var.get():
                    current_ylim_right = self.ax_right.get_ylim()
            
            # Clear the figure
            self.figure.clear()
            
            # Create primary axis
            self.ax = self.figure.add_subplot(111)
            
            # Reconnect zoom callbacks after clearing
            self.ax.callbacks.connect('xlim_changed', self.on_xlims_change)
            self.ax.callbacks.connect('ylim_changed', self.on_ylims_change)
            
            # Check if we need a secondary Y-axis
            has_right_axis = any(self.channel_yaxis_vars[i].get() == "Right" 
                                and self.channel_selection_vars[i].get() 
                                for i in range(self.total_channels))
            
            # Create secondary axis if needed
            if has_right_axis:
                self.ax_right = self.ax.twinx()
            else:
                self.ax_right = None
            
            # Plot channels on appropriate axes
            left_plotted = False
            right_plotted = False
            
            for i in range(self.total_channels):
                if self.channel_selection_vars[i].get():
                    linestyle, marker = self.get_plot_style(i)
                    color = self.channel_color_vars[i].get()
                    label = self.channel_label_vars[i].get()
                    yaxis_side = self.channel_yaxis_vars[i].get()
                    
                    # Choose which axis to plot on
                    if yaxis_side == "Right" and self.ax_right is not None:
                        current_ax = self.ax_right
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
            
            # Set X-axis limits
            if not self.x_axis_auto_var.get() and current_xlim is not None:
                # Manual mode: restore user's zoom
                self.ax.set_xlim(current_xlim)
            
            # Configure left Y-axis
            left_ylabel = self.left_yaxis_title_var.get()
            self.ax.set_ylabel(left_ylabel, fontsize=12, color='black')
            self.ax.tick_params(axis='y', labelcolor='black')
            self.ax.grid(True, alpha=0.3)
            
            # Set left Y-axis limits - SIMPLIFIED LOGIC
            if self.left_y_auto_var.get():
                # AUTO MODE - always use calculated range (ignore user_zoomed for Y when auto is ON)
                self.ax.set_ylim(left_auto_min, left_auto_max)
            else:
                # MANUAL MODE - use fixed range from entry boxes, or preserve user zoom
                if current_ylim_left is not None:
                    self.ax.set_ylim(current_ylim_left)
                else:
                    try:
                        left_min = float(self.left_y_min_var.get())
                        left_max = float(self.left_y_max_var.get())
                        self.ax.set_ylim(left_min, left_max)
                    except ValueError:
                        self.ax.set_ylim(left_auto_min, left_auto_max)
            
            # Configure right Y-axis if it exists
            if self.ax_right is not None and right_plotted:
                right_ylabel = self.right_yaxis_title_var.get()
                self.ax_right.set_ylabel(right_ylabel, fontsize=12, color='black')
                self.ax_right.tick_params(axis='y', labelcolor='black')
                
                # Set right Y-axis limits - SIMPLIFIED LOGIC
                if self.right_y_auto_var.get():
                    # AUTO MODE - always use calculated range
                    self.ax_right.set_ylim(right_auto_min, right_auto_max)
                else:
                    # MANUAL MODE - use fixed range from entry boxes, or preserve user zoom
                    if current_ylim_right is not None:
                        self.ax_right.set_ylim(current_ylim_right)
                    else:
                        try:
                            right_min = float(self.right_y_min_var.get())
                            right_max = float(self.right_y_max_var.get())
                            self.ax_right.set_ylim(right_min, right_max)
                        except ValueError:
                            self.ax_right.set_ylim(right_auto_min, right_auto_max)
            
            # Combine legends from both axes
            if left_plotted or right_plotted:
                lines_1, labels_1 = self.ax.get_legend_handles_labels()
                if self.ax_right is not None and right_plotted:
                    lines_2, labels_2 = self.ax_right.get_legend_handles_labels()
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
        finally:
            # Clear flag after update is complete
            self._updating_plot = False
        
        # Schedule next update
        self.root.after(1000, self.update_plot)
        
    def on_closing(self):
        """Handle window closing"""
        if self.acquisition_running:
            # Create custom confirmation dialog
            dialog = tk.Toplevel(self.root)
            dialog.title("Exit Application")
            dialog.geometry("500x200")
            dialog.resizable(False, False)
            
            # Make it modal
            dialog.transient(self.root)
            dialog.grab_set()
            
            # Center the dialog
            dialog.update_idletasks()
            x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (250)
            y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (100)
            dialog.geometry(f"500x200+{x}+{y}")
            
            # Variable to store result
            result = [False]
            
            # Main container
            main_frame = ttk.Frame(dialog)
            main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
            
            # Warning message
            message_frame = ttk.Frame(main_frame)
            message_frame.pack(fill=tk.BOTH, expand=True)
            
            warning_label = ttk.Label(message_frame, 
                     text="⚠️  Acquisition is Running!", 
                     font=("Arial", 14, "bold"),
                     foreground="red")
            warning_label.pack(pady=(0, 15))
            
            message_label = ttk.Label(message_frame, 
                     text="Data acquisition is currently running.\n\nDo you want to stop and exit?",
                     font=("Arial", 10),
                     justify=tk.CENTER)
            message_label.pack(pady=(0, 20))
            
            # Button frame at the bottom
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
            
            # Configure button frame columns
            button_frame.columnconfigure(0, weight=1)
            button_frame.columnconfigure(1, weight=1)
            
            def on_cancel():
                result[0] = False
                dialog.destroy()
            
            def on_exit():
                result[0] = True
                dialog.destroy()
            
            # Create buttons
            cancel_button = ttk.Button(button_frame, text="No, Continue Recording", 
                                       command=on_cancel)
            cancel_button.grid(row=0, column=0, padx=10, pady=5, sticky=(tk.W, tk.E))
            
            exit_button = ttk.Button(button_frame, text="Yes, Stop and Exit", 
                                    command=on_exit)
            exit_button.grid(row=0, column=1, padx=10, pady=5, sticky=(tk.W, tk.E))
            
            # Set focus to Cancel button and make it default
            cancel_button.focus_set()
            dialog.bind('<Return>', lambda e: on_cancel())
            dialog.bind('<Escape>', lambda e: on_cancel())
            dialog.protocol("WM_DELETE_WINDOW", on_cancel)
            
            # Force update
            dialog.update_idletasks()
            
            # Wait for dialog
            dialog.wait_window()
            
            if result[0]:
                self.acquisition_running = False
                self.close_files()
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
            # If X-axis auto is unchecked, it means user wants to maintain zoom
            if not self.x_axis_auto_var.get():
                self.user_zoomed = True
            else:
                self.user_zoomed = False
                self.plot_xlim = None
            
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
        self.update_plot_from_ring()

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

    def add_module_config(self, log=True):
        """Add a new module configuration UI"""
        module_index = len(self.modules)
        
        # Create module frame
        module_frame = ttk.LabelFrame(self.modules_container, 
                                      text=f"Module {module_index + 1}", 
                                      padding="10")
        module_frame.pack(fill=tk.X, pady=5, padx=5)
        
        # Module data structure
        module_data = {
            'frame': module_frame,
            'device_name_var': tk.StringVar(value=f"cDAQ2Mod{module_index + 1}"),
            'start_channel_var': tk.IntVar(value=0),
            'num_channels_var': tk.IntVar(value=4),
            'index': module_index
        }
        
        # Device Name
        ttk.Label(module_frame, text="Device Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
        device_entry = ttk.Entry(module_frame, textvariable=module_data['device_name_var'], width=18)
        device_entry.grid(row=0, column=1, pady=5, padx=5)
        
        # Start Channel
        ttk.Label(module_frame, text="Start Channel (ai):").grid(row=1, column=0, sticky=tk.W, pady=5)
        start_ch_spinbox = ttk.Spinbox(module_frame, from_=0, to=31, 
                                       textvariable=module_data['start_channel_var'], 
                                       width=16)
        start_ch_spinbox.grid(row=1, column=1, pady=5, padx=5)
        
        # Number of Channels
        ttk.Label(module_frame, text="Number of Channels:").grid(row=2, column=0, sticky=tk.W, pady=5)
        num_ch_spinbox = ttk.Spinbox(module_frame, from_=1, to=32, 
                                     textvariable=module_data['num_channels_var'], 
                                     width=16)
        num_ch_spinbox.grid(row=2, column=1, pady=5, padx=5)
        
        # Delete button
        delete_btn = ttk.Button(module_frame, text="🗑️ Delete Module", 
                               command=lambda idx=module_index: self.delete_module_config(idx),
                               width=18)
        delete_btn.grid(row=3, column=0, columnspan=2, pady=(10, 0))
        module_data['delete_button'] = delete_btn
        
        # Store module data
        self.modules.append(module_data)
        
        # Update module numbering
        self.update_module_numbers()
        
        # Only log if requested (not during initial setup)
        if log:
            self.log_status(f"Module {module_index + 1} configuration added")
    
    def delete_module_config(self, module_index):
        """Delete a module configuration"""
        if len(self.modules) <= 1:
            messagebox.showwarning("Warning", "Cannot delete the last module!\n\nAt least one module is required.")
            return
        
        # Find the actual module by index (not position in list)
        module_to_delete = None
        list_position = -1
        for i, module in enumerate(self.modules):
            if module['index'] == module_index:
                module_to_delete = module
                list_position = i
                break
        
        if module_to_delete is None:
            return
        
        # Confirm deletion
        response = messagebox.askyesno(
            "Delete Module",
            f"Are you sure you want to delete {module_to_delete['device_name_var'].get()}?"
        )
        
        if response:
            # Remove from GUI
            module_to_delete['frame'].destroy()
            
            # Remove from list
            self.modules.pop(list_position)
            
            # Update module numbers
            self.update_module_numbers()
            
            self.log_status(f"Module configuration deleted")
    
    def update_module_numbers(self):
        """Update module frame titles after add/delete"""
        for i, module in enumerate(self.modules):
            module['frame'].config(text=f"Module {i + 1}")
    
    def apply_all_modules(self):
        """Apply configuration for all modules and rebuild channel UI"""
        try:
            if len(self.modules) == 0:
                messagebox.showerror("Error", "No modules configured!")
                return
            
            # Build channel list from all modules
            all_channels = []
            module_info = []
            
            for module in self.modules:
                device_name = module['device_name_var'].get().strip()
                num_channels = module['num_channels_var'].get()
                start_channel = module['start_channel_var'].get()
                
                if not device_name:
                    messagebox.showerror("Error", f"Module {module['index'] + 1}: Device name cannot be empty!")
                    return
                
                if num_channels < 1 or num_channels > 32:
                    messagebox.showerror("Error", f"Module {module['index'] + 1}: Number of channels must be between 1 and 32!")
                    return
                
                # Build channels for this module
                for i in range(num_channels):
                    channel_name = f"{device_name}/ai{start_channel + i}"
                    all_channels.append(channel_name)
                    module_info.append({
                        'device': device_name,
                        'channel_index': start_channel + i,
                        'module_index': module['index']
                    })
            
            # Store configuration
            self.channels = all_channels
            self.module_info = module_info
            self.total_channels = len(all_channels)
            
            self.update_acquisition_rate()
            self.rebuild_ring_buffer() # NEW
            
########### Just a verificaion - can be removed if needed
            self.log_status(f"Ring buffer: channels={self.ring.channels}, capacity={self.ring.capacity} samples")
            
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
            self.channel_yaxis_vars = []
            self.temp_labels = []
            self.stats_labels = []
            self.temperature_data = [[] for _ in range(self.total_channels)]
            self.channel_name_entries = []
            self.temp_display_labels = []
            
            # Create channel configuration UI for all channels
            for i in range(self.total_channels):
                info = module_info[i]
                device_name = info['device']
                channel_index = info['channel_index']
                
                ch_frame = ttk.LabelFrame(self.channels_container, 
                                         text=f"Ch{i}: {device_name}/ai{channel_index}", 
                                         padding="5")
                ch_frame.pack(fill=tk.X, pady=5, padx=5)
                
                # Enable checkbox
                var = tk.BooleanVar(value=True)
                self.channel_selection_vars.append(var)
                ttk.Checkbutton(ch_frame, text="Enable", variable=var).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=2)
                
                # Custom Label
                ttk.Label(ch_frame, text="Label:").grid(row=1, column=0, sticky=tk.W, pady=2)
                label_var = tk.StringVar(value=f"{device_name}_ai{channel_index}")
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
                
                # Y-Axis selection
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
                
                channel_display_label = ttk.Label(temp_frame, text=f"{device_name}_ai{channel_index}:", 
                                                 font=("Arial", 10, "bold"), width=20)
                channel_display_label.pack(side=tk.LEFT, padx=5)
                
                temp_label = ttk.Label(temp_frame, text="-- °C", font=("Arial", 11), foreground="darkgreen")
                temp_label.pack(side=tk.LEFT, padx=5)
                self.temp_labels.append(temp_label)
                self.temp_display_labels.append(channel_display_label)
                
                # Create statistics display
                stats_frame_inner = ttk.Frame(self.stats_display_frame)
                stats_frame_inner.pack(fill=tk.X, pady=2)
                
                ttk.Label(stats_frame_inner, text=f"Ch{i}:", width=5).pack(side=tk.LEFT)
                label = ttk.Label(stats_frame_inner, text="Min: -- | Max: -- | Avg: --", font=("Arial", 9))
                label.pack(side=tk.LEFT, padx=5)
                self.stats_labels.append(label)
            
            # Rebuild Run tab channel controls
            self.rebuild_run_tab_controls()
            
            # Log summary
            summary = f"Configuration Applied: {self.total_channels} total channels across {len(self.modules)} module(s)"
            for module in self.modules:
                summary += f"\n  - {module['device_name_var'].get()}: {module['num_channels_var'].get()} channels"
            
            self.log_status(summary)
            messagebox.showinfo("Success", f"Configuration Applied!\n\n{self.total_channels} channels configured across {len(self.modules)} module(s)")
            
        except Exception as e:
            self.log_status(f"Error applying configuration: {str(e)}")
            messagebox.showerror("Error", f"Failed to apply configuration:\n{str(e)}")
            import traceback
            traceback.print_exc()

    def browse_data_path(self):
        """Browse for data save directory"""
        directory = filedialog.askdirectory(
            initialdir=self.data_path_var.get(),
            title="Select Data Save Directory"
        )
        if directory:
            self.data_path_var.set(directory)
            self.log_status(f"Data path set to: {directory}")
    
    def browse_config_path(self):
        """Browse for configuration file directory"""
        directory = filedialog.askdirectory(
            initialdir=self.config_path_var.get(),
            title="Select Configuration File Directory"
        )
        if directory:
            self.config_path_var.set(directory)
            self.log_status(f"Config path set to: {directory}")
    
    def on_xlims_change(self, event_ax):
        """Called when x-axis limits change (user zoomed)"""
        if hasattr(self, 'ax') and event_ax == self.ax:
            # Only mark as user zoomed if we're not currently updating the plot
            if hasattr(self, '_updating_plot') and self._updating_plot:
                return  # Ignore programmatic changes
            self.plot_xlim = self.ax.get_xlim()
            self.user_zoomed = True
            print("X-axis: User zoom detected")
    
    def on_ylims_change(self, event_ax):
        """Called when y-axis limits change (user zoomed)"""
        if hasattr(self, 'ax') and event_ax == self.ax:
            # Only mark as user zoomed if we're not currently updating the plot
            if hasattr(self, '_updating_plot') and self._updating_plot:
                return  # Ignore programmatic changes
            self.plot_ylim_left = self.ax.get_ylim()
            self.user_zoomed = True
            print("Y-axis: User zoom detected")
            self.user_zoomed = True
    
    def reset_zoom(self):
        """Reset zoom to auto mode"""
        self.user_zoomed = False
        self.plot_xlim = None
        self.plot_ylim_left = None
        self.plot_ylim_right = None
        self.x_axis_auto_var.set(True)
        self.left_y_auto_var.set(True)
        self.right_y_auto_var.set(True)
        self.log_status("Zoom reset to auto mode")

    def on_axis_auto_change(self):
        """Called when any axis auto checkbox is toggled"""
        # When switching from manual to auto on X-axis, reset zoom
        if self.x_axis_auto_var.get():
            self.user_zoomed = False
            self.plot_xlim = None
        
        # Log the change
        status_parts = []
        if self.x_axis_auto_var.get():
            status_parts.append("X-axis: Auto")
        else:
            status_parts.append("X-axis: Manual zoom enabled")
        
        if self.left_y_auto_var.get():
            status_parts.append("Left Y-axis: Auto")
        else:
            status_parts.append("Left Y-axis: Fixed range")
        
        if self.right_y_auto_var.get():
            status_parts.append("Right Y-axis: Auto")
        else:
            status_parts.append("Right Y-axis: Fixed range")
        
        self.log_status(" | ".join(status_parts))
        
    def rebuild_ring_buffer(self):
        """
        (Re)create the in-memory ring buffer for plotting/stats.
        Keeps only last 7 days @ acquisition_rate (1 Hz preferred).
        """
        # We want capacity based on configured acquisition rate (Hz).
        # If acquisition_rate isn't set yet or is invalid, assume 1 Hz.
        try:
            hz = float(self.acquisition_rate)
            if hz <= 0:
                hz = 1.0
        except Exception:
            hz = 1.0
    
        # Capacity = window_seconds * Hz, plus a small margin
        capacity = int(self.max_plot_window_seconds * hz) + 10
        channels = int(self.total_channels)
    
        self.ring = MultiChannelRingBuffer(channels=channels, capacity=capacity, dtype=np.float32)
        
    def get_plot_refresh_seconds(self) -> int:
        """
        Parse plot refresh interval dropdown text like '2 sec' -> 2.
        Default to 2 seconds if parsing fails.
        """
        try:
            s = self.plot_refresh_interval_var.get().strip().lower()
            # expected formats: '1 sec', '2 sec', ...
            n = int(s.split()[0])
            return max(1, n)
        except Exception:
            return 2

    def init_plot_lines(self):
        """Initialize plot axes + line objects once (call when starting acquisition or when config changes)."""
        # Create axes if not present
        if not hasattr(self, "figure") or self.figure is None:
            return
    
        self.figure.clear()
        self.ax = self.figure.add_subplot(111)
        self.ax.grid(True, alpha=0.3)
    
        self.ax_right = None
        self.lines_left = [None] * self.total_channels
        self.lines_right = [None] * self.total_channels
    
        # Decide if any channel is assigned to right axis
        has_right = any(
            self.channel_selection_vars[i].get() and self.channel_yaxis_vars[i].get() == "Right"
            for i in range(self.total_channels)
        )
        if has_right:
            self.ax_right = self.ax.twinx()
    
        # Create Line2D objects for enabled channels (others stay None)
        for i in range(self.total_channels):
            if not self.channel_selection_vars[i].get():
                continue
    
            linestyle, marker = self.get_plot_style(i)
            color = self.channel_color_vars[i].get()
            label = self.channel_label_vars[i].get()
            side = self.channel_yaxis_vars[i].get()
    
            target_ax = self.ax_right if (side == "Right" and self.ax_right is not None) else self.ax
    
            # Start with empty data; we will set_data in updates
            (line,) = target_ax.plot(
                [],
                [],
                color=color,
                linestyle=linestyle if linestyle else "None",
                marker=marker if marker else None,
                markersize=4 if marker else 0,
                linewidth=2,
                label=label,
            )
    
            if target_ax is self.ax:
                self.lines_left[i] = line
            else:
                self.lines_right[i] = line
    
        # Titles/labels
        self.ax.set_title(self.plot_title_var.get(), fontsize=14, fontweight="bold")
        self.ax.set_xlabel("Time", fontsize=12)
        self.ax.set_ylabel(self.left_yaxis_title_var.get(), fontsize=12)
    
        if self.ax_right is not None:
            self.ax_right.set_ylabel(self.right_yaxis_title_var.get(), fontsize=12)
    
        # Legend: build once initially (we’ll refresh legend only when needed in Step 5 if you want)
        handles, labels = self.ax.get_legend_handles_labels()
        if self.ax_right is not None:
            h2, l2 = self.ax_right.get_legend_handles_labels()
            handles += h2
            labels += l2
        if handles:
            self.ax.legend(handles, labels, loc="best", fontsize=9, framealpha=0.9)
    
        self.canvas.draw_idle()

    def update_plot_from_ring(self):
        """Efficient plot update from ring buffer with downsampling."""
        if not self.acquisition_running:
            return
    
        refresh_s = self.get_plot_refresh_seconds()
    
        try:
            x, data_ds, n_raw, stride, hz = self.get_window_snapshot()
            if data_ds is None or len(x) == 0:
                self.root.after(self.get_plot_refresh_seconds() * 1000, self.update_plot_from_ring)
                return
            
            n_ds = data_ds.shape[1]
            eff_rate = hz / stride if stride > 0 else hz
            
            
            if self.ring is None or self.ring.count == 0:
                self.root.after(refresh_s * 1000, self.update_plot_from_ring)
                return
    
            window_s = self.get_time_window_seconds()
    
            # number of samples to fetch ~= window_s * acquisition_rate (Hz)
            # use ring capacity logic; fetch a bit more then trim by time if needed
            try:
                hz = float(self.acquisition_rate) if self.acquisition_rate > 0 else 1.0
            except Exception:
                hz = 1.0
    
            n_want = int(window_s * hz) + 5
            snap = self.ring.snapshot_last(n_want)
    
            if snap.count == 0:
                self.root.after(refresh_s * 1000, self.update_plot_from_ring)
                return
    
            # Convert time to matplotlib-friendly datetime objects only after downsampling
            # First trim by time window precisely (since rate may drift)
            t_end = snap.times[-1]
            t_start = t_end - window_s
            # times are sorted in snapshot; find first index >= t_start
            idx0 = int(np.searchsorted(snap.times, t_start, side="left"))
            times = snap.times[idx0:]
            data = snap.data[:, idx0:]
            n_raw = times.shape[0]
            if n_raw <= 1:
                self.root.after(refresh_s * 1000, self.update_plot_from_ring)
                return
    
            # Downsample per channel to max_plot_points_per_channel using stride
            max_pts = int(self.max_plot_points_per_channel)
            stride = int(np.ceil(n_raw / max_pts)) if n_raw > max_pts else 1
            times_ds = times[::stride]
            data_ds = data[:, ::stride]
            n_ds = times_ds.shape[0]
    
            # Effective plot rate (per channel)
            eff_rate = (hz / stride) if stride > 0 else hz
    
            # Convert to datetimes for x-axis
            x = [datetime.fromtimestamp(ts) for ts in times_ds]
    
            # Update lines
            y_left_min = None
            y_left_max = None
            y_right_min = None
            y_right_max = None
    
            for i in range(self.total_channels):
                if not self.channel_selection_vars[i].get():
                    continue
    
                y = data_ds[i, :]
    
                # Skip if all nan/empty
                if y.size == 0:
                    continue
    
                side = self.channel_yaxis_vars[i].get()
                line = self.lines_right[i] if (side == "Right" and self.ax_right is not None) else self.lines_left[i]
                if line is None:
                    continue
    
                line.set_data(x, y)
    
                # Track min/max for autoscale (only for enabled channels)
                ymin = float(np.nanmin(y))
                ymax = float(np.nanmax(y))
    
                if side == "Right":
                    y_right_min = ymin if y_right_min is None else min(y_right_min, ymin)
                    y_right_max = ymax if y_right_max is None else max(y_right_max, ymax)
                else:
                    y_left_min = ymin if y_left_min is None else min(y_left_min, ymin)
                    y_left_max = ymax if y_left_max is None else max(y_left_max, ymax)
            
            if self.plot_needs_rebuild:
                self.plot_needs_rebuild = False
                self.init_plot_lines()
                
            # X-axis handling
            if self.x_axis_auto_var.get():
                self.ax.set_xlim(x[0], x[-1])
    
            # Y-axis handling with margin
            def apply_auto_ylim(ax, ymin, ymax):
                if ymin is None or ymax is None:
                    return
                r = ymax - ymin
                if r == 0:
                    r = 1.0
                m = r * 0.05
                ax.set_ylim(ymin - m, ymax + m)
    
            # Left Y
            if self.left_y_auto_var.get():
                apply_auto_ylim(self.ax, y_left_min, y_left_max)
            else:
                # manual fixed range
                try:
                    self.ax.set_ylim(float(self.left_y_min_var.get()), float(self.left_y_max_var.get()))
                except Exception:
                    pass
    
            # Right Y
            if self.ax_right is not None:
                if self.right_y_auto_var.get():
                    apply_auto_ylim(self.ax_right, y_right_min, y_right_max)
                else:
                    try:
                        self.ax_right.set_ylim(float(self.right_y_min_var.get()), float(self.right_y_max_var.get()))
                    except Exception:
                        pass
    
            # Update titles/labels (lightweight)
            self.ax.set_title(self.plot_title_var.get(), fontsize=14, fontweight="bold")
            self.ax.set_ylabel(self.left_yaxis_title_var.get(), fontsize=12)
            if self.ax_right is not None:
                self.ax_right.set_ylabel(self.right_yaxis_title_var.get(), fontsize=12)
    
            self.figure.autofmt_xdate()
    
            # Update info label
            self.plot_info_var.set(
                f"Plot: {n_ds}/{n_raw} pts | stride={stride} | eff.rate={eff_rate:.4g} Hz | refresh={refresh_s}s"
            )
    
            self.canvas.draw_idle()
    
        except Exception as e:
            # Keep UI alive even if plot update fails
            # print(f"Plot update error: {e}")
            pass
    
        self.root.after(refresh_s * 1000, self.update_plot_from_ring)

    def get_window_snapshot(self):
        """
        Returns (times_ds, data_ds, n_raw, stride, hz) for the current time-window,
        downsampled to max_plot_points_per_channel.
        times_ds are datetime objects for plotting/stats display.
        data_ds shape: (channels, n_ds)
        """
        if self.ring is None or self.ring.count == 0:
            return [], None, 0, 1, 1.0
    
        window_s = self.get_time_window_seconds()
        try:
            hz = float(self.acquisition_rate) if self.acquisition_rate and self.acquisition_rate > 0 else 1.0
        except Exception:
            hz = 1.0
    
        n_want = int(window_s * hz) + 5
        snap = self.ring.snapshot_last(n_want)
        if snap.count == 0:
            return [], None, 0, 1, hz
    
        t_end = snap.times[-1]
        t_start = t_end - window_s
        idx0 = int(np.searchsorted(snap.times, t_start, side="left"))
        times = snap.times[idx0:]
        data = snap.data[:, idx0:]
        n_raw = times.shape[0]
        if n_raw <= 1:
            return [], None, n_raw, 1, hz
    
        max_pts = int(self.max_plot_points_per_channel)
        stride = int(np.ceil(n_raw / max_pts)) if n_raw > max_pts else 1
    
        times_ds = times[::stride]
        data_ds = data[:, ::stride]
    
        x_dt = [datetime.fromtimestamp(ts) for ts in times_ds]
        return x_dt, data_ds, n_raw, stride, hz

    def update_statistics_from_ring(self):
        """
        Update per-channel stats using the currently selected time window.
        Uses ring buffer snapshot (bounded, fast).
        """
        try:
            x, data_ds, n_raw, stride, hz = self.get_window_snapshot()
            if data_ds is None or data_ds.shape[1] == 0:
                return
    
            # Stats computed on downsampled window for speed.
            # If you want exact stats, we can compute on raw window later.
            unit_symbol = self.get_temp_unit_symbol()
    
            for i in range(self.total_channels):
                # Only compute stats for selected channels (as requested)
                if not self.channel_selection_vars[i].get():
                    continue
    
                y = data_ds[i, :]
                if y.size == 0:
                    continue
    
                # NaN-safe
                y_min = float(np.nanmin(y))
                y_max = float(np.nanmax(y))
                y_avg = float(np.nanmean(y))
    
                if i < len(self.stats_labels):
                    self.stats_labels[i].config(
                        text=f"Min: {y_min:.2f} {unit_symbol} | Max: {y_max:.2f} {unit_symbol} | Avg: {y_avg:.2f} {unit_symbol}"
                    )
        except Exception:
            # Keep UI safe
            pass

    def stats_timer_loop(self):
        if not self.acquisition_running:
            return
        self.update_statistics_from_ring()
        self.root.after(5000, self.stats_timer_loop)  # 5 seconds

    def request_plot_rebuild(self):
        """Mark plot for rebuild on next refresh (safe from UI callbacks)."""
        self.plot_needs_rebuild = True

def main():
    root = tk.Tk()
    app = ThermocopleDAQGUI(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()
