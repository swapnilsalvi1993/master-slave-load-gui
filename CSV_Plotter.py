import tkinter as tk
from tkinter import filedialog, ttk, colorchooser, messagebox
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.ticker import MultipleLocator
from itertools import cycle
import json
from datetime import datetime, timedelta
import io
import os
from pathlib import Path
import matplotlib.dates as mdates

# Excel date conversion constant (Excel epoch: January 1, 1900)
EXCEL_EPOCH = datetime(1899, 12, 30)  # Note: Excel incorrectly treats 1900 as a leap year

# Try to import nptdms for TDMS support
try:
    from nptdms import TdmsFile
    TDMS_AVAILABLE = True
except ImportError:
    TDMS_AVAILABLE = False
    print("Warning: nptdms not installed. TDMS mode will not be available.")
    print("Install with: pip install npTDMS")

DEBOUNCE_MS = 1000  # debounce interval for live preview (milliseconds)


class ScrollableFrame:
    """
    Generic scrollable frame: canvas + inner frame + scrollbars.
    """
    def __init__(self, master, width=None, height=None):
        self.container = ttk.Frame(master)
        self.canvas = tk.Canvas(self.container, borderwidth=0, highlightthickness=0)
        self.v_scroll = ttk.Scrollbar(self.container, orient="vertical", command=self.canvas.yview)
        self.h_scroll = ttk.Scrollbar(self.container, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)

        self.v_scroll.pack(side="right", fill="y")
        self.h_scroll.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.inner = ttk.Frame(self.canvas)
        self.window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        if width:
            self.container.config(width=width)
            self.canvas.config(width=width)
        if height:
            self.container.config(height=height)
            self.canvas.config(height=height)

        self.inner.bind("<Configure>", lambda e: self._on_frame_configure())
        self.canvas.bind("<Configure>", lambda e: self._on_canvas_configure())
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel, add="+")

    def pack(self, **kwargs): self.container.pack(**kwargs)
    def grid(self, **kwargs): self.container.grid(**kwargs)
    def place(self, **kwargs): self.container.place(**kwargs)

    def _on_frame_configure(self):
        try:
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        except Exception:
            pass

    def _on_canvas_configure(self):
        try:
            canvas_w = self.canvas.winfo_width()
            self.canvas.itemconfig(self.window_id, width=canvas_w)
        except Exception:
            pass

    def _on_mousewheel(self, event):
        # vertical scroll if vertical overflow else horizontal
        try:
            if self.canvas.winfo_height() < self.inner.winfo_height():
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            else:
                if self.canvas.winfo_width() < self.inner.winfo_width():
                    self.canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass
        
        
class CSVPlotter:
    def __init__(self, root):
        self.root = root
        self.root.title("Interactive CSV/TDMS Plotter")
        self.df = None

        # Data source mode: 'csv' or 'tdms'
        self.data_mode = tk.StringVar(value='csv')
        
        # TDMS-specific attributes
        self.tdms_folder = None
        self.tdms_files = []
        self.tdms_groups = {}  # {group_name: {channel_name: data_array}}
        self.tdms_time_channel = None  # Store time/index data
        self.tdms_time_offset_hours = 0.0  # Time offset between files (set to 0, files concatenated directly)

        # per-channel frames
        self.line_properties_frames = []
        self.right_line_properties_frames = []
        self.right2_line_properties_frames = []

        # persistent property maps keyed by column name
        self.left_channel_props = {}
        self.right_channel_props = {}
        self.right2_channel_props = {}

        # detected header index for PEC-style CSV
        self.detected_header_line_index = None

        # color cycles
        self.color_cycle = cycle(plt.rcParams['axes.prop_cycle'].by_key()['color'])
        self.right_color_cycle = cycle(plt.rcParams['axes.prop_cycle'].by_key()['color'])
        self.right2_color_cycle = cycle(plt.rcParams['axes.prop_cycle'].by_key()['color'])

        # defaults
        self.default_plot_width = 12.0
        self.default_plot_height = 8.0

        # preview scheduling
        self._preview_after_id = None
        self._preview_cancel_requested = False
        self._explicit_preview_request = False

        # UI state variables
        self.live_preview_var = tk.BooleanVar(value=True)
        self.downsample_live_var = tk.BooleanVar(value=True)
        self.max_preview_points = tk.IntVar(value=5000)
        self.marker_size = tk.DoubleVar(value=6.0)
        self.global_line_width = tk.DoubleVar(value=1.0)
        self.left_axis_color = tk.StringVar(value="#000000")
        self.right_axis_color = tk.StringVar(value="#1f77b4")
        self.right2_axis_color = tk.StringVar(value="#d62728")
        self.right2_pos = tk.DoubleVar(value=1.2)

        self.create_widgets()

    def create_widgets(self):
        # top-level layout
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

        container = ttk.Frame(self.root)
        container.grid(row=0, column=0, sticky="nsew")
        container.columnconfigure(0, weight=3)
        container.columnconfigure(1, weight=2)
        container.rowconfigure(0, weight=1)

        # left (controls)
        left_wrap = ScrollableFrame(container)
        left_wrap.container.grid(row=0, column=0, sticky="nsew", padx=(8,4), pady=8)
        self.main_frame = left_wrap.inner
        self.main_frame.columnconfigure(0, weight=1)

        # right (preview)
        right_wrap = ScrollableFrame(container)
        right_wrap.container.grid(row=0, column=1, sticky="nsew", padx=(4,8), pady=8)
        self.preview_parent = right_wrap.inner

        # ========== MODE SWITCHER ==========
        mode_frame = ttk.LabelFrame(self.main_frame, text="Data Source Mode")
        mode_frame.pack(pady=5, fill="x")
        
        ttk.Radiobutton(mode_frame, text="CSV File", variable=self.data_mode, 
                       value='csv', command=self.on_mode_change).pack(side="left", padx=10)
        
        tdms_text = "TDMS Folder (Multiple Files)" if TDMS_AVAILABLE else "TDMS Folder (Not Available)"
        tdms_state = "normal" if TDMS_AVAILABLE else "disabled"
        ttk.Radiobutton(mode_frame, text=tdms_text, variable=self.data_mode, 
                       value='tdms', command=self.on_mode_change, state=tdms_state).pack(side="left", padx=10)
        
        if not TDMS_AVAILABLE:
            ttk.Label(mode_frame, text="Install npTDMS: pip install npTDMS", 
                     foreground="red").pack(side="left", padx=10)

        # ========== CSV/TDMS File Selection ==========
        self.file_frame = ttk.LabelFrame(self.main_frame, text="CSV File Selection")
        self.file_frame.pack(pady=5, fill="x")
        
        # CSV widgets
        self.csv_frame = ttk.Frame(self.file_frame)
        self.csv_frame.pack(fill="x", padx=5, pady=5)
        ttk.Label(self.csv_frame, text="CSV File:").pack(side="left")
        self.file_entry = ttk.Entry(self.csv_frame, width=48)
        self.file_entry.pack(side="left", padx=5, fill="x", expand=True)
        ttk.Button(self.csv_frame, text="Browse", command=self.browse_file).pack(side="left", padx=4)
        
        # TDMS widgets
        self.tdms_frame = ttk.Frame(self.file_frame)
        ttk.Label(self.tdms_frame, text="TDMS Folder:").pack(side="left")
        self.tdms_folder_entry = ttk.Entry(self.tdms_frame, width=48)
        self.tdms_folder_entry.pack(side="left", padx=5, fill="x", expand=True)
        ttk.Button(self.tdms_frame, text="Browse Folder", command=self.browse_tdms_folder).pack(side="left", padx=4)
        
        # Reset button (common)
        ttk.Button(self.file_frame, text="Reset Data Section", command=self.reset_csv).pack(pady=5)
        
        # TDMS file list
        self.tdms_list_frame = ttk.LabelFrame(self.main_frame, text="TDMS Files in Folder (concatenated with 0 hour offset)")
        
        ttk.Label(self.tdms_list_frame, text="Select Files:").grid(row=0, column=0, sticky="nw", padx=5, pady=5)
        self.tdms_files_listbox = tk.Listbox(self.tdms_list_frame, selectmode="multiple", height=8, exportselection=False)
        self.tdms_files_listbox.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        tdms_scroll = ttk.Scrollbar(self.tdms_list_frame, orient="vertical", command=self.tdms_files_listbox.yview)
        tdms_scroll.grid(row=0, column=2, sticky="ns", pady=5)
        self.tdms_files_listbox.configure(yscrollcommand=tdms_scroll.set)
        self.tdms_list_frame.columnconfigure(1, weight=1)
        
        # Buttons for TDMS file selection
        button_frame = ttk.Frame(self.tdms_list_frame)
        button_frame.grid(row=1, column=1, pady=5, sticky="ew")
        ttk.Button(button_frame, text="Select All", command=self.select_all_tdms).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Clear Selection", command=self.clear_tdms_selection).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Load Selected TDMS Files", 
                  command=self.load_selected_tdms).pack(side="left", padx=5)
        
        # Left Y-axis selection
        col_frame = ttk.LabelFrame(self.main_frame, text="Left Y-axis Selection")
        col_frame.pack(pady=5, fill="x")
        col_frame.columnconfigure(1, weight=1)
        ttk.Label(col_frame, text="X-axis:").grid(row=0, column=0, sticky="w")
        self.x_combo = ttk.Combobox(col_frame, state="readonly")
        self.x_combo.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        self.x_combo.bind("<<ComboboxSelected>>", lambda e: self.schedule_preview())
        
        ttk.Label(col_frame, text="Y-axis:").grid(row=1, column=0, sticky="nw")
        self.y_listbox = tk.Listbox(col_frame, selectmode="multiple", height=10, exportselection=False)
        self.y_listbox.grid(row=1, column=1, padx=(5,2), pady=2, sticky="nsew")
        self.y_listbox_scroll = ttk.Scrollbar(col_frame, orient="vertical", command=self.y_listbox.yview)
        self.y_listbox_scroll.grid(row=1, column=2, sticky="ns", padx=(0,4), pady=2)
        self.y_listbox.configure(yscrollcommand=self.y_listbox_scroll.set)
        self.y_listbox.bind("<<ListboxSelect>>", lambda e: self.update_line_properties())
        
        ttk.Button(col_frame, text="Select All", command=self.select_all_left).grid(row=2, column=1, sticky="w", padx=4, pady=(4,2))
        ttk.Button(col_frame, text="Move Up", command=lambda: self.move_selected_in_listbox(self.y_listbox, direction=-1)).grid(row=2, column=1, sticky="e", padx=(0,80), pady=(4,2))
        ttk.Button(col_frame, text="Move Down", command=lambda: self.move_selected_in_listbox(self.y_listbox, direction=1)).grid(row=2, column=1, sticky="e", padx=(0,4), pady=(4,2))
        ttk.Button(col_frame, text="Reset Axis Section", command=self.reset_axis).grid(row=3, column=1, sticky="e", pady=2, padx=4)
        
        # Left per-channel properties
        self.line_prop_frame = ttk.LabelFrame(self.main_frame, text="Per-Channel Label + Properties (Left Y-axis)")
        self.line_prop_frame.pack(pady=5, fill="both")
        ttk.Button(self.line_prop_frame, text="Reset Line Properties", command=self.reset_line_properties).pack(anchor="e", padx=4, pady=2)
        
        # Right axis 1 selection
        right_col_frame = ttk.LabelFrame(self.main_frame, text="Right Y-axis 1 Selection")
        right_col_frame.pack(pady=5, fill="x")
        right_col_frame.columnconfigure(1, weight=1)
        ttk.Label(right_col_frame, text="Right Y-axis 1 Columns:").grid(row=0, column=0, sticky="nw")
        self.right_y_listbox = tk.Listbox(right_col_frame, selectmode="multiple", height=5, exportselection=False)
        self.right_y_listbox.grid(row=0, column=1, padx=(5,2), pady=2, sticky="nsew")
        self.right_y_listbox_scroll = ttk.Scrollbar(right_col_frame, orient="vertical", command=self.right_y_listbox.yview)
        self.right_y_listbox_scroll.grid(row=0, column=2, sticky="ns", padx=(0,4), pady=2)
        self.right_y_listbox.configure(yscrollcommand=self.right_y_listbox_scroll.set)
        self.right_y_listbox.bind("<<ListboxSelect>>", lambda e: self.update_right_line_properties())
        ttk.Button(right_col_frame, text="Move Up", command=lambda: self.move_selected_in_listbox(self.right_y_listbox, direction=-1)).grid(row=1, column=1, sticky="w", padx=4, pady=(4,2))
        ttk.Button(right_col_frame, text="Move Down", command=lambda: self.move_selected_in_listbox(self.right_y_listbox, direction=1)).grid(row=1, column=1, sticky="e", padx=4, pady=(4,2))
        ttk.Button(right_col_frame, text="Reset Right Y-axis 1 Section", command=self.reset_right_axis).grid(row=2, column=1, sticky="e", pady=2, padx=4)
        
        self.right_line_prop_frame = ttk.LabelFrame(self.main_frame, text="Per-Channel Label + Properties (Right Y-axis 1)")
        self.right_line_prop_frame.pack(pady=5, fill="both")
        ttk.Button(self.right_line_prop_frame, text="Reset Right Line Properties", command=self.reset_right_line_properties).pack(anchor="e", padx=4, pady=2)
        
        # Right axis 2 selection
        right2_col_frame = ttk.LabelFrame(self.main_frame, text="Right Y-axis 2 Selection")
        right2_col_frame.pack(pady=5, fill="x")
        right2_col_frame.columnconfigure(1, weight=1)
        ttk.Label(right2_col_frame, text="Right Y-axis 2 Columns:").grid(row=0, column=0, sticky="nw")
        self.right2_y_listbox = tk.Listbox(right2_col_frame, selectmode="multiple", height=5, exportselection=False)
        self.right2_y_listbox.grid(row=0, column=1, padx=(5,2), pady=2, sticky="nsew")
        self.right2_y_listbox_scroll = ttk.Scrollbar(right2_col_frame, orient="vertical", command=self.right2_y_listbox.yview)
        self.right2_y_listbox_scroll.grid(row=0, column=2, sticky="ns", padx=(0,4), pady=2)
        self.right2_y_listbox.configure(yscrollcommand=self.right2_y_listbox_scroll.set)
        self.right2_y_listbox.bind("<<ListboxSelect>>", lambda e: self.update_right2_line_properties())
        ttk.Button(right2_col_frame, text="Move Up", command=lambda: self.move_selected_in_listbox(self.right2_y_listbox, direction=-1)).grid(row=1, column=1, sticky="w", padx=4, pady=(4,2))
        ttk.Button(right2_col_frame, text="Move Down", command=lambda: self.move_selected_in_listbox(self.right2_y_listbox, direction=1)).grid(row=1, column=1, sticky="e", padx=4, pady=(4,2))
        ttk.Button(right2_col_frame, text="Reset Right Y-axis 2 Section", command=self.reset_right2_axis).grid(row=2, column=1, sticky="e", pady=2, padx=4)
        ttk.Label(right2_col_frame, text="Right2 Position (axes coord):").grid(row=3, column=0, sticky="w", pady=(6,0))
        right2_pos_spin = ttk.Spinbox(right2_col_frame, from_=1.0, to=3.0, increment=0.1, textvariable=self.right2_pos, width=8, command=self.schedule_preview)
        right2_pos_spin.grid(row=3, column=1, sticky="w", padx=(4,2), pady=(6,0))
        right2_pos_spin.bind("<KeyRelease>", lambda e: self.schedule_preview())
        
        self.right2_line_prop_frame = ttk.LabelFrame(self.main_frame, text="Per-Channel Label + Properties (Right Y-axis 2)")
        self.right2_line_prop_frame.pack(pady=5, fill="both")
        ttk.Button(self.right2_line_prop_frame, text="Reset Right2 Line Properties", command=self.reset_right2_line_properties).pack(anchor="e", padx=4, pady=2)
        
        # Axis Colors
        axis_ctrl_frame = ttk.LabelFrame(self.main_frame, text="Axis Colors")
        axis_ctrl_frame.pack(pady=5, fill="x")
        ttk.Label(axis_ctrl_frame, text="Left Axis Color:").grid(row=0, column=0, sticky="w")
        self.left_axis_swatch = tk.Label(axis_ctrl_frame, background=self.left_axis_color.get(), width=2, relief="sunken")
        self.left_axis_swatch.grid(row=0, column=1, sticky="w", padx=(4,2))
        ttk.Button(axis_ctrl_frame, text="Choose", command=lambda: self.choose_axis_color('left', self.left_axis_swatch)).grid(row=0, column=2, padx=4)
        
        ttk.Label(axis_ctrl_frame, text="Right Axis Color:").grid(row=1, column=0, sticky="w")
        self.right_axis_swatch = tk.Label(axis_ctrl_frame, background=self.right_axis_color.get(), width=2, relief="sunken")
        self.right_axis_swatch.grid(row=1, column=1, sticky="w", padx=(4,2))
        ttk.Button(axis_ctrl_frame, text="Choose", command=lambda: self.choose_axis_color('right', self.right_axis_swatch)).grid(row=1, column=2, padx=4)
        
        ttk.Label(axis_ctrl_frame, text="Right2 Axis Color:").grid(row=2, column=0, sticky="w")
        self.right2_axis_swatch = tk.Label(axis_ctrl_frame, background=self.right2_axis_color.get(), width=2, relief="sunken")
        self.right2_axis_swatch.grid(row=2, column=1, sticky="w", padx=(4,2))
        ttk.Button(axis_ctrl_frame, text="Choose", command=lambda: self.choose_axis_color('right2', self.right2_axis_swatch)).grid(row=2, column=2, padx=4)
        ttk.Button(axis_ctrl_frame, text="Reset Axis Colors", command=self.reset_axis_colors).grid(row=0, column=3, rowspan=2, padx=8, sticky="e")
        
        # ------------------ Labels & Legend (scrollable) ------------------
        numeric_width = 7
        
        label_frame = ttk.LabelFrame(self.main_frame, text="Labels, Legend & other Global Settings")
        label_frame.pack(pady=5, fill="x")
        
        label_canvas = tk.Canvas(label_frame, borderwidth=0, highlightthickness=0, height=240)
        label_v_scroll = ttk.Scrollbar(label_frame, orient="vertical", command=label_canvas.yview)
        label_h_scroll = ttk.Scrollbar(label_frame, orient="horizontal", command=label_canvas.xview)
        label_canvas.configure(yscrollcommand=label_v_scroll.set, xscrollcommand=label_h_scroll.set)
        label_v_scroll.pack(side="right", fill="y")
        label_h_scroll.pack(side="bottom", fill="x")
        label_canvas.pack(side="left", fill="both", expand=True)
        label_inner = ttk.Frame(label_canvas)
        label_window = label_canvas.create_window((0, 0), window=label_inner, anchor="nw")
        label_inner.bind("<Configure>", lambda e: label_canvas.configure(scrollregion=label_canvas.bbox("all")))
        label_canvas.bind("<Configure>", lambda e: label_canvas.itemconfig(label_window, width=label_canvas.winfo_width()))
        label_canvas.bind_all("<MouseWheel>", lambda ev: self._section_mousewheel(ev, label_canvas, label_inner), add="+")
        
        for i in range(10):
            label_inner.columnconfigure(i, weight=1 if i in (1,3,5,7,9) else 0)
        
        ttk.Label(label_inner, text="Title:").grid(row=0, column=0, sticky="w")
        self.title_entry = ttk.Entry(label_inner, width=20); self.title_entry.grid(row=0, column=1, sticky="ew", padx=2)
        self.title_entry.bind("<KeyRelease>", lambda e: self.schedule_preview())
        
        ttk.Label(label_inner, text="X-label:").grid(row=0, column=2, sticky="w")
        self.xlabel_entry = ttk.Entry(label_inner, width=15); self.xlabel_entry.grid(row=0, column=3, sticky="ew", padx=2)
        self.xlabel_entry.bind("<KeyRelease>", lambda e: self.schedule_preview())
        
        ttk.Label(label_inner, text="Left Y-label:").grid(row=0, column=4, sticky="w")
        self.ylabel_entry = ttk.Entry(label_inner, width=15); self.ylabel_entry.grid(row=0, column=5, sticky="ew", padx=2)
        self.ylabel_entry.bind("<KeyRelease>", lambda e: self.schedule_preview())
        
        ttk.Label(label_inner, text="Right Y-label:").grid(row=0, column=6, sticky="w")
        self.right_ylabel_entry = ttk.Entry(label_inner, width=15); self.right_ylabel_entry.grid(row=0, column=7, sticky="ew", padx=2)
        self.right_ylabel_entry.bind("<KeyRelease>", lambda e: self.schedule_preview())
        
        ttk.Label(label_inner, text="Right2 Y-label:").grid(row=0, column=8, sticky="w")
        self.right2_ylabel_entry = ttk.Entry(label_inner, width=15); self.right2_ylabel_entry.grid(row=0, column=9, sticky="ew", padx=2)
        self.right2_ylabel_entry.bind("<KeyRelease>", lambda e: self.schedule_preview())
        
        ttk.Label(label_inner, text="Left Legend Pos:").grid(row=1, column=0, sticky="w")
        self.legend_loc_left = ttk.Combobox(label_inner, values=[
            "best","upper right","upper left","lower right","lower left","center",
            "outside top","outside bottom","outside left","outside right"], state="readonly")
        self.legend_loc_left.current(0); self.legend_loc_left.grid(row=1, column=1, sticky="ew", padx=2)
        self.legend_loc_left.bind("<<ComboboxSelected>>", lambda e: self.schedule_preview())
        
        ttk.Label(label_inner, text="Right Legend Pos:").grid(row=1, column=2, sticky="w")
        self.legend_loc_right = ttk.Combobox(label_inner, values=[
            "best","upper right","upper left","lower right","lower left","center",
            "outside top","outside bottom","outside left","outside right"], state="readonly")
        self.legend_loc_right.current(0); self.legend_loc_right.grid(row=1, column=3, sticky="ew", padx=2)
        self.legend_loc_right.bind("<<ComboboxSelected>>", lambda e: self.schedule_preview())
        
        ttk.Label(label_inner, text="Right2 Legend Pos:").grid(row=1, column=4, sticky="w")
        self.legend_loc_right2 = ttk.Combobox(label_inner, values=[
            "best","upper right","upper left","lower right","lower left","center",
            "outside top","outside bottom","outside left","outside right"], state="readonly")
        self.legend_loc_right2.current(0); self.legend_loc_right2.grid(row=1, column=5, sticky="ew", padx=2)
        self.legend_loc_right2.bind("<<ComboboxSelected>>", lambda e: self.schedule_preview())
        
        ttk.Label(label_inner, text="Left Legend Columns:").grid(row=2, column=0, sticky="w")
        self.legend_cols_left = tk.IntVar(value=1)
        ttk.Spinbox(label_inner, from_=1, to=10, textvariable=self.legend_cols_left, width=5, command=self.schedule_preview).grid(row=2, column=1, sticky="w", padx=2)
        ttk.Label(label_inner, text="Right Legend Columns:").grid(row=2, column=2, sticky="w")
        self.legend_cols_right = tk.IntVar(value=1)
        ttk.Spinbox(label_inner, from_=1, to=10, textvariable=self.legend_cols_right, width=5, command=self.schedule_preview).grid(row=2, column=3, sticky="w", padx=2)
        ttk.Label(label_inner, text="Right2 Legend Columns:").grid(row=2, column=4, sticky="w")
        self.legend_cols_right2 = tk.IntVar(value=1)
        ttk.Spinbox(label_inner, from_=1, to=10, textvariable=self.legend_cols_right2, width=5, command=self.schedule_preview).grid(row=2, column=5, sticky="w", padx=2)
        
        ttk.Label(label_inner, text="Left Legend X:").grid(row=3, column=0, sticky="w")
        self.legend_x_left = tk.StringVar(value=""); ttk.Entry(label_inner, textvariable=self.legend_x_left, width=numeric_width, justify="center").grid(row=3, column=1, sticky="w", padx=2)
        self.legend_x_left.trace_add("write", lambda *a: self.schedule_preview())
        ttk.Label(label_inner, text="Left Legend Y:").grid(row=3, column=2, sticky="w")
        self.legend_y_left = tk.StringVar(value=""); ttk.Entry(label_inner, textvariable=self.legend_y_left, width=numeric_width, justify="center").grid(row=3, column=3, sticky="w", padx=2)
        self.legend_y_left.trace_add("write", lambda *a: self.schedule_preview())
        
        ttk.Label(label_inner, text="Right Legend X:").grid(row=3, column=4, sticky="w")
        self.legend_x_right = tk.StringVar(value=""); ttk.Entry(label_inner, textvariable=self.legend_x_right, width=numeric_width, justify="center").grid(row=3, column=5, sticky="w", padx=2)
        self.legend_x_right.trace_add("write", lambda *a: self.schedule_preview())
        
        ttk.Label(label_inner, text="Right Legend Y:").grid(row=3, column=6, sticky="w")
        self.legend_y_right = tk.StringVar(value=""); ttk.Entry(label_inner, textvariable=self.legend_y_right, width=numeric_width, justify="center").grid(row=3, column=7, sticky="w", padx=2)
        self.legend_y_right.trace_add("write", lambda *a: self.schedule_preview())
        
        ttk.Label(label_inner, text="Right2 Legend X:").grid(row=4, column=0, sticky="w")
        self.legend_x_right2 = tk.StringVar(value=""); ttk.Entry(label_inner, textvariable=self.legend_x_right2, width=numeric_width, justify="center").grid(row=4, column=1, sticky="w", padx=2)
        self.legend_x_right2.trace_add("write", lambda *a: self.schedule_preview())
        ttk.Label(label_inner, text="Right2 Legend Y:").grid(row=4, column=2, sticky="w")
        self.legend_y_right2 = tk.StringVar(value=""); ttk.Entry(label_inner, textvariable=self.legend_y_right2, width=numeric_width, justify="center").grid(row=4, column=3, sticky="w", padx=2)
        self.legend_y_right2.trace_add("write", lambda *a: self.schedule_preview())
        
        ttk.Label(label_inner, text="Global Marker Size:").grid(row=2, column=6, sticky="w")
        marker_spin = ttk.Spinbox(label_inner, from_=0.1, to=50.0, increment=0.5, textvariable=self.marker_size, width=6, command=self.schedule_preview)
        marker_spin.grid(row=2, column=7, sticky="w", padx=2); marker_spin.bind("<KeyRelease>", lambda e: self.schedule_preview())
        
        ttk.Label(label_inner, text="Global Line Width:").grid(row=3, column=8, sticky="w")
        global_width_spin = ttk.Spinbox(label_inner, from_=0.1, to=20.0, increment=0.1, textvariable=self.global_line_width, width=6, command=self.schedule_preview)
        global_width_spin.grid(row=3, column=9, sticky="w", padx=2); global_width_spin.bind("<KeyRelease>", lambda e: self.schedule_preview())
        
        ttk.Label(label_inner, text="Global Font Size:").grid(row=1, column=6, sticky="w")
        self.font_size = tk.IntVar(value=12)
        ttk.Spinbox(label_inner, from_=8, to=24, textvariable=self.font_size, width=6, command=self.schedule_preview).grid(row=1, column=7, sticky="w", padx=2)
        self.grid_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(label_inner, text="Show Grid", variable=self.grid_var, command=self.schedule_preview).grid(row=1, column=8, sticky="w", padx=2)
        ttk.Button(label_inner, text="Reset Labels Section", command=self.reset_labels).grid(row=1, column=9, sticky="e", padx=2)
        
        # ------------------ Axis Limits (scrollable) ------------------
        limit_frame = ttk.LabelFrame(self.main_frame, text="Axis Limits (optional)")
        limit_frame.pack(pady=5, fill="x")
        
        limit_canvas = tk.Canvas(limit_frame, borderwidth=0, highlightthickness=0, height=140)
        limit_v_scroll = ttk.Scrollbar(limit_frame, orient="vertical", command=limit_canvas.yview)
        limit_h_scroll = ttk.Scrollbar(limit_frame, orient="horizontal", command=limit_canvas.xview)
        limit_canvas.configure(yscrollcommand=limit_v_scroll.set, xscrollcommand=limit_h_scroll.set)
        limit_v_scroll.pack(side="right", fill="y")
        limit_h_scroll.pack(side="bottom", fill="x")
        limit_canvas.pack(side="left", fill="both", expand=True)
        limit_inner = ttk.Frame(limit_canvas)
        limit_window = limit_canvas.create_window((0, 0), window=limit_inner, anchor="nw")
        limit_inner.bind("<Configure>", lambda e: limit_canvas.configure(scrollregion=limit_canvas.bbox("all")))
        limit_canvas.bind("<Configure>", lambda e: limit_canvas.itemconfig(limit_window, width=limit_canvas.winfo_width()))
        limit_canvas.bind_all("<MouseWheel>", lambda ev: self._section_mousewheel(ev, limit_canvas, limit_inner), add="+")
        
        for i in range(18):
            limit_inner.columnconfigure(i, weight=1 if i % 2 == 1 else 0)
        
        ttk.Label(limit_inner, text="X-min:").grid(row=0, column=0)
        self.xmin_entry = ttk.Entry(limit_inner, width=numeric_width, justify="center"); self.xmin_entry.grid(row=0, column=1); self.xmin_entry.bind("<KeyRelease>", lambda e: self.schedule_preview())
        ttk.Label(limit_inner, text="X-max:").grid(row=0, column=2)
        self.xmax_entry = ttk.Entry(limit_inner, width=numeric_width, justify="center"); self.xmax_entry.grid(row=0, column=3); self.xmax_entry.bind("<KeyRelease>", lambda e: self.schedule_preview())
        ttk.Label(limit_inner, text="Left Y-min:").grid(row=0, column=4)
        self.ymin_entry = ttk.Entry(limit_inner, width=numeric_width, justify="center"); self.ymin_entry.grid(row=0, column=5); self.ymin_entry.bind("<KeyRelease>", lambda e: self.schedule_preview())
        ttk.Label(limit_inner, text="Left Y-max:").grid(row=0, column=6)
        self.ymax_entry = ttk.Entry(limit_inner, width=numeric_width, justify="center"); self.ymax_entry.grid(row=0, column=7); self.ymax_entry.bind("<KeyRelease>", lambda e: self.schedule_preview())
        ttk.Label(limit_inner, text="Right Y-min:").grid(row=0, column=8)
        self.right_ymin_entry = ttk.Entry(limit_inner, width=numeric_width, justify="center"); self.right_ymin_entry.grid(row=0, column=9); self.right_ymin_entry.bind("<KeyRelease>", lambda e: self.schedule_preview())
        ttk.Label(limit_inner, text="Right Y-max:").grid(row=0, column=10)
        self.right_ymax_entry = ttk.Entry(limit_inner, width=numeric_width, justify="center"); self.right_ymax_entry.grid(row=0, column=11); self.right_ymax_entry.bind("<KeyRelease>", lambda e: self.schedule_preview())
        
        ttk.Label(limit_inner, text="Right2 Y-min:").grid(row=0, column=12)
        self.right2_ymin_entry = ttk.Entry(limit_inner, width=numeric_width, justify="center"); self.right2_ymin_entry.grid(row=0, column=13); self.right2_ymin_entry.bind("<KeyRelease>", lambda e: self.schedule_preview())
        ttk.Label(limit_inner, text="Right2 Y-max:").grid(row=0, column=14)
        self.right2_ymax_entry = ttk.Entry(limit_inner, width=numeric_width, justify="center"); self.right2_ymax_entry.grid(row=0, column=15); self.right2_ymax_entry.bind("<KeyRelease>", lambda e: self.schedule_preview())
        
        ttk.Label(limit_inner, text="X-interval:").grid(row=1, column=0)
        self.xinterval_entry = ttk.Entry(limit_inner, width=numeric_width, justify="center"); self.xinterval_entry.grid(row=1, column=1); self.xinterval_entry.bind("<KeyRelease>", lambda e: self.schedule_preview())
        ttk.Label(limit_inner, text="Left Y-interval:").grid(row=1, column=2)
        self.yinterval_entry = ttk.Entry(limit_inner, width=numeric_width, justify="center"); self.yinterval_entry.grid(row=1, column=3); self.yinterval_entry.bind("<KeyRelease>", lambda e: self.schedule_preview())
        ttk.Label(limit_inner, text="Right Y-interval:").grid(row=1, column=4)
        self.right_yinterval_entry = ttk.Entry(limit_inner, width=numeric_width, justify="center"); self.right_yinterval_entry.grid(row=1, column=5); self.right_yinterval_entry.bind("<KeyRelease>", lambda e: self.schedule_preview())
        ttk.Label(limit_inner, text="Right2 Y-interval:").grid(row=1, column=6)
        self.right2_yinterval_entry = ttk.Entry(limit_inner, width=numeric_width, justify="center"); self.right2_yinterval_entry.grid(row=1, column=7); self.right2_yinterval_entry.bind("<KeyRelease>", lambda e: self.schedule_preview())
        
        ttk.Button(limit_inner, text="Reset Axis Limits Section", command=self.reset_limits).grid(row=0, column=16, padx=5)
        
        # ------------------ Preview area ------------------
        preview_frame = ttk.LabelFrame(self.preview_parent, text="Plot Preview")
        preview_frame.pack(pady=5, fill="both", expand=True)
        self.preview_canvas_container = tk.Frame(preview_frame)
        self.preview_canvas_container.pack(fill="both", expand=True)
        
        self.fig = plt.Figure(figsize=(self.default_plot_width, self.default_plot_height))
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.preview_canvas_container)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.pack(side="left", fill="both", expand=True)
        
        controls_frame = ttk.Frame(preview_frame)
        controls_frame.pack(anchor="w", pady=2, fill="x")
        ttk.Checkbutton(controls_frame, text="Enable Live Preview", variable=self.live_preview_var, command=self._on_live_toggle).grid(row=0, column=0, padx=(2,8), sticky="w")
        ttk.Checkbutton(controls_frame, text="Downsample Live Preview", variable=self.downsample_live_var, command=self.schedule_preview).grid(row=0, column=1, padx=(2,8), sticky="w")
        ttk.Label(controls_frame, text="Max pts/series:").grid(row=0, column=2, padx=(2,2))
        self.max_pts_spin = ttk.Spinbox(controls_frame, from_=100, to=1000000, increment=100, textvariable=self.max_preview_points, width=8, command=self.schedule_preview)
        self.max_pts_spin.grid(row=0, column=3, padx=(2,8)); self.max_pts_spin.bind("<KeyRelease>", lambda e: self.schedule_preview())
        ttk.Button(controls_frame, text="Preview Plot", command=lambda: self._explicit_preview()).grid(row=0, column=4, padx=(2,8))
        ttk.Button(controls_frame, text="Cancel Preview", command=self.cancel_preview).grid(row=0, column=5, padx=(2,8))
        
        size_frame = ttk.Frame(preview_frame); size_frame.pack(anchor="w", pady=2, fill="x")
        ttk.Label(size_frame, text="Plot Width (inches):").grid(row=0, column=0)
        self.plot_width = tk.DoubleVar(value=self.default_plot_width)
        w_plot = ttk.Entry(size_frame, textvariable=self.plot_width, width=6); w_plot.grid(row=0, column=1); w_plot.bind("<KeyRelease>", lambda e: self.schedule_preview())
        ttk.Label(size_frame, text="Plot Height (inches):").grid(row=0, column=2)
        self.plot_height = tk.DoubleVar(value=self.default_plot_height)
        h_plot = ttk.Entry(size_frame, textvariable=self.plot_height, width=6); h_plot.grid(row=0, column=3); h_plot.bind("<KeyRelease>", lambda e: self.schedule_preview())
        ttk.Button(size_frame, text="Preview Plot", command=self.preview_plot).grid(row=0, column=4, padx=6)
        
        # Configuration & Save
        config_frame = ttk.LabelFrame(self.preview_parent, text="Configuration"); config_frame.pack(pady=5, fill="x")
        ttk.Button(config_frame, text="Save Configuration", command=self.save_configuration).pack(side="left", padx=4)
        ttk.Button(config_frame, text="Load Configuration", command=self.load_configuration).pack(side="left", padx=4)
        ttk.Button(config_frame, text="Reset All", command=self.reset_all).pack(side="right", padx=4)
        
        save_frame = ttk.LabelFrame(self.preview_parent, text="Save Plot"); save_frame.pack(pady=5, fill="x")
        ttk.Label(save_frame, text="Select PNG Quality (DPI):").pack(side="left")
        self.dpi_option = ttk.Combobox(save_frame, values=[100,150,200,300,600], state="readonly"); self.dpi_option.current(2); self.dpi_option.pack(side="left", padx=5)
        ttk.Button(save_frame, text="Save PNG", command=self.save_png).pack(side="left", padx=5)
        
        # Initial mode setup
        self.on_mode_change()
        
        # ========== MODE SWITCHING ==========
    def on_mode_change(self):
        """Handle switching between CSV and TDMS modes"""
        mode = self.data_mode.get()
        
        if mode == 'csv':
            self.file_frame.config(text="CSV File Selection")
            self.csv_frame.pack(fill="x", padx=5, pady=5)
            self.tdms_frame.pack_forget()
            self.tdms_list_frame.pack_forget()
        else:  # tdms
            self.file_frame.config(text="TDMS Folder Selection")
            self.csv_frame.pack_forget()
            self.tdms_frame.pack(fill="x", padx=5, pady=5)
            self.tdms_list_frame.pack(after=self.file_frame, pady=5, fill="x")
        
        # Clear current data
        self.df = None
        self.reset_csv()
    
    # ========== TDMS FUNCTIONS ==========
    def browse_tdms_folder(self):
        """Browse for a folder containing TDMS files"""
        if not TDMS_AVAILABLE:
            messagebox.showerror("Error", "npTDMS library not installed.\nInstall with: pip install npTDMS")
            return
            
        folder_path = filedialog.askdirectory(title="Select Folder Containing TDMS Files")
        if folder_path:
            self.tdms_folder_entry.delete(0, "end")
            self.tdms_folder_entry.insert(0, folder_path)
            self.scan_tdms_folder(folder_path)
    
    def scan_tdms_folder(self, folder_path):
        """Scan folder for TDMS files and populate listbox"""
        try:
            self.tdms_folder = Path(folder_path)
            self.tdms_files = sorted(self.tdms_folder.glob("*.tdms"))
            
            # Update listbox
            self.tdms_files_listbox.delete(0, tk.END)
            for tdms_file in self.tdms_files:
                self.tdms_files_listbox.insert(tk.END, tdms_file.name)
            
            if not self.tdms_files:
                messagebox.showwarning("No TDMS Files", f"No .tdms files found in {folder_path}")
            else:
                messagebox.showinfo("TDMS Files Found", f"Found {len(self.tdms_files)} TDMS file(s)")
                # Auto-select all files
                for i in range(len(self.tdms_files)):
                    self.tdms_files_listbox.selection_set(i)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to scan folder:\n{e}")
    
    def select_all_tdms(self):
        """Select all TDMS files in listbox"""
        self.tdms_files_listbox.selection_set(0, tk.END)
    
    def clear_tdms_selection(self):
        """Clear TDMS file selection"""
        self.tdms_files_listbox.selection_clear(0, tk.END)
    
    def load_selected_tdms(self):
        """Load selected TDMS files and merge into DataFrame"""
        if not TDMS_AVAILABLE:
            messagebox.showerror("Error", "npTDMS library not installed")
            return
            
        selected_indices = self.tdms_files_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("No Selection", "Please select at least one TDMS file")
            return
        
        try:
            selected_files = [self.tdms_files[i] for i in selected_indices]
            self.load_tdms_files(selected_files)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load TDMS files:\n{e}")

    def convert_excel_date_column(self, df):
        """
        Convert Excel date/time format columns to datetime objects for fast plotting.
        Looks for columns with 'date' or 'time' in name and Excel-like numeric values.
        """
        for col in df.columns:
            col_lower = col.lower()
            # Check if column name suggests it's a date/time column
            if any(keyword in col_lower for keyword in ['date', 'time', 'timestamp']):
                try:
                    # Check if the column contains numeric values (Excel date format)
                    if pd.api.types.is_numeric_dtype(df[col]):
                        # Convert Excel date to datetime objects (keep as datetime, not string)
                        # Excel stores dates as days since December 30, 1899
                        df[col + '_Converted'] = pd.to_datetime(
                            EXCEL_EPOCH + pd.to_timedelta(df[col], 'D')
                        )
                        # Keep as datetime64 for fast plotting, don't convert to string
                except Exception as e:
                    # If conversion fails, skip this column
                    print(f"Could not convert column {col}: {e}")
                    continue
        return df
    
    def load_tdms_files(self, tdms_file_paths):
        """
        Load multiple TDMS files and merge them into a single DataFrame with time offset.
        Assumes all files have the same channel structure.
        Files are concatenated with specified time offset between them.
        """
        all_dataframes = []
        file_info = []
        
        # Get time offset in hours (hardcoded to 0)
        time_offset_hours = self.tdms_time_offset_hours
        
        for file_idx, file_path in enumerate(tdms_file_paths):
            try:
                tdms_file = TdmsFile.read(file_path)
                
                # Get all groups and channels
                data_dict = {}
                time_data = None
                
                for group in tdms_file.groups():
                    group_name = group.name
                    
                    for channel in group.channels():
                        channel_name = channel.name
                        # Create unique column name: Group/Channel
                        col_name = f"{group_name}/{channel_name}"
                        
                        # Get channel data
                        data = channel[:]
                        
                        # Check if this is a time channel
                        if 'time' in channel_name.lower() or 'timestamp' in channel_name.lower():
                            time_data = data
                        
                        data_dict[col_name] = data
                
                # Create DataFrame from this file
                if data_dict:
                    df_file = pd.DataFrame(data_dict)
                    
                    # Convert Excel date/time columns if present
                    df_file = self.convert_excel_date_column(df_file)
                    
                    # Add time offset for files after the first one
                    if file_idx > 0:
                        # Calculate time offset in seconds for each subsequent file
                        offset_seconds = file_idx * time_offset_hours * 3600
                        
                        # If there's a time column, add offset to it
                        time_cols = [col for col in df_file.columns if 'time' in col.lower() or 'timestamp' in col.lower()]
                        for time_col in time_cols:
                            try:
                                df_file[time_col] = df_file[time_col] + offset_seconds
                            except:
                                pass
                        
                        # If no time column found, we'll create a continuous index later
                    
                    all_dataframes.append(df_file)
                    file_info.append({
                        'name': file_path.name,
                        'rows': len(df_file),
                        'offset_hours': file_idx * time_offset_hours
                    })
                    
            except Exception as e:
                messagebox.showwarning("File Error", f"Could not load {file_path.name}:\n{e}")
                continue
        
        if not all_dataframes:
            messagebox.showerror("Error", "No data could be loaded from selected files")
            return
        
        # Concatenate all dataframes
        try:
            self.df = pd.concat(all_dataframes, ignore_index=True)
            
            # Convert any Excel date columns in the final dataframe
            self.df = self.convert_excel_date_column(self.df)
            
            # Check if there's a time column
            time_cols = [col for col in self.df.columns if 'time' in col.lower() or 'timestamp' in col.lower()]
            
            if not time_cols:
                # No time column exists, create a continuous time index
                total_rows = len(self.df)
                
                # Assuming uniform sampling, create time array
                # If you know the sampling rate, adjust accordingly
                # For now, we'll create an index that accounts for the time offset
                
                cumulative_time = []
                for idx, df_file in enumerate(all_dataframes):
                    offset_seconds = idx * time_offset_hours * 3600
                    # Create time array for this file (assuming 1 second intervals)
                    file_time = np.arange(len(df_file)) + offset_seconds
                    cumulative_time.extend(file_time)
                
                self.df.insert(0, 'Time_Continuous', cumulative_time)
            
            # If still no usable time/index column, create simple index
            if 'Index' not in self.df.columns and not time_cols and 'Time_Continuous' not in self.df.columns:
                self.df.insert(0, 'Index', range(len(self.df)))
            
            # Populate UI
            self.populate_ui_from_dataframe()
            
            # Create info message
            info_msg = f"Loaded {len(tdms_file_paths)} TDMS file(s)\n"
            info_msg += f"Total rows: {len(self.df)}\n"
            info_msg += f"Channels: {len(self.df.columns)}\n"
            info_msg += f"Time offset: {time_offset_hours} hours between files\n\n"
            info_msg += "File details:\n"
            for info in file_info:
                info_msg += f"  {info['name']}: {info['rows']} rows, offset: {info['offset_hours']:.1f}h\n"
            
            messagebox.showinfo("Success", info_msg)
            
            self.schedule_preview()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to merge TDMS data:\n{e}")
    
    def populate_ui_from_dataframe(self):
        """Populate combo boxes and listboxes from loaded DataFrame"""
        if self.df is None:
            return
            
        # Populate column selectors
        columns = list(self.df.columns)
        self.x_combo["values"] = columns
        
        # Clear listboxes
        self.y_listbox.delete(0, "end")
        self.right_y_listbox.delete(0, "end")
        self.right2_y_listbox.delete(0, "end")
        
        # Populate listboxes
        for col in columns:
            self.y_listbox.insert("end", col)
            self.right_y_listbox.insert("end", col)
            self.right2_y_listbox.insert("end", col)
        
        # Auto-select time/index as X-axis
        # Prioritize converted datetime columns
        converted_cols = [col for col in columns if col.endswith('_Converted')]
        
        if converted_cols:
            self.x_combo.set(converted_cols[0])
        elif 'Time_Continuous' in columns:
            self.x_combo.set('Time_Continuous')
        elif 'Index' in columns:
            self.x_combo.set('Index')
        else:
            # Look for any time-related column
            time_cols = [col for col in columns if 'time' in col.lower() or 'timestamp' in col.lower()]
            if time_cols:
                self.x_combo.set(time_cols[0])
            elif columns:
                self.x_combo.current(0)
        
        # Clear property frames
        self.reset_line_properties()
        self.reset_right_line_properties()
        self.reset_right2_line_properties()
    
    # small helper for section-specific scrolling
    def _section_mousewheel(self, event, canvas_widget, inner_frame):
        try:
            if canvas_widget.winfo_height() < inner_frame.winfo_height():
                canvas_widget.yview_scroll(int(-1 * (event.delta / 120)), "units")
            else:
                if canvas_widget.winfo_width() < inner_frame.winfo_width():
                    canvas_widget.xview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass
    
    # ---------------- Helpers ----------------
    def choose_axis_color(self, axis, swatch_label):
        color_code = colorchooser.askcolor(title="Choose Axis Color")[1]
        if not color_code:
            return
        if axis == 'left':
            self.left_axis_color.set(color_code)
        elif axis == 'right':
            self.right_axis_color.set(color_code)
        elif axis == 'right2':
            self.right2_axis_color.set(color_code)
        try:
            swatch_label.configure(background=color_code)
        except Exception:
            pass
        self.schedule_preview()
    
    def reset_axis_colors(self):
        self.left_axis_color.set("#000000")
        self.right_axis_color.set("#1f77b4")
        self.right2_axis_color.set("#d62728")
        try:
            self.left_axis_swatch.configure(background=self.left_axis_color.get())
            self.right_axis_swatch.configure(background=self.right_axis_color.get())
            self.right2_axis_swatch.configure(background=self.right2_axis_color.get())
        except Exception:
            pass
        self.schedule_preview()
    
    def choose_marker_color(self, row_frame, swatch_label):
        color_code = colorchooser.askcolor(title="Choose Marker Color")[1]
        if color_code:
            row_frame.marker_color = color_code
            try:
                swatch_label.configure(background=color_code)
            except Exception:
                pass
            self.sync_channel_props_from_frames()
            self.schedule_preview()
    
    def choose_color(self, frame_row, swatch_label):
        color_code = colorchooser.askcolor(title="Choose Line Color")[1]
        if color_code:
            frame_row.line_color = color_code
            try:
                swatch_label.configure(background=color_code)
            except Exception:
                pass
            self.sync_channel_props_from_frames()
            self.schedule_preview()
    
    def select_all_left(self):
        try:
            self.y_listbox.selection_set(0, tk.END)
            self.y_listbox.see(0)
            self.update_line_properties()
        except Exception:
            pass
    
    # ---------------- PEC cycler fallback loader ----------------
    def load_csv_with_dynamic_header(self, file_path, min_consistent_lines=3, min_cols=5):
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.read().splitlines()
    
        # drop leading empty lines
        while lines and lines[0].strip() == "":
            lines.pop(0)
    
        candidate_idx = None
        for i in range(len(lines)):
            parts = lines[i].split(',')
            col_count = len(parts)
            if col_count < min_cols:
                continue
            consistent = True
            for k in range(1, min_consistent_lines):
                j = i + k
                if j >= len(lines):
                    consistent = False
                    break
                parts_next = lines[j].split(',')
                if len(parts_next) != col_count:
                    consistent = False
                    break
            if consistent:
                candidate_idx = i
                break
    
        if candidate_idx is None:
            raise ValueError("Could not auto-detect data header line in PEC-style CSV.")
    
        reconstructed = "\n".join(lines[candidate_idx:])
        sio = io.StringIO(reconstructed)
        df = pd.read_csv(sio)
        self.detected_header_line_index = candidate_idx
        return df
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files","*.csv"), ("All files","*.*")])
        if file_path:
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, file_path)
            self.load_csv(file_path)
    
    def load_csv(self, file_path):
        self.detected_header_line_index = None
        try:
            self.df = pd.read_csv(file_path)
            # Convert any Excel date columns
            self.df = self.convert_excel_date_column(self.df)
        except Exception as e:
            if "Error tokenizing data" in str(e) or "expected" in str(e):
                try:
                    self.df = self.load_csv_with_dynamic_header(file_path)
                    # Convert any Excel date columns
                    self.df = self.convert_excel_date_column(self.df)
                    messagebox.showinfo("PEC CSV Loaded", 
                        f"Loaded PEC-style CSV starting from detected header line {self.detected_header_line_index}.\n"
                        f"File: {os.path.basename(file_path)}")
                except Exception as e2:
                    messagebox.showerror("Error", 
                        f"Failed to parse CSV:\n{e}\nFallback parsing also failed:\n{e2}")
                    return
            else:
                messagebox.showerror("Error", str(e))
                return
    
        # populate UI
        self.populate_ui_from_dataframe()
        self.schedule_preview()
    
    # ---------------- Channel props sync helpers ----------------
    def sync_channel_props_from_frames(self):
        try:
            for row in self.line_properties_frames:
                col = getattr(row, "_col_name", None)
                if not col:
                    continue
                try:
                    self.left_channel_props[col] = {
                        "label": row._linked_label_var.get() if hasattr(row, "_linked_label_var") else col,
                        "line_color": getattr(row, "line_color", ""),
                        "marker_color": getattr(row, "marker_color", "") or "",
                        "style": row.line_style.get() if hasattr(row, "line_style") else "",
                        "marker": row.marker.get() if hasattr(row, "marker") else "",
                        "scatter": row.scatter_mode.get() if hasattr(row, "scatter_mode") else False
                    }
                except Exception:
                    continue
            for row in self.right_line_properties_frames:
                col = getattr(row, "_col_name", None)
                if not col:
                    continue
                try:
                    self.right_channel_props[col] = {
                        "label": row._linked_label_var.get() if hasattr(row, "_linked_label_var") else col,
                        "line_color": getattr(row, "line_color", ""),
                        "marker_color": getattr(row, "marker_color", "") or "",
                        "style": row.line_style.get() if hasattr(row, "line_style") else "",
                        "marker": row.marker.get() if hasattr(row, "marker") else "",
                        "scatter": row.scatter_mode.get() if hasattr(row, "scatter_mode") else False
                    }
                except Exception:
                    continue
            for row in self.right2_line_properties_frames:
                col = getattr(row, "_col_name", None)
                if not col:
                    continue
                try:
                    self.right2_channel_props[col] = {
                        "label": row._linked_label_var.get() if hasattr(row, "_linked_label_var") else col,
                        "line_color": getattr(row, "line_color", ""),
                        "marker_color": getattr(row, "marker_color", "") or "",
                        "style": row.line_style.get() if hasattr(row, "line_style") else "",
                        "marker": row.marker.get() if hasattr(row, "marker") else "",
                        "scatter": row.scatter_mode.get() if hasattr(row, "scatter_mode") else False
                    }
                except Exception:
                    continue
        except Exception:
            pass
    
    def apply_props_to_row(self, row, col, props_map):
        try:
            row._col_name = col
            p = props_map.get(col, {})
            if 'label' in p and hasattr(row, "_linked_label_var"):
                try:
                    row._linked_label_var.set(p.get('label', row._linked_label_var.get()))
                except Exception:
                    pass
            if 'line_color' in p and p.get('line_color'):
                try:
                    row.line_color = p.get('line_color')
                    if hasattr(row, "_swatch") and row._swatch:
                        row._swatch.configure(background=row.line_color)
                except Exception:
                    pass
            mc = p.get('marker_color', "")
            if mc:
                try:
                    row.marker_color = mc
                    if hasattr(row, "_marker_swatch") and row._marker_swatch:
                        row._marker_swatch.configure(background=mc)
                except Exception:
                    pass
            if 'style' in p and hasattr(row, 'line_style'):
                try:
                    row.line_style.set(p.get('style', row.line_style.get()))
                except Exception:
                    pass
            if 'marker' in p and hasattr(row, 'marker'):
                try:
                    row.marker.set(p.get('marker', row.marker.get()))
                except Exception:
                    pass
            if 'scatter' in p and hasattr(row, 'scatter_mode'):
                try:
                    row.scatter_mode.set(p.get('scatter', False))
                except Exception:
                    pass
        except Exception:
            pass
    
    # ---------------- Move selected items with preservation ----------------
    def move_selected_in_listbox(self, listbox, direction=1):
        try:
            self.sync_channel_props_from_frames()
            items = list(listbox.get(0, tk.END))
            sel = list(listbox.curselection())
            if not sel:
                return
            n = len(items)
            if direction < 0:
                for idx in sel:
                    if idx == 0:
                        continue
                    items[idx-1], items[idx] = items[idx], items[idx-1]
                new_sel = [max(0, i-1) for i in sel]
            else:
                for idx in reversed(sel):
                    if idx >= n-1:
                        continue
                    items[idx+1], items[idx] = items[idx], items[idx+1]
                new_sel = [min(n-1, i+1) for i in sel]
            listbox.delete(0, tk.END)
            for it in items:
                listbox.insert(tk.END, it)
            listbox.selection_clear(0, tk.END)
            for i in new_sel:
                listbox.selection_set(i)
            if listbox is self.y_listbox:
                self.update_line_properties()
            elif listbox is self.right_y_listbox:
                self.update_right_line_properties()
            elif listbox is self.right2_y_listbox:
                self.update_right2_line_properties()
        except Exception:
            pass
        
        # ---------------- Update Line Properties ----------------
    def update_line_properties(self):
        self.sync_channel_props_from_frames()
        self.reset_line_properties()
        selected_indices = self.y_listbox.curselection()
        y_cols = [self.y_listbox.get(i) for i in selected_indices]
    
        for col in y_cols:
            row = ttk.Frame(self.line_prop_frame)
            row.pack(fill="x", pady=1, padx=2)
            row._col_name = col
            ttk.Label(row, text=col, width=20).pack(side="left", padx=2)
            label_var = tk.StringVar(value=col)
            label_entry = ttk.Entry(row, textvariable=label_var, width=30)
            label_entry.pack(side="left", padx=4)
            label_entry.bind("<KeyRelease>", lambda e: (self.sync_channel_props_from_frames(), self.schedule_preview()))
            
            # Scatter plot checkbox
            scatter_var = tk.BooleanVar(value=False)
            scatter_cb = ttk.Checkbutton(row, text="Scatter", variable=scatter_var, 
                                        command=lambda: (self.sync_channel_props_from_frames(), self.schedule_preview()))
            scatter_cb.pack(side="left", padx=(8,4))
            
            style_var = tk.StringVar(value="-")
            ttk.Label(row, text="Style:").pack(side="left", padx=(8,2))
            style_cb = ttk.Combobox(row, values=["-","--","-.",":"], textvariable=style_var, width=4)
            style_cb.pack(side="left")
            style_cb.bind("<<ComboboxSelected>>", lambda e: (self.sync_channel_props_from_frames(), self.schedule_preview()))
            marker_var = tk.StringVar(value="None")
            ttk.Label(row, text="Marker:").pack(side="left", padx=(6,2))
            marker_cb = ttk.Combobox(row, values=["None","o","s","^","*","x","+","d","v","<",">","p","h"], textvariable=marker_var, width=4)
            marker_cb.pack(side="left")
            marker_cb.bind("<<ComboboxSelected>>", lambda e: (self.sync_channel_props_from_frames(), self.schedule_preview()))
            color = next(self.color_cycle)
            swatch = tk.Label(row, background=color, width=2, relief="sunken")
            swatch.pack(side="left", padx=(6,2))
            ttk.Button(row, text="Choose", command=lambda r=row, s=swatch: self.choose_color(r, s)).pack(side="left", padx=(2,6))
            marker_color_swatch = tk.Label(row, background=color, width=2, relief="raised")
            marker_color_swatch.pack(side="left", padx=(4,2))
            row.marker_color = ""
            row._marker_swatch = marker_color_swatch
            ttk.Button(row, text="Marker Color", command=lambda r=row, s=marker_color_swatch: self.choose_marker_color(r, s)).pack(side="left", padx=(2,6))
            row.line_color = color
            row.line_style = style_var
            row.marker = marker_var
            row.scatter_mode = scatter_var
            row._linked_label_var = label_var
            row._swatch = swatch
            self.apply_props_to_row(row, col, self.left_channel_props)
            self.line_properties_frames.append(row)
        self.schedule_preview()
    
    def update_right_line_properties(self):
        self.sync_channel_props_from_frames()
        self.reset_right_line_properties()
        selected_indices = self.right_y_listbox.curselection()
        right_cols = [self.right_y_listbox.get(i) for i in selected_indices]
        for col in right_cols:
            row = ttk.Frame(self.right_line_prop_frame)
            row.pack(fill="x", pady=1, padx=2)
            row._col_name = col
            ttk.Label(row, text=col, width=20).pack(side="left", padx=2)
            label_var = tk.StringVar(value=col)
            label_entry = ttk.Entry(row, textvariable=label_var, width=30)
            label_entry.pack(side="left", padx=4)
            label_entry.bind("<KeyRelease>", lambda e: (self.sync_channel_props_from_frames(), self.schedule_preview()))
            
            # Scatter plot checkbox
            scatter_var = tk.BooleanVar(value=False)
            scatter_cb = ttk.Checkbutton(row, text="Scatter", variable=scatter_var, 
                                        command=lambda: (self.sync_channel_props_from_frames(), self.schedule_preview()))
            scatter_cb.pack(side="left", padx=(8,4))
            
            style_var = tk.StringVar(value="-")
            ttk.Label(row, text="Style:").pack(side="left", padx=(8,2))
            style_cb = ttk.Combobox(row, values=["-","--","-.",":"], textvariable=style_var, width=4)
            style_cb.pack(side="left")
            style_cb.bind("<<ComboboxSelected>>", lambda e: (self.sync_channel_props_from_frames(), self.schedule_preview()))
            marker_var = tk.StringVar(value="None")
            ttk.Label(row, text="Marker:").pack(side="left", padx=(6,2))
            marker_cb = ttk.Combobox(row, values=["None","o","s","^","*","x","+","d","v","<",">","p","h"], textvariable=marker_var, width=4)
            marker_cb.pack(side="left")
            marker_cb.bind("<<ComboboxSelected>>", lambda e: (self.sync_channel_props_from_frames(), self.schedule_preview()))
            color = next(self.right_color_cycle)
            swatch = tk.Label(row, background=color, width=2, relief="sunken")
            swatch.pack(side="left", padx=(6,2))
            ttk.Button(row, text="Choose", command=lambda r=row, s=swatch: self.choose_color(r, s)).pack(side="left", padx=(2,6))
            marker_color_swatch = tk.Label(row, background=color, width=2, relief="raised")
            marker_color_swatch.pack(side="left", padx=(4,2))
            row.marker_color = ""
            row._marker_swatch = marker_color_swatch
            ttk.Button(row, text="Marker Color", command=lambda r=row, s=marker_color_swatch: self.choose_marker_color(r, s)).pack(side="left", padx=(2,6))
            row.line_color = color
            row.line_style = style_var
            row.marker = marker_var
            row.scatter_mode = scatter_var
            row._linked_label_var = label_var
            row._swatch = swatch
            self.apply_props_to_row(row, col, self.right_channel_props)
            self.right_line_properties_frames.append(row)
        self.schedule_preview()
    
    def update_right2_line_properties(self):
        self.sync_channel_props_from_frames()
        self.reset_right2_line_properties()
        selected_indices = self.right2_y_listbox.curselection()
        right2_cols = [self.right2_y_listbox.get(i) for i in selected_indices]
        for col in right2_cols:
            row = ttk.Frame(self.right2_line_prop_frame)
            row.pack(fill="x", pady=1, padx=2)
            row._col_name = col
            ttk.Label(row, text=col, width=20).pack(side="left", padx=2)
            label_var = tk.StringVar(value=col)
            label_entry = ttk.Entry(row, textvariable=label_var, width=30)
            label_entry.pack(side="left", padx=4)
            label_entry.bind("<KeyRelease>", lambda e: (self.sync_channel_props_from_frames(), self.schedule_preview()))
            
            # Scatter plot checkbox
            scatter_var = tk.BooleanVar(value=False)
            scatter_cb = ttk.Checkbutton(row, text="Scatter", variable=scatter_var, 
                                        command=lambda: (self.sync_channel_props_from_frames(), self.schedule_preview()))
            scatter_cb.pack(side="left", padx=(8,4))
            
            style_var = tk.StringVar(value="-")
            ttk.Label(row, text="Style:").pack(side="left", padx=(8,2))
            style_cb = ttk.Combobox(row, values=["-","--","-.",":"], textvariable=style_var, width=4)
            style_cb.pack(side="left")
            style_cb.bind("<<ComboboxSelected>>", lambda e: (self.sync_channel_props_from_frames(), self.schedule_preview()))
            marker_var = tk.StringVar(value="None")
            ttk.Label(row, text="Marker:").pack(side="left", padx=(6,2))
            marker_cb = ttk.Combobox(row, values=["None","o","s","^","*","x","+","d","v","<",">","p","h"], textvariable=marker_var, width=4)
            marker_cb.pack(side="left")
            marker_cb.bind("<<ComboboxSelected>>", lambda e: (self.sync_channel_props_from_frames(), self.schedule_preview()))
            color = next(self.right2_color_cycle)
            swatch = tk.Label(row, background=color, width=2, relief="sunken")
            swatch.pack(side="left", padx=(6,2))
            ttk.Button(row, text="Choose", command=lambda r=row, s=swatch: self.choose_color(r, s)).pack(side="left", padx=(2,6))
            marker_color_swatch = tk.Label(row, background=color, width=2, relief="raised")
            marker_color_swatch.pack(side="left", padx=(4,2))
            row.marker_color = ""
            row._marker_swatch = marker_color_swatch
            ttk.Button(row, text="Marker Color", command=lambda r=row, s=marker_color_swatch: self.choose_marker_color(r, s)).pack(side="left", padx=(2,6))
            row.line_color = color
            row.line_style = style_var
            row.marker = marker_var
            row.scatter_mode = scatter_var
            row._linked_label_var = label_var
            row._swatch = swatch
            self.apply_props_to_row(row, col, self.right2_channel_props)
            self.right2_line_properties_frames.append(row)
        self.schedule_preview()
        
    # ---------------- Preview ----------------
    def _on_live_toggle(self):
        if not self.live_preview_var.get():
            if self._preview_after_id is not None:
                try:
                    self.root.after_cancel(self._preview_after_id)
                except Exception:
                    pass
                self._preview_after_id = None
        else:
            self.schedule_preview()

    def schedule_preview(self):
        if not self.live_preview_var.get():
            return
        if self._preview_after_id is not None:
            try:
                self.root.after_cancel(self._preview_after_id)
            except Exception:
                pass
        self._preview_after_id = self.root.after(DEBOUNCE_MS, self.preview_plot)

    def cancel_preview(self):
        if self._preview_after_id is not None:
            try:
                self.root.after_cancel(self._preview_after_id)
            except Exception:
                pass
            self._preview_after_id = None
        self._preview_cancel_requested = True
        self._explicit_preview_request = False

    def _explicit_preview(self):
        if self._preview_after_id is not None:
            try:
                self.root.after_cancel(self._preview_after_id)
            except Exception:
                pass
            self._preview_after_id = None
        self._explicit_preview_request = True
        self._preview_cancel_requested = False
        self.root.after(0, self.preview_plot)

    def preview_plot(self):
        self._preview_after_id = None
        explicit = getattr(self, "_explicit_preview_request", False)
        self._explicit_preview_request = False

        if self.df is None:
            try:
                self.fig.clf()
                self.canvas.draw()
            except Exception:
                pass
            return

        if self._preview_cancel_requested:
            self._preview_cancel_requested = False
            return

        x_col = self.x_combo.get()
        selected_indices = self.y_listbox.curselection()
        right_indices = self.right_y_listbox.curselection()
        right2_indices = self.right2_y_listbox.curselection()
        if not x_col or (not selected_indices and not right_indices and not right2_indices):
            try:
                self.fig.clf()
                self.canvas.draw()
            except Exception:
                pass
            return

        y_cols = [self.y_listbox.get(i) for i in selected_indices]
        right_cols = [self.right_y_listbox.get(i) for i in right_indices]
        right2_cols = [self.right2_y_listbox.get(i) for i in right2_indices]

        # collect labels from frames
        labels = []
        for i, col in enumerate(y_cols):
            label = col
            if i < len(self.line_properties_frames):
                try:
                    var = self.line_properties_frames[i]._linked_label_var
                    if var.get().strip():
                        label = var.get().strip()
                except Exception:
                    pass
            labels.append(label)

        right_labels = []
        for i, col in enumerate(right_cols):
            label = col
            if i < len(self.right_line_properties_frames):
                try:
                    var = self.right_line_properties_frames[i]._linked_label_var
                    if var.get().strip():
                        label = var.get().strip()
                except Exception:
                    pass
            right_labels.append(label)

        right2_labels = []
        for i, col in enumerate(right2_cols):
            label = col
            if i < len(self.right2_line_properties_frames):
                try:
                    var = self.right2_line_properties_frames[i]._linked_label_var
                    if var.get().strip():
                        label = var.get().strip()
                except Exception:
                    pass
            right2_labels.append(label)

        try:
            pw = float(self.plot_width.get())
            ph = float(self.plot_height.get())
        except Exception:
            pw = self.default_plot_width
            ph = self.default_plot_height

        n_rows = len(self.df.index) if self.df is not None else 0
        do_downsample = (not explicit) and self.downsample_live_var.get() and (n_rows > max(1, self.max_preview_points.get()))

        def get_xy_arrays(col_name):
            try:
                if do_downsample:
                    max_pts = max(1, int(self.max_preview_points.get()))
                    n = n_rows
                    idx = np.linspace(0, n-1, num=min(n, max_pts), dtype=int)
                    x_arr = self.df[self.x_combo.get()].values[idx]
                    y_arr = self.df[col_name].values[idx].astype(float)
                    
                    # Convert datetime64 to matplotlib date format for fast plotting
                    if pd.api.types.is_datetime64_any_dtype(x_arr):
                        x_arr = mdates.date2num(pd.to_datetime(x_arr))
                    
                    return x_arr, y_arr
                else:
                    x_arr = self.df[self.x_combo.get()].values
                    y_arr = self.df[col_name].values.astype(float)
                    
                    # Convert datetime64 to matplotlib date format for fast plotting
                    if pd.api.types.is_datetime64_any_dtype(x_arr):
                        x_arr = mdates.date2num(pd.to_datetime(x_arr))
                    
                    return x_arr, y_arr
            except Exception:
                return np.array([]), np.array([])

        plt.close(self.fig)
        if self._preview_cancel_requested:
            self._preview_cancel_requested = False
            return
        self.fig, self.ax = plt.subplots(figsize=(pw, ph))

        try:
            global_lw = float(self.global_line_width.get())
        except Exception:
            global_lw = 1.0

        # left plots
        for y_col, label, row in zip(y_cols, labels, self.line_properties_frames):
            if self._preview_cancel_requested:
                self._preview_cancel_requested = False
                try:
                    self.fig.clf()
                    self.canvas.draw()
                except Exception:
                    pass
                return
            try:
                # Check if scatter mode is enabled
                is_scatter = hasattr(row, 'scatter_mode') and row.scatter_mode.get()
                
                marker = None if (row.marker.get() == "None") else row.marker.get()
                
                # For scatter mode, force marker if none selected
                if is_scatter and (marker is None or marker == "None"):
                    marker = "o"
                
                x_arr, y_arr = get_xy_arrays(y_col)
                if x_arr.size == 0:
                    continue
                mcolor = getattr(row, "marker_color", "") or row.line_color
                
                # Scatter mode: no line, only markers
                if is_scatter:
                    self.ax.scatter(x_arr, y_arr, s=self.marker_size.get()**2, 
                                   c=mcolor, marker=marker, label=label, edgecolors=mcolor)
                else:
                    self.ax.plot(x_arr, y_arr, linestyle=row.line_style.get(), linewidth=global_lw,
                                 color=row.line_color, marker=marker, markersize=self.marker_size.get(),
                                 markerfacecolor=mcolor, markeredgecolor=mcolor, label=label)
            except Exception:
                continue

        try:
            left_col = self.left_axis_color.get()
            self.ax.spines["left"].set_color(left_col)
            self.ax.yaxis.label.set_color(left_col)
            self.ax.tick_params(axis="y", colors=left_col)
        except Exception:
            pass

        ax2 = None
        if right_cols:
            ax2 = self.ax.twinx()
            for y_col, label, row in zip(right_cols, right_labels, self.right_line_properties_frames):
                if self._preview_cancel_requested:
                    self._preview_cancel_requested = False
                    try:
                        self.fig.clf()
                        self.canvas.draw()
                    except Exception:
                        pass
                    return
                try:
                    # Check if scatter mode is enabled
                    is_scatter = hasattr(row, 'scatter_mode') and row.scatter_mode.get()
                    
                    marker = None if (row.marker.get() == "None") else row.marker.get()
                    
                    # For scatter mode, force marker if none selected
                    if is_scatter and (marker is None or marker == "None"):
                        marker = "o"
                    
                    x_arr, y_arr = get_xy_arrays(y_col)
                    if x_arr.size == 0:
                        continue
                    mcolor = getattr(row, "marker_color", "") or row.line_color
                    
                    # Scatter mode: no line, only markers
                    if is_scatter:
                        ax2.scatter(x_arr, y_arr, s=self.marker_size.get()**2, 
                                   c=mcolor, marker=marker, label=label, edgecolors=mcolor)
                    else:
                        ax2.plot(x_arr, y_arr, linestyle=row.line_style.get(), linewidth=global_lw,
                                 color=row.line_color, marker=marker, markersize=self.marker_size.get(),
                                 markerfacecolor=mcolor, markeredgecolor=mcolor, label=label)
                except Exception:
                    continue
            try:
                right_col = self.right_axis_color.get()
                ax2.spines["right"].set_color(right_col)
                ax2.yaxis.label.set_color(right_col)
                ax2.tick_params(axis="y", colors=right_col)
            except Exception:
                pass
            try:
                ax2.set_ylabel(self.right_ylabel_entry.get() or ", ".join(right_labels), fontsize=self.font_size.get())
            except Exception:
                pass

        ax3 = None
        if right2_cols:
            ax3 = self.ax.twinx()
            try:
                pos = float(self.right2_pos.get())
                ax3.spines["right"].set_position(("axes", pos))
            except Exception:
                pass
            try:
                ax3.set_frame_on(True)
            except Exception:
                pass
            for y_col, label, row in zip(right2_cols, right2_labels, self.right2_line_properties_frames):
                if self._preview_cancel_requested:
                    self._preview_cancel_requested = False
                    try:
                        self.fig.clf()
                        self.canvas.draw()
                    except Exception:
                        pass
                    return
                try:
                    # Check if scatter mode is enabled
                    is_scatter = hasattr(row, 'scatter_mode') and row.scatter_mode.get()
                    
                    marker = None if (row.marker.get() == "None") else row.marker.get()
                    
                    # For scatter mode, force marker if none selected
                    if is_scatter and (marker is None or marker == "None"):
                        marker = "o"
                    
                    x_arr, y_arr = get_xy_arrays(y_col)
                    if x_arr.size == 0:
                        continue
                    mcolor = getattr(row, "marker_color", "") or row.line_color
                    
                    # Scatter mode: no line, only markers
                    if is_scatter:
                        ax3.scatter(x_arr, y_arr, s=self.marker_size.get()**2, 
                                   c=mcolor, marker=marker, label=label, edgecolors=mcolor)
                    else:
                        ax3.plot(x_arr, y_arr, linestyle=row.line_style.get(), linewidth=global_lw,
                                 color=row.line_color, marker=marker, markersize=self.marker_size.get(),
                                 markerfacecolor=mcolor, markeredgecolor=mcolor, label=label)
                except Exception:
                    continue
            try:
                right2_col = self.right2_axis_color.get()
                ax3.spines["right"].set_color(right2_col)
                ax3.yaxis.label.set_color(right2_col)
                ax3.tick_params(axis="y", colors=right2_col)
            except Exception:
                pass
            try:
                ax3.set_ylabel(self.right2_ylabel_entry.get() or ", ".join(right2_labels), fontsize=self.font_size.get())
            except Exception:
                pass

        # axis limits
        try:
            xmin = float(self.xmin_entry.get()) if self.xmin_entry.get() else None
            xmax = float(self.xmax_entry.get()) if self.xmax_entry.get() else None
            ymin = float(self.ymin_entry.get()) if self.ymin_entry.get() else None
            ymax = float(self.ymax_entry.get()) if self.ymax_entry.get() else None
            self.ax.set_xlim(xmin, xmax)
            self.ax.set_ylim(ymin, ymax)
            if ax2:
                rymin = float(self.right_ymin_entry.get()) if self.right_ymin_entry.get() else None
                rymax = float(self.right_ymax_entry.get()) if self.right_ymax_entry.get() else None
                ax2.set_ylim(rymin, rymax)
            if ax3:
                ry2min = float(self.right2_ymin_entry.get()) if self.right2_ymin_entry.get() else None
                ry2max = float(self.right2_ymax_entry.get()) if self.right2_ymax_entry.get() else None
                ax3.set_ylim(ry2min, ry2max)
        except ValueError:
            pass

        try:
            if self.xinterval_entry.get().strip():
                xint = float(self.xinterval_entry.get())
                if xint > 0:
                    self.ax.xaxis.set_major_locator(MultipleLocator(xint))
            if self.yinterval_entry.get().strip():
                yint = float(self.yinterval_entry.get())
                if yint > 0:
                    self.ax.yaxis.set_major_locator(MultipleLocator(yint))
            if ax2 and self.right_yinterval_entry.get().strip():
                ryint = float(self.right_yinterval_entry.get())
                if ryint > 0:
                    ax2.yaxis.set_major_locator(MultipleLocator(ryint))
            if ax3 and self.right2_yinterval_entry.get().strip():
                ry2int = float(self.right2_yinterval_entry.get())
                if ry2int > 0:
                    ax3.yaxis.set_major_locator(MultipleLocator(ry2int))
        except Exception:
            pass

        try:
            if ax3 is not None:
                right_margin = 0.70
            elif ax2 is not None:
                right_margin = 0.82
            else:
                right_margin = 0.92
            right_margin = max(0.5, min(0.95, right_margin))
            self.fig.subplots_adjust(right=right_margin)
        except Exception:
            pass

        fs = self.font_size.get()
        try:
            self.ax.set_title(self.title_entry.get(), fontsize=fs)
            self.ax.set_xlabel(self.xlabel_entry.get() or x_col, fontsize=fs)
            
            # Format x-axis if using converted datetime format
            if x_col.endswith('_Converted'):
                # Use matplotlib date formatter for better performance
                self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y/%m/%d\n%H:%M:%S'))
                self.ax.xaxis.set_major_locator(mdates.AutoDateLocator())
                self.fig.autofmt_xdate(rotation=45)
            self.ax.set_ylabel(self.ylabel_entry.get() or ", ".join(labels), fontsize=fs)
            self.ax.tick_params(axis="both", labelsize=fs)
            if ax2:
                ax2.tick_params(axis="both", labelsize=fs)
            if ax3:
                ax3.tick_params(axis="both", labelsize=fs)
            if self.grid_var.get():
                self.ax.grid(True)
        except Exception:
            pass

        def place_legend(ax, legend_pos, cols, x_override=None, y_override=None):
            outside = False
            loc_map = {
                "outside top": "upper center",
                "outside bottom": "lower center",
                "outside left": "center left",
                "outside right": "center right"
            }
            if legend_pos and legend_pos.startswith("outside"):
                outside = True
                loc = loc_map.get(legend_pos, "best")
            else:
                loc = legend_pos or "best"
            try:
                xo = float(x_override) if (x_override is not None and x_override != "") else None
                yo = float(y_override) if (y_override is not None and y_override != "") else None
            except Exception:
                xo = yo = None
            if xo is not None and yo is not None:
                ax.legend(loc=loc, bbox_to_anchor=(xo, yo), fontsize=fs, ncol=cols)
            elif outside:
                if "top" in legend_pos:
                    ax.legend(loc=loc, bbox_to_anchor=(0.5,1.15), fontsize=fs, ncol=cols)
                elif "bottom" in legend_pos:
                    ax.legend(loc=loc, bbox_to_anchor=(0.5,-0.3), fontsize=fs, ncol=cols)
                elif "left" in legend_pos:
                    ax.legend(loc=loc, bbox_to_anchor=(-0.3,0.5), fontsize=fs, ncol=cols)
                elif "right" in legend_pos:
                    ax.legend(loc=loc, bbox_to_anchor=(1.2,0.5), fontsize=fs, ncol=cols)
                else:
                    ax.legend(loc=loc, fontsize=fs, ncol=cols)
            else:
                ax.legend(loc=loc, fontsize=fs, ncol=cols)

        try:
            place_legend(self.ax, self.legend_loc_left.get(), self.legend_cols_left.get(), 
                        self.legend_x_left.get(), self.legend_y_left.get())
            if ax2:
                place_legend(ax2, self.legend_loc_right.get(), self.legend_cols_right.get(), 
                            self.legend_x_right.get(), self.legend_y_right.get())
            if ax3:
                place_legend(ax3, self.legend_loc_right2.get(), self.legend_cols_right2.get(), 
                            self.legend_x_right2.get(), self.legend_y_right2.get())
        except Exception:
            pass

        try:
            self.canvas.figure = self.fig
            self.canvas.draw()
        except Exception:
            pass
        
    # ---------------- Save / Load configuration ----------------
    def save_configuration(self):
        config = {}
        config['data_mode'] = self.data_mode.get()
        config['tdms_time_offset_hours'] = self.tdms_time_offset_hours
        config['title'] = self.title_entry.get()
        config['xlabel'] = self.xlabel_entry.get()
        config['ylabel'] = self.ylabel_entry.get()
        config['right_ylabel'] = self.right_ylabel_entry.get()
        config['right2_ylabel'] = self.right2_ylabel_entry.get()
        config['legend_loc_left'] = self.legend_loc_left.get()
        config['legend_loc_right'] = self.legend_loc_right.get()
        config['legend_loc_right2'] = self.legend_loc_right2.get()
        config['legend_cols_left'] = int(self.legend_cols_left.get())
        config['legend_cols_right'] = int(self.legend_cols_right.get())
        config['legend_cols_right2'] = int(self.legend_cols_right2.get())
        config['legend_x_left'] = self.legend_x_left.get()
        config['legend_y_left'] = self.legend_y_left.get()
        config['legend_x_right'] = self.legend_x_right.get()
        config['legend_y_right'] = self.legend_y_right.get()
        config['legend_x_right2'] = self.legend_x_right2.get()
        config['legend_y_right2'] = self.legend_y_right2.get()
        config['font_size'] = int(self.font_size.get())
        config['show_grid'] = bool(self.grid_var.get())
        config['live_preview'] = bool(self.live_preview_var.get())
        config['downsample_live'] = bool(self.downsample_live_var.get())
        config['max_preview_points'] = int(self.max_preview_points.get())
        config['marker_size'] = float(self.marker_size.get())
        config['plot_width'] = float(self.plot_width.get())
        config['plot_height'] = float(self.plot_height.get())
        config['global_line_width'] = float(self.global_line_width.get())
        config['right2_pos'] = float(self.right2_pos.get())
        config['axis_colors'] = {
            'left': self.left_axis_color.get(),
            'right': self.right_axis_color.get(),
            'right2': self.right2_axis_color.get()
        }

        config['left_channel_props'] = []
        for row in self.line_properties_frames:
            try:
                config['left_channel_props'].append({
                    'col': getattr(row, "_col_name", ""),
                    'label': row._linked_label_var.get() if hasattr(row, "_linked_label_var") else "",
                    'line_color': getattr(row, 'line_color', ""),
                    'marker_color': getattr(row, 'marker_color', "") or "",
                    'style': row.line_style.get() if hasattr(row, 'line_style') else "",
                    'marker': row.marker.get() if hasattr(row, 'marker') else "",
                    'scatter': row.scatter_mode.get() if hasattr(row, 'scatter_mode') else False
                })
            except Exception:
                continue

        config['right_channel_props'] = []
        for row in self.right_line_properties_frames:
            try:
                config['right_channel_props'].append({
                    'col': getattr(row, "_col_name", ""),
                    'label': row._linked_label_var.get() if hasattr(row, "_linked_label_var") else "",
                    'line_color': getattr(row, 'line_color', ""),
                    'marker_color': getattr(row, 'marker_color', "") or "",
                    'style': row.line_style.get() if hasattr(row, 'line_style') else "",
                    'marker': row.marker.get() if hasattr(row, 'marker') else "",
                    'scatter': row.scatter_mode.get() if hasattr(row, 'scatter_mode') else False
                })
            except Exception:
                continue

        config['right2_channel_props'] = []
        for row in self.right2_line_properties_frames:
            try:
                config['right2_channel_props'].append({
                    'col': getattr(row, "_col_name", ""),
                    'label': row._linked_label_var.get() if hasattr(row, "_linked_label_var") else "",
                    'line_color': getattr(row, 'line_color', ""),
                    'marker_color': getattr(row, 'marker_color', "") or "",
                    'style': row.line_style.get() if hasattr(row, 'line_style') else "",
                    'marker': row.marker.get() if hasattr(row, 'marker') else "",
                    'scatter': row.scatter_mode.get() if hasattr(row, 'scatter_mode') else False
                })
            except Exception:
                continue

        config['axis_limits'] = {
            'xmin': self.xmin_entry.get(),
            'xmax': self.xmax_entry.get(),
            'ymin': self.ymin_entry.get(),
            'ymax': self.ymax_entry.get(),
            'right_ymin': self.right_ymin_entry.get(),
            'right_ymax': self.right_ymax_entry.get(),
            'right2_ymin': self.right2_ymin_entry.get(),
            'right2_ymax': self.right2_ymax_entry.get(),
            'xinterval': self.xinterval_entry.get(),
            'yinterval': self.yinterval_entry.get(),
            'right_yinterval': self.right_yinterval_entry.get(),
            'right2_yinterval': self.right2_yinterval_entry.get()
        }

        # save detected header line index
        config['detected_header_line_index'] = self.detected_header_line_index

        default_name = f"csv_tdms_plotter_config_{datetime.utcnow().strftime('%Y%m%dT%H%M%SZ')}.json"
        filepath = filedialog.asksaveasfilename(
            defaultextension=".json",
            initialfile=default_name,
            filetypes=[("JSON files", "*.json")]
        )
        if not filepath:
            return
        try:
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(config, f, indent=2)
            messagebox.showinfo("Saved", f"Configuration saved to {filepath}")
        except Exception as e:
            messagebox.showerror("Save Error", str(e))

    def load_configuration(self):
        path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        except Exception as e:
            messagebox.showerror("Load Error", f"Could not read configuration: {e}")
            return

        try:
            # Load data mode if present
            if 'data_mode' in cfg:
                self.data_mode.set(cfg['data_mode'])
                self.on_mode_change()
            
            # Load TDMS time offset (informational only, hardcoded to 0)
            if 'tdms_time_offset_hours' in cfg:
                try:
                    self.tdms_time_offset_hours = float(cfg.get('tdms_time_offset_hours', 0.0))
                except Exception:
                    self.tdms_time_offset_hours = 0.0
            
            self.title_entry.delete(0, "end")
            self.title_entry.insert(0, cfg.get('title', ''))
            self.xlabel_entry.delete(0, "end")
            self.xlabel_entry.insert(0, cfg.get('xlabel', ''))
            self.ylabel_entry.delete(0, "end")
            self.ylabel_entry.insert(0, cfg.get('ylabel', ''))
            self.right_ylabel_entry.delete(0, "end")
            self.right_ylabel_entry.insert(0, cfg.get('right_ylabel', ''))
            self.right2_ylabel_entry.delete(0, "end")
            self.right2_ylabel_entry.insert(0, cfg.get('right2_ylabel', ''))

            self.legend_loc_left.set(cfg.get('legend_loc_left', self.legend_loc_left.get()))
            self.legend_loc_right.set(cfg.get('legend_loc_right', self.legend_loc_right.get()))
            self.legend_loc_right2.set(cfg.get('legend_loc_right2', self.legend_loc_right2.get()))
            
            try:
                self.legend_cols_left.set(cfg.get('legend_cols_left', self.legend_cols_left.get()))
                self.legend_cols_right.set(cfg.get('legend_cols_right', self.legend_cols_right.get()))
                self.legend_cols_right2.set(cfg.get('legend_cols_right2', self.legend_cols_right2.get()))
            except Exception:
                pass
            
            self.legend_x_left.set(cfg.get('legend_x_left', self.legend_x_left.get()))
            self.legend_y_left.set(cfg.get('legend_y_left', self.legend_y_left.get()))
            self.legend_x_right.set(cfg.get('legend_x_right', self.legend_x_right.get()))
            self.legend_y_right.set(cfg.get('legend_y_right', self.legend_y_right.get()))
            self.legend_x_right2.set(cfg.get('legend_x_right2', self.legend_x_right2.get()))
            self.legend_y_right2.set(cfg.get('legend_y_right2', self.legend_y_right2.get()))

            try:
                self.font_size.set(cfg.get('font_size', self.font_size.get()))
            except Exception:
                pass
            
            self.grid_var.set(cfg.get('show_grid', bool(self.grid_var.get())))
            self.live_preview_var.set(cfg.get('live_preview', bool(self.live_preview_var.get())))
            self.downsample_live_var.set(cfg.get('downsample_live', bool(self.downsample_live_var.get())))
            
            try:
                self.max_preview_points.set(int(cfg.get('max_preview_points', self.max_preview_points.get())))
            except Exception:
                pass
            
            try:
                self.marker_size.set(float(cfg.get('marker_size', self.marker_size.get())))
            except Exception:
                pass
            
            try:
                self.plot_width.set(float(cfg.get('plot_width', self.plot_width.get())))
                self.plot_height.set(float(cfg.get('plot_height', self.plot_height.get())))
            except Exception:
                pass
            
            try:
                self.global_line_width.set(float(cfg.get('global_line_width', self.global_line_width.get())))
            except Exception:
                pass
            
            try:
                self.right2_pos.set(float(cfg.get('right2_pos', self.right2_pos.get())))
            except Exception:
                pass

            axis_colors = cfg.get('axis_colors', {})
            if 'left' in axis_colors:
                self.left_axis_color.set(axis_colors['left'])
                try:
                    self.left_axis_swatch.configure(background=axis_colors['left'])
                except Exception:
                    pass
            if 'right' in axis_colors:
                self.right_axis_color.set(axis_colors['right'])
                try:
                    self.right_axis_swatch.configure(background=axis_colors['right'])
                except Exception:
                    pass
            if 'right2' in axis_colors:
                self.right2_axis_color.set(axis_colors['right2'])
                try:
                    self.right2_axis_swatch.configure(background=axis_colors['right2'])
                except Exception:
                    pass

            axis_limits = cfg.get('axis_limits', {})
            limit_entries = {
                'xmin': self.xmin_entry,
                'xmax': self.xmax_entry,
                'ymin': self.ymin_entry,
                'ymax': self.ymax_entry,
                'right_ymin': self.right_ymin_entry,
                'right_ymax': self.right_ymax_entry,
                'right2_ymin': self.right2_ymin_entry,
                'right2_ymax': self.right2_ymax_entry,
                'xinterval': self.xinterval_entry,
                'yinterval': self.yinterval_entry,
                'right_yinterval': self.right_yinterval_entry,
                'right2_yinterval': self.right2_yinterval_entry
            }
            
            for key, entry_widget in limit_entries.items():
                try:
                    val = axis_limits.get(key, "")
                    entry_widget.delete(0, "end")
                    if val is not None:
                        entry_widget.insert(0, str(val))
                except Exception:
                    pass

            try:
                self.detected_header_line_index = cfg.get('detected_header_line_index', None)
            except Exception:
                self.detected_header_line_index = None

            # Load channel properties
            try:
                self.left_channel_props.clear()
                for p in cfg.get('left_channel_props', []):
                    col = p.get('col') or p.get('label') or ""
                    if not col:
                        continue
                    self.left_channel_props[col] = {
                        "label": p.get('label', col),
                        "line_color": p.get('line_color', ""),
                        "marker_color": p.get('marker_color', "") or "",
                        "style": p.get('style', "-"),
                        "marker": p.get('marker', "None"),
                        "scatter": p.get('scatter', False)
                    }
            except Exception:
                pass
            
            try:
                self.right_channel_props.clear()
                for p in cfg.get('right_channel_props', []):
                    col = p.get('col') or p.get('label') or ""
                    if not col:
                        continue
                    self.right_channel_props[col] = {
                        "label": p.get('label', col),
                        "line_color": p.get('line_color', ""),
                        "marker_color": p.get('marker_color', "") or "",
                        "style": p.get('style', "-"),
                        "marker": p.get('marker', "None"),
                        "scatter": p.get('scatter', False)
                    }
            except Exception:
                pass
            
            try:
                self.right2_channel_props.clear()
                for p in cfg.get('right2_channel_props', []):
                    col = p.get('col') or p.get('label') or ""
                    if not col:
                        continue
                    self.right2_channel_props[col] = {
                        "label": p.get('label', col),
                        "line_color": p.get('line_color', ""),
                        "marker_color": p.get('marker_color', "") or "",
                        "style": p.get('style', "-"),
                        "marker": p.get('marker', "None"),
                        "scatter": p.get('scatter', False)
                    }
            except Exception:
                pass

            self.schedule_preview()
            messagebox.showinfo("Loaded", f"Configuration loaded from {path}")
        except Exception as e:
            messagebox.showerror("Load Error", f"Could not apply configuration: {e}")
            
    # ---------------- Save PNG ----------------
    def save_png(self):
        if self.df is None:
            messagebox.showwarning("No plot", "Preview a plot first!")
            return
        filename = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG files", "*.png")]
        )
        if not filename:
            return
        
        try:
            pw = float(self.plot_width.get())
            ph = float(self.plot_height.get())
        except Exception:
            pw = self.default_plot_width
            ph = self.default_plot_height

        fig_save, ax_save = plt.subplots(figsize=(pw, ph))
        x_col = self.x_combo.get()
        selected_indices = self.y_listbox.curselection()
        right_indices = self.right_y_listbox.curselection()
        right2_indices = self.right2_y_listbox.curselection()
        y_cols = [self.y_listbox.get(i) for i in selected_indices]
        right_cols = [self.right_y_listbox.get(i) for i in right_indices]
        right2_cols = [self.right2_y_listbox.get(i) for i in right2_indices]

        try:
            global_lw = float(self.global_line_width.get())
        except Exception:
            global_lw = 1.0

        # Plot left axis
        # Prepare X-axis data once
        x_data = self.df[x_col].values
        if pd.api.types.is_datetime64_any_dtype(x_data):
            x_data = mdates.date2num(pd.to_datetime(x_data))
        
        for y_col, row in zip(y_cols, self.line_properties_frames):
            try:
                # Check if scatter mode is enabled
                is_scatter = hasattr(row, 'scatter_mode') and row.scatter_mode.get()
                
                marker = None if (row.marker.get() == "None") else row.marker.get()
                
                # For scatter mode, force marker if none selected
                if is_scatter and (marker is None or marker == "None"):
                    marker = "o"
                
                ydata = self.df[y_col].values.astype(float)
                mcolor = getattr(row, "marker_color", "") or row.line_color
                
                # Scatter mode: no line, only markers
                if is_scatter:
                    ax_save.scatter(x_data, ydata, s=self.marker_size.get()**2, 
                                   c=mcolor, marker=marker, label=row._linked_label_var.get(), 
                                   edgecolors=mcolor)
                else:
                    ax_save.plot(
                        x_data, ydata,
                        linestyle=row.line_style.get(),
                        linewidth=global_lw,
                        color=row.line_color,
                        marker=marker,
                        markersize=self.marker_size.get(),
                        markerfacecolor=mcolor,
                        markeredgecolor=mcolor,
                        label=row._linked_label_var.get()
                    )
            except Exception:
                continue

        try:
            left_col = self.left_axis_color.get()
            ax_save.spines["left"].set_color(left_col)
            ax_save.yaxis.label.set_color(left_col)
            ax_save.tick_params(axis="y", colors=left_col)
        except Exception:
            pass

        ax2_save = None
        if right_cols:
            ax2_save = ax_save.twinx()
            for y_col, row in zip(right_cols, self.right_line_properties_frames):
                try:
                    # Check if scatter mode is enabled
                    is_scatter = hasattr(row, 'scatter_mode') and row.scatter_mode.get()
                    
                    marker = None if (row.marker.get() == "None") else row.marker.get()
                    
                    # For scatter mode, force marker if none selected
                    if is_scatter and (marker is None or marker == "None"):
                        marker = "o"
                    
                    ydata = self.df[y_col].values.astype(float)
                    mcolor = getattr(row, "marker_color", "") or row.line_color
                    
                    # Scatter mode: no line, only markers
                    if is_scatter:
                        ax2_save.scatter(x_data, ydata, s=self.marker_size.get()**2, 
                                       c=mcolor, marker=marker, label=row._linked_label_var.get(), 
                                       edgecolors=mcolor)
                    else:
                        ax2_save.plot(
                            x_data, ydata,
                            linestyle=row.line_style.get(),
                            linewidth=global_lw,
                            color=row.line_color,
                            marker=marker,
                            markersize=self.marker_size.get(),
                            markerfacecolor=mcolor,
                            markeredgecolor=mcolor,
                            label=row._linked_label_var.get()
                        )
                except Exception:
                    continue
            try:
                right_col = self.right_axis_color.get()
                ax2_save.spines["right"].set_color(right_col)
                ax2_save.yaxis.label.set_color(right_col)
                ax2_save.tick_params(axis="y", colors=right_col)
            except Exception:
                pass
            try:
                ax2_save.set_ylabel(
                    self.right_ylabel_entry.get() or ", ".join([r._linked_label_var.get() for r in self.right_line_properties_frames]),
                    fontsize=self.font_size.get()
                )
            except Exception:
                pass

        ax3_save = None
        if right2_cols:
            ax3_save = ax_save.twinx()
            try:
                pos = float(self.right2_pos.get())
                ax3_save.spines["right"].set_position(("axes", pos))
            except Exception:
                pass
            for y_col, row in zip(right2_cols, self.right2_line_properties_frames):
                try:
                    # Check if scatter mode is enabled
                    is_scatter = hasattr(row, 'scatter_mode') and row.scatter_mode.get()
                    
                    marker = None if (row.marker.get() == "None") else row.marker.get()
                    
                    # For scatter mode, force marker if none selected
                    if is_scatter and (marker is None or marker == "None"):
                        marker = "o"
                    
                    ydata = self.df[y_col].values.astype(float)
                    mcolor = getattr(row, "marker_color", "") or row.line_color
                    
                    # Scatter mode: no line, only markers
                    if is_scatter:
                        ax3_save.scatter(x_data, ydata, s=self.marker_size.get()**2, 
                                       c=mcolor, marker=marker, label=row._linked_label_var.get(), 
                                       edgecolors=mcolor)
                    else:
                        ax3_save.plot(
                            x_data, ydata,
                            linestyle=row.line_style.get(),
                            linewidth=global_lw,
                            color=row.line_color,
                            marker=marker,
                            markersize=self.marker_size.get(),
                            markerfacecolor=mcolor,
                            markeredgecolor=mcolor,
                            label=row._linked_label_var.get()
                        )
                except Exception:
                    continue
            try:
                right2_col = self.right2_axis_color.get()
                ax3_save.spines["right"].set_color(right2_col)
                ax3_save.yaxis.label.set_color(right2_col)
                ax3_save.tick_params(axis="y", colors=right2_col)
            except Exception:
                pass
            try:
                ax3_save.set_ylabel(
                    self.right2_ylabel_entry.get() or ", ".join([r._linked_label_var.get() for r in self.right2_line_properties_frames]),
                    fontsize=self.font_size.get()
                )
            except Exception:
                pass

        # Apply axis limits
        try:
            xmin = float(self.xmin_entry.get()) if self.xmin_entry.get() else None
            xmax = float(self.xmax_entry.get()) if self.xmax_entry.get() else None
            ymin = float(self.ymin_entry.get()) if self.ymin_entry.get() else None
            ymax = float(self.ymax_entry.get()) if self.ymax_entry.get() else None
            ax_save.set_xlim(xmin, xmax)
            ax_save.set_ylim(ymin, ymax)
            if ax2_save:
                rymin = float(self.right_ymin_entry.get()) if self.right_ymin_entry.get() else None
                rymax = float(self.right_ymax_entry.get()) if self.right_ymax_entry.get() else None
                ax2_save.set_ylim(rymin, rymax)
            if ax3_save:
                ry2min = float(self.right2_ymin_entry.get()) if self.right2_ymin_entry.get() else None
                ry2max = float(self.right2_ymax_entry.get()) if self.right2_ymax_entry.get() else None
                ax3_save.set_ylim(ry2min, ry2max)
        except Exception:
            pass

        # Apply tick intervals
        try:
            if self.xinterval_entry.get().strip():
                xint = float(self.xinterval_entry.get())
                if xint > 0:
                    ax_save.xaxis.set_major_locator(MultipleLocator(xint))
            if self.yinterval_entry.get().strip():
                yint = float(self.yinterval_entry.get())
                if yint > 0:
                    ax_save.yaxis.set_major_locator(MultipleLocator(yint))
            if ax2_save and self.right_yinterval_entry.get().strip():
                ryint = float(self.right_yinterval_entry.get())
                if ryint > 0:
                    ax2_save.yaxis.set_major_locator(MultipleLocator(ryint))
            if ax3_save and self.right2_yinterval_entry.get().strip():
                ry2int = float(self.right2_yinterval_entry.get())
                if ry2int > 0:
                    ax3_save.yaxis.set_major_locator(MultipleLocator(ry2int))
        except Exception:
            pass

        # Adjust margins for multiple axes
        try:
            if ax3_save is not None:
                right_margin = 0.70
            elif ax2_save is not None:
                right_margin = 0.82
            else:
                right_margin = 0.92
            right_margin = max(0.5, min(0.95, right_margin))
            fig_save.subplots_adjust(right=right_margin)
        except Exception:
            pass

        fs = self.font_size.get()
        try:
            ax_save.set_title(self.title_entry.get(), fontsize=fs)
            ax_save.set_xlabel(self.xlabel_entry.get() or x_col, fontsize=fs)
            
            # Format x-axis if using converted datetime format
            if x_col.endswith('_Converted'):
                # Use matplotlib date formatter for better performance
                ax_save.xaxis.set_major_formatter(mdates.DateFormatter('%Y/%m/%d\n%H:%M:%S'))
                ax_save.xaxis.set_major_locator(mdates.AutoDateLocator())
                fig_save.autofmt_xdate(rotation=45)
                fig_save.tight_layout()
            ax_save.set_ylabel(
                self.ylabel_entry.get() or ", ".join([r._linked_label_var.get() for r in self.line_properties_frames]),
                fontsize=fs
            )
            ax_save.tick_params(axis="both", labelsize=fs)
            if ax2_save:
                ax2_save.tick_params(axis="both", labelsize=fs)
            if ax3_save:
                ax3_save.tick_params(axis="both", labelsize=fs)
            if self.grid_var.get():
                ax_save.grid(True)
        except Exception:
            pass

        def place_legend_save(ax, legend_pos, cols, x_override=None, y_override=None):
            outside = False
            loc_map = {
                "outside top": "upper center",
                "outside bottom": "lower center",
                "outside left": "center left",
                "outside right": "center right"
            }
            if legend_pos and legend_pos.startswith("outside"):
                outside = True
                loc = loc_map.get(legend_pos, "best")
            else:
                loc = legend_pos or "best"
            try:
                xo = float(x_override) if (x_override is not None and x_override != "") else None
                yo = float(y_override) if (y_override is not None and y_override != "") else None
            except Exception:
                xo = yo = None
            if xo is not None and yo is not None:
                ax.legend(loc=loc, bbox_to_anchor=(xo, yo), fontsize=fs, ncol=cols)
            elif outside:
                if "top" in legend_pos:
                    ax.legend(loc=loc, bbox_to_anchor=(0.5, 1.15), fontsize=fs, ncol=cols)
                elif "bottom" in legend_pos:
                    ax.legend(loc=loc, bbox_to_anchor=(0.5, -0.3), fontsize=fs, ncol=cols)
                elif "left" in legend_pos:
                    ax.legend(loc=loc, bbox_to_anchor=(-0.3, 0.5), fontsize=fs, ncol=cols)
                elif "right" in legend_pos:
                    ax.legend(loc=loc, bbox_to_anchor=(1.2, 0.5), fontsize=fs, ncol=cols)
                else:
                    ax.legend(loc=loc, fontsize=fs, ncol=cols)
            else:
                ax.legend(loc=loc, fontsize=fs, ncol=cols)

        try:
            place_legend_save(
                ax_save,
                self.legend_loc_left.get(),
                self.legend_cols_left.get(),
                self.legend_x_left.get(),
                self.legend_y_left.get()
            )
            if ax2_save:
                place_legend_save(
                    ax2_save,
                    self.legend_loc_right.get(),
                    self.legend_cols_right.get(),
                    self.legend_x_right.get(),
                    self.legend_y_right.get()
                )
            if ax3_save:
                place_legend_save(
                    ax3_save,
                    self.legend_loc_right2.get(),
                    self.legend_cols_right2.get(),
                    self.legend_x_right2.get(),
                    self.legend_y_right2.get()
                )
        except Exception:
            pass

        try:
            dpi = int(self.dpi_option.get())
        except Exception:
            dpi = 200
        
        try:
            fig_save.savefig(filename, dpi=dpi, bbox_inches="tight")
            plt.close(fig_save)
            messagebox.showinfo("Saved", f"Plot saved as {filename} at {dpi} DPI")
        except Exception as e:
            messagebox.showerror("Save Error", str(e))
            
    # ---------------- Resets ----------------
    def reset_all(self):
        try:
            self.df = None
            self.file_entry.delete(0, "end")
            self.tdms_folder_entry.delete(0, "end")
            self.x_combo.set("")
            self.x_combo["values"] = []
            self.y_listbox.delete(0, tk.END)
            self.right_y_listbox.delete(0, tk.END)
            self.right2_y_listbox.delete(0, tk.END)
            self.tdms_files_listbox.delete(0, tk.END)
            self.reset_line_properties()
            self.reset_right_line_properties()
            self.reset_right2_line_properties()
            self.left_channel_props.clear()
            self.right_channel_props.clear()
            self.right2_channel_props.clear()
            self.detected_header_line_index = None
            self.tdms_files = []
            self.tdms_folder = None
            self.tdms_time_offset_hours = 0.0
            self.reset_labels()
            self.reset_limits()
            self.reset_axis_colors()
            self.live_preview_var.set(True)
            self.downsample_live_var.set(True)
            self.max_preview_points.set(5000)
            self.global_line_width.set(1.0)
            self.marker_size.set(6.0)
            self.plot_width.set(self.default_plot_width)
            self.plot_height.set(self.default_plot_height)
            self.right2_pos.set(1.2)
            try:
                self.dpi_option.current(2)
            except Exception:
                pass
            self.schedule_preview()
        except Exception:
            pass

    def reset_csv(self):
        self.file_entry.delete(0, "end")
        self.tdms_folder_entry.delete(0, "end")
        self.df = None
        self.x_combo.set("")
        self.y_listbox.delete(0, "end")
        self.right_y_listbox.delete(0, "end")
        self.right2_y_listbox.delete(0, "end")
        self.tdms_files_listbox.delete(0, "end")
        self.tdms_files = []
        self.tdms_folder = None
        self.reset_line_properties()
        self.reset_right_line_properties()
        self.reset_right2_line_properties()
        self.schedule_preview()

    def reset_axis(self):
        self.x_combo.set("")
        self.y_listbox.selection_clear(0, "end")
        self.reset_line_properties()
        self.schedule_preview()

    def reset_right_axis(self):
        self.right_y_listbox.selection_clear(0, "end")
        self.reset_right_line_properties()
        self.schedule_preview()

    def reset_right2_axis(self):
        self.right2_y_listbox.selection_clear(0, "end")
        self.reset_right2_line_properties()
        self.schedule_preview()

    def reset_line_properties(self):
        for w in self.line_properties_frames:
            try:
                w.destroy()
            except Exception:
                pass
        self.line_properties_frames.clear()
        self.schedule_preview()

    def reset_right_line_properties(self):
        for w in self.right_line_properties_frames:
            try:
                w.destroy()
            except Exception:
                pass
        self.right_line_properties_frames.clear()
        self.schedule_preview()

    def reset_right2_line_properties(self):
        for w in self.right2_line_properties_frames:
            try:
                w.destroy()
            except Exception:
                pass
        self.right2_line_properties_frames.clear()
        self.schedule_preview()

    def reset_labels(self):
        try:
            self.title_entry.delete(0, "end")
            self.xlabel_entry.delete(0, "end")
            self.ylabel_entry.delete(0, "end")
            self.right_ylabel_entry.delete(0, "end")
            self.right2_ylabel_entry.delete(0, "end")
        except Exception:
            pass
        try:
            self.legend_loc_left.current(0)
            self.legend_loc_right.current(0)
            self.legend_loc_right2.current(0)
            self.legend_cols_left.set(1)
            self.legend_cols_right.set(1)
            self.legend_cols_right2.set(1)
            self.legend_x_left.set("")
            self.legend_y_left.set("")
            self.legend_x_right.set("")
            self.legend_y_right.set("")
            self.legend_x_right2.set("")
            self.legend_y_right2.set("")
        except Exception:
            pass
        try:
            self.font_size.set(12)
            self.grid_var.set(True)
            self.marker_size.set(6.0)
            self.left_axis_color.set("#000000")
            self.right_axis_color.set("#1f77b4")
            self.right2_axis_color.set("#d62728")
            self.right2_pos.set(1.2)
            self.left_axis_swatch.configure(background=self.left_axis_color.get())
            self.right_axis_swatch.configure(background=self.right_axis_color.get())
            self.right2_axis_swatch.configure(background=self.right2_axis_color.get())
        except Exception:
            pass
        self.schedule_preview()

    def reset_limits(self):
        for entry in [
            self.xmin_entry, self.xmax_entry,
            self.ymin_entry, self.ymax_entry,
            self.right_ymin_entry, self.right_ymax_entry,
            self.right2_ymin_entry, self.right2_ymax_entry,
            self.xinterval_entry, self.yinterval_entry,
            self.right_yinterval_entry, self.right2_yinterval_entry
        ]:
            try:
                entry.delete(0, "end")
            except Exception:
                pass
        self.schedule_preview()


# ========== MAIN ENTRY POINT ==========
if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("1600x1000")
    app = CSVPlotter(root)
    root.mainloop()
