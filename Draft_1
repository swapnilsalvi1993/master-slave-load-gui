from __future__ import annotations

import csv
import json
import os
import queue
import threading
import time
import traceback
from collections import deque
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import List, Optional, Tuple, Union
import smtplib
from email.mime.text import MIMEText

# timezone support (CST/CDT via America/Chicago)
try:
    from zoneinfo import ZoneInfo
    CHI_TZ = ZoneInfo("America/Chicago")
except Exception:
    CHI_TZ = None

# serial optional (pyserial)
try:
    import serial
    from serial.serialutil import SerialException
except Exception:
    serial = None
    SerialException = Exception

# GUI
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog, colorchooser
import tkinter.font as tkfont

# plotting
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends. backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.ticker import MaxNLocator

# --- Defaults and constants ---
DEFAULT_MASTER_PORT = "COM10"
DEFAULT_SLAVE_PORT = "COM8"
DEFAULT_BAUD_RATE = 9600
DEFAULT_LOG_INTERVAL = 60
DEFAULT_CUTOFF_VOLTAGE = 180
DEFAULT_RECORD_INTERVAL_S = 1
DEFAULT_ROTATE_HOURS = 6
DEFAULT_IMBALANCE_LIMIT = 0.10  # 10%
DEFAULT_IMBALANCE_TRANSITION = 0.25  # 25%
DEFAULT_TRANSITION_DURATION = 120  # seconds
DEFAULT_VOLTAGE_DIFF_LIMIT = 0.10  # 10%

# Default power profile:  (Total Power W, Duration hours)
DEFAULT_PROFILE = [
    (384.615, 24),
    (28.846, 6),
    (1923.077, 24),
    (384.615, 12),
    (2940.769, 2),
    (300.769, 5),
]

UNITS = [("days", 86400), ("hours", 3600), ("minutes", 60), ("seconds", 1)]
LINE_STYLES = {"solid": "-", "dashed": "--", "dotted": ":", "dashdot": "-. "}
COLOR_CHOICES = ["black", "red", "blue", "green", "orange", "purple", "brown", "grey"]

# Add this: 
MARKER_STYLES = {
    "circle": "o",
    "square": "s", 
    "triangle": "^",
    "diamond": "D",
    "plus": "+",
    "x": "x",
    "none": ""
}

# --- Data models ---
@dataclass
class ProfileRow:
    total_power: float = 0.0  # Total power (will be split 50/50)
    duration_value: Union[int, float] = 0
    duration_unit: str = "seconds"
    cutoff:  bool = False


@dataclass
class Settings:
    master_port: str = DEFAULT_MASTER_PORT
    slave_port: str = DEFAULT_SLAVE_PORT
    baud_rate: int = DEFAULT_BAUD_RATE
    log_interval:  int = DEFAULT_LOG_INTERVAL
    cutoff_voltage: float = DEFAULT_CUTOFF_VOLTAGE
    dry_run: bool = False
    timeout: float = 1.0
    profile: List[ProfileRow] = field(default_factory=list)
    cutoff_safety_enabled: bool = False
    imbalance_check_enabled: bool = True  # NEW
    imbalance_limit: float = DEFAULT_IMBALANCE_LIMIT
    imbalance_transition:   float = DEFAULT_IMBALANCE_TRANSITION
    transition_duration: int = DEFAULT_TRANSITION_DURATION
    voltage_diff_limit:   float = DEFAULT_VOLTAGE_DIFF_LIMIT
    email_enabled: bool = False
    email_sender: str = ""
    email_receiver: str = ""
    email_password: str = ""
    email_subject: str = "IT8514B+ Load Imbalance Alert"


# --- Instrument Controller (for master or slave) ---
class InstrumentController:
    def __init__(self, port: str, baud_rate: int, log_fn, name: str = "Load"):
        self.port = port
        self.baud_rate = baud_rate
        self.ser = None
        self.log = log_fn
        self.name = name

    def open(self, timeout: float = 1.0):
        if serial is None:
            self.log(f"[DRY-RUN] {self.name} serial not opened (pyserial missing).")
            self.ser = None
            return
        try:
            self.ser = serial.Serial(port=self.port, baudrate=self.baud_rate, timeout=timeout)
            time.sleep(0.2)
            try:
                self.ser.reset_input_buffer()
                self.ser.reset_output_buffer()
            except Exception: 
                pass
            self.log(f"{self.name}:  Opened serial {self.port} @ {self.baud_rate} baud.")
        except SerialException as e:
            self.ser = None
            self.log(f"[ERROR] {self.name}:  Unable to open serial port: {e}")
            raise

    def close(self):
        if self.ser:
            try:
                self.ser.close()
                self.log(f"{self. name}: Closed serial port.")
            except Exception as e:
                self.log(f"[WARN] {self.name}: Error closing serial:  {e}")
        self.ser = None

    def send_cmd(self, cmd: str, delay: float = 0.05):
        if self.ser is None:
            self.log(f"[DRY-RUN] {self.name} > {cmd}")
            time.sleep(delay)
            return
        try:
            self. ser.write((cmd + "\n").encode("utf-8"))
            time.sleep(delay)
            return
        except SerialException as e:
            self.log(f"[ERROR] {self.name}: Write failed '{cmd}': {e}")
            raise

    def read_response(self) -> Optional[str]:
        if self. ser is None:
            return None
        try:
            line = self.ser.readline()
            if not line:
                return None
            return line.decode("utf-8", errors="replace").strip()
        except SerialException as e:
            self. log(f"[ERROR] {self.name}: Read failed: {e}")
            return None

    def query_float(self, qcmd: str, attempts: int = 3, delay_between: float = 0.05) -> Optional[float]:
        for _ in range(attempts):
            try:
                self.send_cmd(qcmd)
            except Exception:
                time.sleep(delay_between)
                continue
            resp = self.read_response()
            if not resp:
                time.sleep(delay_between)
                continue
            token = resp.split()[0]. replace(",", "")
            try:
                return float(token)
            except ValueError:
                time.sleep(delay_between)
                continue
        return None

    # SCPI-like commands
    def set_remote(self): self.send_cmd("SYST:REM")
    def set_local(self): self.send_cmd("SYST:LOC")
    def set_func_cp(self): self.send_cmd("FUNC CP")
    def set_func_pow(self): self.send_cmd("FUNC POW")
    def set_power(self, p: float): self.send_cmd(f"POW {p}")
    def input_on(self): self.send_cmd("INPUT ON")
    def load_on(self): self.send_cmd("LOAD ON")
    def input_off(self): self.send_cmd("INPUT OFF")
    def load_off(self): self.send_cmd("LOAD OFF")
    def read_voltage(self) -> Optional[float]:  return self.query_float("MEAS:VOLT? ")
    def read_current(self) -> Optional[float]: return self.query_float("MEAS:CURR?")
    def read_power(self) -> Optional[float]: return self.query_float("MEAS:POW?")
    def set_parallel_master(self): self.send_cmd("CONF:PARA:MODE MASTER")
    def set_parallel_slave(self): self.send_cmd("CONF: PARA:MODE SLAVE")
    def parallel_on(self): self.send_cmd("CONF:PARA ON")
    def parallel_off(self): self.send_cmd("CONF:PARA OFF")
    def get_idn(self) -> Optional[str]: 
        self.send_cmd("*IDN?")
        return self.read_response()

    def set_mode_cw(self):
        """Set instrument to Constant Power (CW) mode"""
        try:
            # Try format 1: MODE: CW (no space)
            cmd = "MODE:CW\n"
            self. ser.write(cmd.encode('ascii'))
            time.sleep(0.2)
            self.log(f"{self.name}:  Sent MODE:CW command")
            
            # Verify by reading back mode
            self.ser.write(b"MODE?\n")
            time.sleep(0.2)
            if self.ser.in_waiting:
                response = self.ser.read(self.ser.in_waiting).decode('ascii').strip()
                self.log(f"{self.name}: Current mode: {response}")
                if "CW" in response. upper() or "POW" in response.upper():
                    self.log(f"{self.name}: ✓ Successfully set to CW mode")
                    return
            
            # If that didn't work, try format 2: MODE POW
            self.log(f"{self.name}:  Trying alternate command:  MODE POW")
            cmd = "MODE POW\n"
            self.ser.write(cmd.encode('ascii'))
            time.sleep(0.2)
            
            # Verify again
            self.ser.write(b"MODE?\n")
            time.sleep(0.2)
            if self.ser.in_waiting:
                response = self.ser.read(self.ser.in_waiting).decode('ascii').strip()
                self.log(f"{self.name}: Current mode: {response}")
                
        except Exception as e:
            self.log(f"{self.name}: [ERROR] Setting CW mode: {e}")

# --- Email Alert System ---
def send_email_alert(settings: Settings, subject: str, body: str, log_fn):
    """Send an alert email via SMTP."""
    if not settings.email_enabled or not settings.email_password:
        log_fn("[INFO] Email alerts not enabled or password not set.")
        return
    try:
        msg = MIMEText(body)
        msg['Subject'] = subject
        msg['From'] = settings.email_sender
        msg['To'] = settings.email_receiver

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server. login(settings.email_sender, settings.email_password)
            server.send_message(msg)

        log_fn("✓ Alert email sent successfully.")
    except Exception as e:
        log_fn(f"✗ Failed to send email alert: {e}")


# --- GUI row widget for profile editor ---
class ProfileRowWidget: 
    def __init__(self, parent, index: int, row: ProfileRow, on_delete, on_move_up, on_move_down):
        self.parent = parent
        self.index = index
        self.row = row
        self.on_delete = on_delete
        self.on_move_up = on_move_up
        self.on_move_down = on_move_down
        self.frame = ttk.Frame(parent)

        # Total Power entry
        self.power_var = tk.StringVar(value=str(self.row.total_power))
        self.power_entry = ttk. Entry(self.frame, width=12, textvariable=self. power_var)
        self.power_entry.grid(row=0, column=0, padx=4, pady=2)

        # Time entry
        self.time_var = tk.StringVar(value=str(self.row.duration_value))
        self.time_entry = ttk.Entry(self. frame, width=10, textvariable=self. time_var)
        self.time_entry.grid(row=0, column=1, padx=4, pady=2)

        # Unit menu
        self.unit_var = tk.StringVar(value=self.row.duration_unit)
        self.unit_menu = ttk.OptionMenu(self.frame, self.unit_var, self.row.duration_unit, *[u for u, _ in UNITS])
        self.unit_menu.grid(row=0, column=2, padx=4, pady=2)

        # Cutoff checkbox
        self.cutoff_var = tk.BooleanVar(value=self.row.cutoff)
        self.cutoff_cb = ttk.Checkbutton(self.frame, text="Cutoff", variable=self.cutoff_var, command=self._on_cutoff_toggle)
        self.cutoff_cb.grid(row=0, column=3, padx=6, pady=2)

        # Move up/down/delete
        self.up_btn = ttk.Button(self. frame, text="↑", width=3, command=lambda: on_move_up(self.index))
        self.up_btn.grid(row=0, column=4, padx=1)
        self.down_btn = ttk.Button(self. frame, text="↓", width=3, command=lambda: on_move_down(self.index))
        self.down_btn.grid(row=0, column=5, padx=1)
        self.del_btn = ttk.Button(self.frame, text="Delete", width=6, command=lambda: on_delete(self.index))
        self.del_btn.grid(row=0, column=6, padx=4)

        # Header labels above first widget
        if index == 0:
            header = ttk.Frame(parent)
            header.grid(row=0, column=0, sticky="w", pady=(4, 0))
            ttk.Label(header, text="Total Power (W)", width=12).grid(row=0, column=0, padx=4)
            ttk.Label(header, text="Time", width=10).grid(row=0, column=1, padx=4)
            ttk.Label(header, text="Unit", width=10).grid(row=0, column=2, padx=4)
            ttk.Label(header, text="", width=8).grid(row=0, column=3, padx=4)

        self._on_cutoff_toggle(initial=True)

    def _on_cutoff_toggle(self, initial=False):
        checked = self.cutoff_var. get()
        if checked:
            self.time_entry.grid_remove()
            self.unit_menu.grid_remove()
        else:
            self.time_entry.grid()
            self.unit_menu. grid()
        if not initial:
            self.row.cutoff = checked

    def grid(self, row_idx):
        self.frame.grid(row=row_idx, column=0, sticky="w", pady=2)

    def update_index(self, new_index):
        self.index = new_index

    def get_value(self) -> ProfileRow:
        try:
            power = float(self.power_var.get())
        except Exception:
            power = 0.0
        cutoff = bool(self.cutoff_var. get())
        if cutoff:
            return ProfileRow(total_power=power, duration_value=0, duration_unit="seconds", cutoff=True)
        try:
            dval = float(self.time_var.get())
        except Exception:
            dval = 0
        dunit = self.unit_var.get()
        return ProfileRow(total_power=power, duration_value=dval, duration_unit=dunit, cutoff=False)

    def destroy(self):
        self.frame.destroy()


# --- Worker thread for master-slave test ---
class MasterSlaveWorkerThread(threading.Thread):
    def __init__(self, settings:  Settings, csv_folder: Optional[str], rotate_hours: float,
                 record_interval_s: int, log_queue: queue.Queue, plot_queue: Optional[queue.Queue] = None):
        super().__init__(daemon=True)
        self.settings = settings
        self.csv_folder = csv_folder
        self.rotate_hours = rotate_hours
        self.record_interval_s = max(1, int(record_interval_s))
        self.log_queue = log_queue
        self.plot_queue = plot_queue
        self._stop_event = threading.Event()
        self.master = InstrumentController(settings. master_port, settings.baud_rate, self._log, "MASTER")
        self.slave = InstrumentController(settings.slave_port, settings.baud_rate, self._log, "SLAVE")
        self.csv_file = None
        self. csv_writer = None
        self.current_csv_start = None
        self.test_start_time = None
        self. csv_lock = threading.Lock()
        
        # NEW: Log file handling
        self.log_file = None
        self.log_writer = None
        self.log_lock = threading.Lock()
        
        # Energy/Capacity tracking
        self.cumulative_energy_wh = 0.0
        self.cumulative_capacity_ah = 0.0
        self.last_measurement_time = None
        self.last_total_power = 0.0
        self.last_total_current = 0.0
        
        # Differential capacity tracking (dQ/dV)
        self.last_avg_voltage = None
        self.last_capacity_ah = 0.0
        self. dq_dv = 0.0
        self.dq_dv_voltage = 0.0
        self.dq_dv_valid = False

    def _open_log_file(self):
        """Open a new log CSV file with timestamp-based naming"""
        if not self.csv_folder:
            return
        
        try:
            os.makedirs(self.csv_folder, exist_ok=True)
            start = self._now_central()
            fname = start.strftime("master_slave_load_%Y%m%d_%H%M%S_central_test_logs.csv")
            path = os.path.join(self.csv_folder, fname)
            
            f = open(path, "w", newline="", encoding="utf-8")
            writer = csv.writer(f)
            writer.writerow(["timestamp_central", "relative_seconds", "log_level", "message"])
            
            self.log_file = f
            self. log_writer = writer
            self._log(f"Opened log file: {path}")
        except Exception as e:
            print(f"[ERROR] Could not open log file: {e}")
    
    def _write_log_entry(self, message: str):
        """Write a log entry to the log CSV file"""
        if not self. log_writer:
            return
        
        try:
            ts = self._format_ts(self._now_central())
            
            # Calculate relative time if test has started
            if self.test_start_time: 
                rel = int((self._now_central() - self.test_start_time).total_seconds())
            else:
                rel = 0
            
            # Determine log level from message
            log_level = "INFO"
            if "[ERROR]" in message or "[EXCEPTION]" in message:
                log_level = "ERROR"
            elif "[ALERT]" in message or "[WARNING]" in message or "⚠" in message:
                log_level = "WARNING"
            elif "[CUTOFF]" in message: 
                log_level = "CUTOFF"
            elif "===" in message or "═══" in message:
                log_level = "MILESTONE"
            elif "[DEBUG]" in message:
                log_level = "DEBUG"
            
            with self.log_lock:
                self.log_writer.writerow([ts, rel, log_level, message])
                self. log_file.flush()
        except Exception as e:
            print(f"[ERROR] Writing log entry: {e}")
    
    def _close_log_file(self):
        """Close the log file"""
        if self.log_file:
            try:
                with self.log_lock:
                    self.log_file.close()
                    self. log_file = None
                    self.log_writer = None
                self._log("Log file closed")
            except Exception as e:
                print(f"[ERROR] Closing log file: {e}")

    def request_stop(self):
        """Request the worker thread to stop"""
        self._stop_event.set()  # CHANGED from _stop to _stop_event

    def _now_central(self) -> datetime: 
        if CHI_TZ is not None:
            return datetime.now(CHI_TZ)
        else:
            return datetime.now()

    def _format_ts(self, dt: datetime) -> str:
        try:
            return dt.isoformat()
        except Exception: 
            return dt.strftime("%Y-%m-%d %H:%M:%S")

    def _log(self, msg: str):
        """Log message to both queue and CSV file"""
        try:
            self.log_queue.put_nowait(msg)
        except Exception: 
            pass
        
        # Also write to log file
        if self.log_writer:
            self._write_log_entry(msg)

    def _open_new_csv(self):
        if not self.csv_folder:
            return
        os.makedirs(self.csv_folder, exist_ok=True)
        start = self._now_central()
        fname = start. strftime("master_slave_load_%Y%m%d_%H%M%S_central.csv")
        path = os.path.join(self.csv_folder, fname)
        f = open(path, "w", newline="", encoding="utf-8")
        writer = csv.writer(f)
        writer.writerow([
            "timestamp_central",
            "relative_seconds",
            "step_index",
            "total_set_power_w",
            "master_voltage_v",
            "master_current_a",
            "master_power_w",
            "slave_voltage_v",
            "slave_current_a",
            "slave_power_w",
            "total_power_w",
            "total_current_a",
            "avg_voltage_v",
            "power_imbalance_%",
            "current_imbalance_%",
            "voltage_imbalance_%",
            "cumulative_energy_wh",
            "cumulative_capacity_ah",
            "dq_dv_ah_per_v",  # NEW
            "dq_dv_at_voltage_v"  # NEW
        ])
        self.csv_file = f
        self.csv_writer = writer
        self.current_csv_start = start
        self._log(f"Opened CSV:  {path}")

    def _close_csv(self):
        if self.csv_file:
            try:
                self.csv_file.flush()
                self.csv_file.close()
                self._log("Closed CSV file.")
            except Exception as e:
                self._log(f"[WARN] Closing CSV failed: {e}")
        self.csv_file = None
        self.csv_writer = None
        self.current_csv_start = None

    def _maybe_rotate_csv(self):
        """Check if CSV needs rotation and open new file if needed"""
        if not self.current_csv_start or not self.csv_file:
            return
        
        elapsed = (self._now_central() - self.current_csv_start).total_seconds() / 3600.0
        if elapsed >= self.rotate_hours:
            self._log(f"Rotating CSV after {elapsed:.2f} hours")
            self._close_csv()
            self._open_new_csv()
            
            # NEW: Also rotate log file
            self._close_log_file()
            self._open_log_file()

    def _write_csv_row(self, step_idx:  int, total_set_power:  float, mv, mi, mp, sv, si, sp, 
                       power_imb, current_imb, voltage_imb, total_power, total_current, avg_voltage,
                       cumulative_energy_wh, cumulative_capacity_ah, dq_dv, dq_dv_voltage):
        if not self.csv_writer:
            return
        ts = self._format_ts(self._now_central())
        rel = int((self._now_central() - self.test_start_time).total_seconds())
        
        row = [
            ts, rel, step_idx, total_set_power,
            "" if mv is None else f"{mv:.6g}",
            "" if mi is None else f"{mi:.6g}",
            "" if mp is None else f"{mp:.6g}",
            "" if sv is None else f"{sv:.6g}",
            "" if si is None else f"{si:.6g}",
            "" if sp is None else f"{sp:.6g}",
            f"{total_power:.6g}",
            f"{total_current:.6g}",
            f"{avg_voltage:.6g}",
            f"{power_imb:.3f}",
            f"{current_imb:.3f}",
            f"{voltage_imb:.3f}",
            f"{cumulative_energy_wh:.6f}",
            f"{cumulative_capacity_ah:.6f}",
            f"{dq_dv:.6f}",  # NEW
            f"{dq_dv_voltage:.6f}"  # NEW
        ]
        try:
            with self.csv_lock:
                self.csv_writer.writerow(row)
                self.csv_file.flush()
        except Exception as e:
            self._log(f"[ERROR] Writing CSV row: {e}")

    def _push_plot_point(self, step_idx: int, mv, mi, mp, sv, si, sp, 
                        cumulative_energy_wh, cumulative_capacity_ah,
                        dq_dv, dq_dv_voltage, dq_dv_valid):
        """Push plot point to queue"""
        if self.plot_queue is None:
            return
        try:
            rel = int((self._now_central() - self.test_start_time).total_seconds())
            # Send exactly 13 values
            self.plot_queue.put((
                rel,                      # 1
                step_idx,                 # 2
                mv,                       # 3
                mi,                       # 4
                mp,                       # 5
                sv,                       # 6
                si,                       # 7
                sp,                       # 8
                cumulative_energy_wh,     # 9
                cumulative_capacity_ah,   # 10
                dq_dv,                    # 11
                dq_dv_voltage,            # 12
                dq_dv_valid               # 13
            ))
        except Exception as e:
            self._log(f"[ERROR] Pushing plot point: {e}")

    def _calculate_energy_increment(self, current_power, current_current, delta_time_hours):
        # Trapezoidal integration:  (P1 + P2)/2 * dt
        avg_power = (self.last_total_power + current_power) / 2.0
        avg_current = (self.last_total_current + current_current) / 2.0
        
        energy_wh = avg_power * delta_time_hours
        capacity_ah = avg_current * delta_time_hours
        
        return energy_wh, capacity_ah
    
    def _calculate_dq_dv(self, current_voltage, current_capacity):
        # Need at least one previous measurement
        if self.last_avg_voltage is None:
            self.last_avg_voltage = current_voltage
            self.last_capacity_ah = current_capacity
            return 0.0, current_voltage, False
        
        # Calculate differences
        dV = current_voltage - self.last_avg_voltage
        dQ = current_capacity - self.last_capacity_ah
        
        # Avoid division by very small voltage changes (noise threshold)
        voltage_threshold = 0.01  # 10 mV threshold
        
        if abs(dV) < voltage_threshold:
            # Voltage hasn't changed enough - keep last valid value
            return self. dq_dv, self. dq_dv_voltage, self.dq_dv_valid
        
        # Calculate dQ/dV
        dq_dv = dQ / dV
        
        # Calculate voltage at midpoint for better representation
        dq_dv_voltage = (current_voltage + self.last_avg_voltage) / 2.0
        
        # Update last values
        self.last_avg_voltage = current_voltage
        self.last_capacity_ah = current_capacity
        
        # Store for next calculation
        self. dq_dv = dq_dv
        self.dq_dv_voltage = dq_dv_voltage
        self.dq_dv_valid = True
        
        return dq_dv, dq_dv_voltage, True

    def _safe_shutdown(self, reason: Optional[str] = None):
        """Safe shutdown of loads - ramp down power, turn off, exit remote"""
        self._log("⚠ Initiating safe shutdown sequence...")
        
        if reason:
            self._log(f"  Reason: {reason}")
        
        # Step 1: Ramp power down to zero
        self._log("  Step 1: Ramping power to zero...")
        for name, controller in [("MASTER", self.master), ("SLAVE", self.slave)]:
            try:
                controller.set_power(0.0)
                self._log(f"    ✓ {name} power set to 0W")
                time.sleep(0.5)  # Brief delay for command to process
            except Exception as e: 
                self._log(f"    ✗ {name} power ramp failed: {e}")
        
        # Step 2: Turn off loads
        self._log("  Step 2: Turning loads OFF...")
        for name, controller in [("MASTER", self.master), ("SLAVE", self.slave)]:
            try:
                controller.load_off()
                self._log(f"    ✓ {name} LOAD OFF")
                time.sleep(0.2)
            except Exception as e:
                self._log(f"    ✗ {name} load off failed: {e}")
            
            try:
                controller. input_off()
                self._log(f"    ✓ {name} INPUT OFF")
                time.sleep(0.2)
            except Exception as e: 
                self._log(f"    ✗ {name} input off failed: {e}")
        
        # Step 3: Disable parallel mode
        self._log("  Step 3: Disabling parallel mode...")
        for name, controller in [("MASTER", self.master), ("SLAVE", self.slave)]:
            try:
                controller.parallel_off()
                self._log(f"    ✓ {name} PARALLEL OFF")
                time.sleep(0.2)
            except Exception as e:
                self._log(f"    ✗ {name} parallel off failed: {e}")
        
        # Step 4: Return to LOCAL mode
        self._log("  Step 4: Returning to LOCAL control...")
        for name, controller in [("MASTER", self.master), ("SLAVE", self.slave)]:
            try:
                controller.set_local()
                self._log(f"    ✓ {name} set to LOCAL mode")
                time.sleep(0.2)
            except Exception as e:
                self._log(f"    ✗ {name} set local failed: {e}")
        
        self._log("✓ Safe shutdown sequence complete.  Equipment is safe.")
        
        # Send email alert if configured and there was a reason
        if reason and self.settings.email_enabled: 
            body = (f"Equipment Safety Shutdown\n\n"
                   f"Reason: {reason}\n\n"
                   f"Shutdown sequence completed:\n"
                   f"1. Power ramped to 0W\n"
                   f"2. Loads turned OFF\n"
                   f"3. Parallel mode disabled\n"
                   f"4. Equipment returned to LOCAL mode\n\n"
                   f"Equipment is now in safe state.")
            send_email_alert(self.settings, self.settings.email_subject, body, self._log)

    def run(self):
        try:
            self._log("=== Master-Slave Worker Thread Starting ===")
            
            # Open log file
            self._open_log_file()
            
            # Open both connections
            try:
                self.master.open(self.settings.timeout)
                self.slave.open(self.settings.timeout)
            except Exception as e:
                self._log(f"[ERROR] Opening instruments: {e}")
                return

            # Get IDs
            try:
                master_id = self.master.get_idn()
                slave_id = self.slave.get_idn()
                self._log(f"Master ID: {master_id}")
                self._log(f"Slave  ID: {slave_id}")
            except Exception: 
                pass

            # Configure master-slave parallel mode
            try:
                self. master.set_remote()
                self.slave.set_remote()
                self.master.set_parallel_master()
                self.slave. set_parallel_slave()
                self. master.parallel_on()
                self.slave.parallel_on()
                self. master.set_func_pow()
                self.slave.set_func_pow()
                self._log("✓ Master-Slave parallel mode configured")
            except Exception as e:
                self._log(f"[WARN] Could not configure parallel mode: {e}")

            self.test_start_time = self._now_central()
            if self.csv_folder:
                try:
                    self._open_new_csv()
                except Exception as e:
                    self._log(f"[ERROR] Opening CSV:  {e}")

            # Execute profile
            for idx, row in enumerate(self.settings.profile, start=1):
                # Check if stop requested before starting next step
                if self._stop_event.is_set():
                    self._log("Stop requested before starting next step.")
                    self._safe_shutdown("Stop requested by user")
                    return  # Exit - safe shutdown already done
                    break

                total_power_setting = float(row. total_power)
                master_power = total_power_setting / 2.0  # 50/50 split
                slave_power = total_power_setting / 2.0
                total_set_power = total_power_setting

                if row.cutoff:
                    dur_desc = f"until voltage < {self.settings.cutoff_voltage} V"
                else:
                    unit_seconds = dict(UNITS)[row.duration_unit]
                    dur_seconds = float(row.duration_value) * unit_seconds
                    dur_desc = f"for {dur_seconds} seconds"

                self._log(f"═══ Step {idx}:  {total_set_power} W total ({master_power}W per load) {dur_desc} ═══")

                # Set power and enable
                try:
                    self.master.set_power(master_power)
                    self.slave.set_power(slave_power)
                    self.master.load_on()
                    self.slave.load_on()
                except Exception as e:
                    self._log(f"[ERROR] Failed to start step {idx}: {e}")
                    try:
                        self.master. load_off()
                        self.slave.load_off()
                    except Exception: 
                        pass
                    continue

                # Transition phase tracking
                transition_start = time.time()
                current_imbalance_limit = self.settings.imbalance_transition
                self._log(f"  Transition phase:  {current_imbalance_limit*100}% imbalance allowed for {self.settings.transition_duration}s")

                step_start = time.time()

                if row.cutoff:
                    # Run until cutoff
                    while not self._stop_event.is_set():
                        # Initialize all variables at the start
                        total_power = 0.0
                        total_current = 0.0
                        avg_voltage = 0.0
                        power_imb = 0.0
                        current_imb = 0.0
                        voltage_imb = 0.0
                        dq_dv = 0.0
                        dq_dv_voltage = 0.0
                        dq_dv_valid = False
                        
                        # Get measurements
                        mv = self.master.read_voltage()
                        mi = self.master.read_current()
                        mp = self.master.read_power()
                        sv = self.slave.read_voltage()
                        si = self.slave.read_current()
                        sp = self.slave.read_power()

                        # Calculate totals
                        if mp is not None and sp is not None:
                            total_power = mp + sp
                        elif mp is not None:
                            total_power = mp
                        elif sp is not None:
                            total_power = sp
                        
                        if mi is not None and si is not None:
                            total_current = mi + si
                        elif mi is not None: 
                            total_current = mi
                        elif si is not None:
                            total_current = si
                        
                        if mv is not None and sv is not None:
                            avg_voltage = (mv + sv) / 2.0
                        elif mv is not None:
                            avg_voltage = mv
                        elif sv is not None:
                            avg_voltage = sv

                        # Calculate cumulative energy and capacity
                        current_time = time.time()
                        if self.last_measurement_time is not None:
                            delta_time_seconds = current_time - self.last_measurement_time
                            delta_time_hours = delta_time_seconds / 3600.0
                            
                            energy_increment, capacity_increment = self._calculate_energy_increment(
                                total_power, total_current, delta_time_hours
                            )
                            
                            self.cumulative_energy_wh += energy_increment
                            self.cumulative_capacity_ah += capacity_increment
                        
                        self.last_measurement_time = current_time
                        self.last_total_power = total_power
                        self.last_total_current = total_current

                        # Calculate differential capacity dQ/dV
                        dq_dv, dq_dv_voltage, dq_dv_valid = self._calculate_dq_dv(
                            avg_voltage, self.cumulative_capacity_ah
                        )

                        self._log(f"  M:   V={mv}V I={mi}A P={mp}W | S:   V={sv}V I={si}A P={sp}W | "
                                 f"Total:   {total_power:.2f}W, {total_current:.2f}A | "
                                 f"Energy:  {self.cumulative_energy_wh:.3f}Wh, Cap: {self.cumulative_capacity_ah:.3f}Ah | "
                                 f"dQ/dV: {dq_dv:.4f} Ah/V @ {dq_dv_voltage:.2f}V")

                        # Check transition phase
                        if time.time() - transition_start > self.settings.transition_duration:
                            if current_imbalance_limit != self.settings.imbalance_limit:
                                current_imbalance_limit = self.settings.imbalance_limit
                                self._log(f"  → Normal phase:  {current_imbalance_limit*100}% imbalance threshold")

                        # Calculate imbalances
                        power_imb, current_imb, voltage_imb = self._calculate_imbalances(mp, sp, mi, si, mv, sv)

                        # Check imbalances ONLY if enabled
                        if self.settings. imbalance_check_enabled: 
                            # Check imbalances (power)
                            if mp is not None and sp is not None and mp > 0 and sp > 0:
                                if power_imb > current_imbalance_limit * 100:
                                    reason = f"Power imbalance >{current_imbalance_limit*100}%:  M={mp}W S={sp}W ({power_imb:.1f}%)"
                                    self._log(f"[ALERT] {reason}")
                                    self._safe_shutdown(reason)
                                    self.request_stop()
                                    break

                            # Check imbalances (current)
                            if mi is not None and si is not None and mi > 0 and si > 0:
                                if current_imb > current_imbalance_limit * 100:
                                    reason = f"Current imbalance >{current_imbalance_limit*100}%: M={mi}A S={si}A ({current_imb:.1f}%)"
                                    self._log(f"[ALERT] {reason}")
                                    self._safe_shutdown(reason)
                                    self.request_stop()
                                    break

                            # Check imbalances (voltage)
                            if mv is not None and sv is not None and mv > 0 and sv > 0:
                                if voltage_imb > self.settings.voltage_diff_limit * 100:
                                    reason = f"Voltage imbalance >{self.settings.voltage_diff_limit*100}%: M={mv}V S={sv}V ({voltage_imb:.1f}%)"
                                    self._log(f"[ALERT] {reason}")
                                    self._safe_shutdown(reason)
                                    self.request_stop()
                                    break
                        else:
                            # Log high imbalance as warning only (not shutdown)
                            if power_imb > 50:  # Log if very high
                                self._log(f"[WARNING] High power imbalance: {power_imb:.1f}% (monitoring disabled)")
                            if current_imb > 50:
                                self._log(f"[WARNING] High current imbalance: {current_imb:. 1f}% (monitoring disabled)")
                            if voltage_imb > 20:
                                self._log(f"[WARNING] High voltage difference: {voltage_imb:.1f}% (monitoring disabled)")

                        # Check voltage cutoff
                        if self.settings.cutoff_safety_enabled: 
                            if (mv is not None and mv <= self.settings.cutoff_voltage) or \
                               (sv is not None and sv <= self. settings.cutoff_voltage):
                                reason = f"Voltage cutoff: M={mv}V S={sv}V <= {self.settings.cutoff_voltage}V"
                                self._log(f"[CUTOFF] {reason}")
                                self._safe_shutdown(reason)
                                self.request_stop()
                                break

                        # Write to CSV with ALL parameters
                        if self.csv_writer:
                            self._maybe_rotate_csv()
                            self._write_csv_row(idx, total_set_power, mv, mi, mp, sv, si, sp, 
                                              power_imb, current_imb, voltage_imb,
                                              total_power, total_current, avg_voltage,
                                              self.cumulative_energy_wh, self.cumulative_capacity_ah,
                                              dq_dv, dq_dv_voltage)

                        # Push to plot queue
                        self._push_plot_point(idx, mv, mi, mp, sv, si, sp,
                                            self.cumulative_energy_wh, self.cumulative_capacity_ah,
                                            dq_dv, dq_dv_voltage, dq_dv_valid)

                        # At the end of the loop, check for stop
                        for _ in range(max(1, int(self.record_interval_s))):
                            if self._stop_event.is_set():
                                self._log("⏸ Stop requested by user")
                                break
                            time.sleep(1)
                        
                        # If stop was requested, break out of main loop
                        if self._stop_event.is_set():
                            break
                    
                    # If we exited due to stop request, perform safe shutdown
                    if self._stop_event.is_set():
                        self._safe_shutdown("Stop requested by user")
                        return  # Exit the run method

                else:
                    # Fixed duration
                    unit_seconds = dict(UNITS)[row.duration_unit]
                    dur_seconds = float(row.duration_value) * unit_seconds
                    
                    while (time.time() - step_start) < dur_seconds and not self._stop_event.is_set():
                        # Initialize all variables at the start
                        total_power = 0.0
                        total_current = 0.0
                        avg_voltage = 0.0
                        power_imb = 0.0
                        current_imb = 0.0
                        voltage_imb = 0.0
                        dq_dv = 0.0
                        dq_dv_voltage = 0.0
                        dq_dv_valid = False
                        
                        elapsed = int(time.time() - step_start)
                        
                        # Get measurements
                        mv = self.master.read_voltage()
                        mi = self.master.read_current()
                        mp = self.master.read_power()
                        sv = self.slave.read_voltage()
                        si = self.slave. read_current()
                        sp = self.slave.read_power()

                        # Calculate totals
                        if mp is not None and sp is not None:
                            total_power = mp + sp
                        elif mp is not None:
                            total_power = mp
                        elif sp is not None:
                            total_power = sp
                        
                        if mi is not None and si is not None:
                            total_current = mi + si
                        elif mi is not None:
                            total_current = mi
                        elif si is not None: 
                            total_current = si
                        
                        if mv is not None and sv is not None: 
                            avg_voltage = (mv + sv) / 2.0
                        elif mv is not None:
                            avg_voltage = mv
                        elif sv is not None:
                            avg_voltage = sv

                        # Calculate cumulative energy and capacity
                        current_time = time.time()
                        if self.last_measurement_time is not None: 
                            delta_time_seconds = current_time - self.last_measurement_time
                            delta_time_hours = delta_time_seconds / 3600.0
                            
                            energy_increment, capacity_increment = self._calculate_energy_increment(
                                total_power, total_current, delta_time_hours
                            )
                            
                            self.cumulative_energy_wh += energy_increment
                            self.cumulative_capacity_ah += capacity_increment
                        
                        self.last_measurement_time = current_time
                        self.last_total_power = total_power
                        self.last_total_current = total_current

                        # Calculate differential capacity dQ/dV
                        dq_dv, dq_dv_voltage, dq_dv_valid = self._calculate_dq_dv(
                            avg_voltage, self. cumulative_capacity_ah
                        )

                        self._log(f"  [{elapsed}s/{int(dur_seconds)}s] M:  V={mv}V I={mi}A P={mp}W | S:  V={sv}V I={si}A P={sp}W | "
                                 f"Total:  {total_power:.2f}W, {total_current:.2f}A | "
                                 f"Energy: {self.cumulative_energy_wh:.3f}Wh, Cap: {self.cumulative_capacity_ah:.3f}Ah | "
                                 f"dQ/dV:  {dq_dv:.4f} Ah/V @ {dq_dv_voltage:.2f}V")

                        # Check transition phase
                        if time.time() - transition_start > self.settings.transition_duration:
                            if current_imbalance_limit != self.settings.imbalance_limit:
                                current_imbalance_limit = self.settings.imbalance_limit
                                self._log(f"  → Normal phase: {current_imbalance_limit*100}% imbalance threshold")

                        # Calculate imbalances
                        power_imb, current_imb, voltage_imb = self._calculate_imbalances(mp, sp, mi, si, mv, sv)

                        # Check imbalances ONLY if enabled
                        if self.settings. imbalance_check_enabled: 
                            # Check imbalances (power)
                            if mp is not None and sp is not None and mp > 0 and sp > 0:
                                if power_imb > current_imbalance_limit * 100:
                                    reason = f"Power imbalance >{current_imbalance_limit*100}%:  M={mp}W S={sp}W ({power_imb:.1f}%)"
                                    self._log(f"[ALERT] {reason}")
                                    self._safe_shutdown(reason)
                                    self.request_stop()
                                    break

                            # Check imbalances (current)
                            if mi is not None and si is not None and mi > 0 and si > 0:
                                if current_imb > current_imbalance_limit * 100:
                                    reason = f"Current imbalance >{current_imbalance_limit*100}%: M={mi}A S={si}A ({current_imb:.1f}%)"
                                    self._log(f"[ALERT] {reason}")
                                    self._safe_shutdown(reason)
                                    self.request_stop()
                                    break

                            # Check imbalances (voltage)
                            if mv is not None and sv is not None and mv > 0 and sv > 0:
                                if voltage_imb > self.settings.voltage_diff_limit * 100:
                                    reason = f"Voltage imbalance >{self.settings.voltage_diff_limit*100}%: M={mv}V S={sv}V ({voltage_imb:.1f}%)"
                                    self._log(f"[ALERT] {reason}")
                                    self._safe_shutdown(reason)
                                    self.request_stop()
                                    break
                        else:
                            # Log high imbalance as warning only (not shutdown)
                            if power_imb > 50:  # Log if very high
                                self._log(f"[WARNING] High power imbalance: {power_imb:.1f}% (monitoring disabled)")
                            if current_imb > 50:
                                self._log(f"[WARNING] High current imbalance: {current_imb:.1f}% (monitoring disabled)")
                            if voltage_imb > 20:
                                self._log(f"[WARNING] High voltage difference: {voltage_imb:.1f}% (monitoring disabled)")

                        # Check voltage cutoff
                        if self.settings.cutoff_safety_enabled:
                            if (mv is not None and mv <= self.settings.cutoff_voltage) or \
                               (sv is not None and sv <= self.settings.cutoff_voltage):
                                reason = f"Voltage cutoff: M={mv}V S={sv}V <= {self.settings.cutoff_voltage}V"
                                self._log(f"[CUTOFF] {reason}")
                                self._safe_shutdown(reason)
                                self.request_stop()
                                break

                        # Write to CSV
                        if self.csv_writer:
                            self._maybe_rotate_csv()
                            self._write_csv_row(idx, total_set_power, mv, mi, mp, sv, si, sp,
                                              power_imb, current_imb, voltage_imb,
                                              total_power, total_current, avg_voltage,
                                              self.cumulative_energy_wh, self.cumulative_capacity_ah,
                                              dq_dv, dq_dv_voltage)

                        # Push to plot
                        self._push_plot_point(idx, mv, mi, mp, sv, si, sp,
                                            self.cumulative_energy_wh, self. cumulative_capacity_ah,
                                            dq_dv, dq_dv_voltage, dq_dv_valid)

                        for _ in range(max(1, int(self.record_interval_s))):
                            if self._stop_event.is_set():
                                self._log("⏸ Stop requested by user")
                                break
                            time.sleep(1)
                        
                        # If stop was requested, break out
                        if self._stop_event.is_set():
                            break
                    
                    # If we exited due to stop request, perform safe shutdown
                    if self._stop_event.is_set():
                        self._safe_shutdown("Stop requested by user")
                        return  # Exit the run method

                # Step complete - turn off for this step
                try:
                    self. master.set_power(0.0)
                    self.slave.set_power(0.0)
                    time.sleep(0.5)
                    self.master.load_off()
                    self.slave.load_off()
                    self._log(f"✓ Step {idx} complete; Power set to 0, LOADS OFF.")
                except Exception as e:  
                    self._log(f"[WARN] Could not turn LOADS OFF after step {idx}: {e}")

            # ALL STEPS COMPLETED - Final safe shutdown
            self._log("═══ All Profile Steps Completed ═══")
            self._safe_shutdown("Mission profile completed successfully")

        except Exception as exc:
            self._log(f"[EXCEPTION] Worker:  {exc}")
            self._log(traceback.format_exc())
            self._safe_shutdown(f"Exception occurred: {exc}")
        finally:
            # Always ensure safe shutdown happened
            try:
                # Double-check everything is off and local
                for controller in [self.master, self. slave]:
                    try:
                        controller.set_power(0.0)
                        controller.load_off()
                        controller.input_off()
                        controller.parallel_off()
                        controller.set_local()
                    except Exception: 
                        pass
            except Exception:
                pass
            
            # Close CSV
            try:
                self._close_csv()
            except Exception:
                pass
            
            # NEW: Close log file
            try:
                self._close_log_file()
            except Exception: 
                pass
            
            # Close serial connections
            try:
                self. master.close()
                self. slave.close()
            except Exception:
                pass
            
            self._log("=== Worker Finished ===")

    def _calculate_imbalances(self, mp, sp, mi, si, mv, sv):
        power_imb = 0.0
        current_imb = 0.0
        voltage_imb = 0.0

        if mp is not None and sp is not None: 
            avg_power = (mp + sp) / 2.0
            if avg_power > 0:
                power_imb = abs(mp - sp) / avg_power * 100

        if mi is not None and si is not None:
            avg_current = (mi + si) / 2.0
            if avg_current > 0:
                current_imb = abs(mi - si) / avg_current * 100

        if mv is not None and sv is not None:
            avg_voltage = (mv + sv) / 2.0
            if avg_voltage > 0:
                voltage_imb = abs(mv - sv) / avg_voltage * 100

        return power_imb, current_imb, voltage_imb


# --- Main Application ---
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Master-Slave Electronic Load - Profile Editor, CSV Logger & Real-Time Plot")
        self.geometry("1400x960")

        # Settings
        self.settings = Settings()
        self.settings.profile = []
        for p in DEFAULT_PROFILE:
            power = float(p[0])
            hours = float(p[1])
            self.settings.profile.append(ProfileRow(total_power=power, duration_value=hours, duration_unit="hours", cutoff=False))

        # Queues and worker
        self.log_queue:  queue.Queue = queue.Queue()
        self.plot_queue: queue.Queue = queue.Queue()
        self.worker:  Optional[MasterSlaveWorkerThread] = None

        # State
        self.row_widgets:  List[ProfileRowWidget] = []
        self. csv_folder: Optional[str] = None
        self.config_save_folder: Optional[str] = None
        self.config_filename_var = tk.StringVar(value="master_slave_config.txt")
        self.config_load_path_var = tk.StringVar(value="")
        self.rotate_hours_var = tk. DoubleVar(value=DEFAULT_ROTATE_HOURS)
        self.record_interval_var = tk.IntVar(value=DEFAULT_RECORD_INTERVAL_S)
        self.cutoff_safety_var = tk.BooleanVar(value=self.settings.cutoff_safety_enabled)
        self.imbalance_limit_var = tk.DoubleVar(value=DEFAULT_IMBALANCE_LIMIT * 100)
        self.imbalance_transition_var = tk.DoubleVar(value=DEFAULT_IMBALANCE_TRANSITION * 100)
        self.transition_duration_var = tk.IntVar(value=DEFAULT_TRANSITION_DURATION)
        self.voltage_diff_limit_var = tk.DoubleVar(value=DEFAULT_VOLTAGE_DIFF_LIMIT * 100)
        self.cutoff_safety_var = tk.BooleanVar(value=self.settings.cutoff_safety_enabled)
        self.imbalance_check_enabled_var = tk.BooleanVar(value=True)  # NEW: Enable imbalance checking by default
        self.imbalance_limit_var = tk.DoubleVar(value=DEFAULT_IMBALANCE_LIMIT * 100)

        # Email settings
        self.email_enabled_var = tk.BooleanVar(value=False)
        self.email_sender_var = tk.StringVar(value="")
        self.email_receiver_var = tk.StringVar(value="")
        self.email_password_var = tk.StringVar(value="")

        # Plot data arrays (totals only)
        self.plot_times:  List[float] = []
        self.plot_total_power: List[float] = []
        self.plot_total_current: List[float] = []
        self.plot_avg_voltage: List[float] = []
        self.plot_cumulative_energy: List[float] = []
        self.plot_cumulative_capacity: List[float] = []
        
        # NEW: Differential capacity plot data (dQ/dV vs Voltage)
        self.plot_dqdv_voltage: List[float] = []  # X-axis: Voltage
        self.plot_dqdv_values: List[float] = []   # Y-axis: dQ/dV
        
        # Style controls for each line
        self.dqdv_color_var = tk.StringVar(value="darkred")
        self.dqdv_width_var = tk. DoubleVar(value=2.0)
        self.dqdv_style_var = tk.StringVar(value="solid")
        self.dqdv_marker_var = tk.StringVar(value="circle")  

        self.power_color_var = tk.StringVar(value="blue")
        self.power_width_var = tk.DoubleVar(value=3.0)
        self.power_style_var = tk.StringVar(value="solid")
        
        self.current_color_var = tk.StringVar(value="red")
        self.current_width_var = tk.DoubleVar(value=2.0)
        self.current_style_var = tk.StringVar(value="dashed")
        
        self.voltage_color_var = tk.StringVar(value="green")
        self.voltage_width_var = tk.DoubleVar(value=2.0)
        self.voltage_style_var = tk.StringVar(value="solid")
        
        self.energy_color_var = tk.StringVar(value="purple")  
        self.energy_width_var = tk.DoubleVar(value=2.0)  
        self. energy_style_var = tk.StringVar(value="solid")  

        # Status
        self.status_top_lines = deque(maxlen=20)

        # Config note
        self.config_note_fontsize = tk.IntVar(value=14)

        # Build UI
        self._build_ui()
        self._periodic_log_poll()
        self._periodic_plot_poll()

    def _build_ui(self):
        # Create main notebook for tabs
        main_notebook = ttk.Notebook(self)
        main_notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # === TAB 1: Configuration & Control ===
        config_tab = ttk.Frame(main_notebook)
        main_notebook.add(config_tab, text="Configuration & Control")
        
        # Top area (inside config tab)
        top_frame = ttk.Frame(config_tab)
        top_frame.pack(side=tk.TOP, fill=tk.X, padx=8, pady=6)
    
        # Settings frame (left)
        settings_fr = ttk. Labelframe(top_frame, text="Master-Slave Settings", padding=(8, 6))
        settings_fr.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 8))
    
        ttk.Label(settings_fr, text="Master COM: ").grid(row=0, column=0, sticky="w")
        self.master_port_entry = ttk.Entry(settings_fr, width=12)
        self.master_port_entry.insert(0, self.settings.master_port)
        self.master_port_entry.grid(row=0, column=1, pady=2)
    
        ttk.Label(settings_fr, text="Slave COM:").grid(row=1, column=0, sticky="w")
        self.slave_port_entry = ttk. Entry(settings_fr, width=12)
        self.slave_port_entry.insert(0, self.settings.slave_port)
        self.slave_port_entry.grid(row=1, column=1, pady=2)
    
        ttk.Label(settings_fr, text="Baud: ").grid(row=2, column=0, sticky="w")
        self.baud_entry = ttk.Entry(settings_fr, width=12)
        self.baud_entry.insert(0, str(self.settings.baud_rate))
        self.baud_entry.grid(row=2, column=1, pady=2)
    
        self.dry_var = tk.BooleanVar(value=self.settings.dry_run)
        ttk.Checkbutton(settings_fr, text="Dry run", variable=self.dry_var).grid(row=3, column=0, columnspan=2, sticky="w", pady=2)
    
        ttk.Label(settings_fr, text="Cutoff Voltage (V):").grid(row=4, column=0, sticky="w")
        self.cutoff_entry = ttk.Entry(settings_fr, width=12)
        self.cutoff_entry. insert(0, str(self. settings.cutoff_voltage))
        self.cutoff_entry. grid(row=4, column=1, pady=2)
        ttk. Checkbutton(settings_fr, text="Enable", variable=self.cutoff_safety_var).grid(row=4, column=2, padx=6)
    
        ttk. Separator(settings_fr, orient=tk.HORIZONTAL).grid(row=5, column=0, columnspan=3, sticky="ew", pady=6)

        # NEW: Imbalance Check Enable/Disable
        ttk.Label(settings_fr, text="Imbalance Monitoring:", font=('Arial', 9, 'bold')).grid(row=6, column=0, columnspan=3, sticky="w", pady=(5, 2))
        ttk.Checkbutton(settings_fr, text="Enable Imbalance Shutdown", 
                       variable=self.imbalance_check_enabled_var,
                       command=self._toggle_imbalance_controls).grid(row=7, column=0, columnspan=3, sticky="w", pady=2)

        ttk.Label(settings_fr, text="Imbalance Limit (%):").grid(row=8, column=0, sticky="w")
        self.imbalance_limit_entry = ttk.Entry(settings_fr, textvariable=self.imbalance_limit_var, width=12)
        self.imbalance_limit_entry.grid(row=8, column=1, pady=2)

        ttk.Label(settings_fr, text="Transition Limit (%):").grid(row=9, column=0, sticky="w")
        self.imbalance_transition_entry = ttk.Entry(settings_fr, textvariable=self.imbalance_transition_var, width=12)
        self.imbalance_transition_entry.grid(row=9, column=1, pady=2)

        ttk.Label(settings_fr, text="Transition Time (s):").grid(row=10, column=0, sticky="w")
        self.transition_duration_entry = ttk.Entry(settings_fr, textvariable=self.transition_duration_var, width=12)
        self.transition_duration_entry. grid(row=10, column=1, pady=2)

        ttk.Label(settings_fr, text="Voltage Diff Limit (%):").grid(row=11, column=0, sticky="w")
        self.voltage_diff_entry = ttk.Entry(settings_fr, textvariable=self.voltage_diff_limit_var, width=12)
        self.voltage_diff_entry.grid(row=11, column=1, pady=2)

        ttk. Separator(settings_fr, orient=tk.HORIZONTAL).grid(row=12, column=0, columnspan=3, sticky="ew", pady=6)
    
        ttk.Label(settings_fr, text="CSV Folder:").grid(row=13, column=0, sticky="w")
        self.csv_folder_label = ttk.Label(settings_fr, text="(not set - REQUIRED! )", 
                                          width=24, foreground="red", font=('Arial', 9, 'bold'))
        self.csv_folder_label.grid(row=13, column=1, sticky="w")
        ttk.Button(settings_fr, text="Select.. .", command=self._select_csv_folder).grid(row=13, column=2, padx=2, pady=2)
    
        ttk.Label(settings_fr, text="Rotate CSV (hours):").grid(row=14, column=0, sticky="w")
        ttk. Spinbox(settings_fr, from_=0, to=168, increment=0.5, textvariable=self.rotate_hours_var, width=10).grid(row=14, column=1, pady=2)
    
        ttk.Label(settings_fr, text="Record interval (s):").grid(row=15, column=0, sticky="w")
        ttk.Spinbox(settings_fr, from_=1, to=3600, textvariable=self.record_interval_var, width=10).grid(row=15, column=1, pady=2)
    
        ttk.Separator(settings_fr, orient=tk.HORIZONTAL).grid(row=16, column=0, columnspan=3, sticky="ew", pady=6)
    
        # Email settings
        ttk. Checkbutton(settings_fr, text="Enable Email Alerts", variable=self.email_enabled_var).grid(row=17, column=0, columnspan=2, sticky="w")
        
        ttk.Label(settings_fr, text="Sender Email:").grid(row=18, column=0, sticky="w")
        ttk.Entry(settings_fr, textvariable=self.email_sender_var, width=20).grid(row=18, column=1, columnspan=2, pady=2)
    
        ttk.Label(settings_fr, text="Receiver Email: ").grid(row=19, column=0, sticky="w")
        ttk.Entry(settings_fr, textvariable=self.email_receiver_var, width=20).grid(row=19, column=1, columnspan=2, pady=2)
    
        ttk.Label(settings_fr, text="Password:").grid(row=20, column=0, sticky="w")
        ttk.Entry(settings_fr, textvariable=self.email_password_var, show="*", width=20).grid(row=20, column=1, columnspan=2, pady=2)
    
        ttk.Button(settings_fr, text="Apply Settings", command=self._apply_settings).grid(row=21, column=0, columnspan=3, pady=(8, 2), sticky="ew")
    
        # Config save/load
        ttk. Separator(settings_fr, orient=tk. HORIZONTAL).grid(row=22, column=0, columnspan=3, sticky="ew", pady=6)
        ttk.Label(settings_fr, text="Config Filename:").grid(row=23, column=0, sticky="w")
        ttk.Entry(settings_fr, textvariable=self.config_filename_var, width=20).grid(row=23, column=1, columnspan=2, pady=2, sticky="ew")
    
        ttk.Label(settings_fr, text="Save Folder:").grid(row=24, column=0, sticky="w")
        self.config_save_folder_label = ttk.Label(settings_fr, text="(not set)", width=24)
        self.config_save_folder_label.grid(row=24, column=1, sticky="w")
        ttk.Button(settings_fr, text="Choose...", command=self._select_config_save_folder).grid(row=24, column=2, padx=2, pady=2)
    
        ttk.Label(settings_fr, text="Load Config:").grid(row=26, column=0, sticky="w")
        ttk.Entry(settings_fr, textvariable=self.config_load_path_var, width=20).grid(row=26, column=1, pady=2, sticky="ew")
        ttk.Button(settings_fr, text="Browse...", command=self._select_config_file).grid(row=26, column=2, padx=2, pady=2)
    
        ttk.Button(settings_fr, text="Load Configuration", command=self._load_configuration).grid(row=27, column=0, columnspan=3, pady=(4, 2), sticky="ew")
        ttk.Button(settings_fr, text="Save Configuration", command=self._save_configuration).grid(row=25, column=0, columnspan=3, pady=(4, 2), sticky="ew")
    
        # Profile editor (center)
        profile_fr = ttk. Labelframe(top_frame, text="Power Profile (Total Power - 50/50 Split)", padding=(8, 6))
        profile_fr.pack(side=tk.LEFT, fill=tk. BOTH, expand=True)
    
        toolbar = ttk.Frame(profile_fr)
        toolbar.pack(side=tk.TOP, fill=tk.X)
        ttk.Button(toolbar, text="Add Row", command=self._add_row).pack(side=tk.LEFT, padx=4)
        ttk.Button(toolbar, text="Clear All", command=self._clear_all_rows).pack(side=tk.LEFT, padx=4)

        canvas = tk.Canvas(profile_fr, height=360)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk. Scrollbar(profile_fr, orient="vertical", command=canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.configure(yscrollcommand=scrollbar.set)
        self.rows_container = ttk.Frame(canvas)
        self.rows_container.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.rows_container, anchor="nw")
    
        for r in self.settings.profile:
            self._append_row_widget(r)

        # Control buttons
        control_button_frame = ttk. Frame(top_frame)
        control_button_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)
        
        ttk.Label(control_button_frame, text="Test Control:", font=('Arial', 11, 'bold')).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        
        self.start_btn = ttk.Button(control_button_frame, text="▶ Start Test", 
                                     command=self._start, 
                                     width=20)
        self.start_btn. grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        
        self.stop_btn = ttk.Button(control_button_frame, text="⏹ Stop Test", 
                                    command=self._stop, 
                                    state=tk.DISABLED, 
                                    width=20)
        self.stop_btn.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        
        # Make columns expand equally
        control_button_frame. columnconfigure(0, weight=1)
        control_button_frame.columnconfigure(1, weight=1)
    
        # # Control (right)
        # control_fr = ttk.Labelframe(top_frame, text="Control", padding=(8, 6))
        # control_fr.pack(side=tk.LEFT, fill=tk.Y, padx=(8, 0))
    
        # self.start_btn = ttk.Button(control_fr, text="Start", width=16, command=self._start)
        # self.start_btn. pack(pady=6)
        # self.stop_btn = ttk.Button(control_fr, text="Stop", width=16, command=self._stop, state=tk.DISABLED)
        # self.stop_btn.pack(pady=4)
        # self.status_label = ttk.Label(control_fr, text="Idle", foreground="blue")
        # self.status_label.pack(pady=6)
        # self.status_label = ttk.Label(control_fr, text="Idle", foreground="blue")
        # self.status_label.pack(pady=6)
        
        # === Test Logs Section (renamed from Status) ===
        log_frame = ttk.LabelFrame(top_frame, text="Test Logs", padding=(10, 10))
        log_frame.pack(side=tk. BOTTOM, fill=tk. BOTH, expand=True, padx=10, pady=(5, 10))

        # Status indicator at top
        status_display_frame = ttk.Frame(log_frame)
        status_display_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 5))
        
        ttk.Label(status_display_frame, text="Status:", font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=(0, 10))
        self.status_label = ttk.Label(status_display_frame, text="Idle", font=('Arial', 10, 'bold'), foreground="blue")
        self.status_label.pack(side=tk.LEFT)

        # Log text area with scrollbar
        log_scroll_frame = ttk.Frame(log_frame)
        log_scroll_frame.pack(fill=tk.BOTH, expand=True)
        
        log_scrollbar = ttk.Scrollbar(log_scroll_frame)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.log_text = tk.Text(log_scroll_frame, height=15, wrap=tk.WORD, 
                                yscrollcommand=log_scrollbar.set,
                                font=('Consolas', 9))
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scrollbar.config(command=self.log_text.yview)
        
        # Log controls
        log_control_frame = ttk.Frame(log_frame)
        log_control_frame.pack(side=tk. BOTTOM, fill=tk.X, pady=(5, 0))
        
        ttk. Button(log_control_frame, text="Clear Logs", 
                   command=self._clear_log_display).pack(side=tk.LEFT, padx=2)
        ttk.Label(log_control_frame, text="(Logs are automatically saved to CSV)", 
                  font=('Arial', 8, 'italic'), foreground='gray').pack(side=tk.LEFT, padx=10)
    
        # # Status (top-right)
        # status_fr = ttk. Labelframe(top_frame, text="Test Logs", padding=(8, 6))
        # status_fr.pack(side=tk.RIGHT, fill=tk.Y)
        # self.status_top_widget = scrolledtext.ScrolledText(status_fr, wrap=tk.WORD, state=tk.DISABLED, width=40, height=12)
        # self.status_top_widget.pack(fill=tk. BOTH, expand=True)
    
        # Config note in config tab
        note_fr = ttk.Labelframe(config_tab, text="Configuration Note", padding=(6, 6))
        note_fr.pack(side=tk.TOP, fill=tk.X, padx=8, pady=(6, 0))
        
        note_inner = ttk.Frame(note_fr)
        note_inner.pack(fill=tk.BOTH, expand=True)
        
        self.config_note_text = tk.Text(note_inner, wrap=tk.WORD, height=6)
        self.config_note_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.config_note_text.configure(state=tk.DISABLED)
    
        fs_fr = ttk.Frame(note_fr)
        fs_fr.pack(fill=tk.X, pady=(6, 0))
        ttk.Label(fs_fr, text="Font size:").pack(side=tk.LEFT)
        ttk.Spinbox(fs_fr, from_=8, to=48, increment=1, textvariable=self.config_note_fontsize, width=5, command=self._apply_note_font).pack(side=tk.LEFT, padx=6)
        ttk.Button(fs_fr, text="Refresh", command=self._update_config_note).pack(side=tk.LEFT, padx=4)
    
        # # Log area in config tab
        # log_fr = ttk.Labelframe(config_tab, text="Log (Central Time)", padding=(8, 6))
        # log_fr.pack(side=tk. TOP, fill=tk. BOTH, expand=True, padx=8, pady=(6, 8))
        # self.log_widget = scrolledtext.ScrolledText(log_fr, wrap=tk.WORD, state=tk. DISABLED)
        # self.log_widget. pack(fill=tk.BOTH, expand=True)
    
        # === TAB 2: Real-Time Plots ===
        plot_tab = ttk.Frame(main_notebook)
        main_notebook.add(plot_tab, text="Real-Time Plots")
        
        # Plot controls at top
        plot_control_frame = ttk.Frame(plot_tab)
        plot_control_frame.pack(side=tk.TOP, fill=tk.X, padx=8, pady=6)
        
        ttk. Label(plot_control_frame, text="Master-Slave Combined Real-Time Monitoring", 
                  font=('Arial', 14, 'bold')).pack(side=tk.LEFT, padx=10)
        
        ttk.Button(plot_control_frame, text="Clear Plot", 
                   command=self._clear_plots).pack(side=tk.RIGHT, padx=5)
        ttk.Button(plot_control_frame, text="Export Plot Image", 
                   command=self._save_plot_screenshot).pack(side=tk.RIGHT, padx=5)
        ttk.Button(plot_control_frame, text="Export Plot Data (CSV)", 
                   command=self._export_plot_data).pack(side=tk.RIGHT, padx=5)
        ttk.Button(plot_control_frame, text="Debug Plot", 
                   command=self._debug_plot_data).pack(side=tk.RIGHT, padx=5)
        
        # Main plot area
        plot_main_frame = ttk.Frame(plot_tab)
        plot_main_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=8, pady=6)

        # Create single plot with multiple y-axes
        self.fig = Figure(figsize=(12, 7), dpi=100)
        self.ax_power = self.fig.add_subplot(111)
        
        # Create secondary y-axes
        self.ax_current = self.ax_power.twinx()
        self.ax_voltage = self.ax_power.twinx()
        self.ax_energy = self.ax_power.twinx()
        
        # Offset the right spines
        self.ax_voltage.spines['right'].set_position(('outward', 60))
        self.ax_energy.spines['right'].set_position(('outward', 120))
        
        # Set axis z-orders and make backgrounds transparent
        self.ax_power.set_zorder(1)
        self.ax_current.set_zorder(2)
        self.ax_voltage.set_zorder(3)
        self.ax_energy. set_zorder(4)
        
        # Make all backgrounds transparent except base
        self.ax_current.patch.set_visible(False)
        self.ax_voltage.patch.set_visible(False)
        self.ax_energy.patch.set_visible(False)
        
        # Disable offset/scientific notation on all axes
        self.ax_power.ticklabel_format(style='plain', axis='y', useOffset=False)
        self.ax_current.ticklabel_format(style='plain', axis='y', useOffset=False)
        self.ax_voltage.ticklabel_format(style='plain', axis='y', useOffset=False)
        self.ax_energy.ticklabel_format(style='plain', axis='y', useOffset=False)
        
        # Labels
        self.ax_power. set_xlabel("Time (s)", fontsize=11, fontweight='bold')
        self.ax_power.set_ylabel("Total Power (W)", fontsize=11, fontweight='bold', 
                                  color=self.power_color_var.get())
        self.ax_current.set_ylabel("Total Current (A)", fontsize=11, fontweight='bold', 
                                    color=self.current_color_var.get())
        self.ax_voltage.set_ylabel("Avg Voltage (V)", fontsize=11, fontweight='bold', 
                                    color=self.voltage_color_var. get())
        self.ax_energy.set_ylabel("Energy (Wh)", fontsize=11, fontweight='bold',
                                   color=self.energy_color_var.get())
        
        # Tick colors
        self.ax_power. tick_params(axis='y', labelcolor=self.power_color_var.get(), labelsize=9)
        self.ax_current.tick_params(axis='y', labelcolor=self.current_color_var.get(), labelsize=9)
        self.ax_voltage.tick_params(axis='y', labelcolor=self. voltage_color_var.get(), labelsize=9)
        self.ax_energy.tick_params(axis='y', labelcolor=self.energy_color_var. get(), labelsize=9)
        
        # Grid (only on primary axis)
        self.ax_power.grid(True, alpha=0.3, linestyle='--')
        
        # Create line objects - ALL with distinct colors and z-orders
        (self. line_power,) = self.ax_power.plot([], [], 
                                                 color=self.power_color_var. get(), 
                                                 linestyle=LINE_STYLES[self.power_style_var. get()], 
                                                 linewidth=self.power_width_var.get(),
                                                 marker='o', markersize=3,
                                                 label="Total Power (W)",
                                                 zorder=10)
        
        (self.line_current,) = self.ax_current.plot([], [], 
                                                     color=self.current_color_var.get(), 
                                                     linestyle=LINE_STYLES[self.current_style_var.get()], 
                                                     linewidth=self.current_width_var.get(),
                                                     marker='s', markersize=3,
                                                     label="Total Current (A)",
                                                     zorder=9)
        
        (self. line_voltage,) = self.ax_voltage.plot([], [], 
                                                     color=self.voltage_color_var. get(), 
                                                     linestyle=LINE_STYLES[self.voltage_style_var. get()], 
                                                     linewidth=self.voltage_width_var.get(),
                                                     marker='^', markersize=3,
                                                     label="Avg Voltage (V)",
                                                     zorder=8)
        
        (self.line_energy,) = self.ax_energy.plot([], [], 
                                                   color=self.energy_color_var.get(), 
                                                   linestyle=LINE_STYLES[self.energy_style_var.get()], 
                                                   linewidth=self.energy_width_var. get(),
                                                   marker='d', markersize=3,
                                                   label="Energy (Wh)",
                                                   zorder=7)
        
        # Combined legend
        lines = [self.line_power, self.line_current, self. line_voltage, self.line_energy]
        labels = [l.get_label() for l in lines]
        self.ax_power.legend(lines, labels, loc='upper left', fontsize=9, framealpha=0.9)
        
        self.fig.tight_layout()
        
        # # After creating all axes, set their z-orders
        # self.ax_power.set_zorder(4)  # Base axis
        # self.ax_current. set_zorder(3)
        # self.ax_voltage.set_zorder(2)
        # self.ax_energy.set_zorder(1)
        
        # # Make backgrounds transparent so lines show through
        # self.ax_current.patch.set_visible(False)
        # self.ax_voltage.patch.set_visible(False)
        # self.ax_energy.patch.set_visible(False)
        
        self.fig.tight_layout()

        self.canvas = FigureCanvasTkAgg(self.fig, master=plot_main_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        toolbar_mpl = NavigationToolbar2Tk(self.canvas, plot_main_frame)
        toolbar_mpl.update()
        toolbar_mpl.pack(side=tk. BOTTOM, fill=tk.X)

        # Right panel for plot styles
        plot_right_panel = ttk.Frame(plot_tab)
        plot_right_panel.pack(side=tk. RIGHT, fill=tk.Y, padx=8, pady=6)

        # Create scrollable frame for styles
        style_canvas = tk.Canvas(plot_right_panel, width=200)
        style_scrollbar = ttk.Scrollbar(plot_right_panel, orient="vertical", command=style_canvas. yview)
        style_scrollable_frame = ttk.Frame(style_canvas)
        
        style_scrollable_frame.bind(
            "<Configure>",
            lambda e: style_canvas.configure(scrollregion=style_canvas.bbox("all"))
        )
        
        style_canvas.create_window((0, 0), window=style_scrollable_frame, anchor="nw")
        style_canvas.configure(yscrollcommand=style_scrollbar. set)
        
        style_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        style_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        style_fr = ttk. Labelframe(style_scrollable_frame, text="Line Appearance", padding=(10, 10))
        style_fr.pack(fill=tk. BOTH, expand=True, padx=5, pady=5)

        # # [Keep all existing Power, Current, Voltage controls as before]
        # # Power line controls
        # ttk.Label(style_fr, text="━━━ Total Power ━━━", 
        #           font=('Arial', 10, 'bold')).pack(pady=(5, 10))
        
        # ttk. Label(style_fr, text="Color: ").pack(anchor='w', padx=5)
        # self.power_color_display = tk.Canvas(style_fr, width=100, height=30, 
        #                                       bg=self.power_color_var. get(), 
        #                                       relief=tk.RIDGE, borderwidth=2)
        # self.power_color_display.pack(pady=3)
        # ttk.Button(style_fr, text="Choose Color", 
        #            command=lambda: self._choose_line_color(self.power_color_var, 
        #                                                    self.power_color_display)).pack(pady=3)
        
        # ttk.Label(style_fr, text="Width:").pack(anchor='w', padx=5, pady=(5, 0))
        # ttk.Scale(style_fr, from_=0.5, to=5.0, variable=self.power_width_var, 
        #           orient=tk.HORIZONTAL).pack(fill=tk.X, padx=5, pady=3)
        
        # ttk.Label(style_fr, text="Style:").pack(anchor='w', padx=5, pady=(5, 0))
        # power_style_combo = ttk.Combobox(style_fr, textvariable=self.power_style_var, 
        #                                  values=list(LINE_STYLES.keys()), state='readonly', width=15)
        # power_style_combo.pack(pady=3)
        
        # ttk. Separator(style_fr, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        # # Current line controls
        # ttk.Label(style_fr, text="━━━ Total Current ━━━", 
        #           font=('Arial', 10, 'bold')).pack(pady=(5, 10))
        
        # ttk.Label(style_fr, text="Color:").pack(anchor='w', padx=5)
        # self.current_color_display = tk.Canvas(style_fr, width=100, height=30, 
        #                                         bg=self.current_color_var.get(), 
        #                                         relief=tk.RIDGE, borderwidth=2)
        # self.current_color_display.pack(pady=3)
        # ttk.Button(style_fr, text="Choose Color", 
        #            command=lambda: self._choose_line_color(self.current_color_var, 
        #                                                    self.current_color_display)).pack(pady=3)
        
        # ttk. Label(style_fr, text="Width:").pack(anchor='w', padx=5, pady=(5, 0))
        # ttk.Scale(style_fr, from_=0.5, to=5.0, variable=self.current_width_var, 
        #           orient=tk. HORIZONTAL).pack(fill=tk.X, padx=5, pady=3)
        
        # ttk.Label(style_fr, text="Style:").pack(anchor='w', padx=5, pady=(5, 0))
        # current_style_combo = ttk. Combobox(style_fr, textvariable=self. current_style_var, 
        #                                    values=list(LINE_STYLES.keys()), state='readonly', width=15)
        # current_style_combo.pack(pady=3)
        
        # ttk.Separator(style_fr, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        # # Voltage line controls
        # ttk.Label(style_fr, text="━━━ Avg Voltage ━━━", 
        #           font=('Arial', 10, 'bold')).pack(pady=(5, 10))
        
        # ttk. Label(style_fr, text="Color:").pack(anchor='w', padx=5)
        # self.voltage_color_display = tk.Canvas(style_fr, width=100, height=30, 
        #                                         bg=self.voltage_color_var.get(), 
        #                                         relief=tk.RIDGE, borderwidth=2)
        # self.voltage_color_display.pack(pady=3)
        # ttk.Button(style_fr, text="Choose Color", 
        #            command=lambda: self._choose_line_color(self.voltage_color_var, 
        #                                                    self.voltage_color_display)).pack(pady=3)
        
        # ttk. Label(style_fr, text="Width:").pack(anchor='w', padx=5, pady=(5, 0))
        # ttk.Scale(style_fr, from_=0.5, to=5.0, variable=self.voltage_width_var, 
        #           orient=tk. HORIZONTAL).pack(fill=tk.X, padx=5, pady=3)
        
        # ttk.Label(style_fr, text="Style:").pack(anchor='w', padx=5, pady=(5, 0))
        # voltage_style_combo = ttk.Combobox(style_fr, textvariable=self.voltage_style_var, 
        #                                    values=list(LINE_STYLES.keys()), state='readonly', width=15)
        # voltage_style_combo.pack(pady=3)
        
        # ttk.Separator(style_fr, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        # # NEW: Energy line controls
        # ttk.Label(style_fr, text="━━━ Cumulative Energy ━━━", 
        #           font=('Arial', 10, 'bold')).pack(pady=(5, 10))
        
        # ttk. Label(style_fr, text="Color:").pack(anchor='w', padx=5)
        # self.energy_color_display = tk.Canvas(style_fr, width=100, height=30, 
        #                                        bg=self.energy_color_var.get(), 
        #                                        relief=tk.RIDGE, borderwidth=2)
        # self.energy_color_display.pack(pady=3)
        # ttk.Button(style_fr, text="Choose Color", 
        #            command=lambda: self._choose_line_color(self.energy_color_var, 
        #                                                    self.energy_color_display)).pack(pady=3)
        
        # ttk. Label(style_fr, text="Width:").pack(anchor='w', padx=5, pady=(5, 0))
        # ttk.Scale(style_fr, from_=0.5, to=5.0, variable=self.energy_width_var, 
        #           orient=tk. HORIZONTAL).pack(fill=tk.X, padx=5, pady=3)
        
        # ttk.Label(style_fr, text="Style:").pack(anchor='w', padx=5, pady=(5, 0))
        # energy_style_combo = ttk.Combobox(style_fr, textvariable=self.energy_style_var, 
        #                                   values=list(LINE_STYLES.keys()), state='readonly', width=15)
        # energy_style_combo.pack(pady=3)
        
        # ttk.Separator(style_fr, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=15)
        
        # ttk.Button(style_fr, text="Apply All Styles", command=self._apply_line_styles).pack(pady=10, fill=tk.X, padx=10)

        # Statistics panel
        stats_fr = ttk. Labelframe(plot_right_panel, text="Statistics", padding=(10, 10))
        stats_fr.pack(side=tk.TOP, fill=tk.X, pady=(10, 0))

        ttk.Label(stats_fr, text="Data Points:", font=('Arial', 9, 'bold')).grid(row=0, column=0, sticky='w', pady=2)
        self.stats_points_label = ttk.Label(stats_fr, text="0")
        self.stats_points_label.grid(row=0, column=1, sticky='e', pady=2)

        ttk.Label(stats_fr, text="Elapsed Time:", font=('Arial', 9, 'bold')).grid(row=1, column=0, sticky='w', pady=2)
        self.stats_time_label = ttk.Label(stats_fr, text="0 s")
        self.stats_time_label.grid(row=1, column=1, sticky='e', pady=2)

        ttk.Label(stats_fr, text="Avg Total Power:", font=('Arial', 9, 'bold')).grid(row=2, column=0, sticky='w', pady=2)
        self.stats_avg_power_label = ttk. Label(stats_fr, text="0.0 W")
        self.stats_avg_power_label. grid(row=2, column=1, sticky='e', pady=2)

        ttk.Label(stats_fr, text="Avg Total Current:", font=('Arial', 9, 'bold')).grid(row=3, column=0, sticky='w', pady=2)
        self.stats_avg_current_label = ttk.Label(stats_fr, text="0.0 A")
        self.stats_avg_current_label.grid(row=3, column=1, sticky='e', pady=2)

        ttk.Label(stats_fr, text="Avg Voltage:", font=('Arial', 9, 'bold')).grid(row=4, column=0, sticky='w', pady=2)
        self.stats_avg_voltage_label = ttk. Label(stats_fr, text="0.0 V")
        self.stats_avg_voltage_label.grid(row=4, column=1, sticky='e', pady=2)

        # NEW: Energy statistics
        ttk. Separator(stats_fr, orient=tk.HORIZONTAL).grid(row=5, column=0, columnspan=2, sticky='ew', pady=5)
        
        ttk.Label(stats_fr, text="Total Energy:", font=('Arial', 9, 'bold')).grid(row=6, column=0, sticky='w', pady=2)
        self.stats_energy_label = ttk.Label(stats_fr, text="0.0 Wh", foreground='purple')
        self.stats_energy_label.grid(row=6, column=1, sticky='e', pady=2)

        ttk.Label(stats_fr, text="Total Capacity:", font=('Arial', 9, 'bold')).grid(row=7, column=0, sticky='w', pady=2)
        self.stats_capacity_label = ttk.Label(stats_fr, text="0.0 Ah", foreground='purple')
        self.stats_capacity_label. grid(row=7, column=1, sticky='e', pady=2)

        # === TAB 3:  Differential Capacity (dQ/dV) Analysis ===
        dqdv_tab = ttk.Frame(main_notebook)
        main_notebook.add(dqdv_tab, text="Differential Capacity (dQ/dV)")
        
        # dQ/dV controls at top
        dqdv_control_frame = ttk.Frame(dqdv_tab)
        dqdv_control_frame.pack(side=tk.TOP, fill=tk.X, padx=8, pady=6)
        
        ttk.Label(dqdv_control_frame, text="Differential Capacity Analysis (dQ/dV vs Voltage)", 
                  font=('Arial', 14, 'bold')).pack(side=tk.LEFT, padx=10)
        
        ttk.Button(dqdv_control_frame, text="Clear dQ/dV Plot", 
                   command=self._clear_dqdv_plot).pack(side=tk.RIGHT, padx=5)
        ttk.Button(dqdv_control_frame, text="Export dQ/dV Image", 
                   command=self._save_dqdv_screenshot).pack(side=tk.RIGHT, padx=5)
        ttk.Button(dqdv_control_frame, text="Export dQ/dV Data (CSV)", 
                   command=self._export_dqdv_data).pack(side=tk.RIGHT, padx=5)
        
        # Info label
        info_frame = ttk.Frame(dqdv_tab)
        info_frame.pack(side=tk.TOP, fill=tk.X, padx=8, pady=(0, 6))
        info_label = ttk.Label(info_frame, 
                               text="dQ/dV (Differential Capacity) reveals battery phase transitions, degradation, and internal chemistry behavior.",
                               font=('Arial', 9, 'italic'), foreground='gray')
        info_label. pack(side=tk.LEFT, padx=10)
        
        # Main dQ/dV plot area
        dqdv_plot_frame = ttk.Frame(dqdv_tab)
        dqdv_plot_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=8, pady=6)

        # Create dQ/dV plot (dQ/dV vs Voltage)
        self.fig_dqdv = Figure(figsize=(12, 7), dpi=100)
        self.ax_dqdv = self.fig_dqdv.add_subplot(111)
        
        # **NEW: Disable offset/scientific notation**
        self.ax_dqdv.ticklabel_format(style='plain', axis='both', useOffset=False)
        
        # Labels
        self. ax_dqdv.set_xlabel("Voltage (V)", fontsize=12, fontweight='bold')
        self.ax_dqdv. set_ylabel("dQ/dV (Ah/V)", fontsize=12, fontweight='bold', 
                                color=self.dqdv_color_var.get())
        self.ax_dqdv.set_title("Differential Capacity Analysis", fontsize=13, fontweight='bold', pad=15)
        
        # Grid
        self.ax_dqdv.grid(True, alpha=0.4, linestyle='--', linewidth=0.8)
        self.ax_dqdv.axhline(y=0, color='black', linestyle='-', linewidth=0.8, alpha=0.3)
        
        # Create line object with safe defaults
        (self.line_dqdv,) = self.ax_dqdv.plot([], [], 
                                               color=self.dqdv_color_var.get(), 
                                               linestyle=LINE_STYLES.get(self.dqdv_style_var.get(), "-"), 
                                               linewidth=self.dqdv_width_var.get(),
                                               marker=MARKER_STYLES.get(self.dqdv_marker_var.get(), "o"),
                                               markersize=5,
                                               label="dQ/dV")
        
        self.ax_dqdv.legend(loc='upper right', fontsize=10, framealpha=0.9)
        self.fig_dqdv.tight_layout()

        self.canvas_dqdv = FigureCanvasTkAgg(self.fig_dqdv, master=dqdv_plot_frame)
        self.canvas_dqdv.draw()
        self.canvas_dqdv.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        toolbar_dqdv = NavigationToolbar2Tk(self.canvas_dqdv, dqdv_plot_frame)
        toolbar_dqdv.update()
        toolbar_dqdv.pack(side=tk.BOTTOM, fill=tk. X)

        # Right panel for dQ/dV plot styles and info
        dqdv_right_panel = ttk.Frame(dqdv_tab)
        dqdv_right_panel.pack(side=tk.RIGHT, fill=tk.Y, padx=8, pady=6)

        # # Style controls
        # dqdv_style_fr = ttk. Labelframe(dqdv_right_panel, text="Plot Appearance", padding=(10, 10))
        # dqdv_style_fr.pack(side=tk.TOP, fill=tk.X)

        # ttk.Label(dqdv_style_fr, text="Line Color:", font=('Arial', 9, 'bold')).pack(anchor='w', pady=(5, 2))
        # self.dqdv_color_display = tk.Canvas(dqdv_style_fr, width=100, height=30, 
        #                                      bg=self.dqdv_color_var.get(), 
        #                                      relief=tk.RIDGE, borderwidth=2)
        # self.dqdv_color_display.pack(pady=3)
        # ttk.Button(dqdv_style_fr, text="Choose Color", 
        #            command=lambda: self._choose_dqdv_color()).pack(pady=3)
        
        # ttk.Label(dqdv_style_fr, text="Line Width:", font=('Arial', 9, 'bold')).pack(anchor='w', pady=(8, 2))
        # ttk.Scale(dqdv_style_fr, from_=0.5, to=5.0, variable=self.dqdv_width_var, 
        #           orient=tk.HORIZONTAL).pack(fill=tk.X, padx=5, pady=3)
        
        # ttk.Label(dqdv_style_fr, text="Line Style:", font=('Arial', 9, 'bold')).pack(anchor='w', pady=(8, 2))
        # dqdv_style_combo = ttk.Combobox(dqdv_style_fr, textvariable=self. dqdv_style_var, 
        #                                 values=list(LINE_STYLES.keys()), state='readonly', width=15)
        # dqdv_style_combo.pack(pady=3)
        
        # ttk. Label(dqdv_style_fr, text="Marker Style:", font=('Arial', 9, 'bold')).pack(anchor='w', pady=(8, 2))
        # dqdv_marker_combo = ttk. Combobox(dqdv_style_fr, textvariable=self.dqdv_marker_var, 
        #                                  values=list(MARKER_STYLES.keys()), state='readonly', width=15)
        # dqdv_marker_combo.pack(pady=3)
        
        # ttk.Button(dqdv_style_fr, text="Apply Styles", 
        #            command=self._apply_dqdv_styles).pack(pady=15, fill=tk.X, padx=10)

        # dQ/dV Statistics
        dqdv_stats_fr = ttk.Labelframe(dqdv_right_panel, text="dQ/dV Statistics", padding=(10, 10))
        dqdv_stats_fr.pack(side=tk.TOP, fill=tk.X, pady=(10, 0))

        ttk.Label(dqdv_stats_fr, text="Data Points:", font=('Arial', 9, 'bold')).grid(row=0, column=0, sticky='w', pady=2)
        self.dqdv_points_label = ttk.Label(dqdv_stats_fr, text="0")
        self.dqdv_points_label.grid(row=0, column=1, sticky='e', pady=2)

        ttk.Label(dqdv_stats_fr, text="Current dQ/dV:", font=('Arial', 9, 'bold')).grid(row=1, column=0, sticky='w', pady=2)
        self.dqdv_current_label = ttk.Label(dqdv_stats_fr, text="0. 0 Ah/V")
        self.dqdv_current_label.grid(row=1, column=1, sticky='e', pady=2)

        ttk.Label(dqdv_stats_fr, text="At Voltage:", font=('Arial', 9, 'bold')).grid(row=2, column=0, sticky='w', pady=2)
        self.dqdv_voltage_label = ttk.Label(dqdv_stats_fr, text="0.0 V")
        self.dqdv_voltage_label.grid(row=2, column=1, sticky='e', pady=2)

        ttk.Label(dqdv_stats_fr, text="Max dQ/dV:", font=('Arial', 9, 'bold')).grid(row=3, column=0, sticky='w', pady=2)
        self.dqdv_max_label = ttk.Label(dqdv_stats_fr, text="0.0 Ah/V", foreground='darkred')
        self.dqdv_max_label.grid(row=3, column=1, sticky='e', pady=2)

        ttk.Label(dqdv_stats_fr, text="Min dQ/dV:", font=('Arial', 9, 'bold')).grid(row=4, column=0, sticky='w', pady=2)
        self.dqdv_min_label = ttk.Label(dqdv_stats_fr, text="0.0 Ah/V", foreground='darkblue')
        self.dqdv_min_label.grid(row=4, column=1, sticky='e', pady=2)

        # Info panel
        dqdv_info_fr = ttk.Labelframe(dqdv_right_panel, text="About dQ/dV", padding=(8, 8))
        dqdv_info_fr.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(10, 0))
        
        info_text = tk.Text(dqdv_info_fr, wrap=tk. WORD, width=25, height=14, font=('Arial', 9))  # Changed:  width=25, font size=8
        info_text. pack(fill=tk.BOTH, expand=True)
        info_text.insert(tk.END, 
            "Differential Capacity:\n\n"
            "• Rate of capacity change vs voltage\n\n"
            "• Peaks show phase transitions\n\n"
            "• Calculated as:\n"
            "  dQ/dV = ΔQ / ΔV\n"
            "  Q = capacity (Ah)\n"
            "  V = voltage (V)\n\n"
            "• Best for constant current tests"
        )
        info_text.config(state=tk. DISABLED)

    def _clear_log_display(self):
        """Clear the log display (doesn't affect saved log file)"""
        self.log_text.delete('1.0', tk.END)
        self._log("Log display cleared (saved logs unaffected)")

    def _clear_plots(self):
        """Clear all plot data"""
        self.plot_times.clear()
        self.plot_total_power.clear()
        self.plot_total_current.clear()
        self.plot_avg_voltage.clear()
        self.plot_cumulative_energy.clear()
        self.plot_cumulative_capacity.clear()
        self.plot_dqdv_voltage.clear()  # NEW
        self.plot_dqdv_values.clear()  # NEW
        self._update_plot_lines(redraw=True)
        self._update_dqdv_plot(redraw=True)  # NEW
        self._log("All plots cleared")

    def _export_plot_data(self):
        """Export combined plot data to """
        if not self.plot_times:
            messagebox.showwarning("No Data", "No plot data to export")
            return
        
        path = filedialog.asksaveasfilename(
            title="Export Plot Data",
            defaultextension=".csv",
            filetypes=[("CSV files", "*. csv"), ("All files", "*.*")],
            initialfile=f"combined_plot_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        
        if not path:
            return
        
        try:
            with open(path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['Time_s', 'Total_Power_W', 'Total_Current_A', 'Avg_Voltage_V',
                               'Cumulative_Energy_Wh', 'Cumulative_Capacity_Ah'])  # NEW columns
                
                for i in range(len(self.plot_times)):
                    writer.writerow([
                        self.plot_times[i],
                        self.plot_total_power[i],
                        self.plot_total_current[i],
                        self.plot_avg_voltage[i],
                        self.plot_cumulative_energy[i],  # NEW
                        self.plot_cumulative_capacity[i]  # NEW
                    ])
            
            messagebox.showinfo("Success", f"Plot data exported to:\n{path}")
            self._log(f"Combined plot data exported:  {path}")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))
            self._log(f"[ERROR] Export failed: {e}")
            
    def _debug_plot_data(self):
        """Debug method to check plot data"""
        self._log("=== PLOT DEBUG INFO ===")
        self._log(f"Data points: {len(self.plot_times)}")
        
        if self.plot_times:
            self._log(f"Time range: {min(self.plot_times)} to {max(self.plot_times)} seconds")
        
        if self.plot_total_power:
            power_vals = [p for p in self. plot_total_power if p and not (p != p)]
            if power_vals:
                self._log(f"Power:  min={min(power_vals):.2f}W, max={max(power_vals):.2f}W, count={len(power_vals)}")
            else:
                self._log("Power:  NO VALID VALUES")
        else:
            self._log("Power:  EMPTY ARRAY")
            
        if self.plot_total_current:
            current_vals = [c for c in self. plot_total_current if c and not (c != c)]
            if current_vals:
                self._log(f"Current: min={min(current_vals):.2f}A, max={max(current_vals):.2f}A, count={len(current_vals)}")
        
        if self.plot_avg_voltage:
            voltage_vals = [v for v in self.plot_avg_voltage if v and not (v != v)]
            if voltage_vals:
                self._log(f"Voltage: min={min(voltage_vals):.2f}V, max={max(voltage_vals):.2f}V, count={len(voltage_vals)}")
        
        if self.plot_cumulative_energy:
            energy_vals = [e for e in self.plot_cumulative_energy if e is not None]
            if energy_vals:
                self._log(f"Energy: min={min(energy_vals):.4f}Wh, max={max(energy_vals):.4f}Wh, count={len(energy_vals)}")
        
        # Check line visibility
        self._log(f"Power line visible: {self.line_power.get_visible()}")
        self._log(f"Power line color: {self.line_power. get_color()}")
        self._log(f"Power line width: {self.line_power. get_linewidth()}")
        self._log(f"Power line zorder: {self.line_power.get_zorder()}")
        
        # Check axis limits
        self._log(f"Power axis Y limits: {self.ax_power.get_ylim()}")
        self._log(f"Power axis X limits: {self.ax_power.get_xlim()}")
        
        self._log("=== END DEBUG ===")
        
    def _toggle_imbalance_controls(self):
        """Enable or disable imbalance control entries based on checkbox"""
        try:
            if self.imbalance_check_enabled_var.get():
                # Enable all imbalance entries
                self.imbalance_limit_entry.config(state='normal')
                self.imbalance_transition_entry.config(state='normal')
                self. transition_duration_entry.config(state='normal')
                self.voltage_diff_entry.config(state='normal')
                self._log("Imbalance monitoring ENABLED")
            else:
                # Disable all imbalance entries
                self.imbalance_limit_entry.config(state='disabled')
                self.imbalance_transition_entry.config(state='disabled')
                self. transition_duration_entry.config(state='disabled')
                self. voltage_diff_entry.config(state='disabled')
                self._log("⚠ WARNING: Imbalance monitoring DISABLED - system will not shut down on imbalance!")
        except Exception as e:
            print(f"[ERROR] Toggling imbalance controls: {e}")

    # Profile row helpers
    def _add_row(self):
        """Add a new row"""
        # First, save current values from existing widgets
        self._sync_profile_from_widgets()
        
        # Add new row
        new_row = ProfileRow(total_power=100.0, duration_value=1, duration_unit="hours", cutoff=False)
        self.settings.profile.append(new_row)
        self._append_row_widget(new_row)
        self._log("Added new profile row")

    def _delete_row(self, index):
        """Delete a row"""
        if 0 <= index < len(self.row_widgets):
            # Save current values first
            self._sync_profile_from_widgets()
            
            # Remove from both GUI and data
            widget = self.row_widgets. pop(index)
            widget. destroy()
            self.settings.profile.pop(index)
            
            # Rebuild to fix indices
            self._rebuild_row_widgets()
            self._log(f"Deleted row {index + 1}")

    def _move_up(self, index):
        """Move row up"""
        if index <= 0 or index >= len(self.row_widgets):
            return
        
        # Save current values first
        self._sync_profile_from_widgets()
        
        # Swap in profile
        self.settings.profile[index - 1], self.settings.profile[index] = \
            self.settings.profile[index], self.settings.profile[index - 1]
        
        # Rebuild widgets
        self._rebuild_row_widgets()
        self._log(f"Moved row {index + 1} up")

    def _move_down(self, index):
        """Move row down"""
        if index < 0 or index >= len(self.row_widgets) - 1:
            return
        
        # Save current values first
        self._sync_profile_from_widgets()
        
        # Swap in profile
        self.settings.profile[index], self.settings.profile[index + 1] = \
            self. settings.profile[index + 1], self.settings.profile[index]
        
        # Rebuild widgets
        self._rebuild_row_widgets()
        self._log(f"Moved row {index + 1} down")

    def _clear_all_rows(self):
        """Clear all rows"""
        if not messagebox.askyesno("Clear all", "Remove all profile rows?"):
            return
        for w in self.row_widgets:
            w.destroy()
        self.row_widgets.clear()
        self.settings.profile.clear()
        self._log("Cleared all profile rows.")

    def _rebuild_row_widgets(self):
        """Rebuild all row widgets from profile data"""
        # Destroy all existing widgets
        for w in self. row_widgets:
            w. destroy()
        self.row_widgets = []
        
        # Recreate from profile data
        for model in self.settings.profile:
            self._append_row_widget(model)

    def _sync_profile_from_widgets(self):
        """Sync profile data from current widget values"""
        for i, widget in enumerate(self.row_widgets):
            if i < len(self.settings. profile):
                self.settings.profile[i] = widget.get_value()

    def _refresh_row_indices(self):
        """Refresh row indices after changes"""
        for i, w in enumerate(self.row_widgets):
            w.update_index(i)
            w.grid(row_idx=i + 1)

    def _append_row_widget(self, row_model:  ProfileRow):
        """Append a row widget"""
        idx = len(self.row_widgets)
        widget = ProfileRowWidget(
            self.rows_container, 
            idx, 
            row_model, 
            on_delete=self._delete_row, 
            on_move_up=self._move_up, 
            on_move_down=self._move_down
        )
        widget.grid(row_idx=idx + 1)
        self.row_widgets.append(widget)
        self._refresh_row_indices()

    # CSV/config helpers
    def _select_csv_folder(self):
        """Select CSV folder for logging"""
        path = filedialog.askdirectory(title="Select folder to save CSV files")
        if not path:
            return
        self.csv_folder = path
        self.csv_folder_label. config(text=path, foreground="green", font=('Arial', 9))
        self._log(f"✓ CSV folder set:   {path}")

    def _select_config_save_folder(self):
        path = filedialog.askdirectory(title="Select folder to save configuration")
        if not path:
            return
        self.config_save_folder = path
        try:
            self.config_save_folder_label.config(text=path)
        except Exception: 
            pass
        self._log(f"Config save folder set: {path}")

    def _select_config_file(self):
        path = filedialog.askopenfilename(title="Select configuration file", filetypes=[("Text files", "*.txt"), ("JSON files", "*.json"), ("All files", "*.*")])
        if not path:
            return
        self.config_load_path_var.set(path)
        self._log(f"Selected config:  {path}")

    def _save_configuration(self):
        """Save configuration to file"""
        filename = self.config_filename_var.get().strip()
        folder = getattr(self, "config_save_folder", None)
        if not filename:
            messagebox.showerror("Filename required", "Enter a filename.")
            return
        if not folder:
            messagebox.showerror("Folder required", "Choose a folder.")
            return
        if not (filename.lower().endswith(".txt") or filename.lower().endswith(".json")):
            filename = filename + ".txt"
        path = os.path.join(folder, filename)
        
        try:
            self._apply_settings()
        except Exception as e:
            self._log(f"[ERROR] Cannot save:   {e}")
            return

        cfg = {
            "master_port": self.settings.master_port,
            "slave_port": self.settings. slave_port,
            "baud_rate": self.settings. baud_rate,
            "dry_run": self.settings.dry_run,
            "cutoff_voltage": self.settings.cutoff_voltage,
            "cutoff_safety_enabled": self. settings.cutoff_safety_enabled,
            "imbalance_check_enabled": self.settings.imbalance_check_enabled,  # NEW
            "imbalance_limit": self.settings.imbalance_limit,
            "imbalance_transition":  self.settings.imbalance_transition,
            "transition_duration": self.settings.transition_duration,
            "voltage_diff_limit": self.settings. voltage_diff_limit,
            "csv_folder": self.csv_folder,
            "rotate_hours": float(self.rotate_hours_var.get()),
            "record_interval_s": int(self.record_interval_var.get()),
            "email_enabled": self.settings.email_enabled,
            "email_sender": self.settings.email_sender,
            "email_receiver": self.settings.email_receiver,
            "email_subject":  self.settings.email_subject,
            "profile":  [
                {
                    "total_power": float(r.total_power),
                    "duration_value": float(r.duration_value),
                    "duration_unit": r.duration_unit,
                    "cutoff": bool(r.cutoff)
                } for r in self.settings.profile
            ],
            "plot_styles": {
                "power_color": self.power_color_var.get(),
                "power_width": self.power_width_var.get(),
                "power_style": self.power_style_var.get(),
                "current_color": self.current_color_var.get(),
                "current_width": self.current_width_var.get(),
                "current_style": self. current_style_var.get(),
                "voltage_color": self.voltage_color_var.get(),
                "voltage_width":  self.voltage_width_var. get(),
                "voltage_style": self.voltage_style_var.get(),
                "energy_color": self.energy_color_var.get(),
                "energy_width": self.energy_width_var.get(),
                "energy_style": self.energy_style_var.get(),
                "dqdv_color": self.dqdv_color_var.get(),
                "dqdv_width": self.dqdv_width_var.get(),
                "dqdv_style": self. dqdv_style_var. get(),
                "dqdv_marker": self.dqdv_marker_var.get()
            },
            "config_note": self._get_config_note_text(),
            "config_filename": self.config_filename_var.get()
        }
        
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2)
            messagebox.showinfo("Saved", f"Configuration saved to {path}")
            self._log(f"Configuration saved:   {path}")
            self._update_config_note()
        except Exception as e: 
            messagebox.showerror("Save error", f"Failed to save:   {e}")
            self._log(f"[ERROR] Save failed:  {e}")

    def _load_configuration(self):
        """Load configuration from file"""
        path = self.config_load_path_var.get().strip()
        if not path:
            messagebox.showerror("No file", "Choose a configuration file.")
            return
        if not os.path.isfile(path):
            messagebox.showerror("Not found", f"File not found:  {path}")
            return
        
        try:
            with open(path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        except Exception as e:
            messagebox.showerror("Load error", f"Failed to read:   {e}")
            return

        try:
            self.master_port_entry.delete(0, tk.END)
            self.master_port_entry.insert(0, str(cfg. get("master_port", DEFAULT_MASTER_PORT)))
            
            self.slave_port_entry.delete(0, tk.END)
            self.slave_port_entry.insert(0, str(cfg.get("slave_port", DEFAULT_SLAVE_PORT)))
            
            self.baud_entry.delete(0, tk.END)
            self.baud_entry.insert(0, str(cfg.get("baud_rate", DEFAULT_BAUD_RATE)))
            
            self. dry_var.set(bool(cfg.get("dry_run", False)))
            
            self.cutoff_entry.delete(0, tk.END)
            self.cutoff_entry.insert(0, str(cfg.get("cutoff_voltage", DEFAULT_CUTOFF_VOLTAGE)))
            
            self.cutoff_safety_var.set(bool(cfg. get("cutoff_safety_enabled", False)))
            self.imbalance_check_enabled_var.set(bool(cfg.get("imbalance_check_enabled", True)))  # NEW
            self.imbalance_limit_var.set(float(cfg.get("imbalance_limit", DEFAULT_IMBALANCE_LIMIT)) * 100)
            self.imbalance_transition_var.set(float(cfg.get("imbalance_transition", DEFAULT_IMBALANCE_TRANSITION)) * 100)
            self.transition_duration_var.set(int(cfg.get("transition_duration", DEFAULT_TRANSITION_DURATION)))
            self.voltage_diff_limit_var.set(float(cfg.get("voltage_diff_limit", DEFAULT_VOLTAGE_DIFF_LIMIT)) * 100)

            csv_folder = cfg.get("csv_folder", None)
            if csv_folder: 
                self.csv_folder = csv_folder
                self.csv_folder_label.config(text=csv_folder)
            else:
                self.csv_folder = None
                self.csv_folder_label.config(text="(not set)")

            self.rotate_hours_var.set(float(cfg.get("rotate_hours", DEFAULT_ROTATE_HOURS)))
            self.record_interval_var.set(int(cfg.get("record_interval_s", DEFAULT_RECORD_INTERVAL_S)))

            self.email_enabled_var.set(bool(cfg.get("email_enabled", False)))
            self.email_sender_var.set(cfg.get("email_sender", ""))
            self.email_receiver_var.set(cfg.get("email_receiver", ""))

            self._toggle_imbalance_controls()

            raw_profile = cfg.get("profile", [])
            new_profile: List[ProfileRow] = []
            for entry in raw_profile:
                power = float(entry.get("total_power", 0.0))
                cutoff = bool(entry.get("cutoff", False))
                dur_unit = entry.get("duration_unit", "seconds")
                dur_val = float(entry.get("duration_value", 0.0))
                if cutoff:
                    new_profile.append(ProfileRow(total_power=power, duration_value=0, duration_unit="seconds", cutoff=True))
                else:
                    if dur_unit not in [u for u, _ in UNITS]:
                        dur_unit = "seconds"
                    new_profile.append(ProfileRow(total_power=power, duration_value=dur_val, duration_unit=dur_unit, cutoff=False))

            if not new_profile:
                raise ValueError("Profile is empty.")

            self.settings.profile = new_profile
            self._rebuild_row_widgets()

            # Load plot styles
            styles = cfg.get("plot_styles", {})
            if styles:
                self.power_color_var.set(styles. get("power_color", "blue"))
                self.power_width_var.set(styles.get("power_width", 2.0))
                self.power_style_var.set(styles. get("power_style", "solid"))
                
                self.current_color_var.set(styles.get("current_color", "red"))
                self.current_width_var. set(styles.get("current_width", 2.0))
                self.current_style_var.set(styles.get("current_style", "solid"))
                
                self. voltage_color_var.set(styles.get("voltage_color", "green"))
                self.voltage_width_var.set(styles.get("voltage_width", 2.0))
                self. voltage_style_var.set(styles.get("voltage_style", "solid"))
                
                self.energy_color_var.set(styles.get("energy_color", "purple"))
                self.energy_width_var.set(styles.get("energy_width", 2.0))
                self.energy_style_var.set(styles. get("energy_style", "solid"))
                
                self. dqdv_color_var.set(styles.get("dqdv_color", "darkred"))
                self.dqdv_width_var.set(styles.get("dqdv_width", 2.0))
                self.dqdv_style_var.set(styles.get("dqdv_style", "solid"))
                self.dqdv_marker_var.set(styles.get("dqdv_marker", "circle"))

            note = cfg.get("config_note", "")
            if note: 
                self._set_config_note_text(note)
            cfg_fname = cfg.get("config_filename", "")
            if cfg_fname:
                self.config_filename_var.set(cfg_fname)

            self._apply_line_styles()
            if hasattr(self, '_apply_dqdv_styles'):
                self._apply_dqdv_styles()
                
            messagebox.showinfo("Loaded", f"Configuration loaded from {path}")
            self._log(f"Configuration loaded:  {path}")
        except Exception as e:
            messagebox.showerror("Apply error", f"Failed to apply:   {e}")
            self._log(f"[ERROR] Apply failed: {e}")
            import traceback
            self._log(traceback.format_exc())

    # Logging
    def _now_central(self) -> datetime:
        if CHI_TZ is not None:
            return datetime.now(CHI_TZ)
        else:
            return datetime.now()

    def _format_ts(self, dt: datetime) -> str:
        try:
            return dt.isoformat()
        except Exception:
            return dt.strftime("%Y-%m-%d %H:%M:%S")

    def _log(self, msg: str):
        """Log message to UI (App class version)"""
        try:
            timestamp = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
            full_msg = f"{timestamp} {msg}"
            
            # Add to UI text widget
            if hasattr(self, 'log_text'):
                self.log_text.insert(tk.END, full_msg + "\n")
                self.log_text. see(tk.END)
            
            # Print to console as backup
            print(full_msg)
        except Exception as e:
            print(f"[ERROR] Logging failed: {e}")

    def _periodic_log_poll(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self. log_widget.configure(state=tk.NORMAL)
                self.log_widget.insert(tk.END, msg + "\n")
                self.log_widget.see(tk.END)
                self.log_widget.configure(state=tk.DISABLED)
                
                self.status_top_lines.append(msg)
                self.status_top_widget.configure(state=tk.NORMAL)
                self.status_top_widget. delete("1.0", tk.END)
                lines = list(self.status_top_lines)
                self.status_top_widget.insert(tk.END, "\n".join(lines))
                self.status_top_widget. see(tk.END)
                self.status_top_widget.configure(state=tk.DISABLED)
        except queue.Empty:
            pass
        self.after(200, self._periodic_log_poll)

    def _clear_log(self):
        self.log_widget.configure(state=tk. NORMAL)
        self.log_widget.delete("1.0", tk.END)
        self.log_widget.configure(state=tk.DISABLED)
        self.status_top_lines.clear()
        self.status_top_widget.configure(state=tk.NORMAL)
        self.status_top_widget.delete("1.0", tk.END)
        self.status_top_widget.configure(state=tk. DISABLED)

    def _export_log(self):
        path = filedialog.asksaveasfilename(title="Export log", defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if not path: 
            return
        text = self.log_widget.get("1.0", tk.END)
        try:
            with open(path, "w", encoding="utf-8") as f:
                f. write(text)
            messagebox.showinfo("Exported", f"Log exported to {path}")
        except Exception as e:
            messagebox.showerror("Export error", str(e))

    # Apply settings
    def _apply_settings(self):
        """Apply settings from GUI"""
        try:
            self.settings.master_port = self.master_port_entry.get().strip() or DEFAULT_MASTER_PORT
            self.settings.slave_port = self.slave_port_entry.get().strip() or DEFAULT_SLAVE_PORT
            self.settings. baud_rate = int(self.baud_entry.get().strip())
            self.settings.dry_run = bool(self.dry_var.get())
            self.settings.cutoff_voltage = float(self.cutoff_entry.get().strip())
            self.settings.cutoff_safety_enabled = bool(self.cutoff_safety_var.get())
            self.settings.imbalance_check_enabled = bool(self.imbalance_check_enabled_var.get())
            self.settings.imbalance_limit = float(self.imbalance_limit_var.get()) / 100.0
            self. settings.imbalance_transition = float(self.imbalance_transition_var.get()) / 100.0
            self.settings.transition_duration = int(self.transition_duration_var.get())
            self.settings.voltage_diff_limit = float(self.voltage_diff_limit_var.get()) / 100.0
            
            self.settings.email_enabled = bool(self.email_enabled_var.get())
            self.settings.email_sender = self.email_sender_var.get().strip()
            self.settings. email_receiver = self.email_receiver_var.get().strip()
            self.settings.email_password = self.email_password_var.get().strip()

            # Sync profile from widgets
            self._sync_profile_from_widgets()
            
            # Validate profile
            if not self. settings.profile: 
                raise ValueError("Profile must contain at least one row.")
            
            # Log warning if imbalance check is disabled
            if not self.settings. imbalance_check_enabled: 
                self._log("⚠⚠⚠ WARNING: Imbalance monitoring is DISABLED!  ⚠⚠⚠")
            
            self._log("Settings applied.")
            self._update_config_note()
            return True
        except Exception as e:
            messagebox.showerror("Error applying settings", str(e))
            self._log(f"[ERROR] Applying settings: {e}")
            raise

    # Start/Stop
    def _start(self):
        """Start the test with automatic pretest conditioning"""
        if self. worker and self.worker.is_alive():
            messagebox.showwarning("Already running", "Test is in progress.")
            return
        
        # Check if CSV folder is selected
        if not self.csv_folder or not os.path.isdir(self.csv_folder):
            response = messagebox.askyesno(
                "CSV Folder Not Set",
                "CSV folder is not selected or invalid.\n\n"
                "Data logging will NOT be saved to CSV files.\n\n"
                "Do you want to select a CSV folder now?\n\n"
                "Click 'Yes' to select folder\n"
                "Click 'No' to continue without CSV logging (NOT RECOMMENDED)",
                icon='warning'
            )
            if response:
                self._select_csv_folder()
                if not self.csv_folder or not os.path.isdir(self.csv_folder):
                    messagebox.showerror(
                        "Cannot Start",
                        "CSV folder must be selected to ensure data is recorded.\n\n"
                        "Test cannot start without valid CSV folder."
                    )
                    self._log("[ERROR] Test start aborted: CSV folder not selected")
                    return
            else:
                messagebox.showwarning(
                    "Test Cancelled",
                    "Test cancelled. Please select a CSV folder to ensure data recording."
                )
                self._log("[WARNING] Test start cancelled:  CSV folder not selected")
                return
        
        # Apply settings
        try:
            self._apply_settings()
        except Exception: 
            return
        
        # Confirm start with user - mention conditioning
        response = messagebox.askyesno(
            "Start Test",
            "Ready to start test.\n\n"
            "The system will automatically:\n"
            "1. Condition instruments (REMOTE, CW mode, 0W, etc.)\n"
            "2. Execute the power profile\n"
            "3. Log all data to CSV files\n\n"
            "Continue? ",
            icon='question'
        )
        
        if not response:
            self._log("Test start cancelled by user")
            return
        
        # Disable start button and update status
        self.start_btn. config(state=tk.DISABLED)
        self.stop_btn.config(state=tk. DISABLED)  # Disable until conditioning completes
        self. status_label.config(text="Conditioning instruments.. .", foreground="orange")
        self._log("═══ Starting Test Sequence ═══")
        
        # Run conditioning in a separate thread, then start test
        import threading
        conditioning_thread = threading.Thread(target=self._condition_and_start, daemon=True)
        conditioning_thread.start()
    
    def _condition_and_start(self):
        """Run conditioning, then start the test (runs in separate thread)"""
        master = None
        slave = None
        conditioning_success = False
        
        try: 
            # === CONDITIONING PHASE ===
            self._log("─── Phase 1: Instrument Conditioning ───")
            
            # Create instrument controllers
            self._log("Opening connections to instruments...")
            master = InstrumentController(
                self.settings. master_port, 
                self.settings.baud_rate, 
                self._log, 
                "MASTER"
            )
            slave = InstrumentController(
                self.settings.slave_port, 
                self.settings.baud_rate, 
                self._log, 
                "SLAVE"
            )
            
            # Open connections
            master.open(self.settings.timeout)
            slave.open(self.settings.timeout)
            self._log("✓ Connections established")
            
            # Step 1: Set to REMOTE mode
            self._log("Step 1: Setting REMOTE mode...")
            master.set_remote()
            slave.set_remote()
            time.sleep(0.5)
            self._log("  ✓ Both instruments in REMOTE mode")
            
            # Step 2: Set to Constant Power (CW) mode
            self._log("Step 2: Setting Constant Power (CW) mode...")
            master.set_mode_cw()
            slave.set_mode_cw()
            time.sleep(0.5)
            self._log("  ✓ Both instruments in CW mode")
            
            # Step 3: Set power to 0W
            self._log("Step 3: Setting power to 0W...")
            master.set_power(0.0)
            slave.set_power(0.0)
            time.sleep(0.5)
            self._log("  ✓ Power set to 0W")
            
            # Step 4: Turn INPUT ON
            self._log("Step 4: Turning INPUT ON...")
            master.input_on()
            slave.input_on()
            time.sleep(0.5)
            self._log("  ✓ INPUT ON for both instruments")
            
            # Step 5: Enable PARALLEL mode
            self._log("Step 5: Enabling PARALLEL mode...")
            master.parallel_on()
            slave.parallel_on()
            time.sleep(0.5)
            self._log("  ✓ PARALLEL mode enabled")
            
            # Step 6: Turn LOAD ON
            self._log("Step 6: Turning LOAD ON...")
            master.load_on()
            slave.load_on()
            time.sleep(0.5)
            self._log("  ✓ LOAD ON for both instruments")
            
            # Verify configuration
            self._log("Verifying configuration...")
            time.sleep(1.0)
            
            # Read back power settings
            master_power = master.read_power()
            slave_power = slave.read_power()
            self._log(f"  Master Power: {master_power}W")
            self._log(f"  Slave Power: {slave_power}W")
            
            conditioning_success = True
            self._log("✓ Conditioning COMPLETE - instruments ready")
            
        except Exception as e:
            self._log(f"[ERROR] Conditioning failed:  {e}")
            import traceback
            self._log(traceback.format_exc())
            
            # Try to return to safe state
            self._log("Attempting to return to safe state...")
            try:
                if master:
                    master.set_power(0.0)
                    master.load_off()
                    master. set_local()
                if slave:
                    slave.set_power(0.0)
                    slave.load_off()
                    slave.set_local()
                self._log("  ✓ Instruments returned to safe state")
            except Exception:
                pass
        
        finally:
            # Close conditioning connections
            try:
                if master:
                    master.close()
                if slave:
                    slave.close()
                self._log("✓ Conditioning connections closed")
            except Exception: 
                pass
        
        # Update UI and start test if conditioning succeeded
        if conditioning_success: 
            self. after(0, self._start_test_after_conditioning)
        else:
            self.after(0, self._conditioning_failed)
    
    def _start_test_after_conditioning(self):
        """Start the actual test after successful conditioning (called on main thread)"""
        try:
            self._log("─── Phase 2: Starting Test Execution ───")
            
            csv_folder = self.csv_folder
            rotate_hours = float(self.rotate_hours_var.get())
            record_interval_s = int(self.record_interval_var.get())
            
            # Create and start worker thread
            self.worker = MasterSlaveWorkerThread(
                self.settings, 
                csv_folder, 
                rotate_hours, 
                record_interval_s, 
                self. log_queue, 
                plot_queue=self.plot_queue
            )
            self.worker.start()
            
            # Update UI
            self. start_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk. NORMAL)
            self.status_label.config(text="Running Test", foreground="green")
            self._log("═══ Test Execution Started ═══")
            self._log(f"✓ CSV logging:  {csv_folder}")
            
            # Reset plot arrays
            self.plot_times.clear()
            self.plot_total_power.clear()
            self.plot_total_current.clear()
            self.plot_avg_voltage.clear()
            self.plot_cumulative_energy.clear()
            self.plot_cumulative_capacity.clear()
            self.plot_dqdv_voltage.clear()
            self.plot_dqdv_values.clear()
            self._update_plot_lines(redraw=True)
            if hasattr(self, '_update_dqdv_plot'):
                self._update_dqdv_plot(redraw=True)
                
        except Exception as e:
            self._log(f"[ERROR] Failed to start test: {e}")
            import traceback
            self._log(traceback.format_exc())
            self._conditioning_failed()
    
    def _conditioning_failed(self):
        """Called when conditioning fails (called on main thread)"""
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.status_label.config(text="Conditioning Failed", foreground="red")
        
        messagebox.showerror(
            "Test Start Failed",
            "Instrument conditioning failed.\n\n"
            "Please check:\n"
            "• Instruments are powered on\n"
            "• COM ports are correct\n"
            "• Cables are connected\n"
            "• No other software is using the ports\n\n"
            "Check the Test Logs for details."
        )

    def _stop(self):
        """Stop the test and safely shut down equipment"""
        if not self.worker:  
            messagebox.showinfo("Not running", "No test is running.")
            return
        
        # Check if worker is actually alive
        try:
            worker_alive = self.worker.is_alive()
        except Exception as e:
            self._log(f"[ERROR] Checking worker status: {e}")
            worker_alive = False
        
        if not worker_alive:
            messagebox.showinfo("Not running", "No test is currently running.")
            self. start_btn.config(state=tk. NORMAL)
            self.stop_btn.config(state=tk. DISABLED)
            self.status_label.config(text="Idle", foreground="blue")
            return
        
        # Request stop
        self. worker.request_stop()
        self._log("═══ STOP REQUESTED ═══")
        self._log("⏸ Initiating safe shutdown...")
        self._log("  • Ramping power to 0W")
        self._log("  • Turning loads OFF")
        self._log("  • Disabling parallel mode")
        self._log("  • Returning to LOCAL control")
        self.status_label.config(text="Shutting down safely.. .", foreground="orange")
        
        # Disable stop button to prevent multiple clicks
        self.stop_btn. config(state=tk.DISABLED)
        
        # Poll for worker finish
        self.after(500, self._poll_worker_finish)

    def _poll_worker_finish(self):
        """Poll until worker thread finishes"""
        if not self.worker:
            self._finalize_stop()
            return
        
        try:
            worker_alive = self.worker.is_alive()
        except Exception: 
            worker_alive = False
        
        if worker_alive:
            # Still running, check again
            self.after(500, self._poll_worker_finish)
            return
        
        # Worker has stopped
        self._finalize_stop()
    
    def _finalize_stop(self):
        """Finalize the stop process and reset UI"""
        self. start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk. DISABLED)
        self.status_label.config(text="Idle - Safe State", foreground="blue")
        self._log("═══ STOP COMPLETE ═══")
        self._log("✓ Equipment in safe state:")
        self._log("  • Power: 0W")
        self._log("  • Loads: OFF")
        self._log("  • Mode: LOCAL")
        self._log("✓ Ready to start new test.")
        
        messagebox.showinfo(
            "Safe Shutdown Complete", 
            "Equipment has been safely shut down:\n\n"
            "✓ Power ramped to 0W\n"
            "✓ Loads turned OFF\n"
            "✓ Parallel mode disabled\n"
            "✓ Returned to LOCAL control\n\n"
            "System is ready for next test."
        )

    def _on_quit(self):
        """Handle application quit"""
        if self.worker:
            try:
                worker_alive = self.worker.is_alive()
            except Exception: 
                worker_alive = False
            
            if worker_alive: 
                if not messagebox.askyesno("Quit", "Test is running. Stop and quit?"):
                    return
                
                self. worker.request_stop()
                self._log("Quit requested - stopping worker...")
                
                # Wait for worker to stop (with timeout)
                for _ in range(30):  # 3 seconds max
                    try:
                        if not self.worker.is_alive():
                            break
                    except Exception:
                        break
                    time.sleep(0.1)
        
        self._log("Application closing...")
        self.destroy()

    # Plot updates
    def _periodic_plot_poll(self):
        """Poll plot queue and update with combined totals"""
        updated = False
        
        try:
            while True: 
                queue_data = self. plot_queue.get_nowait()
                
                if len(queue_data) != 13:
                    self._log(f"[ERROR] Plot queue data length:  {len(queue_data)} (expected 13)")
                    continue
                
                rel, step_idx, mv, mi, mp, sv, si, sp, energy_wh, capacity_ah, dq_dv, dq_dv_voltage, dq_dv_valid = queue_data
                
                # Calculate totals
                total_power = 0.0
                total_current = 0.0
                avg_voltage = 0.0
                
                if mp is not None and sp is not None:
                    total_power = mp + sp
                elif mp is not None: 
                    total_power = mp
                elif sp is not None: 
                    total_power = sp
                
                if mi is not None and si is not None: 
                    total_current = mi + si
                elif mi is not None:
                    total_current = mi
                elif si is not None:
                    total_current = si
                
                if mv is not None and sv is not None: 
                    avg_voltage = (mv + sv) / 2.0
                elif mv is not None:
                    avg_voltage = mv
                elif sv is not None:  
                    avg_voltage = sv
                
                self. plot_times.append(rel)
                self.plot_total_power.append(total_power)
                self.plot_total_current.append(total_current)
                self.plot_avg_voltage.append(avg_voltage)
                self.plot_cumulative_energy.append(energy_wh if energy_wh is not None else 0.0)
                self.plot_cumulative_capacity.append(capacity_ah if capacity_ah is not None else 0.0)
                
                # Store dQ/dV data if valid
                if dq_dv_valid and abs(dq_dv) < 1000:
                    self.plot_dqdv_voltage.append(dq_dv_voltage)
                    self.plot_dqdv_values.append(dq_dv)
                
                updated = True
                
        except queue.Empty:
            pass
        except Exception as e:  
            self._log(f"[ERROR] Plot poll exception:  {e}")

        if updated:
            try:  
                self._update_plot_lines()
                if hasattr(self, '_update_dqdv_plot'):
                    self._update_dqdv_plot()
            except Exception as e:
                self._log(f"[ERROR] Plot update exception:  {e}")
        
        self.after(500, self._periodic_plot_poll)

    def _update_dqdv_plot(self, redraw: bool = False):
        """Update dQ/dV plot"""
        if not self.plot_dqdv_voltage:
            return
        
        try: 
            # Update line data
            self.line_dqdv. set_data(self.plot_dqdv_voltage, self. plot_dqdv_values)
            
            # Update axis limits
            try:
                self.ax_dqdv.relim()
                self.ax_dqdv. autoscale_view(True, True, True)
                
                # Re-apply plain formatting
                self.ax_dqdv.ticklabel_format(style='plain', axis='both', useOffset=False)
            except Exception:
                pass
            
            # Update statistics
            if hasattr(self, 'dqdv_points_label'):
                self. dqdv_points_label.config(text=str(len(self.plot_dqdv_values)))
                
                if self.plot_dqdv_values:
                    current_dqdv = self.plot_dqdv_values[-1]
                    current_voltage = self.plot_dqdv_voltage[-1]
                    max_dqdv = max(self.plot_dqdv_values)
                    min_dqdv = min(self.plot_dqdv_values)
                    
                    self.dqdv_current_label.config(text=f"{current_dqdv:.4f} Ah/V")
                    self.dqdv_voltage_label.config(text=f"{current_voltage:.3f} V")
                    self.dqdv_max_label.config(text=f"{max_dqdv:.4f} Ah/V")
                    self.dqdv_min_label.config(text=f"{min_dqdv:.4f} Ah/V")
            
            self.fig_dqdv.tight_layout()
            
            if redraw: 
                self.canvas_dqdv.draw_idle()
            else:
                self. canvas_dqdv.draw()
                
        except Exception as e:
            self._log(f"[ERROR] in _update_dqdv_plot: {e}")
    
    def _apply_dqdv_styles(self):
        """Apply styles to dQ/dV plot"""
        try:
            dqdv_color = self. dqdv_color_var.get()
            self.line_dqdv.set_color(dqdv_color)
            self.line_dqdv.set_linestyle(LINE_STYLES. get(self.dqdv_style_var.get(), "-"))
            self.line_dqdv.set_linewidth(float(self.dqdv_width_var.get()))
            self.line_dqdv.set_marker(MARKER_STYLES.get(self.dqdv_marker_var.get(), "o"))
            
            self.ax_dqdv.set_ylabel("dQ/dV (Ah/V)", fontsize=12, fontweight='bold', color=dqdv_color)
            self.ax_dqdv.tick_params(axis='y', labelcolor=dqdv_color)
            
            self._update_dqdv_plot(redraw=True)
            self._log("Applied dQ/dV plot styles.")
        except Exception as e: 
            self._log(f"[ERROR] Applying dQ/dV styles: {e}")
    
    def _choose_dqdv_color(self):
        """Choose color for dQ/dV line"""
        col = colorchooser.askcolor(title="Choose dQ/dV line color", initialcolor=self.dqdv_color_var.get())
        if col and col[1]:
            self. dqdv_color_var. set(col[1])
            self. dqdv_color_display.config(bg=col[1])
            self._apply_dqdv_styles()
    
    def _clear_dqdv_plot(self):
        """Clear dQ/dV plot data"""
        self.plot_dqdv_voltage.clear()
        self.plot_dqdv_values.clear()
        self._update_dqdv_plot(redraw=True)
        self._log("dQ/dV plot cleared")
    
    def _save_dqdv_screenshot(self):
        """Save dQ/dV plot screenshot"""
        path = filedialog.asksaveasfilename(
            title="Save dQ/dV Plot",
            defaultextension=". png",
            filetypes=[("PNG files", "*.png"), ("JPEG files", "*.jpg;*.jpeg"), ("All files", "*.*")],
            initialfile=f"dqdv_plot_{datetime. now().strftime('%Y%m%d_%H%M%S')}.png"
        )
        
        if not path:
            return
        
        try:
            self.fig_dqdv.savefig(path, dpi=300, bbox_inches='tight')
            messagebox.showinfo("Success", f"dQ/dV plot saved to:\n{path}")
            self._log(f"dQ/dV plot saved:  {path}")
        except Exception as e:
            messagebox.showerror("Save Error", str(e))
            self._log(f"[ERROR] Failed to save dQ/dV plot:  {e}")
    
    def _export_dqdv_data(self):
        """Export dQ/dV data to CSV"""
        if not self.plot_dqdv_voltage:
            messagebox.showwarning("No Data", "No dQ/dV data to export")
            return
        
        path = filedialog.asksaveasfilename(
            title="Export dQ/dV Data",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=f"dqdv_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        
        if not path:
            return
        
        try:
            with open(path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['Voltage_V', 'dQ_dV_Ah_per_V'])
                
                for i in range(len(self.plot_dqdv_voltage)):
                    writer.writerow([
                        self.plot_dqdv_voltage[i],
                        self.plot_dqdv_values[i]
                    ])
            
            messagebox. showinfo("Success", f"dQ/dV data exported to:\n{path}")
            self._log(f"dQ/dV data exported: {path}")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))
            self._log(f"[ERROR] dQ/dV export failed: {e}")

    def _update_plot_lines(self, redraw: bool = False):
        """Update plot lines with combined data"""
        if not self.plot_times:
            return
        
        try:
            x = self.plot_times
            
            # Update ALL line data
            self.line_power. set_data(x, self. plot_total_power)
            self.line_current.set_data(x, self. plot_total_current)
            self.line_voltage.set_data(x, self.plot_avg_voltage)
            self.line_energy.set_data(x, self.plot_cumulative_energy)
            
            # Ensure all lines are visible
            self.line_power.set_visible(True)
            self.line_current.set_visible(True)
            self.line_voltage.set_visible(True)
            self.line_energy.set_visible(True)
            
            x_max = max(x) if x else 10
            
            # Update x-axis limits
            self.ax_power.set_xlim(0, max(10, x_max))
            
            # Autoscale ALL y-axes with data-driven limits
            try:
                # Power axis
                if self.plot_total_power:
                    power_vals = [p for p in self. plot_total_power if p and not (p != p) and p >= 0]
                    if power_vals:
                        p_min = min(power_vals)
                        p_max = max(power_vals)
                        margin = (p_max - p_min) * 0.1 if p_max > p_min else max(10, p_max * 0.1)
                        self.ax_power.set_ylim(max(0, p_min - margin), p_max + margin)
                
                # Current axis
                if self.plot_total_current:
                    current_vals = [c for c in self.plot_total_current if c and not (c != c) and c >= 0]
                    if current_vals:
                        c_min = min(current_vals)
                        c_max = max(current_vals)
                        margin = (c_max - c_min) * 0.1 if c_max > c_min else max(1, c_max * 0.1)
                        self.ax_current.set_ylim(max(0, c_min - margin), c_max + margin)
                
                # Voltage axis
                if self.plot_avg_voltage:
                    voltage_vals = [v for v in self.plot_avg_voltage if v and not (v != v) and v >= 0]
                    if voltage_vals:
                        v_min = min(voltage_vals)
                        v_max = max(voltage_vals)
                        margin = (v_max - v_min) * 0.1 if v_max > v_min else max(1, v_max * 0.1)
                        self.ax_voltage.set_ylim(max(0, v_min - margin), v_max + margin)
                
                # Energy axis
                if self.plot_cumulative_energy:
                    energy_vals = [e for e in self.plot_cumulative_energy if e is not None and e >= 0]
                    if energy_vals: 
                        e_min = min(energy_vals)
                        e_max = max(energy_vals)
                        margin = (e_max - e_min) * 0.1 if e_max > e_min else max(0.1, e_max * 0.1)
                        self. ax_energy.set_ylim(max(0, e_min - margin), e_max + margin)
                
                # Re-apply plain formatting after setting limits
                self.ax_power.ticklabel_format(style='plain', axis='y', useOffset=False)
                self. ax_current.ticklabel_format(style='plain', axis='y', useOffset=False)
                self.ax_voltage.ticklabel_format(style='plain', axis='y', useOffset=False)
                self.ax_energy.ticklabel_format(style='plain', axis='y', useOffset=False)
            except Exception as e:
                self._log(f"[DEBUG] Axis scaling error: {e}")

            # Update statistics
            if hasattr(self, 'stats_points_label'):
                self.stats_points_label.config(text=str(len(x)))
                self.stats_time_label. config(text=f"{x_max:.1f} s" if x else "0 s")
                
                # Calculate averages
                valid_power = [p for p in self. plot_total_power if p and not (p != p)]
                valid_current = [c for c in self.plot_total_current if c and not (c != c)]
                valid_voltage = [v for v in self. plot_avg_voltage if v and not (v != v)]
                
                avg_power = sum(valid_power) / len(valid_power) if valid_power else 0.0
                avg_current = sum(valid_current) / len(valid_current) if valid_current else 0.0
                avg_voltage = sum(valid_voltage) / len(valid_voltage) if valid_voltage else 0.0
                
                self.stats_avg_power_label.config(text=f"{avg_power:.2f} W")
                self. stats_avg_current_label. config(text=f"{avg_current:.2f} A")
                self.stats_avg_voltage_label.config(text=f"{avg_voltage:.2f} V")
                
                # Update energy statistics
                if self.plot_cumulative_energy:
                    latest_energy = self.plot_cumulative_energy[-1]
                    latest_capacity = self.plot_cumulative_capacity[-1]
                    self. stats_energy_label.config(text=f"{latest_energy:.4f} Wh")
                    self. stats_capacity_label.config(text=f"{latest_capacity:.4f} Ah")

            self.fig.tight_layout()

            if redraw:
                self.canvas.draw_idle()
            else:
                self.canvas. draw()
                
        except Exception as e: 
            self._log(f"[ERROR] in _update_plot_lines: {e}")
            import traceback
            self._log(traceback.format_exc())

    # Style application
    def _choose_color(self, color_var):
        col = colorchooser.askcolor(title="Choose color", initialcolor=color_var.get())
        if col and col[1]:
            color_var.set(col[1])
            # Update color displays
            if hasattr(self, 'master_color_display') and color_var == self.master_color_var:
                self.master_color_display.config(bg=col[1])
            if hasattr(self, 'slave_color_display') and color_var == self.slave_color_var:
                self. slave_color_display.config(bg=col[1])
            self._apply_line_styles()
    
    def _apply_line_styles(self):
        """Apply line styles to all plot lines"""
        try:
            # Update power line
            power_color = self.power_color_var.get()
            self.line_power.set_color(power_color)
            self.line_power.set_linestyle(LINE_STYLES. get(self.power_style_var.get(), "-"))
            self.line_power.set_linewidth(float(self.power_width_var.get()))
            self.line_power.set_zorder(10)  # Ensure it's on top
            self.ax_power.set_ylabel("Total Power (W)", fontsize=11, fontweight='bold', color=power_color)
            self.ax_power.tick_params(axis='y', labelcolor=power_color, labelsize=9)
            
            # Update current line
            current_color = self.current_color_var.get()
            self.line_current.set_color(current_color)
            self.line_current.set_linestyle(LINE_STYLES.get(self.current_style_var.get(), "-"))
            self.line_current.set_linewidth(float(self.current_width_var. get()))
            self.line_current.set_zorder(9)
            self.ax_current.set_ylabel("Total Current (A)", fontsize=11, fontweight='bold', color=current_color)
            self.ax_current. tick_params(axis='y', labelcolor=current_color, labelsize=9)
            
            # Update voltage line
            voltage_color = self.voltage_color_var.get()
            self.line_voltage.set_color(voltage_color)
            self.line_voltage.set_linestyle(LINE_STYLES.get(self.voltage_style_var. get(), "-"))
            self.line_voltage.set_linewidth(float(self.voltage_width_var.get()))
            self.line_voltage.set_zorder(8)
            self.ax_voltage.set_ylabel("Avg Voltage (V)", fontsize=11, fontweight='bold', color=voltage_color)
            self.ax_voltage.tick_params(axis='y', labelcolor=voltage_color, labelsize=9)
            
            # Update energy line
            energy_color = self.energy_color_var.get()
            self.line_energy.set_color(energy_color)
            self.line_energy.set_linestyle(LINE_STYLES.get(self.energy_style_var.get(), "-"))
            self.line_energy.set_linewidth(float(self.energy_width_var.get()))
            self.line_energy.set_zorder(7)
            self.ax_energy.set_ylabel("Energy (Wh)", fontsize=11, fontweight='bold', color=energy_color)
            self.ax_energy.tick_params(axis='y', labelcolor=energy_color, labelsize=9)
            
            # Re-apply plain formatting
            self. ax_power.ticklabel_format(style='plain', axis='y', useOffset=False)
            self.ax_current.ticklabel_format(style='plain', axis='y', useOffset=False)
            self.ax_voltage.ticklabel_format(style='plain', axis='y', useOffset=False)
            self.ax_energy.ticklabel_format(style='plain', axis='y', useOffset=False)
            
            # Update legend
            lines = [self. line_power, self.line_current, self.line_voltage, self.line_energy]
            labels = [l.get_label() for l in lines]
            self.ax_power.legend(lines, labels, loc='upper left', fontsize=9, framealpha=0.9)
            
            self._update_plot_lines(redraw=True)
            self._log("Applied plot styles.")
        except Exception as e: 
            self._log(f"[ERROR] Applying plot styles:  {e}")
            
    def _choose_line_color(self, color_var, display_canvas):
        """Choose color for a specific line"""
        col = colorchooser.askcolor(title="Choose line color", initialcolor=color_var.get())
        if col and col[1]:
            color_var.set(col[1])
            display_canvas.config(bg=col[1])
            self._apply_line_styles()

    # Config note
    def _get_config_note_text(self) -> str:
        """Generate configuration note text"""
        fname = self.config_filename_var. get().strip()
        note = f"Master-Slave Configuration\n"
        note += f"==========================\n\n"
        note += f"Config File: {fname if fname else 'Not set'}\n\n"
        note += f"Master Port:  {self.settings. master_port}\n"
        note += f"Slave Port:  {self.settings.slave_port}\n"
        note += f"Baud Rate: {self.settings.baud_rate}\n\n"
        note += f"Imbalance Monitoring: {'ENABLED' if self. settings.imbalance_check_enabled else 'DISABLED ⚠'}\n"
        note += f"Imbalance Limit: {self.settings. imbalance_limit*100:.1f}%\n"  # Fixed:  no space
        note += f"Transition Limit: {self.settings.imbalance_transition*100:.1f}%\n"  # Fixed: no space
        note += f"Transition Duration: {self.settings.transition_duration}s\n\n"
        note += f"Voltage Cutoff: {self.settings.cutoff_voltage}V\n"
        note += f"Cutoff Safety: {'Enabled' if self.settings.cutoff_safety_enabled else 'Disabled'}\n\n"
        note += f"Email Alerts: {'Enabled' if self.settings. email_enabled else 'Disabled'}\n"
        note += f"Profile Steps: {len(self.settings.profile)}\n"
        return note

    def _set_config_note_text(self, text: str):
        self.config_note_text. configure(state=tk.NORMAL)
        self.config_note_text. delete("1.0", tk. END)
        self.config_note_text.insert(tk. END, text)
        self._apply_note_font()
        self.config_note_text.configure(state=tk.DISABLED)

    def _update_config_note(self):
        """Update configuration note display"""
        try:
            note_text = self._get_config_note_text()
            if hasattr(self, 'config_note_text'):
                self.config_note_text.config(state=tk.NORMAL)
                self.config_note_text.delete('1.0', tk.END)
                self.config_note_text.insert('1.0', note_text)
                self.config_note_text.config(state=tk. DISABLED)
        except Exception as e:
            self._log(f"[ERROR] Updating config note: {e}")

    def _apply_note_font(self):
        sz = int(self.config_note_fontsize. get())
        f = tkfont.Font(family="TkDefaultFont", size=sz)
        self.config_note_text.configure(font=f)

    # Save screenshot
    def _save_plot_screenshot(self):
        path = filedialog.asksaveasfilename(title="Save plot", defaultextension=".png",
                                            filetypes=[("PNG files", "*.png"), ("JPEG files", "*.jpg;*.jpeg"), ("All files", "*.*")])
        if not path:
            return
        try:
            self.fig.savefig(path)
            messagebox.showinfo("Saved", f"Plot saved to {path}")
            self._log(f"Plot screenshot saved:  {path}")
        except Exception as e:
            messagebox.showerror("Save error", str(e))
            self._log(f"[ERROR] Failed to save plot: {e}")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
