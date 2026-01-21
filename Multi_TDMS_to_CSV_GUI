# -*- coding: utf-8 -*-
"""
Created on Fri Nov 21 16:44:43 2025

@author: ssalvi
"""


import os
import re
import sys
import threading
import queue
import traceback
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from tkinter.scrolledtext import ScrolledText

import pandas as pd
import pandas.api.types as ptypes
import nptdms
import numpy as np

PREFERRED_TIMESTAMP_NAMES = [
    "/TimeStamps/TIME 10 Hz",
    "Date/Time (Excel Format)",
    "Date/Time",
    "DateTime",
    "Timestamp",
    "TimeStamp",
    "Time",
]
TICK_NAME_PATTERN = r"(?i)tick"
SECONDS_PER_DAY = 24 * 3600


def flatten_and_clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    if isinstance(df.columns, pd.MultiIndex):
        cols = []
        for col in df.columns:
            parts = [str(p).strip() for p in col if p is not None and str(p).strip() != ""]
            cols.append("/".join(parts) if parts else "")
        df.columns = cols
    else:
        df.columns = [str(c) for c in df.columns]
    cleaned = []
    for c in df.columns:
        c2 = re.sub(r"^\s+|\s+$", "", c)
        c2 = c2.replace("'", "").replace('"', "")
        c2 = re.sub(r"[\\]+", "/", c2)
        c2 = re.sub(r"[<>:|?*]", "_", c2)
        cleaned.append(c2)
    df.columns = cleaned
    return df


def _parse_time_string_to_seconds(s: str):
    if s is None:
        return float("nan")
    s = str(s).strip()
    if s == "":
        return float("nan")
    if ":" in s:
        parts = s.split(":")
        try:
            if len(parts) == 2:
                mm = int(parts[0]); ss = float(parts[1]); return mm * 60.0 + ss
            if len(parts) == 3:
                hh = int(parts[0]); mm = int(parts[1]); ss = float(parts[2])
                return hh * 3600.0 + mm * 60.0 + ss
        except Exception:
            try:
                if len(parts) >= 2:
                    mm = float(parts[-2]); ss = float(parts[-1]); return mm * 60.0 + ss
            except Exception:
                return float("nan")
    try:
        return float(s)
    except Exception:
        return float("nan")


def ts_series_to_seconds_and_unit(ts_series: pd.Series):
    try:
        if ptypes.is_datetime64_any_dtype(ts_series.dtype):
            sec = ts_series.astype("int64") / 1e9
            return sec.astype(float), "datetime_epoch_seconds"
    except Exception:
        pass
    try:
        if ptypes.is_timedelta64_dtype(ts_series.dtype):
            return ts_series.dt.total_seconds().astype(float), "timedelta_seconds"
    except Exception:
        pass
    num = pd.to_numeric(ts_series, errors="coerce")
    valid_frac = (num.notna().sum() / len(ts_series)) if len(ts_series) > 0 else 0.0
    if valid_frac > 0.5:
        med = num.median()
        if med > 1e12:
            return (num / 1e9).astype(float), "epoch_ns"
        if med > 10000:
            return (num * SECONDS_PER_DAY).astype(float), "days"
        return num.astype(float), "seconds"
    parsed = ts_series.apply(_parse_time_string_to_seconds)
    if parsed.notna().sum() > 0:
        return parsed.astype(float), "hms_or_mmss"
    return num.astype(float), "seconds"


def convert_timestamp_series_to_relative_seconds(ts_series: pd.Series, baseline_value_raw=None):
    abs_seconds_series, unit = ts_series_to_seconds_and_unit(ts_series)
    if abs_seconds_series.isna().all():
        return None, None, None, "No parsable timestamps."
    baseline_seconds = None
    if baseline_value_raw is not None:
        try:
            if isinstance(baseline_value_raw, (pd.Timestamp,)):
                baseline_seconds = pd.to_datetime(baseline_value_raw).astype("int64") / 1e9
            else:
                baseline_num = float(baseline_value_raw)
                if unit == "epoch_ns":
                    baseline_seconds = baseline_num / 1e9
                elif unit == "days":
                    baseline_seconds = baseline_num * SECONDS_PER_DAY
                else:
                    baseline_seconds = baseline_num
        except Exception:
            if isinstance(baseline_value_raw, str) and ":" in baseline_value_raw:
                baseline_seconds = _parse_time_string_to_seconds(baseline_value_raw)
            else:
                baseline_seconds = None
    if baseline_seconds is None:
        idx = abs_seconds_series.first_valid_index()
        if idx is None:
            return None, None, None, "No valid timestamp index to set baseline."
        baseline_seconds = float(abs_seconds_series.loc[idx])
    baseline_seconds = float(baseline_seconds)
    time_sec = (abs_seconds_series - baseline_seconds).astype(float).round(6)
    return time_sec, baseline_seconds, unit, None


class ConverterGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("multi_tdms_to_CSV - GUI (with downsample option)")
        self.minsize(980, 680)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.queue = queue.Queue()
        self.worker_thread = None
        self.stop_event = threading.Event()
        self._after_id = None
        self._is_closing = False

        # detection state
        self.scanned_channels = []
        self.detected_timestamp_col = None
        self.detected_tick_col = None
        self.first_file_baseline_raw = None
        self.first_file_baseline_source = None

        self._build_ui()
        self._after_id = self.after(200, self._process_queue)

    def _build_ui(self):
        pad = 8
        frm = ttk.Frame(self, padding=pad)
        frm.pack(fill=tk.BOTH, expand=True)

        top = ttk.Frame(frm); top.pack(fill=tk.X, pady=(0, pad))
        ttk.Label(top, text="Source folder:").grid(row=0, column=0, sticky=tk.W)
        self.src_var = tk.StringVar(); ttk.Entry(top, textvariable=self.src_var, width=80).grid(row=0, column=1, padx=6)
        ttk.Button(top, text="Browse...", command=self.browse_source).grid(row=0, column=2)
        ttk.Label(top, text="Destination folder:").grid(row=1, column=0, sticky=tk.W)
        self.dst_var = tk.StringVar(); ttk.Entry(top, textvariable=self.dst_var, width=80).grid(row=1, column=1, padx=6)
        ttk.Button(top, text="Browse...", command=self.browse_dest).grid(row=1, column=2)

        opts = ttk.Labelframe(frm, text="Options (Global trigger across all files)", padding=pad)
        opts.pack(fill=tk.X, pady=(0, pad))
        ttk.Label(opts, text="Trigger channel (optional):").grid(row=0, column=0, sticky=tk.W)
        self.trigger_combobox = ttk.Combobox(opts, values=[], width=64); self.trigger_combobox.set(""); self.trigger_combobox.grid(row=0, column=1, sticky=tk.W)
        ttk.Label(opts, text="Trigger Signal (optional):").grid(row=1, column=0, sticky=tk.W)
        self.trigger_signal_var = tk.StringVar(value="8"); ttk.Entry(opts, textvariable=self.trigger_signal_var, width=12).grid(row=1, column=1, sticky=tk.W)
        ttk.Label(opts, text="Output frequency (Hz) (optional):").grid(row=2, column=0, sticky=tk.W)
        self.output_freq_var = tk.StringVar(value="")  # empty means no downsample
        ttk.Entry(opts, textvariable=self.output_freq_var, width=12).grid(row=2, column=1, sticky=tk.W)
        ttk.Label(opts, text="(e.g. 1 for 1 Hz from 10 Hz)").grid(row=2, column=1, sticky=tk.W, padx=(120,0))
        self.sort_files_var = tk.BooleanVar(value=True); ttk.Checkbutton(opts, text="Sort files alphabetically", variable=self.sort_files_var).grid(row=3, column=0, sticky=tk.W)
        self.streaming_var = tk.BooleanVar(value=False); ttk.Checkbutton(opts, text="Memory-safe streaming (append per file)", variable=self.streaming_var).grid(row=3, column=1, sticky=tk.W)
        ttk.Label(opts, text="Output filename:").grid(row=4, column=0, sticky=tk.W, pady=(6, 0))
        self.outname_var = tk.StringVar(value=""); ttk.Entry(opts, textvariable=self.outname_var, width=50).grid(row=4, column=1, sticky=tk.W, pady=(6, 0))

        middle = ttk.Frame(frm); middle.pack(fill=tk.BOTH, expand=True, pady=(0, pad))
        left = ttk.Frame(middle); left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Label(left, text="TDMS files found:").pack(anchor=tk.W)
        self.file_listbox = tk.Listbox(left, height=8); self.file_listbox.pack(fill=tk.BOTH, padx=(0, 6), pady=(4, 0))
        ttk.Label(left, text="Available channels (select to include in final CSV):").pack(anchor=tk.W, pady=(6, 0))
        self.channel_listbox = tk.Listbox(left, selectmode=tk.MULTIPLE, height=12); self.channel_listbox.pack(fill=tk.BOTH, padx=(0, 6), pady=(4, 0))
        ch_btns = ttk.Frame(left); ch_btns.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(ch_btns, text="Select All", command=self.select_all_channels).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(ch_btns, text="Clear", command=self.clear_channel_selection).pack(side=tk.LEFT)

        right = ttk.Frame(middle, width=320); right.pack(side=tk.RIGHT, fill=tk.Y)
        ttk.Button(right, text="Scan Folder", command=self.scan_folder).pack(fill=tk.X, pady=(0, 6))
        ttk.Button(right, text="Start Conversion", command=lambda: self.start_conversion(downsample=False)).pack(fill=tk.X, pady=(0, 6))
        ttk.Button(right, text="Start Conversion (with frequency)", command=lambda: self.start_conversion(downsample=True)).pack(fill=tk.X, pady=(0, 6))
        ttk.Button(right, text="Stop", command=self.request_stop).pack(fill=tk.X, pady=(0, 6))
        ttk.Button(right, text="Exit", command=self.on_close).pack(fill=tk.X, pady=(0, 6))

        bottom = ttk.Frame(frm); bottom.pack(fill=tk.BOTH, expand=True)
        self.progress = ttk.Progressbar(bottom, mode="determinate"); self.progress.pack(fill=tk.X, pady=(6, 4))
        ttk.Label(bottom, text="Log:").pack(anchor=tk.W)
        self.log_widget = ScrolledText(bottom, height=12); self.log_widget.pack(fill=tk.BOTH, expand=True)

    def browse_source(self):
        p = filedialog.askdirectory(title="Select Folder Containing TDMS Files")
        if p:
            self.src_var.set(p); base = os.path.basename(os.path.normpath(p))
            if not self.outname_var.get(): self.outname_var.set(f"{base}_Combined.csv")
            self.scan_folder()

    def browse_dest(self):
        p = filedialog.askdirectory(title="Select Destination Folder")
        if p: self.dst_var.set(p)

    def select_all_channels(self): self.channel_listbox.select_set(0, tk.END)
    def clear_channel_selection(self): self.channel_listbox.select_clear(0, tk.END)

    def scan_folder(self):
        src = self.src_var.get().strip()
        self.file_listbox.delete(0, tk.END); self.channel_listbox.delete(0, tk.END); self.trigger_combobox['values'] = []
        self.scanned_channels = []; self.detected_timestamp_col = None; self.detected_tick_col = None
        self.first_file_baseline_raw = None; self.first_file_baseline_source = None

        if not src or not os.path.isdir(src):
            messagebox.showwarning("No folder", "Please select a valid source folder first."); return

        files = [f for f in os.listdir(src) if f.lower().endswith(".tdms")]
        if self.sort_files_var.get(): files = sorted(files)
        for f in files: self.file_listbox.insert(tk.END, f)
        self.log(f"Scanning folder: {src} — {len(files)} .tdms files found")

        channels = []; first_seen = False
        for fname in files:
            fpath = os.path.join(src, fname)
            try:
                with nptdms.TdmsFile.open(fpath) as td:
                    df = td.as_dataframe()
                df = flatten_and_clean_columns(df)
                try:
                    config_idx = next(idx for idx, c in enumerate(df.columns) if "Config Tree" in c)
                    df = df.iloc[:, :config_idx]
                except Exception:
                    pass
                for c in df.columns:
                    if c not in channels: channels.append(c)
                if not first_seen:
                    ts = None; tkcol = None
                    for name in PREFERRED_TIMESTAMP_NAMES:
                        if name in df.columns:
                            ts = name; break
                    if ts is None:
                        for c in df.columns:
                            lc = c.lower()
                            if ("date" in lc and "time" in lc) or "timestamp" in lc or re.search(r"\btime\b", lc):
                                ts = c; break
                    for c in df.columns:
                        if re.search(TICK_NAME_PATTERN, c):
                            tkcol = c; break
                    self.detected_timestamp_col = ts; self.detected_tick_col = tkcol
                    if ts and ts in df.columns:
                        try:
                            idx = df[ts].first_valid_index()
                            if idx is not None:
                                self.first_file_baseline_raw = df[ts].loc[idx]; self.first_file_baseline_source = "timestamp"
                                self.log(f"Captured timestamp baseline from '{ts}': {self.first_file_baseline_raw}")
                        except Exception:
                            pass
                    elif tkcol and tkcol in df.columns:
                        try:
                            tser = pd.to_numeric(df[tkcol], errors="coerce")
                            if tser.notna().any():
                                self.first_file_baseline_raw = float(tser.loc[tser.first_valid_index()]); self.first_file_baseline_source = "tick"
                                self.log(f"Captured tick baseline from '{tkcol}': {self.first_file_baseline_raw}")
                        except Exception:
                            pass
                    first_seen = True
            except Exception as e:
                self.log(f"Warning scanning {fname}: {e}")

        channels.sort(); self.scanned_channels = channels
        for c in channels: self.channel_listbox.insert(tk.END, c)
        self.trigger_combobox['values'] = [""] + channels; self.trigger_combobox.set("")
        self.log(f"Scan complete. timestamp_col={self.detected_timestamp_col}, tick_col={self.detected_tick_col}, baseline={self.first_file_baseline_raw} (source={self.first_file_baseline_source})")

    def start_conversion(self, downsample=False):
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("Busy", "Conversion already running."); return

        src = self.src_var.get().strip(); dst = self.dst_var.get().strip()
        if not src or not os.path.isdir(src): messagebox.showerror("Source folder required", "Please select a valid source folder."); return
        if not dst or not os.path.isdir(dst): messagebox.showerror("Destination folder required", "Please select a valid destination folder."); return

        files = [f for f in os.listdir(src) if f.lower().endswith(".tdms")]
        if self.sort_files_var.get(): files = sorted(files)
        if not files: messagebox.showwarning("No files", "No .tdms files found in the source folder."); return

        outname = self.outname_var.get().strip()
        if not outname: outname = os.path.basename(os.path.normpath(src)) + "_Combined.csv"
        outpath = os.path.join(dst, outname)

        sel = self.channel_listbox.curselection()
        selected_channels = [self.scanned_channels[i] for i in sel] if sel else list(self.scanned_channels)

        trigger_channel = self.trigger_combobox.get().strip() or None
        try:
            trigger_signal = float(self.trigger_signal_var.get().strip()) if self.trigger_signal_var.get().strip() != "" else None
        except Exception:
            trigger_signal = None; self.log("Trigger Signal parse failed; trigger-based baseline will not be used.")

        # if downsample requested and streaming on, block
        if downsample and self.streaming_var.get():
            ok = messagebox.askyesno("Streaming incompatible with downsample",
                                     "You selected streaming and also requested downsampling. Downsampling requires combined data. Disable streaming and continue?")
            if not ok:
                return
            self.streaming_var.set(False)

        # parse output frequency
        output_freq = None
        if downsample:
            val = self.output_freq_var.get().strip()
            if val != "":
                try:
                    output_freq = float(val)
                    if output_freq <= 0:
                        raise ValueError()
                except Exception:
                    messagebox.showerror("Invalid Output Frequency", "Please enter a valid positive number for Output frequency (Hz).")
                    return
            else:
                messagebox.showerror("Output Frequency required", "Please enter Output frequency (Hz) when choosing 'Start Conversion (with frequency)'.")
                return

        options = {
            "src": src, "files": files, "outpath": outpath,
            "trigger_col": trigger_channel,
            "timestamp_col": self.detected_timestamp_col, "tick_col": self.detected_tick_col,
            "streaming": self.streaming_var.get(), "selected_channels": selected_channels,
            "first_file_baseline_raw": self.first_file_baseline_raw, "first_file_baseline_source": self.first_file_baseline_source,
            "trigger_signal": trigger_signal, "output_freq": output_freq,
        }

        try: self.log_widget.delete("1.0", tk.END)
        except Exception: pass
        self.progress["value"] = 0; self.progress["maximum"] = len(files)
        self.stop_event.clear()
        self.worker_thread = threading.Thread(target=self.worker_convert, args=(options,))
        self.worker_thread.daemon = True; self.worker_thread.start()

    def request_stop(self):
        if self.worker_thread and self.worker_thread.is_alive():
            self.stop_event.set(); self.log("Stop requested — worker will stop after current operation.")
        else:
            self.log("No active task to stop.")

    def on_close(self):
        self._is_closing = True; self.stop_event.set()
        try:
            if self._after_id: self.after_cancel(self._after_id); self._after_id = None
        except Exception:
            pass
        try: self.destroy()
        except Exception:
            pass

    def log(self, msg):
        try: self.queue.put(("log", msg))
        except Exception: pass

    def _process_queue(self):
        try:
            while True:
                item = self.queue.get_nowait()
                typ, payload = item
                if typ == "log":
                    try: self.log_widget.insert(tk.END, payload + "\n"); self.log_widget.see(tk.END)
                    except tk.TclError: return
                elif typ == "progress":
                    try: self.progress["value"] = payload
                    except tk.TclError: return
                elif typ == "done":
                    try: messagebox.showinfo("Done", payload)
                    except tk.TclError: pass
                elif typ == "error":
                    try: messagebox.showerror("Error", payload)
                    except tk.TclError: pass
        except queue.Empty: pass
        except tk.TclError: return
        try:
            if not self._is_closing and self.winfo_exists(): self._after_id = self.after(200, self._process_queue)
            else: self._after_id = None
        except tk.TclError:
            self._after_id = None

    def worker_convert(self, options):
        src = options["src"]; files = options["files"]; outpath = options["outpath"]
        trigger_col = options["trigger_col"]; timestamp_col = options["timestamp_col"]; tick_col = options["tick_col"]
        streaming = options["streaming"]; selected_channels = options["selected_channels"]
        first_file_baseline_raw = options.get("first_file_baseline_raw", None)
        first_file_baseline_source = options.get("first_file_baseline_source", None)
        trigger_signal = options.get("trigger_signal", None)
        output_freq = options.get("output_freq", None)

        try:
            self.log(f"Starting conversion: {len(files)} files -> {outpath}")
            first_write = True
            time_cols = ["Time (sec)", "Time (min)", "Time (hour)"]
            template_cols = time_cols + selected_channels + ["Source File"]

            combined_out_parts = []
            combined_full_parts = []

            # Determine scanned/derived baseline preference
            global_tick_baseline = None
            global_timestamp_baseline_raw = None
            if first_file_baseline_source == "timestamp" and first_file_baseline_raw is not None:
                global_timestamp_baseline_raw = first_file_baseline_raw; self.log(f"Using scanned timestamp baseline for global: {global_timestamp_baseline_raw}")
            elif first_file_baseline_source == "tick" and first_file_baseline_raw is not None:
                try: global_tick_baseline = float(first_file_baseline_raw); self.log(f"Using scanned tick baseline for global: {global_tick_baseline}")
                except Exception: global_tick_baseline = None

            if global_timestamp_baseline_raw is None and global_tick_baseline is None and len(files) > 0:
                try:
                    with nptdms.TdmsFile.open(os.path.join(src, files[0])) as td0:
                        df0 = flatten_and_clean_columns(td0.as_dataframe())
                    if timestamp_col and timestamp_col in df0.columns:
                        idx = df0[timestamp_col].first_valid_index()
                        if idx is not None:
                            global_timestamp_baseline_raw = df0[timestamp_col].loc[idx]; self.log(f"Derived timestamp baseline from first file: {global_timestamp_baseline_raw}")
                    if global_timestamp_baseline_raw is None and tick_col and tick_col in df0.columns:
                        t0 = pd.to_numeric(df0[tick_col], errors="coerce")
                        if t0.notna().any():
                            global_tick_baseline = float(t0.loc[t0.first_valid_index()]); self.log(f"Derived tick baseline from first file: {global_tick_baseline}")
                except Exception as e:
                    self.log(f"Warning deriving baseline from first file: {e}")

            # Read files and build combined minimal frames + output placeholders
            for idx, fname in enumerate(files, start=1):
                if self.stop_event.is_set(): self.log("Stopping as requested."); break
                fpath = os.path.join(src, fname); self.log(f"Reading ({idx}/{len(files)}): {fname}")
                try:
                    with nptdms.TdmsFile.open(fpath) as td: df_full = td.as_dataframe()
                except Exception as e:
                    self.log(f"❌ Error reading {fname}: {e}"); continue

                df_full = flatten_and_clean_columns(df_full)
                try:
                    config_idx = next(i for i, c in enumerate(df_full.columns) if "Config Tree" in c)
                    df_full = df_full.iloc[:, :config_idx]
                except Exception:
                    pass

                cols_needed = []
                if timestamp_col and timestamp_col in df_full.columns: cols_needed.append(timestamp_col)
                if tick_col and tick_col in df_full.columns: cols_needed.append(tick_col)
                if trigger_col:
                    if trigger_col in df_full.columns: cols_needed.append(trigger_col)
                    else:
                        seg = trigger_col.split("/")[-1].strip()
                        cand = next((c for c in df_full.columns if seg in c), None)
                        if cand: cols_needed.append(cand)
                mini = df_full.loc[:, cols_needed].copy() if cols_needed else pd.DataFrame(index=df_full.index)
                combined_full_parts.append(mini)

                keep = [c for c in selected_channels if c in df_full.columns]
                df_out = df_full.loc[:, keep].copy() if keep else pd.DataFrame(index=df_full.index)
                for c in selected_channels:
                    if c not in df_out.columns: df_out[c] = pd.NA
                if selected_channels: df_out = df_out[selected_channels]
                df_out["Source File"] = fname
                df_out.insert(0, "Time (sec)", pd.NA); df_out.insert(1, "Time (min)", pd.NA); df_out.insert(2, "Time (hour)", pd.NA)
                for col in template_cols:
                    if col not in df_out.columns: df_out[col] = pd.NA
                df_out = df_out[template_cols]

                if streaming:
                    if trigger_col:
                        self.log("Streaming incompatible with trigger-based global baseline (should have been blocked earlier).")
                        self.log("Skipping file in streaming mode.")
                        continue
                    mode = "w" if first_write else "a"; header = first_write
                    try:
                        df_out.to_csv(outpath, index=False, mode=mode, header=header, columns=template_cols)
                        self.log(f"Written (stream) {fname} -> {outpath} (header={header})")
                    except Exception as e:
                        self.log(f"❌ Error writing {fname} to CSV: {e}")
                    first_write = False
                else:
                    combined_out_parts.append(df_out)

                self.queue.put(("progress", idx))

            # Combine and compute global times
            if not streaming and not self.stop_event.is_set():
                if len(combined_out_parts) == 0:
                    self.log("No data collected; nothing to write.")
                else:
                    combined_out_df = pd.concat(combined_out_parts, ignore_index=True)
                    combined_full_df = pd.concat(combined_full_parts, ignore_index=True) if combined_full_parts else pd.DataFrame()

                    # alignment
                    if len(combined_out_df) != len(combined_full_df):
                        self.log(f"⚠️ Length mismatch: combined_out_df={len(combined_out_df)} rows, combined_full_df={len(combined_full_df)} rows")
                        n = min(len(combined_out_df), len(combined_full_df))
                        combined_out_df = combined_out_df.iloc[:n].reset_index(drop=True)
                        combined_full_df = combined_full_df.iloc[:n].reset_index(drop=True)
                        self.log(f"Trimmed both to {n} rows for alignment")
                    else:
                        combined_out_df = combined_out_df.reset_index(drop=True)
                        combined_full_df = combined_full_df.reset_index(drop=True)

                    baseline_assigned = False
                    chosen_trigger_col = None

                    if trigger_col:
                        if trigger_col in combined_full_df.columns:
                            chosen_trigger_col = trigger_col
                        else:
                            seg = trigger_col.split("/")[-1].strip()
                            cand = next((c for c in combined_full_df.columns if c.endswith(seg)), None)
                            if cand is None:
                                cand = next((c for c in combined_full_df.columns if seg in c), None)
                            if cand:
                                chosen_trigger_col = cand
                                self.log(f"Trigger column name '{trigger_col}' not found; using candidate '{cand}'")
                            else:
                                self.log(f"Trigger column '{trigger_col}' not found in combined columns (showing first 20): {list(combined_full_df.columns)[:20]}")

                    # trigger detection
                    if chosen_trigger_col and chosen_trigger_col in combined_full_df.columns and trigger_signal is not None:
                        thr = float(trigger_signal)
                        trig_raw = combined_full_df[chosen_trigger_col]
                        trig_numeric = pd.to_numeric(trig_raw, errors="coerce")
                        self.log(f"Trigger column used for detection: '{chosen_trigger_col}'. Numeric samples: {int(trig_numeric.notna().sum())}/{len(trig_numeric)}")
                        preview_idx = trig_numeric.dropna().index[:40].tolist()
                        preview = [(int(i), float(trig_numeric.loc[i])) for i in preview_idx]
                        if preview: self.log(f"Trigger sample preview (index,value) first {len(preview)} numeric: {preview}")

                        prev = trig_numeric.shift(1, fill_value=thr - 1.0)
                        rising_mask = (trig_numeric > thr) & (prev <= thr)

                        if rising_mask.any():
                            trans_idx = int(np.where(rising_mask.values)[0][0])
                            self.log(f"Rising-edge detected at combined index {trans_idx}, trigger value={float(trig_numeric.loc[trans_idx])}")
                        else:
                            gt_mask = trig_numeric > thr
                            if gt_mask.any():
                                trans_idx = int(np.where(gt_mask.values)[0][0])
                                self.log(f"No rising-edge; first > threshold at combined index {trans_idx}, trigger value={float(trig_numeric.loc[trans_idx])}")
                            else:
                                trans_idx = None
                                self.log(f"No trigger value > {thr} found across combined data")

                        if trans_idx is not None:
                            # timestamp baseline from same series
                            if timestamp_col and timestamp_col in combined_full_df.columns:
                                ts_series = combined_full_df[timestamp_col]
                                abs_seconds_series, unit = ts_series_to_seconds_and_unit(ts_series)
                                if abs_seconds_series.notna().any() and pd.notna(abs_seconds_series.loc[trans_idx]):
                                    baseline_seconds = float(abs_seconds_series.loc[trans_idx])
                                    tsec_series = (abs_seconds_series - baseline_seconds).astype(float).round(6)
                                    combined_out_df["Time (sec)"] = tsec_series.values
                                    combined_out_df["Time (min)"] = (tsec_series / 60.0).round(6).values
                                    combined_out_df["Time (hour)"] = (tsec_series / 3600.0).round(6).values
                                    baseline_assigned = True
                                    self.log(f"Global baseline assigned at combined index {trans_idx} using timestamp; baseline_seconds={baseline_seconds} (unit={unit})")
                                else:
                                    self.log("Timestamp series present but could not derive absolute seconds at trigger row; will try tick fallback")
                            if not baseline_assigned and tick_col and tick_col in combined_full_df.columns:
                                tick_series = pd.to_numeric(combined_full_df[tick_col], errors="coerce")
                                if tick_series.notna().any() and pd.notna(tick_series.loc[trans_idx]):
                                    baseline_tick = float(tick_series.loc[trans_idx])
                                    tsec_series = ((tick_series - baseline_tick) / 1000.0).astype(float).round(6)
                                    combined_out_df["Time (sec)"] = tsec_series.values
                                    combined_out_df["Time (min)"] = (tsec_series / 60.0).round(6).values
                                    combined_out_df["Time (hour)"] = (tsec_series / 3600.0).round(6).values
                                    baseline_assigned = True
                                    self.log(f"Global baseline assigned at combined index {trans_idx} using tick baseline={baseline_tick}")
                                else:
                                    self.log("Tick series present but could not derive numeric baseline at trigger row")

                    # fallback if no trigger baseline assigned
                    if not baseline_assigned:
                        if global_timestamp_baseline_raw is not None and timestamp_col and timestamp_col in combined_full_df.columns:
                            ts_series = combined_full_df[timestamp_col]
                            abs_seconds_series, unit = ts_series_to_seconds_and_unit(ts_series)
                            if abs_seconds_series.notna().any():
                                try:
                                    baseline_seconds = None
                                    if isinstance(global_timestamp_baseline_raw, (pd.Timestamp,)):
                                        baseline_seconds = pd.to_datetime(global_timestamp_baseline_raw).astype("int64") / 1e9
                                    else:
                                        baseline_num = float(global_timestamp_baseline_raw)
                                        if unit == "days":
                                            baseline_seconds = baseline_num * SECONDS_PER_DAY
                                        elif unit == "epoch_ns":
                                            baseline_seconds = baseline_num / 1e9
                                        else:
                                            baseline_seconds = baseline_num
                                    baseline_seconds = float(baseline_seconds)
                                except Exception:
                                    idx0 = abs_seconds_series.first_valid_index()
                                    baseline_seconds = float(abs_seconds_series.loc[idx0])
                                tsec_series = (abs_seconds_series - baseline_seconds).astype(float).round(6)
                                combined_out_df["Time (sec)"] = tsec_series.values
                                combined_out_df["Time (min)"] = (tsec_series / 60.0).round(6).values
                                combined_out_df["Time (hour)"] = (tsec_series / 3600.0).round(6).values
                                baseline_assigned = True
                                self.log(f"Global time computed using scanned/derived timestamp baseline (baseline_seconds={baseline_seconds})")
                        if not baseline_assigned and global_tick_baseline is not None and tick_col and tick_col in combined_full_df.columns:
                            tick_series = pd.to_numeric(combined_full_df[tick_col], errors="coerce")
                            if tick_series.notna().any():
                                tsec_series = ((tick_series - float(global_tick_baseline)) / 1000.0).astype(float).round(6)
                                combined_out_df["Time (sec)"] = tsec_series.values
                                combined_out_df["Time (min)"] = (tsec_series / 60.0).round(6).values
                                combined_out_df["Time (hour)"] = (tsec_series / 3600.0).round(6).values
                                baseline_assigned = True
                                self.log(f"Global time computed using scanned/derived tick baseline {global_tick_baseline}")
                        if not baseline_assigned and timestamp_col and timestamp_col in combined_full_df.columns:
                            ts_series = combined_full_df[timestamp_col]
                            abs_seconds_series, unit = ts_series_to_seconds_and_unit(ts_series)
                            if abs_seconds_series.notna().any():
                                idx0 = abs_seconds_series.first_valid_index()
                                baseline_seconds = float(abs_seconds_series.loc[idx0])
                                tsec_series = (abs_seconds_series - baseline_seconds).astype(float).round(6)
                                combined_out_df["Time (sec)"] = tsec_series.values
                                combined_out_df["Time (min)"] = (tsec_series / 60.0).round(6).values
                                combined_out_df["Time (hour)"] = (tsec_series / 3600.0).round(6).values
                                baseline_assigned = True
                                self.log("Global time computed using first timestamp in combined data as baseline")
                        if not baseline_assigned:
                            self.log("No usable global timestamp or tick baseline found; Time columns left as NaN")

                    # Downsample if requested
                    if output_freq is not None and baseline_assigned:
                        try:
                            freq = float(output_freq)
                            if freq <= 0:
                                raise ValueError("Output frequency must be positive")
                            # use Time (sec) as numeric
                            time_sec_series = pd.to_numeric(combined_out_df["Time (sec)"], errors="coerce")
                            valid_mask = time_sec_series.notna()
                            if not valid_mask.any():
                                self.log("Cannot downsample: Time (sec) column contains no numeric values.")
                            else:
                                tmin = float(time_sec_series[valid_mask].iloc[0])
                                tmax = float(time_sec_series[valid_mask].iloc[-1])
                                dt = 1.0 / freq
                                # build target times from tmin to tmax inclusive
                                n_steps = int(np.floor((tmax - tmin) / dt)) + 1
                                targets = tmin + np.arange(n_steps) * dt
                                # ensure time_sec_series is numpy array
                                times = time_sec_series.values.astype(float)
                                # find nearest index for each target
                                # use searchsorted for speed (requires sorted times)
                                if not np.all(np.diff(times) >= 0):
                                    # not monotonic: fallback to brute force nearest
                                    chosen_idx = []
                                    for tt in targets:
                                        diffs = np.abs(times - tt)
                                        chosen_idx.append(int(np.nanargmin(diffs)))
                                else:
                                    idxs = np.searchsorted(times, targets, side="left")
                                    chosen_idx = []
                                    for k, tt in enumerate(targets):
                                        i = idxs[k]
                                        if i == 0:
                                            chosen_idx.append(0)
                                        elif i >= len(times):
                                            chosen_idx.append(len(times) - 1)
                                        else:
                                            # pick closer of i-1 and i
                                            if abs(times[i] - tt) < abs(times[i - 1] - tt):
                                                chosen_idx.append(i)
                                            else:
                                                chosen_idx.append(i - 1)
                                # keep unique indices in order
                                chosen_idx_unique = []
                                last = None
                                for i in chosen_idx:
                                    if i != last:
                                        chosen_idx_unique.append(i)
                                        last = i
                                ds_df = combined_out_df.iloc[chosen_idx_unique].reset_index(drop=True)
                                combined_out_df = ds_df
                                self.log(f"Downsampled combined CSV to {freq} Hz -> {len(combined_out_df)} rows")
                        except Exception as e:
                            self.log(f"❌ Error during downsampling: {e}")

                    # finalize and save
                    for col in template_cols:
                        if col not in combined_out_df.columns:
                            combined_out_df[col] = pd.NA
                    combined_out_df = combined_out_df[template_cols]
                    try:
                        combined_out_df.to_csv(outpath, index=False, columns=template_cols)
                        self.log(f"Saved combined CSV -> {outpath}")
                    except Exception as e:
                        self.log(f"❌ Error saving combined CSV: {e}")

            if self.stop_event.is_set():
                self.log("Conversion stopped by user."); self.queue.put(("done", "Conversion stopped by user."))
            else:
                self.log("Conversion completed."); self.queue.put(("done", f"Conversion completed. Output: {outpath}"))
        except Exception as exc:
            tb = traceback.format_exc(); self.log(f"Unexpected error: {exc}\n{tb}"); self.queue.put(("error", f"Unexpected error: {exc}"))
        finally:
            try: self.queue.put(("progress", len(files)))
            except Exception: pass


def main():
    app = ConverterGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
