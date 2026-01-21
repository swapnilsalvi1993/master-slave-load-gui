# -*- coding: utf-8 -*-
"""
Created on Tue Nov 25 23:49:00 2025

@author: ssalvi
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import os
import math

# ---------------------------
# Helpers for token handling
# ---------------------------
def token_display(t):
    if t["type"] == "col":
        return f"[COL] {t['value']}"
    elif t["type"] == "num":
        return f"[NUM] {t['value']}"
    elif t["type"] == "op":
        return f"[OP] {t['value']}"
    else:
        return str(t)

def is_operand(t):
    return t["type"] in ("col", "num")

# ---------------------------
# Main Application
# ---------------------------
class CSVExprApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CSV Expression Builder + Battery Metrics")
        self.geometry("1100x650")
        self.minsize(900, 600)

        self.df = None
        self.csv_path = None
        self.tokens = []  # expression token list: [{'type':'col'|'num'|'op', 'value':...}, ...]

        self._build_ui()

    def _build_ui(self):
        # Top toolbar
        toolbar = ttk.Frame(self)
        toolbar.pack(side=tk.TOP, fill=tk.X, padx=6, pady=6)

        ttk.Button(toolbar, text="Open CSV", command=self.open_csv).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="Save CSV As...", command=self.save_as).pack(side=tk.LEFT, padx=(6,0))
        ttk.Button(toolbar, text="Battery Metrics", command=self.open_battery_metrics_dialog).pack(side=tk.LEFT, padx=(12,0))

        # Main panes
        main_pane = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        # Left: Columns & reordering
        left_frame = ttk.Labelframe(main_pane, text="DataFrame Columns (double-click to add to expression)")
        left_frame.config(width=260)
        main_pane.add(left_frame, weight=0)

        self.cols_listbox = tk.Listbox(left_frame, exportselection=False)
        self.cols_listbox.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.cols_listbox.bind("<Double-Button-1>", lambda e: self.add_selected_col_to_expr())

        cols_btn_frame = ttk.Frame(left_frame)
        cols_btn_frame.pack(fill=tk.X, padx=6, pady=(0,6))
        ttk.Button(cols_btn_frame, text="Add to Expression", command=self.add_selected_col_to_expr).pack(side=tk.LEFT)
        ttk.Button(cols_btn_frame, text="Move Up", command=self.move_selected_column_up).pack(side=tk.LEFT, padx=4)
        ttk.Button(cols_btn_frame, text="Move Down", command=self.move_selected_column_down).pack(side=tk.LEFT)
        ttk.Button(cols_btn_frame, text="Apply Column Order →", command=self.apply_column_order).pack(side=tk.LEFT, padx=6)

        # Center: Expression builder
        center_frame = ttk.Labelframe(main_pane, text="Expression Builder")
        center_frame.config(width=360)
        main_pane.add(center_frame, weight=1)

        expr_label = ttk.Label(center_frame, text="Expression tokens (order matters):")
        expr_label.pack(anchor=tk.W, padx=6, pady=(6,0))

        self.expr_listbox = tk.Listbox(center_frame, exportselection=False)
        self.expr_listbox.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        expr_btn_frame = ttk.Frame(center_frame)
        expr_btn_frame.pack(fill=tk.X, padx=6, pady=(0,6))

        # Operators dropdown
        self.op_var = tk.StringVar(value="+")
        op_box = ttk.Combobox(expr_btn_frame, textvariable=self.op_var, values=["+", "-", "*", "/"], width=3, state="readonly")
        op_box.pack(side=tk.LEFT)
        ttk.Button(expr_btn_frame, text="Add Operator", command=self.add_operator).pack(side=tk.LEFT, padx=4)

        # Add number
        self.num_var = tk.StringVar(value="0")
        num_entry = ttk.Entry(expr_btn_frame, textvariable=self.num_var, width=10)
        num_entry.pack(side=tk.LEFT, padx=(10,0))
        ttk.Button(expr_btn_frame, text="Add Number", command=self.add_number).pack(side=tk.LEFT, padx=4)
        ttk.Button(expr_btn_frame, text="Edit Number", command=self.edit_number).pack(side=tk.LEFT, padx=4)

        # Move/remove
        ttk.Button(expr_btn_frame, text="Move Up", command=self.move_token_up).pack(side=tk.LEFT, padx=(12,4))
        ttk.Button(expr_btn_frame, text="Move Down", command=self.move_token_down).pack(side=tk.LEFT)
        ttk.Button(expr_btn_frame, text="Remove", command=self.remove_token).pack(side=tk.LEFT, padx=(8,0))

        # Right: Data preview & column order
        right_frame = ttk.Labelframe(main_pane, text="Preview / Column Order")
        right_frame.config(width=380)
        main_pane.add(right_frame, weight=0)

        preview_label = ttk.Label(right_frame, text="Data preview (first 200 rows):")
        preview_label.pack(anchor=tk.W, padx=6, pady=(6,0))

        # Treeview for preview
        self.preview = ttk.Treeview(right_frame, show="headings")
        self.preview.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.preview_scroll = ttk.Scrollbar(right_frame, orient="vertical", command=self.preview.yview)
        self.preview.configure(yscrollcommand=self.preview_scroll.set)
        self.preview_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Column order listbox
        order_frame = ttk.Frame(right_frame)
        order_frame.pack(fill=tk.X, padx=6, pady=(0,6))
        ttk.Label(order_frame, text="Column order (reorder and apply):").pack(anchor=tk.W)

        self.order_listbox = tk.Listbox(order_frame, exportselection=False)
        self.order_listbox.pack(fill=tk.X, pady=(4,4))
        order_btns = ttk.Frame(order_frame)
        order_btns.pack(fill=tk.X)
        ttk.Button(order_btns, text="Up", command=self.move_order_up).pack(side=tk.LEFT)
        ttk.Button(order_btns, text="Down", command=self.move_order_down).pack(side=tk.LEFT, padx=6)

        # Bottom controls: compute and new column name
        bottom = ttk.Frame(self)
        bottom.pack(side=tk.BOTTOM, fill=tk.X, padx=6, pady=6)

        ttk.Label(bottom, text="New column name:").pack(side=tk.LEFT)
        self.newcol_var = tk.StringVar()
        ttk.Entry(bottom, textvariable=self.newcol_var, width=30).pack(side=tk.LEFT, padx=(6,12))
        ttk.Button(bottom, text="Compute & Add Column", command=self.compute_and_add).pack(side=tk.LEFT)
        ttk.Button(bottom, text="Preview Expression (show as formula)", command=self.show_formula).pack(side=tk.LEFT, padx=6)

        # Status bar
        self.status_var = tk.StringVar(value="No file loaded.")
        status = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status.pack(side=tk.BOTTOM, fill=tk.X)

    # -----------------------
    # File operations
    # -----------------------
    def open_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
            df = pd.read_csv(path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read CSV:\n{e}")
            return
        self.df = df
        self.csv_path = path
        self.tokens.clear()
        self.update_ui_after_load()
        self.status_var.set(f"Loaded: {os.path.basename(path)} — {len(df)} rows, {len(df.columns)} cols")

    def save_as(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a CSV first.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files","*.csv"), ("All files","*.*")])
        if not path:
            return
        try:
            self.df.to_csv(path, index=False)
            messagebox.showinfo("Saved", f"Saved to: {path}")
            self.status_var.set(f"Saved: {path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save CSV:\n{e}")

    # -----------------------
    # UI update methods
    # -----------------------
    def update_ui_after_load(self):
        # columns listbox
        self.cols_listbox.delete(0, tk.END)
        for c in list(self.df.columns):
            self.cols_listbox.insert(tk.END, c)
        # order listbox
        self.order_listbox.delete(0, tk.END)
        for c in list(self.df.columns):
            self.order_listbox.insert(tk.END, c)
        # expression listbox
        self.refresh_expr_listbox()
        # preview
        self.refresh_preview()

    def refresh_expr_listbox(self):
        self.expr_listbox.delete(0, tk.END)
        for t in self.tokens:
            self.expr_listbox.insert(tk.END, token_display(t))

    def refresh_preview(self):
        # clear treeview
        for col in self.preview["columns"]:
            self.preview.heading(col, text="")
        self.preview.delete(*self.preview.get_children())
        self.preview["columns"] = []
        if self.df is None:
            return
        max_cols = min(12, len(self.df.columns))
        cols = list(self.df.columns)[:max_cols]
        self.preview["columns"] = cols
        for c in cols:
            self.preview.heading(c, text=c)
            self.preview.column(c, width=100, anchor=tk.CENTER)

        for i, row in enumerate(self.df.head(200).itertuples(index=False, name=None)):
            display_row = [self._short_str(v) for v in row[:max_cols]]
            self.preview.insert("", tk.END, values=display_row)

    def _short_str(self, v, maxlen=60):
        s = str(v)
        return (s[:maxlen-3] + "...") if len(s) > maxlen else s

    # -----------------------
    # Column order controls
    # -----------------------
    def move_selected_column_up(self):
        sel = self.cols_listbox.curselection()
        if not sel: return
        i = sel[0]
        if i == 0: return
        val = self.cols_listbox.get(i)
        self.cols_listbox.delete(i)
        self.cols_listbox.insert(i-1, val)
        self.cols_listbox.select_set(i-1)
        # keep order list synced
        self.sync_order_from_cols()

    def move_selected_column_down(self):
        sel = self.cols_listbox.curselection()
        if not sel: return
        i = sel[0]
        if i >= self.cols_listbox.size()-1: return
        val = self.cols_listbox.get(i)
        self.cols_listbox.delete(i)
        self.cols_listbox.insert(i+1, val)
        self.cols_listbox.select_set(i+1)
        self.sync_order_from_cols()

    def sync_order_from_cols(self):
        # keep order_listbox equal to cols_listbox
        self.order_listbox.delete(0, tk.END)
        for i in range(self.cols_listbox.size()):
            self.order_listbox.insert(tk.END, self.cols_listbox.get(i))

    def move_order_up(self):
        sel = self.order_listbox.curselection()
        if not sel: return
        i = sel[0]
        if i == 0: return
        val = self.order_listbox.get(i)
        self.order_listbox.delete(i)
        self.order_listbox.insert(i-1, val)
        self.order_listbox.select_set(i-1)

    def move_order_down(self):
        sel = self.order_listbox.curselection()
        if not sel: return
        i = sel[0]
        if i >= self.order_listbox.size()-1: return
        val = self.order_listbox.get(i)
        self.order_listbox.delete(i)
        self.order_listbox.insert(i+1, val)
        self.order_listbox.select_set(i+1)

    def apply_column_order(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a CSV first.")
            return
        new_order = [self.order_listbox.get(i) for i in range(self.order_listbox.size())]
        # verify same set of columns
        if set(new_order) != set(self.df.columns):
            messagebox.showerror("Invalid order", "The column order must contain the same columns as the DataFrame.")
            return
        self.df = self.df[new_order]
        # update cols_listbox to match
        self.cols_listbox.delete(0, tk.END)
        for c in new_order:
            self.cols_listbox.insert(tk.END, c)
        self.refresh_preview()
        self.status_var.set("Applied new column order.")

    # -----------------------
    # Expression builder controls
    # -----------------------
    def add_selected_col_to_expr(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a CSV first.")
            return
        sel = self.cols_listbox.curselection()
        if not sel:
            messagebox.showinfo("Select a column", "Select a column to add to the expression.")
            return
        col = self.cols_listbox.get(sel[0])
        self.tokens.append({"type":"col", "value":col})
        self.refresh_expr_listbox()

    def add_operator(self):
        op = self.op_var.get().strip()
        if op not in ("+", "-", "*", "/"):
            messagebox.showerror("Invalid operator", "Operator must be one of +, -, *, /.")
            return
        self.tokens.append({"type":"op", "value":op})
        self.refresh_expr_listbox()

    def add_number(self):
        v = self.num_var.get().strip()
        try:
            num = float(v)
        except Exception:
            messagebox.showerror("Invalid number", "Enter a valid numeric value.")
            return
        # keep as simple string but store numeric
        self.tokens.append({"type":"num", "value":num})
        self.refresh_expr_listbox()

    def edit_number(self):
        sel = self.expr_listbox.curselection()
        if not sel:
            messagebox.showinfo("Select token", "Select a numeric token in the expression to edit.")
            return
        idx = sel[0]
        t = self.tokens[idx]
        if t["type"] != "num":
            messagebox.showerror("Wrong token", "Selected token is not a numeric value.")
            return
        nv = simpledialog.askstring("Edit number", "Enter new numeric value:", initialvalue=str(t["value"]))
        if nv is None:
            return
        try:
            num = float(nv)
        except Exception:
            messagebox.showerror("Invalid number", "Enter a valid numeric value.")
            return
        self.tokens[idx]["value"] = num
        self.refresh_expr_listbox()

    def move_token_up(self):
        sel = self.expr_listbox.curselection()
        if not sel: return
        i = sel[0]
        if i == 0: return
        t = self.tokens.pop(i)
        self.tokens.insert(i-1, t)
        self.refresh_expr_listbox()
        self.expr_listbox.select_set(i-1)

    def move_token_down(self):
        sel = self.expr_listbox.curselection()
        if not sel: return
        i = sel[0]
        if i >= len(self.tokens)-1: return
        t = self.tokens.pop(i)
        self.tokens.insert(i+1, t)
        self.refresh_expr_listbox()
        self.expr_listbox.select_set(i+1)

    def remove_token(self):
        sel = self.expr_listbox.curselection()
        if not sel: return
        i = sel[0]
        self.tokens.pop(i)
        self.refresh_expr_listbox()

    def show_formula(self):
        if not self.tokens:
            messagebox.showinfo("Expression", "No tokens in expression.")
            return
        s = " ".join(token_display(t) for t in self.tokens)
        messagebox.showinfo("Expression Tokens", s)

    # -----------------------
    # Evaluation & adding column
    # -----------------------
    def compute_and_add(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a CSV first.")
            return
        if not self.tokens:
            messagebox.showinfo("No expression", "Build an expression first.")
            return
        # validate tokens pattern: operand (op operand)*
        if len(self.tokens) % 2 == 0:
            messagebox.showerror("Invalid expression", "Expression must be in form: operand operator operand ...")
            return
        for i, t in enumerate(self.tokens):
            if i % 2 == 0:
                if not is_operand(t):
                    messagebox.showerror("Invalid expression", "Expected a column or number at position %d." % (i+1))
                    return
            else:
                if t["type"] != "op":
                    messagebox.showerror("Invalid expression", "Expected an operator at position %d." % (i+1))
                    return
        # Get new column name
        new_col = self.newcol_var.get().strip()
        if not new_col:
            new_col = simpledialog.askstring("New column name", "Enter new column name:")
            if not new_col:
                messagebox.showinfo("Needs name", "Operation cancelled: a column name is required.")
                return
            self.newcol_var.set(new_col)

        try:
            result = self._evaluate_tokens()
        except Exception as e:
            messagebox.showerror("Evaluation error", f"Could not evaluate expression:\n{e}")
            return

        # place result into df
        self.df[new_col] = result
        # update lists and preview
        self.cols_listbox.insert(tk.END, new_col)
        self.order_listbox.insert(tk.END, new_col)
        self.refresh_preview()
        self.status_var.set(f"Added column '{new_col}'.")
        messagebox.showinfo("Success", f"New column '{new_col}' added to DataFrame.")

    def _evaluate_tokens(self):
        # Evaluate sequentially: ((a op b) op c) ...
        # token_to_series will convert num -> Series of constant
        def token_to_series(t):
            if t["type"] == "col":
                if t["value"] not in self.df.columns:
                    raise ValueError(f"Column '{t['value']}' not found in DataFrame.")
                return self.df[t["value"]]
            elif t["type"] == "num":
                return pd.Series([float(t["value"])] * len(self.df), index=self.df.index)
            else:
                raise ValueError("Unexpected token type when converting to series: " + str(t))

        # Start with first operand
        left = token_to_series(self.tokens[0])
        for i in range(1, len(self.tokens), 2):
            op = self.tokens[i]["value"]
            right = token_to_series(self.tokens[i+1])

            if op == "+":
                left = left + right
            elif op == "-":
                left = left - right
            elif op == "*":
                left = left * right
            elif op == "/":
                left = left / right
            else:
                raise ValueError("Unsupported operator: " + str(op))
        return left

    # -----------------------
    # Battery metrics dialog & computation
    # -----------------------
    def open_battery_metrics_dialog(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a CSV first.")
            return
        dlg = BatteryMetricsDialog(self, list(self.df.columns))
        self.wait_window(dlg.top)
        if dlg.result is None:
            return  # cancelled
        # dlg.result is a dict with keys: current_col, voltage_col, time_col, time_unit, sort_by_time, names, overwrite
        try:
            self._compute_battery_metrics(
                current_col=dlg.result["current_col"],
                voltage_col=dlg.result["voltage_col"],
                time_col=dlg.result["time_col"],
                time_unit=dlg.result["time_unit"],
                sort_by_time=dlg.result["sort_by_time"],
                power_name=dlg.result["power_name"],
                cap_name=dlg.result["cap_name"],
                energy_name=dlg.result["energy_name"],
                overwrite_csv=dlg.result["overwrite"]
            )
        except Exception as e:
            messagebox.showerror("Computation error", f"Could not compute battery metrics:\n{e}")

    def _compute_battery_metrics(self, current_col, voltage_col, time_col,
                                 time_unit="seconds", sort_by_time=False,
                                 power_name="Power_W", cap_name="CumCapacity_Ah",
                                 energy_name="CumEnergy_Wh", overwrite_csv=False):
        # Validate presence
        for col in (current_col, voltage_col, time_col):
            if col not in self.df.columns:
                raise ValueError(f"Column '{col}' not found in DataFrame.")

        # Work on a copy or view depending on sort option
        df = self.df

        if sort_by_time:
            try:
                df = df.sort_values(by=time_col).reset_index(drop=True)
            except Exception as e:
                raise ValueError(f"Could not sort by time column '{time_col}': {e}")

        # Convert to numeric
        current = pd.to_numeric(df[current_col], errors='coerce')
        voltage = pd.to_numeric(df[voltage_col], errors='coerce')
        time_series = pd.to_numeric(df[time_col], errors='coerce')

        if current.isna().any() or voltage.isna().any() or time_series.isna().any():
            # warn but continue; user can inspect NaNs
            messagebox.showwarning("NaNs present", "Some values in current/voltage/time columns could not be converted to numeric (they become NaN). Computation will proceed and produce NaN where input is invalid.")

        # Determine dt in seconds
        dt = time_series.diff().fillna(0).astype(float)
        if time_unit == "minutes":
            dt_seconds = dt * 60.0
        elif time_unit == "hours":
            dt_seconds = dt * 3600.0
        else:  # seconds
            dt_seconds = dt

        # Compute power
        power = voltage * current  # W if V and A

        # Capacity increment in Ah: current (A) * dt_seconds / 3600
        cap_inc = current * dt_seconds / 3600.0
        cum_cap = cap_inc.cumsum()

        # Energy increment in Wh: power (W) * dt_seconds / 3600
        energy_inc = power * dt_seconds / 3600.0
        cum_energy = energy_inc.cumsum()

        # Insert into DataFrame (matching original df index if sorted)
        # If we sorted, df is a sorted copy; we will map results back to original order
        if sort_by_time:
            # create series aligned to sorted df index; then reindex to original df index using time/order mapping
            # Simpler: attach to sorted df and then reindex by original index if original had unique index.
            # We will assign by position: assume user wants per-row result in current order (sorted) OR overwrite original order?
            # To preserve original row order, we'll create Series with same index as sorted df and then map to original index via index positions.
            # But since we reset_index after sorting, the original index is lost. To keep it simple and intuitive, we will add columns to the sorted df
            # and then re-merge into the original by matching a unique row ID.
            df_with_id = df.copy()
            df_with_id["_row_id_batt_metrics"] = range(len(df_with_id))
            df_with_id[power_name] = power
            df_with_id[cap_name] = cum_cap
            df_with_id[energy_name] = cum_energy
            # Now construct mapping to original by time and nearest match? This can be ambiguous if times repeat.
            # Instead, we'll warn user that when sorting is enabled, the DataFrame will be replaced with the sorted version (stable assignment).
            # Replace self.df with sorted df_with_id (dropping the helper id)
            df_with_id = df_with_id.drop(columns=["_row_id_batt_metrics"])
            self.df = df_with_id
        else:
            # assign directly
            self.df[power_name] = power.values
            self.df[cap_name] = cum_cap.values
            self.df[energy_name] = cum_energy.values

        # Refresh UI (columns lists and preview)
        self.update_ui_after_load()

        # Optionally overwrite original CSV
        if overwrite_csv and self.csv_path:
            try:
                self.df.to_csv(self.csv_path, index=False)
                self.status_var.set(f"Added power/capacity/energy and overwritten: {os.path.basename(self.csv_path)}")
                messagebox.showinfo("Saved", f"Computed metrics added and file overwritten:\n{self.csv_path}")
            except Exception as e:
                messagebox.showerror("Save error", f"Computed metrics were added to the DataFrame but could not overwrite the CSV:\n{e}")
        else:
            self.status_var.set("Computed battery metrics (not saved).")
            messagebox.showinfo("Done", f"Computed metrics added as columns: {power_name}, {cap_name}, {energy_name}.\nSave the file using 'Save CSV As...' or overwrite when asked next time.")

# ---------------------------
# Battery metrics selection dialog
# ---------------------------
class BatteryMetricsDialog:
    def __init__(self, parent, columns):
        self.parent = parent
        self.columns = columns
        self.result = None
        top = self.top = tk.Toplevel(parent)
        top.title("Battery Metrics — select columns & options")
        top.grab_set()
        top.resizable(False, False)

        row = 0
        ttk.Label(top, text="Current column (A):").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.current_cb = ttk.Combobox(top, values=self.columns, state="readonly", width=36)
        self.current_cb.grid(row=row, column=1, padx=8, pady=6)
        self.current_cb.set(self._guess_column(["current", "i", "amp", "amps", "current_a"]))

        row += 1
        ttk.Label(top, text="Voltage column (V):").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.voltage_cb = ttk.Combobox(top, values=self.columns, state="readonly", width=36)
        self.voltage_cb.grid(row=row, column=1, padx=8, pady=6)
        self.voltage_cb.set(self._guess_column(["voltage", "v", "volt", "volts", "voltage_v"]))

        row += 1
        ttk.Label(top, text="Time column:").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.time_cb = ttk.Combobox(top, values=self.columns, state="readonly", width=36)
        self.time_cb.grid(row=row, column=1, padx=8, pady=6)
        self.time_cb.set(self._guess_column(["time", "t", "timestamp", "seconds", "sec"]))

        row += 1
        ttk.Label(top, text="Time units:").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.time_unit_var = tk.StringVar(value="seconds")
        time_unit_cb = ttk.Combobox(top, textvariable=self.time_unit_var, values=["seconds", "minutes", "hours"], state="readonly", width=12)
        time_unit_cb.grid(row=row, column=1, sticky="w", padx=8, pady=6)

        row += 1
        self.sort_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="Sort rows by time before computing", variable=self.sort_var).grid(row=row, column=0, columnspan=2, sticky="w", padx=8, pady=6)

        row += 1
        ttk.Label(top, text="Power column name:").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.power_name_var = tk.StringVar(value="Power_W")
        ttk.Entry(top, textvariable=self.power_name_var, width=38).grid(row=row, column=1, padx=8, pady=6)

        row += 1
        ttk.Label(top, text="Cumulative capacity name:").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.cap_name_var = tk.StringVar(value="CumCapacity_Ah")
        ttk.Entry(top, textvariable=self.cap_name_var, width=38).grid(row=row, column=1, padx=8, pady=6)

        row += 1
        ttk.Label(top, text="Cumulative energy name:").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.energy_name_var = tk.StringVar(value="CumEnergy_Wh")
        ttk.Entry(top, textvariable=self.energy_name_var, width=38).grid(row=row, column=1, padx=8, pady=6)

        row += 1
        self.overwrite_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="Overwrite the loaded CSV file with new columns (if loaded)", variable=self.overwrite_var).grid(row=row, column=0, columnspan=2, sticky="w", padx=8, pady=6)

        row += 1
        btn_frame = ttk.Frame(top)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=(6,10))
        ttk.Button(btn_frame, text="Compute", command=self.on_ok).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_frame, text="Cancel", command=self.on_cancel).pack(side=tk.LEFT, padx=8)

    def _guess_column(self, candidates):
        low = [c.lower() for c in self.columns]
        for cand in candidates:
            if cand in low:
                # return the original casing column name
                return self.columns[low.index(cand)]
        # fallback to first column if present
        return self.columns[0] if self.columns else ""

    def on_ok(self):
        cur = self.current_cb.get().strip()
        volt = self.voltage_cb.get().strip()
        tcol = self.time_cb.get().strip()
        if not cur or not volt or not tcol:
            messagebox.showerror("Missing", "Please select current, voltage and time columns.")
            return
        # collect result
        self.result = {
            "current_col": cur,
            "voltage_col": volt,
            "time_col": tcol,
            "time_unit": self.time_unit_var.get(),
            "sort_by_time": bool(self.sort_var.get()),
            "power_name": self.power_name_var.get().strip() or "Power_W",
            "cap_name": self.cap_name_var.get().strip() or "CumCapacity_Ah",
            "energy_name": self.energy_name_var.get().strip() or "CumEnergy_Wh",
            "overwrite": bool(self.overwrite_var.get())
        }
        self.top.destroy()

    def on_cancel(self):
        self.result = None
        self.top.destroy()

# ---------------------------
# Run the app
# ---------------------------
if __name__ == "__main__":
    app = CSVExprApp()
    app.mainloop()
