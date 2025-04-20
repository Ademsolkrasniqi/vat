# Full corrected code for VatEntryApp

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap.tooltip import ToolTip
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
import json
import os
import datetime
from collections import defaultdict
import calendar
import traceback # Import here for use in except blocks

# --- Constants ---
DATA_FILE = 'vat_entry_data.json'
DEFAULT_VAT_RATE = Decimal('0.18') # Example: 18%
APP_TITLE = "VAT Entry System"
CURRENCY_SYMBOL = "€"
COMPANIES = ["SOL", "HELO"]
TRANSACTION_TYPES = ["Sales", "Imports", "Local"]
# Generate month names based on current locale, fallback to English if needed
try:
    # Attempt locale-aware month names (might require locale setup on system)
    import locale
    # Set locale ideally based on system, or specify e.g., locale.LC_TIME, 'en_US.UTF-8'
    # For simplicity, we'll use default locale setting here.
    # locale.setlocale(locale.LC_TIME, '') # Use system's default locale setting
    MONTH_NAMES = [calendar.month_name[i] for i in range(1, 13)]
    if not MONTH_NAMES[0]: # If first month name is empty (locale issue), fallback
        raise ImportError
except ImportError:
    # Fallback to English month names
    MONTH_NAMES = ["", "January", "February", "March", "April", "May", "June",
                   "July", "August", "September", "October", "November", "December"][1:]


# --- Helper Functions ---
def quantize_decimal(d):
    """Rounds a Decimal to 2 decimal places for currency."""
    if d is None: return Decimal(0)
    try:
        return Decimal(d).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (TypeError, InvalidOperation, ValueError):
        return Decimal(0)

def format_curr(amount):
    """Formats a Decimal amount as currency."""
    try:
        # Basic formatting, adjust locale/pattern if needed for specific conventions
        return f"{CURRENCY_SYMBOL}{quantize_decimal(amount):,.2f}"
    except Exception:
        return f"{CURRENCY_SYMBOL} Error"

def get_month_year_str(year, month_name):
    """Returns 'YYYY-MM' string"""
    try:
        # Find month number (case-insensitive matching might be good)
        month_num = MONTH_NAMES.index(month_name) + 1
        return f"{int(year)}-{month_num:02d}"
    except (ValueError, TypeError):
         # Fallback if month name invalid or year not int
        today = datetime.date.today()
        return f"{today.year}-{today.month:02d}"


def get_prev_month_year_str(year, month_name):
    """Gets the previous month's YYYY-MM string"""
    try:
        month_num = MONTH_NAMES.index(month_name) + 1
        current_date = datetime.date(int(year), month_num, 1)
        # Go back one day to get into the previous month
        prev_month_date = current_date - datetime.timedelta(days=1)
        return prev_month_date.strftime('%Y-%m')
    except (ValueError, TypeError):
         # Fallback if month name invalid or year not int
        today = datetime.date.today()
        prev_month = today.replace(day=1) - datetime.timedelta(days=1)
        return prev_month.strftime('%Y-%m')

# --- Main Application Class ---
class VatEntryApp:
    def __init__(self, root):
        self.root = root
        self.style = tb.Style()
        self.root.title(APP_TITLE)
        # self.root.geometry("1100x700") # Adjust as needed

        self.vat_rate = DEFAULT_VAT_RATE
        self._calculating = False
        self.editing_id = None # Store ID of entry being edited

        # --- Data Store ---
        self.transactions = []
        self.carry_forward_data = defaultdict(lambda: defaultdict(Decimal)) # {company: {YYYY-MM: surplus}}
        self.known_counterparties = set()

        # --- Input Form Variables ---
        current_year = datetime.date.today().year
        # Handle case where current month name might not be in generated list if locale failed
        current_month_index = datetime.date.today().month -1
        current_month_name = MONTH_NAMES[current_month_index] if current_month_index < len(MONTH_NAMES) else MONTH_NAMES[0]

        self.company_var = tk.StringVar(value=COMPANIES[0])
        self.year_var = tk.IntVar(value=current_year)
        self.month_var = tk.StringVar(value=current_month_name)
        self.transaction_type_var = tk.StringVar(value=TRANSACTION_TYPES[0])
        self.invoice_var = tk.StringVar()
        self.counterparty_var = tk.StringVar()
        self.v_me_tvsh_var = tk.StringVar() # Total Amount (w VAT)
        self.v_pa_tvsh_var = tk.StringVar() # Base Amount (w/o VAT)
        self.tvsh_var = tk.StringVar()      # VAT Amount

        # --- Connect Tracers ---
        self.v_pa_tvsh_var.trace_add("write", self._calculate_from_base)
        self.tvsh_var.trace_add("write", self._calculate_from_vat)
        self.v_me_tvsh_var.trace_add("write", self._calculate_from_total)

        # --- Listeners for Live Data Update ---
        self.company_var.trace_add("write", self._update_live_data_display)
        self.year_var.trace_add("write", self._update_live_data_display)
        self.month_var.trace_add("write", self._update_live_data_display)

        # --- Load Initial Data ---
        # Load data *before* creating widgets that depend on it (like counterparty list)
        self._load_data()

        # --- UI Structure ---
        # Top frame for input and live data
        top_frame = tb.Frame(root, padding=10)
        top_frame.pack(fill=X, padx=5, pady=5)

        input_frame = self._create_input_form(top_frame)
        input_frame.pack(side=LEFT, fill=Y, padx=(0, 10))

        live_data_frame = self._create_live_data_display(top_frame)
        live_data_frame.pack(side=LEFT, fill=BOTH, expand=True)

        # Bottom frame for the table
        table_frame = tb.LabelFrame(root, text="Table of Entries", padding=10)
        table_frame.pack(fill=BOTH, expand=True, padx=10, pady=(0, 10))
        self._create_entry_table(table_frame)

        # --- Initialize Displays ---
        self._populate_treeview()
        self._update_live_data_display() # Initial calculation based on loaded data/defaults

        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

        # --- Setup Keyboard Navigation Order ---
        self._setup_keyboard_navigation()


    # --- UI Creation Methods ---
    def _create_input_form(self, parent):
        form = tb.LabelFrame(parent, text="Entry", bootstyle="secondary", padding=10)

        # Labels and Entries using grid
        row_counter = 0
        tb.Label(form, text="Company:").grid(row=row_counter, column=0, padx=5, pady=3, sticky=W)
        self.company_combo = tb.Combobox(form, textvariable=self.company_var, values=COMPANIES, state="readonly", width=10)
        self.company_combo.grid(row=row_counter, column=1, padx=5, pady=3, sticky=W+E)
        row_counter += 1

        tb.Label(form, text="Month/Year:").grid(row=row_counter, column=0, padx=5, pady=3, sticky=W)
        month_year_frame = tb.Frame(form) # Frame to hold month and year together
        month_year_frame.grid(row=row_counter, column=1, padx=5, pady=3, sticky=W+E)
        self.month_combo = tb.Combobox(month_year_frame, textvariable=self.month_var, values=MONTH_NAMES, state="readonly", width=12)
        self.month_combo.pack(side=LEFT, fill=X, expand=True)
        self.year_spin = tb.Spinbox(month_year_frame, from_=2020, to=datetime.date.today().year + 5, textvariable=self.year_var, width=6)
        self.year_spin.pack(side=LEFT, padx=(5,0))
        row_counter += 1

        tb.Label(form, text="Transaction:").grid(row=row_counter, column=0, padx=5, pady=3, sticky=W)
        self.trans_combo = tb.Combobox(form, textvariable=self.transaction_type_var, values=TRANSACTION_TYPES, state="readonly", width=10)
        self.trans_combo.grid(row=row_counter, column=1, padx=5, pady=3, sticky=W+E)
        row_counter += 1

        tb.Label(form, text="Invoice Number:").grid(row=row_counter, column=0, padx=5, pady=3, sticky=W)
        self.invoice_entry = tb.Entry(form, textvariable=self.invoice_var)
        self.invoice_entry.grid(row=row_counter, column=1, padx=5, pady=3, sticky=W+E)
        row_counter += 1

        tb.Label(form, text="Counterparty:").grid(row=row_counter, column=0, padx=5, pady=3, sticky=W)
        # Initialize with loaded counterparties
        self.counterparty_combo = tb.Combobox(form, textvariable=self.counterparty_var, values=sorted(list(self.known_counterparties)), width=25)
        self.counterparty_combo.grid(row=row_counter, column=1, padx=5, pady=3, sticky=W+E)
        self.counterparty_combo.bind('<KeyRelease>', self._filter_counterparties) # Update list on typing
        row_counter += 1

        # Amounts
        tb.Label(form, text="v. me tvsh (Total):").grid(row=row_counter, column=0, padx=5, pady=3, sticky=W)
        self.v_me_tvsh_entry = tb.Entry(form, textvariable=self.v_me_tvsh_var, justify=RIGHT)
        self.v_me_tvsh_entry.grid(row=row_counter, column=1, padx=5, pady=3, sticky=W+E)
        ToolTip(self.v_me_tvsh_entry, "Enter amount INCLUDING VAT")
        row_counter += 1

        tb.Label(form, text="v. pa tvsh (Base):").grid(row=row_counter, column=0, padx=5, pady=3, sticky=W)
        self.v_pa_tvsh_entry = tb.Entry(form, textvariable=self.v_pa_tvsh_var, justify=RIGHT)
        self.v_pa_tvsh_entry.grid(row=row_counter, column=1, padx=5, pady=3, sticky=W+E)
        ToolTip(self.v_pa_tvsh_entry, "Enter amount EXCLUDING VAT")
        row_counter += 1

        tb.Label(form, text="tvsh (VAT Amount):").grid(row=row_counter, column=0, padx=5, pady=3, sticky=W)
        self.tvsh_entry = tb.Entry(form, textvariable=self.tvsh_var, justify=RIGHT)
        self.tvsh_entry.grid(row=row_counter, column=1, padx=5, pady=3, sticky=W+E)
        ToolTip(self.tvsh_entry, "Enter VAT amount directly")
        row_counter += 1

        # Done/Update Button
        self.done_button = tb.Button(form, text="Done (Add Entry)", command=self._add_or_update_entry, bootstyle=SUCCESS)
        self.done_button.grid(row=row_counter, column=0, columnspan=2, pady=10)

        # Cancel Edit Button (Initially hidden)
        self.cancel_button = tb.Button(form, text="Cancel Edit", command=self._cancel_edit, bootstyle=WARNING)
        # self.cancel_button.grid(...) # Grid it when needed

        return form

    def _create_live_data_display(self, parent):
        live_frame = tb.LabelFrame(parent, text="LIVE DATA", bootstyle="info", padding=10)

        # Configure grid
        live_frame.columnconfigure(1, weight=1) # Allow amount columns to expand a bit
        live_frame.columnconfigure(2, weight=1)
        live_frame.columnconfigure(3, weight=1)
        live_frame.columnconfigure(5, weight=1) # Balance col

        # Header Row
        headers = ["Company", "v. me tvsh", "v. pa tvsh", "tvsh", "Prev. Surplus", "Balance (Due/Surplus)"]
        for i, header in enumerate(headers):
            anchor = E if i > 0 else W
            style = SUCCESS if i > 0 else DEFAULT # Make headers stand out slightly
            lbl = tb.Label(live_frame, text=header, anchor=anchor, bootstyle=style, font="-weight bold")
            lbl.grid(row=0, column=i, padx=5, pady=(0, 5), sticky=EW)

        # Data Rows (Store labels in a dictionary for easy access)
        self.live_labels = {} # {company: {type/total: {col_name: label_widget}}}

        row_idx = 1
        for company in COMPANIES:
            tb.Label(live_frame, text=company, font="-weight bold").grid(row=row_idx, column=0, padx=5, pady=2, sticky=W)
            self.live_labels[company] = {}

            # Transaction Type rows
            for trans_type in TRANSACTION_TYPES:
                self.live_labels[company][trans_type] = {}
                tb.Label(live_frame, text=trans_type).grid(row=row_idx, column=0, padx=(20, 5), pady=1, sticky=W) # Indent type
                for i, col_name in enumerate(["v_me_tvsh", "v_pa_tvsh", "tvsh"]):
                    lbl = tb.Label(live_frame, text=format_curr(0), anchor=E)
                    lbl.grid(row=row_idx, column=i + 1, padx=5, pady=1, sticky=EW)
                    self.live_labels[company][trans_type][col_name] = lbl
                row_idx += 1

            # Summary rows (Previous Surplus and Balance)
            self.live_labels[company]['summary'] = {}
            summary_start_row = row_idx - len(TRANSACTION_TYPES) # Align with the first trans type row
            summary_rowspan = len(TRANSACTION_TYPES) # Span across the trans type rows

            surplus_lbl = tb.Label(live_frame, text=format_curr(0), anchor=E)
            surplus_lbl.grid(row=summary_start_row, column=4, padx=5, pady=1, sticky="nsew", rowspan=summary_rowspan) # Use nsew for vertical span
            self.live_labels[company]['summary']['prev_surplus'] = surplus_lbl

            balance_lbl = tb.Label(live_frame, text=format_curr(0), anchor=E, font="-weight bold")
            balance_lbl.grid(row=summary_start_row, column=5, padx=5, pady=1, sticky="nsew", rowspan=summary_rowspan) # Use nsew
            self.live_labels[company]['summary']['balance'] = balance_lbl

            # Add separator below each company block
            tb.Separator(live_frame, orient=HORIZONTAL).grid(row=row_idx, column=0, columnspan=6, pady=(5,10), sticky=EW)
            row_idx += 1 # Increment after the separator for the next company

        return live_frame

    def _create_entry_table(self, parent):
        table_controls_frame = tb.Frame(parent)
        table_controls_frame.pack(fill=X, pady=(0, 5))

        # Add Edit/Delete buttons here, associated with Treeview selection
        edit_btn = tb.Button(table_controls_frame, text="Edit", command=self._edit_entry, bootstyle="warning-outline", width=8)
        edit_btn.pack(side=RIGHT, padx=5)
        delete_btn = tb.Button(table_controls_frame, text="Delete", command=self._delete_selected_entry, bootstyle="danger-outline", width=8)
        delete_btn.pack(side=RIGHT, padx=5)


        cols = ("month_year", "company", "invoice_no", "counterparty", "v_pa_tvsh", "tvsh", "v_me_tvsh") # Reordered to match excel
        col_names = ("Month", "Company", "Invoice No.", "Counterparty", f"Base {CURRENCY_SYMBOL}", f"VAT {CURRENCY_SYMBOL}", f"Total {CURRENCY_SYMBOL}")
        col_widths = (70, 60, 100, 180, 100, 100, 100)
        col_anchors = (W, W, W, W, E, E, E)

        scrollbar_y = tb.Scrollbar(parent, orient=VERTICAL)
        scrollbar_x = tb.Scrollbar(parent, orient=HORIZONTAL)

        self.tree = tb.Treeview(
            parent, columns=cols, show='headings', bootstyle="primary",
            yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set, height=10
        )
        scrollbar_y.config(command=self.tree.yview)
        scrollbar_x.config(command=self.tree.xview)
        scrollbar_y.pack(side=RIGHT, fill=Y)
        scrollbar_x.pack(side=BOTTOM, fill=X)

        for i, col in enumerate(cols):
            self.tree.heading(col, text=col_names[i], anchor=col_anchors[i])
            self.tree.column(col, width=col_widths[i], anchor=col_anchors[i], stretch=(col not in ['month_year', 'company']))

        self.tree.pack(fill=BOTH, expand=True)
        self.tree.bind("<Double-1>", self._edit_entry) # Edit on double-click

    # --- Keyboard Navigation ---
    def _setup_keyboard_navigation(self):
        self.input_widgets_order = [
            self.company_combo,
            self.month_combo,
            self.year_spin,
            self.trans_combo,
            self.invoice_entry,
            self.counterparty_combo,
            self.v_me_tvsh_entry,
            self.v_pa_tvsh_entry,
            self.tvsh_entry,
            self.done_button
        ]
        for i, widget in enumerate(self.input_widgets_order):
            # Use lambda with default argument to capture current value of i
            widget.bind("<Return>", lambda event, next_widget_idx=(i + 1): self._focus_next_widget(next_widget_idx))
            widget.bind("<KP_Enter>", lambda event, next_widget_idx=(i + 1): self._focus_next_widget(next_widget_idx)) # Numpad Enter

    def _focus_next_widget(self, next_widget_idx):
        if next_widget_idx < len(self.input_widgets_order):
            next_widget = self.input_widgets_order[next_widget_idx]
            next_widget.focus_set()
            # Select text in Entry/Spinbox for easy overwriting
            if isinstance(next_widget, (ttk.Entry, ttk.Spinbox)):
                 next_widget.select_range(0, tk.END)
            elif isinstance(next_widget, ttk.Combobox):
                 # Optional: Open dropdown? Or just focus? Just focus is simpler.
                 pass
        else:
             # If it's the last widget (Done button), simulate click
             self.done_button.invoke()
             # Move focus back to invoice_entry after adding/updating
             if hasattr(self, 'invoice_entry'): # Check if exists
                 self.invoice_entry.focus_set()
                 self.invoice_entry.select_range(0, tk.END)

        return "break" # Prevent default Return key behavior (like adding newline in text widgets)

    # --- Calculation Logic ---
    def _safely_get_decimal(self, var):
        """Safely converts a tk Variable's value to Decimal, returns None on failure."""
        try:
            value_str = var.get()
            return Decimal(value_str) if value_str else None # Handle empty string
        except (InvalidOperation, ValueError, TypeError):
            return None # Return None if conversion fails

    def _set_calculation_flag(self, state):
        """Sets the calculation flag to prevent infinite recursion."""
        self._calculating = state

    def _calculate_from_base(self, *args):
        """Calculates VAT and Total when Base amount is entered."""
        if self._calculating or not hasattr(self, 'v_pa_tvsh_entry'): return
        self._set_calculation_flag(True)
        base = self._safely_get_decimal(self.v_pa_tvsh_var)
        if base is not None and base >= 0:
            vat = quantize_decimal(base * self.vat_rate)
            total = quantize_decimal(base + vat)
            if hasattr(self, 'tvsh_entry'): self.tvsh_var.set(f"{vat:.2f}")
            if hasattr(self, 'v_me_tvsh_entry'): self.v_me_tvsh_var.set(f"{total:.2f}")
        elif hasattr(self, 'tvsh_entry') and hasattr(self, 'v_me_tvsh_entry') and \
             not self.tvsh_entry.focus_get() and not self.v_me_tvsh_entry.focus_get():
            # Clear others only if base is invalid AND other fields aren't the source of input
            self.tvsh_var.set("")
            self.v_me_tvsh_var.set("")
        self._set_calculation_flag(False)

    def _calculate_from_vat(self, *args):
        """Calculates Base and Total when VAT amount is entered."""
        if self._calculating or not hasattr(self, 'tvsh_entry'): return
        self._set_calculation_flag(True)
        vat = self._safely_get_decimal(self.tvsh_var)
        if vat is not None and vat >= 0 and self.vat_rate > 0:
            base = quantize_decimal(vat / self.vat_rate)
            total = quantize_decimal(base + vat)
            if hasattr(self, 'v_pa_tvsh_entry'): self.v_pa_tvsh_var.set(f"{base:.2f}")
            if hasattr(self, 'v_me_tvsh_entry'): self.v_me_tvsh_var.set(f"{total:.2f}")
        elif hasattr(self, 'v_pa_tvsh_entry') and hasattr(self, 'v_me_tvsh_entry') and \
             not self.v_pa_tvsh_entry.focus_get() and not self.v_me_tvsh_entry.focus_get():
            self.v_pa_tvsh_var.set("")
            self.v_me_tvsh_var.set("")
        self._set_calculation_flag(False)

    def _calculate_from_total(self, *args):
        """Calculates Base and VAT when Total amount is entered."""
        if self._calculating or not hasattr(self, 'v_me_tvsh_entry'): return
        self._set_calculation_flag(True)
        total = self._safely_get_decimal(self.v_me_tvsh_var)
        denominator = (1 + self.vat_rate)
        if total is not None and total >= 0 and denominator > 0:
            base = quantize_decimal(total / denominator)
            vat = quantize_decimal(total - base) # More accurate than base*rate after division
            if hasattr(self, 'v_pa_tvsh_entry'): self.v_pa_tvsh_var.set(f"{base:.2f}")
            if hasattr(self, 'tvsh_entry'): self.tvsh_var.set(f"{vat:.2f}")
        elif hasattr(self, 'v_pa_tvsh_entry') and hasattr(self, 'tvsh_entry') and \
             not self.v_pa_tvsh_entry.focus_get() and not self.tvsh_entry.focus_get():
            self.v_pa_tvsh_var.set("")
            self.tvsh_var.set("")
        self._set_calculation_flag(False)


    # --- Data Handling & Actions ---
    def _validate_inputs(self):
        """Validates input fields and returns status and calculated Decimal values."""
        errors = []
        month_year = get_month_year_str(self.year_var.get(), self.month_var.get())

        # Amounts are most critical
        base = self._safely_get_decimal(self.v_pa_tvsh_var)
        vat = self._safely_get_decimal(self.tvsh_var)
        total = self._safely_get_decimal(self.v_me_tvsh_var)

        if base is None and vat is None and total is None:
             errors.append("Enter at least one amount (Base, VAT, or Total).")
        elif base is None or vat is None or total is None:
             # This can happen if calculation failed or only one field was entered and cleared
             errors.append("Amounts invalid or incomplete. Check calculations.")
        elif base < 0 or vat < 0 or total < 0:
             errors.append("Amounts cannot be negative.")
        # Check consistency with tolerance
        elif abs(base + vat - total) > Decimal('0.01'):
             errors.append(f"Amounts inconsistent (Base {base:.2f} + VAT {vat:.2f} != Total {total:.2f}).")

        if errors:
            messagebox.showerror("Input Error", "\n".join(errors))
            return False, None, None, None, None
        # Return validated values
        return True, month_year, quantize_decimal(base), quantize_decimal(vat), quantize_decimal(total)

    def _add_or_update_entry(self):
        """Handles adding a new entry or updating an existing one."""
        is_valid, month_year, base_val, vat_val, total_val = self._validate_inputs()
        if not is_valid:
            return

        counterparty_name = self.counterparty_var.get().strip()

        entry_data = {
            # Use get() with defaults for robustness, though variables should exist
            "company": self.company_var.get(),
            "month_year": month_year,
            "transaction_type": self.transaction_type_var.get(),
            "invoice_no": self.invoice_var.get().strip(),
            "counterparty": counterparty_name,
            "base": float(base_val), # Store as float for JSON compatibility
            "vat": float(vat_val),
            "total": float(total_val)
        }

        if self.editing_id: # --- Update existing entry ---
            found = False
            for i, trans in enumerate(self.transactions):
                if trans.get('id') == self.editing_id:
                    entry_data['id'] = self.editing_id # Preserve original ID
                    self.transactions[i] = entry_data
                    found = True
                    break
            if not found:
                 messagebox.showerror("Error", "Could not find the entry to update. Please cancel edit.")
                 # Optionally call self._cancel_edit() here too
                 return
            self._cancel_edit() # Reset form state after successful update
        else: # --- Add new entry ---
             entry_data['id'] = datetime.datetime.now().strftime("%Y%m%d%H%M%S%f") # Unique ID based on time
             self.transactions.append(entry_data)

        # Update known counterparties list
        if counterparty_name: # Only add non-empty names
            self.known_counterparties.add(counterparty_name)
            # Update combobox values
            if hasattr(self, 'counterparty_combo'): # Check if widget exists
                 self.counterparty_combo['values'] = sorted(list(self.known_counterparties))

        self._populate_treeview()
        self._update_live_data_display() # Recalculate live data and balances
        self._save_data() # Save after every successful add/update
        self._clear_entry_fields() # Clear only specific fields for next entry
        # Keep focus logic in _focus_next_widget triggered by button invoke

    def _clear_entry_fields(self):
        """Clears only fields specific to one transaction, keeps persistent ones."""
        self._set_calculation_flag(True) # Prevent calculations during clear
        self.invoice_var.set("")
        self.counterparty_var.set("")
        self.v_me_tvsh_var.set("")
        self.v_pa_tvsh_var.set("")
        self.tvsh_var.set("")
        self._set_calculation_flag(False)

    def _edit_entry(self, event=None): # Allow calling without event (from button)
        """Populates the form with data from the selected Treeview item for editing."""
        selected_items = self.tree.selection()
        if not selected_items:
            # Optional: provide feedback if Edit button clicked with no selection
            # messagebox.showinfo("Edit", "Select an entry in the table to edit.")
            return
        item_id = selected_items[0] # Get Treeview internal item ID
        entry_id_tag = self.tree.item(item_id, 'tags')

        if not entry_id_tag:
             messagebox.showerror("Error", "Selected item is missing its ID tag.")
             return
        entry_id = entry_id_tag[0] # Get our unique transaction ID from the tag

        # Find the transaction data in our list
        entry_to_edit = None
        for trans in self.transactions:
             if trans.get('id') == entry_id:
                  entry_to_edit = trans
                  break

        if not entry_to_edit:
             messagebox.showerror("Error", f"Could not find transaction data for ID {entry_id}.")
             return

        # --- Populate the form ---
        self._set_calculation_flag(True) # Prevent calculations during population
        self.editing_id = entry_id # Set mode to Update

        try:
            # Attempt to parse month/year - handle potential errors
            year_str, month_str = entry_to_edit.get('month_year', '').split('-')
            month_num = int(month_str)
            # Ensure month index is valid for MONTH_NAMES list
            month_name = MONTH_NAMES[month_num - 1] if 0 < month_num <= len(MONTH_NAMES) else self.month_var.get()
            self.year_var.set(int(year_str))
            self.month_var.set(month_name)
        except (ValueError, AttributeError, IndexError):
             # Keep existing month/year if parsing fails
             print(f"Warning: Could not parse month_year '{entry_to_edit.get('month_year')}' for editing.")
             pass # Keep existing selections

        self.company_var.set(entry_to_edit.get('company', COMPANIES[0]))
        self.transaction_type_var.set(entry_to_edit.get('transaction_type', TRANSACTION_TYPES[0]))
        self.invoice_var.set(entry_to_edit.get('invoice_no', ''))
        self.counterparty_var.set(entry_to_edit.get('counterparty', ''))
        # Populate amounts directly using stored values
        self.v_me_tvsh_var.set(f"{entry_to_edit.get('total', 0.0):.2f}")
        self.v_pa_tvsh_var.set(f"{entry_to_edit.get('base', 0.0):.2f}")
        self.tvsh_var.set(f"{entry_to_edit.get('vat', 0.0):.2f}")

        self._set_calculation_flag(False) # Re-enable calculations

        # --- Change UI state for editing ---
        self.done_button.config(text="Update Entry", bootstyle=PRIMARY)
        # Place cancel button (assuming done_button grid info is reliable)
        done_grid_info = self.done_button.grid_info()
        if done_grid_info: # Check if button has been gridded
            self.cancel_button.grid(row=done_grid_info['row'] + 1, column=0, columnspan=2, pady=5, sticky=EW)
        else: # Fallback placement if grid info not ready (shouldn't happen often)
             self.cancel_button.grid(row=9, column=0, columnspan=2, pady=5, sticky=EW)

        self.invoice_entry.focus_set() # Focus for editing

    def _cancel_edit(self):
        """Cancels the editing state, resets the form."""
        self.editing_id = None
        self.done_button.config(text="Done (Add Entry)", bootstyle=SUCCESS)
        self.cancel_button.grid_remove() # Hide cancel button
        self._clear_entry_fields() # Clear only specific fields
        self.tree.selection_remove(self.tree.selection()) # Deselect


    def _delete_selected_entry(self):
        """Deletes selected entries from the Treeview and data list."""
        selected_items = self.tree.selection() # Get Treeview item IDs
        if not selected_items:
            messagebox.showwarning("No Selection", "Select one or more entries in the table to delete.")
            return

        confirm = messagebox.askyesno("Confirm Delete", f"Delete {len(selected_items)} selected transaction(s)? This cannot be undone.")
        if not confirm:
            return

        ids_to_delete = set() # Use a set for efficient lookup
        for item_id in selected_items:
            tag = self.tree.item(item_id, 'tags')
            if tag:
                ids_to_delete.add(tag[0]) # Add our unique transaction ID

        if not ids_to_delete:
             messagebox.showwarning("Deletion Issue", "Could not retrieve IDs for selected items.")
             return

        initial_count = len(self.transactions)
        # Filter the list, keeping transactions whose ID is NOT in the set
        self.transactions = [t for t in self.transactions if t.get('id') not in ids_to_delete]

        deleted_count = initial_count - len(self.transactions)
        if deleted_count > 0:
             self._populate_treeview()
             self._update_live_data_display() # Recalculate live data and balances
             self._save_data() # Save changes
             # Optional: Give feedback
             # print(f"Deleted {deleted_count} transaction(s).")
        else:
             messagebox.showwarning("Deletion Issue", "Selected item(s) were not found in the data list.")

    def _populate_treeview(self):
        """Clears and refills the Treeview with current transaction data."""
        # Keep track of selection before clearing
        selected_ids = {self.tree.item(item, 'tags')[0] for item in self.tree.selection() if self.tree.item(item, 'tags')}

        for item in self.tree.get_children(): self.tree.delete(item)

        # Sort by Month Desc, then by internal ID (time) Desc for consistency
        for entry in sorted(self.transactions, key=lambda x: (x.get('month_year', '0000-00'), x.get('id', '')), reverse=True):
            try:
                # Ensure values exist before formatting
                base_val = entry.get('base', 0.0)
                vat_val = entry.get('vat', 0.0)
                total_val = entry.get('total', 0.0)

                values = (
                    entry.get('month_year', ''),
                    entry.get('company', ''),
                    entry.get('invoice_no', ''),
                    entry.get('counterparty', ''),
                    f"{float(base_val):.2f}", # Treeview needs strings
                    f"{float(vat_val):.2f}",
                    f"{float(total_val):.2f}"
                )
                entry_id = entry.get('id', '')
                # Insert and store the unique ID in the 'tags' attribute
                iid = self.tree.insert('', END, values=values, tags=(entry_id,))
                # Re-select if it was selected before
                if entry_id in selected_ids:
                    self.tree.selection_add(iid)
            except Exception as e:
                 print(f"Error populating treeview row for entry {entry.get('id', 'N/A')}: {e}")


    # Use the robust version of _update_live_data_display from the previous correction
    # (Inside VatEntryApp class)

       # (Inside VatEntryApp class)

    # (Inside VatEntryApp class)

    def _update_live_data_display(self, *args):
        """Recalculates and displays totals and balance for the selected month/company."""
        try:
            # Get current selections
            selected_company = self.company_var.get()
            selected_year = self.year_var.get()
            selected_month_name = self.month_var.get()
            current_month_year = get_month_year_str(selected_year, selected_month_name)
            prev_month_year = get_prev_month_year_str(selected_year, selected_month_name)

            # --- Recalculate Monthly Totals ---
            monthly_totals = defaultdict(lambda: defaultdict(lambda: defaultdict(Decimal)))
            for trans in self.transactions:
                if trans.get('month_year') == current_month_year and \
                   'company' in trans and 'transaction_type' in trans:
                    comp = trans['company']
                    t_type = trans['transaction_type']
                    if comp in COMPANIES and t_type in TRANSACTION_TYPES:
                        base = quantize_decimal(trans.get('base', 0))
                        vat = quantize_decimal(trans.get('vat', 0))
                        total = quantize_decimal(trans.get('total', 0))
                        if not isinstance(monthly_totals[comp][t_type], dict):
                            monthly_totals[comp][t_type] = defaultdict(Decimal)
                        monthly_totals[comp][t_type]['v_pa_tvsh'] += base
                        monthly_totals[comp][t_type]['tvsh'] += vat
                        monthly_totals[comp][t_type]['v_me_tvsh'] += total
            # --- End Recalculation ---

            # --- Update UI Labels and Carry Forward ---
            for company in COMPANIES:
                company_month_data = monthly_totals.get(company, {})

                # Get current month's VAT contributions
                sales_vat = company_month_data.get("Sales", {}).get('tvsh', Decimal(0))
                imports_vat = company_month_data.get("Imports", {}).get('tvsh', Decimal(0))
                local_vat = company_month_data.get("Local", {}).get('tvsh', Decimal(0))

                # --- Get Previous Month's Surplus (Carry Forward) ---
                # Retrieve the stored value (negative or zero)
                prev_surplus_carried_forward = self.carry_forward_data.get(company, {}).get(prev_month_year, Decimal(0))

                # --- Calculate Current Month's Net VAT and Balance ---
                current_month_net_vat = sales_vat - imports_vat - local_vat
                current_balance = prev_surplus_carried_forward + current_month_net_vat

                # --- Update UI Labels ---
                # Update transaction type labels
                for t_type in TRANSACTION_TYPES:
                    type_data = company_month_data.get(t_type, {})
                    if company in self.live_labels and t_type in self.live_labels[company]:
                         if 'v_me_tvsh' in self.live_labels[company][t_type]: self.live_labels[company][t_type]['v_me_tvsh'].config(text=format_curr(type_data.get('v_me_tvsh', 0)))
                         if 'v_pa_tvsh' in self.live_labels[company][t_type]: self.live_labels[company][t_type]['v_pa_tvsh'].config(text=format_curr(type_data.get('v_pa_tvsh', 0)))
                         if 'tvsh' in self.live_labels[company][t_type]: self.live_labels[company][t_type]['tvsh'].config(text=format_curr(type_data.get('tvsh', 0)))

                # Update summary labels
                if company in self.live_labels and 'summary' in self.live_labels[company]:
                    # --- DISPLAY CHANGE HERE: Show the ACTUAL carried forward value (negative or zero) ---
                    if 'prev_surplus' in self.live_labels[company]['summary']:
                         self.live_labels[company]['summary']['prev_surplus'].config(text=format_curr(prev_surplus_carried_forward)) # REMOVED abs()

                    # Display the final balance
                    if 'balance' in self.live_labels[company]['summary']:
                         self.live_labels[company]['summary']['balance'].config(text=format_curr(current_balance))

                         # Set balance label style (Negative = Surplus = Green, Positive = Due = Red)
                         if current_balance < 0: style = SUCCESS
                         elif current_balance > 0: style = DANGER
                         else: style = DEFAULT
                         self.live_labels[company]['summary']['balance'].config(bootstyle=style)

                # --- Update Carry Forward Data for *Next* Month (Logic remains the same) ---
                if company not in self.carry_forward_data:
                    self.carry_forward_data[company] = defaultdict(Decimal)

                if current_balance < 0:
                    # Store the negative balance (surplus) for the *current* month
                    self.carry_forward_data[company][current_month_year] = quantize_decimal(current_balance)
                else:
                    # If DUE (positive) or ZERO, ensure no surplus is carried forward *from this month*
                    if current_month_year in self.carry_forward_data.get(company, {}):
                        del self.carry_forward_data[company][current_month_year]
            # --- End Company Loop ---

        except Exception as e:
             print(f"Error updating live data display: {e}")
             traceback.print_exc()
    def _filter_counterparties(self, event):
        """Filters the counterparty combobox list based on typed input."""
        typed = self.counterparty_var.get().lower()
        if not typed:
            # Show full sorted list if input is empty
            filtered_list = sorted(list(self.known_counterparties))
        else:
            # Show matches (case-insensitive)
            filtered_list = sorted([cp for cp in self.known_counterparties if typed in cp.lower()])

        # Prevent list from becoming empty if no matches, maybe show original list then?
        # Or just show the (potentially empty) filtered list. Current way shows empty if no match.
        if hasattr(self, 'counterparty_combo'): # Check widget exists
            self.counterparty_combo['values'] = filtered_list

    # --- Save/Load ---
    def _save_data(self):
        """Saves current transactions, carry-forward data, and counterparties to JSON."""
        # Convert defaultdicts and Decimals to standard types for JSON
        data_to_save = {
            "transactions": self.transactions,
            # Store carry_forward only if it contains data
            "carry_forward": {comp: {my: str(surplus) for my, surplus in months.items()}
                              for comp, months in self.carry_forward_data.items() if months},
            "counterparties": sorted(list(self.known_counterparties))
        }
        try:
            with open(DATA_FILE, 'w', encoding='utf-8') as f:
                json.dump(data_to_save, f, indent=2, ensure_ascii=False) # Use ensure_ascii=False for non-latin chars
            # print(f"Data saved to {DATA_FILE}") # Optional feedback
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save data to {DATA_FILE}:\n{e}")

    # Use the robust version of _load_data from the previous correction
    def _load_data(self):
        """Loads data from JSON file, performing type checks."""
        if not os.path.exists(DATA_FILE):
            print(f"Data file {DATA_FILE} not found. Starting fresh.")
            self.transactions = []
            self.carry_forward_data = defaultdict(lambda: defaultdict(Decimal))
            self.known_counterparties = set()
            return

        try:
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)

            # Load transactions with checks
            loaded_transactions = data.get("transactions", [])
            if isinstance(loaded_transactions, list):
                 self.transactions = [t for t in loaded_transactions if isinstance(t, dict)]
                 if len(self.transactions) != len(loaded_transactions):
                      print("Warning: Some invalid transaction entries removed during load.")
            else:
                print("Warning: Transactions data in file is not a list. Resetting.")
                self.transactions = []

            # Load carry_forward with checks and conversion
            loaded_carry_forward = data.get("carry_forward", {})
            self.carry_forward_data = defaultdict(lambda: defaultdict(Decimal))
            if isinstance(loaded_carry_forward, dict):
                for comp, months in loaded_carry_forward.items():
                    if isinstance(months, dict):
                        for my, surplus_str in months.items():
                            try: self.carry_forward_data[comp][my] = Decimal(surplus_str)
                            except: print(f"Warning: Invalid carry forward value '{surplus_str}' for {comp}/{my}")
                    else: print(f"Warning: Invalid carry forward month data for company {comp}")
            else: print("Warning: Carry forward data in file is not a dictionary.")

            # Load counterparties with checks
            loaded_counterparties = data.get("counterparties", [])
            if isinstance(loaded_counterparties, list):
                self.known_counterparties = set(cp for cp in loaded_counterparties if isinstance(cp, str))
            else:
                print("Warning: Counterparties data in file is not a list.")
                self.known_counterparties = set()

            # Update counterparty combobox list *after* loading
            if hasattr(self, 'counterparty_combo'):
                self.counterparty_combo['values'] = sorted(list(self.known_counterparties))
            # print(f"Data loaded successfully from {DATA_FILE}") # Optional feedback

        except json.JSONDecodeError:
             messagebox.showwarning("Load Warning", f"{DATA_FILE} is corrupt or empty. Starting fresh.")
             self.transactions, self.carry_forward_data, self.known_counterparties = [], defaultdict(lambda: defaultdict(Decimal)), set()
        except Exception as e:
            messagebox.showerror("Load Error", f"Could not load data from {DATA_FILE}:\n{e}")
            self.transactions, self.carry_forward_data, self.known_counterparties = [], defaultdict(lambda: defaultdict(Decimal)), set()

    def _on_closing(self):
        """Saves data automatically when the window is closed."""
        self._save_data() # Save automatically on close
        self.root.destroy()

# --- Run the Application ---
# THIS BLOCK MUST START AT THE VERY BEGINNING OF THE LINE (COLUMN 0)
if __name__ == "__main__":
    app_theme = 'litera' # Default theme ('cosmo', 'flatly', 'litera', 'minty', 'lumen', 'sandstone', 'united', 'yeti', 'pulse', 'journal', 'darkly', 'superhero', 'solar', 'cyborg', 'vapor')
    try:
        root = tb.Window(themename=app_theme)
        # Optional: Set minimum window size
        root.minsize(width=950, height=600)
        app = VatEntryApp(root)
        root.mainloop()
    except Exception as e:
        print(f"FATAL: An error occurred during application startup: {e}")
        # Import traceback here only if needed
        # import traceback # Already imported at top level
        traceback.print_exc()
        # Fallback Tkinter error message
        try:
            # Create a simple Tk root only for the error message box
            err_root = tk.Tk()
            err_root.withdraw() # Hide the empty window
            messagebox.showerror("Startup Error", f"Failed to start:\n{e}\n\nSee console for details.")
            err_root.destroy() # Clean up the error window
        except:
            # If even showing the error fails, just pass
            pass
