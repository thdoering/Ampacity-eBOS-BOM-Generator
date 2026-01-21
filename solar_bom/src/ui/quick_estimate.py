import tkinter as tk
from tkinter import ttk
from typing import Optional
from pathlib import Path
import json

class QuickEstimate(ttk.Frame):
    """Quick estimation tool for early-stage project sizing"""
    
    def __init__(self, parent, current_project=None):
        super().__init__(parent)
        self.current_project = current_project
        self.pricing_data = self.load_pricing_data()
        self.setup_ui()

    def load_pricing_data(self):
        """Load pricing data from JSON file"""
        try:
            pricing_path = Path('data/pricing_data.json')
            if pricing_path.exists():
                with open(pricing_path, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading pricing data: {e}")
        return {}

    def round_whip_length(self, raw_length_ft):
        """Apply 5% waste factor and round up to nearest 5ft increment (min 10ft)"""
        WASTE_FACTOR = 1.05  # 5% adder
        length_with_waste = raw_length_ft * WASTE_FACTOR
        # Round up to nearest 5ft
        rounded = 5 * ((length_with_waste + 5 - 0.1) // 5 + 1)
        # Minimum 10ft
        return max(10, int(rounded))
    
    def setup_ui(self):
        """Create and arrange UI components"""
        # Main container with padding
        main_container = ttk.Frame(self, padding="10")
        main_container.pack(fill='both', expand=True)
        
        # Title
        title_label = ttk.Label(main_container, text="Quick Estimate", font=('Helvetica', 14, 'bold'))
        title_label.pack(anchor='w', pady=(0, 10))
        
        # Description
        desc_label = ttk.Label(main_container, text="Early-stage BOM estimation for bid and preliminary designs")
        desc_label.pack(anchor='w', pady=(0, 20))
        
        # Input frame
        input_frame = ttk.LabelFrame(main_container, text="Project Inputs", padding="10")
        input_frame.pack(fill='x', pady=(0, 10))
        
        # Combiner Boxes
        combiner_frame = ttk.Frame(input_frame)
        combiner_frame.pack(fill='x', pady=5)
        
        ttk.Label(combiner_frame, text="Number of Combiner Boxes:").pack(side='left')
        self.combiner_count_var = tk.StringVar(value="0")
        combiner_spinbox = ttk.Spinbox(
            combiner_frame, 
            from_=0, 
            to=1000, 
            textvariable=self.combiner_count_var,
            width=10
        )
        combiner_spinbox.pack(side='left', padx=(10, 0))
        
        # Combiner Unit Price
        combiner_price_frame = ttk.Frame(input_frame)
        combiner_price_frame.pack(fill='x', pady=5)
        
        ttk.Label(combiner_price_frame, text="Combiner Unit Price ($):").pack(side='left')
        default_combiner_price = self.get_default_combiner_price()
        self.combiner_price_var = tk.StringVar(value=f"{default_combiner_price:.2f}")
        combiner_price_entry = ttk.Entry(combiner_price_frame, textvariable=self.combiner_price_var, width=10)
        combiner_price_entry.pack(side='left', padx=(10, 0))
        
        # Show source of price
        price_source = "(from pricing data)" if default_combiner_price > 0 else "(manual entry)"
        self.combiner_price_source_label = ttk.Label(combiner_price_frame, text=price_source, font=('Helvetica', 8), foreground='gray')
        self.combiner_price_source_label.pack(side='left', padx=(5, 0))
        
        # Row Spacing
        row_spacing_frame = ttk.Frame(input_frame)
        row_spacing_frame.pack(fill='x', pady=5)
        
        ttk.Label(row_spacing_frame, text="Row Spacing (ft):").pack(side='left')
        self.row_spacing_var = tk.StringVar(value="20")
        row_spacing_spinbox = ttk.Spinbox(
            row_spacing_frame, 
            from_=1, 
            to=100, 
            textvariable=self.row_spacing_var,
            width=10
        )
        row_spacing_spinbox.pack(side='left', padx=(10, 0))
        
        # Max Rows from Combiner
        max_rows_frame = ttk.Frame(input_frame)
        max_rows_frame.pack(fill='x', pady=5)
        
        ttk.Label(max_rows_frame, text="Max Rows from Combiner:").pack(side='left')
        self.max_rows_var = tk.StringVar(value="4")
        max_rows_spinbox = ttk.Spinbox(
            max_rows_frame, 
            from_=1, 
            to=50, 
            textvariable=self.max_rows_var,
            width=10
        )
        max_rows_spinbox.pack(side='left', padx=(10, 0))
        ttk.Label(max_rows_frame, text="(furthest row from combiner)", font=('Helvetica', 8), foreground='gray').pack(side='left', padx=(5, 0))

        # Separator before module info
        ttk.Separator(input_frame, orient='horizontal').pack(fill='x', pady=10)
        
        # Module Info Section
        module_section_label = ttk.Label(input_frame, text="Module Info", font=('Helvetica', 10, 'bold'))
        module_section_label.pack(anchor='w', pady=(0, 5))
        
        # Module width
        module_width_frame = ttk.Frame(input_frame)
        module_width_frame.pack(fill='x', pady=5)
        
        ttk.Label(module_width_frame, text="Module Width (mm):").pack(side='left')
        self.module_width_var = tk.StringVar(value="1134")
        ttk.Entry(module_width_frame, textvariable=self.module_width_var, width=10).pack(side='left', padx=(10, 0))
        ttk.Label(module_width_frame, text="(dimension along tracker)", font=('Helvetica', 8), foreground='gray').pack(side='left', padx=(5, 0))
        
        # Modules per string
        modules_per_string_frame = ttk.Frame(input_frame)
        modules_per_string_frame.pack(fill='x', pady=5)
        
        ttk.Label(modules_per_string_frame, text="Modules per String:").pack(side='left')
        self.modules_per_string_var = tk.StringVar(value="28")
        ttk.Spinbox(
            modules_per_string_frame, 
            from_=1, 
            to=100, 
            textvariable=self.modules_per_string_var,
            width=10
        ).pack(side='left', padx=(10, 0))
        
        # Separator
        ttk.Separator(input_frame, orient='horizontal').pack(fill='x', pady=10)
        
        # Tracker Configurations Section
        tracker_section_label = ttk.Label(input_frame, text="Tracker Configurations", font=('Helvetica', 10, 'bold'))
        tracker_section_label.pack(anchor='w', pady=(0, 5))
        
        # Tracker table frame
        tracker_table_frame = ttk.Frame(input_frame)
        tracker_table_frame.pack(fill='x', pady=5)
        
        # Headers
        ttk.Label(tracker_table_frame, text="Strings/Tracker", width=15).grid(row=0, column=0, padx=5, pady=2)
        ttk.Label(tracker_table_frame, text="Quantity", width=10).grid(row=0, column=1, padx=5, pady=2)
        ttk.Label(tracker_table_frame, text="Harness Config", width=12).grid(row=0, column=2, padx=5, pady=2)
        ttk.Label(tracker_table_frame, text="", width=5).grid(row=0, column=3, padx=5, pady=2)  # For remove button
        
        # Container for tracker rows
        self.tracker_rows_frame = ttk.Frame(tracker_table_frame)
        self.tracker_rows_frame.grid(row=1, column=0, columnspan=3, sticky='ew')
        
        # List to track tracker row data
        self.tracker_rows = []
        
        # Help text for harness config
        harness_help = ttk.Label(input_frame, text="Harness Config: e.g., '3' for 3-string harness, '2+1' for 2-string + 1-string", 
                                  font=('Helvetica', 8), foreground='gray')
        harness_help.pack(anchor='w', pady=(5, 0))

        # Add button
        add_tracker_btn = ttk.Button(input_frame, text="+ Add Tracker Type", command=self.add_tracker_row)
        add_tracker_btn.pack(anchor='w', pady=(5, 0))
        
        # Add one default row
        self.add_tracker_row()
        
        # Calculate button
        calc_btn = ttk.Button(main_container, text="Calculate Estimate", command=self.calculate_estimate)
        calc_btn.pack(anchor='w', pady=(10, 10))
        
        # Results frame
        results_frame = ttk.LabelFrame(main_container, text="Estimated BOM", padding="10")
        results_frame.pack(fill='both', expand=True, pady=(0, 0))
        
        # Results treeview
        columns = ('item', 'quantity', 'unit', 'unit_price', 'total_price')
        self.results_tree = ttk.Treeview(results_frame, columns=columns, show='headings', height=10)
        self.results_tree.heading('item', text='Item')
        self.results_tree.heading('quantity', text='Quantity')
        self.results_tree.heading('unit', text='Unit')
        self.results_tree.heading('unit_price', text='Unit Price')
        self.results_tree.heading('total_price', text='Total Price')
        self.results_tree.column('item', width=200, anchor='w')
        self.results_tree.column('quantity', width=80, anchor='center')
        self.results_tree.column('unit', width=50, anchor='center')
        self.results_tree.column('unit_price', width=80, anchor='e')
        self.results_tree.column('total_price', width=100, anchor='e')
        self.results_tree.pack(fill='both', expand=True)

    def add_tracker_row(self):
        """Add a new tracker configuration row"""
        row_idx = len(self.tracker_rows)
        
        row_frame = ttk.Frame(self.tracker_rows_frame)
        row_frame.pack(fill='x', pady=2)
        
        # Strings per tracker dropdown
        strings_var = tk.StringVar(value="2")
        strings_combo = ttk.Combobox(row_frame, textvariable=strings_var, values=["1", "2", "3", "4", "5", "6"], width=12, state='readonly')
        strings_combo.pack(side='left', padx=5)
        
        # Quantity spinbox
        qty_var = tk.StringVar(value="0")
        qty_spinbox = ttk.Spinbox(row_frame, from_=0, to=10000, textvariable=qty_var, width=8)
        qty_spinbox.pack(side='left', padx=5)
        
        # Harness config dropdown (will be populated based on strings selection)
        harness_var = tk.StringVar(value="2")
        harness_combo = ttk.Combobox(row_frame, textvariable=harness_var, width=10, state='readonly')
        harness_combo.pack(side='left', padx=5)
        
        # Remove button
        remove_btn = ttk.Button(row_frame, text="X", width=3, 
                                command=lambda: self.remove_tracker_row(row_frame, row_data))
        remove_btn.pack(side='left', padx=5)
        
        # Store row data
        row_data = {
            'frame': row_frame,
            'strings_var': strings_var,
            'qty_var': qty_var,
            'harness_var': harness_var,
            'harness_combo': harness_combo
        }
        self.tracker_rows.append(row_data)
        
        # Update harness options when strings changes
        def update_harness_options(*args):
            try:
                num_strings = int(strings_var.get())
                options = self.get_harness_combinations(num_strings)
                harness_combo['values'] = options
                if options:
                    harness_var.set(options[0])  # Default to first option (single harness)
            except ValueError:
                pass
        
        strings_var.trace('w', update_harness_options)
        update_harness_options()  # Initialize
    
    def remove_tracker_row(self, frame, row_data):
        """Remove a tracker configuration row"""
        if len(self.tracker_rows) > 1:  # Keep at least one row
            frame.destroy()
            self.tracker_rows.remove(row_data)

    def get_harness_combinations(self, num_strings):
        """Get valid harness combinations for a given number of strings"""
        if num_strings == 1:
            return ["1"]
        elif num_strings == 2:
            return ["2", "1+1"]
        elif num_strings == 3:
            return ["3", "2+1", "1+1+1"]
        elif num_strings == 4:
            return ["4", "3+1", "2+2", "2+1+1", "1+1+1+1"]
        elif num_strings == 5:
            return ["5", "4+1", "3+2", "3+1+1", "2+2+1", "2+1+1+1", "1+1+1+1+1"]
        elif num_strings == 6:
            return ["6", "5+1", "4+2", "4+1+1", "3+3", "3+2+1", "3+1+1+1", 
                    "2+2+2", "2+2+1+1", "2+1+1+1+1", "1+1+1+1+1+1"]
        else:
            return [str(num_strings)]
        
    def parse_harness_config(self, config_str):
        """Parse harness config string like '2+1' into list of integers [2, 1]"""
        try:
            return [int(x) for x in config_str.split('+')]
        except:
            return []

    def calculate_estimate(self):
        """Calculate and display the BOM estimate"""
        # Clear previous results
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        # Get inputs
        try:
            combiner_count = int(self.combiner_count_var.get())
        except ValueError:
            combiner_count = 0
        
        try:
            row_spacing = float(self.row_spacing_var.get())
        except ValueError:
            row_spacing = 20.0
        
        try:
            max_rows = int(self.max_rows_var.get())
        except ValueError:
            max_rows = 4
        
        # Get tracker data
        total_trackers = 0
        tracker_summary = []
        for row in self.tracker_rows:
            try:
                strings = int(row['strings_var'].get())
                qty = int(row['qty_var'].get())
                if qty > 0:
                    total_trackers += qty
                    tracker_summary.append({'strings': strings, 'qty': qty})
            except ValueError:
                continue
        
        # === COMBINER BOXES ===
        self.running_total = 0.0  # Track grand total
        
        if combiner_count > 0:
            unit_price = self.get_combiner_price()
            total_price = unit_price * combiner_count
            self.running_total += total_price
            self.results_tree.insert('', 'end', values=(
                'Combiner Boxes', 
                combiner_count, 
                'ea',
                f'${unit_price:,.2f}',
                f'${total_price:,.2f}'
            ))
        
        # === TRACKERS ===
        for tracker in tracker_summary:
            self.results_tree.insert('', 'end', values=(
                f"{tracker['strings']}-String Trackers", 
                tracker['qty'], 
                'ea'
            ))

        # === HARNESSES ===
        harness_counts = {}  # {num_strings: quantity}
        for row in self.tracker_rows:
            try:
                qty = int(row['qty_var'].get())
                harness_config = row['harness_var'].get()
                if qty > 0 and harness_config:
                    harness_sizes = self.parse_harness_config(harness_config)
                    for size in harness_sizes:
                        if size not in harness_counts:
                            harness_counts[size] = 0
                        harness_counts[size] += qty
            except ValueError:
                continue
        
        if harness_counts:
            self.results_tree.insert('', 'end', values=('--- HARNESSES ---', '', ''))
            for size in sorted(harness_counts.keys(), reverse=True):
                self.results_tree.insert('', 'end', values=(
                    f"{size}-String Harness",
                    harness_counts[size],
                    'ea'
                ))
        
        # === EXTENDERS ===
        # Each harness needs 2 extenders:
        # - Short side: ~10ft (harness ends near row end)
        # - Long side: one string length back from row end
        total_harnesses = sum(harness_counts.values()) if harness_counts else 0
        
        if total_harnesses > 0:
            try:
                module_width_mm = float(self.module_width_var.get())
                modules_per_string = int(self.modules_per_string_var.get())
                
                # Convert module width to feet
                module_width_ft = module_width_mm / 304.8  # mm to ft
                
                # String length in feet
                string_length_ft = module_width_ft * modules_per_string
                
                # Short extender: 10ft standard
                short_extender_raw = 10
                short_extender_length = self.round_whip_length(short_extender_raw)
                
                # Long extender: string length
                long_extender_length = self.round_whip_length(string_length_ft)
                
                self.results_tree.insert('', 'end', values=('--- EXTENDERS ---', '', ''))
                self.results_tree.insert('', 'end', values=(
                    f"Extender {short_extender_length}ft (short side)",
                    total_harnesses,
                    'ea'
                ))
                self.results_tree.insert('', 'end', values=(
                    f"Extender {long_extender_length}ft (long side)",
                    total_harnesses,
                    'ea'
                ))
            except ValueError:
                pass
        
        if total_trackers > 0:
            self.results_tree.insert('', 'end', values=('Total Trackers', total_trackers, 'ea'))
        
        # === WHIPS ===
        # Assumption: combiner is in middle of block
        # Each row distance has 2 whips (one on each side of combiner)
        if combiner_count > 0 and max_rows > 0:
            self.results_tree.insert('', 'end', values=('--- WHIPS ---', '', ''))
            
            total_whip_length = 0
            whips_per_combiner = 0
            
            for row_num in range(1, max_rows + 1):
                raw_whip_length = row_num * row_spacing
                whip_length = self.round_whip_length(raw_whip_length)
                whips_at_length = 2  # 2 whips per row distance (both sides of combiner)
                total_at_length = whips_at_length * combiner_count
                length_subtotal = whip_length * total_at_length
                
                self.results_tree.insert('', 'end', values=(
                    f"Whip {whip_length}ft", 
                    total_at_length, 
                    'ea'
                ))
                
                total_whip_length += length_subtotal
                whips_per_combiner += whips_at_length
            
            self.results_tree.insert('', 'end', values=(
                'Total Whip Length', 
                f"{total_whip_length:.0f}", 
                'ft'
            ))

    def get_default_combiner_price(self):
        """Get default combiner box price from pricing data, or 0 if not available"""
        try:
            combiner_data = self.pricing_data.get('combiner_boxes', {})
            # Try to get first available price
            for key, value in combiner_data.items():
                if isinstance(value, dict) and 'price' in value:
                    return float(value['price'])
                elif isinstance(value, (int, float)):
                    return float(value)
            return 0.0
        except:
            return 0.0
    
    def get_combiner_price(self):
        """Get combiner price from user input field"""
        try:
            return float(self.combiner_price_var.get())
        except ValueError:
            return 0.0