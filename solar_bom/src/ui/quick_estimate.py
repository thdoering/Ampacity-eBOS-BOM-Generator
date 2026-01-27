import tkinter as tk
from tkinter import ttk
from typing import Optional, Dict, List, Any
from pathlib import Path
import json
import uuid
from datetime import datetime


class QuickEstimate(ttk.Frame):
    """Quick estimation tool for early-stage project sizing with hierarchical structure"""
    
    def __init__(self, parent, current_project=None, estimate_id=None, on_save=None):
        super().__init__(parent)
        self.current_project = current_project
        self.estimate_id = estimate_id
        self.on_save = on_save
        self.pricing_data = self.load_pricing_data()
        
        # Data structure for subarrays and blocks
        self.subarrays: Dict[str, Dict[str, Any]] = {}
        
        # Global settings defaults
        self.module_width_default = 1134
        self.modules_per_string_default = 28
        self.row_spacing_default = 20.0
        
        # Track currently selected item
        self.selected_item_id = None
        self.selected_item_type = None  # 'subarray' or 'block'
        
        self.setup_ui()
        
        # Load existing estimate or create default
        if self.estimate_id and self.current_project:
            self.load_estimate()
        else:
            # Add a default subarray with one block to start
            self.add_subarray(auto_add_block=True)

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

    def generate_id(self, prefix: str) -> str:
        """Generate a unique ID with a prefix"""
        return f"{prefix}_{uuid.uuid4().hex[:8]}"
    
    def get_next_block_name(self, subarray_id: str, source_name: str) -> str:
        """Generate the next logical block name based on the source name pattern"""
        import re
        
        # Try to find a trailing number (with optional leading zeros)
        match = re.match(r'^(.*?)(\d+)$', source_name)
        
        if not match:
            # No number found, just append (Copy)
            return f"{source_name} (Copy)"
        
        prefix = match.group(1)
        number_str = match.group(2)
        number = int(number_str)
        
        # Determine the padding (e.g., "01" is width 2, "001" is width 3)
        padding = len(number_str)
        
        # Get existing block names in this subarray
        existing_names = set()
        if subarray_id in self.subarrays:
            for block_data in self.subarrays[subarray_id]['blocks'].values():
                existing_names.add(block_data['name'])
        
        # Find the next available number
        next_number = number + 1
        while True:
            new_name = f"{prefix}{str(next_number).zfill(padding)}"
            if new_name not in existing_names:
                return new_name
            next_number += 1
            # Safety limit to prevent infinite loop
            if next_number > number + 1000:
                return f"{source_name} (Copy)"
            
    def disable_combobox_scroll(self, combobox):
        """Prevent combobox from responding to mouse wheel"""
        def _ignore_scroll(event):
            return "break"
        
        combobox.bind("<MouseWheel>", _ignore_scroll)
        # For Linux compatibility
        combobox.bind("<Button-4>", _ignore_scroll)
        combobox.bind("<Button-5>", _ignore_scroll)

    def update_string_count(self):
        """Update the string count label for the current block"""
        if not hasattr(self, 'string_count_label') or not hasattr(self, 'current_block'):
            return
        
        total_strings = 0
        total_trackers = 0
        
        for tracker in self.current_block.get('trackers', []):
            qty = tracker.get('quantity', 0)
            strings = tracker.get('strings', 0)
            total_trackers += qty
            total_strings += qty * strings
        
        self.string_count_label.config(text=f"Total Strings: {total_strings:,}  |  Total Trackers: {total_trackers:,}")

    # ==================== Data Management ====================
    
    def add_subarray(self, auto_add_block: bool = False) -> str:
        """Add a new subarray and return its ID"""
        subarray_id = self.generate_id("subarray")
        subarray_num = len(self.subarrays) + 1
        
        self.subarrays[subarray_id] = {
            'name': f"Subarray {subarray_num}",
            'transformer_mva': 4.0,
            'blocks': {}
        }
        
        # Add to tree view
        self.tree.insert('', 'end', subarray_id, text=f"Subarray {subarray_num}", open=True)
        
        # Optionally add a default block
        if auto_add_block:
            self.add_block(subarray_id)
        
        # Select the new subarray
        self.tree.selection_set(subarray_id)
        self.on_tree_select(None)
        
        return subarray_id

    def add_block(self, subarray_id: str) -> str:
        """Add a new block to a subarray and return its ID"""
        if subarray_id not in self.subarrays:
            return None
        
        block_id = self.generate_id("block")
        block_num = len(self.subarrays[subarray_id]['blocks']) + 1
        
        self.subarrays[subarray_id]['blocks'][block_id] = {
            'name': f"Block {block_num}",
            'type': 'combiner',  # 'combiner' or 'string_inverter'
            'num_rows': 4,
            'trackers': [
                {'strings': 2, 'quantity': 1, 'harness_config': '2'}
            ]
        }
        
        # Add to tree view under the subarray
        self.tree.insert(subarray_id, 'end', block_id, text=f"Block {block_num}")
        
        # Expand the parent subarray
        self.tree.item(subarray_id, open=True)
        
        # Select the new block
        self.tree.selection_set(block_id)
        self.on_tree_select(None)
        
        return block_id

    def delete_selected_item(self):
        """Delete the currently selected subarray or block"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item_id = selection[0]
        parent_id = self.tree.parent(item_id)
        
        if parent_id == '':
            # It's a subarray
            if item_id in self.subarrays:
                del self.subarrays[item_id]
                self.tree.delete(item_id)
        else:
            # It's a block
            if parent_id in self.subarrays:
                if item_id in self.subarrays[parent_id]['blocks']:
                    del self.subarrays[parent_id]['blocks'][item_id]
                    self.tree.delete(item_id)
        
        # Clear the details panel
        self.clear_details_panel()

    def get_subarray_for_block(self, block_id: str) -> Optional[str]:
        """Find which subarray a block belongs to"""
        for subarray_id, subarray_data in self.subarrays.items():
            if block_id in subarray_data['blocks']:
                return subarray_id
        return None

    def round_whip_length(self, raw_length_ft):
        """Apply 5% waste factor and round up to nearest 5ft increment (min 10ft)"""
        WASTE_FACTOR = 1.05
        length_with_waste = raw_length_ft * WASTE_FACTOR
        rounded = 5 * ((length_with_waste + 5 - 0.1) // 5 + 1)
        return max(10, int(rounded))

    def get_default_combiner_price(self):
        """Get default combiner box price from pricing data"""
        try:
            combiner_data = self.pricing_data.get('combiner_boxes', {})
            for key, value in combiner_data.items():
                if isinstance(value, dict) and 'price' in value:
                    return float(value['price'])
                elif isinstance(value, (int, float)):
                    return float(value)
            return 0.0
        except:
            return 0.0

    def get_harness_options(self, num_strings):
        """Get valid harness configuration options for a given number of strings"""
        if num_strings == 1:
            return ["1"]
        elif num_strings == 2:
            return ["2", "1+1"]
        elif num_strings == 3:
            return ["3", "2+1"]
        elif num_strings == 4:
            return ["4", "3+1", "2+2"]
        elif num_strings == 5:
            return ["5", "4+1", "3+2"]
        elif num_strings == 6:
            return ["6", "5+1", "4+2", "3+3"]
        elif num_strings == 7:
            return ["7", "6+1", "5+2", "4+3"]
        elif num_strings == 8:
            return ["8", "7+1", "6+2", "5+3", "4+4"]
        elif num_strings == 9:
            return ["9", "8+1", "7+2", "6+3", "5+4"]
        elif num_strings == 10:
            return ["10", "9+1", "8+2", "7+3", "6+4", "5+5"]
        elif num_strings == 11:
            return ["11", "10+1", "9+2", "8+3", "7+4", "6+5"]
        elif num_strings == 12:
            return ["12", "11+1", "10+2", "9+3", "8+4", "7+5", "6+6"]
        elif num_strings == 13:
            return ["13", "12+1", "11+2", "10+3", "9+4", "8+5", "7+6"]
        elif num_strings == 14:
            return ["14", "13+1", "12+2", "11+3", "10+4", "9+5", "8+6", "7+7"]
        elif num_strings == 15:
            return ["15", "14+1", "13+2", "12+3", "11+4", "10+5", "9+6", "8+7"]
        elif num_strings == 16:
            return ["16", "15+1", "14+2", "13+3", "12+4", "11+5", "10+6", "9+7", "8+8"]
        else:
            return [str(num_strings)]

    def parse_harness_config(self, config_str):
        """Parse harness config string like '2+1' into list of integers [2, 1]"""
        try:
            return [int(x) for x in config_str.split('+')]
        except:
            return []
        
    def load_estimate(self):
        """Load estimate data from the project"""
        if not self.current_project or not self.estimate_id:
            self.add_subarray(auto_add_block=True)
            return
        
        estimate_data = self.current_project.quick_estimates.get(self.estimate_id)
        if not estimate_data:
            self.add_subarray(auto_add_block=True)
            return
        
        # Load global settings
        self.module_width_default = estimate_data.get('module_width_mm', 1134)
        self.modules_per_string_default = estimate_data.get('modules_per_string', 28)
        self.row_spacing_default = estimate_data.get('row_spacing_ft', 20.0)
        
        # Update UI if already set up
        if hasattr(self, 'module_width_var'):
            self.module_width_var.set(str(self.module_width_default))
        if hasattr(self, 'modules_per_string_var'):
            self.modules_per_string_var.set(str(self.modules_per_string_default))
        if hasattr(self, 'row_spacing_var'):
            self.row_spacing_var.set(str(self.row_spacing_default))
        
        # Load subarrays
        saved_subarrays = estimate_data.get('subarrays', {})
        
        if not saved_subarrays:
            # No saved data, create default
            self.add_subarray(auto_add_block=True)
            return
        
        # Clear existing tree items
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.subarrays.clear()
        
        # Reconstruct subarrays and blocks
        for subarray_id, subarray_data in saved_subarrays.items():
            # Add to internal data
            self.subarrays[subarray_id] = {
                'name': subarray_data.get('name', 'Subarray'),
                'transformer_mva': subarray_data.get('transformer_mva', 4.0),
                'blocks': {}
            }
            
            # Add to tree
            self.tree.insert('', 'end', subarray_id, text=subarray_data.get('name', 'Subarray'), open=True)
            
            # Load blocks
            for block_id, block_data in subarray_data.get('blocks', {}).items():
                self.subarrays[subarray_id]['blocks'][block_id] = {
                    'name': block_data.get('name', 'Block'),
                    'type': block_data.get('type', 'combiner'),
                    'num_rows': block_data.get('num_rows', 4),
                    'trackers': block_data.get('trackers', [{'strings': 2, 'quantity': 1, 'harness_config': '2'}])
                }
                
                # Add block to tree
                self.tree.insert(subarray_id, 'end', block_id, text=block_data.get('name', 'Block'))
        
        # Select first item
        first_items = self.tree.get_children()
        if first_items:
            self.tree.selection_set(first_items[0])
            self.on_tree_select(None)

    def save_estimate(self):
        """Save estimate data to the project"""
        if not self.current_project or not self.estimate_id:
            return
        
        # Get current estimate or create new one
        if self.estimate_id not in self.current_project.quick_estimates:
            self.current_project.quick_estimates[self.estimate_id] = {
                'name': 'Unnamed Estimate',
                'created_date': datetime.now().isoformat(),
            }
        
        estimate_data = self.current_project.quick_estimates[self.estimate_id]
        
        # Update global settings
        try:
            estimate_data['module_width_mm'] = float(self.module_width_var.get())
        except ValueError:
            estimate_data['module_width_mm'] = 1134
        
        try:
            estimate_data['modules_per_string'] = int(self.modules_per_string_var.get())
        except ValueError:
            estimate_data['modules_per_string'] = 28
        
        try:
            estimate_data['row_spacing_ft'] = float(self.row_spacing_var.get())
        except ValueError:
            estimate_data['row_spacing_ft'] = 20.0
        
        # Update modified date
        estimate_data['modified_date'] = datetime.now().isoformat()
        
        # Save subarrays
        estimate_data['subarrays'] = {}
        
        for subarray_id, subarray_data in self.subarrays.items():
            estimate_data['subarrays'][subarray_id] = {
                'name': subarray_data['name'],
                'transformer_mva': subarray_data['transformer_mva'],
                'blocks': {}
            }
            
            for block_id, block_data in subarray_data['blocks'].items():
                estimate_data['subarrays'][subarray_id]['blocks'][block_id] = {
                    'name': block_data['name'],
                    'type': block_data['type'],
                    'num_rows': block_data['num_rows'],
                    'trackers': block_data['trackers']
                }
        
        # Notify callback
        if self.on_save:
            self.on_save()

    # ==================== UI Setup ====================
    
    def setup_ui(self):
        """Create and arrange UI components"""
        # Main container with padding
        main_container = ttk.Frame(self, padding="10")
        main_container.pack(fill='both', expand=True)
        
        # Title
        title_label = ttk.Label(main_container, text="Quick Estimate", font=('Helvetica', 14, 'bold'))
        title_label.pack(anchor='w', pady=(0, 5))
        
        # Description
        desc_label = ttk.Label(main_container, text="Early-stage BOM estimation for bid and preliminary designs")
        desc_label.pack(anchor='w', pady=(0, 10))
        
        # Main content area - PanedWindow for resizable split
        paned = ttk.PanedWindow(main_container, orient='horizontal')
        paned.pack(fill='both', expand=True, pady=(0, 10))
        
        # Left panel - Tree view
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=1)
        
        # Right panel - Details
        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=3)
        
        self.setup_tree_panel(left_frame)
        self.setup_details_panel(right_frame)
        
        # Bottom section - Calculate button and results
        bottom_frame = ttk.Frame(main_container)
        bottom_frame.pack(fill='both', expand=True)
        
        # Global settings frame (applies to all calculations)
        settings_frame = ttk.LabelFrame(bottom_frame, text="Global Settings", padding="5")
        settings_frame.pack(fill='x', pady=(0, 10))
        
        settings_inner = ttk.Frame(settings_frame)
        settings_inner.pack(fill='x')
        
        ttk.Label(settings_inner, text="Module Width (mm):").pack(side='left', padx=(0, 5))
        self.module_width_var = tk.StringVar(value=str(getattr(self, 'module_width_default', 1134)))
        ttk.Entry(settings_inner, textvariable=self.module_width_var, width=8).pack(side='left', padx=(0, 15))
        
        ttk.Label(settings_inner, text="Modules per String:").pack(side='left', padx=(0, 5))
        self.modules_per_string_var = tk.StringVar(value=str(getattr(self, 'modules_per_string_default', 28)))
        ttk.Spinbox(settings_inner, from_=1, to=100, textvariable=self.modules_per_string_var, width=6).pack(side='left', padx=(0, 15))
        
        ttk.Label(settings_inner, text="Row Spacing (ft):").pack(side='left', padx=(0, 5))
        self.row_spacing_var = tk.StringVar(value=str(getattr(self, 'row_spacing_default', 20.0)))
        ttk.Spinbox(settings_inner, from_=1, to=100, textvariable=self.row_spacing_var, width=6).pack(side='left')
        
        # Calculate button
        calc_btn = ttk.Button(bottom_frame, text="Calculate Estimate", command=self.calculate_estimate)
        calc_btn.pack(anchor='w', pady=(0, 10))
        
        # Results frame
        results_frame = ttk.LabelFrame(bottom_frame, text="Estimated BOM (Rolled-Up Totals)", padding="10")
        results_frame.pack(fill='both', expand=True)
        
        # Results treeview
        columns = ('item', 'quantity', 'unit')
        self.results_tree = ttk.Treeview(results_frame, columns=columns, show='headings', height=8)
        self.results_tree.heading('item', text='Item')
        self.results_tree.heading('quantity', text='Quantity')
        self.results_tree.heading('unit', text='Unit')
        self.results_tree.column('item', width=250, anchor='w')
        self.results_tree.column('quantity', width=100, anchor='center')
        self.results_tree.column('unit', width=60, anchor='center')
        
        # Scrollbar for results
        results_scroll = ttk.Scrollbar(results_frame, orient='vertical', command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=results_scroll.set)
        
        self.results_tree.pack(side='left', fill='both', expand=True)
        results_scroll.pack(side='right', fill='y')

    def setup_tree_panel(self, parent):
        """Setup the left tree view panel"""
        # Tree frame with label
        tree_label = ttk.Label(parent, text="Project Structure", font=('Helvetica', 10, 'bold'))
        tree_label.pack(anchor='w', pady=(0, 5))
        
        # Tree view
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill='both', expand=True)
        
        self.tree = ttk.Treeview(tree_frame, show='tree', selectmode='browse')
        self.tree.pack(side='left', fill='both', expand=True)
        
        # Scrollbar
        tree_scroll = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        tree_scroll.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=tree_scroll.set)
        
        # Bind selection event
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        
        # Buttons frame
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill='x', pady=(10, 0))
        
        add_subarray_btn = ttk.Button(btn_frame, text="+ Subarray", command=lambda: self.add_subarray(auto_add_block=True))
        add_subarray_btn.pack(side='left', padx=(0, 5))
        
        add_block_btn = ttk.Button(btn_frame, text="+ Block", command=self.add_block_to_selected)
        add_block_btn.pack(side='left', padx=(0, 5))
        
        copy_block_btn = ttk.Button(btn_frame, text="Copy Block", command=self.copy_selected_block)
        copy_block_btn.pack(side='left', padx=(0, 5))
        
        delete_btn = ttk.Button(btn_frame, text="Delete", command=self.delete_selected_item)
        delete_btn.pack(side='left')

    def add_block_to_selected(self):
        """Add a block to the currently selected subarray"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item_id = selection[0]
        parent_id = self.tree.parent(item_id)
        
        # Determine which subarray to add to
        if parent_id == '':
            # Selected item is a subarray
            subarray_id = item_id
        else:
            # Selected item is a block, get its parent subarray
            subarray_id = parent_id
        
        if subarray_id in self.subarrays:
            self.add_block(subarray_id)

    def copy_selected_block(self):
        """Copy the currently selected block with all its configurations"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item_id = selection[0]
        parent_id = self.tree.parent(item_id)
        
        # Only copy if a block is selected (not a subarray)
        if parent_id == '':
            # Selected item is a subarray, not a block
            return
        
        subarray_id = parent_id
        source_block_id = item_id
        
        if subarray_id not in self.subarrays:
            return
        if source_block_id not in self.subarrays[subarray_id]['blocks']:
            return
        
        # Get source block data
        source_block = self.subarrays[subarray_id]['blocks'][source_block_id]
        
        # Create new block ID
        new_block_id = self.generate_id("block")
        
        # Deep copy the tracker configurations
        import copy
        copied_trackers = copy.deepcopy(source_block['trackers'])
        
        # Create new block with copied data
        new_block_name = self.get_next_block_name(subarray_id, source_block['name'])
        
        self.subarrays[subarray_id]['blocks'][new_block_id] = {
            'name': new_block_name,
            'type': source_block['type'],
            'num_rows': source_block['num_rows'],
            'trackers': copied_trackers
        }
        
        # Add to tree view under the same subarray
        self.tree.insert(subarray_id, 'end', new_block_id, text=new_block_name)
        
        # Select the new block
        self.tree.selection_set(new_block_id)
        self.on_tree_select(None)

    def setup_details_panel(self, parent):
        """Setup the right details panel"""
        # Details label
        self.details_label = ttk.Label(parent, text="Details", font=('Helvetica', 10, 'bold'))
        self.details_label.pack(anchor='w', pady=(0, 5))
        
        # Container for dynamic content
        self.details_container = ttk.Frame(parent)
        self.details_container.pack(fill='both', expand=True)
        
        # Placeholder
        self.placeholder_label = ttk.Label(self.details_container, text="Select a subarray or block to view details", foreground='gray')
        self.placeholder_label.pack(pady=20)

    def clear_details_panel(self):
        """Clear the details panel"""
        for widget in self.details_container.winfo_children():
            widget.destroy()
        
        self.placeholder_label = ttk.Label(self.details_container, text="Select a subarray or block to view details", foreground='gray')
        self.placeholder_label.pack(pady=20)
        
        self.details_label.config(text="Details")

    def on_tree_select(self, event):
        """Handle tree selection changes"""
        selection = self.tree.selection()
        if not selection:
            self.clear_details_panel()
            return
        
        item_id = selection[0]
        parent_id = self.tree.parent(item_id)
        
        # Clear existing details
        for widget in self.details_container.winfo_children():
            widget.destroy()
        
        if parent_id == '':
            # It's a subarray
            self.selected_item_id = item_id
            self.selected_item_type = 'subarray'
            self.show_subarray_details(item_id)
        else:
            # It's a block
            self.selected_item_id = item_id
            self.selected_item_type = 'block'
            self.show_block_details(parent_id, item_id)

    def show_subarray_details(self, subarray_id: str):
        """Show details panel for a subarray"""
        if subarray_id not in self.subarrays:
            return
        
        subarray = self.subarrays[subarray_id]
        self.details_label.config(text=f"Subarray: {subarray['name']}")
        
        # Create form
        form_frame = ttk.Frame(self.details_container, padding="10")
        form_frame.pack(fill='x')
        
        # Subarray Name
        ttk.Label(form_frame, text="Name:").grid(row=0, column=0, sticky='w', pady=5)
        name_var = tk.StringVar(value=subarray['name'])
        name_entry = ttk.Entry(form_frame, textvariable=name_var, width=30)
        name_entry.grid(row=0, column=1, sticky='w', pady=5, padx=(10, 0))
        
        def update_name(*args):
            subarray['name'] = name_var.get()
            self.tree.item(subarray_id, text=name_var.get())
            self.details_label.config(text=f"Subarray: {name_var.get()}")
        name_var.trace_add('write', update_name)
        
        # Transformer MVA
        ttk.Label(form_frame, text="Transformer (MVA):").grid(row=1, column=0, sticky='w', pady=5)
        mva_var = tk.StringVar(value=str(subarray['transformer_mva']))
        mva_spinbox = ttk.Spinbox(form_frame, from_=0.5, to=100, increment=0.5, textvariable=mva_var, width=10)
        mva_spinbox.grid(row=1, column=1, sticky='w', pady=5, padx=(10, 0))
        
        def update_mva(*args):
            try:
                subarray['transformer_mva'] = float(mva_var.get())
            except ValueError:
                pass
        mva_var.trace_add('write', update_mva)
        
        # Summary info
        summary_frame = ttk.LabelFrame(self.details_container, text="Summary", padding="10")
        summary_frame.pack(fill='x', pady=(20, 0), padx=10)
        
        num_blocks = len(subarray['blocks'])
        total_trackers = sum(
            sum(t['quantity'] for t in block['trackers'])
            for block in subarray['blocks'].values()
        )
        
        ttk.Label(summary_frame, text=f"Blocks: {num_blocks}").pack(anchor='w')
        ttk.Label(summary_frame, text=f"Total Trackers: {total_trackers}").pack(anchor='w')

    def show_block_details(self, subarray_id: str, block_id: str):
        """Show details panel for a block"""
        if subarray_id not in self.subarrays:
            return
        if block_id not in self.subarrays[subarray_id]['blocks']:
            return
        
        block = self.subarrays[subarray_id]['blocks'][block_id]
        self.details_label.config(text=f"Block: {block['name']}")
        
        # Create scrollable frame for block details
        canvas = tk.Canvas(self.details_container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.details_container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        def _bind_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        def _unbind_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
        
        # Bind when mouse enters canvas, unbind when it leaves
        canvas.bind("<Enter>", _bind_mousewheel)
        canvas.bind("<Leave>", _unbind_mousewheel)
        scrollable_frame.bind("<Enter>", _bind_mousewheel)
        scrollable_frame.bind("<Leave>", _unbind_mousewheel)
        
        # Form frame
        form_frame = ttk.Frame(scrollable_frame, padding="10")
        form_frame.pack(fill='x')
        
        # Block Name
        ttk.Label(form_frame, text="Name:").grid(row=0, column=0, sticky='w', pady=5)
        name_var = tk.StringVar(value=block['name'])
        name_entry = ttk.Entry(form_frame, textvariable=name_var, width=25)
        name_entry.grid(row=0, column=1, sticky='w', pady=5, padx=(10, 0))
        
        def update_name(*args):
            block['name'] = name_var.get()
            self.tree.item(block_id, text=name_var.get())
            self.details_label.config(text=f"Block: {name_var.get()}")
        name_var.trace_add('write', update_name)
        
        # Block Type
        ttk.Label(form_frame, text="Type:").grid(row=1, column=0, sticky='w', pady=5)
        type_var = tk.StringVar(value=block['type'])
        type_combo = ttk.Combobox(form_frame, textvariable=type_var, values=['combiner', 'string_inverter'], state='readonly', width=15)
        type_combo.grid(row=1, column=1, sticky='w', pady=5, padx=(10, 0))
        self.disable_combobox_scroll(type_combo)
        
        def update_type(*args):
            block['type'] = type_var.get()
        type_var.trace_add('write', update_type)
        
        # Number of Rows
        ttk.Label(form_frame, text="Rows from Device:").grid(row=2, column=0, sticky='w', pady=5)
        rows_var = tk.StringVar(value=str(block['num_rows']))
        rows_spinbox = ttk.Spinbox(form_frame, from_=1, to=50, textvariable=rows_var, width=10)
        rows_spinbox.grid(row=2, column=1, sticky='w', pady=5, padx=(10, 0))
        ttk.Label(form_frame, text="(furthest row from combiner/string inverter)", font=('Helvetica', 8), foreground='gray').grid(row=2, column=2, sticky='w', padx=(5, 0))
        
        def update_rows(*args):
            try:
                block['num_rows'] = int(rows_var.get())
            except ValueError:
                pass
        rows_var.trace_add('write', update_rows)
        
        # Tracker Configurations Section
        tracker_frame = ttk.LabelFrame(scrollable_frame, text="Tracker Configurations", padding="10")
        tracker_frame.pack(fill='x', pady=(15, 0), padx=10)
        
        # Store references for tracker rows
        self.current_block_tracker_frame = tracker_frame
        self.current_block_id = block_id
        self.current_subarray_id = subarray_id
        self.current_block = block
        
        # String count display
        count_frame = ttk.Frame(tracker_frame)
        count_frame.pack(fill='x', pady=(0, 10))
        self.string_count_label = ttk.Label(count_frame, text="Total Strings: 0", font=('Helvetica', 10, 'bold'))
        self.string_count_label.pack(side='left')
        
        # Headers
        header_frame = ttk.Frame(tracker_frame)
        header_frame.pack(fill='x')
        ttk.Label(header_frame, text="Strings/Tracker", width=14).pack(side='left', padx=2)
        ttk.Label(header_frame, text="Quantity", width=10).pack(side='left', padx=2)
        ttk.Label(header_frame, text="Harness Config", width=14).pack(side='left', padx=2)
        ttk.Label(header_frame, text="", width=8).pack(side='left', padx=2)
        
        # Container for tracker rows
        self.tracker_rows_container = ttk.Frame(tracker_frame)
        self.tracker_rows_container.pack(fill='x', pady=(5, 0))
        
        # Add existing tracker rows
        for i, tracker in enumerate(block['trackers']):
            self.add_tracker_row_ui(block, i, tracker)
        
        # Update initial string count
        self.update_string_count()
        
        # Add tracker button
        add_btn = ttk.Button(tracker_frame, text="+ Add Tracker Type", 
                            command=lambda: self.add_new_tracker_to_block(block))
        add_btn.pack(anchor='w', pady=(10, 0))

    def add_tracker_row_ui(self, block: dict, index: int, tracker: dict):
        """Add a tracker configuration row to the UI"""
        row_frame = ttk.Frame(self.tracker_rows_container)
        row_frame.pack(fill='x', pady=2)
        
        # Strings dropdown
        strings_var = tk.StringVar(value=str(tracker['strings']))
        strings_combo = ttk.Combobox(row_frame, textvariable=strings_var, 
                                     values=["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16"], 
                                     width=12, state='readonly')
        strings_combo.pack(side='left', padx=2)
        self.disable_combobox_scroll(strings_combo)
        
        # Quantity
        qty_var = tk.StringVar(value=str(tracker['quantity']))
        qty_spinbox = ttk.Spinbox(row_frame, from_=0, to=10000, textvariable=qty_var, width=8)
        qty_spinbox.pack(side='left', padx=2)
        
        # Harness config
        harness_var = tk.StringVar(value=tracker['harness_config'])
        harness_options = self.get_harness_options(tracker['strings'])
        harness_combo = ttk.Combobox(row_frame, textvariable=harness_var, 
                                     values=harness_options, width=12)
        harness_combo.pack(side='left', padx=2)
        self.disable_combobox_scroll(harness_combo)
        
        # Update harness options when strings change
        def on_strings_change(*args):
            try:
                new_strings = int(strings_var.get())
                tracker['strings'] = new_strings
                new_options = self.get_harness_options(new_strings)
                harness_combo['values'] = new_options
                if harness_var.get() not in new_options:
                    harness_var.set(new_options[0])
                    tracker['harness_config'] = new_options[0]
            except ValueError:
                pass
            self.update_string_count()
        strings_var.trace_add('write', on_strings_change)
        
        # Update data when values change
        def on_qty_change(*args):
            try:
                tracker['quantity'] = int(qty_var.get())
            except ValueError:
                pass
            self.update_string_count()
        qty_var.trace_add('write', on_qty_change)
        
        def on_harness_change(*args):
            tracker['harness_config'] = harness_var.get()
        harness_var.trace_add('write', on_harness_change)
        
        # Remove button
        def remove_tracker():
            if tracker in block['trackers']:
                block['trackers'].remove(tracker)
            row_frame.destroy()
            self.update_string_count()
        
        remove_btn = ttk.Button(row_frame, text="✕", width=3, command=remove_tracker)
        remove_btn.pack(side='left', padx=2)

    def add_new_tracker_to_block(self, block: dict):
        """Add a new tracker configuration to a block"""
        new_tracker = {'strings': 2, 'quantity': 0, 'harness_config': '2'}
        block['trackers'].append(new_tracker)
        self.add_tracker_row_ui(block, len(block['trackers']) - 1, new_tracker)

    # ==================== Calculation Methods ====================

    def calculate_estimate(self):
        """Calculate and display the rolled-up BOM estimate"""
        # Clear previous results
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        # Aggregated totals
        totals = {
            'combiners': 0,
            'string_inverters': 0,
            'trackers_by_string': {},  # {num_strings: quantity}
            'harnesses_by_size': {},   # {num_strings: quantity}
            'whips_by_length': {},     # {length_ft: quantity}
            'extenders_short': 0,
            'extenders_long': 0,
            'total_whip_length': 0,
        }
        
        # Get module info for extender calculations
        try:
            module_width_mm = float(self.module_width_var.get())
            modules_per_string = int(self.modules_per_string_var.get())
            module_width_ft = module_width_mm / 304.8
            string_length_ft = module_width_ft * modules_per_string
        except ValueError:
            module_width_ft = 3.72  # Default ~1134mm
            modules_per_string = 28
            string_length_ft = module_width_ft * modules_per_string
        
        # Process each subarray and block
        for subarray_id, subarray in self.subarrays.items():
            for block_id, block in subarray['blocks'].items():
                # Count block types
                if block['type'] == 'combiner':
                    totals['combiners'] += 1
                else:
                    totals['string_inverters'] += 1
                
                # Process trackers in this block
                block_total_harnesses = 0
                
                for tracker in block['trackers']:
                    strings = tracker['strings']
                    qty = tracker['quantity']
                    harness_config = tracker['harness_config']
                    
                    if qty <= 0:
                        continue
                    
                    # Count trackers by string count
                    if strings not in totals['trackers_by_string']:
                        totals['trackers_by_string'][strings] = 0
                    totals['trackers_by_string'][strings] += qty
                    
                    # Count harnesses by size
                    harness_sizes = self.parse_harness_config(harness_config)
                    for size in harness_sizes:
                        if size not in totals['harnesses_by_size']:
                            totals['harnesses_by_size'][size] = 0
                        totals['harnesses_by_size'][size] += qty
                        block_total_harnesses += qty
                
                # Calculate whips for this block
                num_rows = block['num_rows']
                try:
                    row_spacing = float(self.row_spacing_var.get())
                except ValueError:
                    row_spacing = 20.0
                
                if num_rows > 0:
                    for row_num in range(1, num_rows + 1):
                        raw_whip_length = row_num * row_spacing
                        whip_length = self.round_whip_length(raw_whip_length)
                        whips_at_length = 2  # 2 whips per row (both sides of block)
                        
                        if whip_length not in totals['whips_by_length']:
                            totals['whips_by_length'][whip_length] = 0
                        totals['whips_by_length'][whip_length] += whips_at_length
                        
                        totals['total_whip_length'] += whip_length * whips_at_length
                
                # Calculate extenders for this block (2 per harness - short and long)
                totals['extenders_short'] += block_total_harnesses
                totals['extenders_long'] += block_total_harnesses
        
        # Calculate extender lengths
        short_extender_length = self.round_whip_length(10)  # Standard 10ft short side
        long_extender_length = self.round_whip_length(string_length_ft)
        
        # ==================== Display Results ====================
        
        # Combiner Boxes / String Inverters
        if totals['combiners'] > 0:
            self.results_tree.insert('', 'end', values=('Combiner Boxes', totals['combiners'], 'ea'))
        
        if totals['string_inverters'] > 0:
            self.results_tree.insert('', 'end', values=('String Inverters', totals['string_inverters'], 'ea'))
        
        # Trackers
        if totals['trackers_by_string']:
            self.results_tree.insert('', 'end', values=('--- TRACKERS ---', '', ''))
            total_trackers = 0
            for strings in sorted(totals['trackers_by_string'].keys()):
                qty = totals['trackers_by_string'][strings]
                total_trackers += qty
                self.results_tree.insert('', 'end', values=(f"{strings}-String Trackers", qty, 'ea'))
            self.results_tree.insert('', 'end', values=('Total Trackers', total_trackers, 'ea'))
        
        # Harnesses
        if totals['harnesses_by_size']:
            self.results_tree.insert('', 'end', values=('--- HARNESSES ---', '', ''))
            for size in sorted(totals['harnesses_by_size'].keys(), reverse=True):
                qty = totals['harnesses_by_size'][size]
                self.results_tree.insert('', 'end', values=(f"{size}-String Harness", qty, 'ea'))
        
        # Extenders
        total_extenders = totals['extenders_short'] + totals['extenders_long']
        if total_extenders > 0:
            self.results_tree.insert('', 'end', values=('--- EXTENDERS ---', '', ''))
            self.results_tree.insert('', 'end', values=(
                f"Extender {short_extender_length}ft (short side)", 
                totals['extenders_short'], 
                'ea'
            ))
            self.results_tree.insert('', 'end', values=(
                f"Extender {long_extender_length}ft (long side)", 
                totals['extenders_long'], 
                'ea'
            ))
        
        # Whips
        if totals['whips_by_length']:
            self.results_tree.insert('', 'end', values=('--- WHIPS ---', '', ''))
            for length in sorted(totals['whips_by_length'].keys()):
                qty = totals['whips_by_length'][length]
                self.results_tree.insert('', 'end', values=(f"Whip {length}ft", qty, 'ea'))
            self.results_tree.insert('', 'end', values=(
                'Total Whip Length', 
                f"{totals['total_whip_length']:,}", 
                'ft'
            ))


class QuickEstimateDialog(tk.Toplevel):
    """Dialog wrapper for Quick Estimate tool"""
    
    def __init__(self, parent, current_project=None, estimate_id=None, on_save=None):
        super().__init__(parent)
        self.title("Quick Estimate")
        self.current_project = current_project
        self.estimate_id = estimate_id
        self.on_save = on_save
        
        # Set dialog size
        self.geometry("1100x850")
        self.minsize(900, 700)
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
        # Create the Quick Estimate frame inside the dialog
        self.quick_estimate = QuickEstimate(
            self, 
            current_project=current_project,
            estimate_id=estimate_id,
            on_save=self._handle_save
        )
        self.quick_estimate.pack(fill='both', expand=True)
        
        # Add button frame at bottom
        button_frame = ttk.Frame(self)
        button_frame.pack(fill='x', padx=10, pady=10)
        
        # Save button
        save_btn = ttk.Button(button_frame, text="Save", command=self.save_and_close)
        save_btn.pack(side='right', padx=(5, 0))
        
        # Close button (save on close)
        close_btn = ttk.Button(button_frame, text="Close", command=self.save_and_close)
        close_btn.pack(side='right')
        
        # Center the dialog on the parent window
        self.center_on_parent(parent)
        
        # Focus on the dialog
        self.focus_set()
        
        # Handle window close button (X)
        self.protocol("WM_DELETE_WINDOW", self.save_and_close)
        
        # Wait for window to close before returning
        self.wait_window(self)
    
    def _handle_save(self):
        """Internal save handler"""
        if self.on_save:
            self.on_save()
    
    def save_and_close(self):
        """Save the estimate and close the dialog"""
        self.quick_estimate.save_estimate()
        self.destroy()
    
    def center_on_parent(self, parent):
        """Center the dialog on the parent window"""
        self.update_idletasks()
        
        # Get parent geometry
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        
        # Get dialog size
        dialog_width = self.winfo_width()
        dialog_height = self.winfo_height()
        
        # Calculate position
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        # Ensure dialog is on screen
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        x = max(0, min(x, screen_width - dialog_width))
        y = max(0, min(y, screen_height - dialog_height))
        
        self.geometry(f"+{x}+{y}")