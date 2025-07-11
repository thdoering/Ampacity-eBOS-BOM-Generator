import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, List, Optional, Set
import json
from ..models.block import BlockConfig, DeviceType
from ..models.device import HarnessConnection, CombinerBoxConfig
from ..utils.calculations import STANDARD_FUSE_SIZES, calculate_required_cable_size_for_current

class DeviceConfigurator(ttk.Frame):
    """UI for configuring combiner boxes and other devices"""
    
    def __init__(self, parent, project_manager):
        super().__init__(parent)
        self.project_manager = project_manager
        self.current_project = None
        self.combiner_configs: Dict[str, CombinerBoxConfig] = {}
        self.edited_cells: Set[str] = set()  # Track manually edited cells
        
        self.setup_ui()
        
    def setup_ui(self):
        """Set up the user interface"""
        # Main container
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Title
        title_label = ttk.Label(main_frame, text="Device Configuration", 
                               font=('TkDefaultFont', 16, 'bold'))
        title_label.pack(pady=(0, 10))
        
        # Instructions
        instructions = ttk.Label(main_frame, 
                               text="Configure fuses, cables, and breakers for combiner boxes. "
                                    "Red values indicate cable size mismatches with wiring configuration.",
                               wraplength=800)
        instructions.pack(pady=(0, 10))
        
        # Create treeview for the table
        self.create_treeview(main_frame)
        
        # Control buttons
        self.create_control_buttons(main_frame)
        
        # Status bar
        self.status_var = tk.StringVar(value="No project loaded")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.pack(fill=tk.X, pady=(10, 0))
    
    def create_treeview(self, parent):
        """Create the treeview table"""
        # Frame for treeview and scrollbars
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        # Create treeview
        columns = ('Tracker', 'Harness', '# Strings', 'Module Isc', 'NEC Factor', 
                  'Harness Current', 'Fuse Size', 'Cable Size', 'Total Current', 'Breaker Size')
        
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='tree headings',
                                yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Configure scrollbars
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        
        # Configure column headings
        self.tree.heading('#0', text='Combiner', anchor='center')
        self.tree.heading('Tracker', text='Tracker', anchor='center')
        self.tree.heading('Harness', text='Harness', anchor='center')
        self.tree.heading('# Strings', text='# Strings', anchor='center')
        self.tree.heading('Module Isc', text='Module Isc', anchor='center')
        self.tree.heading('NEC Factor', text='NEC Safety Factor', anchor='center')
        self.tree.heading('Harness Current', text='Harness Current', anchor='center')
        self.tree.heading('Fuse Size', text='Fuse Size', anchor='center')
        self.tree.heading('Cable Size', text='Cable Size', anchor='center')
        self.tree.heading('Total Current', text='Total Current', anchor='center')
        self.tree.heading('Breaker Size', text='Breaker Size', anchor='center')
        
        # Configure column widths
        self.tree.column('#0', width=100, stretch=False)
        self.tree.column('Tracker', width=70, stretch=False)
        self.tree.column('Harness', width=70, stretch=False)
        self.tree.column('# Strings', width=80, stretch=False)
        self.tree.column('Module Isc', width=90, stretch=False)
        self.tree.column('NEC Factor', width=120, stretch=False)
        self.tree.column('Harness Current', width=120, stretch=False)
        self.tree.column('Fuse Size', width=80, stretch=False)
        self.tree.column('Cable Size', width=90, stretch=False)
        self.tree.column('Total Current', width=100, stretch=False)
        self.tree.column('Breaker Size', width=100, stretch=False)
        
        # Configure column alignments - center everything
        self.tree.column('#0', anchor='center')  # Combiner column
        for col in columns:
            self.tree.column(col, anchor='center')
        
        # Configure tags for styling
        self.tree.tag_configure('mismatch', foreground='red')
        self.tree.tag_configure('edited', background='#ffffcc')  # Light yellow for edited cells
        self.tree.tag_configure('warning', foreground='orange')
        self.tree.tag_configure('combiner_header', font=('TkDefaultFont', 10, 'bold'))
        
        # Pack treeview and scrollbars
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Bind double-click for editing
        self.tree.bind('<Double-Button-1>', self.on_double_click)
    
    def create_control_buttons(self, parent):
        """Create control buttons"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Refresh button
        ttk.Button(button_frame, text="Refresh", 
                  command=self.refresh_display).pack(side=tk.LEFT, padx=5)
        
        # Reset to calculated button
        ttk.Button(button_frame, text="Reset Selected to Calculated", 
                  command=self.reset_selected).pack(side=tk.LEFT, padx=5)
        
        # Export button
        ttk.Button(button_frame, text="Export Configuration", 
                  command=self.export_configuration).pack(side=tk.LEFT, padx=5)
        
        # Warning label
        self.warning_label = ttk.Label(button_frame, text="", foreground='red')
        self.warning_label.pack(side=tk.RIGHT, padx=5)
    
    def load_project(self, project):
        """Load a project and generate device configurations"""
        self.current_project = project
        self.combiner_configs.clear()
        self.edited_cells.clear()
        
        if not project:
            self.status_var.set("No project loaded")
            self.tree.delete(*self.tree.get_children())
            return
        
        # Generate configurations for all combiner boxes
        self.generate_combiner_configs()
        
        # Update display
        self.refresh_display()
        
        # Update status
        combiner_count = len(self.combiner_configs)
        self.status_var.set(f"Loaded {combiner_count} combiner box(es)")
    
    def generate_combiner_configs(self):
        """Generate combiner box configurations from project blocks"""
        if not self.current_project:
            return
        
        # Iterate through all blocks
        for block_id, block in self.current_project.blocks.items():
            # Skip if not a combiner box
            if block.device_type != DeviceType.COMBINER_BOX:
                continue
            
            # Create combiner config
            combiner_id = f"{block_id}"
            combiner_config = CombinerBoxConfig(
                combiner_id=combiner_id,
                block_id=block_id,
                connections=[]
            )
            
            # Get module Isc
            module_isc = 0
            if block.tracker_template and block.tracker_template.module_spec:
                module_isc = block.tracker_template.module_spec.isc
            
            # Process each tracker
            tracker_positions = sorted(block.tracker_positions, key=lambda t: (t.y, t.x))
            
            # Check if we have custom harness groupings
            if hasattr(block.wiring_config, 'harness_groupings') and block.wiring_config.harness_groupings:
                # Process each tracker
                for idx, tracker_pos in enumerate(tracker_positions):
                    tracker_id = f"T{idx+1:02d}"
                    
                    if not tracker_pos.template:
                        continue
                    
                    strings_in_tracker = tracker_pos.template.strings_per_tracker
                    
                    # Look for harness configuration for this tracker's string count
                    if strings_in_tracker in block.wiring_config.harness_groupings:
                        # Use the harness configuration for this string count
                        harness_list = block.wiring_config.harness_groupings[strings_in_tracker]
                        
                        # Create connections for each harness
                        for h_idx, harness in enumerate(harness_list):
                            harness_id = f"H{h_idx+1:02d}"
                            
                            connection = HarnessConnection(
                                block_id=block_id,
                                tracker_id=tracker_id,
                                harness_id=harness_id,
                                num_strings=len(harness.string_indices),
                                module_isc=module_isc,
                                actual_cable_size=getattr(harness, 'cable_size', 
                                                        block.wiring_config.harness_cable_size)
                            )
                            
                            combiner_config.connections.append(connection)
                    else:
                        # No harness configuration for this string count - create default
                        connection = HarnessConnection(
                            block_id=block_id,
                            tracker_id=tracker_id,
                            harness_id="H01",
                            num_strings=strings_in_tracker,
                            module_isc=module_isc,
                            actual_cable_size=block.wiring_config.harness_cable_size
                        )
                        
                        combiner_config.connections.append(connection)
            else:
                # No custom harness groupings - create default connections
                for idx, tracker_pos in enumerate(tracker_positions):
                    tracker_id = f"T{idx+1:02d}"
                    
                    if tracker_pos.template:
                        connection = HarnessConnection(
                            block_id=block_id,
                            tracker_id=tracker_id,
                            harness_id="H01",
                            num_strings=tracker_pos.template.strings_per_tracker,
                            module_isc=module_isc,
                            actual_cable_size=block.wiring_config.harness_cable_size if block.wiring_config else "8 AWG"
                        )
                        
                        combiner_config.connections.append(connection)
            
            # Recalculate totals
            combiner_config.calculate_totals()
            
            # Store config
            self.combiner_configs[combiner_id] = combiner_config
    
    def refresh_display(self):
        """Refresh the treeview display"""
        # Clear existing items
        self.tree.delete(*self.tree.get_children())
        
        # Add combiner boxes
        for combiner_id in sorted(self.combiner_configs.keys()):
            config = self.combiner_configs[combiner_id]
            
            # Add combiner box as parent item
            combiner_item = self.tree.insert('', 'end', text=combiner_id, 
                                           tags=('combiner_header',))
            
            # Add connections as child items
            for conn in config.connections:
                tags = []
                
                # Check for cable mismatch
                if conn.is_cable_size_mismatch():
                    tags.append('mismatch')
                
                # Check for edited cells
                cell_ids = [
                    f"{combiner_id}_{conn.tracker_id}_{conn.harness_id}_fuse",
                    f"{combiner_id}_{conn.tracker_id}_{conn.harness_id}_cable"
                ]
                if any(cell_id in self.edited_cells for cell_id in cell_ids):
                    tags.append('edited')
                
                # Check for warnings (fuse > 90A)
                if conn.get_display_fuse_size() > 90:
                    tags.append('warning')
                
                # Format values
                values = (
                    conn.tracker_id,
                    conn.harness_id,
                    conn.num_strings,
                    f"{conn.module_isc:.2f}",
                    f"{conn.nec_factor}",
                    f"{conn.harness_current:.2f}",
                    f"{conn.get_display_fuse_size()}",
                    conn.get_display_cable_size(),
                    "",  # Total current only on first row
                    ""   # Breaker size only on first row
                )
                
                self.tree.insert(combiner_item, 'end', values=values, tags=tuple(tags))
            
            # Update first row with total current and breaker size
            children = self.tree.get_children(combiner_item)
            if children:
                first_child = children[0]
                values = list(self.tree.item(first_child, 'values'))
                values[8] = f"{config.total_input_current:.2f}"
                values[9] = f"{config.get_display_breaker_size()}"
                self.tree.item(first_child, values=values)
        
        # Expand all combiner boxes
        for item in self.tree.get_children():
            self.tree.item(item, open=True)
        
        # Update warnings
        self.update_warnings()
    
    def on_double_click(self, event):
        """Handle double-click for editing"""
        # Get clicked item and column
        item = self.tree.identify('item', event.x, event.y)
        column = self.tree.identify('column', event.x, event.y)
        
        if not item or not column:
            return
        
        # Check if it's a child item (connection)
        parent = self.tree.parent(item)
        if not parent:
            return  # Clicked on combiner header
        
        # Get column index
        col_idx = int(column.replace('#', '')) - 1
        
        # Only allow editing fuse size (col 6) and cable size (col 7)
        if col_idx not in [6, 7]:
            return
        
        # Get current values
        values = self.tree.item(item, 'values')
        combiner_id = self.tree.item(parent, 'text')
        
        # Find the connection
        config = self.combiner_configs.get(combiner_id)
        if not config:
            return
        
        # Determine which connection this is
        children = list(self.tree.get_children(parent))
        conn_idx = children.index(item)
        
        if conn_idx >= len(config.connections):
            return
        
        connection = config.connections[conn_idx]
        
        # Edit the value
        if col_idx == 6:  # Fuse size
            self.edit_fuse_size(connection, combiner_id)
        elif col_idx == 7:  # Cable size
            self.edit_cable_size(connection, combiner_id)
    
    def edit_fuse_size(self, connection: HarnessConnection, combiner_id: str):
        """Edit fuse size for a connection"""
        # Create dialog
        dialog = tk.Toplevel(self)
        dialog.title("Edit Fuse Size")
        dialog.geometry("300x150")
        
        # Center the dialog
        dialog.transient(self)
        dialog.grab_set()
        
        # Content
        ttk.Label(dialog, text=f"Tracker: {connection.tracker_id}, Harness: {connection.harness_id}").pack(pady=10)
        ttk.Label(dialog, text=f"Calculated: {connection.calculated_fuse_size}A").pack()
        
        # Fuse size selection
        frame = ttk.Frame(dialog)
        frame.pack(pady=10)
        
        ttk.Label(frame, text="Fuse Size:").grid(row=0, column=0, padx=5)
        
        fuse_var = tk.StringVar(value=str(connection.get_display_fuse_size()))
        fuse_combo = ttk.Combobox(frame, textvariable=fuse_var, state='readonly', width=10)
        fuse_combo['values'] = [str(s) for s in STANDARD_FUSE_SIZES]
        fuse_combo.grid(row=0, column=1, padx=5)
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        def save():
            try:
                new_size = int(fuse_var.get())
                
                # Validate
                if new_size < connection.harness_current:
                    messagebox.showerror("Invalid Size", 
                                       f"Fuse size must be at least {connection.harness_current:.1f}A "
                                       f"for the harness current.")
                    return
                
                # Update connection
                if new_size == connection.calculated_fuse_size:
                    connection.user_fuse_size = None
                    connection.fuse_manually_set = False
                    cell_id = f"{combiner_id}_{connection.tracker_id}_{connection.harness_id}_fuse"
                    self.edited_cells.discard(cell_id)
                else:
                    connection.user_fuse_size = new_size
                    connection.fuse_manually_set = True
                    cell_id = f"{combiner_id}_{connection.tracker_id}_{connection.harness_id}_fuse"
                    self.edited_cells.add(cell_id)
                
                # Recalculate cable size
                connection.calculated_cable_size = connection._calculate_cable_size()
                
                self.refresh_display()
                dialog.destroy()
                
            except ValueError:
                messagebox.showerror("Error", "Invalid fuse size")
        
        ttk.Button(button_frame, text="Save", command=save).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def edit_cable_size(self, connection: HarnessConnection, combiner_id: str):
        """Edit cable size for a connection"""
        # Create dialog
        dialog = tk.Toplevel(self)
        dialog.title("Edit Cable Size")
        dialog.geometry("350x200")
        
        # Center the dialog
        dialog.transient(self)
        dialog.grab_set()
        
        # Content
        ttk.Label(dialog, text=f"Tracker: {connection.tracker_id}, Harness: {connection.harness_id}").pack(pady=10)
        ttk.Label(dialog, text=f"Calculated: {connection.calculated_cable_size}").pack()
        ttk.Label(dialog, text=f"Wiring Config: {connection.actual_cable_size}", 
                 foreground='red' if connection.is_cable_size_mismatch() else 'black').pack()
        
        # Cable size selection
        frame = ttk.Frame(dialog)
        frame.pack(pady=10)
        
        ttk.Label(frame, text="Cable Size:").grid(row=0, column=0, padx=5)
        
        cable_var = tk.StringVar(value=connection.get_display_cable_size())
        cable_combo = ttk.Combobox(frame, textvariable=cable_var, state='readonly', width=12)
        cable_combo['values'] = ["10 AWG", "8 AWG", "6 AWG", "4 AWG"]
        cable_combo.grid(row=0, column=1, padx=5)
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        def save():
            new_size = cable_var.get()
            
            # Validate ampacity
            from ..utils.calculations import get_cable_ampacity
            ampacity = get_cable_ampacity(new_size)
            required_ampacity = connection.get_display_fuse_size()
            
            if ampacity < required_ampacity:
                messagebox.showerror("Invalid Size", 
                                   f"Cable ampacity ({ampacity}A) must be at least "
                                   f"equal to fuse rating ({required_ampacity}A).")
                return
            
            # Update connection
            if new_size == connection.calculated_cable_size:
                connection.user_cable_size = None
                connection.cable_manually_set = False
                cell_id = f"{combiner_id}_{connection.tracker_id}_{connection.harness_id}_cable"
                self.edited_cells.discard(cell_id)
            else:
                connection.user_cable_size = new_size
                connection.cable_manually_set = True
                cell_id = f"{combiner_id}_{connection.tracker_id}_{connection.harness_id}_cable"
                self.edited_cells.add(cell_id)
            
            self.refresh_display()
            dialog.destroy()
        
        ttk.Button(button_frame, text="Save", command=save).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def reset_selected(self):
        """Reset selected items to calculated values"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("No Selection", "Please select items to reset.")
            return
        
        for item in selection:
            parent = self.tree.parent(item)
            if not parent:
                continue  # Skip combiner headers
            
            # Get combiner and connection
            combiner_id = self.tree.item(parent, 'text')
            config = self.combiner_configs.get(combiner_id)
            if not config:
                continue
            
            # Find connection index
            children = list(self.tree.get_children(parent))
            conn_idx = children.index(item)
            
            if conn_idx >= len(config.connections):
                continue
            
            connection = config.connections[conn_idx]
            
            # Reset to calculated values
            connection.user_fuse_size = None
            connection.fuse_manually_set = False
            connection.user_cable_size = None
            connection.cable_manually_set = False
            
            # Remove from edited cells
            cell_ids = [
                f"{combiner_id}_{connection.tracker_id}_{connection.harness_id}_fuse",
                f"{combiner_id}_{connection.tracker_id}_{connection.harness_id}_cable"
            ]
            for cell_id in cell_ids:
                self.edited_cells.discard(cell_id)
        
        self.refresh_display()
    
    def update_warnings(self):
        """Update warning messages"""
        warnings = []
        
        # Check for high fuse ratings
        for config in self.combiner_configs.values():
            for conn in config.connections:
                if conn.get_display_fuse_size() > 90:
                    warnings.append(f"{config.combiner_id} {conn.tracker_id} {conn.harness_id}: "
                                  f"Fuse size {conn.get_display_fuse_size()}A exceeds 90A")
        
        # Update warning label
        if warnings:
            self.warning_label.config(text=f"⚠ {len(warnings)} warning(s)")
        else:
            self.warning_label.config(text="")
    
    def export_configuration(self):
        """Export device configuration to JSON"""
        if not self.combiner_configs:
            messagebox.showinfo("No Configuration", "No device configuration to export.")
            return
        
        # Prepare export data
        export_data = {}
        for combiner_id, config in self.combiner_configs.items():
            connections = []
            for conn in config.connections:
                connections.append({
                    'tracker_id': conn.tracker_id,
                    'harness_id': conn.harness_id,
                    'num_strings': conn.num_strings,
                    'module_isc': conn.module_isc,
                    'calculated_fuse_size': conn.calculated_fuse_size,
                    'user_fuse_size': conn.user_fuse_size,
                    'calculated_cable_size': conn.calculated_cable_size,
                    'user_cable_size': conn.user_cable_size,
                    'actual_cable_size': conn.actual_cable_size
                })
            
            export_data[combiner_id] = {
                'block_id': config.block_id,
                'connections': connections,
                'calculated_breaker_size': config.calculated_breaker_size,
                'user_breaker_size': config.user_breaker_size,
                'total_input_current': config.total_input_current
            }
        
        # Show in dialog for copying
        dialog = tk.Toplevel(self)
        dialog.title("Export Configuration")
        dialog.geometry("600x400")
        
        text_widget = tk.Text(dialog, wrap=tk.NONE)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Add JSON data
        json_str = json.dumps(export_data, indent=2)
        text_widget.insert('1.0', json_str)
        text_widget.config(state='disabled')
        
        # Copy button
        def copy_to_clipboard():
            self.clipboard_clear()
            self.clipboard_append(json_str)
            messagebox.showinfo("Copied", "Configuration copied to clipboard.")
        
        ttk.Button(dialog, text="Copy to Clipboard", 
                  command=copy_to_clipboard).pack(pady=10)