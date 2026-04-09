import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, List, Optional, Set
import json
import re
from ..models.block import BlockConfig, DeviceType, WiringType
from ..models.device import HarnessConnection, CombinerBoxConfig
from ..utils.calculations import STANDARD_FUSE_SIZES

class DeviceConfigurator(ttk.Frame):
    """UI for configuring combiner boxes and other devices"""
    
    def __init__(self, parent, project_manager):
        super().__init__(parent)
        self.project_manager = project_manager
        self.current_project = None
        self.combiner_configs: Dict[str, CombinerBoxConfig] = {}
        self.edited_cells: Set[str] = set()  # Track manually edited cells
        self.data_source = 'blocks'  # 'blocks' or 'quick_estimate'
        
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
                                    "Red values indicate cable size mismatches with wiring configuration.")
        instructions.pack(anchor=tk.W)
        
        # NEC Safety Factor control frame
        nec_frame = ttk.Frame(main_frame)
        nec_frame.pack(fill=tk.X, pady=(10, 10))
        
        ttk.Label(nec_frame, text="NEC Safety Factor:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.nec_factor_var = tk.StringVar(value="1.56")
        self.nec_factor_entry = ttk.Entry(
            nec_frame,
            textvariable=self.nec_factor_var,
            width=6
        )
        self.nec_factor_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        # Bind Enter key and focus out to trigger update
        self.nec_factor_entry.bind('<Return>', lambda e: self.on_nec_factor_changed())
        self.nec_factor_entry.bind('<FocusOut>', lambda e: self.on_nec_factor_changed())
        
        ttk.Label(nec_frame, text="(Default: 1.56 = 125% × 125%)", 
                 foreground='gray').pack(side=tk.LEFT)
        
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
        
        # Bind click for inline editing
        self.tree.bind('<Button-1>', self.on_click)
    
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
        
        # Data source toggle
        ttk.Separator(button_frame, orient='vertical').pack(side=tk.LEFT, padx=10, fill=tk.Y)
        ttk.Label(button_frame, text="Data Source:").pack(side=tk.LEFT, padx=(5, 2))
        
        self.data_source_var = tk.StringVar(value='blocks')
        ttk.Radiobutton(
            button_frame, text="Block/Wiring Config",
            variable=self.data_source_var, value='blocks',
            command=self._on_data_source_changed
        ).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(
            button_frame, text="Quick Estimate",
            variable=self.data_source_var, value='quick_estimate',
            command=self._on_data_source_changed
        ).pack(side=tk.LEFT, padx=2)
        
        # Expand/Collapse All buttons
        ttk.Button(button_frame, text="Collapse All", 
                command=self.collapse_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Expand All", 
                command=self.expand_all).pack(side=tk.LEFT, padx=5)
        
        # Add separator and whips control
        ttk.Separator(button_frame, orient='vertical').pack(side=tk.LEFT, padx=10, fill=tk.Y)
        
        # Add whips control
        ttk.Label(button_frame, text="Combiner Whips:").pack(side=tk.LEFT, padx=5)
        
        self.use_whips_var = tk.BooleanVar(value=True)  # Default to using whips
        whips_check = ttk.Checkbutton(
            button_frame,
            text="Use 3ft Whips",
            variable=self.use_whips_var,
            command=self.on_whips_changed
        )
        whips_check.pack(side=tk.LEFT, padx=5)
        
        # Warning label
        self.warning_label = ttk.Label(button_frame, text="", foreground='red')
        self.warning_label.pack(side=tk.RIGHT, padx=5)

    def collapse_all(self):
        """Collapse all combiner boxes in the tree"""
        for item in self.tree.get_children():
            self.tree.item(item, open=False)

    def expand_all(self):
        """Expand all combiner boxes in the tree"""
        for item in self.tree.get_children():
            self.tree.item(item, open=True)

    def on_nec_factor_changed(self):
        """Handle NEC safety factor change"""
        if not self.current_project:
            return
        
        try:
            new_factor = float(self.nec_factor_var.get())
            
            # Validate range
            if new_factor < 1.0 or new_factor > 2.0:
                return
            
            # Check if value actually changed
            current_factor = getattr(self.current_project, 'nec_safety_factor', 1.56)
            if abs(new_factor - current_factor) < 0.001:
                return  # No change, skip update
            
            # Update project
            self.current_project.nec_safety_factor = new_factor
            
            # Update existing combiner configs instead of regenerating
            for combiner_id, config in self.combiner_configs.items():
                for conn in config.connections:
                    # Update NEC factor
                    conn.nec_factor = new_factor
                    # Recalculate derived values
                    conn.harness_current = conn.num_strings * conn.module_isc * new_factor
                    conn.calculated_fuse_size = conn._calculate_fuse_size()
                    conn.calculated_cable_size = conn._calculate_cable_size()
                
                # Recalculate combiner totals
                config.calculate_totals()
                
                # Reapply fuse uniformity rule
                if config.connections:
                    max_fuse = max(c.calculated_fuse_size for c in config.connections)
                    for conn in config.connections:
                        conn.calculated_fuse_size = max_fuse
                        if conn.user_fuse_size and conn.user_fuse_size < max_fuse:
                            conn.user_fuse_size = None
                            conn.fuse_manually_set = False
            
            # Refresh display
            self.refresh_display()
            
            # Trigger autosave
            if hasattr(self, 'main_app') and hasattr(self.main_app, 'autosave_project'):
                self.main_app.autosave_project()
                
            self.status_var.set(f"NEC factor updated to {new_factor:.2f}")
            
        except tk.TclError:
            # Invalid value in spinbox, ignore
            pass

    def on_whips_changed(self):
        """Handle whips checkbox change"""
        use_whips = self.use_whips_var.get()
        
        # Update all combiner configs
        for combiner_id, config in self.combiner_configs.items():
            config.use_whips = use_whips
            config.whip_length_ft = 3 if use_whips else 0
        
        # Save to project
        self.save_configuration_to_project()
        
        # Refresh display (optional - whips don't show in the main tree)
        # self.refresh_display()
    
    def load_project(self, project):
        """Load a project and generate device configurations"""
        self.current_project = project
        self.combiner_configs.clear()
        self.edited_cells.clear()
        
        if not project:
            self.status_var.set("No project loaded")
            self.tree.delete(*self.tree.get_children())
            return
        
        # Set NEC factor from project (or default)
        nec_factor = getattr(project, 'nec_safety_factor', 1.56)
        self.nec_factor_var.set(f"{nec_factor:.2f}")
        
        # Restore data source setting
        saved_source = getattr(project, 'device_config_source', 'blocks')
        self.data_source = saved_source
        if hasattr(self, 'data_source_var'):
            self.data_source_var.set(saved_source)
        
        if saved_source == 'quick_estimate' and hasattr(project, 'device_configs') and project.device_configs:
            # Load saved QE-sourced configs directly (QE hasn't calculated yet at load time)
            self._load_qe_configs_from_saved(project.device_configs)
        else:
            # Generate configurations from blocks
            try:
                self.generate_combiner_configs()
            except Exception as e:
                import traceback
                print(f"[DeviceConfig load_project] generate_combiner_configs FAILED: {e}")
                traceback.print_exc()

            # Load saved device configurations if they exist
            if hasattr(project, 'device_configs') and project.device_configs:
                if not self.combiner_configs:
                    self._load_qe_configs_from_saved(project.device_configs)
                else:
                    self.load_saved_configurations(project.device_configs)

        # Update display
        self.refresh_display()
        
        # Update status
        source_label = "Quick Estimate" if saved_source == 'quick_estimate' else "Block/Wiring Config"
        combiner_count = len(self.combiner_configs)
        self.status_var.set(f"Loaded {combiner_count} combiner box(es) from {source_label}")

    def _load_qe_configs_from_saved(self, saved_configs):
        """Rebuild CombinerBoxConfig objects from saved QE-sourced device configs."""
        self.combiner_configs.clear()
        self.edited_cells.clear()
        
        for combiner_id, saved_config in saved_configs.items():
            connections = []
            for saved_conn in saved_config.get('connections', []):
                connection = HarnessConnection(
                    block_id=saved_conn.get('block_id', 'QE'),
                    tracker_id=saved_conn['tracker_id'],
                    harness_id=saved_conn['harness_id'],
                    num_strings=saved_conn['num_strings'],
                    module_isc=saved_conn['module_isc'],
                    nec_factor=saved_conn['nec_factor'],
                    actual_cable_size=saved_conn.get('actual_cable_size', '10 AWG'),
                )
                
                # Restore user overrides
                if saved_conn.get('user_fuse_size'):
                    connection.user_fuse_size = saved_conn['user_fuse_size']
                    connection.fuse_manually_set = saved_conn.get('fuse_manually_set', True)
                    if connection.fuse_manually_set:
                        cell_id = f"{combiner_id}_{connection.tracker_id}_{connection.harness_id}_fuse"
                        self.edited_cells.add(cell_id)
                
                if saved_conn.get('user_cable_size'):
                    connection.user_cable_size = saved_conn['user_cable_size']
                    connection.cable_manually_set = saved_conn.get('cable_manually_set', True)
                    if connection.cable_manually_set:
                        cell_id = f"{combiner_id}_{connection.tracker_id}_{connection.harness_id}_cable"
                        self.edited_cells.add(cell_id)
                
                connections.append(connection)
            
            config = CombinerBoxConfig(
                combiner_id=combiner_id,
                block_id=saved_config.get('block_id', 'QE'),
                connections=connections,
                use_whips=saved_config.get('use_whips', True),
                whip_length_ft=saved_config.get('whip_length_ft', 3),
            )
            
            # Restore breaker overrides
            if saved_config.get('user_breaker_size'):
                config.user_breaker_size = saved_config['user_breaker_size']
                config.breaker_manually_set = saved_config.get('breaker_manually_set', False)
                if config.breaker_manually_set:
                    self.edited_cells.add(f"{combiner_id}_breaker")
            
            # Apply fuse uniformity rule
            if config.connections:
                max_fuse = max(conn.calculated_fuse_size for conn in config.connections)
                for conn in config.connections:
                    conn.calculated_fuse_size = max_fuse
            
            self.combiner_configs[combiner_id] = config

    def load_saved_configurations(self, saved_configs: Dict[str, dict]):
        """Load saved device configurations"""
        for combiner_id, saved_config in saved_configs.items():
            if combiner_id not in self.combiner_configs:
                continue
            
            config = self.combiner_configs[combiner_id]
            
            # Update breaker settings
            if 'user_breaker_size' in saved_config:
                config.user_breaker_size = saved_config.get('user_breaker_size')
                config.breaker_manually_set = saved_config.get('breaker_manually_set', False)
                if config.breaker_manually_set:
                    self.edited_cells.add(f"{combiner_id}_breaker")

            # Update whips settings
            if 'use_whips' in saved_config:
                config.use_whips = saved_config.get('use_whips', True)
                config.whip_length_ft = saved_config.get('whip_length_ft', 3)
                
                # Update UI checkbox if it exists
                if hasattr(self, 'use_whips_var'):
                    self.use_whips_var.set(config.use_whips)
            
            # Update connections
            saved_connections = saved_config.get('connections', [])
            for i, saved_conn in enumerate(saved_connections):
                if i >= len(config.connections):
                    continue
                
                conn = config.connections[i]
                
                # Update user overrides
                if saved_conn.get('user_fuse_size'):
                    conn.user_fuse_size = saved_conn['user_fuse_size']
                    conn.fuse_manually_set = saved_conn.get('fuse_manually_set', True)
                    if conn.fuse_manually_set:
                        cell_id = f"{combiner_id}_{conn.tracker_id}_{conn.harness_id}_fuse"
                        self.edited_cells.add(cell_id)
                
                if saved_conn.get('user_cable_size'):
                    conn.user_cable_size = saved_conn['user_cable_size']
                    conn.cable_manually_set = saved_conn.get('cable_manually_set', True)
                    if conn.cable_manually_set:
                        cell_id = f"{combiner_id}_{conn.tracker_id}_{conn.harness_id}_cable"
                        self.edited_cells.add(cell_id)

    def _on_data_source_changed(self):
        """Handle data source toggle between Block/Wiring Config and Quick Estimate."""
        new_source = self.data_source_var.get()
        
        if new_source == self.data_source:
            return  # No change
        
        # Warn if there are manual edits
        if self.edited_cells:
            if not messagebox.askyesno(
                "Switch Data Source?",
                "Switching data sources will discard your manual edits.\n\nContinue?"
            ):
                # Revert the radio button
                self.data_source_var.set(self.data_source)
                return
        
        self.data_source = new_source
        
        # Always persist the source preference
        if self.current_project:
            self.current_project.device_config_source = new_source
        
        if new_source == 'quick_estimate':
            self.load_from_quick_estimate()
        else:
            # Reload from Block/Wiring Configurator
            self.combiner_configs.clear()
            self.edited_cells.clear()
            if self.current_project:
                try:
                    self.generate_combiner_configs()
                except (AttributeError, TypeError) as e:
                    import traceback
                    print(f"[DeviceConfig] generate_combiner_configs failed: {e}")
                    traceback.print_exc()
                if hasattr(self.current_project, 'device_configs') and self.current_project.device_configs:
                    # If we have saved configs and generate failed, load from saved
                    if not self.combiner_configs:
                        self._load_qe_configs_from_saved(self.current_project.device_configs)
                    else:
                        self.load_saved_configurations(self.current_project.device_configs)
                self.refresh_display()
                combiner_count = len(self.combiner_configs)
                self.status_var.set(f"Loaded {combiner_count} combiner box(es) from Block/Wiring Config")
    
    def load_from_quick_estimate(self):
        """Load combiner box configurations from Quick Estimate data."""
        # Get the QE widget via main_app
        qe_widget = None
        if hasattr(self, 'main_app') and hasattr(self.main_app, 'quick_estimate_widget'):
            qe_widget = self.main_app.quick_estimate_widget
        
        if not qe_widget:
            messagebox.showwarning(
                "Not Available",
                "Quick Estimate is not available. Please run a Quick Estimate first."
            )
            return
        
        assignments = getattr(qe_widget, 'last_combiner_assignments', [])
        if not assignments:
            messagebox.showinfo(
                "No Data",
                "No combiner assignments found.\n\n"
                "Please run Calculate Estimate in the Quick Estimate tab first\n"
                "(using Centralized String or Central Inverter topology)."
            )
            return
        
        # Clear existing
        self.combiner_configs.clear()
        self.edited_cells.clear()
        
        # NEC factor — use project setting if available
        nec_factor = 1.56
        if self.current_project:
            nec_factor = getattr(self.current_project, 'nec_safety_factor', 1.56)
            self.nec_factor_var.set(f"{nec_factor:.2f}")
        
        # Convert each assignment into a CombinerBoxConfig
        for cb_data in assignments:
            cb_name = cb_data['combiner_name']
            combiner_id = cb_name  # Use the display name as the ID
            
            connections = []
            for conn_data in cb_data['connections']:
                connection = HarnessConnection(
                    block_id='QE',  # Mark as Quick Estimate sourced
                    tracker_id=conn_data['tracker_label'],
                    harness_id=conn_data['harness_label'],
                    num_strings=conn_data['num_strings'],
                    module_isc=conn_data['module_isc'],
                    nec_factor=conn_data['nec_factor'],
                    actual_cable_size=conn_data.get('wire_gauge', '10 AWG'),
                )
                connections.append(connection)
            
            combiner_config = CombinerBoxConfig(
                combiner_id=combiner_id,
                block_id='QE',
                connections=connections,
                use_whips=self.use_whips_var.get(),
                whip_length_ft=3 if self.use_whips_var.get() else 0,
            )
            
            # Let the calculated breaker size stand — user can override manually
            # (QE breaker dropdown is a global default, not per-CB)
            
            # Apply fuse uniformity rule
            if combiner_config.connections:
                max_fuse = max(conn.calculated_fuse_size for conn in combiner_config.connections)
                for conn in combiner_config.connections:
                    conn.calculated_fuse_size = max_fuse
                    if conn.user_fuse_size and conn.user_fuse_size < max_fuse:
                        conn.user_fuse_size = None
                        conn.fuse_manually_set = False
            
            self.combiner_configs[combiner_id] = combiner_config
        
        # Refresh display
        self.refresh_display()
        self.update_warnings()
        
        # Save to project
        self.save_configuration_to_project()
        
        # Update status and ensure toggle reflects source
        self.data_source = 'quick_estimate'
        if hasattr(self, 'data_source_var'):
            self.data_source_var.set('quick_estimate')
        
        total_connections = sum(len(c.connections) for c in self.combiner_configs.values())
        self.status_var.set(
            f"Loaded {len(self.combiner_configs)} combiner box(es) "
            f"with {total_connections} connections from Quick Estimate"
        )

    def generate_combiner_configs(self):
        """Generate combiner box configurations from project blocks"""
        if not self.current_project:
            return
        
        # Iterate through all blocks
        for block_id, block in self.current_project.blocks.items():
            # Skip if block is still a raw dict (not yet deserialized)
            if isinstance(block, dict):
                device_type_val = block.get('device_type', '')
                if device_type_val != DeviceType.COMBINER_BOX.value:
                    continue
                # Can't fully process a raw dict block, skip it
                continue
            
            # Skip if not a combiner box
            if block.device_type != DeviceType.COMBINER_BOX:
                continue
            
            # Create combiner config
            combiner_id = f"{block_id}"
            combiner_config = CombinerBoxConfig(
                combiner_id=combiner_id,
                block_id=block_id,
                connections=[],
                use_whips=True,  # Default to using whips
                whip_length_ft=3  # Default 3ft whips
            )
            
            # Get module Isc
            module_isc = 0
            if block.tracker_template and block.tracker_template.module_spec:
                module_isc = block.tracker_template.module_spec.isc
            
            # Process each tracker
            tracker_positions = sorted(block.tracker_positions, key=lambda t: (t.y, t.x))
            
            # Check wiring type - String Homerun creates individual string connections
            if block.wiring_config and block.wiring_config.wiring_type == WiringType.HOMERUN:
                # String Homerun: one connection per string per tracker (num_strings=1 each)
                for idx, tracker_pos in enumerate(tracker_positions):
                    tracker_id = f"T{idx+1:02d}"
                    
                    if not tracker_pos.template:
                        continue
                    
                    strings_in_tracker = int(tracker_pos.template.strings_per_tracker)
                    whip_size = block.wiring_config.whip_cable_size if block.wiring_config else "10 AWG"
                    
                    for s_idx in range(strings_in_tracker):
                        string_id = f"S{s_idx+1:02d}"
                        
                        connection = HarnessConnection(
                            block_id=block_id,
                            tracker_id=tracker_id,
                            harness_id=string_id,
                            num_strings=1,
                            module_isc=module_isc,
                            nec_factor=self.current_project.nec_safety_factor,
                            actual_cable_size=whip_size
                        )
                        
                        combiner_config.connections.append(connection)
            
            # Wire Harness mode: check if we have custom harness groupings
            elif hasattr(block.wiring_config, 'harness_groupings') and block.wiring_config.harness_groupings:
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
                            
                            # Get whip cable size - first check harness-specific, then fall back to block default
                            whip_size = getattr(harness, 'whip_cable_size', '')
                            if not whip_size:
                                whip_size = block.wiring_config.whip_cable_size if block.wiring_config else "8 AWG"
                            
                            connection = HarnessConnection(
                                block_id=block_id,
                                tracker_id=tracker_id,
                                harness_id=harness_id,
                                num_strings=len(harness.string_indices),
                                module_isc=module_isc,
                                nec_factor=self.current_project.nec_safety_factor,
                                actual_cable_size=whip_size
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
                            nec_factor=self.current_project.nec_safety_factor,
                            actual_cable_size=block.wiring_config.whip_cable_size if block.wiring_config else "8 AWG"
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
                            nec_factor=self.current_project.nec_safety_factor,
                            actual_cable_size=block.wiring_config.whip_cable_size if block.wiring_config else "8 AWG"
                        )
                        
                        combiner_config.connections.append(connection)
            
            # Recalculate totals
            combiner_config.calculate_totals()
                       
            # Apply fuse uniformity rule - all fuses should be the max required
            if combiner_config.connections:
                max_fuse = max(conn.calculated_fuse_size for conn in combiner_config.connections)
                for conn in combiner_config.connections:
                    conn.calculated_fuse_size = max_fuse
                    # Reset user override if it's less than the new max
                    if conn.user_fuse_size and conn.user_fuse_size < max_fuse:
                        conn.user_fuse_size = None
                        conn.fuse_manually_set = False
            
            # Add to configurations
            self.combiner_configs[combiner_id] = combiner_config
    
    def refresh_display(self):
        """Refresh the treeview display"""
        # Clear existing items
        self.tree.delete(*self.tree.get_children())
        
        # Add combiner boxes
        for combiner_id in sorted(self.combiner_configs.keys(), key=lambda x: [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', x)]):
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
    
    def on_click(self, event):
        """Handle click for inline editing"""
        # Identify the clicked cell
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
            
        # Get the item and column
        item = self.tree.identify('item', event.x, event.y)
        column = self.tree.identify('column', event.x, event.y)
        
        if not item or not column:
            return
        
        # Get column index
        col_idx = int(column.replace('#', '')) - 1
        
        # Only allow editing fuse size (col 6), cable size (col 7), and breaker size (col 9)
        if col_idx not in [6, 7, 9]:
            return
        
        # Check if it's a combiner header or connection row
        parent = self.tree.parent(item)
        
        # Handle breaker size editing (only on first child row)
        if col_idx == 9:
            if not parent:
                return  # Skip combiner headers
            
            # Only allow editing on first child
            children = list(self.tree.get_children(parent))
            if item != children[0]:
                return
            
            combiner_id = self.tree.item(parent, 'text')
            self.create_inline_breaker_combo(item, column, combiner_id)
            return
        
        # Handle fuse and cable size editing
        if not parent:
            return  # Skip combiner headers
        
        # Get combiner and connection info
        combiner_id = self.tree.item(parent, 'text')
        config = self.combiner_configs.get(combiner_id)
        if not config:
            return
        
        # Find connection index
        children = list(self.tree.get_children(parent))
        conn_idx = children.index(item)
        
        if conn_idx >= len(config.connections):
            return
        
        connection = config.connections[conn_idx]
        
        # Create inline editor based on column
        if col_idx == 6:  # Fuse size
            self.create_inline_fuse_combo(item, column, connection, combiner_id)
        elif col_idx == 7:  # Cable size
            self.create_inline_cable_combo(item, column, connection, combiner_id)

    def create_inline_fuse_combo(self, item, column, connection, combiner_id):
        """Create inline combobox for fuse size editing"""
        # Get the bounding box of the cell
        bbox = self.tree.bbox(item, column)
        if not bbox:
            return
        
        # Create combobox
        fuse_var = tk.StringVar(value=str(connection.get_display_fuse_size()))
        combo = ttk.Combobox(self.tree, textvariable=fuse_var, state='readonly', width=8)
        combo['values'] = [str(s) for s in STANDARD_FUSE_SIZES]
        
        # Position the combobox
        combo.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        
        def save_fuse(event=None):
            try:
                new_size = int(fuse_var.get())
                
                # Validate
                if new_size < connection.harness_current:
                    messagebox.showerror("Invalid Size", 
                                    f"Fuse size must be at least {connection.harness_current:.1f}A")
                    combo.destroy()
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
                self.save_configuration_to_project()
                combo.destroy()
                
            except ValueError:
                combo.destroy()
        
        def cancel(event=None):
            combo.destroy()
        
        combo.bind('<<ComboboxSelected>>', save_fuse)
        combo.bind('<Return>', save_fuse)
        combo.bind('<Escape>', cancel)
        combo.bind('<FocusOut>', cancel)
        combo.focus_set()
        combo.event_generate('<Button-1>')  # Open dropdown immediately

    def create_inline_cable_combo(self, item, column, connection, combiner_id):
        """Create inline combobox for cable size editing"""
        # Get the bounding box of the cell
        bbox = self.tree.bbox(item, column)
        if not bbox:
            return
        
        # Create combobox
        cable_var = tk.StringVar(value=connection.get_display_cable_size())
        combo = ttk.Combobox(self.tree, textvariable=cable_var, state='readonly', width=10)
        combo['values'] = ["14 AWG", "12 AWG", "10 AWG", "8 AWG", "6 AWG", "4 AWG", "2 AWG", "1/0 AWG", "2/0 AWG", "4/0 AWG"]
        
        # Position the combobox
        combo.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        
        def save_cable(event=None):
            new_size = cable_var.get()
            
            # Validate ampacity
            from ..utils.cable_sizing import get_cable_ampacity
            ampacity = get_cable_ampacity(new_size)
            required_ampacity = connection.get_display_fuse_size()

            if ampacity < required_ampacity:
                messagebox.showerror("Invalid Size", 
                                    f"Cable ampacity ({ampacity}A) must be at least "
                                    f"equal to fuse rating ({required_ampacity}A).")
                combo.destroy()
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
            self.save_configuration_to_project()
            combo.destroy()
        
        def cancel(event=None):
            combo.destroy()
        
        combo.bind('<<ComboboxSelected>>', save_cable)
        combo.bind('<Return>', save_cable)
        combo.bind('<Escape>', cancel)
        combo.bind('<FocusOut>', cancel)
        combo.focus_set()
        combo.event_generate('<Button-1>')  # Open dropdown immediately

    def create_inline_breaker_combo(self, item, column, combiner_id):
        """Create inline combobox for breaker size editing"""
        config = self.combiner_configs.get(combiner_id)
        if not config:
            return
        
        # Get the bounding box of the cell
        bbox = self.tree.bbox(item, column)
        if not bbox:
            return
        
        # Create combobox
        from ..utils.calculations import STANDARD_BREAKER_SIZES
        breaker_var = tk.StringVar(value=str(config.get_display_breaker_size()))
        combo = ttk.Combobox(self.tree, textvariable=breaker_var, state='readonly', width=10)
        combo['values'] = [str(s) for s in STANDARD_BREAKER_SIZES]
        
        # Position the combobox
        combo.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        
        def save_breaker(event=None):
            try:
                new_size = int(breaker_var.get())
                
                # Validate
                if new_size < config.total_input_current:
                    messagebox.showerror("Invalid Size", 
                                    f"Breaker size must be at least {config.total_input_current:.1f}A")
                    combo.destroy()
                    return
                
                # Update config
                if new_size == config.calculated_breaker_size:
                    config.user_breaker_size = None
                    config.breaker_manually_set = False
                    cell_id = f"{combiner_id}_breaker"
                    self.edited_cells.discard(cell_id)
                else:
                    config.user_breaker_size = new_size
                    config.breaker_manually_set = True
                    cell_id = f"{combiner_id}_breaker"
                    self.edited_cells.add(cell_id)
                
                self.refresh_display()
                self.save_configuration_to_project()
                combo.destroy()
                
            except ValueError:
                combo.destroy()
        
        def cancel(event=None):
            combo.destroy()
        
        combo.bind('<<ComboboxSelected>>', save_breaker)
        combo.bind('<Return>', save_breaker)
        combo.bind('<Escape>', cancel)
        combo.bind('<FocusOut>', cancel)
        combo.focus_set()
        combo.event_generate('<Button-1>')  # Open dropdown immediately

    def save_configuration_to_project(self):
        """Save device configuration to the current project"""
        if not self.current_project:
            return
        
        # Convert combiner configs to serializable format
        device_configs = {}
        for combiner_id, config in self.combiner_configs.items():
            connections = []
            for conn in config.connections:
                connections.append({
                    'block_id': conn.block_id,
                    'tracker_id': conn.tracker_id,
                    'harness_id': conn.harness_id,
                    'num_strings': conn.num_strings,
                    'module_isc': conn.module_isc,
                    'nec_factor': conn.nec_factor,
                    'actual_cable_size': conn.actual_cable_size,
                    'calculated_fuse_size': conn.calculated_fuse_size,
                    'user_fuse_size': conn.user_fuse_size,
                    'calculated_cable_size': conn.calculated_cable_size,
                    'user_cable_size': conn.user_cable_size,
                    'fuse_manually_set': conn.fuse_manually_set,
                    'cable_manually_set': conn.cable_manually_set
                })
            
            device_configs[combiner_id] = {
                'combiner_id': config.combiner_id,
                'block_id': config.block_id,
                'connections': connections,
                'calculated_breaker_size': config.calculated_breaker_size,
                'user_breaker_size': config.user_breaker_size,
                'breaker_manually_set': config.breaker_manually_set,
                'total_input_current': config.total_input_current,
                'use_whips': getattr(config, 'use_whips', True),
                'whip_length_ft': getattr(config, 'whip_length_ft', 3)
            }
        
        # Save to project (include data source setting)
        self.current_project.device_configs = device_configs
        self.current_project.device_config_source = self.data_source
        
        # Trigger autosave if available
        if hasattr(self, 'main_app') and hasattr(self.main_app, 'autosave_project'):
            self.main_app.autosave_project()

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