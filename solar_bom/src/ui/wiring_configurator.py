import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Dict, ClassVar, List
from ..models.block import BlockConfig, WiringType, CollectionPoint, WiringConfig, HarnessGroup
from ..models.module import ModuleOrientation
from ..models.tracker import TrackerPosition
from ..utils.calculations import get_ampacity_for_wire_gauge, calculate_nec_current, wire_harness_compatibility
from ..utils.cable_sizing import CABLE_SIZE_ORDER, calculate_all_cable_sizes
from ..utils.calculations import calculate_fuse_size

class CollapsibleFrame(ttk.Frame):
    """A collapsible frame widget with instant show/hide"""
    
    def __init__(self, parent, text="", start_collapsed=True, **kwargs):
        super().__init__(parent, **kwargs)
        
        self.is_collapsed = start_collapsed
        self.harness_tree_items = {}  # Maps tree item IDs to harness keys
        self.AWG_SIZES = { 
            "10 AWG": 10,
            "8 AWG": 8,
            "6 AWG": 6,
            "4 AWG": 4,
            "2 AWG": 2,
            "1/0 AWG": 0,
            "2/0 AWG": -1,
            "4/0 AWG": -3
        }

        self.harness_tree_items = {}  # Maps tree item IDs to harness keys
        self.harness_cable_edited_cells = set()  # Track edited cells

        # Header frame
        self.header_frame = ttk.Frame(self)
        self.header_frame.pack(fill=tk.X, padx=5, pady=2)
        
        # Toggle button
        self.toggle_btn = ttk.Label(self.header_frame, text="▶" if start_collapsed else "▼", 
                                   cursor="hand2", font=('TkDefaultFont', 10))
        self.toggle_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.toggle_btn.bind("<Button-1>", lambda e: self.toggle())
        
        # Title label
        self.title_label = ttk.Label(self.header_frame, text=text, font=('TkDefaultFont', 10, 'bold'))
        self.title_label.pack(side=tk.LEFT)
        self.title_label.bind("<Button-1>", lambda e: self.toggle())
        
        # Content frame
        self.content_frame = ttk.LabelFrame(self, text="", padding="5")
        
        if start_collapsed:
            # Start with content hidden
            self.content_frame.pack_forget()
        else:
            self.content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=2)
    
    def toggle(self):
        """Toggle the collapsed state"""
        if self.is_collapsed:
            self.expand()
        else:
            self.collapse()
    
    def expand(self):
        """Expand the frame"""
        if not self.is_collapsed:
            return
            
        self.is_collapsed = False
        self.toggle_btn.config(text="▼")
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=2)
    
    def collapse(self):
        """Collapse the frame"""
        if self.is_collapsed:
            return
            
        self.is_collapsed = True
        self.toggle_btn.config(text="▶")
        self.content_frame.pack_forget()

class WiringConfigurator(tk.Toplevel):

    # AWG to mm² conversion - Extended with more sizes
    AWG_SIZES: ClassVar[Dict[str, float]] = {
        "10 AWG": 5.26,
        "8 AWG": 8.37,
        "6 AWG": 13.30,
        "4 AWG": 21.15,
        "2 AWG": 33.62,
        "1/0 AWG": 53.49,
        "2/0 AWG": 67.43,
        "4/0 AWG": 107.22
    }

    # Standard fuse ratings in amps
    FUSE_RATINGS: ClassVar[List[int]] = [5, 10, 15, 20, 25, 30, 35, 40, 45]

    def __init__(self, parent, block: BlockConfig):
        super().__init__(parent)
        self.parent = parent
        self.block = block
        self.project = getattr(parent, 'current_project', None)

        self.parent_notify_blocks_changed = getattr(parent, '_notify_blocks_changed', None)

        self.scale_factor = 10.0  # Starting scale (10 pixels per meter)
        self.pan_x = 0  # Pan offset in pixels
        self.pan_y = 0
        self.panning = False
        self.selected_whips = set()  # Set of (tracker_id, harness_idx, polarity) tuples
        self.dragging = False
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.drag_whips = False
        self.selection_box = None
        self.routing_mode_var = tk.StringVar(value="realistic")
        self.extended_extender_points = {}  # Track extended extender points
        
        # Set up window properties
        self.title("Wiring Configuration")
        self.geometry("1800x1000")
        self.minsize(1000, 600)
        
        # Initialize UI
        self.setup_ui()
        
        # Make window modal
        self.transient(parent)
        # self.grab_set()
        
        # Position window relative to parent
        x = parent.winfo_rootx() + 50
        y = parent.winfo_rooty() - 75
        self.geometry(f"+{x}+{y}")

    def world_to_canvas(self, world_x, world_y):
        """Convert world coordinates to canvas coordinates"""
        scale = self.get_canvas_scale()
        canvas_x = 20 + self.pan_x + world_x * scale
        canvas_y = 20 + self.pan_y + world_y * scale
        return canvas_x, canvas_y

    def canvas_to_world(self, canvas_x, canvas_y):
        """Convert canvas coordinates to world coordinates"""
        scale = self.get_canvas_scale()
        world_x = (canvas_x - 20 - self.pan_x) / scale
        world_y = (canvas_y - 20 - self.pan_y) / scale
        return world_x, world_y

    def world_point_to_canvas(self, point):
        """Convert a single world coordinate point (x, y) to canvas coordinates"""
        return self.world_to_canvas(point[0], point[1])

    def world_route_to_canvas_points(self, route):
        """Convert a route (list of world coordinate points) to flat canvas coordinates list"""
        canvas_points = []
        for x, y in route:
            canvas_x, canvas_y = self.world_to_canvas(x, y)
            canvas_points.extend([canvas_x, canvas_y])
        return canvas_points
        
    def setup_ui(self):
        """Create and arrange UI components"""
        # Main container with padding
        main_container = ttk.Frame(self, padding="10")
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        main_container.grid_columnconfigure(1, weight=1)
        main_container.grid_rowconfigure(0, weight=1)
        
        # Create main sections
        self.setup_controls_frame(main_container)
        self.setup_canvas_frame(main_container)
        self.setup_bottom_buttons(main_container)
        
        # Update UI based on initial wiring type
        self.update_ui_for_wiring_type()
        
        # Initialize with existing configuration if available
        self.load_existing_configuration()

    def setup_controls_frame(self, main_container):
        """Set up the left side controls frame"""
        controls_frame = ttk.LabelFrame(main_container, text="Wiring Configuration", padding="5")
        controls_frame.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.setup_wiring_type_selection(controls_frame)
        self.setup_wiring_mode_selection(controls_frame)
        self.setup_cable_specifications(controls_frame)
        self.setup_harness_configuration(controls_frame)
        self.setup_routing_controls(controls_frame)
        self.setup_harness_cable_table(controls_frame)

    def setup_canvas_frame(self, main_container):
        """Set up the right side canvas frame"""
        canvas_frame = ttk.LabelFrame(main_container, text="Wiring Layout", padding="5")
        canvas_frame.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.canvas = tk.Canvas(canvas_frame, width=800, height=600, bg='white')
        self.canvas.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.setup_legend(canvas_frame)
        self.setup_warning_panel()
        self.setup_canvas_bindings()

    def setup_bottom_buttons(self, main_container):
        """Set up bottom buttons"""
        button_frame = ttk.Frame(main_container)
        button_frame.grid(row=1, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Apply", command=self.apply_configuration).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.cancel).grid(row=0, column=1, padx=5)

    def setup_wiring_type_selection(self, controls_frame):
        """Set up wiring type selection dropdown"""
        ttk.Label(controls_frame, text="Wiring Type:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.wiring_type_var = tk.StringVar(value=WiringType.HARNESS.value)
        # Initialize wiring mode variable
        # Get project reference - parent is BlockConfigurator
        project = getattr(self.parent, 'current_project', None)
        default_mode = 'daisy_chain'
        if project and hasattr(project, 'wiring_mode'):
            default_mode = project.wiring_mode
        self.wiring_mode_var = tk.StringVar(value=default_mode)
        wiring_type_combo = ttk.Combobox(controls_frame, textvariable=self.wiring_type_var, state='readonly')
        wiring_type_combo['values'] = [t.value for t in WiringType]
        wiring_type_combo.grid(row=0, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        wiring_type_combo.bind('<<ComboboxSelected>>', self.on_wiring_type_change)
    
    def setup_wiring_mode_selection(self, controls_frame):
        """Set up wiring mode selection (daisy-chain vs leapfrog)"""
        mode_frame = ttk.Frame(controls_frame)
        mode_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        ttk.Label(mode_frame, text="Wiring Mode:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        
        radio_frame = ttk.Frame(mode_frame)
        radio_frame.grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)
        
        ttk.Radiobutton(radio_frame, text="Daisy-chain", 
                       variable=self.wiring_mode_var,
                       value="daisy_chain",
                       command=self.on_wiring_mode_change).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(radio_frame, text="Leapfrog", 
                       variable=self.wiring_mode_var,
                       value="leapfrog",
                       command=self.on_wiring_mode_change).pack(side=tk.LEFT, padx=5)

    def setup_cable_specifications(self, controls_frame):
        """Set up cable specifications section"""
        # Create collapsible frame
        cable_collapsible = CollapsibleFrame(controls_frame, text="Default Cable Sizes", start_collapsed=True)
        cable_collapsible.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Store reference for visibility control
        self.cable_collapsible = cable_collapsible
        
        # Use the content_frame inside the collapsible
        cable_frame = cable_collapsible.content_frame

        # Add note about defaults
        note_text = "These sizes are used as defaults for new harness configurations"
        note_label = ttk.Label(cable_frame, text=note_text, font=('TkDefaultFont', 9, 'italic'))
        note_label.grid(row=0, column=0, columnspan=2, padx=5, pady=(0, 5), sticky=tk.W)

        # String Cable Size
        ttk.Label(cable_frame, text="String Cable Size:").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.string_cable_size_var = tk.StringVar(value="10 AWG")
        string_cable_combo = ttk.Combobox(cable_frame, textvariable=self.string_cable_size_var, state='readonly', width=10)
        string_cable_combo['values'] = list(self.AWG_SIZES.keys())
        string_cable_combo.grid(row=1, column=1, padx=5, pady=2)
        self.string_cable_size_var.trace('w', lambda *args: self.draw_wiring_layout())

        # Wire Harness Size
        self.harness_frame = ttk.Frame(cable_frame)
        self.harness_frame.grid(row=2, column=0, columnspan=2, padx=0, pady=5, sticky=(tk.W, tk.E))
        ttk.Label(self.harness_frame, text="Harness Cable Size:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.harness_cable_size_var = tk.StringVar(value="8 AWG")
        harness_cable_combo = ttk.Combobox(self.harness_frame, textvariable=self.harness_cable_size_var, state='readonly', width=10)
        harness_cable_combo['values'] = list(self.AWG_SIZES.keys())
        harness_cable_combo.grid(row=0, column=1, padx=5, pady=2)
        self.harness_cable_size_var.trace('w', lambda *args: self.draw_wiring_layout())

        # Extender Cable Size
        self.extender_frame = ttk.Frame(cable_frame)
        self.extender_frame.grid(row=3, column=0, columnspan=2, padx=0, pady=5, sticky=(tk.W, tk.E))
        ttk.Label(self.extender_frame, text="Extender Cable Size:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.extender_cable_size_var = tk.StringVar(value="8 AWG")
        extender_cable_combo = ttk.Combobox(self.extender_frame, textvariable=self.extender_cable_size_var, state='readonly', width=10)
        extender_cable_combo['values'] = list(self.AWG_SIZES.keys())
        extender_cable_combo.grid(row=0, column=1, padx=5, pady=2)
        self.extender_cable_size_var.trace('w', lambda *args: self.draw_wiring_layout())

        # Whip Cable Size
        self.whip_frame = ttk.Frame(cable_frame)
        self.whip_frame.grid(row=4, column=0, columnspan=2, padx=0, pady=5, sticky=(tk.W, tk.E))
        ttk.Label(self.whip_frame, text="Whip Cable Size:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.whip_cable_size_var = tk.StringVar(value="8 AWG")
        whip_cable_combo = ttk.Combobox(self.whip_frame, textvariable=self.whip_cable_size_var, state='readonly', width=10)
        whip_cable_combo['values'] = list(self.AWG_SIZES.keys())
        whip_cable_combo.grid(row=0, column=1, padx=5, pady=2)
        self.whip_cable_size_var.trace('w', lambda *args: self.draw_wiring_layout())

        # Current label toggle button
        self.show_current_labels_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(cable_frame, text="Show Current Labels", 
                variable=self.show_current_labels_var,
                command=self.toggle_current_labels).grid(
                row=5, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)

    def setup_harness_configuration(self, controls_frame):
        """Set up harness configuration section"""
        # Create collapsible frame
        harness_collapsible = CollapsibleFrame(controls_frame, text="Tracker Harness Configuration", start_collapsed=False)
        harness_collapsible.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Store reference to collapsible frame
        self.harness_collapsible = harness_collapsible
        
        # Use the content_frame inside the collapsible
        self.harness_config_frame = harness_collapsible.content_frame

        # String count selector
        ttk.Label(self.harness_config_frame, text="Configure trackers with:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.string_count_var = tk.StringVar()
        self.string_count_combobox = ttk.Combobox(self.harness_config_frame, textvariable=self.string_count_var, state='readonly')
        self.string_count_combobox.grid(row=0, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        self.string_count_combobox.bind('<<ComboboxSelected>>', self.on_string_count_selected)

        # Add tracker count label
        self.tracker_count_label = ttk.Label(self.harness_config_frame, text="")
        self.tracker_count_label.grid(row=0, column=2, padx=5, pady=2, sticky=tk.W)

        # String grouping frame
        self.string_grouping_frame = ttk.Frame(self.harness_config_frame)
        self.string_grouping_frame.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky=(tk.W, tk.E))

        # Add a subframe for string checkboxes
        self.string_check_frame = ttk.Frame(self.string_grouping_frame)
        self.string_check_frame.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Add Create Harness button
        ttk.Button(self.string_grouping_frame, text="Create Harness from Selected", 
                command=self.create_harness_from_selected).grid(row=1, column=0, padx=5, pady=5)

        # Initially hide harness configuration if not in harness mode
        if self.wiring_type_var.get() != WiringType.HARNESS.value:
            harness_collapsible.grid_remove()

    def setup_harness_cable_table(self, controls_frame):
        """Set up harness cable configuration table"""
        # Create collapsible frame
        cable_table_collapsible = CollapsibleFrame(controls_frame, text="Harness Cable Configuration", start_collapsed=False)
        cable_table_collapsible.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Store reference
        self.cable_table_collapsible = cable_table_collapsible
        
        # Use the content_frame inside the collapsible
        table_frame = cable_table_collapsible.content_frame
        
        # Instructions
        instructions = ttk.Label(table_frame, 
                            text="Configure cable sizes for each harness type. Right-click to reset to defaults.",
                            font=('TkDefaultFont', 9, 'italic'))
        instructions.grid(row=0, column=0, padx=5, pady=(0, 5), sticky=tk.W)
        
        # Create treeview
        columns = ('String Cable', 'Harness Cable', 'Extender Cable', 'Whip Cable', 'Fuse Size', 'Recommended')
        self.harness_cable_tree = ttk.Treeview(table_frame, columns=columns, height=8, show='tree headings')
        self.harness_cable_tree.grid(row=1, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Change cursor when hovering over editable cells
        def on_motion(event):
            region = self.harness_cable_tree.identify_region(event.x, event.y)
            column = self.harness_cable_tree.identify('column', event.x, event.y)
            
            if region == "cell" and column in ['#1', '#2', '#3', '#4', '#5']:
                self.harness_cable_tree.configure(cursor="hand2")
            else:
                self.harness_cable_tree.configure(cursor="")

        self.harness_cable_tree.bind('<Motion>', on_motion)
        
        # Configure columns
        self.harness_cable_tree.column('#0', width=200, minwidth=150)  # Harness type
        self.harness_cable_tree.column('String Cable', width=100, minwidth=80)
        self.harness_cable_tree.column('Harness Cable', width=100, minwidth=80)
        self.harness_cable_tree.column('Extender Cable', width=100, minwidth=80)
        self.harness_cable_tree.column('Whip Cable', width=100, minwidth=80)
        self.harness_cable_tree.column('Fuse Size', width=80, minwidth=60)
        self.harness_cable_tree.column('Recommended', width=100, minwidth=80)
        
        # Configure column headings
        self.harness_cable_tree.heading('#0', text='Harness Type')
        self.harness_cable_tree.heading('String Cable', text='String Cable')
        self.harness_cable_tree.heading('Harness Cable', text='Harness Cable')
        self.harness_cable_tree.heading('Extender Cable', text='Extender Cable')
        self.harness_cable_tree.heading('Whip Cable', text='Whip Cable')
        self.harness_cable_tree.heading('Fuse Size', text='Fuse Size')
        self.harness_cable_tree.heading('Recommended', text='Recommended Whip')
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient='vertical', command=self.harness_cable_tree.yview)
        scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.harness_cable_tree.configure(yscrollcommand=scrollbar.set)
        
        # Tags for styling
        self.harness_cable_tree.tag_configure('edited', foreground='blue')
        self.harness_cable_tree.tag_configure('undersized', background='#ffcccc')
        
        # Bind events for inline editing
        self.harness_cable_tree.bind('<Button-1>', self.on_harness_cable_click)
        self.harness_cable_tree.bind('<Button-3>', self.show_harness_cable_context_menu)
        
        # Track edited cells
        self.harness_cable_edited_cells = set()
        
        # Context menu
        self.harness_cable_menu = tk.Menu(self.harness_cable_tree, tearoff=0)
        self.harness_cable_menu.add_command(label="Reset Cable Sizes to Defaults", command=self.reset_harness_cable_sizes)
        self.harness_cable_menu.add_separator()
        self.harness_cable_menu.add_command(label="Delete Harness", command=self.delete_selected_harness)

        # Add Delete All Harnesses button below the table
        button_frame = ttk.Frame(table_frame)
        button_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=(5, 0), sticky=(tk.W, tk.E))
        ttk.Button(button_frame, text="Delete All Harnesses", 
                command=self.delete_all_harnesses).pack(side=tk.LEFT, padx=5)

    def update_harness_cable_table(self):
        """Update the harness cable configuration table"""
        # Clear existing items
        for item in self.harness_cable_tree.get_children():
            self.harness_cable_tree.delete(item)
        
        # Clear the mapping dictionary
        self.harness_tree_items = {}
        
        if not self.block or not self.block.wiring_config:
            return
        
        if self.block.wiring_config.wiring_type != WiringType.HARNESS:
            return
        
        # Group harnesses by string count
        harness_groups = {}
        
        # Only show actual harness groupings that have been created
        if hasattr(self.block.wiring_config, 'harness_groupings') and self.block.wiring_config.harness_groupings:
            # Count actual harnesses by their string configuration
            for string_count, harness_list in self.block.wiring_config.harness_groupings.items():
                for harness_idx, harness in enumerate(harness_list):
                    # Use actual string count from the harness configuration
                    actual_string_count = len(harness.string_indices)
                    key = f"{actual_string_count}_string_{harness_idx}"
                    
                    # Count how many trackers have this configuration
                    tracker_count = sum(1 for pos in self.block.tracker_positions 
                                    if len(pos.strings) == string_count)
                    
                    if key not in harness_groups:
                        harness_groups[key] = {
                            'string_count': actual_string_count,
                            'count': 0,
                            'harness': harness
                        }
                    harness_groups[key]['count'] += tracker_count
        
        # Add rows for each harness group
        for key, group in sorted(harness_groups.items()):
            string_count = group['string_count']
            harness = group['harness']
            count = group['count']
            
            # Get cable sizes
            if harness:
                string_size = harness.string_cable_size or "10 AWG"
                harness_size = harness.cable_size or "10 AWG"
                extender_size = harness.extender_cable_size or "8 AWG"
                whip_size = harness.whip_cable_size or "8 AWG"
            else:
                # Use defaults
                string_size = "10 AWG"
                harness_size = "10 AWG"
                extender_size = "8 AWG"
                whip_size = "8 AWG"
            
            # Get fuse size and quantity for this harness
            fuse_size = "N/A"  # Default for 1-string harnesses
            if string_count > 1:
                if harness and hasattr(harness, 'fuse_rating_amps'):
                    # Show fuse rating and quantity (one per string)
                    fuse_size = f"{harness.fuse_rating_amps}A ({string_count}x)"
                else:
                    # Calculate default based on Imp
                    default_fuse = self.calculate_recommended_fuse_size(list(range(string_count)))
                    fuse_size = f"{default_fuse}A ({string_count}x)"
            
            # Calculate recommended size for whip
            recommended = self.calculate_recommended_whip_size(string_count)
            
            # Determine tags
            tags = []
            if harness and any([
                harness.string_cable_size,
                harness.extender_cable_size,
                harness.whip_cable_size
            ]):
                tags.append('edited')
            
            # Check if whip is undersized compared to recommended
            if self.is_cable_undersized(whip_size, string_count):
                tags.append('undersized')
            
            # Format label
            label = f"{string_count}-string harnesses ({count} total)"
            
            # Insert row
            item = self.harness_cable_tree.insert('', 'end', text=label,
                                                values=(string_size, harness_size, extender_size, whip_size, fuse_size, recommended),
                                                tags=tuple(tags))

            # Store harness reference in dictionary
            self.harness_tree_items[item] = key

    def calculate_recommended_whip_size(self, string_count):
        """Calculate recommended cable size for whip based on string count"""
        # Get module Isc from block's tracker template
        if not self.block or not self.block.tracker_template or not self.block.tracker_template.module_spec:
            return "8 AWG"
        
        # Calculate current: num_strings × module_Isc × 1.25 (NEC factor)
        module_isc = self.block.tracker_template.module_spec.isc
        total_current = string_count * module_isc * 1.25
        
        # Use the cable sizing utility - whip carries same current as harness
        from ..utils.cable_sizing import calculate_whip_cable_size
        return calculate_whip_cable_size(string_count, module_isc, 1.25)

    def is_cable_undersized(self, cable_size, string_count):
        """Check if cable size is undersized for the given string count"""
        # Get module Isc from block's tracker template
        if not self.block or not self.block.tracker_template or not self.block.tracker_template.module_spec:
            return False
        
        # Calculate required current
        module_isc = self.block.tracker_template.module_spec.isc
        required_current = string_count * module_isc
        
        # Use cable sizing utility to validate
        from ..utils.cable_sizing import validate_cable_size_for_current
        return not validate_cable_size_for_current(cable_size, required_current, 1.25)
    
    def on_harness_cable_click(self, event):
        """Handle click for inline editing - improved UX"""
        # Identify the clicked region
        region = self.harness_cable_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        
        # Get the item and column
        item = self.harness_cable_tree.identify('item', event.x, event.y)
        column = self.harness_cable_tree.identify('column', event.x, event.y)
        
        if not item or not column:
            return
        
        # Only allow editing cable size columns and fuse column (not recommended)
        if column not in ['#1', '#2', '#3', '#4', '#5']:  # string, harness, extender, whip cables, fuse
            return
        
        # Get the cell's bounding box
        bbox = self.harness_cable_tree.bbox(item, column)
        if not bbox:
            return
        
        # Get column index
        col_map = {'#1': 0, '#2': 1, '#3': 2, '#4': 3, '#5': 4}
        col_idx = col_map.get(column)
        
        # Get current value
        values = self.harness_cable_tree.item(item, 'values')
        current_value = values[col_idx]
        
        # Get harness key from dictionary
        harness_key = self.harness_tree_items.get(item)
        if not harness_key:
            return
        
        # Handle fuse column differently
        if col_idx == 4:  # Fuse column
            self.create_harness_fuse_combo(item, column, current_value, harness_key)
        else:  # Cable columns
            cable_types = ['string', 'harness', 'extender', 'whip']
            cable_type = cable_types[col_idx]
            self.create_harness_cable_combo(item, column, current_value, harness_key, cable_type)

    def create_harness_cable_combo(self, item, column, current_value, harness_key, cable_type):
        """Create inline combobox for cable size editing"""
        # Get the bounding box of the cell
        bbox = self.harness_cable_tree.bbox(item, column)
        if not bbox:
            return
        
        # Create combobox
        cable_var = tk.StringVar(value=current_value)
        combo = ttk.Combobox(self.harness_cable_tree, textvariable=cable_var, state='readonly', width=10)
        combo['values'] = list(self.AWG_SIZES.keys())
        
        # Position the combobox
        combo.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        
        def save_cable(event=None):
            new_size = cable_var.get()
            
            # Update the harness configuration
            self.update_harness_cable_size(harness_key, cable_type, new_size)
            
            # Mark as edited
            self.harness_cable_edited_cells.add(f"{harness_key}_{cable_type}")
            
            # Refresh display
            self.update_harness_cable_table()
            self.draw_wiring_layout()
            self.notify_wiring_changed()
            
            combo.destroy()
        
        def cancel(event=None):
            combo.destroy()
        
        combo.bind('<<ComboboxSelected>>', save_cable)
        combo.bind('<Return>', save_cable)
        combo.bind('<Escape>', cancel)
        combo.bind('<FocusOut>', cancel)
        combo.focus_set()
        combo.event_generate('<Button-1>')  # Open dropdown immediately

    def create_harness_fuse_combo(self, item, column, current_value, harness_key):
        """Create inline combobox for fuse size editing"""
        # Parse the harness key to check if this is a multi-string harness
        parts = harness_key.split('_')
        string_count = int(parts[0])
        harness_idx = int(parts[2])
        
        # Don't allow fuse editing for single-string harnesses
        if string_count <= 1:
            return
        
        # Get the bounding box of the cell
        bbox = self.harness_cable_tree.bbox(item, column)
        if not bbox:
            return
        
        # Extract current fuse value (remove 'A' suffix)
        current_fuse = current_value.replace('A', '') if current_value != 'N/A' else '15'
        
        # Create combobox
        fuse_var = tk.StringVar(value=current_fuse)
        combo = ttk.Combobox(self.harness_cable_tree, textvariable=fuse_var, state='readonly', width=10)
        
        # Calculate minimum fuse size based on module Imp
        if hasattr(self.block, 'tracker_template') and self.block.tracker_template.module_spec:
            module_imp = self.block.tracker_template.module_spec.imp
            # Find next standard fuse size above Imp
            combo['values'] = [str(r) for r in self.FUSE_RATINGS if r > module_imp]
        else:
            combo['values'] = [str(r) for r in self.FUSE_RATINGS]
        
        # Position the combobox
        combo.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        
        def save_fuse(event=None):
            new_fuse = int(fuse_var.get())
            
            # Update fuse size for ALL harnesses of the same string count
            updated = False
            for tracker_string_count, harness_list in self.block.wiring_config.harness_groupings.items():
                for harness in harness_list:
                    actual_string_count = len(harness.string_indices)
                    
                    # Update all harnesses with matching string count
                    if actual_string_count == string_count:
                        harness.fuse_rating_amps = new_fuse
                        harness.use_fuse = True
                        updated = True
            
            if updated:
                # Mark as edited
                self.harness_cable_edited_cells.add(f"{harness_key}_fuse")
                
                # Refresh display
                self.update_harness_cable_table()
                self.draw_wiring_layout()
                self.notify_wiring_changed()
            
            combo.destroy()
        
        def cancel(event=None):
            combo.destroy()
        
        combo.bind('<<ComboboxSelected>>', save_fuse)
        combo.bind('<Return>', save_fuse)
        combo.bind('<Escape>', cancel)
        combo.bind('<FocusOut>', cancel)
        combo.focus_set()
        combo.event_generate('<Button-1>')  

    def update_harness_cable_size(self, harness_key, cable_type, new_size):
        """Update cable size for a specific harness group"""
        
        # Parse the harness key
        parts = harness_key.split('_')
        string_count = int(parts[0])
        harness_idx = int(parts[2])
                
        # Ensure harness groupings exist
        if not hasattr(self.block.wiring_config, 'harness_groupings'):
            return        
        
        # Find all harnesses with the matching string count
        updated = False
        for tracker_string_count, harness_list in self.block.wiring_config.harness_groupings.items():
            
            for idx, harness in enumerate(harness_list):
                actual_string_count = len(harness.string_indices)
                
                # Update all harnesses with matching string count
                if actual_string_count == string_count:                    
                    if cable_type == 'string':
                        harness.string_cable_size = new_size
                    elif cable_type == 'harness':
                        harness.cable_size = new_size
                    elif cable_type == 'extender':
                        harness.extender_cable_size = new_size
                    elif cable_type == 'whip':
                        harness.whip_cable_size = new_size
                    
                    updated = True
        
        if updated:
            # Force a redraw of the wiring
            self.draw_wiring_layout()
            self.notify_wiring_changed()

    def show_harness_cable_context_menu(self, event):
        """Show context menu for harness cable table"""
        # Select the item under cursor
        item = self.harness_cable_tree.identify('item', event.x, event.y)
        if item:
            self.harness_cable_tree.selection_set(item)
            self.harness_cable_menu.post(event.x_root, event.y_root)

    def reset_harness_cable_sizes(self):
        """Reset selected harness cable sizes to defaults"""
        selected = self.harness_cable_tree.selection()
        if not selected:
            return
        
        for item in selected:
            harness_key = self.harness_tree_items.get(item)
            
            # Parse the harness key
            parts = harness_key.split('_')
            string_count = int(parts[0])
            harness_idx = int(parts[2])
            
            # Get the harness
            if (string_count in self.block.wiring_config.harness_groupings and
                harness_idx < len(self.block.wiring_config.harness_groupings[string_count])):
                
                harness = self.block.wiring_config.harness_groupings[string_count][harness_idx]
                
                # Reset to defaults by removing custom attributes
                if hasattr(harness, 'string_cable_size'):
                    delattr(harness, 'string_cable_size')
                if hasattr(harness, 'extender_cable_size'):
                    delattr(harness, 'extender_cable_size')
                if hasattr(harness, 'whip_cable_size'):
                    delattr(harness, 'whip_cable_size')
                
                # Reset cable_size to default
                harness.cable_size = self.harness_cable_size_var.get()
                
                # Clear edited markers
                for cable_type in ['string', 'harness', 'extender', 'whip']:
                    cell_id = f"{harness_key}_{cable_type}"
                    self.harness_cable_edited_cells.discard(cell_id)
        
        # Refresh display
        self.update_harness_cable_table()
        self.draw_wiring_layout()
        self.notify_wiring_changed()

    def delete_selected_harness(self):
        """Delete the selected harness from configuration"""
        selected = self.harness_cable_tree.selection()
        if not selected:
            return
        
        # Confirm deletion
        if not messagebox.askyesno("Delete Harness", "Are you sure you want to delete the selected harness?"):
            return
        
        # Collect all harnesses to delete with their details
        harnesses_to_delete = []
        for item in selected:
            harness_key = self.harness_tree_items.get(item)            
            if not harness_key:
                continue
                
            # Parse the harness key
            parts = harness_key.split('_')
            actual_string_count = int(parts[0])  # Number of strings IN the harness
            global_harness_idx = int(parts[2])   # Global index across all trackers
            
            # Build a list of ALL harnesses with this actual string count
            all_matching_harnesses = []
            for tracker_string_count, harness_list in self.block.wiring_config.harness_groupings.items():
                for harness in harness_list:
                    if len(harness.string_indices) == actual_string_count:
                        all_matching_harnesses.append((tracker_string_count, harness.string_indices))
                        
            # Find the harness at the global index
            if global_harness_idx < len(all_matching_harnesses):
                tracker_string_count, string_indices = all_matching_harnesses[global_harness_idx]
                harnesses_to_delete.append((tracker_string_count, string_indices))
                
        # Now delete harnesses by matching their string_indices
        for tracker_string_count, target_string_indices in harnesses_to_delete:
            if tracker_string_count in self.block.wiring_config.harness_groupings:
                original_list = self.block.wiring_config.harness_groupings[tracker_string_count]                
                new_list = [h for h in original_list if h.string_indices != target_string_indices]                
                if new_list:
                    self.block.wiring_config.harness_groupings[tracker_string_count] = new_list
                else:
                    del self.block.wiring_config.harness_groupings[tracker_string_count]
        
        # Update display - force complete refresh
        self.harness_cable_tree.delete(*self.harness_cable_tree.get_children())
        self.harness_tree_items.clear()
        self.update_harness_cable_table()
        self.draw_wiring_layout()
        self.notify_wiring_changed()

    def delete_all_harnesses(self):
        """Delete all harnesses from the configuration"""
        if not hasattr(self.block.wiring_config, 'harness_groupings') or not self.block.wiring_config.harness_groupings:
            messagebox.showinfo("Info", "No harnesses to delete")
            return
        
        # Count total harnesses
        total_harnesses = sum(len(harness_list) for harness_list in self.block.wiring_config.harness_groupings.values())
        
        # Confirm deletion
        if not messagebox.askyesno("Delete All Harnesses", 
                                f"Are you sure you want to delete all {total_harnesses} harnesses?\n\nThis action cannot be undone."):
            return
        
        # Clear all harness groupings
        self.block.wiring_config.harness_groupings.clear()
        
        # Clear edited cells tracking
        self.harness_cable_edited_cells.clear()
        
        # Update display
        self.harness_cable_tree.delete(*self.harness_cable_tree.get_children())
        self.harness_tree_items.clear()
        self.update_harness_cable_table()
        self.draw_wiring_layout()
        self.notify_wiring_changed()
        
        messagebox.showinfo("Success", f"Deleted all {total_harnesses} harnesses")
        
    def setup_routing_controls(self, controls_frame):
        """Set up remaining routing controls section (whip controls moved to bottom)"""
        # This method now intentionally left minimal since whip controls moved to canvas area
        pass

    def setup_legend(self, canvas_frame):
        """Set up color legend and whip point controls"""
        # Container for both legend and controls
        bottom_controls_frame = ttk.Frame(canvas_frame)
        bottom_controls_frame.grid(row=1, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Configure grid weights so controls expand properly
        bottom_controls_frame.columnconfigure(0, weight=1)
        bottom_controls_frame.columnconfigure(1, weight=1)
        
        # Color legend on the left
        legend_frame = ttk.LabelFrame(bottom_controls_frame, text="Wire Color Legend", padding="5")
        legend_frame.grid(row=0, column=0, padx=(0, 5), pady=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Create legend canvas
        self.legend_canvas = tk.Canvas(legend_frame, width=280, height=120, bg='white')
        self.legend_canvas.grid(row=0, column=0, padx=5, pady=5)

        # Draw the legend
        self.draw_legend()

        # Add current labels toggle to legend frame
        self.show_current_labels_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(legend_frame, text="Show Current Labels", 
                        variable=self.show_current_labels_var,
                        command=self.toggle_current_labels).grid(
                        row=1, column=0, padx=5, pady=5, sticky=tk.W)
        
        # Whip point controls on the right
        whip_controls_frame = ttk.LabelFrame(bottom_controls_frame, text="Whip Point Controls", padding="5")
        whip_controls_frame.grid(row=0, column=1, padx=(5, 0), pady=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Routing mode controls
        routing_frame = ttk.Frame(whip_controls_frame)
        routing_frame.grid(row=0, column=0, columnspan=2, padx=0, pady=2, sticky=(tk.W, tk.E))
        
        self.routing_mode_var = tk.StringVar(value="realistic")
        ttk.Radiobutton(routing_frame, text="Realistic Routing", 
                    variable=self.routing_mode_var, value="realistic",
                    command=self.draw_wiring_layout).grid(row=0, column=0, padx=5, sticky=tk.W)
        ttk.Radiobutton(routing_frame, text="Conceptual Routing", 
                    variable=self.routing_mode_var, value="conceptual",
                    command=self.draw_wiring_layout).grid(row=0, column=1, padx=5, sticky=tk.W)
        
        # Movement constraint controls
        constraint_frame = ttk.Frame(whip_controls_frame)
        constraint_frame.grid(row=1, column=0, columnspan=2, padx=0, pady=2, sticky=(tk.W, tk.E))
        
        self.snap_5ft_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(constraint_frame, text="Snap to 5ft increments", 
                    variable=self.snap_5ft_var).grid(row=0, column=0, padx=5, sticky=tk.W)
        
        self.right_angle_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(constraint_frame, text="Constrain to N/S movement", 
               variable=self.right_angle_var).grid(row=0, column=1, padx=5, sticky=tk.W)
        
        # Reset buttons
        button_frame = ttk.Frame(whip_controls_frame)
        button_frame.grid(row=2, column=0, columnspan=2, padx=0, pady=5, sticky=(tk.W, tk.E))
        
        ttk.Button(button_frame, text="Reset Selected", 
                command=self.reset_selected_whips).grid(row=0, column=0, padx=5, sticky=tk.W)
        ttk.Button(button_frame, text="Reset All Whips", 
                command=self.reset_all_whips).grid(row=0, column=1, padx=5, sticky=tk.W)

    def setup_warning_panel(self):
        """Set up warning panel"""
        self.warning_panel = tk.Frame(self.canvas, bg='white', bd=1, relief=tk.RAISED)
        self.warning_panel.place(relx=1.0, rely=1.0, x=-5, y=-5, anchor='se')
        
        # Add header
        self.warning_header = tk.Label(self.warning_panel, text="Wire Warnings", 
                                    bg='#f0f0f0', font=('Arial', 9, 'bold'),
                                    padx=5, pady=2, anchor='w', width=30)
        self.warning_header.pack(side=tk.TOP, fill=tk.X)
        
        # Add scrollable warning list
        self.warning_frame = tk.Frame(self.warning_panel, bg='white')
        self.warning_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        # Store warnings with their associated wire IDs
        self.wire_warnings = {}
        self.current_labels = {}
        self.dragging_label = None
        self.label_drag_start = None

    def setup_canvas_bindings(self):
        """Set up canvas event bindings"""
        # Bind canvas resize
        self.canvas.bind('<Configure>', self.on_canvas_resize)

        # Add mouse wheel binding for zoom (only when over canvas)
        def on_canvas_mousewheel(event):
            # Only zoom if we're actually over the canvas, not the controls
            if self.canvas.winfo_containing(event.x_root, event.y_root) == self.canvas:
                self.on_mouse_wheel(event)
                
        self.canvas.bind('<MouseWheel>', on_canvas_mousewheel)  # Windows
        self.canvas.bind('<Button-4>', on_canvas_mousewheel)    # Linux scroll up
        self.canvas.bind('<Button-5>', on_canvas_mousewheel)    # Linux scroll down

        # Pan bindings
        self.canvas.bind('<Button-2>', self.start_pan)  # Middle mouse button
        self.canvas.bind('<B2-Motion>', self.update_pan)
        self.canvas.bind('<ButtonRelease-2>', self.end_pan)
        # Alternative right-click pan
        self.canvas.bind('<Button-3>', self.start_pan)  
        self.canvas.bind('<B3-Motion>', self.update_pan)
        self.canvas.bind('<ButtonRelease-3>', self.end_pan)

        # Bindings for whip point interaction
        self.canvas.bind('<Button-1>', self.on_canvas_click)
        self.canvas.bind('<B1-Motion>', self.on_canvas_drag)
        self.canvas.bind('<ButtonRelease-1>', self.on_canvas_release)
        
        # Add keyboard shortcuts
        self.canvas.bind('<Delete>', self.reset_selected_whips)
        self.canvas.bind('<Control-a>', self.select_all_whips)
        self.canvas.bind('<Escape>', self.clear_selection)
        
        # Add right-click menu
        self.whip_menu = tk.Menu(self.canvas, tearoff=0)
        self.whip_menu.add_command(label="Reset to Default", command=self.reset_selected_whips)
        self.canvas.bind('<Button-3>', self.show_context_menu)

    def load_existing_configuration(self):
        """Load existing configuration if available"""
        if self.block.wiring_config:
            # Set wiring type
            self.wiring_type_var.set(self.block.wiring_config.wiring_type.value)
            
            # Set cable sizes
            if hasattr(self.block.wiring_config, 'string_cable_size'):
                self.string_cable_size_var.set(self.block.wiring_config.string_cable_size)
            
            if hasattr(self.block.wiring_config, 'harness_cable_size'):
                self.harness_cable_size_var.set(self.block.wiring_config.harness_cable_size)

            if hasattr(self.block.wiring_config, 'whip_cable_size'):
                self.whip_cable_size_var.set(self.block.wiring_config.whip_cable_size)

            if hasattr(self.block.wiring_config, 'extender_cable_size'):
                self.extender_cable_size_var.set(self.block.wiring_config.extender_cable_size)

            # Restore routing mode
            if hasattr(self.block.wiring_config, 'routing_mode'):
                self.routing_mode_var.set(self.block.wiring_config.routing_mode)
            
            # Update UI based on wiring type
            self.update_ui_for_wiring_type()
            
            # Load harness cable table if in harness mode
            if self.block.wiring_config.wiring_type == WiringType.HARNESS:
                # Update the harness cable table to show existing configurations
                self.update_harness_cable_table()
                
                # If we have custom harness cable sizes, mark them as edited
                if hasattr(self.block.wiring_config, 'harness_groupings'):
                    for string_count, harness_list in self.block.wiring_config.harness_groupings.items():
                        for harness_idx, harness in enumerate(harness_list):
                            # Check if any custom cable sizes are set
                            if (hasattr(harness, 'string_cable_size') and harness.string_cable_size) or \
                               (hasattr(harness, 'extender_cable_size') and harness.extender_cable_size) or \
                               (hasattr(harness, 'whip_cable_size') and harness.whip_cable_size):
                                # Mark cells as edited
                                actual_string_count = len(harness.string_indices)
                                harness_key = f"{actual_string_count}_string_{harness_idx}"
                                
                                if hasattr(harness, 'string_cable_size') and harness.string_cable_size:
                                    self.harness_cable_edited_cells.add(f"{harness_key}_string")
                                if hasattr(harness, 'extender_cable_size') and harness.extender_cable_size:
                                    self.harness_cable_edited_cells.add(f"{harness_key}_extender")
                                if hasattr(harness, 'whip_cable_size') and harness.whip_cable_size:
                                    self.harness_cable_edited_cells.add(f"{harness_key}_whip")

    def draw_wiring_layout(self):
        """Draw block layout with wiring visualization"""
        # Ensure all tracker positions have project reference for wiring mode
        if self.project:
            for pos in self.block.tracker_positions:
                pos._project_ref = self.project
        # Initialize route storage for Block Configurator
        self.saved_routes_for_block = {}

        # Save current label positions before clearing
        saved_labels = {}
        if hasattr(self, 'current_labels'):
            for label_id, label_info in self.current_labels.items():
                if self.canvas.coords(label_info['text_id']):  # Check if label exists
                    saved_labels[label_id] = label_info['current_pos']

        self.canvas.delete("all")
        # Reset extended extender points
        self.extended_extender_points = {}

        # Reset current labels tracking
        self.current_labels = {}
        
        # Clear previous warnings
        self.clear_warnings()
        
        # Re-create warning panel (since we deleted all canvas items)
        self.setup_warning_panel()

        # Draw trackers and devices
        self.draw_trackers()
        self.draw_device()
        self.draw_device_destination_points()

        # Draw routes based on current routing mode
        self.draw_current_routes()

        # Restore any saved label positions
        if hasattr(self, 'saved_labels') and saved_labels:
            for label_id, saved_pos in saved_labels.items():
                # Find the corresponding new label and move it
                if label_id in self.current_labels:
                    label_info = self.current_labels[label_id]
                    text_id = label_info['text_id']
                    self.canvas.coords(text_id, saved_pos[0], saved_pos[1])
                    label_info['current_pos'] = saved_pos
            
            # Clear the saved positions
            self.saved_labels = {}
                        
    def draw_collection_points(self, pos: TrackerPosition, x: float, y: float, scale: float):
        """Draw collection points for a tracker position"""
        for string in pos.strings:
            # Draw positive and negative source points
            pos_world_x = pos.x + string.positive_source_x
            pos_world_y = pos.y + string.positive_source_y
            neg_world_x = pos.x + string.negative_source_x
            neg_world_y = pos.y + string.negative_source_y

            self.draw_collection_point(pos_world_x, pos_world_y, True, 'collection')
            self.draw_collection_point(neg_world_x, neg_world_y, False, 'collection')
        
        # Draw whip points
        tracker_idx = self.block.tracker_positions.index(pos)
        tracker_id = str(tracker_idx)
        
        # Get whip point positions
        pos_whip = self.get_whip_position(tracker_id, 'positive')
        neg_whip = self.get_whip_position(tracker_id, 'negative')
        
        # Draw positive whip point
        if pos_whip:
            wx, wy = self.world_to_canvas(pos_whip[0], pos_whip[1])
            
            # Highlight if selected
            is_selected = (tracker_id, 'positive') in self.selected_whips
            fill = 'red' if not is_selected else 'orange'
            outline = 'darkred' if not is_selected else 'red'
            size = 3 if not is_selected else 5
            
            self.canvas.create_oval(
                wx - size, wy - size,
                wx + size, wy + size,
                fill=fill, outline=outline,
                tags='whip_point'
            )
            
        # Draw negative whip point
        if neg_whip:
            wx, wy = self.world_to_canvas(neg_whip[0], neg_whip[1])
            
            # Highlight if selected
            is_selected = (tracker_id, 'negative') in self.selected_whips
            fill = 'blue' if not is_selected else 'cyan'
            outline = 'darkblue' if not is_selected else 'blue'
            size = 3 if not is_selected else 5
            
            self.canvas.create_oval(
                wx - size, wy - size,
                wx + size, wy + size,
                fill=fill, outline=outline,
                tags='whip_point'
            )

    def on_wiring_type_change(self, event=None):
        """Handle wiring type selection change"""
        self.update_ui_for_wiring_type()
    
    def on_wiring_mode_change(self):
        """Handle wiring mode change"""
        new_mode = self.wiring_mode_var.get()
        
        # Show warning dialog
        result = messagebox.askyesno(
            "Change Wiring Mode", 
            "Changing the wiring mode will affect all blocks in the project. Do you want to proceed?"
        )
        
        if result:
            # Update project wiring mode
            if self.project:
                self.project.wiring_mode = new_mode
            
            # Recalculate string positions for current block
            for pos in self.block.tracker_positions:
                # Pass project reference so it can access wiring mode
                pos._project_ref = self.project
                pos.calculate_string_positions()
            
            # Clear the canvas and redraw everything
            self.canvas.delete("all")
            self.draw_wiring_layout()
        else:
            # Revert the radio button
            if self.project and hasattr(self.project, 'wiring_mode'):
                self.wiring_mode_var.set(self.project.wiring_mode)
            else:
                self.wiring_mode_var.set('daisy_chain')
        
    def on_canvas_resize(self, event):
        """Handle canvas resize event"""
        if event.width > 1 and event.height > 1:  # Ensure valid dimensions
            self.draw_wiring_layout()
        
    def apply_configuration(self):
        """Apply the current wiring configuration to the block"""
        if not self.block:
            tk.messagebox.showerror("Error", "No block selected")
            self.destroy()
            return
            
        try:
            # Create wiring configuration
            wiring_config = self.create_wiring_configuration()
            
            # Validate configuration
            if not self.validate_configuration():
                return
            
            # Store configuration and notify success
            self.block.wiring_config = wiring_config
            tk.messagebox.showinfo("Success", "Wiring configuration applied successfully")
            
            # Check MPPT capacity and notify parent
            self.perform_final_checks()
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            tk.messagebox.showerror("Error", f"Failed to apply wiring configuration: {str(e)}")

    def create_wiring_configuration(self):
        """Create the WiringConfig instance from current settings"""
        wiring_type = WiringType(self.wiring_type_var.get())
        
        # Create collection points
        positive_collection_points, negative_collection_points, strings_per_collection = self.build_collection_points()
        
        # Get routes based on current routing mode
        cable_routes = self.get_current_routes()
        
        # Get existing custom settings
        custom_whip_points, harness_groupings = self.get_existing_custom_settings()
        
        # If we have existing harness groupings with custom cable sizes, preserve them
        if harness_groupings and hasattr(self.block.wiring_config, 'harness_groupings'):
            # Create a new harness_groupings dict to preserve custom cable sizes
            preserved_harness_groupings = {}
            
            for string_count, existing_harness_list in self.block.wiring_config.harness_groupings.items():
                preserved_harness_groupings[string_count] = []
                
                for existing_harness in existing_harness_list:
                    # Create a new HarnessGroup with preserved cable sizes
                    new_harness = HarnessGroup(
                        string_indices=existing_harness.string_indices,
                        cable_size=existing_harness.cable_size,  # Preserve harness cable size
                        string_cable_size=getattr(existing_harness, 'string_cable_size', ''),
                        extender_cable_size=getattr(existing_harness, 'extender_cable_size', ''),
                        whip_cable_size=getattr(existing_harness, 'whip_cable_size', ''),
                        fuse_rating_amps=getattr(existing_harness, 'fuse_rating_amps', 15),
                        use_fuse=getattr(existing_harness, 'use_fuse', True)
                    )
                    preserved_harness_groupings[string_count].append(new_harness)
            
            harness_groupings = preserved_harness_groupings
        
        # Create the WiringConfig instance
        return WiringConfig(
            wiring_type=wiring_type,
            positive_collection_points=positive_collection_points,
            negative_collection_points=negative_collection_points,
            strings_per_collection=strings_per_collection,
            cable_routes=cable_routes,
            string_cable_size=self.string_cable_size_var.get(),
            harness_cable_size=self.harness_cable_size_var.get(),
            whip_cable_size=self.whip_cable_size_var.get(),
            extender_cable_size=self.extender_cable_size_var.get(),
            custom_whip_points=custom_whip_points,
            harness_groupings=harness_groupings,
            routing_mode=self.routing_mode_var.get()
        )

    def build_collection_points(self):
        """Build collection points from tracker positions"""
        positive_collection_points = []
        negative_collection_points = []
        strings_per_collection = {}
        
        # Process each tracker position to capture collection points
        for idx, pos in enumerate(self.block.tracker_positions):
            if not pos.template:
                continue
                
            # Get whip points for this tracker
            pos_whip = self.get_whip_position(str(idx), 'positive')
            neg_whip = self.get_whip_position(str(idx), 'negative')
            
            # Add collection points
            if pos_whip:
                collection_point = CollectionPoint(
                    x=pos_whip[0],
                    y=pos_whip[1],
                    connected_strings=[s.index for s in pos.strings],
                    current_rating=self.calculate_current_for_segment('whip', num_strings=len(pos.strings))
                )
                positive_collection_points.append(collection_point)
                strings_per_collection[idx] = len(pos.strings)
            
            if neg_whip:
                collection_point = CollectionPoint(
                    x=neg_whip[0],
                    y=neg_whip[1],
                    connected_strings=[s.index for s in pos.strings],
                    current_rating=self.calculate_current_for_segment('whip', num_strings=len(pos.strings))
                )
                negative_collection_points.append(collection_point)
                
        return positive_collection_points, negative_collection_points, strings_per_collection

    def get_existing_custom_settings(self):
        """Get existing custom whip points and harness groupings"""
        custom_whip_points = {}
        if (hasattr(self.block, 'wiring_config') and 
            self.block.wiring_config and 
            hasattr(self.block.wiring_config, 'custom_whip_points')):
            custom_whip_points = self.block.wiring_config.custom_whip_points

        harness_groupings = {}
        if (hasattr(self.block, 'wiring_config') and 
            self.block.wiring_config and 
            hasattr(self.block.wiring_config, 'harness_groupings')):
            harness_groupings = self.block.wiring_config.harness_groupings
            
        return custom_whip_points, harness_groupings

    def validate_configuration(self):
        """Validate the wiring configuration"""
        return True  # Add specific validation logic if needed

    def perform_final_checks(self):
        """Perform final validation checks and notify parent"""
        if not self.validate_mppt_capacity():
            mppt_warning = messagebox.askyesno(
                "MPPT Current Warning",
                "The selected inverter MPPT capacity may be insufficient for the total string current.\n\n"
                "Total string current exceeds the inverter MPPT capacity.\n"
                "This may require reconfiguring the wiring or choosing a different inverter.\n\n"
                "Do you want to apply this configuration anyway?",
                icon='warning'
            )
            
            if not mppt_warning:
                return
                
        if self.parent_notify_blocks_changed:
            self.parent_notify_blocks_changed()
        self.destroy()
        
    def cancel(self):
        """Cancel wiring configuration"""
        self.destroy()

    def on_mouse_wheel(self, event):
        """Handle mouse wheel events for zooming"""
        # Get the current scale
        old_scale = self.scale_factor
        
        # Update scale factor based on scroll direction
        if event.num == 5 or event.delta < 0:  # Scroll down or delta negative
            self.scale_factor = max(5.0, self.scale_factor * 0.9)
        if event.num == 4 or event.delta > 0:  # Scroll up or delta positive
            self.scale_factor = min(50.0, self.scale_factor * 1.1)
            
        # Redraw if scale changed
        if old_scale != self.scale_factor:
            self.draw_wiring_layout()

    def get_canvas_scale(self):
        """Return current scale factor (pixels per meter)"""
        return self.scale_factor
    

    def start_pan(self, event):
        """Start canvas panning"""
        self.panning = True
        self.pan_start_x = event.x
        self.pan_start_y = event.y
        self.canvas.config(cursor="fleur")  # Change cursor to indicate panning

    def update_pan(self, event):
        """Update canvas pan position"""
        if not self.panning:
            return
        
        # Calculate the distance moved
        dx = event.x - self.pan_start_x
        dy = event.y - self.pan_start_y
        
        # Update pan offset
        self.pan_x += dx
        self.pan_y += dy

        # Update start position for next movement
        self.pan_start_x = event.x
        self.pan_start_y = event.y
        
        # Redraw
        self.draw_wiring_layout()

    def end_pan(self, event):
        """End canvas panning"""
        self.panning = False
        self.canvas.config(cursor="")  # Reset cursor

    def get_device_destination_points(self):
        """Get the two main destination points on the device (positive and negative)"""
        if not self.block:
            return None, None
            
        device_width = 0.91  # 3ft in meters
        device_x = self.block.device_x
        device_y = self.block.device_y
        
        # Positive point on left side, halfway up
        positive_point = (device_x, device_y + (device_width / 2))
        
        # Negative point on right side, halfway up
        negative_point = (device_x + device_width, device_y + (device_width / 2))
        
        return positive_point, negative_point

    def draw_device_destination_points(self):
        """Draw the positive and negative destination points on the device"""
        pos_point, neg_point = self.get_device_destination_points()
        if not pos_point or not neg_point:
            return
        
        # Draw positive destination point (red)
        px, py = self.world_to_canvas(pos_point[0], pos_point[1])
        self.canvas.create_oval(
            px - 4, py - 4,
            px + 4, py + 4,
            fill='red',  # Match source point colors
            outline='darkred',
            tags='destination_point'
        )
        
        # Draw negative destination point (blue)
        nx, ny = self.world_to_canvas(neg_point[0], neg_point[1])
        self.canvas.create_oval(
            nx - 4, ny - 4,
            nx + 4, ny + 4,
            fill='blue',  # Match source point colors
            outline='darkblue',
            tags='destination_point'
        )

    def calculate_cable_route(self, source_x: float, source_y: float, 
                            dest_x: float, dest_y: float, 
                            is_positive: bool, route_index: int) -> List[tuple[float, float]]:
        """
        Calculate cable route from source point to destination point.
        
        Args:
            source_x, source_y: Starting coordinates in meters
            dest_x, dest_y: Ending coordinates in meters
            is_positive: True if this is a positive (red) wire, False if negative (blue)
            route_index: Index of this route for offsetting parallel runs
            
        Returns:
            List of (x, y) coordinates defining the route
        """
        route = []
        offset = 0.1 * route_index  # 0.1m offset between parallel runs
        
        # Add starting point
        route.append((source_x, source_y))

        # Use horizontal-then-vertical routing for all configurations
        route.append((dest_x, source_y))  # Horizontal segment
        route.append((dest_x, dest_y))    # Vertical segment
        
        return route
    
    def calculate_node_points(self, pos, is_positive: bool) -> List[tuple[float, float]]:
        """
        Calculate node points for a tracker's wire harness.
        
        Args:
            pos: TrackerPosition object
            is_positive: True for positive (red) nodes, False for negative (blue) nodes
            
        Returns:
            List of (x, y) coordinates for node points
        """
        nodes = []
        horizontal_offset = 0.6  # ~2ft offset from tracker
        
        # Sort strings by y-position
        source_points = []
        for string in pos.strings:
            if is_positive:
                source_points.append((string.positive_source_x, string.positive_source_y))
            else:
                source_points.append((string.negative_source_x, string.negative_source_y))
        
        source_points.sort(key=lambda p: p[1])  # Sort by y-coordinate
        
        # Create node points at same y-level as source points, but offset horizontally
        for i, (sx, sy) in enumerate(source_points):
            node_x = pos.x + sx + (horizontal_offset if not is_positive else -horizontal_offset)
            node_y = pos.y + sy  # Same y-level as source point
            nodes.append((node_x, node_y))
            
        return nodes

    def calculate_whip_points(self, pos, is_positive: bool) -> tuple[float, float]:
        if self.block.device_y is None or self.block.device_x is None:
            return None
            
        horizontal_offset = 0.6  # Same offset as node points
        whip_offset = 2.0  # 2m offset from tracker edge
        
        # Get tracker dimensions
        tracker_dims = pos.template.get_physical_dimensions()
        tracker_height = tracker_dims[0]
        tracker_width = tracker_dims[1]
        
        # Get tracker Y boundaries
        tracker_north_y = pos.y
        tracker_south_y = pos.y + tracker_height
        device_y = self.block.device_y

        # Check if this tracker has multiple harnesses
        tracker_idx = self.block.tracker_positions.index(pos)
        string_count = len(pos.strings)
        has_multiple_harnesses = (hasattr(self.block, 'wiring_config') and 
                                self.block.wiring_config and 
                                hasattr(self.block.wiring_config, 'harness_groupings') and
                                string_count in self.block.wiring_config.harness_groupings and
                                len(self.block.wiring_config.harness_groupings[string_count]) > 1)

        # Calculate Y position based on device location and harness configuration
        if device_y < tracker_north_y:
            # Device north of tracker - place whip at north end
            y = tracker_north_y - whip_offset
        elif device_y > tracker_south_y:
            # Device south of tracker - place whip at south end  
            y = tracker_south_y + whip_offset
        else:
            # Device in middle of tracker
            if has_multiple_harnesses:
                # Multiple harnesses - whip can be at device level
                y = device_y
            else:
                # Single harness - whip must be at tracker end (choose closer end)
                distance_to_north = abs(device_y - tracker_north_y)
                distance_to_south = abs(device_y - tracker_south_y)
                
                if distance_to_north <= distance_to_south:
                    y = tracker_north_y - whip_offset
                else:
                    y = tracker_south_y + whip_offset
        
        # Calculate X position - match node points pattern
        if is_positive:
            x = pos.x - horizontal_offset  # Left side
        else:
            x = pos.x + tracker_width + horizontal_offset  # Right side
        
        return (x, y)

    def get_line_thickness_for_wire(self, wire_gauge: str) -> float:
        """
        Convert wire gauge to appropriate line thickness for display
        
        Args:
            wire_gauge: String representing AWG gauge (e.g., "10 AWG")
            
        Returns:
            float: Line thickness in pixels
        """
        # Map AWG sizes to line thickness (larger wire = thicker line)
        thickness_map = {
            "4 AWG": 5.0,
            "6 AWG": 4.0,
            "8 AWG": 3.0, 
            "10 AWG": 2.0
        }
        
        return thickness_map.get(wire_gauge, 2.0)  # Default to 2.0 if gauge not found

    def calculate_current_for_segment(self, segment_type: str, num_strings: int = 1) -> float:
        """Calculate current flowing through a wire segment based on configuration"""
        
        # Initialize default string current
        string_current = 10.0  # Default fallback value in amps
        
        # Direct access to module spec from block's tracker template
        if (self.block and 
            hasattr(self.block, 'tracker_template') and 
            self.block.tracker_template and 
            hasattr(self.block.tracker_template, 'module_spec') and 
            self.block.tracker_template.module_spec):
            
            module_spec = self.block.tracker_template.module_spec
            string_current = module_spec.imp
        
        result = string_current * num_strings
        
        return result
    
    def validate_harness_cable_sizes(self):
        """Validate all cable sizes for each harness assembly and return warnings"""
        warnings = []
        
        if not self.block or not self.block.wiring_config:
            return warnings
        
        if self.block.wiring_config.wiring_type != WiringType.HARNESS:
            return warnings
        
        # Check each tracker
        for tracker_idx, pos in enumerate(self.block.tracker_positions):
            string_count = len(pos.strings)
            
            # Get harness groupings for this string count
            if (hasattr(self.block.wiring_config, 'harness_groupings') and 
                string_count in self.block.wiring_config.harness_groupings):
                
                for harness_idx, harness in enumerate(self.block.wiring_config.harness_groupings[string_count]):
                    if not harness.string_indices:
                        continue
                    
                    num_strings = len(harness.string_indices)
                    
                    # Validate each cable type
                    # String cable
                    string_size = harness.string_cable_size if harness.string_cable_size else self.block.wiring_config.string_cable_size
                    string_current = self.calculate_current_for_segment('string')
                    string_valid = self.validate_cable_size(string_size, string_current)
                    
                    # Harness cable
                    harness_size = harness.cable_size
                    harness_current = self.calculate_current_for_segment('harness', num_strings)
                    harness_valid = self.validate_cable_size(harness_size, harness_current)
                    
                    # Extender cable
                    extender_size = harness.extender_cable_size if harness.extender_cable_size else self.block.wiring_config.extender_cable_size
                    extender_current = self.calculate_current_for_segment('extender', num_strings)
                    extender_valid = self.validate_cable_size(extender_size, extender_current)
                    
                    # Whip cable
                    whip_size = harness.whip_cable_size if harness.whip_cable_size else self.block.wiring_config.whip_cable_size
                    whip_current = self.calculate_current_for_segment('whip', num_strings)
                    whip_valid = self.validate_cable_size(whip_size, whip_current)
                    
                    # Add warnings for any undersized cables
                    harness_id = f"T{tracker_idx+1:02d}-H{harness_idx+1:02d}"
                    
                    if not string_valid:
                        warnings.append(f"{harness_id} string cable {string_size} undersized")
                    if not harness_valid:
                        warnings.append(f"{harness_id} harness cable {harness_size} undersized")
                    if not extender_valid:
                        warnings.append(f"{harness_id} extender cable {extender_size} undersized")
                    if not whip_valid:
                        warnings.append(f"{harness_id} whip cable {whip_size} undersized")
        
        return warnings

    def validate_cable_size(self, cable_size, current):
        """Validate if a cable size is adequate for the given current"""
        ampacity = get_ampacity_for_wire_gauge(cable_size)
        if ampacity == 0:
            return True  # Unknown size, assume OK
        
        nec_current = calculate_nec_current(current)
        return nec_current <= ampacity

    def add_current_label(self, points, current, is_positive, segment_type='string'):
        """Add current label to a wire segment if enabled"""
        if not self.show_current_labels_var.get():
            return
            
        if len(points) < 2:
            return
            
        # Find midpoint of line segment
        if len(points) == 2:
            mid_x = (points[0][0] + points[1][0]) / 2
            mid_y = (points[0][1] + points[1][1]) / 2
        else:
            # For multi-segment lines, choose a good point (middle segment)
            mid_idx = len(points) // 2
            mid_x = (points[mid_idx-1][0] + points[mid_idx][0]) / 2
            mid_y = (points[mid_idx-1][1] + points[mid_idx][1]) / 2
        
        # Adjust label position slightly above/below line
        offset = -8 if is_positive else 8
        color = 'red' if is_positive else 'blue'
        
        self.canvas.create_text(mid_x, mid_y + offset, text=f"{current:.1f}A", 
                            fill=color, font=('Arial', 8))
        
    def add_homerun_routes(self, cable_routes, pos, string, tracker_idx, string_idx, pos_whip, neg_whip):
        """Add homerun wiring routes to the cable_routes dictionary"""
        # Source to whip routes
        if pos_whip:
            cable_routes[f"pos_src_{tracker_idx}_{string_idx}"] = [
                (pos.x + string.positive_source_x, pos.y + string.positive_source_y),
                pos_whip
            ]
        
        if neg_whip:
            cable_routes[f"neg_src_{tracker_idx}_{string_idx}"] = [
                (pos.x + string.negative_source_x, pos.y + string.negative_source_y),
                neg_whip
            ]
        
        # Whip to device routes
        pos_dest, neg_dest = self.get_device_destination_points()
        if pos_whip and pos_dest:
            cable_routes[f"pos_dev_{tracker_idx}_{string_idx}"] = [
                pos_whip, pos_dest
            ]
        
        if neg_whip and neg_dest:
            cable_routes[f"neg_dev_{tracker_idx}_{string_idx}"] = [
                neg_whip, neg_dest
            ]

    def add_harness_routes(self, cable_routes, pos, string, tracker_idx, string_idx, pos_whip, neg_whip):
        """Add harness wiring routes to the cable_routes dictionary"""
        # Calculate node points
        pos_nodes = self.calculate_node_points(pos, True)
        neg_nodes = self.calculate_node_points(pos, False)
        
        # String to node routes
        if pos_nodes and len(pos_nodes) > string_idx:
            cable_routes[f"pos_node_{tracker_idx}_{string_idx}"] = [
                (pos.x + string.positive_source_x, pos.y + string.positive_source_y),
                pos_nodes[string_idx]
            ]
        
        if neg_nodes and len(neg_nodes) > string_idx:
            cable_routes[f"neg_node_{tracker_idx}_{string_idx}"] = [
                (pos.x + string.negative_source_x, pos.y + string.negative_source_y),
                neg_nodes[string_idx]
            ]
        
        # Only add harness routes once per tracker
        if string_idx == 0:
            # Check if extenders are needed
            needs_extender = self.needs_extender(tracker_idx)
            pos_extender_point = None
            neg_extender_point = None
            
            if needs_extender:
                pos_extender_point = self.get_extender_point(tracker_idx, 'positive')
                neg_extender_point = self.get_extender_point(tracker_idx, 'negative')
            
            # Node-to-node to whip/extender routes (for all nodes)
            if pos_nodes:
                target_point = pos_extender_point if needs_extender else pos_whip
                if target_point:
                    cable_routes[f"pos_harness_{tracker_idx}"] = pos_nodes + [target_point]
            
            if neg_nodes:
                target_point = neg_extender_point if needs_extender else neg_whip
                if target_point:
                    cable_routes[f"neg_harness_{tracker_idx}"] = neg_nodes + [target_point]
            
            # Add extender routes if needed
            if needs_extender:
                if pos_extender_point and pos_whip:
                    cable_routes[f"pos_extender_{tracker_idx}"] = [pos_extender_point, pos_whip]
                
                if neg_extender_point and neg_whip:
                    cable_routes[f"neg_extender_{tracker_idx}"] = [neg_extender_point, neg_whip]
            
            # Whip to device routes
            pos_dest, neg_dest = self.get_device_destination_points()
            if pos_whip and pos_dest:
                cable_routes[f"pos_main_{tracker_idx}"] = [pos_whip, pos_dest]
            
            if neg_whip and neg_dest:
                cable_routes[f"neg_main_{tracker_idx}"] = [neg_whip, neg_dest]

    def on_canvas_click(self, event):
        """Handle canvas click events for whip point selection"""
        if self.panning:  # Don't select while panning
            return
            
        # Convert canvas coordinates to world coordinates
        scale = self.get_canvas_scale()
        world_x = (event.x - 20 - self.pan_x) / scale
        world_y = (event.y - 20 - self.pan_y) / scale
        
        # Check if we clicked on a whip point
        hit_whip = self.get_whip_at_position(world_x, world_y)
        
        if hit_whip:
            # Toggle selection if Ctrl is pressed, otherwise select only this one
            if event.state & 0x4:  # Ctrl key is pressed
                if hit_whip in self.selected_whips:
                    self.selected_whips.remove(hit_whip)
                else:
                    self.selected_whips.add(hit_whip)
            else:
                # Clear previous selection unless Shift is pressed
                if not (event.state & 0x1):  # Shift key is not pressed
                    self.selected_whips.clear()
                self.selected_whips.add(hit_whip)
                
            self.drag_whips = True
            self.drag_start_x = event.x
            self.drag_start_y = event.y
        else:
            # Start selection box if not clicking on a whip point
            if not (event.state & 0x4) and not (event.state & 0x1):  # Neither Ctrl nor Shift
                self.selected_whips.clear()
            self.drag_start_x = event.x
            self.drag_start_y = event.y
            self.dragging = True
            
        self.draw_wiring_layout()

    def on_canvas_drag(self, event):
        """Handle dragging of whip points or selection box"""
        if self.panning:
            return
            
        if self.drag_whips and self.selected_whips:
            # Convert current mouse position to world coordinates
            current_world_x, current_world_y = self.canvas_to_world(event.x, event.y)
            
            # For each selected whip, calculate target position with constraints and snapping
            for whip_info in self.selected_whips:
                self.update_whip_to_target_position(whip_info, current_world_x, current_world_y)
            
            # Force route recalculation with new whip positions
            self.draw_wiring_layout()
        elif self.dragging:
            # Update selection box
            if self.selection_box:
                self.canvas.delete(self.selection_box)
                
            self.selection_box = self.canvas.create_rectangle(
                self.drag_start_x, self.drag_start_y, event.x, event.y,
                outline='blue', dash=(4, 4)
            )

    def update_whip_to_target_position(self, whip_info, target_world_x, target_world_y):
        """Update whip point to target position with all constraints applied"""
        # Get original position when drag started
        if len(whip_info) == 3:
            tracker_id, harness_idx, polarity = whip_info
            original_pos = self.get_harness_whip_current_stored_position(tracker_id, harness_idx, polarity)
            if not original_pos:
                original_pos = self.get_realistic_whip_position(tracker_id, polarity, harness_idx)
        else:
            tracker_id, polarity = whip_info
            original_pos = self.get_whip_current_stored_position(tracker_id, polarity)
            if not original_pos:
                original_pos = self.get_realistic_whip_position(tracker_id, polarity)
        
        if not original_pos:
            return
        
        # In realistic mode, preserve the original centerline X and only allow Y movement
        if self.routing_mode_var.get() == "realistic":
            target_x = original_pos[0]  # Keep the original centerline X position
            target_y = target_world_y   # Allow Y movement
        else:
            # In conceptual mode, allow both X and Y movement
            target_x = target_world_x
            target_y = target_world_y
        
        # Apply N/S constraint if enabled (only allow Y movement)
        if self.right_angle_var.get():
            target_x = original_pos[0]  # Keep original X, only allow Y movement
        
        # Apply 5ft snapping if enabled
        if self.snap_5ft_var.get():
            snap_increment = 1.524  # 5 feet in meters
            target_y = round(target_y / snap_increment) * snap_increment
        
        # Store the new position
        if len(whip_info) == 3:
            self.store_harness_whip_position(tracker_id, harness_idx, polarity, target_x, target_y)
        else:
            self.store_regular_whip_position(tracker_id, polarity, target_x, target_y)

    def store_regular_whip_position(self, tracker_id, polarity, x, y):
        """Store regular whip point position"""
        # Ensure custom_whip_points exists
        if not hasattr(self.block, 'wiring_config') or not self.block.wiring_config:
            from ..models.block import WiringConfig, WiringType
            self.block.wiring_config = WiringConfig(wiring_type=WiringType.HOMERUN)
        
        if not hasattr(self.block.wiring_config, 'custom_whip_points'):
            self.block.wiring_config.custom_whip_points = {}
        
        if tracker_id not in self.block.wiring_config.custom_whip_points:
            self.block.wiring_config.custom_whip_points[tracker_id] = {}
        
        self.block.wiring_config.custom_whip_points[tracker_id][polarity] = (x, y)

    def store_harness_whip_position(self, tracker_id, harness_idx, polarity, x, y):
        """Store harness whip point position"""
        # Ensure custom_harness_whip_points exists
        if not hasattr(self.block, 'wiring_config') or not self.block.wiring_config:
            from ..models.block import WiringConfig, WiringType
            self.block.wiring_config = WiringConfig(wiring_type=WiringType.HOMERUN)
        
        if not hasattr(self.block.wiring_config, 'custom_harness_whip_points'):
            self.block.wiring_config.custom_harness_whip_points = {}
        
        if tracker_id not in self.block.wiring_config.custom_harness_whip_points:
            self.block.wiring_config.custom_harness_whip_points[tracker_id] = {}
        if harness_idx not in self.block.wiring_config.custom_harness_whip_points[tracker_id]:
            self.block.wiring_config.custom_harness_whip_points[tracker_id][harness_idx] = {}
        
        self.block.wiring_config.custom_harness_whip_points[tracker_id][harness_idx][polarity] = (x, y)

    def apply_movement_constraints(self, dx, dy):
        """Apply movement constraints based on routing mode and user settings"""
        # In realistic mode, only allow Y-axis movement (along centerline)
        if self.routing_mode_var.get() == "realistic":
            dx = 0  # Force to centerline, no X movement allowed
        
        # Apply North/South constraint if enabled (vertical movement only)
        if self.right_angle_var.get():
            dx = 0  # Constrain to vertical movement only
        
        return dx, dy

    def apply_5ft_snapping(self, dx, dy):
        """Apply 5-foot increment snapping - but we'll handle this in the position update instead"""
        # For now, just return the original deltas
        # The actual snapping will happen in the position update methods
        return dx, dy
    
    def update_whip_point_position_with_constraints(self, whip_info, dx, dy):
        """Update whip point position with realistic mode centerline constraints"""
        if len(whip_info) == 3:
            tracker_id, harness_idx, polarity = whip_info
            self.update_harness_whip_position_with_constraints(tracker_id, harness_idx, polarity, dx, dy)
        else:
            tracker_id, polarity = whip_info
            self.update_regular_whip_position_with_constraints(tracker_id, polarity, dx, dy)

    def update_regular_whip_position_with_constraints(self, tracker_id, polarity, dx, dy):
        """Update regular whip point position with centerline constraints for realistic mode"""
        # Get current position (custom or default)
        current_pos = self.get_whip_current_stored_position(tracker_id, polarity)
        if not current_pos:
            current_pos = self.get_whip_default_position(tracker_id, polarity)
        
        if not current_pos:
            return
        
        # Calculate new position
        new_x = current_pos[0] + dx
        new_y = current_pos[1] + dy
        
        # In realistic mode, force X to centerline
        if self.routing_mode_var.get() == "realistic":
            centerline_pos = self.get_realistic_whip_position(tracker_id, polarity)
            if centerline_pos:
                new_x = centerline_pos[0]  # Force to centerline X
        
        # Apply 5ft snapping to final position if enabled
        if self.snap_5ft_var.get():
            snap_increment = 1.524  # 5 feet in meters
            new_y = round(new_y / snap_increment) * snap_increment
        
        # Ensure custom_whip_points exists
        if not hasattr(self.block, 'wiring_config') or not self.block.wiring_config:
            from ..models.block import WiringConfig, WiringType
            self.block.wiring_config = WiringConfig(wiring_type=WiringType.HOMERUN)
        
        if not hasattr(self.block.wiring_config, 'custom_whip_points'):
            self.block.wiring_config.custom_whip_points = {}
        
        # Store the new position
        if tracker_id not in self.block.wiring_config.custom_whip_points:
            self.block.wiring_config.custom_whip_points[tracker_id] = {}
        
        self.block.wiring_config.custom_whip_points[tracker_id][polarity] = (new_x, new_y)

    def update_regular_whip_position_with_constraints(self, tracker_id, polarity, dx, dy):
        """Update regular whip point position with centerline constraints for realistic mode"""
        # Get current position (custom or default)
        current_pos = self.get_whip_current_stored_position(tracker_id, polarity)
        if not current_pos:
            current_pos = self.get_whip_default_position(tracker_id, polarity)
        
        if not current_pos:
            return
        
        # Calculate new position
        new_x = current_pos[0] + dx
        new_y = current_pos[1] + dy
        
        # In realistic mode, force X to centerline
        if self.routing_mode_var.get() == "realistic":
            centerline_pos = self.get_realistic_whip_position(tracker_id, polarity)
            if centerline_pos:
                new_x = centerline_pos[0]  # Force to centerline X
        
        # Ensure custom_whip_points exists
        if not hasattr(self.block, 'wiring_config') or not self.block.wiring_config:
            from ..models.block import WiringConfig, WiringType
            self.block.wiring_config = WiringConfig(wiring_type=WiringType.HOMERUN)
        
        if not hasattr(self.block.wiring_config, 'custom_whip_points'):
            self.block.wiring_config.custom_whip_points = {}
        
        # Store the new position
        if tracker_id not in self.block.wiring_config.custom_whip_points:
            self.block.wiring_config.custom_whip_points[tracker_id] = {}
        
        self.block.wiring_config.custom_whip_points[tracker_id][polarity] = (new_x, new_y)

    def update_harness_whip_position_with_constraints(self, tracker_id, harness_idx, polarity, dx, dy):
        """Update harness-specific whip point position with centerline constraints"""
        # Get current position (custom or default)
        current_pos = self.get_harness_whip_current_stored_position(tracker_id, harness_idx, polarity)
        if not current_pos:
            current_pos = self.get_whip_default_position(tracker_id, polarity, harness_idx)
        
        if not current_pos:
            return
        
        # Calculate new position
        new_x = current_pos[0] + dx
        new_y = current_pos[1] + dy
        
        # In realistic mode, force X to centerline
        if self.routing_mode_var.get() == "realistic":
            centerline_pos = self.get_realistic_whip_position(tracker_id, polarity, harness_idx)
            if centerline_pos:
                new_x = centerline_pos[0]  # Force to centerline X
        
        # Ensure custom_harness_whip_points exists
        if not hasattr(self.block, 'wiring_config') or not self.block.wiring_config:
            from ..models.block import WiringConfig, WiringType
            self.block.wiring_config = WiringConfig(wiring_type=WiringType.HOMERUN)
        
        if not hasattr(self.block.wiring_config, 'custom_harness_whip_points'):
            self.block.wiring_config.custom_harness_whip_points = {}
        
        # Store the new position
        if tracker_id not in self.block.wiring_config.custom_harness_whip_points:
            self.block.wiring_config.custom_harness_whip_points[tracker_id] = {}
        if harness_idx not in self.block.wiring_config.custom_harness_whip_points[tracker_id]:
            self.block.wiring_config.custom_harness_whip_points[tracker_id][harness_idx] = {}
        
        self.block.wiring_config.custom_harness_whip_points[tracker_id][harness_idx][polarity] = (new_x, new_y)

    def get_whip_current_stored_position(self, tracker_id, polarity):
        """Get the currently stored custom position for a whip point"""
        if (hasattr(self.block, 'wiring_config') and 
            self.block.wiring_config and 
            hasattr(self.block.wiring_config, 'custom_whip_points') and
            tracker_id in self.block.wiring_config.custom_whip_points and
            polarity in self.block.wiring_config.custom_whip_points[tracker_id]):
            return self.block.wiring_config.custom_whip_points[tracker_id][polarity]
        return None

    def get_harness_whip_current_stored_position(self, tracker_id, harness_idx, polarity):
        """Get the currently stored custom position for a harness whip point"""
        if (hasattr(self.block, 'wiring_config') and 
            self.block.wiring_config and 
            hasattr(self.block.wiring_config, 'custom_harness_whip_points') and
            tracker_id in self.block.wiring_config.custom_harness_whip_points and
            harness_idx in self.block.wiring_config.custom_harness_whip_points[tracker_id] and
            polarity in self.block.wiring_config.custom_harness_whip_points[tracker_id][harness_idx]):
            return self.block.wiring_config.custom_harness_whip_points[tracker_id][harness_idx][polarity]
        return None

    def on_canvas_release(self, event):
        """Handle end of drag operation"""
        if self.panning:
            return
            
        if self.drag_whips:
            self.drag_whips = False
        elif self.dragging and self.selection_box:
            # Convert selection box to world coordinates and select whips within it
            x1 = min(self.drag_start_x, event.x)
            y1 = min(self.drag_start_y, event.y)
            x2 = max(self.drag_start_x, event.x)
            y2 = max(self.drag_start_y, event.y)
            
            # Convert to world coordinates
            wx1, wy1 = self.canvas_to_world(x1, y1)
            wx2, wy2 = self.canvas_to_world(x2, y2)
            
            # Select all whip points within this box
            for tracker_idx, pos in enumerate(self.block.tracker_positions):
                tracker_id = str(tracker_idx)
                
                # Check positive whip
                pos_whip = self.get_whip_position(tracker_id, 'positive')
                if pos_whip:
                    if wx1 <= pos_whip[0] <= wx2 and wy1 <= pos_whip[1] <= wy2:
                        self.selected_whips.add((tracker_id, 'positive'))
                        
                # Check negative whip
                neg_whip = self.get_whip_position(tracker_id, 'negative')
                if neg_whip:
                    if wx1 <= neg_whip[0] <= wx2 and wy1 <= neg_whip[1] <= wy2:
                        self.selected_whips.add((tracker_id, 'negative'))
            
            # Delete selection box
            self.canvas.delete(self.selection_box)
            self.selection_box = None
            
        self.dragging = False
        self.draw_wiring_layout()

    def reset_selected_whips(self, event=None):
        """Reset selected whip points to their default positions"""
        if not self.selected_whips:
            return
            
        for whip_info in self.selected_whips:
            if len(whip_info) == 3:
                # Harness whip point: (tracker_id, harness_idx, polarity)
                tracker_id, harness_idx, polarity = whip_info
                if (hasattr(self.block, 'wiring_config') and 
                    self.block.wiring_config and 
                    hasattr(self.block.wiring_config, 'custom_harness_whip_points') and
                    tracker_id in self.block.wiring_config.custom_harness_whip_points and
                    harness_idx in self.block.wiring_config.custom_harness_whip_points[tracker_id] and
                    polarity in self.block.wiring_config.custom_harness_whip_points[tracker_id][harness_idx]):
                    
                    # Remove custom harness whip position
                    del self.block.wiring_config.custom_harness_whip_points[tracker_id][harness_idx][polarity]
                    
                    # Clean up empty entries
                    if not self.block.wiring_config.custom_harness_whip_points[tracker_id][harness_idx]:
                        del self.block.wiring_config.custom_harness_whip_points[tracker_id][harness_idx]
                    if not self.block.wiring_config.custom_harness_whip_points[tracker_id]:
                        del self.block.wiring_config.custom_harness_whip_points[tracker_id]
            else:
                # Regular whip point: (tracker_id, polarity)
                tracker_id, polarity = whip_info
                if (hasattr(self.block, 'wiring_config') and 
                    self.block.wiring_config and 
                    hasattr(self.block.wiring_config, 'custom_whip_points') and
                    tracker_id in self.block.wiring_config.custom_whip_points and
                    polarity in self.block.wiring_config.custom_whip_points[tracker_id]):
                    
                    # Remove custom position
                    del self.block.wiring_config.custom_whip_points[tracker_id][polarity]
                    
                    # Clean up empty entries
                    if not self.block.wiring_config.custom_whip_points[tracker_id]:
                        del self.block.wiring_config.custom_whip_points[tracker_id]
                        
        self.draw_wiring_layout()

    def reset_all_whips(self, event=None):
        """Reset all whip points to their default positions"""
        if hasattr(self.block, 'wiring_config') and self.block.wiring_config:
            self.block.wiring_config.custom_whip_points = {}
            if hasattr(self.block.wiring_config, 'custom_harness_whip_points'):
                self.block.wiring_config.custom_harness_whip_points = {}
                    
        self.selected_whips.clear()
        self.draw_wiring_layout()
        
    def select_all_whips(self, event=None):
        """Select all whip points"""
        self.selected_whips.clear()
        
        for tracker_idx, _ in enumerate(self.block.tracker_positions):
            tracker_id = str(tracker_idx)
            self.selected_whips.add((tracker_id, 'positive'))
            self.selected_whips.add((tracker_id, 'negative'))
            
        self.draw_wiring_layout()
        return "break"  # Prevent default Ctrl+A behavior
        
    def clear_selection(self, event=None):
        """Clear whip point selection"""
        self.selected_whips.clear()
        self.draw_wiring_layout()

    def get_whip_position(self, tracker_id, polarity, harness_idx=None):
        """Get the current position of a whip point (custom or default)"""
        # Check for custom harness whip points first
        if (harness_idx is not None and 
            hasattr(self.block, 'wiring_config') and 
            self.block.wiring_config and 
            hasattr(self.block.wiring_config, 'custom_harness_whip_points') and
            tracker_id in self.block.wiring_config.custom_harness_whip_points and
            harness_idx in self.block.wiring_config.custom_harness_whip_points[tracker_id] and
            polarity in self.block.wiring_config.custom_harness_whip_points[tracker_id][harness_idx]):
            return self.block.wiring_config.custom_harness_whip_points[tracker_id][harness_idx][polarity]
        
        # Check for custom tracker whip points
        if (hasattr(self.block, 'wiring_config') and 
            self.block.wiring_config and 
            hasattr(self.block.wiring_config, 'custom_whip_points') and
            tracker_id in self.block.wiring_config.custom_whip_points and
            polarity in self.block.wiring_config.custom_whip_points[tracker_id]):
            return self.block.wiring_config.custom_whip_points[tracker_id][polarity]
        
        # Fall back to default position based on routing mode
        if hasattr(self, 'routing_mode_var') and self.routing_mode_var.get() == "realistic":
            return self.get_realistic_whip_position(tracker_id, polarity, harness_idx)
        else:
            return self.get_whip_default_position(tracker_id, polarity, harness_idx)

    def get_realistic_whip_position(self, tracker_id, polarity, harness_idx=None):
        """Get whip position aligned with harness cable for realistic mode"""
        tracker_idx = int(tracker_id)
        if tracker_idx < 0 or tracker_idx >= len(self.block.tracker_positions):
            return None
            
        pos = self.block.tracker_positions[tracker_idx]
        
        # Get tracker dimensions and calculate centerline
        dims = pos.template.get_physical_dimensions()
        tracker_length = dims[0]
        tracker_width = dims[1]
        tracker_center_x = pos.x + (tracker_width / 2)
        
        # IMPORTANT: Use EXACT same offsets as harness drawing
        pos_offset = -0.5  # Same as harness offset
        neg_offset = 0.5   # Same as harness offset
        
        if polarity == 'positive':
            whip_x = tracker_center_x + pos_offset  # Align with positive harness
        else:
            whip_x = tracker_center_x + neg_offset  # Align with negative harness
        
        # Get tracker Y boundaries
        tracker_north_y = pos.y
        tracker_south_y = pos.y + tracker_length
        device_y = self.block.device_y

        # Check if this tracker has multiple harnesses
        tracker_idx = int(tracker_id)
        string_count = len(pos.strings)
        has_multiple_harnesses = (hasattr(self.block, 'wiring_config') and 
                                self.block.wiring_config and 
                                hasattr(self.block.wiring_config, 'harness_groupings') and
                                string_count in self.block.wiring_config.harness_groupings and
                                len(self.block.wiring_config.harness_groupings[string_count]) > 1)

        # Place whip point based on device location and harness configuration
        whip_offset = 0.5  # Offset from tracker end in meters

        if device_y < tracker_north_y:
            # Device north of tracker - place whip at north end
            y_position = tracker_north_y - whip_offset
        elif device_y > tracker_south_y:
            # Device south of tracker - place whip at south end
            y_position = tracker_south_y + whip_offset
        else:
            # Device in middle of tracker
            if has_multiple_harnesses:
                # Multiple harnesses - whip can be at device level
                y_position = device_y
            else:
                # Single harness - whip must be at tracker end (choose closer end)
                distance_to_north = abs(device_y - tracker_north_y)
                distance_to_south = abs(device_y - tracker_south_y)
                
                if distance_to_north <= distance_to_south:
                    y_position = tracker_north_y - whip_offset
                else:
                    y_position = tracker_south_y + whip_offset
        
        result = (whip_x, y_position)
        return result
    
    def get_whip_default_position(self, tracker_id, polarity, harness_idx=None):
        """Calculate the default position for a whip point"""
        tracker_idx = int(tracker_id)
        if tracker_idx < 0 or tracker_idx >= len(self.block.tracker_positions):
            return None
            
        pos = self.block.tracker_positions[tracker_idx]
        is_positive = (polarity == 'positive')
        
        # Use the existing calculation method
        whip_pos = self.calculate_whip_points(pos, is_positive)
        
        # If this is a harness-specific whip, apply an offset
        if harness_idx is not None:
            vertical_spacing = 0.3  # Same as in draw_custom_harnesses
            if whip_pos:
                return (whip_pos[0], whip_pos[1] + harness_idx * vertical_spacing)
        
        return whip_pos
            
    def get_whip_at_position(self, x, y, tolerance=0.5):
        """Find a whip point at the given position within tolerance"""
        # Check each tracker's harness whip points
        for tracker_idx, pos in enumerate(self.block.tracker_positions):
            tracker_id = str(tracker_idx)
            string_count = len(pos.strings)
            
            if (hasattr(self.block, 'wiring_config') and 
                self.block.wiring_config and 
                hasattr(self.block.wiring_config, 'harness_groupings') and
                string_count in self.block.wiring_config.harness_groupings):
                
                for harness_idx, _ in enumerate(self.block.wiring_config.harness_groupings[string_count]):
                    # Check positive harness whip
                    pos_whip = self.get_whip_position(tracker_id, 'positive', harness_idx)
                    if pos_whip:
                        if abs(x - pos_whip[0]) <= tolerance and abs(y - pos_whip[1]) <= tolerance:
                            return (tracker_id, harness_idx, 'positive')
                    
                    # Check negative harness whip
                    neg_whip = self.get_whip_position(tracker_id, 'negative', harness_idx)
                    if neg_whip:
                        if abs(x - neg_whip[0]) <= tolerance and abs(y - pos_whip[1]) <= tolerance:
                            return (tracker_id, harness_idx, 'negative')
            
            # Check regular tracker whip points
            pos_whip = self.get_whip_position(tracker_id, 'positive')
            if pos_whip:
                if abs(x - pos_whip[0]) <= tolerance and abs(y - pos_whip[1]) <= tolerance:
                    return (tracker_id, 'positive')
                    
            neg_whip = self.get_whip_position(tracker_id, 'negative')
            if neg_whip:
                if abs(x - neg_whip[0]) <= tolerance and abs(y - neg_whip[1]) <= tolerance:
                    return (tracker_id, 'negative')
                    
        return None
    
    def get_harness_collection_point(self, tracker_pos, string, harness_idx, is_positive, routing_mode):
        """Get harness collection point coordinates based on routing mode"""
        if routing_mode == "realistic":
            # Realistic: Use tracker centerline with small offset
            dims = tracker_pos.template.get_physical_dimensions()
            tracker_width = dims[1]
            tracker_center_x = tracker_pos.x + (tracker_width / 2)
            
            offset = -0.5 if is_positive else 0.5
            node_x = tracker_center_x + offset
            # Check wiring mode
            wiring_mode = self.project.wiring_mode if self.project and hasattr(self.project, 'wiring_mode') else 'daisy_chain'
            
            # In leapfrog mode, both use positive source Y
            if wiring_mode == 'leapfrog':
                node_y = tracker_pos.y + string.positive_source_y
            else:
                node_y = tracker_pos.y + (string.positive_source_y if is_positive else string.negative_source_y)
            
            return (node_x, node_y)
        else:
            # Conceptual: Use horizontal offset but NO harness-specific vertical spacing
            # Each harness positions itself relative to its own strings naturally
            horizontal_offset = 0.6
            
            # Check wiring mode
            wiring_mode = self.project.wiring_mode if self.project and hasattr(self.project, 'wiring_mode') else 'daisy_chain'
            
            if is_positive:
                node_x = tracker_pos.x + string.positive_source_x - horizontal_offset
                node_y = tracker_pos.y + string.positive_source_y
            else:
                node_x = tracker_pos.x + string.negative_source_x + horizontal_offset
                # In leapfrog mode, use positive source Y (top) for negative too
                if wiring_mode == 'leapfrog':
                    node_y = tracker_pos.y + string.positive_source_y
                else:
                    node_y = tracker_pos.y + string.negative_source_y
                    
            return (node_x, node_y)

    def get_harness_whip_point(self, tracker_id, harness_idx, polarity, routing_mode):
        """Get harness-specific whip point coordinates based on routing mode"""
        # First check for custom positions
        if (hasattr(self.block, 'wiring_config') and 
            self.block.wiring_config and 
            hasattr(self.block.wiring_config, 'custom_harness_whip_points') and
            tracker_id in self.block.wiring_config.custom_harness_whip_points and
            harness_idx in self.block.wiring_config.custom_harness_whip_points[tracker_id] and
            polarity in self.block.wiring_config.custom_harness_whip_points[tracker_id][harness_idx]):
            return self.block.wiring_config.custom_harness_whip_points[tracker_id][harness_idx][polarity]
        
        # For multi-harness, each harness gets its natural whip point based on its strings
        # No artificial vertical spacing - let each harness position naturally
        tracker_idx = int(tracker_id)
        pos = self.block.tracker_positions[tracker_idx]
        string_count = len(pos.strings)
        
        # Get this specific harness's string indices
        if (hasattr(self.block.wiring_config, 'harness_groupings') and 
            string_count in self.block.wiring_config.harness_groupings and
            harness_idx < len(self.block.wiring_config.harness_groupings[string_count])):
            
            harness = self.block.wiring_config.harness_groupings[string_count][harness_idx]
            string_indices = harness.string_indices
            
            if string_indices:
                # Calculate whip point based on this harness's string positions
                avg_string_y = sum(pos.strings[i].positive_source_y if polarity == 'positive' 
                                else pos.strings[i].negative_source_y 
                                for i in string_indices if i < len(pos.strings)) / len(string_indices)
                
                # Get base whip position and adjust for this harness
                base_whip = self.get_whip_position(tracker_id, polarity)
                if base_whip:
                    return (base_whip[0], pos.y + avg_string_y + self.calculate_whip_offset_for_harness(pos, string_indices))
        
        # Fallback to base whip position
        return self.get_whip_position(tracker_id, polarity)

    def calculate_whip_offset_for_harness(self, pos, string_indices):
        """Calculate appropriate whip offset based on harness string positions"""
        # This creates a small offset so whip points don't overlap exactly
        # but keeps each harness positioned relative to its strings
        if not string_indices:
            return 0.0
        
        # Use the average position of the strings in this harness
        avg_string_pos = sum(string_indices) / len(string_indices)
        # Small offset based on string position to avoid exact overlap
        return avg_string_pos * 0.1  # 10cm per string index difference

    def get_string_source_to_harness_route(self, source_point, harness_point, routing_mode):
        """Get route from string source to harness collection point"""
        if routing_mode == "realistic":
            # Realistic: Follow centerline then horizontal to harness point
            tracker_center_x = (source_point[0] + harness_point[0]) / 2  # Approximate centerline
            return [
                source_point,
                (tracker_center_x, source_point[1]),
                (tracker_center_x, harness_point[1]),
                harness_point
            ]
        else:
            # Conceptual: Direct connection
            return [source_point, harness_point]

    def show_context_menu(self, event):
        """Show context menu for whip points"""
        # First, check if we clicked on a whip point
        world_x, world_y = self.canvas_to_world(event.x, event.y)
        
        hit_whip = self.get_whip_at_position(world_x, world_y)
        
        if hit_whip:
            # If hitting a whip point that's not selected, select only this one
            if hit_whip not in self.selected_whips:
                self.selected_whips.clear()
                self.selected_whips.add(hit_whip)
                self.draw_wiring_layout()
            
            # Show the context menu
            self.whip_menu.post(event.x_root, event.y_root)
            return "break"  # Prevent default right-click behavior
    
    def validate_mppt_capacity(self):
        """Check if the inverter MPPT can handle the string current"""
        if not self.block or not self.block.inverter:
            return True
            
        # Calculate total strings
        total_strings = sum(len(pos.strings) for pos in self.block.tracker_positions)
        
        # Get template
        template = self.block.tracker_template
        if not template or not template.module_spec:
            return True
        
        # Calculate per-string current
        string_current = template.module_spec.imp
        
        # Calculate total current 
        total_current = string_current * total_strings
        
        # Get total MPPT capacity
        total_mppt_capacity = sum(ch.max_input_current for ch in self.block.inverter.mppt_channels)
        
        return total_current <= total_mppt_capacity
    
    def add_wire_warning(self, wire_id, message, severity):
        """Add a warning message to the warning panel
        
        Args:
            wire_id: ID of the wire this warning refers to
            message: Warning message text
            severity: 'caution', 'warning' or 'overload'
        """
        # Create unique ID for this warning
        warning_id = f"warning_{len(self.wire_warnings) + 1}"
        
        # Define colors based on severity
        if severity == 'overload':
            bg_color = '#ffdddd'
            icon = "⚠"
        elif severity == 'warning':
            bg_color = '#fff0dd'
            icon = "⚠"
        else:  # caution
            bg_color = '#ffffdd'
            icon = "⚠"
        
        # Create warning entry in panel
        warning_frame = tk.Frame(self.warning_frame, bg=bg_color, padx=2, pady=2)
        warning_frame.pack(side=tk.TOP, fill=tk.X, padx=2, pady=1)
        
        # Add warning text
        warning_label = tk.Label(warning_frame, text=f"{icon} {message}", 
                                bg=bg_color, anchor='w', justify='left')
        warning_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Store mapping between warning and wire
        self.wire_warnings[warning_id] = wire_id
        
        # Bind click event to highlight the wire
        warning_frame.bind("<Button-1>", lambda e, wid=wire_id: self.highlight_wire(wid))
        warning_label.bind("<Button-1>", lambda e, wid=wire_id: self.highlight_wire(wid))
        
        return warning_id

    def highlight_wire(self, wire_id):
        """Highlight a wire when its warning is clicked"""
        self.canvas.delete("wire_highlight")
        
        wire_coords = self.canvas.coords(wire_id)
        if not wire_coords:
            return
        
        # Create highlight with pulsing effect
        highlight_width = self.get_wire_highlight_width(wire_id)
        highlight_color = '#FF00FF'  # Bright magenta
        
        self.canvas.create_line(
            wire_coords,
            width=highlight_width,
            fill=highlight_color,
            dash=(8, 4),
            tags="wire_highlight"
        )
        
        # Start pulsing animation
        self.start_wire_pulse_animation(highlight_width)
        
        # Scroll to wire if needed
        self.scroll_to_wire_if_needed(wire_coords)

    def get_wire_highlight_width(self, wire_id):
        """Get appropriate highlight width for a wire"""
        try:
            current_width = int(float(self.canvas.itemcget(wire_id, 'width')))
        except (ValueError, TypeError):
            current_width = 2
        return current_width * 2 + 4

    def start_wire_pulse_animation(self, base_width):
        """Start the pulsing animation for highlighted wire"""
        self.pulse_highlight(6, base_width)

    def scroll_to_wire_if_needed(self, wire_coords):
        """Scroll canvas to show the wire if it's outside visible area"""
        if len(wire_coords) >= 2:
            mid_x = (wire_coords[0] + wire_coords[2]) / 2
            mid_y = (wire_coords[1] + wire_coords[3]) / 2
            
            # Get canvas dimensions
            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()
            
            # Check if the point is outside the visible area
            if mid_x < 0 or mid_x > canvas_width or mid_y < 0 or mid_y > canvas_height:
                # Adjust pan to center on the point
                self.pan_x = canvas_width/2 - mid_x + 20
                self.pan_y = canvas_height/2 - mid_y + 20
                
                # Redraw with new pan values
                self.draw_wiring_layout()

    def pulse_highlight(self, count, base_width):
        """Create a pulsing effect for the highlight"""
        if count <= 0 or not self.canvas.find_withtag("wire_highlight"):
            return
        
        # Calculate width based on sine wave (pulsing effect)
        import math
        phase = count % 6  # 0 to 5
        pulse_factor = 1 + 0.5 * math.sin(phase * math.pi / 3)  # 1.0 to 1.5
        new_width = int(base_width * pulse_factor)
        
        # Update line width
        self.canvas.itemconfigure("wire_highlight", width=new_width)
        
        # Schedule next pulse
        self.after(150, lambda: self.pulse_highlight(count - 1, base_width))

    def clear_warnings(self):
        """Clear all warnings from the panel"""
        for widget in self.warning_frame.winfo_children():
            widget.destroy()
        self.wire_warnings = {}

    def draw_wire_segment(self, points, wire_gauge, current, is_positive=True, segment_type="string", context_info=None):
        """Draw a wire segment with warnings only for overloads"""
        # Get standard wire properties
        line_thickness = self.get_line_thickness_for_wire(wire_gauge)
        base_color = self.get_wire_color(segment_type, is_positive)
        
        # Draw the wire with standard color
        line_id = self.canvas.create_line(points, fill=base_color, width=line_thickness,
                                        tags=f"wire_{segment_type}")
        
        # Skip warnings if ampacity can't be determined
        ampacity = get_ampacity_for_wire_gauge(wire_gauge)
        if ampacity == 0:
            return line_id
        
        # Calculate load percentage
        nec_current = calculate_nec_current(current)
        load_percent = (nec_current / ampacity) * 100
        
        # Only add warning if it's an overload (>100%)
        if load_percent > 100:
            polarity = "positive" if is_positive else "negative"
            
            # Use context info if provided, otherwise use default
            if context_info:
                warning_text = f"{context_info} {polarity} {wire_gauge}: {load_percent:.0f}% (OVERLOAD)"
            else:
                warning_text = f"{polarity.capitalize()} {segment_type} {wire_gauge}: {load_percent:.0f}% (OVERLOAD)"
            
            self.add_wire_warning(line_id, warning_text, 'overload')
            
            # Optional: Make the line dashed to indicate overload
            self.canvas.itemconfig(line_id, dash=(5, 3))
        
        return line_id
    
    def setup_warning_panel(self):
        """Create warning panel in bottom right corner of canvas"""
        # Create warning panel frame
        self.warning_panel = tk.Frame(self.canvas, bg='white', bd=1, relief=tk.RAISED)
        self.warning_panel.place(relx=1.0, rely=1.0, x=-5, y=-5, anchor='se')
        
        # Add header - changed to be more specific
        self.warning_header = tk.Label(self.warning_panel, text="Wire Overload Issues", 
                                    bg='#ffdddd', font=('Arial', 9, 'bold'),
                                    padx=5, pady=2, anchor='w', width=30)
        self.warning_header.pack(side=tk.TOP, fill=tk.X)
        
        # Add scrollable warning list
        self.warning_frame = tk.Frame(self.warning_panel, bg='white')
        self.warning_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        # Initialize warnings dictionary if it doesn't exist
        if not hasattr(self, 'wire_warnings'):
            self.wire_warnings = {}

    def update_ui_for_wiring_type(self):
        """Update UI elements based on selected wiring type"""
        is_harness = self.wiring_type_var.get() == WiringType.HARNESS.value
        if is_harness:
            self.harness_frame.grid()
            self.extender_frame.grid()
            self.harness_collapsible.grid()
            
            # Hide Default Cable Sizes section in harness mode
            if hasattr(self, 'cable_collapsible'):
                self.cable_collapsible.grid_remove()
            
            # Show harness cable configuration table
            if hasattr(self, 'cable_table_collapsible'):
                self.cable_table_collapsible.grid()
                self.update_harness_cable_table()
                
            self.populate_string_count_combobox()

        else:
            self.harness_frame.grid_remove()
            self.extender_frame.grid_remove()
            self.harness_collapsible.grid_remove()
            
            # Show Default Cable Sizes section in string homerun mode
            if hasattr(self, 'cable_collapsible'):
                self.cable_collapsible.grid()
            
            # Hide harness cable configuration table
            if hasattr(self, 'cable_table_collapsible'):
                self.cable_table_collapsible.grid_remove()

    def populate_string_count_combobox(self):
        """Populate the string count combobox with available tracker configurations"""
        # Count trackers by string count
        tracker_counts = {}
        for pos in self.block.tracker_positions:
            if not pos.strings:
                continue
            string_count = len(pos.strings)
            if string_count not in tracker_counts:
                tracker_counts[string_count] = 0
            tracker_counts[string_count] += 1
        
        # Generate descriptive items
        items = []
        self.string_count_mapping = {}  # Store mapping for display items to string counts
        
        for string_count, count in sorted(tracker_counts.items()):
            item = f"{string_count} string{'s' if string_count != 1 else ''} ({count} tracker{'s' if count != 1 else ''})"
            items.append(item)
            self.string_count_mapping[item] = string_count
        
        self.string_count_combobox['values'] = items
        if items:
            self.string_count_combobox.current(0)
            self.on_string_count_selected()

    def on_string_count_selected(self, event=None):
        """Handle string count selection from combobox"""
        if not self.string_count_var.get() or not hasattr(self, 'string_count_mapping'):
            return
        
        # Get selected string count
        selected_item = self.string_count_var.get()
        if selected_item not in self.string_count_mapping:
            return
        
        string_count = self.string_count_mapping[selected_item]
        
        # Count trackers with this string count
        tracker_count = sum(1 for pos in self.block.tracker_positions if len(pos.strings) == string_count)
        self.tracker_count_label.config(text=f"{tracker_count} tracker{'s' if tracker_count != 1 else ''}")
        
        # Clear string grouping frame
        for widget in self.string_check_frame.winfo_children():
            widget.destroy()
        
        # Create string checkboxes in columns of 8
        self.string_vars = []
        for i in range(string_count):
            var = tk.BooleanVar(value=False)
            self.string_vars.append(var)
            check = ttk.Checkbutton(self.string_check_frame, text=f"String {i+1}", variable=var)
            
            # Calculate column and row: new column every 8 strings
            column = i // 8
            row = i % 8
            check.grid(row=row, column=column, sticky=tk.W, padx=5, pady=2)
        
        # Update harness display
        self.update_harness_display(string_count)

    def update_harness_display(self, string_count):
        """Update the harness cable table after harness changes"""
        # Just update the cable configuration table
        self.update_harness_cable_table()

    def create_harness_from_selected(self):
        """Create a new harness from the selected strings"""
        if not self.string_count_var.get() or not hasattr(self, 'string_count_mapping'):
            return
        
        # Get selected string count
        selected_item = self.string_count_var.get()
        if selected_item not in self.string_count_mapping:
            return
        
        string_count = self.string_count_mapping[selected_item]
        
        # Get selected string indices
        selected_indices = [i for i, var in enumerate(self.string_vars) if var.get()]
        
        if not selected_indices:
            messagebox.showwarning("Warning", "No strings selected")
            return
        
        # Ensure wiring config and harness_groupings are initialized
        if not hasattr(self.block, 'wiring_config') or not self.block.wiring_config:
            self.block.wiring_config = WiringConfig(
                wiring_type=WiringType(self.wiring_type_var.get()),
                positive_collection_points=[],
                negative_collection_points=[],
                strings_per_collection={},
                cable_routes={},
                harness_groupings={}
            )
        
        if not hasattr(self.block.wiring_config, 'harness_groupings'):
            self.block.wiring_config.harness_groupings = {}
        
        if string_count not in self.block.wiring_config.harness_groupings:
            self.block.wiring_config.harness_groupings[string_count] = []
        
        # Check if any selected strings are already in a harness
        used_strings = []
        for harness in self.block.wiring_config.harness_groupings[string_count]:
            used_strings.extend(harness.string_indices)
        
        overlap = set(selected_indices) & set(used_strings)
        if overlap:
            overlap_strings = [str(i+1) for i in overlap]
            messagebox.showwarning("Warning", 
                                f"String(s) {', '.join(overlap_strings)} are already in a harness. "
                                "Please remove them from existing harnesses first.")
            return
        
        # Calculate recommended fuse size
        recommended_fuse = self.calculate_recommended_fuse_size(selected_indices)

        # Create new harness with default cable sizes
        new_harness = HarnessGroup(
            string_indices=selected_indices,
            cable_size="10 AWG",  # Default harness trunk size
            string_cable_size="10 AWG",  # Default string cable size
            extender_cable_size="8 AWG",  # Default extender size
            whip_cable_size="8 AWG",  # Default whip size
            fuse_rating_amps=recommended_fuse,
            use_fuse=len(selected_indices) > 1  # Only use fuses for 2+ strings
        )
        
        # Add to harness groupings
        self.block.wiring_config.harness_groupings[string_count].append(new_harness)
        
        # Update display
        self.update_harness_display(string_count)
        
        # Update the wiring visualization
        self.draw_wiring_layout()
        self.notify_wiring_changed()

    def delete_harness(self, string_count, harness_idx):
        """Delete a harness from the configuration"""
        if string_count in self.block.wiring_config.harness_groupings and \
        harness_idx < len(self.block.wiring_config.harness_groupings[string_count]):
            del self.block.wiring_config.harness_groupings[string_count][harness_idx]
            
            # If no harnesses left, remove the string count entry
            if not self.block.wiring_config.harness_groupings[string_count]:
                del self.block.wiring_config.harness_groupings[string_count]
                
            # Update display
            self.update_harness_display(string_count)
            
            # Update the wiring visualization
            self.draw_wiring_layout()
            self.notify_wiring_changed()

    def draw_device(self):
        """Draw inverter/combiner box"""
        if self.block.device_x is not None and self.block.device_y is not None:
            scale = self.get_canvas_scale()
            device_x, device_y = self.world_to_canvas(self.block.device_x, self.block.device_y)
            device_size = 0.91 * scale  # 3ft = 0.91m
            self.canvas.create_rectangle(
                device_x, device_y,
                device_x + device_size,
                device_y + device_size,
                fill='red', outline='darkred',
                tags='device'
            )
            
    def draw_trackers(self):
        """Draw all trackers with modules and source points"""
        # Ensure all tracker positions have project reference and are recalculated
        if self.project:
            for i, pos in enumerate(self.block.tracker_positions):
                pos._project_ref = self.project
                pos.calculate_string_positions()
        else:
            print("No project reference available in wiring configurator!")
        
        scale = self.get_canvas_scale()
        scale = self.get_canvas_scale()
        
        for pos in self.block.tracker_positions:
            if not pos.template:
                continue
                
            template = pos.template
            
            # Get base coordinates with pan offset
            x_base, y_base = self.world_to_canvas(pos.x, pos.y)
            
            # Get module dimensions based on orientation
            if template.module_orientation == ModuleOrientation.PORTRAIT:
                module_height = template.module_spec.width_mm / 1000
                module_width = template.module_spec.length_mm / 1000
            else:
                module_height = template.module_spec.length_mm / 1000
                module_width = template.module_spec.width_mm / 1000
                
            # Get physical dimensions for torque tube
            dims = template.get_physical_dimensions()
            
            # Draw torque tube through center
            self.canvas.create_line(
                x_base + module_width * scale/2, y_base,
                x_base + module_width * scale/2, y_base + dims[1] * scale,
                width=3, fill='gray'
            )
            
            # Handle different motor placement types
            if template.motor_placement_type == "middle_of_string":
                # Motor is in the middle of a specific string
                current_y = y_base
                
                for string_idx in range(template.strings_per_tracker):
                    if string_idx + 1 == template.motor_string_index:  # This string has the motor (1-based index)
                        # Draw north modules
                        for i in range(template.motor_split_north):
                            self.canvas.create_rectangle(
                                x_base, current_y,
                                x_base + module_width * scale, current_y + module_height * scale,
                                fill='lightblue', outline='blue'
                            )
                            current_y += (module_height + template.module_spacing_m) * scale
                        
                        # Draw motor gap with red circle
                        motor_y = current_y
                        gap_height = template.motor_gap_m * scale
                        circle_radius = min(gap_height / 3, module_width * scale / 4)
                        circle_center_x = x_base + module_width * scale / 2
                        circle_center_y = motor_y + gap_height / 2
                        
                        self.canvas.create_oval(
                            circle_center_x - circle_radius, circle_center_y - circle_radius,
                            circle_center_x + circle_radius, circle_center_y + circle_radius,
                            fill='red', outline='darkred', width=2
                        )
                        current_y += gap_height
                        
                        # Draw south modules
                        for i in range(template.motor_split_south):
                            self.canvas.create_rectangle(
                                x_base, current_y,
                                x_base + module_width * scale, current_y + module_height * scale,
                                fill='lightblue', outline='blue'
                            )
                            current_y += (module_height + template.module_spacing_m) * scale
                    else:
                        # Draw normal string without motor
                        for i in range(template.modules_per_string):
                            self.canvas.create_rectangle(
                                x_base, current_y,
                                x_base + module_width * scale, current_y + module_height * scale,
                                fill='lightblue', outline='blue'
                            )
                            current_y += (module_height + template.module_spacing_m) * scale
            else:
                # Original between_strings logic
                modules_per_string = template.modules_per_string
                motor_position = template.get_motor_position()
                strings_above_motor = motor_position
                strings_below_motor = template.strings_per_tracker - motor_position
                modules_above_motor = modules_per_string * strings_above_motor
                modules_below_motor = modules_per_string * strings_below_motor
                
                # Draw all modules
                y_pos = y_base
                
                # Draw modules above motor
                for i in range(modules_above_motor):
                    self.canvas.create_rectangle(
                        x_base, y_pos,
                        x_base + module_width * scale, 
                        y_pos + module_height * scale,
                        fill='lightblue', outline='blue'
                    )
                    y_pos += (module_height + template.module_spacing_m) * scale
                
                # Draw motor (only if there are strings below)
                if strings_below_motor > 0:
                    motor_y = y_pos
                    gap_height = template.motor_gap_m * scale
                    circle_radius = min(gap_height / 3, module_width * scale / 4)
                    circle_center_x = x_base + module_width * scale / 2
                    circle_center_y = motor_y + gap_height / 2
                    
                    self.canvas.create_oval(
                        circle_center_x - circle_radius, circle_center_y - circle_radius,
                        circle_center_x + circle_radius, circle_center_y + circle_radius,
                        fill='red', outline='darkred', width=2
                    )
                    y_pos += gap_height
                
                # Draw modules below motor
                for i in range(modules_below_motor):
                    self.canvas.create_rectangle(
                        x_base, y_pos,
                        x_base + module_width * scale,
                        y_pos + module_height * scale,
                        fill='lightblue', outline='blue'
                    )
                    y_pos += (module_height + template.module_spacing_m) * scale
        
            # Draw source points for this tracker
            self.draw_collection_points(pos, x_base, y_base, scale)

            # Draw whip points for this tracker
            self.draw_whip_points(pos)
            # Draw extender points for this tracker
            self.draw_extender_points(pos)

    def draw_whip_points(self, pos):
        """Draw whip points for a tracker"""
        tracker_idx = self.block.tracker_positions.index(pos)
        tracker_id = str(tracker_idx)
        string_count = len(pos.strings)
        
        # Draw regular tracker whip points
        pos_whip = self.get_whip_position(tracker_id, 'positive')
        neg_whip = self.get_whip_position(tracker_id, 'negative')
        
        if pos_whip:
            self.draw_whip_point(pos_whip[0], pos_whip[1], tracker_id, 'positive')
        
        if neg_whip:
            self.draw_whip_point(neg_whip[0], neg_whip[1], tracker_id, 'negative')
        
        # Draw harness-specific whip points
        self.draw_harness_specific_whip_points(tracker_id, string_count)

    def draw_harness_specific_whip_points(self, tracker_id, string_count):
        """Draw harness-specific whip points if they exist"""
        if not (hasattr(self.block, 'wiring_config') and 
                self.block.wiring_config and 
                hasattr(self.block.wiring_config, 'harness_groupings') and
                string_count in self.block.wiring_config.harness_groupings):
            return
        
        for harness_idx, _ in enumerate(self.block.wiring_config.harness_groupings[string_count]):
            harness_pos_whip = self.get_whip_position(tracker_id, 'positive', harness_idx)
            harness_neg_whip = self.get_whip_position(tracker_id, 'negative', harness_idx)
            
            if harness_pos_whip:
                self.draw_whip_point(harness_pos_whip[0], harness_pos_whip[1], tracker_id, 'positive', harness_idx)
            
            if harness_neg_whip:
                self.draw_whip_point(harness_neg_whip[0], harness_neg_whip[1], tracker_id, 'negative', harness_idx)

    def draw_string_homerun_wiring(self):
        """Draw string homerun wiring configuration"""
        
        # Sort strings by Y position to determine route indices
        pos_routes = []  # List to store positive route info
        neg_routes = []  # List to store negative route info
        
        # Collect all source points
        for pos in self.block.tracker_positions:
            tracker_idx = self.block.tracker_positions.index(pos)
            pos_whip = self.get_whip_position(str(tracker_idx), 'positive')
            tracker_idx = self.block.tracker_positions.index(pos)
            neg_whip = self.get_whip_position(str(tracker_idx), 'negative')
            
            for string in pos.strings:
                whip_points_valid = pos_whip and neg_whip  # Check both whip points exist
                if whip_points_valid:
                    # First route: source to whip
                    route1 = self.calculate_cable_route(
                        pos.x + string.positive_source_x,
                        pos.y + string.positive_source_y,
                        pos_whip[0], pos_whip[1],
                        True, len(pos_routes)
                    )
                    current = self.calculate_current_for_segment('string')
                    context_info = f"T{tracker_idx+1}-S{len(pos_routes)+1} String"
                    self.draw_wire_route(route1, self.string_cable_size_var.get(), current, True, "string", context_info)
                    
                if neg_whip:
                    route1 = self.calculate_cable_route(
                        pos.x + string.negative_source_x,
                        pos.y + string.negative_source_y,
                        neg_whip[0], neg_whip[1],
                        False, len(neg_routes)
                    )
                    current = self.calculate_current_for_segment('string')
                    context_info = f"T{tracker_idx+1}-S{len(neg_routes)+1} String"
                    self.draw_wire_route(route1, self.string_cable_size_var.get(), current, False, "string", context_info)
                                            
                pos_routes.append({
                    'source_x': pos_whip[0] if pos_whip else (pos.x + string.positive_source_x),
                    'source_y': pos_whip[1] if pos_whip else (pos.y + string.positive_source_y),
                    'going_north': self.block.device_y < (pos_whip[1] if pos_whip else (pos.y + string.positive_source_y))
                })
                neg_routes.append({
                    'source_x': neg_whip[0] if neg_whip else (pos.x + string.negative_source_x),
                    'source_y': neg_whip[1] if neg_whip else (pos.y + string.negative_source_y),
                    'going_north': self.block.device_y < (neg_whip[1] if neg_whip else (pos.y + string.negative_source_y))
                })
        
        # Sort routes by y position within their north/south groups
        pos_routes_north = sorted([r for r in pos_routes if r['going_north']], 
                                key=lambda r: r['source_y'], reverse=True)
        pos_routes_south = sorted([r for r in pos_routes if not r['going_north']], 
                                key=lambda r: r['source_y'])
        neg_routes_north = sorted([r for r in neg_routes if r['going_north']], 
                                key=lambda r: r['source_y'], reverse=True)
        neg_routes_south = sorted([r for r in neg_routes if not r['going_north']], 
                                key=lambda r: r['source_y'])
        
        # Get device destination points
        pos_dest, neg_dest = self.get_device_destination_points()
        
        # Draw routes from whip points to device
        for i, route_info in enumerate(pos_routes_north + pos_routes_south):
            route = self.calculate_cable_route(
                route_info['source_x'],
                route_info['source_y'],
                pos_dest[0],  # Left side of device
                pos_dest[1],  # Device y-position
                True,  # is_positive
                i  # route_index
            )
            current = self.calculate_current_for_segment('whip', num_strings=1)
            context_info = f"T{i+1} Whip"
            self.draw_wire_route(route, self.whip_cable_size_var.get(), current, True, "whip", context_info)
        
        for i, route_info in enumerate(neg_routes_north + neg_routes_south):
            route = self.calculate_cable_route(
                route_info['source_x'],
                route_info['source_y'],
                neg_dest[0],  # Right side of device
                neg_dest[1],  # Device y-position
                False,  # is_positive
                i  # route_index
            )
            current = self.calculate_current_for_segment('whip', num_strings=1)
            context_info = f"T{i+1} Whip"
            self.draw_wire_route(route, self.whip_cable_size_var.get(), current, False, "whip", context_info)

    def draw_wire_harness_wiring(self):
        """Draw wire harness wiring configuration"""
        # Process each tracker
        for pos_idx, pos in enumerate(self.block.tracker_positions):
            tracker_id = str(pos_idx)
            
            # Get whip points for this tracker
            pos_whip = self.get_whip_position(tracker_id, 'positive')
            neg_whip = self.get_whip_position(tracker_id, 'negative')
            
            # Check if we have custom harness groupings for this tracker
            string_count = len(pos.strings)
            has_custom_groupings = (hasattr(self.block.wiring_config, 'harness_groupings') and 
                                string_count in self.block.wiring_config.harness_groupings and 
                                self.block.wiring_config.harness_groupings[string_count])
            
            # Draw either default or custom harness configuration
            if not has_custom_groupings:
                self.draw_default_harness(pos, pos_whip, neg_whip)
            else:
                # Draw custom harnesses
                self.draw_custom_harnesses(pos, pos_whip, neg_whip, string_count)
                
                # Find unconfigured strings and auto-configure them
                all_configured_indices = set()
                for harness in self.block.wiring_config.harness_groupings[string_count]:
                    all_configured_indices.update(harness.string_indices)
                
                unconfigured_indices = [i for i in range(string_count) if i not in all_configured_indices]
                
                if unconfigured_indices:
                    self.draw_auto_harnesses_for_unconfigured(pos, pos_whip, neg_whip, unconfigured_indices)

    def draw_default_harness(self, pos, pos_whip, neg_whip):
        """Draw default harness configuration - single independent harness"""
        # Use the unified harness drawing logic
        string_indices = list(range(len(pos.strings)))
        routing_mode = self.routing_mode_var.get()
        
        # For default harness, we don't have a harness group, so cable sizes come from block config
        self._current_harness_group = None
        self.draw_harness_for_tracker(pos, 0, string_indices, routing_mode, is_default=True)

    def draw_custom_harnesses(self, pos, pos_whip, neg_whip, string_count):
        """Draw custom harness configurations - each harness is completely independent"""
        tracker_idx = self.block.tracker_positions.index(pos)
        
        # Process each harness group independently
        for harness_idx, harness in enumerate(self.block.wiring_config.harness_groupings[string_count]):
            if not harness.string_indices:
                continue
            
            self.draw_single_custom_harness(pos, harness_idx, harness, tracker_idx)

    def draw_single_custom_harness(self, pos, harness_idx, harness, tracker_idx):
        """Draw a single custom harness configuration"""
        string_indices = harness.string_indices
        
        # Step 1: Calculate this harness's own extender and whip points
        pos_extender_point = self.get_extender_point(tracker_idx, 'positive', harness_idx)
        neg_extender_point = self.get_extender_point(tracker_idx, 'negative', harness_idx)
        
        # Each harness has its own whip point (co-located but separate)
        pos_whip_point = self.get_shared_whip_point(tracker_idx, 'positive')
        neg_whip_point = self.get_shared_whip_point(tracker_idx, 'negative')
        
        # Store current harness group for use in drawing methods
        self._current_harness_group = harness

        # Step 2: Calculate and draw collection points for this harness
        routing_mode = self.routing_mode_var.get()
        pos_nodes, neg_nodes = self.calculate_and_draw_harness_nodes(pos, string_indices, harness_idx, routing_mode)
        
        # Step 3: Draw harness connections from collection points to extender point
        if pos_nodes and pos_extender_point:
            self.draw_harness_to_extender_connection(pos_nodes, pos_extender_point, len(string_indices), True, harness, tracker_idx, harness_idx)
        
        if neg_nodes and neg_extender_point:
            self.draw_harness_to_extender_connection(neg_nodes, neg_extender_point, len(string_indices), False, harness, tracker_idx, harness_idx)
        
        # Step 4: Draw extender cables
        self.draw_harness_extender_cables(pos_extender_point, neg_extender_point, pos_whip_point, neg_whip_point, len(string_indices), tracker_idx, harness_idx)
        
        # Step 5: Draw whip to device connections
        self.draw_harness_whip_to_device(tracker_idx, harness_idx, len(string_indices))

        # Clear current harness group
        self._current_harness_group = None

    def calculate_and_draw_harness_nodes(self, pos, string_indices, harness_idx, routing_mode):
        """Calculate and draw collection nodes for a harness"""
        pos_nodes = []
        neg_nodes = []
        
        for string_idx in string_indices:
            if string_idx < len(pos.strings):
                string = pos.strings[string_idx]
                
                pos_node = self.get_harness_collection_point(pos, string, harness_idx, True, routing_mode)
                neg_node = self.get_harness_collection_point(pos, string, harness_idx, False, routing_mode)
                
                pos_nodes.append(pos_node)
                neg_nodes.append(neg_node)
                
                # Draw string to collection point
                self.draw_string_to_harness_connection(pos, string, pos_node, string_idx, True, routing_mode)
                self.draw_string_to_harness_connection(pos, string, neg_node, string_idx, False, routing_mode)
        
        # Draw collection points
        for nx, ny in pos_nodes:
            self.draw_collection_point(nx, ny, True, 'collection')
        
        for nx, ny in neg_nodes:
            self.draw_collection_point(nx, ny, False, 'collection')
        
        return pos_nodes, neg_nodes

    def draw_harness_extender_cables(self, pos_extender_point, neg_extender_point, pos_whip_point, neg_whip_point, num_strings, tracker_idx, harness_idx):
        """Draw extender cables for a harness"""
        current_pos_whip = self.get_whip_position(str(tracker_idx), 'positive', harness_idx)
        current_neg_whip = self.get_whip_position(str(tracker_idx), 'negative', harness_idx)
        
        # Check if we have extended extender points
        extended_pos = self.extended_extender_points.get((tracker_idx, harness_idx, 'positive'), pos_extender_point)
        extended_neg = self.extended_extender_points.get((tracker_idx, harness_idx, 'negative'), neg_extender_point)

        # Get harness group for cable sizes
        harness = None
        string_count = len(self.block.tracker_positions[tracker_idx].strings)
        if (hasattr(self.block.wiring_config, 'harness_groupings') and 
            string_count in self.block.wiring_config.harness_groupings and
            harness_idx < len(self.block.wiring_config.harness_groupings[string_count])):
            harness = self.block.wiring_config.harness_groupings[string_count][harness_idx]

        if extended_pos and current_pos_whip:
            self._current_harness_group = harness
            self.draw_extender_route(extended_pos, current_pos_whip, num_strings, True, tracker_idx, harness_idx)
            self._current_harness_group = None

        if extended_neg and current_neg_whip:
            self._current_harness_group = harness
            self.draw_extender_route(extended_neg, current_neg_whip, num_strings, False, tracker_idx, harness_idx)
            self._current_harness_group = None

    def draw_harness_whip_to_device(self, tracker_idx, harness_idx, num_strings):
        """Draw whip to device connections for a harness"""
        pos_dest, neg_dest = self.get_device_destination_points()
        
        current_pos_whip = self.get_whip_position(str(tracker_idx), 'positive', harness_idx)
        current_neg_whip = self.get_whip_position(str(tracker_idx), 'negative', harness_idx)

        # Get harness group for cable sizes
        harness = None
        string_count = len(self.block.tracker_positions[tracker_idx].strings)
        if (hasattr(self.block.wiring_config, 'harness_groupings') and 
            string_count in self.block.wiring_config.harness_groupings and
            harness_idx < len(self.block.wiring_config.harness_groupings[string_count])):
            harness = self.block.wiring_config.harness_groupings[string_count][harness_idx]
            # Set as current for cable size lookups
            self._current_harness_group = harness

        if current_pos_whip and pos_dest:
            self._current_harness_group = harness
            self.draw_whip_to_device_connection(current_pos_whip, pos_dest, num_strings, True, harness_idx, tracker_idx)
            self._current_harness_group = None

        if current_neg_whip and neg_dest:
            self._current_harness_group = harness
            self.draw_whip_to_device_connection(current_neg_whip, neg_dest, num_strings, False, harness_idx, tracker_idx)
            self._current_harness_group = None

    def draw_auto_harnesses_for_unconfigured(self, pos, pos_whip, neg_whip, unconfigured_indices):
        """Automatically draw harnesses for unconfigured strings"""
        if not unconfigured_indices:
            return
            
        scale = self.get_canvas_scale()
        # Calculate harness index (after all existing harnesses)
        string_count = len(pos.strings)
        harness_idx = len(self.block.wiring_config.harness_groupings[string_count])
        
        # Calculate node positions for unconfigured strings
        pos_nodes = []
        neg_nodes = []
        routing_mode = self.routing_mode_var.get()

        for string_idx in unconfigured_indices:
            if string_idx < len(pos.strings):
                string = pos.strings[string_idx]
                
                # Use position helpers
                pos_node = self.get_harness_collection_point(pos, string, harness_idx, True, routing_mode)
                neg_node = self.get_harness_collection_point(pos, string, harness_idx, False, routing_mode)
                
                pos_nodes.append(pos_node)
                neg_nodes.append(neg_node)
        
        # Draw string connections to nodes
        tracker_idx = self.block.tracker_positions.index(pos)
        for i, string_idx in enumerate(unconfigured_indices):
            if string_idx < len(pos.strings):
                string = pos.strings[string_idx]
                
                # Positive string cable
                self.draw_string_to_harness_connection(pos, string, pos_nodes[i], string_idx, True, routing_mode)
                
                # Negative string cable
                self.draw_string_to_harness_connection(pos, string, neg_nodes[i], string_idx, False, routing_mode)
        
        # Draw node points for this harness
        for nx, ny in pos_nodes:
            self.draw_collection_point(nx, ny, True, 'collection')
        
        for nx, ny in neg_nodes:
            self.draw_collection_point(nx, ny, False, 'collection')
        
        # Get harness-specific whip points using helper
        tracker_id = str(tracker_idx)
        harness_pos_whip = self.get_harness_whip_point(tracker_id, harness_idx, 'positive', routing_mode)
        harness_neg_whip = self.get_harness_whip_point(tracker_id, harness_idx, 'negative', routing_mode)
        
        # Create a dummy harness object for the drawing method with default sizes
        class DummyHarness:
            def __init__(self):
                self.cable_size = "10 AWG"
                self.string_cable_size = "10 AWG"
                self.extender_cable_size = "8 AWG"
                self.whip_cable_size = "8 AWG"

        dummy_harness = DummyHarness()
        
        # Draw harness connections using the same method as custom harnesses
        # This ensures proper sequential connections instead of individual lines
        if pos_nodes and harness_pos_whip:
            self.draw_harness_to_extender_connection(
                pos_nodes, harness_pos_whip, 
                len(unconfigured_indices), True, 
                dummy_harness, tracker_idx, harness_idx
            )
        
        if neg_nodes and harness_neg_whip:
            self.draw_harness_to_extender_connection(
                neg_nodes, harness_neg_whip, 
                len(unconfigured_indices), False, 
                dummy_harness, tracker_idx, harness_idx
            )
        
        # Draw harness whip points
        if harness_pos_whip:
            wx, wy = self.world_to_canvas(harness_pos_whip[0], harness_pos_whip[1])
            self.canvas.create_oval(
                wx - 3, wy - 3,
                wx + 3, wy + 3,
                fill='pink', outline='deeppink',
                tags=f'auto_harness_pos_whip'
            )
        
        if harness_neg_whip:
            wx, wy = self.world_to_canvas(harness_neg_whip[0], harness_neg_whip[1])
            self.canvas.create_oval(
                wx - 3, wy - 3,
                wx + 3, wy + 3,
                fill='teal', outline='darkcyan',
                tags=f'auto_harness_neg_whip'
            )
        
        # Route from harness whip points to device
        pos_dest, neg_dest = self.get_device_destination_points()
        
        # Draw whip to device connections
        if harness_pos_whip and pos_dest:
            self.draw_whip_to_device_connection(
                harness_pos_whip, pos_dest, 
                len(unconfigured_indices), True, harness_idx, tracker_idx
            )
        
        if harness_neg_whip and neg_dest:
            self.draw_whip_to_device_connection(
                harness_neg_whip, neg_dest, 
                len(unconfigured_indices), False, harness_idx, tracker_idx
            )

    def draw_whip_to_device_connection(self, whip_point, device_point, num_strings, is_positive, harness_idx=0, tracker_idx=0):
        """Draw a connection from a whip point to the device"""
        route = self.calculate_cable_route(
            whip_point[0], whip_point[1],
            device_point[0], device_point[1],
            is_positive, harness_idx
        )

        # Get cable size from harness group if available
        cable_size = "8 AWG"  # Default for harness mode
        if hasattr(self, '_current_harness_group') and self._current_harness_group:
            harness_cable_size = getattr(self._current_harness_group, 'whip_cable_size', '')
            if harness_cable_size:
                cable_size = harness_cable_size
        elif self.wiring_type_var.get() == WiringType.HOMERUN.value:
            # Only use block-level default in String Homerun mode
            cable_size = self.whip_cable_size_var.get()

        current = self.calculate_current_for_segment('whip', num_strings)
        context_info = f"T{tracker_idx+1}-H{harness_idx+1} Whip"
        self.draw_wire_route(route, cable_size, current, is_positive, "whip", context_info)

    def get_selected_string_count(self):
        """Get the currently selected string count"""
        if not self.string_count_var.get() or not hasattr(self, 'string_count_mapping'):
            return None
        
        selected_item = self.string_count_var.get()
        if selected_item not in self.string_count_mapping:
            return None
            
        return self.string_count_mapping[selected_item]

    def ensure_wiring_config_exists(self):
        """Ensure wiring config and harness groupings are initialized"""
        if not hasattr(self.block, 'wiring_config') or not self.block.wiring_config:
            from ..models.block import WiringConfig, WiringType
            self.block.wiring_config = WiringConfig(
                wiring_type=WiringType(self.wiring_type_var.get()),
                positive_collection_points=[],
                negative_collection_points=[],
                strings_per_collection={},
                cable_routes={},
                harness_groupings={}
            )

        if not hasattr(self.block.wiring_config, 'harness_groupings'):
            self.block.wiring_config.harness_groupings = {}

    def remove_custom_harness_config(self, string_count):
        """Remove custom harness configuration"""
        if string_count in self.block.wiring_config.harness_groupings:
            del self.block.wiring_config.harness_groupings[string_count]
        
    def calculate_recommended_fuse_size(self, string_indices):
        """Calculate recommended fuse size based on module Imp"""
        if not self.block or not self.block.tracker_template or not self.block.tracker_template.module_spec:
            return 15  # Default
        
        module_imp = self.block.tracker_template.module_spec.imp
        
        # Find the next standard fuse size above Imp
        for fuse_size in self.FUSE_RATINGS:
            if fuse_size > module_imp:
                return fuse_size
        
        return self.FUSE_RATINGS[-1]  # Return largest if none found
    
    def get_current_routes(self):
        """Get routes based on current routing mode (realistic or conceptual)"""
        # Return routes that were saved during drawing (if available)
        if hasattr(self, 'saved_routes_for_block') and self.saved_routes_for_block:
            return self.saved_routes_for_block.copy()
        
        # If no saved routes, return empty dictionary
        return {}
        
    def draw_current_routes(self):
        """Draw routes based on current wiring type (always uses conceptual logic with positioning variants)"""
        if self.wiring_type_var.get() == WiringType.HOMERUN.value:
            self.draw_string_homerun_wiring()
        else:  # Wire Harness configuration
            self.draw_wire_harness_wiring()

    def add_current_label_to_route(self, canvas_points, current, is_positive, segment_type, context_info=None):
        """Add current label to a route with smart positioning and drag capability"""
        if not self.show_current_labels_var.get():
            return
            
        if len(canvas_points) < 4:  # Need at least 2 points (x,y pairs)
            return
            
        # Convert flat list to point pairs for midpoint calculation
        points = [(canvas_points[i], canvas_points[i+1]) for i in range(0, len(canvas_points), 2)]
        
        # Find midpoint
        if len(points) == 1:
            mid_x, mid_y = points[0]
        else:
            mid_idx = len(points) // 2
            if mid_idx == 0:
                mid_x, mid_y = points[0]
            else:
                mid_x = (points[mid_idx-1][0] + points[mid_idx][0]) / 2
                mid_y = (points[mid_idx-1][1] + points[mid_idx][1]) / 2
        
        wire_midpoint = (mid_x, mid_y)
        
        # Smart positioning to avoid overlaps
        label_pos = self.find_smart_label_position(wire_midpoint, is_positive)
        
        color = 'red' if is_positive else 'blue'
        # Create descriptive label with context
        if context_info:
            label_text = f"{context_info}: {current:.1f}A"
        else:
            label_text = f"{segment_type}: {current:.1f}A"
        
        text_id = self.canvas.create_text(label_pos[0], label_pos[1], text=label_text, 
                            fill=color, font=('Arial', 8), tags='current_label')
        
        # Store label info for dragging
        label_id = f"label_{len(self.current_labels)}"
        self.current_labels[label_id] = {
            'text_id': text_id,
            'wire_midpoint': wire_midpoint,
            'current_pos': label_pos,
            'leader_id': None
        }
        
        # Bind drag events to the text
        self.canvas.tag_bind(text_id, '<Button-1>', lambda e, lid=label_id: self.start_label_drag(e, lid))
        self.canvas.tag_bind(text_id, '<B1-Motion>', lambda e, lid=label_id: self.drag_label(e, lid))
        self.canvas.tag_bind(text_id, '<ButtonRelease-1>', lambda e, lid=label_id: self.end_label_drag(e, lid))

    def get_string_count_for_route(self, route_id):
        """Get number of strings for any route based on route ID"""
        # Extract tracker index from route ID
        parts = route_id.split('_')
        if len(parts) >= 3:
            try:
                tracker_idx = int(parts[2])
                if tracker_idx < len(self.block.tracker_positions):
                    return len(self.block.tracker_positions[tracker_idx].strings)
            except (ValueError, IndexError):
                pass
        return 1

    def draw_wire_route(self, route, wire_gauge, current, is_positive=True, segment_type="string", context_info=None):
        """Draw a complete wire route with current calculation and labeling"""
        return self.draw_route_with_properties(route, wire_gauge, current, is_positive, segment_type, context_info)

    def draw_string_to_harness_connection(self, pos, string, harness_node, string_idx, is_positive, routing_mode):
        """Draw connection from string source to harness collection point"""
        # Calculate source point
        if is_positive:
            source_x = pos.x + string.positive_source_x
            source_y = pos.y + string.positive_source_y
        else:
            source_x = pos.x + string.negative_source_x  
            source_y = pos.y + string.negative_source_y
        
        # Get route using helper
        route = self.get_string_source_to_harness_route(
            (source_x, source_y), harness_node, routing_mode
        )
        
        # Get cable size from harness group if available
        cable_size = "10 AWG"  # Default for harness mode
        if hasattr(self, '_current_harness_group') and self._current_harness_group:
            harness_cable_size = getattr(self._current_harness_group, 'string_cable_size', '')
            if harness_cable_size:
                cable_size = harness_cable_size
        elif self.wiring_type_var.get() == WiringType.HOMERUN.value:
            # Only use block-level default in String Homerun mode
            cable_size = self.string_cable_size_var.get()

        # Draw the route
        current = self.calculate_current_for_segment('string')
        tracker_idx = self.block.tracker_positions.index(pos)
        context_info = f"T{tracker_idx+1}-S{string_idx+1} String"
        return self.draw_wire_route(route, cable_size, 
                                current, is_positive, "string", context_info)
    
    def needs_extender(self, tracker_idx, harness_idx=None):
        """Determine if a tracker/harness combination needs an extender"""
        if self.wiring_type_var.get() != WiringType.HARNESS.value:
            return False
        
        pos = self.block.tracker_positions[tracker_idx]
        
        # Case 1: Multi-harness scenario - harness 2, 3, etc. need extenders
        if harness_idx is not None and harness_idx > 0:
            return True
        
        # Case 2: Stacked tracker scenario
        return self.is_stacked_tracker_needing_extender(tracker_idx)

    def is_stacked_tracker_needing_extender(self, tracker_idx):
        """Check if tracker is stacked and needs extender due to being further from device"""
        pos = self.block.tracker_positions[tracker_idx]
        
        # Find other trackers with similar x-coordinate (within 3m)
        stacked_trackers = []
        for i, other_pos in enumerate(self.block.tracker_positions):
            if i != tracker_idx and abs(pos.x - other_pos.x) <= 3.0:
                stacked_trackers.append((i, other_pos))
        
        if not stacked_trackers:
            return False
        
        # Calculate distance from each tracker to device
        device_y = self.block.device_y
        current_distance = abs(pos.y - device_y)
        
        # Check if this tracker is further from device than any other stacked tracker
        for other_idx, other_pos in stacked_trackers:
            other_distance = abs(other_pos.y - device_y)
            if other_distance < current_distance:
                return True
        
        return False

    def get_extender_point(self, tracker_idx, polarity, harness_idx=None):
        """Calculate extender point position (the collection point closest to device)"""
        pos = self.block.tracker_positions[tracker_idx]
        string_count = len(pos.strings)
        
        if harness_idx is not None:
            # For multi-harness scenario, find the collection point closest to device
            if (hasattr(self.block, 'wiring_config') and 
                self.block.wiring_config and 
                hasattr(self.block.wiring_config, 'harness_groupings') and
                string_count in self.block.wiring_config.harness_groupings and
                harness_idx < len(self.block.wiring_config.harness_groupings[string_count])):
                
                harness = self.block.wiring_config.harness_groupings[string_count][harness_idx]
                string_indices = harness.string_indices
                
                # Check if this is a 1-string harness
                if len(string_indices) == 1:
                    # For 1-string harness, return the collection point directly
                    string_idx = string_indices[0]
                    if string_idx < len(pos.strings):
                        string = pos.strings[string_idx]
                        routing_mode = self.routing_mode_var.get()
                        collection_point = self.get_harness_collection_point(pos, string, harness_idx, polarity == 'positive', routing_mode)
                        return collection_point
                
                if string_indices:
                    # For multi-string harness, continue with existing logic
                    routing_mode = self.routing_mode_var.get()
                    collection_points = []
                    
                    for string_idx in string_indices:
                        if string_idx < len(pos.strings):
                            string = pos.strings[string_idx]
                            collection_point = self.get_harness_collection_point(pos, string, harness_idx, polarity == 'positive', routing_mode)
                            collection_points.append((collection_point, string_idx))
                    
                    if collection_points:
                        # Get just the collection point coordinates
                        coord_points = [cp[0] for cp in collection_points]
                        # Find the end closest to device for extender point
                        return self.get_harness_extender_end(tracker_idx, harness_idx, polarity, coord_points)
        
        # Fallback to default whip position
        tracker_id = str(tracker_idx)
        return self.get_whip_default_position(tracker_id, polarity, harness_idx=0)

    def draw_extender_route(self, extender_point, whip_point, num_strings, is_positive, tracker_idx, harness_idx=None):
        """Draw an extender cable route (vertical only)"""
        if not extender_point or not whip_point:
            return
        
        # Extenders route vertically only - straight line
        route = [extender_point, whip_point]
        
        # Convert to canvas points for drawing
        canvas_points = self.world_route_to_canvas_points(route)
        
        # Get cable size from harness group if available
        cable_size = "8 AWG"  # Default
        if hasattr(self, '_current_harness_group') and self._current_harness_group:
            harness_cable_size = getattr(self._current_harness_group, 'extender_cable_size', '')
            if harness_cable_size:
                cable_size = harness_cable_size

        # Draw the wire segment
        current = self.calculate_current_for_segment('extender', num_strings)
        harness_num = harness_idx + 1 if harness_idx is not None else 1
        context_info = f"T{tracker_idx+1}-H{harness_num} Extender"
        line_id = self.draw_wire_route(route, cable_size, current, is_positive, "extender", context_info)

        return line_id
    
    def add_custom_harness_routes(self, cable_routes, pos, tracker_idx):
        """Add routes for custom harness configurations including extenders"""
        string_count = len(pos.strings)
        
        if not (hasattr(self.block, 'wiring_config') and 
                self.block.wiring_config and 
                hasattr(self.block.wiring_config, 'harness_groupings') and
                string_count in self.block.wiring_config.harness_groupings):
            return
        
        # Process each harness group
        for harness_idx, harness in enumerate(self.block.wiring_config.harness_groupings[string_count]):
            string_indices = harness.string_indices
            if not string_indices:
                continue
            
            # Check if this harness needs extenders
            harness_needs_extender = self.needs_extender(tracker_idx, harness_idx)
            
            # Get whip and extender points
            tracker_id = str(tracker_idx)
            routing_mode = self.routing_mode_var.get()
            harness_pos_whip = self.get_harness_whip_point(tracker_id, harness_idx, 'positive', routing_mode)
            harness_neg_whip = self.get_harness_whip_point(tracker_id, harness_idx, 'negative', routing_mode)
            
            pos_extender_point = None
            neg_extender_point = None
            if harness_needs_extender:
                pos_extender_point = self.get_extender_point(tracker_idx, 'positive', harness_idx)
                neg_extender_point = self.get_extender_point(tracker_idx, 'negative', harness_idx)
            
            # Calculate node positions for this harness group
            pos_nodes = []
            neg_nodes = []
            
            for string_idx in string_indices:
                if string_idx < len(pos.strings):
                    string = pos.strings[string_idx]
                    
                    pos_node = self.get_harness_collection_point(pos, string, harness_idx, True, routing_mode)
                    neg_node = self.get_harness_collection_point(pos, string, harness_idx, False, routing_mode)
                    
                    pos_nodes.append(pos_node)
                    neg_nodes.append(neg_node)
                    
                    # Add string to node routes
                    cable_routes[f"pos_node_{tracker_idx}_{harness_idx}_{string_idx}"] = [
                        (pos.x + string.positive_source_x, pos.y + string.positive_source_y),
                        pos_node
                    ]
                    
                    cable_routes[f"neg_node_{tracker_idx}_{harness_idx}_{string_idx}"] = [
                        (pos.x + string.negative_source_x, pos.y + string.negative_source_y),
                        neg_node
                    ]
            
            # Add harness routes using new routing direction
            if pos_nodes:
                target_point = pos_extender_point if harness_needs_extender else harness_pos_whip
                if target_point:
                    # Sort nodes to route from far end toward extender point
                    extender_end = self.get_harness_extender_end(tracker_idx, harness_idx, 'positive', pos_nodes)
                    if extender_end:
                        extender_y = extender_end[1]
                        extender_is_north = extender_y == min(p[1] for p in pos_nodes)
                        if extender_is_north:
                            sorted_nodes = sorted(pos_nodes, key=lambda p: p[1], reverse=True)
                        else:
                            sorted_nodes = sorted(pos_nodes, key=lambda p: p[1], reverse=False)
                        cable_routes[f"pos_harness_{tracker_idx}_{harness_idx}"] = sorted_nodes + [target_point]

            if neg_nodes:
                target_point = neg_extender_point if harness_needs_extender else harness_neg_whip
                if target_point:
                    # Sort nodes to route from far end toward extender point
                    extender_end = self.get_harness_extender_end(tracker_idx, harness_idx, 'negative', neg_nodes)
                    if extender_end:
                        extender_y = extender_end[1]
                        extender_is_north = extender_y == min(p[1] for p in neg_nodes)
                        if extender_is_north:
                            sorted_nodes = sorted(neg_nodes, key=lambda p: p[1], reverse=True)
                        else:
                            sorted_nodes = sorted(neg_nodes, key=lambda p: p[1], reverse=False)
                        cable_routes[f"neg_harness_{tracker_idx}_{harness_idx}"] = sorted_nodes + [target_point]
            
            # Add extender routes if needed
            if harness_needs_extender:
                if pos_extender_point and harness_pos_whip:
                    cable_routes[f"pos_extender_{tracker_idx}_{harness_idx}"] = [pos_extender_point, harness_pos_whip]
                
                if neg_extender_point and harness_neg_whip:
                    cable_routes[f"neg_extender_{tracker_idx}_{harness_idx}"] = [neg_extender_point, harness_neg_whip]
            
            # Add whip to device routes
            pos_dest, neg_dest = self.get_device_destination_points()
            if harness_pos_whip and pos_dest:
                cable_routes[f"pos_main_{tracker_idx}_{harness_idx}"] = [harness_pos_whip, pos_dest]
            
            if harness_neg_whip and neg_dest:
                cable_routes[f"neg_main_{tracker_idx}_{harness_idx}"] = [harness_neg_whip, neg_dest]

    def get_wire_color(self, segment_type, is_positive=True):
        """Get wire color based on segment type and polarity"""
        if is_positive:
            # Warm colors - light to dark based on importance
            colors = {
                'string': '#FFB6C1',    # Light Pink (least important)
                'whip': '#FFA500',      # Orange  
                'harness': '#FF0000',   # Red
                'extender': '#8B0000'   # Dark Red (most important)
            }
        else:
            # Cool colors - light to dark based on importance  
            colors = {
                'string': '#B0FFFF',    # Light Cyan (least important)
                'whip': '#40E0D0',      # Turquoise
                'harness': '#0000FF',   # Blue
                'extender': '#800080'   # Purple (most important)
            }
        
        return colors.get(segment_type, '#FF0000' if is_positive else '#0000FF')
    
    def draw_extender_points(self, pos):
        """Draw extender points for a tracker"""
        tracker_idx = self.block.tracker_positions.index(pos)
        string_count = len(pos.strings)
        
        # Check if this tracker needs extenders
        if not self.needs_extender(tracker_idx):
            return
        
        # Draw extender points for multi-harness scenario
        if self.has_custom_harness_groupings(string_count):
            self.draw_multi_harness_extender_points(tracker_idx, string_count)
        # Draw extender points for stacked tracker scenario
        elif self.is_stacked_tracker_needing_extender(tracker_idx):
            self.draw_stacked_tracker_extender_points(tracker_idx)

    def has_custom_harness_groupings(self, string_count):
        """Check if there are custom harness groupings for the given string count"""
        return (hasattr(self.block, 'wiring_config') and 
                self.block.wiring_config and 
                hasattr(self.block.wiring_config, 'harness_groupings') and
                string_count in self.block.wiring_config.harness_groupings)

    def draw_multi_harness_extender_points(self, tracker_idx, string_count):
        """Draw extender points for multi-harness configuration"""
        for harness_idx, _ in enumerate(self.block.wiring_config.harness_groupings[string_count]):
            if harness_idx > 0:  # Only harness 2, 3, etc. get extenders
                pos_extender = self.get_extender_point(tracker_idx, 'positive', harness_idx)
                neg_extender = self.get_extender_point(tracker_idx, 'negative', harness_idx)
                
                if pos_extender:
                    self.draw_collection_point(pos_extender[0], pos_extender[1], True, 'extender')
                
                if neg_extender:
                    self.draw_collection_point(neg_extender[0], neg_extender[1], False, 'extender')

    def draw_stacked_tracker_extender_points(self, tracker_idx):
        """Draw extender points for stacked tracker configuration"""
        pos_extender = self.get_extender_point(tracker_idx, 'positive')
        neg_extender = self.get_extender_point(tracker_idx, 'negative')
        
        if pos_extender:
            self.draw_collection_point(pos_extender[0], pos_extender[1], True, 'extender')
        
        if neg_extender:
            self.draw_collection_point(neg_extender[0], neg_extender[1], False, 'extender')

    def draw_legend(self):
        """Draw the color legend showing wire types and colors"""
        if not hasattr(self, 'legend_canvas'):
            return
            
        # Clear existing legend
        self.legend_canvas.delete("all")
        
        # Legend title
        self.legend_canvas.create_text(140, 10, text="Wire Colors by Type", 
                                    font=('Arial', 10, 'bold'), anchor='center')
        
        # Wire types in order of importance
        wire_types = ['extender', 'harness', 'whip', 'string']
        type_labels = ['Extender', 'Harness', 'Whip', 'String']
        
        y_start = 25
        line_length = 25
        
        # Draw positive column
        self.legend_canvas.create_text(70, y_start, text="Positive (+)", 
                                    font=('Arial', 9, 'bold'), anchor='center')
        
        for i, (wire_type, label) in enumerate(zip(wire_types, type_labels)):
            y_pos = y_start + 15 + (i * 18)
            color = self.get_wire_color(wire_type, True)
            
            # Draw colored line
            self.legend_canvas.create_line(20, y_pos, 20 + line_length, y_pos, 
                                        fill=color, width=3)
            
            # Draw label
            self.legend_canvas.create_text(50, y_pos, text=label, 
                                        font=('Arial', 8), anchor='w')
        
        # Draw negative column  
        self.legend_canvas.create_text(210, y_start, text="Negative (-)", 
                                    font=('Arial', 9, 'bold'), anchor='center')
        
        for i, (wire_type, label) in enumerate(zip(wire_types, type_labels)):
            y_pos = y_start + 15 + (i * 18)
            color = self.get_wire_color(wire_type, False)
            
            # Draw colored line
            self.legend_canvas.create_line(160, y_pos, 160 + line_length, y_pos, 
                                        fill=color, width=3)
            
            # Draw label
            self.legend_canvas.create_text(190, y_pos, text=label, 
                                        font=('Arial', 8), anchor='w')
        
        
    def get_shared_whip_point(self, tracker_idx, polarity):
        """Get the shared whip point that all harnesses connect to via extenders"""
        # This is the standard whip point location that all harnesses ultimately connect to
        tracker_id = str(tracker_idx)
        return self.get_whip_default_position(tracker_id, polarity, harness_idx=0)
    
    def draw_harness_to_extender_connection(self, collection_points, extender_point, num_strings, is_positive, harness, tracker_idx, harness_idx):
        """Draw harness cables from collection points to extender point"""
        if not collection_points or not extender_point:
            return
        
        # Check if this is a 1-string harness
        if num_strings == 1:
            # For 1-string harness, no harness trunk is drawn
            # The extender point is already at the collection point
            # Just update the extended extender points to match
            self.extended_extender_points[(tracker_idx, harness_idx, 'positive' if is_positive else 'negative')] = collection_points[0]
            return
        
        # Sort collection points to route from far end toward extender point
        extender_point = self.get_harness_extender_end(tracker_idx, harness_idx, is_positive, collection_points)
        if not extender_point:
            return
        
        # Get the opposite polarity's extender point to check if we need a long trunk
        opposite_polarity = 'negative' if is_positive else 'positive'
        opposite_extender_point = self.get_extender_point(tracker_idx, opposite_polarity, harness_idx)
        
        # Determine if we need a long trunk end based on device location
        needs_long_trunk = False
        final_extender_point = extender_point
        
        if opposite_extender_point and abs(extender_point[1] - opposite_extender_point[1]) > 0.1:
            # The extender points are at different y-coordinates
            device_y = self.block.device_y
            
            # Calculate the midpoint between the two harness extender points
            harness_midpoint_y = (extender_point[1] + opposite_extender_point[1]) / 2
            
            # Determine which harness needs the long trunk
            if device_y < min(extender_point[1], opposite_extender_point[1]):
                # Device is north of both harnesses
                needs_long_trunk = not is_positive  # Negative gets long trunk
            elif device_y > max(extender_point[1], opposite_extender_point[1]):
                # Device is south of both harnesses  
                needs_long_trunk = is_positive  # Positive gets long trunk
            else:
                # Device is between harnesses - check which side of center
                if device_y < harness_midpoint_y:
                    # Device is on north side of center
                    needs_long_trunk = not is_positive  # Negative gets long trunk
                else:
                    # Device is on south side of center (or exactly at center)
                    needs_long_trunk = is_positive  # Positive gets long trunk
                
            if needs_long_trunk:
                # Create new extender point at the same y-coordinate as opposite harness
                final_extender_point = (extender_point[0], opposite_extender_point[1])
        
        # Determine if extender is at north or south end
        extender_y = extender_point[1]
        extender_is_north = extender_y == min(p[1] for p in collection_points)

        if extender_is_north:
            # Extender at north, route from south to north  
            sorted_points = sorted(collection_points, key=lambda p: p[1], reverse=True)
        else:
            # Extender at south, route from north to south
            sorted_points = sorted(collection_points, key=lambda p: p[1], reverse=False)
        
        # Connect collection points in sequence, with last one being the original extender point
        for i in range(len(sorted_points)):
            start_point = sorted_points[i]
            
            if i < len(sorted_points) - 1:
                # Connect to next collection point
                end_point = sorted_points[i + 1]
            else:
                # Last point - connect to extender point
                end_point = extender_point
                # Verify they're not the same (they should be for the last collection point)
                if abs(end_point[0] - start_point[0]) < 0.1 and abs(end_point[1] - start_point[1]) < 0.1:
                    # If we need a long trunk, draw from this point to the final extender point
                    if needs_long_trunk and abs(extender_point[1] - final_extender_point[1]) > 0.1:
                        route = self.calculate_cable_route(
                            extender_point[0], extender_point[1],
                            final_extender_point[0], final_extender_point[1],
                            is_positive, len(sorted_points)
                        )
                        current = self.calculate_current_for_segment('harness', num_strings)
                        cable_size = harness.cable_size if hasattr(harness, 'cable_size') else self.harness_cable_size_var.get()
                        context_info = f"T{tracker_idx+1}-H{harness_idx+1} Harness"
                        self.draw_wire_route(route, cable_size, current, is_positive, "harness", context_info)
                    continue  # Don't draw zero-length cable
            
            route = self.calculate_cable_route(
                start_point[0], start_point[1],
                end_point[0], end_point[1],
                is_positive, i
            )
            
            # Calculate accumulated current (more strings as we progress toward device)
            accumulated_strings = len(sorted_points) - i
            current = self.calculate_current_for_segment('harness', num_strings)
            cable_size = harness.cable_size if hasattr(harness, 'cable_size') else self.harness_cable_size_var.get()
            context_info = f"T{tracker_idx+1}-H{harness_idx+1} Harness"
            self.draw_wire_route(route, cable_size, current, is_positive, "harness", context_info)
        
        # If we added a long trunk, draw that final segment
        if needs_long_trunk and abs(extender_point[1] - final_extender_point[1]) > 0.1:
            # Update the extender point for this harness to the extended position
            self.extended_extender_points[(tracker_idx, harness_idx, 'positive' if is_positive else 'negative')] = final_extender_point

    def draw_simple_harness_connection(self, collection_points, extender_point, num_strings, is_positive):
        """Draw simple harness connection for default single harness"""
        if not collection_points or not extender_point:
            return
        
        # Check if this is a 1-string configuration
        if num_strings == 1:
            # For 1-string, no harness trunk is drawn
            # The extender will run directly from the collection point
            return
        
        # Find which end should be the extender point (closest to device)
        extender_point = self.get_harness_extender_end(0, 0, is_positive, collection_points)  # Use tracker 0, harness 0 for default
        if not extender_point:
            return

        # Determine if extender is at north or south end
        extender_y = extender_point[1]
        extender_is_north = extender_y == min(p[1] for p in collection_points)

        # Sort to route from far end toward extender point
        if extender_is_north:
            sorted_points = sorted(collection_points, key=lambda p: p[1], reverse=True)  # South to north
        else:
            sorted_points = sorted(collection_points, key=lambda p: p[1], reverse=False)  # North to south
        
        # Connect collection points in sequence, ending at extender point
        for i in range(len(sorted_points)):
            start_point = sorted_points[i]
            
            if i < len(sorted_points) - 1:
                # Connect to next collection point
                end_point = sorted_points[i + 1]
            else:
                # Last collection point connects to extender point
                end_point = extender_point
            
            route = self.calculate_cable_route(
                start_point[0], start_point[1],
                end_point[0], end_point[1],
                is_positive, 0
            )
            
            # Calculate accumulated current (more strings as we progress)
            accumulated_strings = i + 1
            current = self.calculate_current_for_segment('harness', accumulated_strings)
            tracker_idx = self.block.tracker_positions.index(pos) if hasattr(self, 'current_tracker_pos') else 0
            context_info = f"T{tracker_idx+1}-H1 Harness"
            self.draw_wire_route(route, self.harness_cable_size_var.get(), current, is_positive, "harness", context_info)

    def find_smart_label_position(self, wire_midpoint, is_positive):
        """Find a position for the label that avoids overlaps"""
        base_offset = -15 if is_positive else 15
        
        # Try positions at increasing distances
        for distance in [base_offset, base_offset * 1.5, base_offset * 2]:
            candidate_pos = (wire_midpoint[0], wire_midpoint[1] + distance)
            
            # Check if this position overlaps with existing labels
            overlap = False
            for label_info in self.current_labels.values():
                existing_pos = label_info['current_pos']
                if abs(candidate_pos[0] - existing_pos[0]) < 30 and abs(candidate_pos[1] - existing_pos[1]) < 15:
                    overlap = True
                    break
            
            if not overlap:
                return candidate_pos
        
        # If all positions overlap, use the farthest one
        return (wire_midpoint[0], wire_midpoint[1] + base_offset * 2)
    
    def handle_label_interaction(self, event, label_id, action):
        """Handle all label interactions (start, drag, end)"""
        if action == "start":
            self.dragging_label = label_id
            self.label_drag_start = (event.x, event.y)
        
        elif action == "drag" and self.dragging_label == label_id:
            self.update_label_position(event, label_id)
        
        elif action == "end":
            self.dragging_label = None
            self.label_drag_start = None

    def update_label_position(self, event, label_id):
        """Update label position and leader line during drag"""
        label_info = self.current_labels[label_id]
        text_id = label_info['text_id']
        
        # Update label position
        self.canvas.coords(text_id, event.x, event.y)
        label_info['current_pos'] = (event.x, event.y)
        
        # Update or create leader line
        wire_mid = label_info['wire_midpoint']
        distance = ((event.x - wire_mid[0])**2 + (event.y - wire_mid[1])**2)**0.5
        
        if distance > 20:  # Show leader line
            if label_info['leader_id'] is None:
                leader_id = self.canvas.create_line(
                    wire_mid[0], wire_mid[1], event.x, event.y,
                    fill='gray', width=1, dash=(2, 2), tags='label_leader'
                )
                label_info['leader_id'] = leader_id
            else:
                self.canvas.coords(label_info['leader_id'], 
                                wire_mid[0], wire_mid[1], event.x, event.y)
        else:
            # Remove leader line if too close
            if label_info['leader_id'] is not None:
                self.canvas.delete(label_info['leader_id'])
                label_info['leader_id'] = None

        # Update the binding in add_current_label_to_route method:
        # REPLACE the three bind statements with:
        self.canvas.tag_bind(text_id, '<Button-1>', lambda e, lid=label_id: self.handle_label_interaction(e, lid, "start"))
        self.canvas.tag_bind(text_id, '<B1-Motion>', lambda e, lid=label_id: self.handle_label_interaction(e, lid, "drag"))
        self.canvas.tag_bind(text_id, '<ButtonRelease-1>', lambda e, lid=label_id: self.handle_label_interaction(e, lid, "end"))

    def toggle_current_labels(self):
        """Toggle current labels and reset positions"""
        # Clear existing labels and leaders
        self.canvas.delete('current_label')
        self.canvas.delete('label_leader')
        self.current_labels = {}
        
        # Redraw with new state
        self.draw_wiring_layout()

    def get_harness_extender_end(self, tracker_idx, harness_idx, polarity, collection_points):
        """Determine which end of the harness should be the extender point (closest to device)"""
        if not collection_points:
            return None
        
        device_y = self.block.device_y
        
        # Find northernmost and southernmost collection points
        north_point = min(collection_points, key=lambda p: p[1])  # Smallest Y
        south_point = max(collection_points, key=lambda p: p[1])  # Largest Y
        
        # Calculate distances to device
        north_distance = abs(north_point[1] - device_y)
        south_distance = abs(south_point[1] - device_y)
        
        # Return the end closest to device (for extender point)
        if north_distance < south_distance:
            return north_point
        elif south_distance < north_distance:
            return south_point
        else:
            # Tie - default to north end
            return north_point
        
    def draw_route_with_properties(self, route, wire_gauge, current, is_positive, segment_type, context_info=None):
        """Draw a route with proper wire properties and current labeling"""
        if len(route) < 2:
            return None
            
        # Convert to canvas points
        canvas_points = self.world_route_to_canvas_points(route)
        
        # Save route for Block Configurator if we have context info
        if context_info and hasattr(self, 'saved_routes_for_block'):
            polarity = 'pos' if is_positive else 'neg'
            route_id = f"{polarity}_{segment_type}_{context_info.lower().replace(' ', '_')}"
            self.saved_routes_for_block[route_id] = route.copy()
        
        # Draw the wire segment
        line_id = self.draw_wire_segment(canvas_points, wire_gauge, current, is_positive, segment_type, context_info)
        
        # Add current label if enabled
        if self.show_current_labels_var.get():
            self.add_current_label_to_route(canvas_points, current, is_positive, segment_type, context_info)
        
        return line_id
    
    def draw_collection_point(self, x, y, is_positive, point_type='collection'):
        """Draw a single collection point at the specified coordinates"""
        canvas_x, canvas_y = self.world_to_canvas(x, y)
        
        if point_type == 'collection':
            size = 3
            fill_color = 'red' if is_positive else 'blue'
            outline_color = 'darkred' if is_positive else 'darkblue'
        elif point_type == 'whip':
            size = 3
            fill_color = 'pink' if is_positive else 'teal'
            outline_color = 'deeppink' if is_positive else 'darkcyan'
        else:  # extender
            size = 4
            fill_color = 'yellow' if is_positive else 'lightblue'
            outline_color = 'orange' if is_positive else 'darkblue'
        
        return self.canvas.create_oval(
            canvas_x - size, canvas_y - size,
            canvas_x + size, canvas_y + size,
            fill=fill_color, outline=outline_color,
            tags=f'{point_type}_point'
        )
    
    def draw_whip_point(self, world_x, world_y, tracker_id, polarity, harness_idx=None, selected_whips=None):
        """Draw a single whip point with appropriate styling"""
        canvas_x, canvas_y = self.world_to_canvas(world_x, world_y)
        
        # Determine if selected
        if harness_idx is not None:
            is_selected = (tracker_id, harness_idx, polarity) in (selected_whips or self.selected_whips)
            tag_suffix = f'_{harness_idx}_{polarity}'
        else:
            is_selected = (tracker_id, polarity) in (selected_whips or self.selected_whips)
            tag_suffix = f'_{polarity}'
        
        # Set colors and size based on polarity and selection
        if polarity == 'positive':
            fill_color = 'orange' if is_selected else 'pink'
            outline_color = 'red' if is_selected else 'deeppink'
        else:
            fill_color = 'cyan' if is_selected else 'teal'
            outline_color = 'blue' if is_selected else 'darkcyan'
        
        size = 5 if is_selected else 3
        tag = f'harness_whip_point{tag_suffix}' if harness_idx is not None else 'whip_point'
        
        return self.canvas.create_oval(
            canvas_x - size, canvas_y - size,
            canvas_x + size, canvas_y + size,
            fill=fill_color, outline=outline_color,
            tags=tag
        )
    
    def update_whip_point_position(self, whip_info, dx, dy):
        """Update a whip point position by the given delta"""
        if len(whip_info) == 3:
            tracker_id, harness_idx, polarity = whip_info
            self.update_harness_whip_position(tracker_id, harness_idx, polarity, dx, dy)
        else:
            tracker_id, polarity = whip_info
            self.update_regular_whip_position(tracker_id, polarity, dx, dy)

    def update_harness_whip_position(self, tracker_id, harness_idx, polarity, dx, dy):
        """Update harness-specific whip point position"""
        # Ensure custom_harness_whip_points exists
        if not hasattr(self.block, 'wiring_config') or not self.block.wiring_config:
            return
        if not hasattr(self.block.wiring_config, 'custom_harness_whip_points'):
            self.block.wiring_config.custom_harness_whip_points = {}
        
        # Ensure tracker_id and harness_idx entries exist
        if tracker_id not in self.block.wiring_config.custom_harness_whip_points:
            self.block.wiring_config.custom_harness_whip_points[tracker_id] = {}
        if harness_idx not in self.block.wiring_config.custom_harness_whip_points[tracker_id]:
            self.block.wiring_config.custom_harness_whip_points[tracker_id][harness_idx] = {}
        
        # Get current position or default position
        if polarity not in self.block.wiring_config.custom_harness_whip_points[tracker_id][harness_idx]:
            current_pos = self.get_whip_default_position(tracker_id, polarity, harness_idx)
            if current_pos:
                self.block.wiring_config.custom_harness_whip_points[tracker_id][harness_idx][polarity] = current_pos
        
        # Update position
        if polarity in self.block.wiring_config.custom_harness_whip_points[tracker_id][harness_idx]:
            old_x, old_y = self.block.wiring_config.custom_harness_whip_points[tracker_id][harness_idx][polarity]
            self.block.wiring_config.custom_harness_whip_points[tracker_id][harness_idx][polarity] = (old_x + dx, old_y + dy)

    def update_regular_whip_position(self, tracker_id, polarity, dx, dy):
        """Update regular whip point position"""
        if not hasattr(self.block, 'wiring_config') or not self.block.wiring_config:
            self.block.wiring_config = WiringConfig(
                wiring_type=WiringType(self.wiring_type_var.get()),
                positive_collection_points=[],
                negative_collection_points=[],
                strings_per_collection={},
                cable_routes={},
                custom_whip_points={}
            )
            
        if 'custom_whip_points' not in self.block.wiring_config.__dict__:
            self.block.wiring_config.custom_whip_points = {}
            
        if tracker_id not in self.block.wiring_config.custom_whip_points:
            self.block.wiring_config.custom_whip_points[tracker_id] = {}
            
        # If this is the first time moving this whip, initialize with current position
        if polarity not in self.block.wiring_config.custom_whip_points[tracker_id]:
            current_pos = self.get_whip_default_position(tracker_id, polarity)
            if current_pos:
                self.block.wiring_config.custom_whip_points[tracker_id][polarity] = current_pos
                
        # Update position
        if polarity in self.block.wiring_config.custom_whip_points[tracker_id]:
            old_x, old_y = self.block.wiring_config.custom_whip_points[tracker_id][polarity]
            self.block.wiring_config.custom_whip_points[tracker_id][polarity] = (old_x + dx, old_y + dy)

    def draw_harness_for_tracker(self, pos, harness_idx, string_indices, routing_mode, is_default=False):
        """Draw a single harness (default or custom) for a tracker"""
        tracker_idx = self.block.tracker_positions.index(pos)
        
        # Get the harness configuration
        harness = None
        if not is_default:
            string_count = len(pos.strings)
            if (hasattr(self.block.wiring_config, 'harness_groupings') and 
                string_count in self.block.wiring_config.harness_groupings and
                harness_idx < len(self.block.wiring_config.harness_groupings[string_count])):
                harness = self.block.wiring_config.harness_groupings[string_count][harness_idx]
        
        # For default harnesses, create a harness with default values
        if is_default and harness is None:
            harness = HarnessGroup(
                string_indices=string_indices,
                cable_size="10 AWG",
                string_cable_size="10 AWG",
                extender_cable_size="8 AWG",
                whip_cable_size="8 AWG"
            )
        
        # Calculate collection points
        pos_nodes, neg_nodes = self.calculate_harness_collection_points(pos, string_indices, harness_idx, routing_mode)

        # Set current harness group for cable size lookups
        self._current_harness_group = harness
        
        # Draw string to collection connections
        self.draw_string_to_collection_connections(pos, string_indices, pos_nodes, neg_nodes, routing_mode)
        
        # Draw collection points
        self.draw_harness_collection_points(pos_nodes, neg_nodes)
        
        # Get target points (extender or whip)
        pos_target, neg_target = self.get_harness_target_points(tracker_idx, harness_idx, is_default)
        
        # Draw harness connections
        self.draw_harness_connections(pos_nodes, neg_nodes, pos_target, neg_target, len(string_indices), tracker_idx, harness_idx, is_default)
        
        # Draw extender and whip connections
        self.draw_harness_final_connections(tracker_idx, harness_idx, len(string_indices), is_default)

        # Clear current harness group
        self._current_harness_group = None

    def draw_harness_connections(self, pos_nodes, neg_nodes, pos_target, neg_target, num_strings, tracker_idx, harness_idx, is_default):
        """Draw harness connections from collection points to target points"""
        if is_default:
            # For default harness, use the simple harness connection
            if pos_nodes and pos_target:
                self.draw_simple_harness_connection(pos_nodes, pos_target, num_strings, True)
            
            if neg_nodes and neg_target:
                self.draw_simple_harness_connection(neg_nodes, neg_target, num_strings, False)
        else:
            # For custom harnesses, use the harness-to-extender connection
            # Get the harness object for cable size
            string_count = len(self.block.tracker_positions[tracker_idx].strings)
            harness = None
            if (hasattr(self.block.wiring_config, 'harness_groupings') and 
                string_count in self.block.wiring_config.harness_groupings and
                harness_idx < len(self.block.wiring_config.harness_groupings[string_count])):
                harness = self.block.wiring_config.harness_groupings[string_count][harness_idx]
            
            if pos_nodes and pos_target and harness:
                self.draw_harness_to_extender_connection(pos_nodes, pos_target, num_strings, True, harness, tracker_idx, harness_idx)
            
            if neg_nodes and neg_target and harness:
                self.draw_harness_to_extender_connection(neg_nodes, neg_target, num_strings, False, harness, tracker_idx, harness_idx)

    def draw_harness_final_connections(self, tracker_idx, harness_idx, num_strings, is_default):
        """Draw final connections (extender and whip to device)"""
        pos_dest, neg_dest = self.get_device_destination_points()
        
        if is_default:
            # For default harness, draw whip to device directly
            pos_whip_point = self.get_shared_whip_point(tracker_idx, 'positive')
            neg_whip_point = self.get_shared_whip_point(tracker_idx, 'negative')
            
            if pos_whip_point and pos_dest:
                route = self.calculate_cable_route(
                    pos_whip_point[0], pos_whip_point[1],
                    pos_dest[0], pos_dest[1],
                    True, 0
                )
                current = self.calculate_current_for_segment('whip', num_strings)
                context_info = f"T{tracker_idx+1}-H1 Whip"
                self.draw_wire_route(route, self.whip_cable_size_var.get(), current, True, "whip", context_info)
            
            if neg_whip_point and neg_dest:
                route = self.calculate_cable_route(
                    neg_whip_point[0], neg_whip_point[1],
                    neg_dest[0], neg_dest[1],
                    False, 0
                )
                current = self.calculate_current_for_segment('whip', num_strings)
                context_info = f"T{tracker_idx+1}-H1 Whip"
                self.draw_wire_route(route, self.whip_cable_size_var.get(), current, False, "whip", context_info)
        else:
            # For custom harnesses, this is handled by the draw_harness_whip_to_device method
            # which was already called in draw_harness_for_tracker
            pass

    def calculate_harness_collection_points(self, pos, string_indices, harness_idx, routing_mode):
        """Calculate collection points for a harness"""
        pos_nodes = []
        neg_nodes = []
        
        for string_idx in string_indices:
            if string_idx < len(pos.strings):
                string = pos.strings[string_idx]
                
                pos_node = self.get_harness_collection_point(pos, string, harness_idx, True, routing_mode)
                neg_node = self.get_harness_collection_point(pos, string, harness_idx, False, routing_mode)
                
                pos_nodes.append(pos_node)
                neg_nodes.append(neg_node)
        
        return pos_nodes, neg_nodes

    def draw_string_to_collection_connections(self, pos, string_indices, pos_nodes, neg_nodes, routing_mode):
        """Draw connections from strings to collection points"""
        for i, string_idx in enumerate(string_indices):
            if string_idx < len(pos.strings):
                string = pos.strings[string_idx]
                
                if i < len(pos_nodes):
                    self.draw_string_to_harness_connection(pos, string, pos_nodes[i], string_idx, True, routing_mode)
                if i < len(neg_nodes):
                    self.draw_string_to_harness_connection(pos, string, neg_nodes[i], string_idx, False, routing_mode)

    def draw_harness_collection_points(self, pos_nodes, neg_nodes):
        """Draw collection point markers"""
        for nx, ny in pos_nodes:
            self.draw_collection_point(nx, ny, True, 'collection')
        
        for nx, ny in neg_nodes:
            self.draw_collection_point(nx, ny, False, 'collection')

    def get_harness_target_points(self, tracker_idx, harness_idx, is_default):
        """Get target points for harness connections (extender or whip)"""
        needs_extender = self.needs_extender(tracker_idx, harness_idx) if not is_default else False
        
        if needs_extender:
            pos_target = self.get_extender_point(tracker_idx, 'positive', harness_idx)
            neg_target = self.get_extender_point(tracker_idx, 'negative', harness_idx)
        else:
            pos_target = self.get_shared_whip_point(tracker_idx, 'positive')
            neg_target = self.get_shared_whip_point(tracker_idx, 'negative')
        
        return pos_target, neg_target
    
    def notify_wiring_changed(self):
        """Notify that wiring configuration has changed"""
        # Save state
        if self.parent_notify_blocks_changed:
            self.parent_notify_blocks_changed()
        
        # If the main app has a device configurator, refresh it
        if hasattr(self.parent, 'main_app') and hasattr(self.parent.main_app, 'device_configurator'):
            self.parent.main_app.device_configurator.refresh_display()