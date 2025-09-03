"""
Single Line Diagram Editor Window
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Dict, Optional, Tuple
import math

from ..models.block import BlockConfig
from ..models.project import Project
from ..models.sld import (
    SLDDiagram, SLDElement, SLDConnection, SLDAnnotation,
    SLDElementType, ConnectionPortType, ConnectionPort
)
from ..utils.sld_symbols import ANSISymbols


class SLDEditor(tk.Toplevel):
    """Single Line Diagram Editor Window"""
    
    def __init__(self, parent, blocks: Dict[str, BlockConfig], project: Optional[Project] = None):
        super().__init__(parent)
        
        self.parent = parent
        self.blocks = blocks
        self.project = project
        
        # Window setup
        self.title(f"Single Line Diagram Editor - {project.metadata.name if project else 'New Diagram'}")
        self.geometry("1400x900")
        self.minsize(1000, 600)
        
        # Center the window
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (1400 // 2)
        y = (self.winfo_screenheight() // 2) - (900 // 2)
        self.geometry(f'+{x}+{y}')
        
        # SLD properties
        self.sld_diagram = None
        self.sld_elements = {}  # element_id -> canvas item IDs
        self.selected_element = None
        self.selected_items = []
        
        # Canvas properties
        self.canvas = None
        self.grid_visible = True
        self.grid_size = 10
        self.zoom_level = 1.0
        
        # Interaction properties
        self.dragging = False
        self.drag_data = {"x": 0, "y": 0, "item": None, "element_id": None}
        self.connecting = False
        self.connection_start = None
        self.temp_connection_line = None
        
        # Set up the UI
        self.setup_ui()
        
        # Generate initial SLD from blocks
        self.generate_sld_from_blocks()
        
        # Handle window closing
        self.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def setup_ui(self):
        """Set up the user interface"""
        # Main container
        main_frame = ttk.Frame(self, padding="5")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)
        
        # Create menu bar
        self.create_menu()
        
        # Create toolbar
        self.create_toolbar(main_frame)
        
        # Create main content area with canvas and side panel
        content_frame = ttk.Frame(main_frame)
        content_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(5, 0))
        content_frame.grid_columnconfigure(0, weight=1)
        content_frame.grid_rowconfigure(0, weight=1)
        
        # Create canvas area
        self.create_canvas_area(content_frame)
        
        # Create status bar
        self.create_status_bar(main_frame)
    
    def create_menu(self):
        """Create the menu bar"""
        menubar = tk.Menu(self)
        self.configure(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Export PNG...", command=self.export_png, accelerator="Ctrl+P")
        file_menu.add_command(label="Export PDF...", command=self.export_pdf, accelerator="Ctrl+D")
        file_menu.add_separator()
        file_menu.add_command(label="Close", command=self.on_close, accelerator="Esc")
        
        # Edit menu
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="Select All", command=self.select_all, accelerator="Ctrl+A")
        edit_menu.add_command(label="Delete Selected", command=self.delete_selected, accelerator="Del")
        edit_menu.add_separator()
        edit_menu.add_command(label="Auto Layout", command=self.auto_layout)
        edit_menu.add_command(label="Clear All", command=self.clear_all)
        
        # View menu
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Zoom In", command=lambda: self.zoom(1.25), accelerator="Ctrl++")
        view_menu.add_command(label="Zoom Out", command=lambda: self.zoom(0.8), accelerator="Ctrl+-")
        view_menu.add_command(label="Fit to Window", command=self.fit_to_window, accelerator="Ctrl+0")
        view_menu.add_separator()
        view_menu.add_checkbutton(label="Show Grid", variable=tk.BooleanVar(value=True), 
                                  command=self.toggle_grid)
        
        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Refresh from Blocks", command=self.generate_sld_from_blocks)
        tools_menu.add_command(label="Validate Connections", command=self.validate_connections)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Keyboard Shortcuts", command=self.show_shortcuts)
        help_menu.add_command(label="About SLD Editor", command=self.show_about)
    
    def create_toolbar(self, parent):
        """Create the toolbar"""
        toolbar_frame = ttk.Frame(parent)
        toolbar_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # Layout section
        ttk.Label(toolbar_frame, text="Layout:").grid(row=0, column=0, padx=(5, 2))
        
        ttk.Button(
            toolbar_frame,
            text="Auto Layout",
            command=self.auto_layout
        ).grid(row=0, column=1, padx=2)
        
        ttk.Button(
            toolbar_frame,
            text="Align Left",
            command=lambda: self.align_selected('left')
        ).grid(row=0, column=2, padx=2)
        
        ttk.Button(
            toolbar_frame,
            text="Align Top",
            command=lambda: self.align_selected('top')
        ).grid(row=0, column=3, padx=2)
        
        ttk.Separator(toolbar_frame, orient='vertical').grid(row=0, column=4, padx=10, sticky='ns')
        
        # Zoom section
        ttk.Label(toolbar_frame, text="Zoom:").grid(row=0, column=5, padx=(5, 2))
        
        ttk.Button(
            toolbar_frame,
            text="-",
            width=3,
            command=lambda: self.zoom(0.8)
        ).grid(row=0, column=6, padx=1)
        
        self.zoom_label = ttk.Label(toolbar_frame, text="100%", width=6)
        self.zoom_label.grid(row=0, column=7, padx=2)
        
        ttk.Button(
            toolbar_frame,
            text="+",
            width=3,
            command=lambda: self.zoom(1.25)
        ).grid(row=0, column=8, padx=1)
        
        ttk.Button(
            toolbar_frame,
            text="Fit",
            command=self.fit_to_window
        ).grid(row=0, column=9, padx=2)
        
        ttk.Separator(toolbar_frame, orient='vertical').grid(row=0, column=10, padx=10, sticky='ns')
        
        # Export section
        ttk.Label(toolbar_frame, text="Export:").grid(row=0, column=11, padx=(5, 2))
        
        ttk.Button(
            toolbar_frame,
            text="Export PNG",
            command=self.export_png
        ).grid(row=0, column=12, padx=2)
        
        ttk.Button(
            toolbar_frame,
            text="Export PDF",
            command=self.export_pdf
        ).grid(row=0, column=13, padx=2)
        
        ttk.Separator(toolbar_frame, orient='vertical').grid(row=0, column=14, padx=10, sticky='ns')
        
        # Connection mode toggle
        self.connection_mode_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            toolbar_frame,
            text="Connection Mode",
            variable=self.connection_mode_var,
            command=self.toggle_connection_mode
        ).grid(row=0, column=15, padx=5)
        
        # Grid toggle
        self.grid_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            toolbar_frame,
            text="Show Grid",
            variable=self.grid_var,
            command=self.toggle_grid
        ).grid(row=0, column=16, padx=5)
    
    def create_canvas_area(self, parent):
        """Create the main canvas area"""
        # Canvas frame with border
        canvas_frame = ttk.Frame(parent, relief=tk.SUNKEN, borderwidth=2)
        canvas_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Create canvas
        self.canvas = tk.Canvas(
            canvas_frame,
            bg='white',
            width=1200,
            height=800,
            scrollregion=(0, 0, 2400, 1600)
        )
        self.canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbars
        h_scroll = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        h_scroll.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        v_scroll = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        v_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        self.canvas.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)
        
        # Configure grid weights
        canvas_frame.grid_columnconfigure(0, weight=1)
        canvas_frame.grid_rowconfigure(0, weight=1)
        
        # Bind canvas events
        self.canvas.bind("<Button-1>", self.on_canvas_click)
        self.canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        self.canvas.bind("<Button-3>", self.on_canvas_right_click)
        self.canvas.bind("<Double-1>", self.on_canvas_double_click)
        
        # Mouse wheel for zoom
        self.canvas.bind("<Control-MouseWheel>", self.on_canvas_zoom)  # Windows
        self.canvas.bind("<Control-Button-4>", self.on_canvas_zoom)  # Linux up
        self.canvas.bind("<Control-Button-5>", self.on_canvas_zoom)  # Linux down
        
        # Regular scroll without Ctrl
        self.canvas.bind("<MouseWheel>", self.on_canvas_scroll)
        self.canvas.bind("<Button-4>", self.on_canvas_scroll)
        self.canvas.bind("<Button-5>", self.on_canvas_scroll)
        
        # Middle button for panning
        self.canvas.bind("<Button-2>", self.on_pan_start)
        self.canvas.bind("<B2-Motion>", self.on_pan_motion)
        
        # Keyboard shortcuts
        self.bind("<Delete>", lambda e: self.delete_selected())
        self.bind("<Control-a>", lambda e: self.select_all())
        self.bind("<Control-plus>", lambda e: self.zoom(1.25))
        self.bind("<Control-minus>", lambda e: self.zoom(0.8))
        self.bind("<Control-0>", lambda e: self.fit_to_window())
        self.bind("<Escape>", lambda e: self.on_close())
        
        # Draw initial grid
        self.draw_grid()
    
    def create_status_bar(self, parent):
        """Create the status bar"""
        status_frame = ttk.Frame(parent)
        status_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # Status label
        self.status_label = ttk.Label(status_frame, text="Ready", relief=tk.SUNKEN)
        self.status_label.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 2))
        
        # Coordinates label
        self.coord_label = ttk.Label(status_frame, text="X: 0, Y: 0", relief=tk.SUNKEN, width=20)
        self.coord_label.grid(row=0, column=1, sticky=tk.E, padx=2)
        
        # Element count label
        self.element_count_label = ttk.Label(status_frame, text="Elements: 0", relief=tk.SUNKEN, width=15)
        self.element_count_label.grid(row=0, column=2, sticky=tk.E, padx=2)
        
        # Connection count label
        self.connection_count_label = ttk.Label(status_frame, text="Connections: 0", relief=tk.SUNKEN, width=15)
        self.connection_count_label.grid(row=0, column=3, sticky=tk.E)
        
        status_frame.grid_columnconfigure(0, weight=1)
        
        # Bind mouse motion for coordinate display
        self.canvas.bind("<Motion>", self.on_canvas_motion)
    
    def draw_grid(self):
        """Draw grid lines on canvas"""
        self.canvas.delete("grid")
        
        if not self.grid_visible:
            return
        
        width = 2400
        height = 1600
        
        # Apply zoom to grid
        grid_spacing = self.grid_size * self.zoom_level
        
        # Draw vertical lines
        for x in range(0, int(width * self.zoom_level), int(grid_spacing)):
            line_width = 1
            color = "#E0E0E0"
            
            # Every 5th line is thicker
            if (x / grid_spacing) % 5 == 0:
                line_width = 2
                color = "#D0D0D0"
            
            # Every 10th line is even thicker
            if (x / grid_spacing) % 10 == 0:
                line_width = 2
                color = "#C0C0C0"
            
            self.canvas.create_line(
                x, 0, x, height * self.zoom_level,
                fill=color,
                width=line_width,
                tags="grid"
            )
        
        # Draw horizontal lines
        for y in range(0, int(height * self.zoom_level), int(grid_spacing)):
            line_width = 1
            color = "#E0E0E0"
            
            # Every 5th line is thicker
            if (y / grid_spacing) % 5 == 0:
                line_width = 2
                color = "#D0D0D0"
            
            # Every 10th line is even thicker
            if (y / grid_spacing) % 10 == 0:
                line_width = 2
                color = "#C0C0C0"
            
            self.canvas.create_line(
                0, y, width * self.zoom_level, y,
                fill=color,
                width=line_width,
                tags="grid"
            )
        
        # Keep grid behind other elements
        self.canvas.tag_lower("grid")
    
    def generate_sld_from_blocks(self):
        """Generate SLD diagram from blocks"""
        if not self.blocks:
            self.status_label.configure(text="No blocks available")
            return
        
        # Clear canvas
        self.canvas.delete("all")
        self.draw_grid()
        
        # Create SLD diagram
        self.sld_diagram = SLDDiagram(
            project_id=self.project.metadata.name if self.project else "unnamed",
            diagram_name=f"SLD - {self.project.metadata.name if self.project else 'New'}"
        )
        
        # Parse blocks into SLD elements
        self.parse_blocks_to_sld()
        
        # Auto-layout the elements
        self.sld_diagram.auto_layout()
        
        # Render the diagram
        self.render_sld_diagram()
        
        self.update_status_counts()
        self.status_label.configure(text="SLD generated from blocks")
    
    def parse_blocks_to_sld(self):
        """Parse block configurations into SLD elements at string level"""
        if not self.sld_diagram:
            return
        
        # Track inverters and combiners to avoid duplicates
        inverters_added = {}  # inverter_id -> element
        combiners_added = {}  # combiner_id -> element
        
        # Sort blocks for consistent ordering
        from ..utils.calculations import natural_sort_key
        sorted_block_ids = sorted(self.blocks.keys(), key=natural_sort_key)
        
        for block_id in sorted_block_ids:
            block = self.blocks[block_id]
            
            # Determine if using combiner boxes or direct to inverter
            using_combiners = (hasattr(block, 'device_type') and 
                              str(block.device_type) == "DeviceType.COMBINER_BOX")
            
            # Get harness groupings if available
            harness_groups = []
            if block.wiring_config and hasattr(block.wiring_config, 'harness_groupings'):
                # Organize by harness groups
                for string_count, groups in block.wiring_config.harness_groupings.items():
                    for group in groups:
                        harness_groups.append({
                            'string_indices': group.string_indices,
                            'string_count': len(group.string_indices),
                            'cable_size': group.cable_size,
                            'fuse_rating': group.fuse_rating_amps
                        })
            
            if not harness_groups:
                # No harness groups defined - create default groups per tracker
                if hasattr(block, 'tracker_positions'):
                    for tracker_idx, tracker_pos in enumerate(block.tracker_positions):
                        if hasattr(tracker_pos, 'strings'):
                            string_indices = [s.index for s in tracker_pos.strings]
                            harness_groups.append({
                                'string_indices': string_indices,
                                'string_count': len(string_indices),
                                'cable_size': '10 AWG',
                                'fuse_rating': 15
                            })
            
            # Create string group elements (one per harness or tracker)
            string_elements = []
            for group_idx, group in enumerate(harness_groups):
                # Calculate power for this string group
                strings_in_group = group['string_count']
                
                # Get module info if available
                module_power_w = 550  # Default
                modules_per_string = 30  # Default
                
                if hasattr(block, 'tracker_template') and block.tracker_template:
                    modules_per_string = block.tracker_template.modules_per_string
                    if block.tracker_template.module_spec:
                        module_power_w = block.tracker_template.module_spec.wattage
                
                group_power_kw = (strings_in_group * modules_per_string * module_power_w) / 1000
                
                # Create string group element
                string_element = SLDElement(
                    element_id=f"STRINGS_{block_id}_G{group_idx+1}",
                    element_type=SLDElementType.PV_BLOCK,
                    x=50 + (group_idx * 100),  # Will be repositioned by auto-layout
                    y=50 + (group_idx * 50),
                    width=120,
                    height=80,
                    label=f"{strings_in_group} String{'s' if strings_in_group > 1 else ''}\n{group_power_kw:.1f} kW",
                    source_block_id=block_id,
                    power_kw=group_power_kw
                )
                
                # Add properties
                string_element.properties['string_count'] = strings_in_group
                string_element.properties['harness_size'] = group['cable_size']
                string_element.properties['fuse_rating'] = group['fuse_rating']
                
                # Add connection ports
                string_element.ports.append(ConnectionPort(
                    port_id="dc_positive",
                    port_type=ConnectionPortType.DC_POSITIVE,
                    side="right",
                    offset=0.3,
                    max_current=strings_in_group * 10  # Assuming ~10A per string
                ))
                string_element.ports.append(ConnectionPort(
                    port_id="dc_negative",
                    port_type=ConnectionPortType.DC_NEGATIVE,
                    side="right",
                    offset=0.7,
                    max_current=strings_in_group * 10
                ))
                
                self.sld_diagram.add_element(string_element)
                string_elements.append(string_element)
            
            # Create combiner boxes if needed
            if using_combiners:
                # Determine number of combiners needed
                total_inputs_needed = len(string_elements)
                inputs_per_combiner = block.num_inputs if hasattr(block, 'num_inputs') else 12
                num_combiners = (total_inputs_needed + inputs_per_combiner - 1) // inputs_per_combiner
                
                combiner_elements = []
                for cb_idx in range(num_combiners):
                    cb_id = f"CB_{block_id}_{cb_idx+1}"
                    
                    # Check if already added
                    if cb_id not in combiners_added:
                        cb_element = SLDElement(
                            element_id=cb_id,
                            element_type=SLDElementType.COMBINER_BOX,
                            x=400 + (cb_idx * 100),
                            y=200 + (cb_idx * 100),
                            width=100,
                            height=100,
                            label=f"{block_id}\n{inputs_per_combiner} inputs",
                            source_block_id=block_id
                        )
                        
                        # Add input ports based on actual inputs
                        inputs_for_this_cb = min(inputs_per_combiner, total_inputs_needed - cb_idx * inputs_per_combiner)
                        for i in range(inputs_for_this_cb):
                            cb_element.ports.append(ConnectionPort(
                                port_id=f"input_{i+1}_pos",
                                port_type=ConnectionPortType.DC_POSITIVE,
                                side="left",
                                offset=(i + 1) / (inputs_for_this_cb + 1)
                            ))
                        
                        # Add output ports
                        cb_element.ports.append(ConnectionPort(
                            port_id="output_positive",
                            port_type=ConnectionPortType.DC_POSITIVE,
                            side="right",
                            offset=0.35
                        ))
                        cb_element.ports.append(ConnectionPort(
                            port_id="output_negative",
                            port_type=ConnectionPortType.DC_NEGATIVE,
                            side="right",
                            offset=0.65
                        ))
                        
                        self.sld_diagram.add_element(cb_element)
                        combiners_added[cb_id] = cb_element
                        combiner_elements.append(cb_element)
                
                # Connect strings to combiners
                for idx, string_elem in enumerate(string_elements):
                    # Determine which combiner this string connects to
                    cb_idx = idx // inputs_per_combiner
                    if cb_idx < len(combiner_elements):
                        cb_elem = combiner_elements[cb_idx]
                        input_idx = (idx % inputs_per_combiner) + 1
                        
                        # Create connection
                        connection = SLDConnection(
                            connection_id=f"CONN_{string_elem.element_id}_to_{cb_elem.element_id}",
                            from_element=string_elem.element_id,
                            from_port="dc_positive",
                            to_element=cb_elem.element_id,
                            to_port=f"input_{input_idx}_pos",
                            cable_type="DC",
                            cable_size=string_elem.properties.get('harness_size', '10 AWG'),
                            color="#DC143C"  # Red for positive
                        )
                        self.sld_diagram.add_connection(connection)
            
            # Add inverter if assigned
            if block.inverter:
                inv_key = f"{block.inverter.manufacturer}_{block.inverter.model}"
                
                if inv_key not in inverters_added:
                    inverter = block.inverter
                    
                    # Create inverter element
                    inv_element = SLDElement(
                        element_id=f"INV_{inverter.model.replace(' ', '_')}",
                        element_type=SLDElementType.INVERTER,
                        x=800,
                        y=300,
                        width=180,
                        height=120,
                        label=f"{inverter.manufacturer}\n{inverter.model}\n{inverter.max_ac_power_w/1000:.0f} kW",
                        source_inverter_id=inv_key,
                        power_kw=inverter.max_ac_power_w / 1000
                    )
                    
                    # Add MPPT input ports based on inverter config
                    if hasattr(inverter, 'mppt_channels'):
                        num_mppts = len(inverter.mppt_channels)
                        for mppt_idx in range(num_mppts):
                            # Each MPPT might have multiple string inputs
                            inv_element.ports.append(ConnectionPort(
                                port_id=f"mppt_{mppt_idx+1}_pos",
                                port_type=ConnectionPortType.DC_POSITIVE,
                                side="left",
                                offset=(mppt_idx + 1) / (num_mppts + 1)
                            ))
                    else:
                        # Default DC inputs
                        inv_element.ports.append(ConnectionPort(
                            port_id="dc_positive_in",
                            port_type=ConnectionPortType.DC_POSITIVE,
                            side="left",
                            offset=0.3
                        ))
                        inv_element.ports.append(ConnectionPort(
                            port_id="dc_negative_in",
                            port_type=ConnectionPortType.DC_NEGATIVE,
                            side="left",
                            offset=0.7
                        ))
                    
                    # AC outputs
                    inv_element.ports.append(ConnectionPort(
                        port_id="ac_l1",
                        port_type=ConnectionPortType.AC_L1,
                        side="right",
                        offset=0.25
                    ))
                    inv_element.ports.append(ConnectionPort(
                        port_id="ac_l2",
                        port_type=ConnectionPortType.AC_L2,
                        side="right",
                        offset=0.5
                    ))
                    inv_element.ports.append(ConnectionPort(
                        port_id="ac_l3",
                        port_type=ConnectionPortType.AC_L3,
                        side="right",
                        offset=0.75
                    ))
                    
                    self.sld_diagram.add_element(inv_element)
                    inverters_added[inv_key] = inv_element
                
                # Connect combiners (or strings directly) to inverter
                if using_combiners and combiner_elements:
                    # Connect each combiner to inverter
                    for idx, cb_elem in enumerate(combiner_elements):
                        inv_elem = inverters_added[inv_key]
                        
                        # Determine which MPPT input to use
                        if hasattr(block.inverter, 'mppt_channels'):
                            mppt_idx = idx % len(block.inverter.mppt_channels)
                            port_name = f"mppt_{mppt_idx+1}_pos"
                        else:
                            port_name = "dc_positive_in"
                        
                        connection = SLDConnection(
                            connection_id=f"CONN_{cb_elem.element_id}_to_{inv_elem.element_id}",
                            from_element=cb_elem.element_id,
                            from_port="output_positive",
                            to_element=inv_elem.element_id,
                            to_port=port_name,
                            cable_type="DC",
                            cable_size="4/0 AWG",  # Larger cable for combiner output
                            color="#8B0000"  # Dark red for higher current
                        )
                        self.sld_diagram.add_connection(connection)
                else:
                    # Direct string to inverter connections
                    inv_elem = inverters_added[inv_key]
                    for idx, string_elem in enumerate(string_elements):
                        # Determine which MPPT input
                        if hasattr(block.inverter, 'mppt_channels'):
                            mppt_idx = idx % len(block.inverter.mppt_channels)
                            port_name = f"mppt_{mppt_idx+1}_pos"
                        else:
                            port_name = "dc_positive_in"
                        
                        connection = SLDConnection(
                            connection_id=f"CONN_{string_elem.element_id}_to_{inv_elem.element_id}",
                            from_element=string_elem.element_id,
                            from_port="dc_positive",
                            to_element=inv_elem.element_id,
                            to_port=port_name,
                            cable_type="DC",
                            cable_size=string_elem.properties.get('harness_size', '10 AWG'),
                            color="#DC143C"
                        )
                        self.sld_diagram.add_connection(connection)
    
    def create_block_to_inverter_connections(self, pv_element, inv_element, block):
        """Create connections between PV block and inverter"""
        # Positive connection
        pos_connection = SLDConnection(
            connection_id=f"CONN_{pv_element.element_id}_to_{inv_element.element_id}_pos",
            from_element=pv_element.element_id,
            from_port="dc_positive",
            to_element=inv_element.element_id,
            to_port="dc_positive_in",
            cable_type="DC",
            cable_size=block.wiring_config.string_cable_size if block.wiring_config else "10 AWG",
            voltage=pv_element.voltage_dc,
            current=pv_element.current_dc
        )
        self.sld_diagram.add_connection(pos_connection)
        
        # Negative connection
        neg_connection = SLDConnection(
            connection_id=f"CONN_{pv_element.element_id}_to_{inv_element.element_id}_neg",
            from_element=pv_element.element_id,
            from_port="dc_negative",
            to_element=inv_element.element_id,
            to_port="dc_negative_in",
            cable_type="DC",
            cable_size=block.wiring_config.string_cable_size if block.wiring_config else "10 AWG",
            voltage=pv_element.voltage_dc,
            current=pv_element.current_dc
        )
        self.sld_diagram.add_connection(neg_connection)
    
    def create_pv_to_combiner_connection(self, pv_element, cb_element):
        """Create connection from PV to combiner box"""
        connection = SLDConnection(
            connection_id=f"CONN_{pv_element.element_id}_to_{cb_element.element_id}",
            from_element=pv_element.element_id,
            from_port="dc_positive",
            to_element=cb_element.element_id,
            to_port="input_1",
            cable_type="DC",
            cable_size="10 AWG"
        )
        self.sld_diagram.add_connection(connection)
    
    def render_sld_diagram(self):
        """Render the SLD diagram on canvas"""
        if not self.sld_diagram:
            return
        
        # Clear existing elements
        self.canvas.delete("element")
        self.canvas.delete("connection")
        self.canvas.delete("label")
        self.sld_elements.clear()
        
        # Draw connections first (so they appear behind elements)
        for connection in self.sld_diagram.connections:
            self.draw_connection(connection)
        
        # Draw elements
        for element in self.sld_diagram.elements:
            self.draw_element(element)
        
        # Draw annotations
        for annotation in self.sld_diagram.annotations:
            self.draw_annotation(annotation)
        
        # Update scroll region to encompass all drawn items
        self.update_scroll_region()

    def update_scroll_region(self):
        """Update canvas scroll region to encompass all items"""
        # Get bounding box of all items
        bbox = self.canvas.bbox("all")
        if bbox:
            # Add some padding around the content
            padding = 50
            x1, y1, x2, y2 = bbox
            x1 -= padding
            y1 -= padding
            x2 += padding
            y2 += padding
            
            # Ensure minimum size
            min_width = 2400
            min_height = 1600
            if (x2 - x1) < min_width:
                x2 = x1 + min_width
            if (y2 - y1) < min_height:
                y2 = y1 + min_height
            
            # Update scroll region
            self.canvas.configure(scrollregion=(x1, y1, x2, y2))
    
    def draw_element(self, element: SLDElement):
        """Draw an SLD element on canvas"""
        # Check if this is a string element
        if element.element_id.startswith('STRINGS_'):
            # Draw technical-style string symbol
            result = ANSISymbols.draw_technical_string(
                self.canvas,
                x=element.x,
                y=element.y,
                width=100,
                height=40,
                element_id=element.element_id,
                fill='#FFFFFF',
                outline='#000000',
                outline_width=1
            )
            
            # Store reference
            self.sld_elements[element.element_id] = result
            
        elif element.element_type == SLDElementType.COMBINER_BOX:
            # Draw technical-style combiner
            num_inputs = element.properties.get('num_inputs', 12)
            
            result = ANSISymbols.draw_technical_combiner(
                self.canvas,
                x=element.x,
                y=element.y,
                width=150,
                height=max(120, num_inputs * 12),  # Scale height with inputs
                num_inputs=num_inputs,
                element_id=element.element_id,
                label=element.label,
                fill='#FFFFFF',
                outline='#000000',
                outline_width=2
            )
            
            # Add label text
            label_lines = element.label.split('\n')
            label_y = element.y - 20
            for line in label_lines:
                label_id = self.canvas.create_text(
                    element.x + 75, label_y,
                    text=line,
                    fill='black',
                    font=('Arial', 9, 'bold'),
                    anchor='n',
                    tags=[element.element_id, 'label', f'{element.element_id}_label']
                )
                result['items']['text'].append(label_id)
                result['items']['all'].append(label_id)
                label_y += 15
            
            # Store reference
            self.sld_elements[element.element_id] = result
            
        else:
            # Use existing symbol drawing for other types
            symbol_type_map = {
                SLDElementType.PV_BLOCK: 'pv_array',
                SLDElementType.INVERTER: 'inverter',
                SLDElementType.COMBINER_BOX: 'combiner'
            }
            
            symbol_type = symbol_type_map.get(element.element_type, 'pv_array')
            
            # Draw using ANSI symbols
            result = ANSISymbols.draw_symbol(
                self.canvas,
                symbol_type=symbol_type,
                x=element.x,
                y=element.y,
                width=element.width,
                height=element.height,
                label=element.label,
                element_id=element.element_id,
                fill=element.color,
                outline=element.stroke_color,
                outline_width=element.stroke_width,
                show_ports=False
            )
            
            # Store reference
            self.sld_elements[element.element_id] = result
    
    def draw_connection(self, connection: SLDConnection):
        """Draw a connection between elements with improved visibility"""
        from_element = self.sld_diagram.get_element(connection.from_element)
        to_element = self.sld_diagram.get_element(connection.to_element)
        
        if not from_element or not to_element:
            return
        
        # Get port positions
        from_pos = from_element.get_port_position(connection.from_port)
        to_pos = to_element.get_port_position(connection.to_port)
        
        if not from_pos or not to_pos:
            return
        
        # Calculate orthogonal path with better routing
        path_points = self.calculate_better_orthogonal_path(
            from_pos, to_pos, 
            from_element, to_element,
            connection.from_port, connection.to_port
        )
        
        # Flatten path points for canvas
        points = []
        for point in path_points:
            points.extend(point)
        
        # Simple red line for all connections
        color = '#FF0000'  # Red
        
        # Draw line
        if len(points) >= 4:
            line_id = self.canvas.create_line(
                *points,
                fill=color,
                width=2,
                tags=('connection', connection.connection_id),
                smooth=False
            )
            connection.canvas_items.append(line_id)
            
            # Add cable size label
            mid_idx = len(path_points) // 2
            mid_point = path_points[mid_idx]
            
            label_id = self.canvas.create_text(
                mid_point[0], mid_point[1] - 5,
                text=connection.cable_size,
                fill='black',
                font=('Arial', 8),
                tags=('connection_label', f'{connection.connection_id}_label'),
                anchor='s'
            )
            connection.canvas_items.append(label_id)
    
    def calculate_better_orthogonal_path(self, from_pos, to_pos, from_element=None, to_element=None, from_port=None, to_port=None):
        """Calculate a better orthogonal path with only horizontal and vertical segments"""
        path = []
        
        from_x, from_y = from_pos
        to_x, to_y = to_pos
        
        # Determine port sides for better routing
        from_side = 'right'  # default
        to_side = 'left'    # default
        
        if from_element and from_port:
            port = from_element.get_port(from_port)
            if port:
                from_side = port.side
        
        if to_element and to_port:
            port = to_element.get_port(to_port)
            if port:
                to_side = port.side
        
        # Start at source
        path.append(from_pos)
        
        # Standard routing: go right from source, then route to destination
        h_offset = 30  # Horizontal offset from elements
        
        if from_side == 'right' and to_side == 'left':
            # Most common case: left-to-right flow
            if to_x > from_x + h_offset * 2:
                # Enough space for direct routing
                mid_x = (from_x + to_x) / 2
                
                # Go right from source
                path.append((mid_x, from_y))
                # Go vertically to align with destination
                path.append((mid_x, to_y))
                # Go to destination
                path.append(to_pos)
            else:
                # Need to route around
                path.append((from_x + h_offset, from_y))
                path.append((from_x + h_offset, to_y))
                path.append(to_pos)
        else:
            # Generic orthogonal routing
            # Move horizontally first
            if abs(to_x - from_x) > 5:
                path.append((to_x, from_y))
            
            # Then move vertically
            path.append(to_pos)
        
        return path
    
    def draw_annotation(self, annotation: SLDAnnotation):
        """Draw a text annotation"""
        if not annotation.visible:
            return
        
        text_id = self.canvas.create_text(
            annotation.x,
            annotation.y,
            text=annotation.text,
            fill=annotation.color,
            font=(annotation.font_family, int(annotation.font_size)),
            anchor=annotation.anchor,
            angle=annotation.rotation,
            tags=('annotation', annotation.annotation_id)
        )
        annotation.canvas_item_id = text_id

    def create_symbol_palette(self, parent):
        """Create a palette of available symbols"""
        palette_frame = ttk.LabelFrame(parent, text="Symbols", padding="5")
        palette_frame.grid(row=0, column=1, sticky=(tk.N, tk.S), padx=(5, 0))
        
        # Create small canvas for each symbol type
        symbols = ANSISymbols.get_available_symbols()
        
        for i, symbol_type in enumerate(symbols):
            # Create label
            ttk.Label(palette_frame, text=symbol_type.replace('_', ' ').title()).grid(
                row=i*2, column=0, pady=(5, 0)
            )
            
            # Create mini canvas for preview
            mini_canvas = tk.Canvas(
                palette_frame,
                width=80,
                height=60,
                bg='white',
                relief=tk.RAISED,
                borderwidth=1
            )
            mini_canvas.grid(row=i*2+1, column=0, padx=5, pady=2)
            
            # Draw mini symbol
            ANSISymbols.draw_symbol(
                mini_canvas,
                symbol_type=symbol_type,
                x=5,
                y=5,
                width=70,
                height=50,
                label=""
            )
            
            # Make draggable (implement in future cards)
            mini_canvas.bind("<Button-1>", lambda e, st=symbol_type: self.start_symbol_drag(st))
    
    def update_status_counts(self):
        """Update the status bar counts"""
        # Count elements
        elements = self.canvas.find_withtag("element")
        self.element_count_label.configure(text=f"Elements: {len(elements)}")
        
        # Count connections
        connections = self.canvas.find_withtag("connection")
        self.connection_count_label.configure(text=f"Connections: {len(connections)}")
    
    # Event handlers
    def on_canvas_click(self, event):
        """Handle canvas click"""
        if self.connection_mode_var.get():
            # Connection mode - will be implemented in Card #8
            pass
        else:
            # Selection/drag mode
            canvas_x = self.canvas.canvasx(event.x)
            canvas_y = self.canvas.canvasy(event.y)
            
            # Find clicked item
            item = self.canvas.find_closest(canvas_x, canvas_y)
            if item:
                tags = self.canvas.gettags(item)
                
                # Find the element ID from tags
                element_id = None
                for tag in tags:
                    if tag.startswith(('PV_', 'INV_', 'CB_', 'STRINGS_')):
                        element_id = tag
                        break
                
                if element_id and element_id in self.sld_elements:
                    # Start dragging this entire element
                    self.dragging = True
                    self.drag_data["x"] = canvas_x
                    self.drag_data["y"] = canvas_y
                    self.drag_data["element_id"] = element_id
                    
                    # Highlight ONLY shape items, not text!
                    element_data = self.sld_elements[element_id]
                    if 'items' in element_data and 'shapes' in element_data['items']:
                        # Only highlight shapes (rectangles, lines, etc.)
                        for item_id in element_data['items']['shapes']:
                            # Only set width for non-text items
                            if self.canvas.type(item_id) != 'text':
                                self.canvas.itemconfig(item_id, width=3)
                    
                    self.selected_element = element_id
    
    def on_canvas_drag(self, event):
        """Handle canvas drag"""
        if self.dragging and self.drag_data["element_id"]:
            canvas_x = self.canvas.canvasx(event.x)
            canvas_y = self.canvas.canvasy(event.y)
            
            # Calculate movement
            dx = canvas_x - self.drag_data["x"]
            dy = canvas_y - self.drag_data["y"]
            
            element_id = self.drag_data["element_id"]
            
            # Move all items with the element_id tag
            if element_id in self.sld_elements:
                items_to_move = self.canvas.find_withtag(element_id)
                for item_id in items_to_move:
                    self.canvas.move(item_id, dx, dy)
            
            # Update the SLD element's position in real-time
            sld_element = self.sld_diagram.get_element(element_id)
            if sld_element:
                sld_element.x += dx / self.zoom_level
                sld_element.y += dy / self.zoom_level
                
                # Redraw all connections for this element
                self.redraw_element_connections(element_id)
            
            # Update drag position
            self.drag_data["x"] = canvas_x
            self.drag_data["y"] = canvas_y

    def redraw_element_connections(self, element_id):
        """Redraw all connections for a specific element"""
        if not self.sld_diagram:
            return
        
        # Find all connections involving this element
        connections_to_redraw = []
        for connection in self.sld_diagram.connections:
            if connection.from_element == element_id or connection.to_element == element_id:
                connections_to_redraw.append(connection)
        
        # Delete old connection lines from canvas
        for connection in connections_to_redraw:
            for item_id in connection.canvas_items:
                self.canvas.delete(item_id)
            connection.canvas_items.clear()
        
        # Redraw the connections
        for connection in connections_to_redraw:
            self.draw_connection(connection)
        
        # Update scroll region in case element moved outside current bounds
        self.update_scroll_region()
        
    def on_canvas_release(self, event):
        """Handle canvas mouse release"""
        if self.dragging and self.drag_data["element_id"]:
            element_id = self.drag_data["element_id"]
            
            # Snap to grid if enabled
            if self.grid_visible:
                if element_id in self.sld_elements:
                    element_data = self.sld_elements[element_id]
                    
                    # Get bounding box of all items
                    all_items = element_data['items']['all']
                    if all_items:
                        bbox = self.canvas.bbox(all_items[0])
                        for item_id in all_items[1:]:
                            item_bbox = self.canvas.bbox(item_id)
                            if item_bbox:
                                bbox = (
                                    min(bbox[0], item_bbox[0]),
                                    min(bbox[1], item_bbox[1]),
                                    max(bbox[2], item_bbox[2]),
                                    max(bbox[3], item_bbox[3])
                                )
                        
                        if bbox:
                            # Calculate snap position
                            grid_snap = self.grid_size * self.zoom_level
                            snapped_x = round(bbox[0] / grid_snap) * grid_snap
                            snapped_y = round(bbox[1] / grid_snap) * grid_snap
                            
                            # Calculate offset
                            dx = snapped_x - bbox[0]
                            dy = snapped_y - bbox[1]
                            
                            # Apply snap to all items
                            for item_id in all_items:
                                self.canvas.move(item_id, dx, dy)
                    
                    # Update the SLD diagram element position
                    if self.sld_diagram:
                        elem = self.sld_diagram.get_element(element_id)
                        if elem and bbox:
                            elem.x = bbox[0] / self.zoom_level
                            elem.y = bbox[1] / self.zoom_level
                            
                            # Final redraw of connections to ensure they're properly positioned
                            self.redraw_element_connections(element_id)
            
            # Reset drag state
            self.dragging = False
            self.drag_data = {"x": 0, "y": 0, "item": None, "element_id": None}
            
            # Remove highlight - ONLY from shapes, not text!
            if self.selected_element and self.selected_element in self.sld_elements:
                element_data = self.sld_elements[self.selected_element]
                if 'items' in element_data and 'shapes' in element_data['items']:
                    for item_id in element_data['items']['shapes']:
                        # Only reset width for non-text items
                        if self.canvas.type(item_id) != 'text':
                            self.canvas.itemconfig(item_id, width=2)

    def on_canvas_right_click(self, event):
        """Handle right-click for context menu"""
        # This will be implemented in Card #10
        pass
    
    def on_canvas_double_click(self, event):
        """Handle double-click for editing"""
        # This will be implemented in Card #11
        pass
    
    def on_canvas_motion(self, event):
        """Update coordinate display on mouse motion"""
        canvas_x = int(self.canvas.canvasx(event.x) / self.zoom_level)
        canvas_y = int(self.canvas.canvasy(event.y) / self.zoom_level)
        self.coord_label.configure(text=f"X: {canvas_x}, Y: {canvas_y}")
    
    def on_canvas_zoom(self, event):
        """Handle zoom with Ctrl+scroll"""
        if event.delta > 0 or event.num == 4:
            self.zoom(1.1)
        else:
            self.zoom(0.9)
    
    def on_canvas_scroll(self, event):
        """Handle regular scrolling"""
        if event.delta:
            # Windows
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        else:
            # Linux
            if event.num == 4:
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.canvas.yview_scroll(1, "units")
    
    def on_pan_start(self, event):
        """Start panning with middle mouse"""
        self.canvas.scan_mark(event.x, event.y)
    
    def on_pan_motion(self, event):
        """Pan with middle mouse"""
        self.canvas.scan_dragto(event.x, event.y, gain=1)
    
    # Toolbar actions
    def auto_layout(self):
        """Auto-layout elements"""
        if self.sld_diagram:
            self.sld_diagram.auto_layout()
            # Refresh canvas - will be implemented in Card #5
        self.status_label.configure(text="Auto-layout will be implemented in Card #7")
    
    def align_selected(self, direction):
        """Align selected elements"""
        # This will be implemented with multi-select in Card #6
        self.status_label.configure(text=f"Align {direction} will be implemented in Card #6")
    
    def zoom(self, factor):
        """Zoom the canvas"""
        # Update zoom level
        self.zoom_level *= factor
        self.zoom_level = max(0.1, min(5.0, self.zoom_level))  # Limit zoom range
        
        # Update zoom label
        self.zoom_label.configure(text=f"{int(self.zoom_level * 100)}%")
        
        # Scale all canvas items
        self.canvas.scale("all", 0, 0, factor, factor)
        
        # Update scroll region
        bbox = self.canvas.bbox("all")
        if bbox:
            self.canvas.configure(scrollregion=bbox)
        
        # Redraw grid at new scale
        self.draw_grid()
        
        self.status_label.configure(text=f"Zoom: {int(self.zoom_level * 100)}%")
    
    def fit_to_window(self):
        """Fit diagram to window"""
        bbox = self.canvas.bbox("all")
        if not bbox:
            return
        
        # Get canvas dimensions
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        
        # Calculate required scale
        diagram_width = bbox[2] - bbox[0]
        diagram_height = bbox[3] - bbox[1]
        
        if diagram_width > 0 and diagram_height > 0:
            scale_x = canvas_width / diagram_width * 0.9
            scale_y = canvas_height / diagram_height * 0.9
            scale = min(scale_x, scale_y)
            
            # Reset zoom and apply new scale
            reset_factor = 1.0 / self.zoom_level
            self.canvas.scale("all", 0, 0, reset_factor, reset_factor)
            self.zoom_level = 1.0
            
            # Apply fit scale
            self.zoom(scale)
    
    def toggle_grid(self):
        """Toggle grid visibility"""
        self.grid_visible = self.grid_var.get()
        if self.grid_visible:
            self.draw_grid()
        else:
            self.canvas.delete("grid")
    
    def toggle_connection_mode(self):
        """Toggle connection drawing mode"""
        if self.connection_mode_var.get():
            self.status_label.configure(text="Connection mode - Click elements to connect")
            self.canvas.configure(cursor="cross")
        else:
            self.status_label.configure(text="Ready")
            self.canvas.configure(cursor="")
            # Cancel any in-progress connection
            if self.temp_connection_line:
                self.canvas.delete(self.temp_connection_line)
                self.temp_connection_line = None
    
    def clear_all(self):
        """Clear all elements from canvas"""
        response = messagebox.askyesno("Clear All", 
                                       "Are you sure you want to clear all elements?")
        if response:
            self.canvas.delete("all")
            self.draw_grid()
            self.sld_diagram = None
            self.sld_elements.clear()
            self.update_status_counts()
            self.status_label.configure(text="Canvas cleared")
    
    def select_all(self):
        """Select all elements"""
        elements = self.canvas.find_withtag("element")
        for elem in elements:
            self.canvas.itemconfig(elem, width=3)
        self.selected_items = list(elements)
        self.status_label.configure(text=f"Selected {len(elements)} elements")
    
    def delete_selected(self):
        """Delete selected elements"""
        if self.selected_element:
            # Get tags of selected element
            tags = self.canvas.gettags(self.selected_element)
            for tag in tags:
                if tag not in ["element", "current"]:
                    # Delete element and its label
                    self.canvas.delete(tag)
                    self.canvas.delete(f"{tag}_label")
            
            self.selected_element = None
            self.update_status_counts()
            self.status_label.configure(text="Deleted selected element")
    
    def validate_connections(self):
        """Validate all connections"""
        # This will be implemented with connection system in Card #8
        messagebox.showinfo("Validate", 
                          "Connection validation will be implemented in Card #8")
    
    # Export functions
    def export_png(self):
        """Export diagram as PNG"""
        # This will be implemented in Card #12
        messagebox.showinfo("Export PNG", 
                          "PNG export will be implemented in Card #12")
    
    def export_pdf(self):
        """Export diagram as PDF"""
        # This will be implemented in Card #13
        messagebox.showinfo("Export PDF", 
                          "PDF export will be implemented in Card #13")
    
    # Help functions
    def show_shortcuts(self):
        """Show keyboard shortcuts"""
        shortcuts = """
        Keyboard Shortcuts:
        
        Ctrl+A      - Select All
        Delete      - Delete Selected
        Ctrl++      - Zoom In
        Ctrl+-      - Zoom Out
        Ctrl+0      - Fit to Window
        Ctrl+P      - Export PNG
        Ctrl+D      - Export PDF
        
        Middle Mouse - Pan
        Ctrl+Scroll  - Zoom
        """
        messagebox.showinfo("Keyboard Shortcuts", shortcuts)
    
    def show_about(self):
        """Show about dialog"""
        about_text = """
        Single Line Diagram Editor
        Version 1.0
        
        Part of Solar eBOS BOM Generator
        
        Create professional electrical single line
        diagrams for solar PV systems.
        """
        messagebox.showinfo("About SLD Editor", about_text)
    
    def on_close(self):
        """Handle window closing"""
        # Save diagram to project if modified
        if self.sld_diagram and self.project:
            # This will be implemented with save functionality
            pass
        
        self.destroy()