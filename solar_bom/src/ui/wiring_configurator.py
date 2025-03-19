import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Dict, ClassVar, List
from ..models.block import BlockConfig, WiringType, CollectionPoint, WiringConfig
from ..models.module import ModuleOrientation
from ..models.tracker import TrackerPosition
from ..utils.calculations import get_ampacity_for_wire_gauge, calculate_nec_current, wire_harness_compatibility

class WiringConfigurator(tk.Toplevel):

    # AWG to mm² conversion
    AWG_SIZES: ClassVar[Dict[str, float]] = {
        "4 AWG": 21.15,
        "6 AWG": 13.30,
        "8 AWG": 8.37,
        "10 AWG": 5.26
    }

    def __init__(self, parent, block: BlockConfig):
        super().__init__(parent)
        self.parent = parent
        self.block = block

        self.parent_notify_blocks_changed = getattr(parent, '_notify_blocks_changed', None)

        self.scale_factor = 10.0  # Starting scale (10 pixels per meter)
        self.pan_x = 0  # Pan offset in pixels
        self.pan_y = 0
        self.panning = False
        self.selected_whips = set()  # Set of (tracker_id, polarity) tuples
        self.dragging = False
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.drag_whips = False
        self.selection_box = None
        
        # Set up window properties
        self.title("Wiring Configuration")
        self.geometry("1200x800")
        self.minsize(800, 600)
        
        # Initialize UI
        self.setup_ui()
        
        # Make window modal
        self.transient(parent)
        self.grab_set()
        
        # Position window relative to parent
        x = parent.winfo_rootx() + 50
        y = parent.winfo_rooty() + 50
        self.geometry(f"+{x}+{y}")
        
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
        
        # Left side - Controls
        controls_frame = ttk.LabelFrame(main_container, text="Wiring Configuration", padding="5")
        controls_frame.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Wiring Type Selection
        ttk.Label(controls_frame, text="Wiring Type:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.wiring_type_var = tk.StringVar(value=WiringType.HARNESS.value)
        wiring_type_combo = ttk.Combobox(controls_frame, textvariable=self.wiring_type_var, state='readonly')
        wiring_type_combo['values'] = [t.value for t in WiringType]
        wiring_type_combo.grid(row=0, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        wiring_type_combo.bind('<<ComboboxSelected>>', self.on_wiring_type_change)
        
        # Cable Specifications
        cable_frame = ttk.LabelFrame(controls_frame, text="Cable Specifications", padding="5")
        cable_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=10, sticky=(tk.W, tk.E))

        # String Cable Size
        ttk.Label(cable_frame, text="String Cable Size:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.string_cable_size_var = tk.StringVar(value="10 AWG")
        string_cable_combo = ttk.Combobox(cable_frame, textvariable=self.string_cable_size_var, state='readonly', width=10)
        string_cable_combo['values'] = list(self.AWG_SIZES.keys())
        string_cable_combo.grid(row=0, column=1, padx=5, pady=2)
        self.string_cable_size_var.trace('w', lambda *args: self.draw_wiring_layout())

        # Wire Harness Size
        self.harness_frame = ttk.Frame(cable_frame)
        self.harness_frame.grid(row=1, column=0, columnspan=2, padx=0, pady=5, sticky=(tk.W, tk.E))
        ttk.Label(self.harness_frame, text="Harness Cable Size:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.harness_cable_size_var = tk.StringVar(value="8 AWG")
        harness_cable_combo = ttk.Combobox(self.harness_frame, textvariable=self.harness_cable_size_var, state='readonly', width=10)
        harness_cable_combo['values'] = list(self.AWG_SIZES.keys())
        harness_cable_combo.grid(row=0, column=1, padx=5, pady=2)
        self.harness_cable_size_var.trace('w', lambda *args: self.draw_wiring_layout())

        # Whip Cable Size
        self.whip_frame = ttk.Frame(cable_frame)
        self.whip_frame.grid(row=2, column=0, columnspan=2, padx=0, pady=5, sticky=(tk.W, tk.E))
        ttk.Label(self.whip_frame, text="Whip Cable Size:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.whip_cable_size_var = tk.StringVar(value="6 AWG")
        whip_cable_combo = ttk.Combobox(self.whip_frame, textvariable=self.whip_cable_size_var, state='readonly', width=10)
        whip_cable_combo['values'] = list(self.AWG_SIZES.keys())
        whip_cable_combo.grid(row=0, column=1, padx=5, pady=2)
        self.whip_cable_size_var.trace('w', lambda *args: self.draw_wiring_layout())

        # Current label toggle button
        self.show_current_labels_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(cable_frame, text="Show Current Labels", 
                        variable=self.show_current_labels_var,
                        command=self.draw_wiring_layout).grid(
                        row=3, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
        
        # Right side - Visualization
        canvas_frame = ttk.LabelFrame(main_container, text="Wiring Layout", padding="5")
        canvas_frame.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.canvas = tk.Canvas(canvas_frame, width=800, height=600, bg='white')
        self.canvas.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

         # Add warning panel frame
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
        self.wire_warnings = {}  # Maps warning IDs to wire IDs
        
        # Bind canvas resize
        self.canvas.bind('<Configure>', self.on_canvas_resize)

        # Add mouse wheel binding for zoom
        self.canvas.bind('<MouseWheel>', self.on_mouse_wheel)  # Windows
        self.canvas.bind('<Button-4>', self.on_mouse_wheel)    # Linux scroll up
        self.canvas.bind('<Button-5>', self.on_mouse_wheel)    # Linux scroll down

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
        
        # Add a reset button to the controls
        ttk.Button(controls_frame, text="Reset All Whip Points", 
                command=self.reset_all_whips).grid(
                row=4, column=0, columnspan=2, padx=5, pady=5)
        
        # Bottom buttons
        button_frame = ttk.Frame(main_container)
        button_frame.grid(row=1, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Apply", command=self.apply_configuration).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.cancel).grid(row=0, column=1, padx=5)

        # Update UI based on initial wiring type
        self.update_ui_for_wiring_type()

        # Initialize with existing configuration if available
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
            
            # Update UI based on wiring type
            self.update_ui_for_wiring_type()

        # Check if block has an existing configuration
        if self.block.wiring_config:
            self.wiring_type_var.set(self.block.wiring_config.wiring_type.value)
            self.update_ui_for_wiring_type()
        
    def update_ui_for_wiring_type(self):
        """Update UI elements based on selected wiring type"""
        is_harness = self.wiring_type_var.get() == WiringType.HARNESS.value
        if is_harness:
            self.harness_frame.grid()
        else:
            self.harness_frame.grid_remove()
        self.draw_wiring_layout()

    def draw_wiring_layout(self):
        """Draw block layout with wiring visualization"""
        self.canvas.delete("all")
        scale = self.get_canvas_scale()
        
        # Clear previous warnings
        self.clear_warnings()
        
        # Re-create warning panel (since we deleted all canvas items)
        self.setup_warning_panel()
        
        # Draw existing trackers
        for pos in self.block.tracker_positions:
            if not pos.template:
                continue
                
            template = pos.template
            
            # Get base coordinates with pan offset
            x_base = 20 + self.pan_x + pos.x * scale
            y_base = 20 + self.pan_y + pos.y * scale
            
            # Get module dimensions based on orientation
            if template.module_orientation == ModuleOrientation.PORTRAIT:
                module_height = template.module_spec.width_mm / 1000
                module_width = template.module_spec.length_mm / 1000
            else:
                module_height = template.module_spec.length_mm / 1000
                module_width = template.module_spec.width_mm / 1000
                
            # Draw torque tube through center
            self.canvas.create_line(
                x_base + module_width * scale/2, y_base,
                x_base + module_width * scale/2, y_base + (module_height * template.modules_per_string + 
                    template.module_spacing_m * (template.modules_per_string - 1) + 
                    template.motor_gap_m) * scale,
                width=3, fill='gray'
            )
            
            # Calculate number of modules
            total_modules = template.modules_per_string * template.strings_per_tracker
            modules_per_string = template.modules_per_string
            strings_above_motor = template.strings_per_tracker - 1
            modules_above_motor = modules_per_string * strings_above_motor
            modules_below_motor = modules_per_string
            
            # Draw all modules
            y_pos = y_base
            modules_drawn = 0
            
            # Draw modules above motor
            for i in range(modules_above_motor):
                self.canvas.create_rectangle(
                    x_base, y_pos,
                    x_base + module_width * scale, 
                    y_pos + module_height * scale,
                    fill='lightblue', outline='blue'
                )
                modules_drawn += 1
                y_pos += (module_height + template.module_spacing_m) * scale
            
            # Draw motor
            motor_y = y_pos
            self.canvas.create_oval(
                x_base + module_width * scale/2 - 5, motor_y - 5,
                x_base + module_width * scale/2 + 5, motor_y + 5,
                fill='red'
            )
            y_pos += template.motor_gap_m * scale
            
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
            tracker_idx = self.block.tracker_positions.index(pos)
            pos_whip = self.get_whip_position(str(tracker_idx), 'positive')
            tracker_idx = self.block.tracker_positions.index(pos)
            neg_whip = self.get_whip_position(str(tracker_idx), 'negative')
            
            if pos_whip:
                wx = 20 + self.pan_x + pos_whip[0] * scale
                wy = 20 + self.pan_y + pos_whip[1] * scale
                self.canvas.create_oval(
                    wx - 3, wy - 3,
                    wx + 3, wy + 3,
                    fill='pink',
                    outline='deeppink',
                    tags='whip_point'
                )
            
            if neg_whip:
                wx = 20 + self.pan_x + neg_whip[0] * scale
                wy = 20 + self.pan_y + neg_whip[1] * scale
                self.canvas.create_oval(
                    wx - 3, wy - 3,
                    wx + 3, wy + 3,
                    fill='teal',
                    outline='darkcyan',
                    tags='whip_point'
                )

        # Draw inverter/combiner
        if self.block.device_x is not None and self.block.device_y is not None:
            device_x = 20 + self.pan_x + self.block.device_x * scale
            device_y = 20 + self.pan_y + self.block.device_y * scale
            device_size = 0.91 * scale  # 3ft = 0.91m
            self.canvas.create_rectangle(
                device_x, device_y,
                device_x + device_size,
                device_y + device_size,
                fill='red', outline='darkred',
                tags='device'
            )

        # Draw device destination points
        self.draw_device_destination_points()

        # Draw routes if wiring type is selected
        if self.wiring_type_var.get() == WiringType.HOMERUN.value:
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
                        points = [(20 + self.pan_x + x * scale, 
                                20 + self.pan_y + y * scale) for x, y in route1]
                        if len(points) > 1:
                            current = self.calculate_current_for_segment('string')
                            self.draw_wire_segment(points, self.string_cable_size_var.get(), 
                                                current, is_positive=True, segment_type="string")
                            self.add_current_label(points, current, is_positive=True)
                        
                    if neg_whip:
                        route1 = self.calculate_cable_route(
                            pos.x + string.negative_source_x,
                            pos.y + string.negative_source_y,
                            neg_whip[0], neg_whip[1],
                            False, len(neg_routes)
                        )
                        points = [(20 + self.pan_x + x * scale, 
                                20 + self.pan_y + y * scale) for x, y in route1]
                        if len(points) > 1:
                            current = self.calculate_current_for_segment('string')
                            self.draw_wire_segment(points, self.string_cable_size_var.get(), 
                                                current, is_positive=False, segment_type="string")
                            self.add_current_label(points, current, is_positive=False)
                                                
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
                points = [(20 + self.pan_x + x * scale, 
                        20 + self.pan_y + y * scale) for x, y in route]
                if len(points) > 1:
                    current = self.calculate_current_for_segment('whip', num_strings=1) # For homerun, it's 1 string per route
                    self.draw_wire_segment(points, self.whip_cable_size_var.get(), 
                                        current, is_positive=True, segment_type="whip")
                    self.add_current_label(points, current, is_positive=True, segment_type='whip')
            
            for i, route_info in enumerate(neg_routes_north + neg_routes_south):
                route = self.calculate_cable_route(
                    route_info['source_x'],
                    route_info['source_y'],
                    neg_dest[0],  # Right side of device
                    neg_dest[1],  # Device y-position
                    False,  # is_positive
                    i  # route_index
                )
                points = [(20 + self.pan_x + x * scale, 
                        20 + self.pan_y + y * scale) for x, y in route]
                if len(points) > 1:
                    current = self.calculate_current_for_segment('whip', num_strings=1) # For homerun, it's 1 string per route
                    self.draw_wire_segment(points, self.whip_cable_size_var.get(), 
                                        current, is_positive=False, segment_type="whip")
                    self.add_current_label(points, current, is_positive=False, segment_type='whip')
                    
        else:  # Wire Harness configuration
            # Process each tracker
            for pos in self.block.tracker_positions:
                # Calculate node points
                pos_nodes = self.calculate_node_points(pos, True)  # Positive nodes
                neg_nodes = self.calculate_node_points(pos, False)  # Negative nodes
                
                # Get whip points
                tracker_idx = self.block.tracker_positions.index(pos)
                pos_whip = self.get_whip_position(str(tracker_idx), 'positive')
                tracker_idx = self.block.tracker_positions.index(pos)
                neg_whip = self.get_whip_position(str(tracker_idx), 'negative')
                
                # Draw string cables to node points
                for i, string in enumerate(pos.strings):
                    # Positive string cable
                    source_x = pos.x + string.positive_source_x
                    source_y = pos.y + string.positive_source_y
                    route = self.calculate_cable_route(
                        source_x, source_y,
                        pos_nodes[i][0], pos_nodes[i][1],
                        True, i
                    )
                    points = [(20 + self.pan_x + x * scale, 
                            20 + self.pan_y + y * scale) for x, y in route]
                    if len(points) > 1:
                        current = self.calculate_current_for_segment('string')
                        self.draw_wire_segment(points, self.string_cable_size_var.get(), 
                                            current, is_positive=True, segment_type="string")
                        self.add_current_label(points, current, is_positive=True)
                        
                    # Negative string cable
                    source_x = pos.x + string.negative_source_x
                    source_y = pos.y + string.negative_source_y
                    route = self.calculate_cable_route(
                        source_x, source_y,
                        neg_nodes[i][0], neg_nodes[i][1],
                        False, i
                    )
                    points = [(20 + self.pan_x + x * scale, 
                            20 + self.pan_y + y * scale) for x, y in route]
                    if len(points) > 1:
                        current = self.calculate_current_for_segment('string')
                        self.draw_wire_segment(points, self.string_cable_size_var.get(), 
                                            current, is_positive=False, segment_type="string")
                        self.add_current_label(points, current, is_positive=False)
                
                # Draw node points
                for nx, ny in pos_nodes:
                    x = 20 + self.pan_x + nx * scale
                    y = 20 + self.pan_y + ny * scale
                    self.canvas.create_oval(x-3, y-3, x+3, y+3, fill='red', outline='darkred')
                
                for nx, ny in neg_nodes:
                    x = 20 + self.pan_x + nx * scale
                    y = 20 + self.pan_y + ny * scale
                    self.canvas.create_oval(x-3, y-3, x+3, y+3, fill='blue', outline='darkblue')
                
                # Route from node points through whip points to device
                pos_dest, neg_dest = self.get_device_destination_points()
                
                # Draw node connections - bottom to top for north device
                # Positive harness
                sorted_pos_nodes = sorted(pos_nodes, key=lambda p: p[1], reverse=True)  # Sort by y coordinate
                for i in range(len(sorted_pos_nodes)):
                    start = sorted_pos_nodes[i]
                    if i < len(sorted_pos_nodes) - 1:
                        # Connect to next node
                        end = sorted_pos_nodes[i + 1]
                        route = self.calculate_cable_route(
                            start[0], start[1],
                            end[0], end[1],
                            True, 0
                        )
                    else:
                        # Last (northernmost) node routes to whip point
                        if pos_whip:
                            route = self.calculate_cable_route(
                                start[0], start[1],
                                pos_whip[0], pos_whip[1],
                                True, 0
                            )
                        else:
                            continue
                    points = [(20 + self.pan_x + x * scale, 
                            20 + self.pan_y + y * scale) for x, y in route]
                    if len(points) > 1:
                        # Calculate accumulated current based on position in chain
                        # i+1 strings are accumulated at this point
                        accumulated_strings = i + 1
                        current = self.calculate_current_for_segment('string') * accumulated_strings
                        self.draw_wire_segment(points, self.harness_cable_size_var.get(), 
                                            current, is_positive=True, segment_type="harness")
                        # Add current label showing accumulated current
                        if len(points) >= 2 and self.show_current_labels_var.get():
                            mid_idx = len(points) // 2
                            mid_x = (points[mid_idx-1][0] + points[mid_idx][0]) / 2
                            mid_y = (points[mid_idx-1][1] + points[mid_idx][1]) / 2
                            self.canvas.create_text(mid_x, mid_y-10, text=f"{current:.1f}A", 
                                                fill='red', font=('Arial', 8))

                # Negative harness
                sorted_neg_nodes = sorted(neg_nodes, key=lambda p: p[1], reverse=True)  # Sort by y coordinate
                for i in range(len(sorted_neg_nodes)):
                    start = sorted_neg_nodes[i]
                    if i < len(sorted_neg_nodes) - 1:
                        # Connect to next node
                        end = sorted_neg_nodes[i + 1]
                        route = self.calculate_cable_route(
                            start[0], start[1],
                            end[0], end[1],
                            False, 0
                        )
                    else:
                        # Last (northernmost) node routes to whip point
                        if neg_whip:
                            route = self.calculate_cable_route(
                                start[0], start[1],
                                neg_whip[0], neg_whip[1],
                                False, 0
                            )
                        else:
                            continue
                    points = [(20 + self.pan_x + x * scale, 
                            20 + self.pan_y + y * scale) for x, y in route]
                    if len(points) > 1:
                        # Calculate accumulated current based on position in chain
                        # i+1 strings are accumulated at this point
                        accumulated_strings = i + 1
                        current = self.calculate_current_for_segment('string') * accumulated_strings
                        self.draw_wire_segment(points, self.harness_cable_size_var.get(), 
                                            current, is_positive=False, segment_type="harness")
                        # Add current label showing accumulated current
                        if len(points) >= 2 and self.show_current_labels_var.get():
                            mid_idx = len(points) // 2
                            mid_x = (points[mid_idx-1][0] + points[mid_idx][0]) / 2
                            mid_y = (points[mid_idx-1][1] + points[mid_idx][1]) / 2
                            self.canvas.create_text(mid_x, mid_y+10, text=f"{current:.1f}A", 
                                                fill='blue', font=('Arial', 8))
                                    
                # Route from whip point to device
                pos_dest, neg_dest = self.get_device_destination_points()

                # Draw positive whip to device route
                if pos_whip and pos_dest:
                    route = self.calculate_cable_route(
                        pos_whip[0], pos_whip[1],
                        pos_dest[0], pos_dest[1],
                        True, 0
                    )
                    points = [(20 + self.pan_x + x * scale, 
                            20 + self.pan_y + y * scale) for x, y in route]
                    if len(points) > 1:
                        # Total current for all strings on this tracker
                        total_strings = len(pos.strings)
                        current = self.calculate_current_for_segment('whip', total_strings)
                        self.draw_wire_segment(points, self.whip_cable_size_var.get(), 
                                            current, is_positive=True, segment_type="whip")
                        # Add current label at midpoint of line
                        if len(points) >= 2 and self.show_current_labels_var.get():
                            mid_idx = len(points) // 2
                            mid_x = (points[mid_idx-1][0] + points[mid_idx][0]) / 2
                            mid_y = (points[mid_idx-1][1] + points[mid_idx][1]) / 2
                            self.canvas.create_text(mid_x, mid_y-10, text=f"{current:.1f}A", 
                                                fill='red', font=('Arial', 8))

                # Draw negative whip to device route
                if neg_whip and neg_dest:
                    route = self.calculate_cable_route(
                        neg_whip[0], neg_whip[1],
                        neg_dest[0], neg_dest[1],
                        False, 0
                    )
                    points = [(20 + self.pan_x + x * scale, 
                            20 + self.pan_y + y * scale) for x, y in route]
                    if len(points) > 1:
                        # Total current for all strings on this tracker
                        total_strings = len(pos.strings)
                        current = self.calculate_current_for_segment('whip', total_strings)
                        self.draw_wire_segment(points, self.whip_cable_size_var.get(), 
                                            current, is_positive=False, segment_type="whip")
                        # Add current label at midpoint of line
                        if len(points) >= 2 and self.show_current_labels_var.get():
                            mid_idx = len(points) // 2
                            mid_x = (points[mid_idx-1][0] + points[mid_idx][0]) / 2
                            mid_y = (points[mid_idx-1][1] + points[mid_idx][1]) / 2
                            self.canvas.create_text(mid_x, mid_y+10, text=f"{current:.1f}A", 
                                                fill='blue', font=('Arial', 8))
                        
    def draw_collection_points(self, pos: TrackerPosition, x: float, y: float, scale: float):
        """Draw collection points for a tracker position"""
        for string in pos.strings:
            # Draw positive source point (red circle)
            px = x + string.positive_source_x * scale
            py = y + string.positive_source_y * scale
            self.canvas.create_oval(
                px - 3, py - 3,
                px + 3, py + 3,
                fill='red',
                outline='darkred',
                tags='collection_point'
            )

            # Draw negative source point (blue circle)
            nx = x + string.negative_source_x * scale
            ny = y + string.negative_source_y * scale
            self.canvas.create_oval(
                nx - 3, ny - 3,
                nx + 3, ny + 3,
                fill='blue',
                outline='darkblue',
                tags='collection_point'
            )
        
        # Draw whip points
        tracker_idx = self.block.tracker_positions.index(pos)
        tracker_id = str(tracker_idx)
        
        # Get whip point positions
        pos_whip = self.get_whip_position(tracker_id, 'positive')
        neg_whip = self.get_whip_position(tracker_id, 'negative')
        
        # Draw positive whip point
        if pos_whip:
            wx = 20 + self.pan_x + pos_whip[0] * scale
            wy = 20 + self.pan_y + pos_whip[1] * scale
            
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
            wx = 20 + self.pan_x + neg_whip[0] * scale
            wy = 20 + self.pan_y + neg_whip[1] * scale
            
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
        
    def on_canvas_resize(self, event):
        """Handle canvas resize event"""
        if event.width > 1 and event.height > 1:  # Ensure valid dimensions
            self.draw_wiring_layout()
        
    def apply_configuration(self):
        """Apply wiring configuration to block"""
        if not self.block:
            tk.messagebox.showerror("Error", "No block selected")
            self.destroy()
            return
            
        try:
            # Get the selected wiring type
            wiring_type_str = self.wiring_type_var.get()
            wiring_type = WiringType(wiring_type_str)
            
            # Create collection points
            positive_collection_points = []
            negative_collection_points = []
            strings_per_collection = {}
            cable_routes = {}
            
            # Process each tracker position to capture collection points and routes
            for idx, pos in enumerate(self.block.tracker_positions):
                if not pos.template:
                    continue
                    
                # Get whip points for this tracker
                tracker_idx = self.block.tracker_positions.index(pos)
                pos_whip = self.get_whip_position(str(tracker_idx), 'positive')
                tracker_idx = self.block.tracker_positions.index(pos)
                neg_whip = self.get_whip_position(str(tracker_idx), 'negative')
                
                # Add collection points
                if pos_whip:
                    point_id = idx  # Use tracker index as point ID
                    collection_point = CollectionPoint(
                        x=pos_whip[0],
                        y=pos_whip[1],
                        connected_strings=[s.index for s in pos.strings],
                        current_rating=self.calculate_current_for_segment('whip', num_strings=len(pos.strings))
                    )
                    positive_collection_points.append(collection_point)
                    strings_per_collection[point_id] = len(pos.strings)
                
                if neg_whip:
                    point_id = idx  # Use tracker index as point ID
                    collection_point = CollectionPoint(
                        x=neg_whip[0],
                        y=neg_whip[1],
                        connected_strings=[s.index for s in pos.strings],
                        current_rating=self.calculate_current_for_segment('whip', num_strings=len(pos.strings))
                    )
                    negative_collection_points.append(collection_point)
                
                # Add routes based on wiring type - these methods should populate the cable_routes dict
                for string_idx, string in enumerate(pos.strings):
                    if wiring_type == WiringType.HOMERUN:
                        self.add_homerun_routes(cable_routes, pos, string, idx, string_idx, pos_whip, neg_whip)
                    else:
                        self.add_harness_routes(cable_routes, pos, string, idx, string_idx, pos_whip, neg_whip)
            
            custom_whip_points = {}
            if (hasattr(self.block, 'wiring_config') and 
                self.block.wiring_config and 
                hasattr(self.block.wiring_config, 'custom_whip_points')):
                custom_whip_points = self.block.wiring_config.custom_whip_points

            # Create the WiringConfig instance
            wiring_config = WiringConfig(
                wiring_type=wiring_type,
                positive_collection_points=positive_collection_points,
                negative_collection_points=negative_collection_points,
                strings_per_collection=strings_per_collection,
                cable_routes=cable_routes,
                string_cable_size=self.string_cable_size_var.get(),
                harness_cable_size=self.harness_cable_size_var.get(),
                whip_cable_size=self.whip_cable_size_var.get(),
                custom_whip_points=custom_whip_points
                )
            
            # Store the configuration in the block
            self.block.wiring_config = wiring_config
            
            tk.messagebox.showinfo("Success", "Wiring configuration applied successfully")

            # Check MPPT capacity
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
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            tk.messagebox.showerror("Error", f"Failed to apply wiring configuration: {str(e)}")
        
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
            
        scale = self.get_canvas_scale()
        
        # Draw positive destination point (red)
        px = 20 + self.pan_x + pos_point[0] * scale
        py = 20 + self.pan_y + pos_point[1] * scale
        self.canvas.create_oval(
            px - 4, py - 4,
            px + 4, py + 4,
            fill='red',  # Match source point colors
            outline='darkred',
            tags='destination_point'
        )
        
        # Draw negative destination point (blue)
        nx = 20 + self.pan_x + neg_point[0] * scale
        ny = 20 + self.pan_y + neg_point[1] * scale
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
        
        # Determine if device is north or south of tracker
        device_is_north = self.block.device_y < pos.y
        
        # Calculate Y position
        if device_is_north:
            y = pos.y - whip_offset
        else:
            y = pos.y + tracker_height + whip_offset
        
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
        """
        Calculate current flowing through a wire segment based on configuration
        """
        # Direct access to module spec from block's tracker template
        if (self.block and 
            hasattr(self.block, 'tracker_template') and 
            self.block.tracker_template and 
            hasattr(self.block.tracker_template, 'module_spec') and 
            self.block.tracker_template.module_spec):
            
            module_spec = self.block.tracker_template.module_spec
            
            # Use the Imp value
            string_current = module_spec.imp
        else:
            # Fallback if we can't get the actual current
            string_current = 10.0
        
        # Calculate current based on how many strings are combined
        return string_current * num_strings

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
            # Node-to-node to whip routes (for all nodes)
            if pos_nodes and pos_whip:
                cable_routes[f"pos_harness_{tracker_idx}"] = pos_nodes + [pos_whip]
            
            if neg_nodes and neg_whip:
                cable_routes[f"neg_harness_{tracker_idx}"] = neg_nodes + [neg_whip]
            
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
            tracker_id, polarity = hit_whip
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
            # Calculate movement in world coordinates
            scale = self.get_canvas_scale()
            dx = (event.x - self.drag_start_x) / scale
            dy = (event.y - self.drag_start_y) / scale
            
            # Update custom whip point positions
            for tracker_id, polarity in self.selected_whips:
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
            
            # Update drag start for continuous dragging
            self.drag_start_x = event.x
            self.drag_start_y = event.y
            
            # Redraw
            self.draw_wiring_layout()
        elif self.dragging:
            # Update selection box
            if self.selection_box:
                self.canvas.delete(self.selection_box)
                
            self.selection_box = self.canvas.create_rectangle(
                self.drag_start_x, self.drag_start_y, event.x, event.y,
                outline='blue', dash=(4, 4)
            )

    def on_canvas_release(self, event):
        """Handle end of drag operation"""
        if self.panning:
            return
            
        if self.drag_whips:
            self.drag_whips = False
        elif self.dragging and self.selection_box:
            # Convert selection box to world coordinates and select whips within it
            scale = self.get_canvas_scale()
            x1 = min(self.drag_start_x, event.x)
            y1 = min(self.drag_start_y, event.y)
            x2 = max(self.drag_start_x, event.x)
            y2 = max(self.drag_start_y, event.y)
            
            # Convert to world coordinates
            wx1 = (x1 - 20 - self.pan_x) / scale
            wy1 = (y1 - 20 - self.pan_y) / scale
            wx2 = (x2 - 20 - self.pan_x) / scale
            wy2 = (y2 - 20 - self.pan_y) / scale
            
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
            
        for tracker_id, polarity in self.selected_whips:
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
            if hasattr(self.block.wiring_config, 'custom_whip_points'):
                self.block.wiring_config.custom_whip_points = {}
                
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

    def get_whip_position(self, tracker_id, polarity):
        """Get the current position of a whip point (custom or default)"""
        # Check for custom position first
        if (hasattr(self.block, 'wiring_config') and 
            self.block.wiring_config and 
            hasattr(self.block.wiring_config, 'custom_whip_points') and
            tracker_id in self.block.wiring_config.custom_whip_points and
            polarity in self.block.wiring_config.custom_whip_points[tracker_id]):
            
            return self.block.wiring_config.custom_whip_points[tracker_id][polarity]
            
        # Fall back to default position
        return self.get_whip_default_position(tracker_id, polarity)
        
    def get_whip_default_position(self, tracker_id, polarity):
        """Calculate the default position for a whip point"""
        tracker_idx = int(tracker_id)
        if tracker_idx < 0 or tracker_idx >= len(self.block.tracker_positions):
            return None
            
        pos = self.block.tracker_positions[tracker_idx]
        is_positive = (polarity == 'positive')
        
        # Use the existing calculation method
        if is_positive:
            return self.calculate_whip_points(pos, True)
        else:
            return self.calculate_whip_points(pos, False)
            
    def get_whip_at_position(self, x, y, tolerance=0.5):
        """Find a whip point at the given position within tolerance"""
        # Check each tracker's whip points
        for tracker_idx, pos in enumerate(self.block.tracker_positions):
            tracker_id = str(tracker_idx)
            
            # Check positive whip
            pos_whip = self.get_whip_position(tracker_id, 'positive')
            if pos_whip:
                if abs(x - pos_whip[0]) <= tolerance and abs(y - pos_whip[1]) <= tolerance:
                    return (tracker_id, 'positive')
                    
            # Check negative whip
            neg_whip = self.get_whip_position(tracker_id, 'negative')
            if neg_whip:
                if abs(x - neg_whip[0]) <= tolerance and abs(y - neg_whip[1]) <= tolerance:
                    return (tracker_id, 'negative')
                    
        return None
    
    def show_context_menu(self, event):
        """Show context menu for whip points"""
        # First, check if we clicked on a whip point
        scale = self.get_canvas_scale()
        world_x = (event.x - 20 - self.pan_x) / scale
        world_y = (event.y - 20 - self.pan_y) / scale
        
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
        
    def draw_wire_with_warning(self, points, wire_gauge, current, is_positive=True):
        """Draw a wire segment with appropriate warnings based on loading"""
        # Get standard wire properties
        line_thickness = self.get_line_thickness_for_wire(wire_gauge)
        base_color = 'red' if is_positive else 'blue'
        
        # Draw the wire with standard color
        line_id = self.canvas.create_line(points, fill=base_color, width=line_thickness)
        
        # Check loading against ampacity
        ampacity = get_ampacity_for_wire_gauge(wire_gauge)
        if ampacity == 0:
            return line_id  # Unknown wire size
        
        # Apply NEC factors (125% of continuous current)
        nec_current = calculate_nec_current(current)
        
        # Calculate load percentage
        load_percent = (nec_current / ampacity) * 100
        
        # Add warning indicator based on loading
        if load_percent >= 100:
            indicator_color = 'red'
            bg_color = '#ffeeee'  # Light red background
            indicator_text = "⚠ OVERLOAD"
        elif load_percent >= 80:
            indicator_color = '#dd6600'  # Dark orange - more readable
            bg_color = '#fff0e0'  # Light orange background
            indicator_text = "⚠ WARNING"
        elif load_percent >= 60:
            indicator_color = '#775500'  # Dark yellow/amber - more readable
            bg_color = '#ffffd0'  # Light yellow background
            indicator_text = "⚠ CAUTION"
        else:
            return line_id  # No warning needed
        
        # Find midpoint for warning
        if len(points) >= 2:
            mid_idx = len(points) // 2
            mid_x = (points[mid_idx-1][0] + points[mid_idx][0]) / 2
            mid_y = (points[mid_idx-1][1] + points[mid_idx][1]) / 2
            
            # Create warning indicator
            tag_name = f"wire_warning_{line_id}"
            
            # Create a background rectangle for better visibility
            self.canvas.create_rectangle(
                mid_x - 40, mid_y - 12,
                mid_x + 40, mid_y + 12,
                fill=bg_color, outline=indicator_color,
                tags=tag_name
            )
            
            # Create warning text
            self.canvas.create_text(
                mid_x, mid_y,
                text=f"{indicator_text} ({load_percent:.0f}%)",
                fill=indicator_color,
                font=('Arial', 8, 'bold'),
                tags=tag_name
            )
        
        return line_id
    
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
        # Remove any existing highlights
        self.canvas.delete("wire_highlight")
        
        # Get the wire's coordinates
        wire_coords = self.canvas.coords(wire_id)
        if not wire_coords:
            return
        
        # Get the current width and convert to int before adding
        try:
            current_width = int(float(self.canvas.itemcget(wire_id, 'width')))
        except (ValueError, TypeError):
            current_width = 2  # Default width if conversion fails
        
        # Get wire color to determine a contrasting highlight color
        wire_color = self.canvas.itemcget(wire_id, 'fill')
        highlight_color = '#FF00FF'  # Bright magenta works well for both red and blue wires
        
        # Create dashed outline effect with double the thickness
        self.canvas.create_line(
            wire_coords,
            width=current_width * 2 + 4,
            fill=highlight_color,
            dash=(8, 4),  # Dashed line pattern
            tags="wire_highlight"
        )
        
        # Animate the highlight line (expand and contract)
        self.pulse_highlight(6, current_width * 2 + 4)
        
        # Also scroll to make this wire visible if it's outside view
        if len(wire_coords) >= 2:
            mid_x = (wire_coords[0] + wire_coords[2]) / 2
            mid_y = (wire_coords[1] + wire_coords[3]) / 2
            self.scroll_to_position(mid_x, mid_y)

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

    def scroll_to_position(self, x, y):
        """Scroll the canvas to make a position visible"""
        # Get canvas dimensions
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        
        # Check if the point is outside the visible area
        if x < 0 or x > canvas_width or y < 0 or y > canvas_height:
            # Adjust pan to center on the point
            self.pan_x = canvas_width/2 - x + 20
            self.pan_y = canvas_height/2 - y + 20
            
            # Redraw with new pan values
            self.draw_wiring_layout()

    def flash_highlight(self, count):
        """Flash the highlight by toggling visibility"""
        if count <= 0:
            return
            
        # Toggle visibility
        current_state = self.canvas.itemcget("wire_highlight", 'state')
        new_state = 'hidden' if current_state == 'normal' else 'normal'
        self.canvas.itemconfigure("wire_highlight", state=new_state)
        
        # Schedule next toggle
        self.after(300, lambda: self.flash_highlight(count - 1))

    def clear_warnings(self):
        """Clear all warnings from the panel"""
        for widget in self.warning_frame.winfo_children():
            widget.destroy()
        self.wire_warnings = {}

    def draw_wire_segment(self, points, wire_gauge, current, is_positive=True, segment_type="string"):
        """Draw a wire segment with warnings only for overloads"""
        # Get standard wire properties
        line_thickness = self.get_line_thickness_for_wire(wire_gauge)
        base_color = 'red' if is_positive else 'blue'
        
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
            self.add_wire_warning(
                line_id,
                f"{polarity.capitalize()} {segment_type} {wire_gauge}: {load_percent:.0f}% (OVERLOAD)",
                'overload'
            )
        
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