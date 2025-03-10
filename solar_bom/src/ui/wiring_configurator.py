import tkinter as tk
from tkinter import ttk
from typing import Optional, Dict, ClassVar, List
from ..models.block import BlockConfig, WiringType
from ..models.module import ModuleOrientation
from ..models.tracker import TrackerPosition

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
        self.scale_factor = 10.0  # Starting scale (10 pixels per meter)
        self.pan_x = 0  # Pan offset in pixels
        self.pan_y = 0
        self.panning = False
        
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

        # Current label toggle button
        self.show_current_labels_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(cable_frame, text="Show Current Labels", 
                        variable=self.show_current_labels_var,
                        command=self.draw_wiring_layout).grid(
                        row=2, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
        
        # Right side - Visualization
        canvas_frame = ttk.LabelFrame(main_container, text="Wiring Layout", padding="5")
        canvas_frame.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.canvas = tk.Canvas(canvas_frame, width=800, height=600, bg='white')
        self.canvas.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
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
        
        # Bottom buttons
        button_frame = ttk.Frame(main_container)
        button_frame.grid(row=1, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Apply", command=self.apply_configuration).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.cancel).grid(row=0, column=1, padx=5)

        # Update UI based on initial wiring type
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
            pos_whip = self.calculate_whip_points(pos, True)
            neg_whip = self.calculate_whip_points(pos, False)
            
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
                pos_whip = self.calculate_whip_points(pos, True)
                neg_whip = self.calculate_whip_points(pos, False)
                
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
                            line_thickness = self.get_line_thickness_for_wire(self.string_cable_size_var.get())
                            self.canvas.create_line(points, fill='red', width=line_thickness)
                            current = self.calculate_current_for_segment('string')
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
                            line_thickness = self.get_line_thickness_for_wire(self.string_cable_size_var.get())
                            self.canvas.create_line(points, fill='blue', width=line_thickness)
                            current = self.calculate_current_for_segment('string')
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
                    thickness = self.get_line_thickness_for_wire(self.string_cable_size_var.get())
                    self.canvas.create_line(points, fill='red', width=thickness)
                    current = self.calculate_current_for_segment('whip', num_strings=1) # For homerun, it's 1 string per route
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
                    thickness = self.get_line_thickness_for_wire(self.string_cable_size_var.get())
                    self.canvas.create_line(points, fill='blue', width=thickness)
                    current = self.calculate_current_for_segment('whip', num_strings=1) # For homerun, it's 1 string per route
                    self.add_current_label(points, current, is_positive=True, segment_type='whip')
                    
        else:  # Wire Harness configuration
            # Process each tracker
            for pos in self.block.tracker_positions:
                # Calculate node points
                pos_nodes = self.calculate_node_points(pos, True)  # Positive nodes
                neg_nodes = self.calculate_node_points(pos, False)  # Negative nodes
                
                # Get whip points
                pos_whip = self.calculate_whip_points(pos, True)
                neg_whip = self.calculate_whip_points(pos, False)
                
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
                        line_thickness = self.get_line_thickness_for_wire(self.string_cable_size_var.get())
                        self.canvas.create_line(points, fill='red', width=line_thickness)
                        current = self.calculate_current_for_segment('string')
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
                        line_thickness = self.get_line_thickness_for_wire(self.string_cable_size_var.get())
                        self.canvas.create_line(points, fill='blue', width=line_thickness)
                        current = self.calculate_current_for_segment('string')
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
                        line_thickness = self.get_line_thickness_for_wire(self.harness_cable_size_var.get())
                        line_id = self.canvas.create_line(points, fill='red', width=line_thickness)
                        # Add current label showing accumulated current
                        if len(points) >= 2 and self.show_current_labels_var.get():
                            mid_idx = len(points) // 2
                            mid_x = (points[mid_idx-1][0] + points[mid_idx][0]) / 2
                            mid_y = (points[mid_idx-1][1] + points[mid_idx][1]) / 2
                            # Calculate accumulated current based on position in chain
                            # i+1 strings are accumulated at this point
                            accumulated_strings = i + 1
                            current = self.calculate_current_for_segment('string') * accumulated_strings
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
                        line_thickness = self.get_line_thickness_for_wire(self.harness_cable_size_var.get())
                        line_id = self.canvas.create_line(points, fill='blue', width=line_thickness)
                        # Add current label showing accumulated current
                        if len(points) >= 2 and self.show_current_labels_var.get():
                            mid_idx = len(points) // 2
                            mid_x = (points[mid_idx-1][0] + points[mid_idx][0]) / 2
                            mid_y = (points[mid_idx-1][1] + points[mid_idx][1]) / 2
                            # Calculate accumulated current based on position in chain
                            # i+1 strings are accumulated at this point
                            accumulated_strings = i + 1
                            current = self.calculate_current_for_segment('string') * accumulated_strings
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
                        line_thickness = self.get_line_thickness_for_wire(self.string_cable_size_var.get())
                        line_id = self.canvas.create_line(points, fill='red', width=line_thickness)
                        # Add current label at midpoint of line
                        if len(points) >= 2 and self.show_current_labels_var.get():
                            mid_idx = len(points) // 2
                            mid_x = (points[mid_idx-1][0] + points[mid_idx][0]) / 2
                            mid_y = (points[mid_idx-1][1] + points[mid_idx][1]) / 2
                            # Total current for all strings on this tracker
                            total_strings = len(pos.strings)
                            current = self.calculate_current_for_segment('whip', total_strings)
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
                        line_thickness = self.get_line_thickness_for_wire(self.string_cable_size_var.get())
                        line_id = self.canvas.create_line(points, fill='blue', width=line_thickness)
                        # Add current label at midpoint of line
                        if len(points) >= 2 and self.show_current_labels_var.get():
                            mid_idx = len(points) // 2
                            mid_x = (points[mid_idx-1][0] + points[mid_idx][0]) / 2
                            mid_y = (points[mid_idx-1][1] + points[mid_idx][1]) / 2
                            # Total current for all strings on this tracker
                            total_strings = len(pos.strings)
                            current = self.calculate_current_for_segment('whip', total_strings)
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

    def on_wiring_type_change(self, event=None):
        """Handle wiring type selection change"""
        self.update_ui_for_wiring_type()
        
    def on_canvas_resize(self, event):
        """Handle canvas resize event"""
        if event.width > 1 and event.height > 1:  # Ensure valid dimensions
            self.draw_wiring_layout()
        
    def apply_configuration(self):
        """Apply wiring configuration to block"""
        # TODO: Implement configuration application
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
        
        if self.wiring_type_var.get() == WiringType.HARNESS.value:
            # For harness configuration, route directly to node point
            route.append((dest_x, dest_y))
        else:
            # For homerun configuration, use horizontal-then-vertical routing
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

    def calculate_current_for_segment(self, segment_type: str, strings_combined: int = 1) -> float:
        """
        Calculate current flowing through a wire segment based on configuration
        
        Args:
            segment_type: Type of segment ('string', 'harness', or 'whip')
            strings_combined: Number of strings combined in this segment
            
        Returns:
            float: Current in amperes
        """
        # Get module current value from template
        if not self.block or not self.block.tracker_template or not self.block.tracker_template.module_spec:
            return 0.0
            
        module_spec = self.block.tracker_template.module_spec
        
        # Use Imp as the normal operating current per string
        string_current = module_spec.imp
        
        # Calculate current based on how many strings are combined at this point
        return string_current * strings_combined

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