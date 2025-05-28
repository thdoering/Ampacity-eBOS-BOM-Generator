import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Dict, ClassVar, List
from ..models.block import BlockConfig, WiringType, CollectionPoint, WiringConfig, HarnessGroup
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

    # Standard fuse ratings in amps
    FUSE_RATINGS: ClassVar[List[int]] = [5, 10, 15, 20, 25, 30, 35, 40, 45]

    def __init__(self, parent, block: BlockConfig):
        super().__init__(parent)
        self.parent = parent
        self.block = block

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
        
        # Set up window properties
        self.title("Wiring Configuration")
        self.geometry("1200x800")
        self.minsize(1000, 600)
        
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

        # Add Tracker Harness Configuration section - only visible in Wire Harness mode
        self.harness_config_frame = ttk.LabelFrame(controls_frame, text="Tracker Harness Configuration", padding="5")
        self.harness_config_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=10, sticky=(tk.W, tk.E))

        # String count selector instead of tracker selector
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

        # Harness display frame
        self.harness_display_frame = ttk.LabelFrame(self.harness_config_frame, text="Current Harnesses")
        self.harness_display_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=5, sticky=(tk.W, tk.E))

        # Initially hide harness configuration if not in harness mode
        if self.wiring_type_var.get() != WiringType.HARNESS.value:
            self.harness_config_frame.grid_remove()

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
        self.whip_cable_size_var = tk.StringVar(value="8 AWG")
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
        
        # Add routing controls
        routing_frame = ttk.Frame(controls_frame)
        routing_frame.grid(row=4, column=0, columnspan=2, padx=5, pady=5)

        ttk.Button(routing_frame, text="Reset All Whip Points", 
                command=self.reset_all_whips).grid(row=0, column=0, padx=5)

        # Add realistic/conceptual toggle
        self.routing_mode_var = tk.StringVar(value="realistic")
        ttk.Radiobutton(routing_frame, text="Realistic Routing", 
                    variable=self.routing_mode_var, value="realistic",
                    command=self.draw_wiring_layout).grid(row=0, column=1, padx=5)
        ttk.Radiobutton(routing_frame, text="Conceptual Routing", 
                    variable=self.routing_mode_var, value="conceptual",
                    command=self.draw_wiring_layout).grid(row=0, column=2, padx=5)
        
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

            # Restore routing mode
            if hasattr(self.block.wiring_config, 'routing_mode'):
                self.routing_mode_var.set(self.block.wiring_config.routing_mode)
            
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

        # Draw trackers and devices
        self.draw_trackers()
        self.draw_device()
        self.draw_device_destination_points()

        # Draw routes based on current routing mode
        self.draw_current_routes()
                        
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
            
            # Get routes based on current routing mode
            cable_routes = self.get_current_routes()

            # Process each tracker position to capture collection points
            for idx, pos in enumerate(self.block.tracker_positions):
                if not pos.template:
                    continue
                    
                # Get whip points for this tracker
                tracker_idx = idx
                pos_whip = self.get_whip_position(str(tracker_idx), 'positive')
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
                        
            custom_whip_points = {}
            if (hasattr(self.block, 'wiring_config') and 
                self.block.wiring_config and 
                hasattr(self.block.wiring_config, 'custom_whip_points')):
                custom_whip_points = self.block.wiring_config.custom_whip_points

            # Get harness groupings from the existing configuration
            harness_groupings = {}
            if (hasattr(self.block, 'wiring_config') and 
                self.block.wiring_config and 
                hasattr(self.block.wiring_config, 'harness_groupings')):
                harness_groupings = self.block.wiring_config.harness_groupings

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
                custom_whip_points=custom_whip_points,
                harness_groupings=harness_groupings
                )
            
            # Store the routing mode for BOM warnings
            wiring_config.routing_mode = self.routing_mode_var.get()
            
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
            for whip_info in self.selected_whips:
                if len(whip_info) == 3:
                    tracker_id, harness_idx, polarity = whip_info
                    
                    # Handle harness-specific whip points
                    if harness_idx is not None:
                        # Ensure custom_harness_whip_points exists
                        if not hasattr(self.block, 'wiring_config'):
                            continue
                        if not self.block.wiring_config:
                            continue
                        if not hasattr(self.block.wiring_config, 'custom_harness_whip_points'):
                            self.block.wiring_config.custom_harness_whip_points = {}
                        
                        # Ensure tracker_id entry exists
                        if tracker_id not in self.block.wiring_config.custom_harness_whip_points:
                            self.block.wiring_config.custom_harness_whip_points[tracker_id] = {}
                        
                        # Ensure harness_idx entry exists
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
                    
                else:
                    # Handle legacy format (tracker_id, polarity)
                    tracker_id, polarity = whip_info
                    harness_idx = None
                    
                    # Update regular whip point
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

    def get_whip_position(self, tracker_id, polarity, harness_idx=None):
        """Get the current position of a whip point (custom or default)"""
        # In realistic mode, always use centerline positions
        if hasattr(self, 'routing_mode_var') and self.routing_mode_var.get() == "realistic":
            return self.get_realistic_whip_position(tracker_id, polarity, harness_idx)
        
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
            
        # Fall back to default position
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
        
        # Determine which end of the tracker is closer to the device
        device_y = self.block.device_y
        tracker_north_y = pos.y
        tracker_south_y = pos.y + tracker_length
        
        # Calculate distances to both ends
        distance_to_north = abs(device_y - tracker_north_y)
        distance_to_south = abs(device_y - tracker_south_y)
        
        # Place whip at the appropriate end
        whip_offset = 0.5  # Offset from tracker end in meters
        
        if distance_to_north < distance_to_south:
            # Place on north side (top of tracker)
            y_position = tracker_north_y - whip_offset
        else:
            # Place on south side (bottom of tracker)
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
            node_y = tracker_pos.y + (string.positive_source_y if is_positive else string.negative_source_y)
            
            return (node_x, node_y)
        else:
            # Conceptual: Use existing logic with horizontal offset and vertical spacing
            horizontal_offset = 0.6
            vertical_spacing = 0.3
            harness_offset_y = harness_idx * vertical_spacing
            
            if is_positive:
                node_x = tracker_pos.x + string.positive_source_x - horizontal_offset
                node_y = tracker_pos.y + string.positive_source_y + harness_offset_y
            else:
                node_x = tracker_pos.x + string.negative_source_x + horizontal_offset
                node_y = tracker_pos.y + string.negative_source_y + harness_offset_y
                
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
        
        # Get base whip position
        base_whip = self.get_whip_position(tracker_id, polarity)
        if not base_whip:
            return None
        
        if routing_mode == "realistic":
            # Realistic: All harness whips at same Y position (no vertical spacing)
            return base_whip
        else:
            # Conceptual: Add vertical spacing between harness whips
            vertical_spacing = 0.3
            return (base_whip[0], base_whip[1] + harness_idx * vertical_spacing)

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

    def update_ui_for_wiring_type(self):
        """Update UI elements based on selected wiring type"""
        is_harness = self.wiring_type_var.get() == WiringType.HARNESS.value
        if is_harness:
            self.harness_frame.grid()
            self.harness_config_frame.grid()
            
            # Make sure we have the quick patterns section
            if not hasattr(self, 'quick_patterns_frame'):
                self.setup_quick_patterns_ui()
                
            self.populate_string_count_combobox()
        else:
            self.harness_frame.grid_remove()
            self.harness_config_frame.grid_remove()
        self.draw_wiring_layout()

    def setup_quick_patterns_ui(self):
        """Create UI section for quick harness patterns"""
        # Add before the string grouping frame
        self.quick_patterns_frame = ttk.LabelFrame(self.harness_config_frame, text="Quick Pattern Presets")
        self.quick_patterns_frame.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Add preset buttons
        ttk.Button(self.quick_patterns_frame, text="Split Evenly in 2", 
                command=lambda: self.apply_quick_pattern("split_even_2")).grid(
                row=0, column=0, padx=5, pady=5)
                
        ttk.Button(self.quick_patterns_frame, text="Furthest String Separate", 
                command=lambda: self.apply_quick_pattern("furthest_separate")).grid(
                row=0, column=1, padx=5, pady=5)
                
        ttk.Button(self.quick_patterns_frame, text="Reset to Default", 
                command=lambda: self.apply_quick_pattern("default")).grid(
                row=0, column=2, padx=5, pady=5)
        
        # Move the existing string grouping frame to appear after quick patterns
        if hasattr(self, 'string_grouping_frame'):
            self.string_grouping_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Move the harness display frame to after string grouping
        if hasattr(self, 'harness_display_frame'):
            self.harness_display_frame.grid(row=3, column=0, columnspan=3, padx=5, pady=5, sticky=(tk.W, tk.E))

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
        
        # Create string checkboxes
        self.string_vars = []
        for i in range(string_count):
            var = tk.BooleanVar(value=False)
            self.string_vars.append(var)
            check = ttk.Checkbutton(self.string_check_frame, text=f"String {i+1}", variable=var)
            check.grid(row=i, column=0, sticky=tk.W, padx=5, pady=2)
        
        # Update harness display
        self.update_harness_display(string_count)

    def update_harness_display(self, string_count):
        """Update the display of current harnesses for the selected string count"""
        # Clear current display
        for widget in self.harness_display_frame.winfo_children():
            widget.destroy()
        
        # Check if harness groupings exist for this string count
        if not hasattr(self.block.wiring_config, 'harness_groupings') or \
        string_count not in self.block.wiring_config.harness_groupings or \
        not self.block.wiring_config.harness_groupings[string_count]:
            ttk.Label(self.harness_display_frame, text="Default configuration: all strings in one harness").grid(
                row=0, column=0, padx=5, pady=5)
            return
        
        # Display each harness
        for i, harness in enumerate(self.block.wiring_config.harness_groupings[string_count]):
            # Use a LabelFrame instead of a Frame to better separate harnesses
            harness_frame = ttk.LabelFrame(self.harness_display_frame, text=f"Harness {i+1}")
            harness_frame.grid(row=i, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)
            
            # First row: String info
            # Get string indices as a comma-separated list
            string_indices = [str(idx+1) for idx in harness.string_indices]
            strings_text = f"Strings: {', '.join(string_indices)}"
            ttk.Label(harness_frame, text=strings_text).grid(
                row=0, column=0, columnspan=2, sticky=tk.W, padx=5, pady=2)
            
            # Second row: Cable info
            ttk.Label(harness_frame, text=f"Cable: {harness.cable_size}").grid(
                row=1, column=0, sticky=tk.W, padx=5, pady=2)
            
            # Add fuse information if harness has multiple strings
            has_multiple_strings = len(harness.string_indices) > 1
            use_fuse = getattr(harness, 'use_fuse', has_multiple_strings)
            
            if has_multiple_strings and use_fuse:
                fuse_rating = getattr(harness, 'fuse_rating_amps', 15)
                fuse_text = f"Fuse: {fuse_rating}A ({len(harness.string_indices)} required)"
                ttk.Label(harness_frame, text=fuse_text).grid(
                    row=1, column=1, sticky=tk.W, padx=5, pady=2)
            elif has_multiple_strings:
                ttk.Label(harness_frame, text="No Fuses").grid(
                    row=1, column=1, sticky=tk.W, padx=5, pady=2)
            
            # Third row: Buttons - in their own frame
            button_frame = ttk.Frame(harness_frame)
            button_frame.grid(row=2, column=0, columnspan=2, pady=5)
            
            # Place Edit and Delete buttons side by side
            ttk.Button(button_frame, text="Edit", 
                    command=lambda h=i, sc=string_count: self.edit_harness(sc, h)).grid(
                    row=0, column=0, padx=5)
            
            ttk.Button(button_frame, text="Delete", 
                    command=lambda h=i, sc=string_count: self.delete_harness(sc, h)).grid(
                    row=0, column=1, padx=5)

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

        # Create new harness
        new_harness = HarnessGroup(
            string_indices=selected_indices,
            cable_size=self.harness_cable_size_var.get(),
            fuse_rating_amps=recommended_fuse,
            use_fuse=len(selected_indices) > 1  # Only use fuses for 2+ strings
        )
        
        # Add to harness groupings
        self.block.wiring_config.harness_groupings[string_count].append(new_harness)
        
        # Update display
        self.update_harness_display(string_count)
        
        # Update the wiring visualization
        self.draw_wiring_layout()

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

    def draw_device(self):
        """Draw inverter/combiner box"""
        if self.block.device_x is not None and self.block.device_y is not None:
            scale = self.get_canvas_scale()
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
    def draw_trackers(self):
        """Draw all trackers with modules and source points"""
        scale = self.get_canvas_scale()
        
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
            self.draw_whip_points(pos)

    def draw_whip_points(self, pos):
        """Draw whip points for a tracker"""
        scale = self.get_canvas_scale()
        tracker_idx = self.block.tracker_positions.index(pos)
        tracker_id = str(tracker_idx)
        string_count = len(pos.strings)
        
        # Draw regular tracker whip points
        pos_whip = self.get_whip_position(tracker_id, 'positive')
        neg_whip = self.get_whip_position(tracker_id, 'negative')
        
        if pos_whip:
            wx = 20 + self.pan_x + pos_whip[0] * scale
            wy = 20 + self.pan_y + pos_whip[1] * scale
            
            # Check if selected
            is_selected = (tracker_id, 'positive') in self.selected_whips
            fill_color = 'orange' if is_selected else 'pink'
            outline_color = 'red' if is_selected else 'deeppink'
            size = 5 if is_selected else 3
            
            self.canvas.create_oval(
                wx - size, wy - size,
                wx + size, wy + size,
                fill=fill_color, outline=outline_color,
                tags='whip_point'
            )
        
        if neg_whip:
            wx = 20 + self.pan_x + neg_whip[0] * scale
            wy = 20 + self.pan_y + neg_whip[1] * scale
            
            # Check if selected
            is_selected = (tracker_id, 'negative') in self.selected_whips
            fill_color = 'cyan' if is_selected else 'teal'
            outline_color = 'blue' if is_selected else 'darkcyan'
            size = 5 if is_selected else 3
            
            self.canvas.create_oval(
                wx - size, wy - size,
                wx + size, wy + size,
                fill=fill_color, outline=outline_color,
                tags='whip_point'
            )
        
        # Draw harness-specific whip points
        if (hasattr(self.block, 'wiring_config') and 
            self.block.wiring_config and 
            hasattr(self.block.wiring_config, 'harness_groupings') and
            string_count in self.block.wiring_config.harness_groupings):
            
            for harness_idx, _ in enumerate(self.block.wiring_config.harness_groupings[string_count]):
                harness_pos_whip = self.get_whip_position(tracker_id, 'positive', harness_idx)
                harness_neg_whip = self.get_whip_position(tracker_id, 'negative', harness_idx)
                
                if harness_pos_whip:
                    wx = 20 + self.pan_x + harness_pos_whip[0] * scale
                    wy = 20 + self.pan_y + harness_pos_whip[1] * scale
                    
                    # Check if selected
                    is_selected = (tracker_id, harness_idx, 'positive') in self.selected_whips
                    fill_color = 'orange' if is_selected else 'pink'
                    outline_color = 'red' if is_selected else 'deeppink'
                    size = 5 if is_selected else 3
                    
                    self.canvas.create_oval(
                        wx - size, wy - size,
                        wx + size, wy + size,
                        fill=fill_color, outline=outline_color,
                        tags=f'harness_whip_point_{harness_idx}_positive'
                    )
                
                if harness_neg_whip:
                    wx = 20 + self.pan_x + harness_neg_whip[0] * scale
                    wy = 20 + self.pan_y + harness_neg_whip[1] * scale
                    
                    # Check if selected
                    is_selected = (tracker_id, harness_idx, 'negative') in self.selected_whips
                    fill_color = 'cyan' if is_selected else 'teal'
                    outline_color = 'blue' if is_selected else 'darkcyan'
                    size = 5 if is_selected else 3
                    
                    self.canvas.create_oval(
                        wx - size, wy - size,
                        wx + size, wy + size,
                        fill=fill_color, outline=outline_color,
                        tags=f'harness_whip_point_{harness_idx}_negative'
                    )

    def draw_string_homerun_wiring(self):
        """Draw string homerun wiring configuration"""
        scale = self.get_canvas_scale()
        
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
        """Draw default harness configuration (all strings in one harness)"""
        scale = self.get_canvas_scale()
        
        # Calculate node points using position helpers
        routing_mode = self.routing_mode_var.get()
        pos_nodes = []
        neg_nodes = []

        for i, string in enumerate(pos.strings):
            pos_node = self.get_harness_collection_point(pos, string, 0, True, routing_mode)  # harness_idx=0 for default
            neg_node = self.get_harness_collection_point(pos, string, 0, False, routing_mode)
            pos_nodes.append(pos_node)
            neg_nodes.append(neg_node)

        # Draw string cables to node points
        for i, string in enumerate(pos.strings):
            # Positive string cable
            source_x = pos.x + string.positive_source_x
            source_y = pos.y + string.positive_source_y
            route = self.get_string_source_to_harness_route(
                (source_x, source_y), 
                pos_nodes[i], 
                routing_mode
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
            route = self.get_string_source_to_harness_route(
                (source_x, source_y), 
                neg_nodes[i], 
                routing_mode
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
        
        # Determine device position relative to tracker
        device_is_north = self.block.device_y < pos.y

        # Draw node connections - order depends on device position
        # Positive harness
        if device_is_north:
            # Device north of tracker: connect from south to north (current logic)
            sorted_pos_nodes = sorted(pos_nodes, key=lambda p: p[1], reverse=True)
        else:
            # Device south of tracker: connect from north to south (reversed)
            sorted_pos_nodes = sorted(pos_nodes, key=lambda p: p[1], reverse=False)

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
        if device_is_north:
            # Device north of tracker: connect from south to north (current logic)
            sorted_neg_nodes = sorted(neg_nodes, key=lambda p: p[1], reverse=True)
        else:
            # Device south of tracker: connect from north to south (reversed)
            sorted_neg_nodes = sorted(neg_nodes, key=lambda p: p[1], reverse=False)

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

        # Get device destination points
        pos_dest, neg_dest = self.get_device_destination_points()
        
        # Draw whip-to-device connections for default case
        if pos_whip and pos_dest:
            route = self.calculate_cable_route(
                pos_whip[0], pos_whip[1],
                pos_dest[0], pos_dest[1],
                True, 0
            )
            points = [(20 + self.pan_x + x * scale, 
                    20 + self.pan_y + y * scale) for x, y in route]
            if len(points) > 1:
                # All strings are combined into one harness
                num_strings = len(pos.strings)
                current = self.calculate_current_for_segment('whip', num_strings)
                self.draw_wire_segment(points, self.whip_cable_size_var.get(), 
                                    current, is_positive=True, segment_type="whip")
                self.add_current_label(points, current, is_positive=True, segment_type='whip')
        
        # Same for negative
        if neg_whip and neg_dest:
            route = self.calculate_cable_route(
                neg_whip[0], neg_whip[1],
                neg_dest[0], neg_dest[1],
                False, 0
            )
            points = [(20 + self.pan_x + x * scale, 
                    20 + self.pan_y + y * scale) for x, y in route]
            if len(points) > 1:
                num_strings = len(pos.strings)
                current = self.calculate_current_for_segment('whip', num_strings)
                self.draw_wire_segment(points, self.whip_cable_size_var.get(), 
                                    current, is_positive=False, segment_type="whip")
                self.add_current_label(points, current, is_positive=False, segment_type='whip')

    def draw_custom_harnesses(self, pos, pos_whip, neg_whip, string_count):
        """Draw custom harness configurations"""
        scale = self.get_canvas_scale()
        tracker_idx = self.block.tracker_positions.index(pos)
        tracker_id = str(tracker_idx)
        
        # Process each harness group
        for harness_idx, harness in enumerate(self.block.wiring_config.harness_groupings[string_count]):
            string_indices = harness.string_indices
            if not string_indices:
                continue
            
            # Calculate node positions for this harness group
            pos_nodes = []
            neg_nodes = []
            routing_mode = self.routing_mode_var.get()

            for string_idx in string_indices:
                if string_idx < len(pos.strings):
                    string = pos.strings[string_idx]
                    
                    # Use position helpers
                    pos_node = self.get_harness_collection_point(pos, string, harness_idx, True, routing_mode)
                    neg_node = self.get_harness_collection_point(pos, string, harness_idx, False, routing_mode)
                    
                    pos_nodes.append(pos_node)
                    neg_nodes.append(neg_node)
            
            # Draw string connections to nodes
            for i, string_idx in enumerate(string_indices):
                if string_idx < len(pos.strings):
                    string = pos.strings[string_idx]
                    
                    # Positive string cable
                    source_x = pos.x + string.positive_source_x
                    source_y = pos.y + string.positive_source_y
                    route = self.get_string_source_to_harness_route(
                        (source_x, source_y), 
                        pos_nodes[i], 
                        routing_mode
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
                    route = self.get_string_source_to_harness_route(
                        (source_x, source_y), 
                        neg_nodes[i], 
                        routing_mode
                    )
                    points = [(20 + self.pan_x + x * scale, 
                            20 + self.pan_y + y * scale) for x, y in route]
                    if len(points) > 1:
                        current = self.calculate_current_for_segment('string')
                        self.draw_wire_segment(points, self.string_cable_size_var.get(), 
                                            current, is_positive=False, segment_type="string")
                        self.add_current_label(points, current, is_positive=False)
            
            # Draw node points for this harness
            for nx, ny in pos_nodes:
                x = 20 + self.pan_x + nx * scale
                y = 20 + self.pan_y + ny * scale
                self.canvas.create_oval(x-3, y-3, x+3, y+3, fill='red', outline='darkred')
            
            for nx, ny in neg_nodes:
                x = 20 + self.pan_x + nx * scale
                y = 20 + self.pan_y + ny * scale
                self.canvas.create_oval(x-3, y-3, x+3, y+3, fill='blue', outline='darkblue')
            
            # Get harness-specific whip points using helper
            harness_pos_whip = self.get_harness_whip_point(tracker_id, harness_idx, 'positive', routing_mode)
            harness_neg_whip = self.get_harness_whip_point(tracker_id, harness_idx, 'negative', routing_mode)
            
            # Determine device position relative to tracker
            device_is_north = self.block.device_y < pos.y

            # Draw harness connections
            if pos_nodes:
                # Sort nodes by y-coordinate based on device position
                if device_is_north:
                    # Device north of tracker: connect from south to north
                    sorted_pos_nodes = sorted(pos_nodes, key=lambda p: p[1], reverse=True)
                else:
                    # Device south of tracker: connect from north to south
                    sorted_pos_nodes = sorted(pos_nodes, key=lambda p: p[1], reverse=False)
                
                # Connect nodes in sequence
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
                        # Last node routes to harness whip point
                        if harness_pos_whip:
                            route = self.calculate_cable_route(
                                start[0], start[1],
                                harness_pos_whip[0], harness_pos_whip[1],
                                True, 0
                            )
                        else:
                            continue
                    
                    points = [(20 + self.pan_x + x * scale, 
                            20 + self.pan_y + y * scale) for x, y in route]
                    if len(points) > 1:
                        # Calculate current for this harness
                        strings_in_harness = len(string_indices)
                        current = self.calculate_current_for_segment('string') * strings_in_harness
                        
                        # Draw with the harness's cable size if specified, else use default
                        cable_size = harness.cable_size if hasattr(harness, 'cable_size') else self.harness_cable_size_var.get()
                        
                        self.draw_wire_segment(points, cable_size, 
                                            current, is_positive=True, segment_type="harness")
                        
                        # Add current label
                        if len(points) >= 2 and self.show_current_labels_var.get():
                            mid_idx = len(points) // 2
                            mid_x = (points[mid_idx-1][0] + points[mid_idx][0]) / 2
                            mid_y = (points[mid_idx-1][1] + points[mid_idx][1]) / 2
                            self.canvas.create_text(mid_x, mid_y-10, text=f"{current:.1f}A", 
                                                fill='red', font=('Arial', 8))
            
            if neg_nodes:
                # Same for negative nodes
                if device_is_north:
                    # Device north of tracker: connect from south to north
                    sorted_neg_nodes = sorted(neg_nodes, key=lambda p: p[1], reverse=True)
                else:
                    # Device south of tracker: connect from north to south
                    sorted_neg_nodes = sorted(neg_nodes, key=lambda p: p[1], reverse=False)
                
                for i in range(len(sorted_neg_nodes)):
                    start = sorted_neg_nodes[i]
                    if i < len(sorted_neg_nodes) - 1:
                        end = sorted_neg_nodes[i + 1]
                        route = self.calculate_cable_route(
                            start[0], start[1],
                            end[0], end[1],
                            False, 0
                        )
                    else:
                        if harness_neg_whip:
                            route = self.calculate_cable_route(
                                start[0], start[1],
                                harness_neg_whip[0], harness_neg_whip[1],
                                False, 0
                            )
                        else:
                            continue
                            
                    points = [(20 + self.pan_x + x * scale, 
                            20 + self.pan_y + y * scale) for x, y in route]
                    if len(points) > 1:
                        strings_in_harness = len(string_indices)
                        current = self.calculate_current_for_segment('string') * strings_in_harness
                        
                        cable_size = harness.cable_size if hasattr(harness, 'cable_size') else self.harness_cable_size_var.get()
                        
                        self.draw_wire_segment(points, cable_size, 
                                            current, is_positive=False, segment_type="harness")
                        
                        if len(points) >= 2 and self.show_current_labels_var.get():
                            mid_idx = len(points) // 2
                            mid_x = (points[mid_idx-1][0] + points[mid_idx][0]) / 2
                            mid_y = (points[mid_idx-1][1] + points[mid_idx][1]) / 2
                            self.canvas.create_text(mid_x, mid_y+10, text=f"{current:.1f}A", 
                                                fill='blue', font=('Arial', 8))
            
            # Draw harness whip points
            if harness_pos_whip:
                wx = 20 + self.pan_x + harness_pos_whip[0] * scale
                wy = 20 + self.pan_y + harness_pos_whip[1] * scale
                
                # Check if this harness whip point is selected
                is_selected = (tracker_id, harness_idx, 'positive') in self.selected_whips
                fill_color = 'orange' if is_selected else 'pink'
                outline_color = 'red' if is_selected else 'deeppink'
                size = 5 if is_selected else 3
                
                self.canvas.create_oval(
                    wx - size, wy - size,
                    wx + size, wy + size,
                    fill=fill_color, 
                    outline=outline_color,
                    tags=f'harness_whip_point_{harness_idx}_positive'
                )
            
            if harness_neg_whip:
                wx = 20 + self.pan_x + harness_neg_whip[0] * scale
                wy = 20 + self.pan_y + harness_neg_whip[1] * scale
                
                # Check if this harness whip point is selected
                is_selected = (tracker_id, harness_idx, 'negative') in self.selected_whips
                fill_color = 'cyan' if is_selected else 'teal'
                outline_color = 'blue' if is_selected else 'darkcyan'
                size = 5 if is_selected else 3
                
                self.canvas.create_oval(
                    wx - size, wy - size,
                    wx + size, wy + size,
                    fill=fill_color, 
                    outline=outline_color,
                    tags=f'harness_whip_point_{harness_idx}_negative'
                )
            
            # Route from harness whip points to device
            pos_dest, neg_dest = self.get_device_destination_points()
            
            # Draw positive whip to device route
            if harness_pos_whip and pos_dest:
                self.draw_whip_to_device_connection(
                    harness_pos_whip, pos_dest, 
                    len(string_indices), True, harness_idx
                )
            
            # Draw negative whip to device route
            if harness_neg_whip and neg_dest:
                self.draw_whip_to_device_connection(
                    harness_neg_whip, neg_dest, 
                    len(string_indices), False, harness_idx
                )

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
        for i, string_idx in enumerate(unconfigured_indices):
            if string_idx < len(pos.strings):
                string = pos.strings[string_idx]
                
                # Positive string cable
                source_x = pos.x + string.positive_source_x
                source_y = pos.y + string.positive_source_y
                route = self.get_string_source_to_harness_route(
                    (source_x, source_y), 
                    pos_nodes[i], 
                    routing_mode
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
                route = self.get_string_source_to_harness_route(
                    (source_x, source_y), 
                    neg_nodes[i], 
                    routing_mode
                )
                points = [(20 + self.pan_x + x * scale, 
                        20 + self.pan_y + y * scale) for x, y in route]
                if len(points) > 1:
                    current = self.calculate_current_for_segment('string')
                    self.draw_wire_segment(points, self.string_cable_size_var.get(), 
                                        current, is_positive=False, segment_type="string")
                    self.add_current_label(points, current, is_positive=False)
        
        # Draw node points for this harness
        for nx, ny in pos_nodes:
            x = 20 + self.pan_x + nx * scale
            y = 20 + self.pan_y + ny * scale
            self.canvas.create_oval(x-3, y-3, x+3, y+3, fill='red', outline='darkred')
        
        for nx, ny in neg_nodes:
            x = 20 + self.pan_x + nx * scale
            y = 20 + self.pan_y + ny * scale
            self.canvas.create_oval(x-3, y-3, x+3, y+3, fill='blue', outline='darkblue')
        
        # Generate harness-specific whip points for unconfigured strings
        harness_pos_whip = None
        harness_neg_whip = None
        
        # Get harness-specific whip points using helper
        tracker_id = str(self.block.tracker_positions.index(pos))
        harness_pos_whip = self.get_harness_whip_point(tracker_id, harness_idx, 'positive', routing_mode)
        harness_neg_whip = self.get_harness_whip_point(tracker_id, harness_idx, 'negative', routing_mode)
        
        # Draw harness connections (simplified for unconfigured)
        if pos_nodes:
            # Just connect all nodes directly to whip point
            for i, node_pos in enumerate(pos_nodes):
                if harness_pos_whip:
                    route = self.calculate_cable_route(
                        node_pos[0], node_pos[1],
                        harness_pos_whip[0], harness_pos_whip[1],
                        True, i
                    )
                    points = [(20 + self.pan_x + x * scale, 
                            20 + self.pan_y + y * scale) for x, y in route]
                    if len(points) > 1:
                        current = self.calculate_current_for_segment('string')
                        self.draw_wire_segment(points, self.harness_cable_size_var.get(), 
                                            current, is_positive=True, segment_type="harness")
        
        if neg_nodes:
            # Same for negative
            for i, node_pos in enumerate(neg_nodes):
                if harness_neg_whip:
                    route = self.calculate_cable_route(
                        node_pos[0], node_pos[1],
                        harness_neg_whip[0], harness_neg_whip[1],
                        False, i
                    )
                    points = [(20 + self.pan_x + x * scale, 
                            20 + self.pan_y + y * scale) for x, y in route]
                    if len(points) > 1:
                        current = self.calculate_current_for_segment('string')
                        self.draw_wire_segment(points, self.harness_cable_size_var.get(), 
                                            current, is_positive=False, segment_type="harness")
        
        # Draw harness whip points
        if harness_pos_whip:
            wx = 20 + self.pan_x + harness_pos_whip[0] * scale
            wy = 20 + self.pan_y + harness_pos_whip[1] * scale
            self.canvas.create_oval(
                wx - 3, wy - 3,
                wx + 3, wy + 3,
                fill='pink', outline='deeppink',
                tags=f'auto_harness_pos_whip'
            )
        
        if harness_neg_whip:
            wx = 20 + self.pan_x + harness_neg_whip[0] * scale
            wy = 20 + self.pan_y + harness_neg_whip[1] * scale
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
                len(unconfigured_indices), True, harness_idx
            )
        
        if harness_neg_whip and neg_dest:
            self.draw_whip_to_device_connection(
                harness_neg_whip, neg_dest, 
                len(unconfigured_indices), False, harness_idx
            )

    def draw_whip_to_device_connection(self, whip_point, device_point, num_strings, is_positive, harness_idx=0):
        """Draw a connection from a whip point to the device"""
        scale = self.get_canvas_scale()
        
        route = self.calculate_cable_route(
            whip_point[0], whip_point[1],
            device_point[0], device_point[1],
            is_positive, harness_idx
        )
        
        points = [(20 + self.pan_x + x * scale, 
                20 + self.pan_y + y * scale) for x, y in route]
        
        if len(points) > 1:
            current = self.calculate_current_for_segment('whip', num_strings)
            
            self.draw_wire_segment(points, self.whip_cable_size_var.get(), 
                                current, is_positive=is_positive, segment_type="whip")
            
            if len(points) >= 2 and self.show_current_labels_var.get():
                mid_idx = len(points) // 2
                mid_x = (points[mid_idx-1][0] + points[mid_idx][0]) / 2
                mid_y = (points[mid_idx-1][1] + points[mid_idx][1]) / 2
                offset = -10 if is_positive else 10
                color = 'red' if is_positive else 'blue'
                self.canvas.create_text(mid_x, mid_y + offset, 
                                    text=f"{current:.1f}A", 
                                    fill=color, font=('Arial', 8))
                
    def apply_quick_pattern(self, pattern_type):
        """Apply a predefined harness pattern"""
        if not self.string_count_var.get() or not hasattr(self, 'string_count_mapping'):
            return
        
        # Get selected string count
        selected_item = self.string_count_var.get()
        if selected_item not in self.string_count_mapping:
            return
        
        string_count = self.string_count_mapping[selected_item]
        
        # Initialize harness_groupings if needed
        if not hasattr(self.block.wiring_config, 'harness_groupings'):
            self.block.wiring_config.harness_groupings = {}
        
        # Reset existing harness configuration for this string count
        self.block.wiring_config.harness_groupings[string_count] = []
        
        # Create harnesses based on selected pattern
        if pattern_type == "split_even_2":
            # Split strings evenly into 2 harnesses
            half_point = string_count // 2
            
            # First half
            harness1 = HarnessGroup(
                string_indices=list(range(0, half_point)),
                cable_size=self.harness_cable_size_var.get()
            )
            self.block.wiring_config.harness_groupings[string_count].append(harness1)
            
            # Second half
            harness2 = HarnessGroup(
                string_indices=list(range(half_point, string_count)),
                cable_size=self.harness_cable_size_var.get()
            )
            self.block.wiring_config.harness_groupings[string_count].append(harness2)
        
        elif pattern_type == "furthest_separate":
            # Determine which string is furthest from the device
            pos_dest, neg_dest = self.get_device_destination_points()
            if not pos_dest or not self.block.tracker_positions:
                return
                
            # For each string in the selected tracker, calculate its distance from the device
            furthest_string_idx = 0
            max_distance = 0
            
            # Get the tracker position for this string count
            for pos_idx, pos in enumerate(self.block.tracker_positions):
                if len(pos.strings) != string_count:
                    continue
                    
                # For each string in this tracker
                for i, string in enumerate(pos.strings):
                    # Calculate distance from string source to device
                    source_x = pos.x + string.positive_source_x
                    source_y = pos.y + string.positive_source_y
                    
                    # Pythagorean distance
                    distance = ((source_x - pos_dest[0])**2 + (source_y - pos_dest[1])**2)**0.5
                    
                    if distance > max_distance:
                        max_distance = distance
                        furthest_string_idx = i
            
            # Create two harnesses: one for the furthest string, one for all others
            harness1 = HarnessGroup(
                string_indices=[furthest_string_idx],
                cable_size=self.harness_cable_size_var.get()
            )
            self.block.wiring_config.harness_groupings[string_count].append(harness1)
            
            other_indices = [i for i in range(string_count) if i != furthest_string_idx]
            if other_indices:
                harness2 = HarnessGroup(
                    string_indices=other_indices,
                    cable_size=self.harness_cable_size_var.get()
                )
                self.block.wiring_config.harness_groupings[string_count].append(harness2)
        
        elif pattern_type == "default":
            # Remove all custom harness definitions - will revert to default
            if string_count in self.block.wiring_config.harness_groupings:
                del self.block.wiring_config.harness_groupings[string_count]
        
        # Update the display and redraw
        self.update_harness_display(string_count)
        self.draw_wiring_layout()

    def edit_harness(self, string_count, harness_idx):
        """Edit harness properties including fuse configuration"""
        if string_count not in self.block.wiring_config.harness_groupings or \
        harness_idx >= len(self.block.wiring_config.harness_groupings[string_count]):
            return
            
        harness = self.block.wiring_config.harness_groupings[string_count][harness_idx]
        
        # Create dialog window
        dialog = tk.Toplevel(self)
        dialog.title(f"Edit Harness {harness_idx + 1}")
        dialog.transient(self)
        dialog.grab_set()
        
        # Center dialog
        x = self.winfo_rootx() + 50
        y = self.winfo_rooty() + 50
        dialog.geometry(f"+{x}+{y}")
        
        dialog_frame = ttk.Frame(dialog, padding="10")
        dialog_frame.pack(fill=tk.BOTH, expand=True)
        
        # Cable size
        ttk.Label(dialog_frame, text="Cable Size:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        cable_size_var = tk.StringVar(value=harness.cable_size)
        cable_combo = ttk.Combobox(dialog_frame, textvariable=cable_size_var, state='readonly')
        cable_combo['values'] = list(self.AWG_SIZES.keys())
        cable_combo.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Use fuse option - only for harnesses with multiple strings
        use_fuse_frame = ttk.Frame(dialog_frame)
        use_fuse_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
        
        # Default use_fuse to True for harnesses with multiple strings
        has_multiple_strings = len(harness.string_indices) > 1
        default_use_fuse = getattr(harness, 'use_fuse', has_multiple_strings)
        
        use_fuse_var = tk.BooleanVar(value=default_use_fuse)
        fuse_check = ttk.Checkbutton(use_fuse_frame, text="Use Fuse Protection", 
                    variable=use_fuse_var)
        fuse_check.grid(row=0, column=0, sticky=tk.W)
        
        # Disable checkbox for single-string harnesses
        if not has_multiple_strings:
            fuse_check.configure(state='disabled')
            use_fuse_var.set(False)
        
        # Fuse rating - only enabled if use_fuse is True
        fuse_frame = ttk.LabelFrame(dialog_frame, text="Fuse Configuration")
        fuse_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        ttk.Label(fuse_frame, text="Fuse Rating (A):").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        # Calculate recommended fuse size based on NEC (1.25 × Isc)
        recommended_fuse = self.calculate_recommended_fuse_size(harness.string_indices)
        
        # Use existing fuse rating or recommended
        current_rating = getattr(harness, 'fuse_rating_amps', recommended_fuse)
        fuse_rating_var = tk.StringVar(value=str(current_rating))
        
        fuse_combo = ttk.Combobox(fuse_frame, textvariable=fuse_rating_var, state='readonly')
        fuse_combo['values'] = [str(r) for r in self.FUSE_RATINGS if r >= recommended_fuse]
        fuse_combo.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Recommendation label
        recommendation_label = ttk.Label(fuse_frame, 
                                        text=f"NEC minimum: {recommended_fuse}A (1.25 × Isc)")
        recommendation_label.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
        
        # Add fuse quantity info - one per string for harnesses with multiple strings
        if has_multiple_strings:
            quantity_label = ttk.Label(fuse_frame, 
                                    text=f"Fuses required: {len(harness.string_indices)} (one per string)")
            quantity_label.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
        
        # Update fuse frame state based on use_fuse checkbox
        def update_fuse_frame_state():
            state = 'normal' if use_fuse_var.get() and has_multiple_strings else 'disabled'
            for child in fuse_frame.winfo_children():
                if child.winfo_class() != 'TLabel':  # Don't disable labels
                    child.configure(state=state)
        
        # Call initially and add trace
        update_fuse_frame_state()
        use_fuse_var.trace('w', lambda *args: update_fuse_frame_state())
        
        # Button frame
        button_frame = ttk.Frame(dialog_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=10)
        
        # Save button
        def save_changes():
            # Update harness properties
            harness.cable_size = cable_size_var.get()
            harness.use_fuse = use_fuse_var.get() and has_multiple_strings
            try:
                harness.fuse_rating_amps = int(fuse_rating_var.get())
            except ValueError:
                harness.fuse_rating_amps = recommended_fuse
            
            # Update the display
            self.update_harness_display(string_count)
            dialog.destroy()
        
        ttk.Button(button_frame, text="Save", command=save_changes).grid(
            row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).grid(
            row=0, column=1, padx=5)
        
    def calculate_recommended_fuse_size(self, string_indices):
        """Calculate recommended fuse size based on NEC (1.25 × Isc)"""
        if not self.block.tracker_template or not self.block.tracker_template.module_spec:
            return 15  # Default if no module spec available
        
        # Get Isc from module spec
        isc = self.block.tracker_template.module_spec.isc
        
        # Calculate NEC minimum (1.25 × Isc)
        nec_min = isc * 1.25
        
        # Round up to the nearest standard fuse size
        for rating in self.FUSE_RATINGS:
            if rating >= nec_min:
                return rating
        
        # If it's larger than our highest rating, return the highest
        return self.FUSE_RATINGS[-1]

    def get_current_routes(self):
        """Get routes based on current routing mode (realistic or conceptual)"""
        if self.routing_mode_var.get() == "realistic":
            # For BOM calculation, we need both positive and negative routes
            # even though we might show simplified visuals
            return self.calculate_realistic_routes_for_bom()
        else:
            # Use existing conceptual routing logic
            cable_routes = {}
            
            for idx, pos in enumerate(self.block.tracker_positions):
                if not pos.template:
                    continue
                    
                tracker_idx = idx
                pos_whip = self.get_whip_position(str(tracker_idx), 'positive')
                neg_whip = self.get_whip_position(str(tracker_idx), 'negative')
                
                for string_idx, string in enumerate(pos.strings):
                    if self.wiring_type_var.get() == WiringType.HOMERUN.value:
                        self.add_homerun_routes(cable_routes, pos, string, tracker_idx, string_idx, pos_whip, neg_whip)
                    else:
                        self.add_harness_routes(cable_routes, pos, string, tracker_idx, string_idx, pos_whip, neg_whip)
            
            return cable_routes
        
    def draw_current_routes(self):
        """Draw routes based on current wiring type (always uses conceptual logic with positioning variants)"""
        if self.wiring_type_var.get() == WiringType.HOMERUN.value:
            self.draw_string_homerun_wiring()
        else:  # Wire Harness configuration
            self.draw_wire_harness_wiring()

    def add_current_label_to_route(self, canvas_points, current, is_positive, segment_type):
        """Add current label to a route"""
        if len(canvas_points) < 2:
            return
            
        # Find midpoint
        mid_idx = len(canvas_points) // 2
        mid_x, mid_y = canvas_points[mid_idx]
        
        # Adjust label position
        offset = -8 if is_positive else 8
        color = 'red' if is_positive else 'blue'
        
        self.canvas.create_text(mid_x, mid_y + offset, text=f"{current:.1f}A", 
                            fill=color, font=('Arial', 8))

    def validate_and_show_current_warnings(self):
        """Validate current loads and show warnings for both routing modes"""
        # Clear previous warnings
        self.clear_warnings()
        
        # Get routes for validation
        routes = self.get_current_routes()
        
        # Validate each route for overloads
        for route_id, route_points in routes.items():
            if len(route_points) < 2:
                continue
                
            # Determine wire properties and current
            is_positive = 'pos_' in route_id
            if 'src_' in route_id:
                wire_gauge = self.string_cable_size_var.get()
                current = self.calculate_current_for_segment('string')
            elif 'harness_' in route_id or 'main_' in route_id:
                wire_gauge = self.harness_cable_size_var.get()
                num_strings = self.get_harness_string_count(route_id)
                current = self.calculate_current_for_segment('string') * num_strings
            else:  # whip routes
                wire_gauge = self.whip_cable_size_var.get()
                num_strings = self.get_whip_string_count(route_id)
                current = self.calculate_current_for_segment('whip', num_strings)
            
            # Check for overloads
            from ..utils.calculations import get_ampacity_for_wire_gauge, calculate_nec_current
            ampacity = get_ampacity_for_wire_gauge(wire_gauge)
            if ampacity > 0:
                nec_current = calculate_nec_current(current)
                load_percent = (nec_current / ampacity) * 100
                
                if load_percent > 100:
                    polarity = "positive" if is_positive else "negative"
                    segment_type = "string" if 'src_' in route_id else ("harness" if 'harness_' in route_id or 'main_' in route_id else "whip")
                    
                    self.add_wire_warning(
                        route_id,
                        f"{polarity.capitalize()} {segment_type} {wire_gauge}: {load_percent:.0f}% (OVERLOAD)",
                        'overload'
                    )

    def get_harness_string_count(self, route_id):
        """Get number of strings in a harness based on route ID"""
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

    def get_whip_string_count(self, route_id):
        """Get number of strings for a whip route"""
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
    