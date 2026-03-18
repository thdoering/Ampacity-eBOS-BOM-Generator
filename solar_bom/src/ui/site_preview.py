import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import math
import copy


class SitePreviewWindow(tk.Toplevel):
    """Pop-out window for site layout preview with zoom and pan"""
    
    def __init__(self, parent, inv_summary, topology, colors, groups, enabled_templates, row_spacing_ft,
                 num_devices=0, device_label='CB', initial_inspect=False, pads=None, device_names=None):
        super().__init__(parent)
        self.title("Site Preview — Inverter Allocation")
        self.geometry("1100x750")
        self.minsize(600, 400)
        
        self.inv_summary = inv_summary
        self.topology = topology
        self.colors = colors
        self.groups = groups or []
        self.enabled_templates = enabled_templates or {}
        self.row_spacing_ft = row_spacing_ft
        self.num_devices = num_devices
        self.device_label = device_label
        self.inspect_mode = initial_inspect
        self.selected_device_idx = None
        self.selected_pad_inspect_idx = None  # Pad selected in inspect mode
        self.pads = list(pads) if pads else []  # Deep copy so we don't mutate caller's list
        self.device_names = dict(device_names) if device_names else {}  # {device_idx: "custom_name"}
        self.selected_pad_idx = None
        self.placing_pad = False  # True when in "click to place" mode
        self.assigning_devices = False  # True when Assign Devices dialog is open
        
        # Zoom and pan state
        self.scale = 1.0
        self.pan_x = 0
        self.pan_y = 0
        self.dragging_canvas = False
        self.dragging_group = False
        self.drag_start_x = 0
        self.drag_start_y = 0
        self._drag_moved = False
        self._drag_group_start_x = 0
        self._drag_group_start_y = 0
        self.selected_group_idx = None
        self.align_on_motor = True
        
        # Read lock state from parent (QuickEstimate)
        self.allocation_locked = getattr(self.master, 'allocation_locked', False)
        
        self.setup_ui()
        self.build_layout_data()
        self._recolor_from_cb_assignments()
        self.after(50, self.fit_and_redraw)
    
    def _get_preview_tracker_dims_ft(self, template_ref):
        """Compute physical (width_ft, length_ft) for a tracker from its template.
        
        Width = E-W dimension (across tracker, short side)
        Length = N-S dimension (along tracker, long side)
        
        Returns (width_ft, length_ft) or None if template not found.
        """
        if not template_ref or template_ref not in self.enabled_templates:
            return None
        
        tdata = self.enabled_templates[template_ref]
        module_spec = tdata.get('module_spec', {})
        
        module_length_mm = module_spec.get('length_mm', 2000)
        module_width_mm = module_spec.get('width_mm', 1000)
        orientation = tdata.get('module_orientation', 'Portrait')
        modules_per_string = tdata.get('modules_per_string', 28)
        strings_per_tracker = tdata.get('strings_per_tracker', 2)
        modules_high = tdata.get('modules_high', 1)
        module_spacing_m = tdata.get('module_spacing_m', 0.02)
        has_motor = tdata.get('has_motor', True)
        motor_gap_m = tdata.get('motor_gap_m', 1.0) if has_motor else 0
        
        # Module dimensions along vs across the tracker
        if orientation == 'Portrait':
            # Portrait: module width runs N-S (along tracker), length runs E-W (across)
            mod_along_m = module_width_mm / 1000
            mod_across_m = module_length_mm / 1000
        else:
            # Landscape: module length runs N-S (along tracker), width runs E-W (across)
            mod_along_m = module_length_mm / 1000
            mod_across_m = module_width_mm / 1000
        
        # N-S length: all modules in one string laid end-to-end, times strings, plus gaps and motor
        full_spt = int(strings_per_tracker)
        partial_mods = round((strings_per_tracker - full_spt) * modules_per_string) if strings_per_tracker != full_spt else 0
        modules_in_row = full_spt * modules_per_string + partial_mods
        tracker_length_m = (modules_in_row * mod_along_m + 
                           (modules_in_row - 1) * module_spacing_m +
                           motor_gap_m)
        
        # E-W width: module across dimension times modules_high
        tracker_width_m = mod_across_m * modules_high
        
        # Convert to feet
        m_to_ft = 3.28084

        return (tracker_width_m * m_to_ft, tracker_length_m * m_to_ft)

    def get_motor_position_in_tracker(self, template_ref):
        """Compute the motor's Y offset from the tracker top (north end), in feet.
        
        Returns (motor_y_offset_ft, motor_gap_ft, has_motor) or (0, 0, False).
        """
        if not template_ref or template_ref not in self.enabled_templates:
            return 0, 0, False
        
        tdata = self.enabled_templates[template_ref]
        has_motor = tdata.get('has_motor', True)
        if not has_motor:
            return 0, 0, False
        
        module_spec = tdata.get('module_spec', {})
        module_length_mm = module_spec.get('length_mm', 2000)
        module_width_mm = module_spec.get('width_mm', 1000)
        orientation = tdata.get('module_orientation', 'Portrait')
        modules_per_string = tdata.get('modules_per_string', 28)
        strings_per_tracker = tdata.get('strings_per_tracker', 2)
        module_spacing_m = tdata.get('module_spacing_m', 0.02)
        motor_gap_m = tdata.get('motor_gap_m', 1.0)
        motor_placement = tdata.get('motor_placement_type', 'between_strings')
        motor_position_after_string = tdata.get('motor_position_after_string', None)
        motor_string_index_raw = tdata.get('motor_string_index', None)
        motor_split_north = tdata.get('motor_split_north', modules_per_string // 2)
        
        if orientation == 'Portrait':
            mod_along_m = module_width_mm / 1000
        else:
            mod_along_m = module_length_mm / 1000
        
        m_to_ft = 3.28084
        
        # Partial string on north pushes motor further south
        partial_north_m = 0
        spt_val = tdata.get('strings_per_tracker', 1)
        if spt_val != int(spt_val) and tdata.get('partial_string_side', 'north') == 'north':
            partial_north_mods = round((spt_val - int(spt_val)) * modules_per_string)
            partial_north_m = partial_north_mods * (mod_along_m + module_spacing_m)
        
        if motor_placement == 'between_strings':
            pos_after = motor_position_after_string if motor_position_after_string is not None else (motor_string_index_raw if motor_string_index_raw is not None else 1)
            modules_north = pos_after * modules_per_string
            if modules_north > 0:
                motor_y_m = partial_north_m + (modules_north * mod_along_m + 
                            (modules_north - 1) * module_spacing_m +
                            module_spacing_m)
            else:
                motor_y_m = partial_north_m
        elif motor_placement == 'middle_of_string':
            string_idx = motor_string_index_raw if motor_string_index_raw is not None else 1
            modules_before_split = (string_idx - 1) * modules_per_string + motor_split_north
            motor_y_m = partial_north_m + (modules_before_split * mod_along_m + 
                        (modules_before_split - 1) * module_spacing_m +
                        module_spacing_m)
        else:
            # Fallback: center
            dims = self.get_tracker_dimensions_ft(template_ref)
            if dims:
                return dims[1] / 2, motor_gap_m * m_to_ft, True
            return 0, 0, False
        
        return motor_y_m * m_to_ft, motor_gap_m * m_to_ft, True
    
    def setup_ui(self):
        """Create the preview window UI"""
        # Top bar with controls
        top_bar = ttk.Frame(self, padding="5")
        top_bar.pack(fill='x')
        
        ttk.Button(top_bar, text="Fit to Window", command=self.fit_and_redraw).pack(side='left', padx=2)
        ttk.Button(top_bar, text="Zoom In", command=lambda: self.zoom(1.3)).pack(side='left', padx=2)
        ttk.Button(top_bar, text="Zoom Out", command=lambda: self.zoom(0.7)).pack(side='left', padx=2)
        ttk.Button(top_bar, text="Reset Positions", command=self._reset_positions).pack(side='left', padx=2)
        ttk.Button(top_bar, text="Refresh Allocation", command=self._refresh_allocation).pack(side='left', padx=2)
        
        self.lock_btn = ttk.Button(top_bar, text="Lock Allocation", command=self._toggle_allocation_lock)
        self.lock_btn.pack(side='left', padx=2)
        self._update_lock_button()
        
        ttk.Separator(top_bar, orient='vertical').pack(side='left', fill='y', padx=8, pady=2)
        
        self.align_motor_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            top_bar, text="Align on Motor",
            variable=self.align_motor_var,
            command=self._on_alignment_toggle
        ).pack(side='left', padx=4)
        
        ttk.Separator(top_bar, orient='vertical').pack(side='left', fill='y', padx=8, pady=2)
        
        ttk.Label(top_bar, text="Mode:").pack(side='left', padx=(0, 4))
        
        self.inspect_mode_var = tk.BooleanVar(value=self.inspect_mode)
        
        toggle_frame = ttk.Frame(top_bar)
        toggle_frame.pack(side='left', padx=4)
        
        self.toggle_canvas = tk.Canvas(toggle_frame, width=52, height=24,
                                        highlightthickness=0, bg=top_bar.winfo_toplevel().cget('bg'))
        self.toggle_canvas.pack(side='left')
        self.toggle_canvas.bind('<Button-1>', self._on_toggle_click)
        
        self.toggle_label = ttk.Label(toggle_frame, text="Layout", foreground='#333333')
        self.toggle_label.pack(side='left', padx=(4, 0))
        
        self._draw_toggle()

        # Sync label to initial state
        if self.inspect_mode:
            self.toggle_label.config(text="Inspect", foreground='#4CAF50')
        
        ttk.Separator(top_bar, orient='vertical').pack(side='left', fill='y', padx=8, pady=2)
        
        self.add_pad_btn = ttk.Button(top_bar, text="+ Add Pad", command=self._add_pad)
        self.add_pad_btn.pack(side='left', padx=4)
        
        self.assign_btn = ttk.Button(top_bar, text="Assign Devices", command=self._show_assignment_dialog)
        self.assign_btn.pack(side='left', padx=4)
        
        self.edit_devices_btn = ttk.Button(top_bar, text="Edit Devices", command=self._show_edit_devices_dialog)
        self.edit_devices_btn.pack(side='left', padx=4)

        self.show_routes_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            top_bar, text="Show Routes",
            variable=self.show_routes_var,
            command=self.draw
        ).pack(side='left', padx=4)
        
        self.zoom_label = ttk.Label(top_bar, text="100%")
        self.zoom_label.pack(side='left', padx=10)
        
        # Summary info
        num_inv = self.inv_summary.get('total_inverters', 0)
        total_str = self.inv_summary.get('total_strings', 0)
        actual_ratio = self.inv_summary.get('actual_dc_ac', 0)
        split = self.inv_summary.get('total_split_trackers', 0)
        
        summary_text = f"{num_inv} Inverters  |  {total_str} Strings  |  DC:AC: {actual_ratio:.2f}  |  {split} Split Trackers  |  {self.topology}"
        self.summary_label = ttk.Label(top_bar, text=summary_text, foreground='#333333')
        self.summary_label.pack(side='right', padx=10)
        
        # Canvas
        canvas_frame = ttk.Frame(self)
        canvas_frame.pack(fill='both', expand=True)
        
        self.canvas = tk.Canvas(canvas_frame, bg='white', highlightthickness=0)
        self.canvas.pack(fill='both', expand=True)
        
        # Bind events
        self.canvas.bind('<MouseWheel>', self.on_mousewheel)
        self.canvas.bind('<Button-4>', lambda e: self.zoom(1.1))
        self.canvas.bind('<Button-5>', lambda e: self.zoom(0.9))
        # Left-click: group select/drag only
        self.canvas.bind('<ButtonPress-1>', self.on_press)
        self.canvas.bind('<Double-Button-1>', self._on_device_double_click)
        self.canvas.bind('<B1-Motion>', self.on_motion)
        self.canvas.bind('<ButtonRelease-1>', self.on_release)
        # Middle-click: pan canvas
        self.canvas.bind('<ButtonPress-2>', self.on_pan_press)
        self.canvas.bind('<B2-Motion>', self.on_pan_motion)
        self.canvas.bind('<ButtonRelease-2>', self.on_pan_release)
        self.canvas.bind('<Button-3>', self._on_pad_right_click)
        self.canvas.bind('<Configure>', lambda e: self.draw())
        
        # Bottom legend (rebuildable)
        self.legend_frame = ttk.Frame(self, padding="5")
        self.legend_frame.pack(fill='x')
        self._build_legend()
    
    def _build_legend(self):
        """Build or rebuild the bottom legend with color swatches and allocation summary."""
        # Clear existing legend contents
        for child in self.legend_frame.winfo_children():
            child.destroy()
        
        # Color swatches
        swatch_frame = ttk.Frame(self.legend_frame)
        swatch_frame.pack(anchor='w')
        
        num_inv = self.inv_summary.get('total_inverters', 0)
        max_show = min(num_inv, 15)
        for i in range(max_show):
            color = self.colors[i % len(self.colors)]
            swatch = tk.Canvas(swatch_frame, width=12, height=12, highlightthickness=0)
            swatch.create_rectangle(0, 0, 12, 12, fill=color, outline='#333333')
            swatch.pack(side='left', padx=(0, 2))
            ttk.Label(swatch_frame, text=f"Inv {i+1}", font=('Helvetica', 8)).pack(side='left', padx=(0, 8))
        if num_inv > max_show:
            ttk.Label(swatch_frame, text=f"... +{num_inv - max_show} more",
                     font=('Helvetica', 8, 'italic'), foreground='gray').pack(side='left', padx=(5, 0))
        
        # Allocation summary
        allocation_result = self.inv_summary.get('allocation_result')
        if allocation_result:
            summary = allocation_result['summary']
            max_spi = summary.get('max_strings_per_inverter', 0)
            min_spi = summary.get('min_strings_per_inverter', 0)
            n_larger = summary.get('num_larger_inverters', 0)
            n_smaller = summary.get('num_smaller_inverters', 0)
            split_count = summary.get('total_split_trackers', 0)
            if max_spi == min_spi:
                size_str = f"All inverters: {max_spi} strings"
            else:
                size_str = f"{n_larger} inverters × {max_spi} strings + {n_smaller} inverters × {min_spi} strings"
            spatial_runs = allocation_result.get('spatial_runs', 1)
            runs_str = f"  |  {spatial_runs} spatial run(s)" if spatial_runs > 1 else ""
            ttk.Label(self.legend_frame, text=f"{size_str}  |  {split_count} split tracker(s){runs_str}",
                     font=('Helvetica', 9), foreground='#555555').pack(anchor='w')

    def build_layout_data(self):
        """Build a group-based layout of trackers with physical dimensions from templates.
        
        World units are in feet. Trackers run N-S (Y axis), spaced E-W (X axis).
        Each group has an (x, y) position in world-space representing its top-left corner.
        """
        self.group_layout = []
        self.selected_group_idx = None
        
        allocation_result = self.inv_summary.get('allocation_result')
        if not allocation_result:
            self.world_width = 0
            self.world_height = 0
            return
        
        # Build tracker_idx -> assignments from harness_map
        tracker_map = {}
        for inv_idx, inv in enumerate(allocation_result['inverters']):
            color = self.colors[inv_idx % len(self.colors)]
            for entry in inv['harness_map']:
                tidx = entry['tracker_idx']
                if tidx not in tracker_map:
                    tracker_map[tidx] = {
                        'strings_per_tracker': entry['strings_per_tracker'],
                        'assignments': []
                    }
                tracker_map[tidx]['assignments'].append({
                    'color': color,
                    'strings': entry['strings_taken'],
                    'inv_idx': inv_idx,
                    'start_physical_pos': entry.get('start_physical_pos', -1),
                })
        
        # Sort each tracker's assignments by physical position (north-to-south)
        for tidx in tracker_map:
            assignments = tracker_map[tidx]['assignments']
            if any(a.get('start_physical_pos', -1) >= 0 for a in assignments):
                assignments.sort(key=lambda a: a.get('start_physical_pos', 0))

        # Split tracker_map into groups with template dimensions
        global_idx = 0
        max_tracker_length_ft = 0
        max_tracker_width_ft = 0
        
        # Fallback dimensions for unlinked trackers
        fallback_width_ft = 6.0
        fallback_length_ft = 180.0
        
        for grp_idx, group_data in enumerate(self.groups):
            group_trackers = []
            group_motor_y = None  # motor Y offset for first tracker in group (used for alignment)
            
            for seg in group_data['segments']:
                ref = seg.get('template_ref')
                dims = self._get_preview_tracker_dims_ft(ref)
                
                for _ in range(seg['quantity']):
                    if global_idx in tracker_map:
                        tracker = tracker_map[global_idx].copy()
                        if dims:
                            tracker['width_ft'] = dims[0]
                            tracker['length_ft'] = dims[1]
                        else:
                            tracker['width_ft'] = fallback_width_ft
                            tracker['length_ft'] = fallback_length_ft
                        tracker['template_ref'] = ref
                        
                        # Motor position
                        motor_y, motor_gap, has_motor = self.get_motor_position_in_tracker(ref)
                        tracker['motor_y_ft'] = motor_y
                        tracker['motor_gap_ft'] = motor_gap
                        tracker['has_motor'] = has_motor
                        # Partial string info from template
                        if ref and ref in self.enabled_templates:
                            tdata_ps = self.enabled_templates[ref]
                            raw_spt = tdata_ps.get('strings_per_tracker', 1)
                            if raw_spt != int(raw_spt):
                                mps_ps = tdata_ps.get('modules_per_string', 28)
                                tracker['partial_module_count'] = round((raw_spt - int(raw_spt)) * mps_ps)
                                tracker['partial_string_side'] = tdata_ps.get('partial_string_side', 'north')
                                tracker['full_string_count'] = int(raw_spt)
                            else:
                                tracker['partial_module_count'] = 0
                                tracker['partial_string_side'] = 'north'
                                tracker['full_string_count'] = int(raw_spt)
                        
                        if group_motor_y is None and has_motor:
                            group_motor_y = motor_y

                        group_trackers.append(tracker)
                        
                        max_tracker_width_ft = max(max_tracker_width_ft, tracker['width_ft'])
                        max_tracker_length_ft = max(max_tracker_length_ft, tracker['length_ft'])
                    global_idx += 1
            
            # Group dimensions (bounding box of all its trackers laid out E-W)
            num_trackers = len(group_trackers)
            if num_trackers > 0:
                group_max_width = max(t['width_ft'] for t in group_trackers)
                group_width = group_max_width + (num_trackers - 1) * self.row_spacing_ft
                group_length = max(t['length_ft'] for t in group_trackers)
            else:
                group_width = 0
                group_length = 0
            
            # Compute string length for NS snap offset (from first linked template)
            string_length_ft = 0
            for seg in group_data['segments']:
                ref = seg.get('template_ref')
                if ref and ref in self.enabled_templates:
                    tdata = self.enabled_templates[ref]
                    ms = tdata.get('module_spec', {})
                    mps = tdata.get('modules_per_string', 28)
                    orientation = tdata.get('module_orientation', 'Portrait')
                    if orientation == 'Portrait':
                        mod_along_m = ms.get('width_mm', 1000) / 1000
                    else:
                        mod_along_m = ms.get('length_mm', 2000) / 1000
                    spacing_m = tdata.get('module_spacing_m', 0.02)
                    string_length_ft = (mps * mod_along_m + (mps - 1) * spacing_m) * 3.28084
                    break
            
            # Compute visual bounding box offsets considering motor alignment
            # When align_on_motor is active, trackers shift vertically so their
            # motors match the group's reference motor_y. This affects the actual
            # visual extent of the group.
            ref_motor = group_motor_y or 0
            visual_min_y_offset = 0.0
            visual_max_y_offset = 0.0
            
            # Driveline angle: each tracker offset in Y by t_idx * pitch * tan(angle)
            driveline_angle_deg = group_data.get('driveline_angle', 0.0)
            driveline_angle_rad = math.radians(driveline_angle_deg)
            driveline_tan = math.tan(driveline_angle_rad) if driveline_angle_deg > 0 else 0.0
            
            visual_min_y_base = 0.0
            visual_max_y_base = 0.0
            
            if group_trackers and ref_motor is not None:
                for t_i, t in enumerate(group_trackers):
                    t_motor = t.get('motor_y_ft', 0)
                    t_length_val = t.get('length_ft', group_length)
                    y_offset = ref_motor - t_motor  # Motor alignment shift (no angle)
                    angle_y = t_i * self.row_spacing_ft * driveline_tan
                    # Base bounds (no angle) — for parallelogram overlap checking
                    visual_min_y_base = min(visual_min_y_base, y_offset)
                    visual_max_y_base = max(visual_max_y_base, y_offset + t_length_val)
                    # Full bounds (with angle) — for bounding box and selection highlight
                    visual_min_y_offset = min(visual_min_y_offset, y_offset + angle_y)
                    visual_max_y_offset = max(visual_max_y_offset, y_offset + angle_y + t_length_val)
            
            self.group_layout.append({
                'name': group_data['name'],
                'trackers': group_trackers,
                'width_ft': group_width,
                'length_ft': group_length,
                'motor_y_ft': group_motor_y or 0,
                'string_length_ft': string_length_ft,
                'group_idx': grp_idx,
                'visual_min_y': visual_min_y_offset,
                'visual_max_y': visual_max_y_offset,
                'visual_min_y_base': visual_min_y_base,
                'visual_max_y_base': visual_max_y_base,
                'driveline_angle': driveline_angle_deg,
                'driveline_tan': driveline_tan,
            })

        # Flat list for backward compat
        self.tracker_list = [tracker_map[i] for i in sorted(tracker_map.keys())]
        
        # Store global metrics
        self.tracker_pitch_ft = self.row_spacing_ft
        self.tracker_gap_ft = max(self.row_spacing_ft - max_tracker_width_ft, 1.0)
        self.max_tracker_width_ft = max_tracker_width_ft if max_tracker_width_ft > 0 else fallback_width_ft
        self.max_tracker_length_ft = max_tracker_length_ft if max_tracker_length_ft > 0 else fallback_length_ft
        self.group_ns_gap_ft = self.max_tracker_length_ft * 0.15
        
        # Assign initial positions from saved data or auto-layout
        self._assign_group_positions()
        
        # Compute device (CB/SI) positions per group
        self._compute_device_positions()
        
        # Compute world bounds from actual positions (including devices)
        self._update_world_bounds()
    
    def _assign_group_positions(self):
        """Assign (x, y) positions to each group. Use saved positions if available,
        otherwise auto-layout stacking groups left-to-right."""
        for grp_idx, layout in enumerate(self.group_layout):
            group_data = self.groups[grp_idx] if grp_idx < len(self.groups) else {}
            
            saved_x = group_data.get('position_x')
            saved_y = group_data.get('position_y')
            
            if saved_x is not None and saved_y is not None:
                layout['x'] = saved_x
                layout['y'] = saved_y
            else:
                # Auto-layout: stack groups left to right with spacing
                # Each group is placed one group-width + gap to the right
                group_spacing = self.max_tracker_length_ft * 0.1
                layout['x'] = grp_idx * (layout['width_ft'] + group_spacing)
                layout['y'] = 0
    
    def _compute_device_positions(self):
        """Compute world-space positions for combiner boxes / string inverters.
        
        For Distributed String and Centralized String: derives placement from
        the allocation result so each device maps 1:1 to an inverter.
        
        For Central Inverter: distributes CBs proportionally across groups
        by tracker count.
        
        Device Y is determined by the group's device_position setting:
          - 'north': offset above northernmost tracker edge
          - 'south': offset below southernmost tracker edge  
          - 'middle': at the motor/driveline Y
        """
        self.device_positions = []
        
        if not self.group_layout:
            return
        
        device_width_ft = 4.0
        device_height_ft = 3.0
        offset_ft = 5.0
        
        alloc = self.inv_summary.get('allocation_result', {}) if hasattr(self, 'inv_summary') else {}
        inverters = alloc.get('inverters', [])
        
        if self.topology in ('Distributed String', 'Centralized String') and inverters:
            self._compute_devices_from_allocation(
                inverters, device_width_ft, device_height_ft, offset_ft
            )
        elif self.topology == 'Central Inverter':
            self._compute_devices_proportional(
                device_width_ft, device_height_ft, offset_ft
            )
    
    def _compute_devices_from_allocation(self, inverters, device_width_ft, device_height_ft, offset_ft):
        """Place one device per inverter, positioned at the center of that inverter's trackers."""
        pitch = self.tracker_pitch_ft
        max_width = self.max_tracker_width_ft
        
        # Build a lookup: global_tracker_idx -> (group_idx, local_tracker_idx)
        tracker_to_group = {}
        running = 0
        for grp_idx, grp in enumerate(self.group_layout):
            for local_idx in range(len(grp['trackers'])):
                tracker_to_group[running] = (grp_idx, local_idx)
                running += 1
        
        for inv_idx, inv in enumerate(inverters):
            harness_map = inv.get('harness_map', [])
            if not harness_map:
                continue
            
            # Find which trackers this inverter uses
            inv_tracker_indices = [entry['tracker_idx'] for entry in harness_map]
            
            # Determine majority group
            group_counts = {}
            for tidx in inv_tracker_indices:
                if tidx in tracker_to_group:
                    grp_idx = tracker_to_group[tidx][0]
                    group_counts[grp_idx] = group_counts.get(grp_idx, 0) + 1
            
            if not group_counts:
                continue
            
            primary_grp_idx = max(group_counts, key=group_counts.get)
            group_data = self.group_layout[primary_grp_idx]
            group_source = self.groups[primary_grp_idx] if primary_grp_idx < len(self.groups) else {}
            device_position = group_source.get('device_position', 'middle')
            
            gx = group_data['x']
            gy = group_data['y']
            
            # Compute X from the center of this inverter's trackers within the primary group
            local_indices = []
            for tidx in inv_tracker_indices:
                if tidx in tracker_to_group and tracker_to_group[tidx][0] == primary_grp_idx:
                    local_indices.append(tracker_to_group[tidx][1])
            
            if local_indices:
                center_local = (min(local_indices) + max(local_indices)) / 2.0
                device_x = gx + center_local * pitch + (max_width - device_width_ft) / 2
            else:
                center_local = 0
                device_x = gx
            
            # Driveline angle Y offset based on device's X position in group
            angle_y_offset = center_local * pitch * group_data.get('driveline_tan', 0.0)
            
            # Compute Y based on position setting
            if device_position == 'north':
                vis_min = group_data.get('visual_min_y', 0)
                device_y = gy + vis_min - offset_ft - device_height_ft + angle_y_offset
            elif device_position == 'south':
                vis_max = group_data.get('visual_max_y', group_data['length_ft'])
                device_y = gy + vis_max + offset_ft + angle_y_offset
            else:  # 'middle'
                motor_y = group_data.get('motor_y_ft', group_data['length_ft'] / 2)
                device_y = gy + motor_y - device_height_ft / 2 + angle_y_offset
            
            # Build assigned_strings from this inverter's harness_map
            assigned_strings = {}
            
            # Prefer combiner assignments (have exact start_string_pos from Edit Devices)
            parent_qe = self.master
            cb_assignments = getattr(parent_qe, 'last_combiner_assignments', [])
            if cb_assignments and inv_idx < len(cb_assignments):
                cb = cb_assignments[inv_idx]
                for conn in cb.get('connections', []):
                    tidx = conn['tracker_idx']
                    start = conn.get('start_string_pos', 0)
                    n = conn['num_strings']
                    if tidx not in assigned_strings:
                        assigned_strings[tidx] = set()
                    for s in range(start, start + n):
                        assigned_strings[tidx].add(s)
            else:
                # Fallback to harness_map heuristic
                for entry in harness_map:
                    tidx = entry['tracker_idx']
                    strings_taken = entry['strings_taken']
                    spt = entry['strings_per_tracker']
                    is_split = entry.get('is_split', False)
                    split_pos = entry.get('split_position', 'full')
                    
                    if tidx not in assigned_strings:
                        assigned_strings[tidx] = set()
                    
                    if is_split and split_pos == 'tail':
                        start_idx = spt - strings_taken
                        for s in range(start_idx, spt):
                            assigned_strings[tidx].add(s)
                    else:
                        existing = len(assigned_strings[tidx])
                        for s in range(existing, existing + strings_taken):
                            assigned_strings[tidx].add(s)
            
            dev_idx = len(self.device_positions)
            label = self.device_names.get(dev_idx, f"{self.device_label}-{inv_idx + 1:02d}")
            
            self.device_positions.append({
                'x': device_x,
                'y': device_y,
                'width_ft': device_width_ft,
                'height_ft': device_height_ft,
                'label': label,
                'group_idx': primary_grp_idx,
                'device_position': device_position,
                'assigned_strings': assigned_strings,
            })
    
    def _compute_devices_proportional(self, device_width_ft, device_height_ft, offset_ft):
        """Distribute devices proportionally across groups for Central Inverter topology."""
        pitch = self.tracker_pitch_ft
        max_width = self.max_tracker_width_ft
        
        total_trackers = sum(len(g['trackers']) for g in self.group_layout)
        if total_trackers <= 0 or self.num_devices <= 0:
            return
        
        global_device_idx = 0
        
        for grp_idx, group_data in enumerate(self.group_layout):
            group_trackers = group_data['trackers']
            num_trackers_in_group = len(group_trackers)
            if num_trackers_in_group == 0:
                continue
            
            group_share = num_trackers_in_group / total_trackers
            group_device_count = max(1, round(group_share * self.num_devices))
            remaining = self.num_devices - global_device_idx
            group_device_count = min(group_device_count, remaining)
            
            if group_device_count <= 0:
                continue
            
            group_source = self.groups[grp_idx] if grp_idx < len(self.groups) else {}
            device_position = group_source.get('device_position', 'middle')
            
            gx = group_data['x']
            gy = group_data['y']
            
            # Base Y (before driveline angle offset)
            if device_position == 'north':
                vis_min = group_data.get('visual_min_y', 0)
                base_device_y = gy + vis_min - offset_ft - device_height_ft
            elif device_position == 'south':
                vis_max = group_data.get('visual_max_y', group_data['length_ft'])
                base_device_y = gy + vis_max + offset_ft
            else:
                motor_y = group_data.get('motor_y_ft', group_data['length_ft'] / 2)
                base_device_y = gy + motor_y - device_height_ft / 2
            
            driveline_tan = group_data.get('driveline_tan', 0.0)
            
            # Even spacing within group
            group_global_start = sum(len(self.group_layout[g]['trackers']) for g in range(grp_idx))
            base_group_size = num_trackers_in_group // group_device_count
            extra = num_trackers_in_group % group_device_count
            
            tracker_start = 0
            for dev_i in range(group_device_count):
                sub_size = base_group_size + (1 if dev_i < extra else 0)
                if sub_size <= 0:
                    continue
                
                center_tracker = tracker_start + sub_size / 2.0 - 0.5
                device_x = gx + center_tracker * pitch + (max_width - device_width_ft) / 2
                device_y = base_device_y + center_tracker * pitch * driveline_tan
                
                # All strings in tracker range belong to this CB
                assigned_strings = {}
                for local_idx in range(tracker_start, tracker_start + sub_size):
                    global_idx = group_global_start + local_idx
                    if local_idx < len(group_trackers):
                        spt = group_trackers[local_idx].get('strings_per_tracker', 0)
                        assigned_strings[global_idx] = set(range(int(spt)))
                
                self.device_positions.append({
                    'x': device_x,
                    'y': device_y,
                    'width_ft': device_width_ft,
                    'height_ft': device_height_ft,
                    'label': self.device_names.get(global_device_idx, f"CB-{global_device_idx + 1:02d}"),
                    'group_idx': grp_idx,
                    'device_position': device_position,
                    'assigned_strings': assigned_strings,
                })
                
                global_device_idx += 1
                tracker_start += sub_size
        
        # Fill any remaining devices
        while global_device_idx < self.num_devices and self.device_positions:
            last = self.device_positions[-1].copy()
            last['label'] = self.device_names.get(global_device_idx, f"CB-{global_device_idx + 1:02d}")
            last['x'] += device_width_ft + 2
            last['assigned_strings'] = {}
            self.device_positions.append(last)
            global_device_idx += 1

    def _update_world_bounds(self):
        """Recompute world_width and world_height from actual group positions."""
        if not self.group_layout:
            self.world_width = 0
            self.world_height = 0
            return
        
        min_x = min(g['x'] for g in self.group_layout)
        max_x = max(g['x'] + g['width_ft'] for g in self.group_layout)
        min_y = min(g['y'] for g in self.group_layout)
        max_y = max(g['y'] + g['length_ft'] for g in self.group_layout)
        
        # Include device positions in bounds
        if hasattr(self, 'device_positions') and self.device_positions:
            for dev in self.device_positions:
                min_x = min(min_x, dev['x'])
                max_x = max(max_x, dev['x'] + dev['width_ft'])
                min_y = min(min_y, dev['y'])
                max_y = max(max_y, dev['y'] + dev['height_ft'])

        # Include pads in bounds
        if hasattr(self, 'pads') and self.pads:
            for pad in self.pads:
                pw = pad.get('width_ft', 10.0)
                ph = pad.get('height_ft', 8.0)
                min_x = min(min_x, pad['x'])
                max_x = max(max_x, pad['x'] + pw)
                min_y = min(min_y, pad['y'])
                max_y = max(max_y, pad['y'] + ph)
        
        # Add margin
        margin = self.max_tracker_width_ft * 2
        self.world_min_x = min_x - margin
        self.world_min_y = min_y - margin
        self.world_width = (max_x - min_x) + margin * 2
        self.world_height = (max_y - min_y) + margin * 2
    
    def _save_group_positions(self):
        """Save current group positions back to the group data dicts."""
        for grp_idx, layout in enumerate(self.group_layout):
            if grp_idx < len(self.groups):
                self.groups[grp_idx]['position_x'] = layout['x']
                self.groups[grp_idx]['position_y'] = layout['y']
    
    def fit_to_canvas(self):
        """Calculate scale and pan to fit all content"""
        self.canvas.update_idletasks()
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        
        if cw < 10 or ch < 10:
            cw = 1100
            ch = 750
        
        margin = 40
        
        if self.world_width <= 0 or self.world_height <= 0:
            return
        
        scale_x = (cw - 2 * margin) / self.world_width
        scale_y = (ch - 2 * margin) / self.world_height
        self.scale = min(scale_x, scale_y)
        
        # Center on actual content bounds
        min_x = getattr(self, 'world_min_x', 0)
        min_y = getattr(self, 'world_min_y', 0)
        
        scaled_w = self.world_width * self.scale
        scaled_h = self.world_height * self.scale
        self.pan_x = (cw - scaled_w) / 2 - min_x * self.scale
        self.pan_y = (ch - scaled_h) / 2 - min_y * self.scale
    
    def fit_and_redraw(self):
        """Fit to window and redraw"""
        self.fit_to_canvas()
        self.draw()
    
    def world_to_canvas(self, wx, wy):
        """Convert world coordinates to canvas coordinates"""
        cx = self.pan_x + wx * self.scale
        cy = self.pan_y + wy * self.scale
        return cx, cy
    
    def zoom(self, factor):
        """Zoom in/out centered on the canvas"""
        self.canvas.update_idletasks()
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        
        center_x = cw / 2
        center_y = ch / 2
        
        self.pan_x = center_x - (center_x - self.pan_x) * factor
        self.pan_y = center_y - (center_y - self.pan_y) * factor
        self.scale *= factor
        
        self.zoom_label.config(text=f"{self.scale * 100:.0f}%")
        self.draw()
    
    def on_mousewheel(self, event):
        """Handle mouse wheel zoom"""
        if event.delta > 0:
            self.zoom(1.1)
        else:
            self.zoom(0.9)
    
    def canvas_to_world(self, cx, cy):
        """Convert canvas pixel coordinates to world-space feet."""
        if self.scale == 0:
            return 0, 0
        wx = (cx - self.pan_x) / self.scale
        wy = (cy - self.pan_y) / self.scale
        return wx, wy
    
    def hit_test_group(self, cx, cy):
        """Return the index of the group under canvas coords (cx, cy), or None."""
        wx, wy = self.canvas_to_world(cx, cy)
        # Check in reverse order so topmost (last drawn) is hit first
        for i in range(len(self.group_layout) - 1, -1, -1):
            g = self.group_layout[i]
            vis_min = g.get('visual_min_y', 0)
            vis_max = g.get('visual_max_y', g['length_ft'])
            if (g['x'] <= wx <= g['x'] + g['width_ft'] and
                g['y'] + vis_min <= wy <= g['y'] + vis_max):
                return i
        return None
    
    def on_press(self, event):
        """Handle left mouse press — place pad, select/drag group or pad, or select device."""
        self.drag_start_x = event.x
        self.drag_start_y = event.y
        self._drag_moved = False
        self._dragging_pad = False
        
        # Pad placement mode — click to place
        if self.placing_pad:
            wx, wy = self.canvas_to_world(event.x, event.y)
            self._place_pad_at(wx, wy)
            return
        
        if self.inspect_mode:
            # Check info panel first (topmost layer)
            if hasattr(self, '_info_panel_bounds') and self._info_panel_bounds:
                bx1, by1, bx2, by2 = self._info_panel_bounds
                if bx1 <= event.x <= bx2 and by1 <= event.y <= by2:
                    self._dragging_panel = True
                    self._panel_drag_start = (event.x, event.y)
                    self._panel_drag_dev = self.selected_device_idx
                    return

            # Check pads first (drawn on top)
            hit_pad = self.hit_test_pad(event.x, event.y)
            if hit_pad is not None:
                if self.selected_pad_inspect_idx == hit_pad:
                    self.selected_pad_inspect_idx = None
                else:
                    self.selected_pad_inspect_idx = hit_pad
                self.selected_device_idx = None
                self.dragging_canvas = False
                self.draw()
                self.dragging_group = False
                return
            
            # Then check devices
            hit_dev = self.hit_test_device(event.x, event.y)
            if hit_dev is not None:
                if self.selected_device_idx == hit_dev:
                    self.selected_device_idx = None
                else:
                    self.selected_device_idx = hit_dev
                self.selected_pad_inspect_idx = None
                self.dragging_canvas = False
                self.draw()
            else:
                # Empty space — clear selections and start panning
                self.selected_device_idx = None
                self.selected_pad_inspect_idx = None
                self.dragging_canvas = True
                self.draw()
            self.dragging_group = False
            return
        
        # Layout mode — check pads first (they're on top visually)
        hit_pad = self.hit_test_pad(event.x, event.y)
        if hit_pad is not None:
            self.selected_pad_idx = hit_pad
            self._dragging_pad = True
            pad = self.pads[hit_pad]
            self._drag_pad_start_x = pad['x']
            self._drag_pad_start_y = pad['y']
            self.selected_group_idx = None
            self.dragging_group = False
            self.draw()
            return
        
        self.selected_pad_idx = None
        
        hit = self.hit_test_group(event.x, event.y)
        if hit is not None:
            self.selected_group_idx = hit
            self.dragging_group = True
            g = self.group_layout[hit]
            self._drag_group_start_x = g['x']
            self._drag_group_start_y = g['y']
            self.draw()
        else:
            self.selected_group_idx = None
            self.dragging_group = False
            self.draw()

    def _on_device_double_click(self, event):
        """Handle double-click on canvas — rename device if clicked on one."""
        dev_idx = self.hit_test_device(event.x, event.y)
        if dev_idx is None:
            return
        
        dev = self.device_positions[dev_idx]
        current_name = dev.get('label', f"{self.device_label}-{dev_idx + 1:02d}")
        
        from tkinter import simpledialog
        new_name = simpledialog.askstring(
            "Rename Device",
            f"Enter new name for {current_name}:",
            initialvalue=current_name,
            parent=self
        )
        
        if new_name and new_name.strip():
            new_name = new_name.strip()
            self.device_names[dev_idx] = new_name
            dev['label'] = new_name
            self.draw()
    
    def on_motion(self, event):
        """Handle mouse drag — move group, move pad, or pan canvas."""
        # Panel drag
        if getattr(self, '_dragging_panel', False) and self._panel_drag_start:
            dx = event.x - self._panel_drag_start[0]
            dy = event.y - self._panel_drag_start[1]
            dev_idx = self._panel_drag_dev
            if not hasattr(self, '_info_panel_offset'):
                self._info_panel_offset = {}
            old = self._info_panel_offset.get(dev_idx, (0, -100))
            self._info_panel_offset[dev_idx] = (old[0] + dx, old[1] + dy)
            self._panel_drag_start = (event.x, event.y)
            self.draw()
            return
        dx_px = event.x - self.drag_start_x
        dy_px = event.y - self.drag_start_y
        
        if abs(dx_px) > 3 or abs(dy_px) > 3:
            self._drag_moved = True
        
        if getattr(self, '_dragging_pad', False) and self.selected_pad_idx is not None:
            dx_world = dx_px / self.scale if self.scale != 0 else 0
            dy_world = dy_px / self.scale if self.scale != 0 else 0
            self.pads[self.selected_pad_idx]['x'] = self._drag_pad_start_x + dx_world
            self.pads[self.selected_pad_idx]['y'] = self._drag_pad_start_y + dy_world
            self.draw()
        elif getattr(self, 'dragging_group', False) and self.selected_group_idx is not None:
            dx_world = dx_px / self.scale if self.scale != 0 else 0
            dy_world = dy_px / self.scale if self.scale != 0 else 0
            
            new_x = self._drag_group_start_x + dx_world
            new_y = self._drag_group_start_y + dy_world
            
            shift_held = event.state & 0x1
            if shift_held:
                # Shift held — constrain to N/S movement only (lock X)
                new_x = self._drag_group_start_x
            else:
                # Normal drag — apply snapping
                new_x, new_y = self._snap_group_position(
                    self.selected_group_idx, new_x, new_y
                )
            
            self.group_layout[self.selected_group_idx]['x'] = new_x
            self.group_layout[self.selected_group_idx]['y'] = new_y
            
            self.draw()
    
    def on_release(self, event):
        """Handle mouse release — finalize group or pad position."""
        if getattr(self, '_dragging_panel', False):
            self._dragging_panel = False
            self._panel_drag_start = None
            return

        if getattr(self, '_dragging_pad', False) and self._drag_moved:
            self._update_world_bounds()
        
        if getattr(self, 'dragging_group', False) and self._drag_moved:
            self._update_world_bounds()
            self._save_group_positions()
            
            overlaps = self._check_overlaps()
            if overlaps:
                names = set()
                for i, j in overlaps:
                    names.add(self.group_layout[i].get('name', f'Group {i+1}'))
                    names.add(self.group_layout[j].get('name', f'Group {j+1}'))
        
        self.dragging_group = False
        self.dragging_canvas = False
        self._dragging_pad = False

    def on_pan_press(self, event):
        """Handle middle mouse press — start panning."""
        self.dragging_canvas = True
        self.drag_start_x = event.x
        self.drag_start_y = event.y

    def on_pan_motion(self, event):
        """Handle middle mouse drag — pan canvas."""
        if self.dragging_canvas:
            dx_px = event.x - self.drag_start_x
            dy_px = event.y - self.drag_start_y
            self.pan_x += dx_px
            self.pan_y += dy_px
            self.drag_start_x = event.x
            self.drag_start_y = event.y
            self.draw()

    def on_pan_release(self, event):
        """Handle middle mouse release — stop panning."""
        self.dragging_canvas = False
    
    def _snap_group_position(self, group_idx, raw_x, raw_y):
        """Apply snapping to a group's proposed position.
        
        EW (X): Snap to row-spacing pitch grid.
        NS (Y): Snap motor to align with ANY nearby group's motor.
        Checks all groups and picks the closest snap candidate.
        """
        group = self.group_layout[group_idx]
        my_motor_offset = group.get('motor_y_ft', 0)
        string_len = group.get('string_length_ft', 0)
        
        snapped_x = raw_x
        snapped_y = raw_y
        
        # EW snap: align to pitch grid
        pitch = self.tracker_pitch_ft
        if pitch > 0:
            snapped_x = round(raw_x / pitch) * pitch
        
        # NS snap: check motor alignment against ALL other groups
        snap_threshold = self.max_tracker_length_ft * 0.15
        best_snap_y = None
        best_snap_dist = float('inf')
        
        for i, g in enumerate(self.group_layout):
            if i == group_idx:
                continue
            
            neighbor_motor_world_y = g['y'] + g.get('motor_y_ft', 0)
            
            # Target Y so motors align
            motor_aligned_y = neighbor_motor_world_y - my_motor_offset
            
            # Also compute string-offset positions
            candidates = [
                motor_aligned_y,                     # motor alignment
                motor_aligned_y + string_len,        # offset +1 string south
                motor_aligned_y - string_len,        # offset +1 string north
            ]
            
            for candidate_y in candidates:
                dist = abs(raw_y - candidate_y)
                if dist < best_snap_dist and dist < snap_threshold:
                    best_snap_dist = dist
                    best_snap_y = candidate_y
        
        if best_snap_y is not None:
            snapped_y = best_snap_y
        
        return snapped_x, snapped_y
    
    def draw(self):
        """Draw the site layout with to-scale trackers at their group positions.
        
        X = E-W (tracker width + row spacing gaps)
        Y = N-S (tracker length, north at top)
        World units = feet.
        """
        self.canvas.delete('all')
        
        if not self.group_layout:
            return
        
        pitch = getattr(self, 'tracker_pitch_ft', 20)
        max_width = getattr(self, 'max_tracker_width_ft', 6)
        
        for group_idx, group_data in enumerate(self.group_layout):
            gx = group_data['x']
            gy = group_data['y']
            is_selected = (group_idx == self.selected_group_idx)
            
            # Draw selection highlight behind group (using visual bounds)
            if is_selected:
                pad = max_width * 0.3
                vis_min = group_data.get('visual_min_y', 0)
                vis_max = group_data.get('visual_max_y', group_data['length_ft'])
                hx1, hy1 = self.world_to_canvas(gx - pad, gy + vis_min - pad)
                hx2, hy2 = self.world_to_canvas(
                    gx + group_data['width_ft'] + pad,
                    gy + vis_max + pad
                )
                self.canvas.create_rectangle(
                    hx1, hy1, hx2, hy2,
                    fill='', outline='#4A90D9', width=2, dash=(6, 3)
                )
            
            # Draw group label
            label_x, label_y = self.world_to_canvas(
                gx - max_width * 0.5,
                gy + group_data['length_ft'] / 2
            )
            font_size = max(6, min(11, int(9 * self.scale)))
            self.canvas.create_text(
                label_x, label_y,
                text=group_data['name'], font=('Helvetica', font_size),
                fill='#4A90D9' if is_selected else '#333333', anchor='e'
            )
            
            for t_idx, tracker in enumerate(group_data['trackers']):
                spt = tracker['strings_per_tracker']
                assignments = tracker['assignments']
                t_width = tracker.get('width_ft', max_width)
                t_length = tracker.get('length_ft', 100)
                
                # X position within group: center-to-center = pitch
                tx = gx + t_idx * pitch
                # Center tracker within pitch slot
                tx_offset = (max_width - t_width) / 2 if max_width > t_width else 0
                
                # Driveline angle: offset each tracker in Y
                angle_y_offset = t_idx * pitch * group_data.get('driveline_tan', 0.0)
                
                # Align tracker vertically within group
                if getattr(self, 'align_on_motor', False) and tracker.get('has_motor', False) and group_data.get('motor_y_ft', None) is not None:
                    # Motor alignment: offset so this tracker's motor Y matches group's reference motor Y
                    reference_motor_y = group_data['motor_y_ft']
                    ty = gy + (reference_motor_y - tracker['motor_y_ft']) + angle_y_offset
                else:
                    # Center alignment fallback
                    max_group_length = group_data['length_ft']
                    ty_offset = (max_group_length - t_length) / 2
                    ty = gy + ty_offset + angle_y_offset              
                # Per-string height — adjust for partial strings
                partial_mods = tracker.get('partial_module_count', 0)
                partial_side = tracker.get('partial_string_side', 'north')
                full_str_count = tracker.get('full_string_count', spt)
                mps_for_height = 26  # fallback
                
                ref = tracker.get('template_ref')
                if partial_mods > 0 and ref and ref in self.enabled_templates:
                    mps_for_height = self.enabled_templates[ref].get('modules_per_string', 26)
                
                if partial_mods > 0 and full_str_count > 0:
                    total_mods = full_str_count * mps_for_height + partial_mods
                    module_extent = t_length
                    full_height = (module_extent * mps_for_height / total_mods) if total_mods > 0 else module_extent
                    partial_height = (module_extent * partial_mods / total_mods) if total_mods > 0 else 0
                    
                    # Build height list per effective string slot
                    # Always include the partial band (even for right-of-pair trackers)
                    string_heights = []
                    has_owned_partial = (spt > full_str_count)
                    draw_spt = spt + (1 if not has_owned_partial else 0)  # Add unowned partial band
                    
                    if partial_side == 'north':
                        string_heights.append(partial_height)  # Always draw partial band
                        for _ in range(full_str_count):
                            string_heights.append(full_height)
                    else:  # south
                        for _ in range(full_str_count):
                            string_heights.append(full_height)
                        string_heights.append(partial_height)  # Always draw partial band
                else:
                    string_height = t_length / spt if spt > 0 else t_length
                    string_heights = [string_height] * int(spt)
                
                # Build string colors
                string_colors = []
                for assignment in assignments:
                    for _ in range(assignment['strings']):
                        string_colors.append(assignment['color'])
                
                # Determine global tracker index for device highlighting
                global_tracker_idx = sum(
                    len(self.group_layout[g]['trackers']) for g in range(group_idx)
                ) + t_idx
                
                # Check if we're highlighting a selected device or pad
                highlighting = False
                selected_strings = set()
                
                if self.inspect_mode and hasattr(self, 'device_positions') and self.device_positions:
                    if self.selected_device_idx is not None:
                        highlighting = True
                        dev = self.device_positions[self.selected_device_idx]
                        assigned = dev.get('assigned_strings', {})
                        selected_strings = assigned.get(global_tracker_idx, set())
                    elif self.selected_pad_inspect_idx is not None:
                        highlighting = True
                        # Collect strings from ALL devices assigned to this pad
                        pad = self.pads[self.selected_pad_inspect_idx] if self.selected_pad_inspect_idx < len(self.pads) else None
                        if pad:
                            for dev_idx in pad.get('assigned_devices', []):
                                if dev_idx < len(self.device_positions):
                                    dev = self.device_positions[dev_idx]
                                    assigned = dev.get('assigned_strings', {})
                                    selected_strings.update(assigned.get(global_tracker_idx, set()))
                
                # Draw each string (including unowned partial bands)
                draw_count = len(string_heights)
                for s_idx in range(draw_count):
                    # Determine if this is an unowned partial band
                    is_unowned_partial = (partial_mods > 0 and spt <= full_str_count and
                                         ((partial_side == 'north' and s_idx == 0) or
                                          (partial_side == 'south' and s_idx == draw_count - 1)))
                    
                    if is_unowned_partial:
                        color = '#D4C878'  # Muted gold for unowned partial
                    else:
                        # Map drawing index back to allocation string index
                        # Skip the partial band position to get the right color
                        if partial_mods > 0 and partial_side == 'north':
                            color_idx = s_idx - 1  # Partial is at 0, so shift down
                            if spt > full_str_count and s_idx == 0:
                                color_idx = 0  # Owned partial gets first color
                        elif partial_mods > 0 and partial_side == 'south':
                            if spt > full_str_count and s_idx == draw_count - 1:
                                color_idx = spt - 1  # Owned partial gets last color
                            else:
                                color_idx = s_idx
                        else:
                            color_idx = s_idx
                        color = string_colors[color_idx] if 0 <= color_idx < len(string_colors) else '#D0D0D0'
                    
                    if highlighting:
                        if s_idx in selected_strings:
                            outline_color = '#FF6600'
                            outline_width = 2
                        else:
                            # Dim unselected strings
                            color = '#E0E0E0'
                            outline_color = '#CCCCCC'
                            outline_width = 1
                    elif self.assigning_devices:
                        # Grey out all trackers while Assign Devices dialog is open
                        color = '#E0E0E0'
                        outline_color = '#CCCCCC'
                        outline_width = 1
                    else:
                        outline_color = '#555555'
                        outline_width = 1
                    
                    # Yellow highlight from Edit Devices dialog
                    _hl = getattr(self, '_highlighted_strings', set())
                    if (not is_unowned_partial and _hl and
                            (global_tracker_idx, s_idx) in _hl):
                        color = '#FFFF00'
                        outline_color = '#DAA520'
                        outline_width = 2
                    
                    sy = ty + sum(string_heights[:s_idx])
                    sh = string_heights[s_idx] if s_idx < len(string_heights) else string_heights[-1]
                    
                    sx1, sy1 = self.world_to_canvas(tx + tx_offset, sy)
                    sx2, sy2 = self.world_to_canvas(tx + tx_offset + t_width, sy + sh)
                    
                    self.canvas.create_rectangle(
                        sx1, sy1, sx2, sy2,
                        fill=color, outline=outline_color, width=outline_width
                    )
                
                # Tracker outline
                ox1, oy1 = self.world_to_canvas(tx + tx_offset - 0.5, ty - 0.5)
                ox2, oy2 = self.world_to_canvas(tx + tx_offset + t_width + 0.5, ty + t_length + 0.5)
                
                self.canvas.create_rectangle(
                    ox1, oy1, ox2, oy2,
                    fill='', outline='#222222', width=1
                )
                
                # Motor indicator
                if tracker.get('has_motor', False):
                    motor_y = tracker['motor_y_ft']
                    motor_gap = tracker['motor_gap_ft']
                    
                    motor_world_y = ty + motor_y
                    motor_x1 = tx + tx_offset - 0.3
                    motor_x2 = tx + tx_offset + t_width + 0.3
                    
                    mx1, my1 = self.world_to_canvas(motor_x1, motor_world_y)
                    mx2, my2 = self.world_to_canvas(motor_x2, motor_world_y + motor_gap)
                    
                    self.canvas.create_rectangle(
                        mx1, my1, mx2, my2,
                        fill='#666666', outline='#444444', width=1
                    )
                    
                    motor_cx = (mx1 + mx2) / 2
                    motor_cy = (my1 + my2) / 2
                    dot_r = max(2, min(4, 3 * self.scale))
                    self.canvas.create_oval(
                        motor_cx - dot_r, motor_cy - dot_r,
                        motor_cx + dot_r, motor_cy + dot_r,
                        fill='#FF8800', outline='#CC6600', width=1
                    )
                
                # Tracker label — use global tracker index to match info panel / assignments
                label_cx, label_cy = self.world_to_canvas(
                    tx + tx_offset + t_width / 2,
                    ty + t_length + 2
                )
                pixel_width = abs(ox2 - ox1)
                if pixel_width > 14:
                    lbl_size = max(6, min(9, int(8 * self.scale)))
                    self.canvas.create_text(
                        label_cx, label_cy,
                        text=f"T{global_tracker_idx+1}", font=('Helvetica', lbl_size), fill='#555555'
                    )
        
        # Draw devices (CB/SI)
        self._draw_devices()
        self._draw_device_info_panel()

        # Draw routes (behind pads)
        self._draw_routes()
        
        # Draw pads
        self._draw_pads()
        
        # Motor alignment line (if groups share a motor Y and alignment is on)
        if getattr(self, 'align_on_motor', False):
            self._draw_motor_alignment_lines()
        
        # Overlap warnings
        self._draw_overlap_warnings()

        # Scale bar
        self._draw_scale_bar()
        
        # Compass
        self.canvas.update_idletasks()
        cw = self.canvas.winfo_width()
        compass_x = cw - 30
        compass_y = 30
        arrow_len = 18
        
        self.canvas.create_line(
            compass_x, compass_y + arrow_len,
            compass_x, compass_y - arrow_len,
            fill='#333333', width=2, arrow='last'
        )
        self.canvas.create_text(
            compass_x, compass_y - arrow_len - 8,
            text='N', font=('Helvetica', 9, 'bold'), fill='#333333'
        )
    
    def _draw_motor_alignment_lines(self):
        """Draw a driveline across each group at its motor Y position,
        following the driveline angle if set."""
        for group_data in self.group_layout:
            motor_y = group_data.get('motor_y_ft', 0)
            if motor_y <= 0:
                continue
            
            overhang = self.max_tracker_width_ft * 0.5
            driveline_tan = group_data.get('driveline_tan', 0.0)
            
            left_x = group_data['x'] - overhang
            right_x = group_data['x'] + group_data['width_ft'] + overhang
            
            left_y = group_data['y'] + motor_y
            # Angle offset based on horizontal span from group origin
            right_y = left_y + (right_x - group_data['x']) * driveline_tan
            left_y = left_y + (-overhang) * driveline_tan
            
            x1, y1 = self.world_to_canvas(left_x, left_y)
            x2, y2 = self.world_to_canvas(right_x, right_y)
            
            self.canvas.create_line(
                x1, y1, x2, y2,
                fill='#FF8800', width=2, dash=(6, 3)
            )

    def _on_alignment_toggle(self):
        """Handle motor alignment checkbox toggle."""
        self.align_on_motor = self.align_motor_var.get()
        self.draw()

    def _on_inspect_toggle(self):
        """Handle inspect mode toggle."""
        self.selected_device_idx = None
        self.selected_pad_inspect_idx = None
        if self.inspect_mode:
            self.selected_group_idx = None
        self.draw()

    def _draw_toggle(self):
        """Draw the slider toggle switch on the canvas."""
        self.toggle_canvas.delete('all')
        w, h = 52, 24
        r = h // 2  # radius for rounded ends
        
        if self.inspect_mode:
            # ON state — green track
            track_color = '#4CAF50'
            knob_x = w - r
        else:
            # OFF state — gray track
            track_color = '#BDBDBD'
            knob_x = r
        
        # Draw rounded track
        self.toggle_canvas.create_oval(0, 0, h, h, fill=track_color, outline=track_color)
        self.toggle_canvas.create_oval(w - h, 0, w, h, fill=track_color, outline=track_color)
        self.toggle_canvas.create_rectangle(r, 0, w - r, h, fill=track_color, outline=track_color)
        
        # Draw knob
        knob_r = r - 2
        self.toggle_canvas.create_oval(
            knob_x - knob_r, 2, knob_x + knob_r, h - 2,
            fill='white', outline='#999999', width=1
        )
    
    def _on_toggle_click(self, event=None):
        """Handle click on the toggle switch."""
        if self.inspect_mode:
            # Switching back to Layout — confirm
            if not messagebox.askyesno("Switch Mode",
                                        "Switch back to Layout mode? Groups will be draggable again.",
                                        parent=self):
                return
        
        self.inspect_mode = not self.inspect_mode
        self.inspect_mode_var.set(self.inspect_mode)
        self._draw_toggle()
        self.toggle_label.config(
            text="Inspect" if self.inspect_mode else "Layout",
            foreground='#4CAF50' if self.inspect_mode else '#333333'
        )
        self._on_inspect_toggle()
    
    def hit_test_device(self, cx, cy):
        """Return the index of the device under canvas coords (cx, cy), or None."""
        if not hasattr(self, 'device_positions') or not self.device_positions:
            return None
        wx, wy = self.canvas_to_world(cx, cy)
        for i, dev in enumerate(self.device_positions):
            if (dev['x'] <= wx <= dev['x'] + dev['width_ft'] and
                dev['y'] <= wy <= dev['y'] + dev['height_ft']):
                return i
        return None
    
    def _reset_positions(self):
        """Reset all group positions to auto-layout and clear saved positions."""
        if not messagebox.askyesno("Reset Positions",
                                    "This will reset all group positions to the default layout. Continue?"):
            return
        
        for grp_idx, layout in enumerate(self.group_layout):
            group_spacing = self.max_tracker_length_ft * 0.1
            layout['x'] = grp_idx * (layout['width_ft'] + group_spacing)
            layout['y'] = 0
        
        self._update_world_bounds()
        self._save_group_positions()
        self.fit_and_redraw()

    def _refresh_allocation(self):
        """Re-run string allocation using current group positions, then refresh preview."""
        self._save_group_positions()
        
        parent = self.master
        if hasattr(parent, 'calculate_estimate'):
            # Clear stale combiner assignments so calculate_estimate rebuilds them fresh
            parent.last_combiner_assignments = []
            parent.allocation_locked = False
            parent.locked_allocation_result = None
            self.allocation_locked = False
            self._tracker_physical_order = None
            parent.calculate_estimate()
            
            inv_summary = getattr(parent, 'last_totals', {}).get('inverter_summary', {})
            if inv_summary and inv_summary.get('allocation_result'):
                self.inv_summary = inv_summary
                self.build_layout_data()
                self._build_legend()
                
                # Update top bar summary
                num_inv = inv_summary.get('total_inverters', 0)
                total_str = inv_summary.get('total_strings', 0)
                actual_ratio = inv_summary.get('actual_dc_ac', 0)
                split = inv_summary.get('total_split_trackers', 0)
                spatial_runs = inv_summary.get('allocation_result', {}).get('spatial_runs', 1)
                lock_str = "  |  🔒 LOCKED" if self.allocation_locked else ""
                self.summary_label.config(
                    text=f"{num_inv} Inverters  |  {total_str} Strings  |  DC:AC: {actual_ratio:.2f}  |  "
                         f"{split} Split Trackers  |  {spatial_runs} Run(s)  |  {self.topology}{lock_str}"
                )
                
                self.draw()

    def _toggle_allocation_lock(self):
        """Toggle the allocation lock on/off."""
        parent = self.master
        
        if self.allocation_locked:
            # Unlock
            self.allocation_locked = False
            parent.allocation_locked = False
            parent.locked_allocation_result = None
            self._update_lock_button()
        else:
            # Lock — snapshot the current allocation
            inv_summary = getattr(parent, 'last_totals', {}).get('inverter_summary', {})
            alloc = inv_summary.get('allocation_result')
            if not alloc:
                messagebox.showwarning(
                    "No Allocation",
                    "Run Refresh Allocation first to generate an allocation to lock.",
                    parent=self
                )
                return
            
            import copy
            self.allocation_locked = True
            parent.allocation_locked = True
            parent.locked_allocation_result = copy.deepcopy(alloc)
            self._update_lock_button()

    def _update_lock_button(self):
        """Update the lock button text and style to reflect current state."""
        if self.allocation_locked:
            self.lock_btn.config(text="🔒 Unlock Allocation")
        else:
            self.lock_btn.config(text="🔓 Lock Allocation")

    def _check_overlaps(self):
        """Check for overlapping groups and return list of overlapping pair indices."""
        overlaps = []
        
        for i in range(len(self.group_layout)):
            for j in range(i + 1, len(self.group_layout)):
                gi = self.group_layout[i]
                gj = self.group_layout[j]
                
                i_x1 = gi['x']
                i_x2 = gi['x'] + gi['width_ft']
                j_x1 = gj['x']
                j_x2 = gj['x'] + gj['width_ft']
                
                # Check X overlap first
                if i_x1 >= j_x2 or i_x2 <= j_x1:
                    continue
                
                # Driveline angle: Y bounds shift linearly with X
                i_tan = gi.get('driveline_tan', 0.0)
                j_tan = gj.get('driveline_tan', 0.0)
                i_vis_min = gi.get('visual_min_y_base', gi.get('visual_min_y', 0))
                i_vis_max = gi.get('visual_max_y_base', gi.get('visual_max_y', gi['length_ft']))
                j_vis_min = gj.get('visual_min_y_base', gj.get('visual_min_y', 0))
                j_vis_max = gj.get('visual_max_y_base', gj.get('visual_max_y', gj['length_ft']))
                
                # Check Y overlap at both ends of the X overlap region
                x_overlap_left = max(i_x1, j_x1)
                x_overlap_right = min(i_x2, j_x2)
                
                has_overlap = False
                for x_check in [x_overlap_left, x_overlap_right]:
                    i_y1 = gi['y'] + i_vis_min + (x_check - i_x1) * i_tan
                    i_y2 = gi['y'] + i_vis_max + (x_check - i_x1) * i_tan
                    j_y1 = gj['y'] + j_vis_min + (x_check - j_x1) * j_tan
                    j_y2 = gj['y'] + j_vis_max + (x_check - j_x1) * j_tan
                    
                    if i_y1 < j_y2 and i_y2 > j_y1:
                        has_overlap = True
                        break
                
                if has_overlap:
                    overlaps.append((i, j))
        
        return overlaps
    
    def _draw_overlap_warnings(self):
        """Draw red warning highlights around overlapping groups."""
        overlaps = self._check_overlaps()
        if not overlaps:
            return
        
        # Collect unique group indices that are involved in overlaps
        overlap_indices = set()
        for i, j in overlaps:
            overlap_indices.add(i)
            overlap_indices.add(j)
        
        max_width = getattr(self, 'max_tracker_width_ft', 6)
        pad = max_width * 0.3
        
        for idx in overlap_indices:
            g = self.group_layout[idx]
            vis_min = g.get('visual_min_y', 0)
            vis_max = g.get('visual_max_y', g['length_ft'])
            
            hx1, hy1 = self.world_to_canvas(g['x'] - pad, g['y'] + vis_min - pad)
            hx2, hy2 = self.world_to_canvas(
                g['x'] + g['width_ft'] + pad,
                g['y'] + vis_max + pad
            )
            
            self.canvas.create_rectangle(
                hx1, hy1, hx2, hy2,
                fill='', outline='#FF0000', width=3, dash=(8, 4),
                tags='overlap_warning'
            )
            
            # Warning label
            label_x = (hx1 + hx2) / 2
            label_y = hy1 - 8
            font_size = max(7, min(10, int(9 * self.scale)))
            self.canvas.create_text(
                label_x, label_y,
                text=f"⚠ Overlap", font=('Helvetica', font_size, 'bold'),
                fill='#FF0000', tags='overlap_warning'
            )

    def _draw_devices(self):
        """Draw combiner box / string inverter rectangles on the canvas."""
        if not hasattr(self, 'device_positions') or not self.device_positions:
            return
        
        # Build device_idx -> pad_idx lookup
        device_to_pad = {}
        if self.pads:
            for pad_idx, pad in enumerate(self.pads):
                for dev_idx in pad.get('assigned_devices', []):
                    device_to_pad[dev_idx] = pad_idx
        
        # Pad colors for device outlines (when pads exist)
        PAD_COLORS = ['#C62828', '#1565C0', '#2E7D32', '#E65100', '#6A1B9A',
                      '#00838F', '#AD1457', '#4E342E']
        
        for dev_idx, dev in enumerate(self.device_positions):
            dx = dev['x']
            dy = dev['y']
            dw = dev['width_ft']
            dh = dev['height_ft']
            label = dev['label']
            is_selected = (self.selected_device_idx == dev_idx)
            
            x1, y1 = self.world_to_canvas(dx, dy)
            x2, y2 = self.world_to_canvas(dx + dw, dy + dh)
            
            # Device fill color
            if self.device_label == 'CB':
                fill_color = '#FFB74D' if is_selected else '#FF9800'
            else:
                fill_color = '#64B5F6' if is_selected else '#2196F3'
            
            # Outline color: pad color if assigned, else default
            if dev_idx in device_to_pad and self.pads:
                pad_idx = device_to_pad[dev_idx]
                outline_color = PAD_COLORS[pad_idx % len(PAD_COLORS)]
                outline_width = 3
            else:
                outline_color = '#E65100' if self.device_label == 'CB' else '#0D47A1'
                outline_width = 3 if is_selected else 2
            
            if is_selected:
                outline_width = 4
            
            self.canvas.create_rectangle(
                x1, y1, x2, y2,
                fill=fill_color, outline=outline_color, width=outline_width
            )
            
            # Label — offset above device for readability
            cx = (x1 + x2) / 2
            font_size = max(7, min(14, int(10 * self.scale)))
            label_y = y1 - font_size - 2
            self.canvas.create_text(
                cx, label_y,
                text=label, font=('Helvetica', font_size, 'bold'),
                fill='#333333', anchor='s'
            )

    def _draw_device_info_panel(self):
        """Draw a draggable info panel near the selected device with a connector tail."""
        if not self.inspect_mode or self.selected_device_idx is None:
            return
        if not hasattr(self, 'device_positions') or self.selected_device_idx >= len(self.device_positions):
            return

        parent_qe = self.master
        dev = self.device_positions[self.selected_device_idx]
        dev_label = dev.get('label', f'{self.device_label}-{self.selected_device_idx + 1}')

        # Gather data from parent QE
        ns_offsets = getattr(parent_qe, '_tracker_ns_to_device', {})
        assignments = getattr(parent_qe, 'last_combiner_assignments', [])
        split_details = getattr(parent_qe, '_split_tracker_details', {})
        tracker_seg_map = getattr(parent_qe, '_tracker_to_segment', [])

        # Find CB assignment for this device
        cb_data = None
        for cb in assignments:
            if cb.get('device_idx') == self.selected_device_idx:
                cb_data = cb
                break
        if cb_data is None and self.selected_device_idx < len(assignments):
            cb_data = assignments[self.selected_device_idx]
        if cb_data is None:
            return

        inv_idx = cb_data.get('device_idx', self.selected_device_idx)

        # Build per-tracker detail
        tracker_info = {}  # tidx -> {strings, harnesses, whip_ew, ns_offset, ext_pos, ext_neg}
        for conn in cb_data.get('connections', []):
            tidx = conn['tracker_idx']
            if tidx not in tracker_info:
                tracker_info[tidx] = {
                    'strings': 0, 'harness_labels': [],
                    'whip_ew': 0, 'ns_offset': 0,
                    'ext_pos': [], 'ext_neg': [],
                }
            info = tracker_info[tidx]
            info['strings'] += conn['num_strings']
            info['harness_labels'].append(f"{conn.get('harness_label', '?')}({conn['num_strings']}S)")

        # Compute whip E-W per tracker from device position data
        dev_cx = dev['x'] + dev['width_ft'] / 2
        for tidx in tracker_info:
            # Find tracker world X
            global_idx = 0
            t_x = None
            for grp in self.group_layout:
                for t_i, t in enumerate(grp['trackers']):
                    if global_idx == tidx:
                        t_x = grp['x'] + t_i * self.tracker_pitch_ft
                        break
                    global_idx += 1
                if t_x is not None:
                    break
            if t_x is not None:
                tracker_info[tidx]['whip_ew'] = abs(t_x - dev_cx)

            # N-S inter-row offset
            tracker_info[tidx]['ns_offset'] = ns_offsets.get((tidx, inv_idx), 0)

            # Extender lengths
            if tidx < len(tracker_seg_map):
                seg_info = tracker_seg_map[tidx]
                seg = seg_info['seg']
                device_position = seg_info['device_position']
                ns_off = tracker_info[tidx]['ns_offset']

                signed_ns = ns_offsets.get((tidx, inv_idx), 0)
                if tidx in split_details:
                    for portion in split_details[tidx]['portions']:
                        if portion['inv_idx'] == inv_idx:
                            pairs = parent_qe.calculate_extender_lengths_per_segment(
                                seg, device_position, portion.get('start_pos', 0),
                                target_y_offset=signed_ns,
                                harness_sizes_override=portion['harnesses'])
                            for p_len, n_len in pairs:
                                tracker_info[tidx]['ext_pos'].append(
                                    parent_qe.round_whip_length(p_len))
                                tracker_info[tidx]['ext_neg'].append(
                                    parent_qe.round_whip_length(n_len))
                else:
                    pairs = parent_qe.calculate_extender_lengths_per_segment(
                        seg, device_position, target_y_offset=signed_ns)
                    for p_len, n_len in pairs:
                        tracker_info[tidx]['ext_pos'].append(
                            parent_qe.round_whip_length(p_len))
                        tracker_info[tidx]['ext_neg'].append(
                            parent_qe.round_whip_length(n_len))

        # Build row data for tabular layout
        headers = ['Trkr', 'Str', 'Whip', 'Ext+', 'Ext-']
        rows = []
        for tidx in sorted(tracker_info.keys()):
            info = tracker_info[tidx]
            whip_r = str(parent_qe.round_whip_length(info['whip_ew']) if info['whip_ew'] > 0 else 10)
            ext_p = '/'.join(str(v) for v in info['ext_pos']) if info['ext_pos'] else '--'
            ext_n = '/'.join(str(v) for v in info['ext_neg']) if info['ext_neg'] else '--'
            rows.append([f"T{tidx+1}", str(info['strings']), whip_r, ext_p, ext_n])

        # Calculate column widths from headers + data
        col_widths = [len(h) for h in headers]
        for row in rows:
            for i, val in enumerate(row):
                col_widths[i] = max(col_widths[i], len(val))

        # Build formatted lines
        def fmt_row(vals, widths):
            parts = []
            for i, val in enumerate(vals):
                if i == 0:
                    parts.append(val.ljust(widths[i]))
                else:
                    parts.append(val.rjust(widths[i]))
            return '  '.join(parts)

        lines = []
        lines.append(('header', f"{dev_label}  ({len(cb_data.get('connections', []))} inputs, {cb_data.get('breaker_size', '?')}A)"))
        lines.append(('spacer', ''))
        lines.append(('subheader', fmt_row(headers, col_widths)))

        for row in rows:
            lines.append(('row', fmt_row(row, col_widths)))

        # Totals
        total_whip = sum(parent_qe.round_whip_length(info['whip_ew'])
                         for info in tracker_info.values() if info['whip_ew'] > 0)
        lines.append(('spacer', ''))
        lines.append(('summary', f"Whips total: {total_whip}ft"))

        # Panel dimensions — width measured from actual font rendering
        import tkinter.font as tkfont
        font_size = 9
        line_height = font_size + 5
        pad = 10
        measure_font = tkfont.Font(family='Consolas', size=font_size + 1)
        max_text_px = max(measure_font.measure(text) for _, text in lines) if lines else 200
        panel_width = max_text_px + pad * 3
        panel_height = len(lines) * line_height + pad * 2

        # Anchor point on device (center-top)
        anchor_wx = dev['x'] + dev['width_ft'] / 2
        anchor_wy = dev['y']
        ax, ay = self.world_to_canvas(anchor_wx, anchor_wy)

        # Panel position — use stored drag offset or default above device
        if not hasattr(self, '_info_panel_offset'):
            self._info_panel_offset = {}
        off = self._info_panel_offset.get(self.selected_device_idx, (0, -panel_height - 30))
        px = ax + off[0] - panel_width // 2
        py = ay + off[1]

        # Clamp to canvas
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        px = max(5, min(px, cw - panel_width - 5))
        py = max(5, min(py, ch - panel_height - 5))

        # Store panel bounds for hit testing
        self._info_panel_bounds = (px, py, px + panel_width, py + panel_height)

        # Draw tail connector line from anchor to panel bottom-center
        panel_cx = px + panel_width // 2
        panel_bottom = py + panel_height
        self.canvas.create_line(
            ax, ay, panel_cx, panel_bottom,
            fill='#4444AA', width=2, dash=(4, 3)
        )
        # Small circle at anchor
        r = 4
        self.canvas.create_oval(ax - r, ay - r, ax + r, ay + r,
                                fill='#4444AA', outline='#6666CC', width=1)

        # Panel background with shadow
        self.canvas.create_rectangle(
            px + 3, py + 3, px + panel_width + 3, py + panel_height + 3,
            fill='#111111', outline='', width=0
        )
        self.canvas.create_rectangle(
            px, py, px + panel_width, py + panel_height,
            fill='#1E1E2E', outline='#4444AA', width=2
        )

        # Draw text lines
        for i, (style, text) in enumerate(lines):
            y_pos = py + pad + i * line_height
            if style == 'header':
                self.canvas.create_text(px + pad, y_pos, text=text,
                    font=('Consolas', font_size + 1, 'bold'), fill='#FFCC00', anchor='nw')
            elif style == 'subheader':
                self.canvas.create_text(px + pad, y_pos, text=text,
                    font=('Consolas', font_size, 'bold'), fill='#8888BB', anchor='nw')
            elif style == 'row':
                color = '#FF9944' if '+' in text.split()[-1] and text.split()[-1] != '--' else '#DDDDDD'
                self.canvas.create_text(px + pad, y_pos, text=text,
                    font=('Consolas', font_size), fill=color, anchor='nw')
            elif style == 'summary':
                self.canvas.create_text(px + pad, y_pos, text=text,
                    font=('Consolas', font_size, 'bold'), fill='#66BBFF', anchor='nw')

        # Drag handle indicator (top-right corner)
        hx = px + panel_width - 16
        hy = py + 4
        for row in range(3):
            for col in range(2):
                self.canvas.create_rectangle(
                    hx + col * 5, hy + row * 4,
                    hx + col * 5 + 3, hy + row * 4 + 2,
                    fill='#666688', outline='')

    def _draw_scale_bar(self):
        """Draw a scale bar in the bottom-left corner showing real-world distance."""
        if self.scale <= 0:
            return
        
        self.canvas.update_idletasks()
        ch = self.canvas.winfo_height()
        
        # Pick a nice round scale bar length in feet
        target_px = 120  # target pixel width for the bar
        target_ft = target_px / self.scale
        
        # Round to a nice number
        nice_values = [5, 10, 20, 25, 50, 100, 200, 250, 500, 1000]
        bar_ft = nice_values[0]
        for v in nice_values:
            if v <= target_ft:
                bar_ft = v
            else:
                break
        
        bar_px = bar_ft * self.scale
        
        x1 = 20
        y1 = ch - 25
        x2 = x1 + bar_px
        
        self.canvas.create_line(x1, y1, x2, y1, fill='#333333', width=2)
        self.canvas.create_line(x1, y1 - 5, x1, y1 + 5, fill='#333333', width=2)
        self.canvas.create_line(x2, y1 - 5, x2, y1 + 5, fill='#333333', width=2)
        
        self.canvas.create_text(
            (x1 + x2) / 2, y1 - 10,
            text=f"{bar_ft} ft", font=('Helvetica', 9), fill='#333333'
        )

    def _add_pad(self):
        """Enter pad placement mode — next click on canvas places a new pad."""
        if self.inspect_mode:
            messagebox.showinfo("Locked", "Switch to Layout mode to add pads.", parent=self)
            return
        
        self.placing_pad = True
        self.canvas.config(cursor='crosshair')
        self.add_pad_btn.config(state='disabled')
    
    def _place_pad_at(self, wx, wy):
        """Create a new pad at the given world coordinates."""
        pad_num = len(self.pads) + 1
        # Auto-assign all devices to the first pad if it's the only one
        if len(self.pads) == 0:
            all_device_indices = list(range(len(self.device_positions)))
        else:
            all_device_indices = []
        
        label_char = chr(ord('A') + (pad_num - 1) % 26)
        
        self.pads.append({
            'label': f"Pad {label_char}",
            'x': wx - 5,  # Center the 10ft-wide pad on click
            'y': wy - 4,  # Center the 8ft-tall pad on click
            'width_ft': 10.0,
            'height_ft': 8.0,
            'assigned_devices': all_device_indices,
        })
        
        self.placing_pad = False
        self.canvas.config(cursor='')
        self.add_pad_btn.config(state='normal')
        self.draw()
    
    def _draw_pads(self):
        """Draw inverter pad rectangles on the canvas."""
        if not self.pads:
            return
        
        PAD_COLORS = ['#C62828', '#1565C0', '#2E7D32', '#E65100', '#6A1B9A',
                      '#00838F', '#AD1457', '#4E342E']
        
        for pad_idx, pad in enumerate(self.pads):
            px = pad['x']
            py = pad['y']
            pw = pad.get('width_ft', 10.0)
            ph = pad.get('height_ft', 8.0)
            label = pad.get('label', f'Pad {pad_idx+1}')
            is_selected = (self.selected_pad_idx == pad_idx)
            
            x1, y1 = self.world_to_canvas(px, py)
            x2, y2 = self.world_to_canvas(px + pw, py + ph)
            
            base_color = PAD_COLORS[pad_idx % len(PAD_COLORS)]
            outline_width = 3 if is_selected else 2
            
            self.canvas.create_rectangle(
                x1, y1, x2, y2,
                fill=base_color, outline='#222222', width=outline_width
            )
            
            # Label
            cx = (x1 + x2) / 2
            cy = (y1 + y2) / 2
            font_size = max(6, min(10, int(8 * self.scale)))
            self.canvas.create_text(
                cx, cy,
                text=label, font=('Helvetica', font_size, 'bold'),
                fill='white'
            )
            
            # Device count subtitle
            num_assigned = len(pad.get('assigned_devices', []))
            if num_assigned > 0:
                sub_size = max(5, min(8, int(6 * self.scale)))
                self.canvas.create_text(
                    cx, cy + font_size + 2,
                    text=f"({num_assigned} devices)", font=('Helvetica', sub_size),
                    fill='#CCCCCC'
                )
    
    def hit_test_pad(self, cx, cy):
        """Return the index of the pad under canvas coords, or None."""
        if not self.pads:
            return None
        wx, wy = self.canvas_to_world(cx, cy)
        for i, pad in enumerate(self.pads):
            pw = pad.get('width_ft', 10.0)
            ph = pad.get('height_ft', 8.0)
            if (pad['x'] <= wx <= pad['x'] + pw and
                pad['y'] <= wy <= pad['y'] + ph):
                return i
        return None

    def hit_test_device(self, cx, cy):
        """Return the index of the device under canvas coords, or None."""
        if not hasattr(self, 'device_positions') or not self.device_positions:
            return None
        wx, wy = self.canvas_to_world(cx, cy)
        for i, dev in enumerate(self.device_positions):
            if (dev['x'] <= wx <= dev['x'] + dev['width_ft'] and
                dev['y'] <= wy <= dev['y'] + dev['height_ft']):
                return i
        return None
    
    def _show_assignment_dialog(self):
        """Show a dialog to assign devices to pads."""
        if not self.pads:
            messagebox.showinfo("No Pads", "Add at least one pad first.", parent=self)
            return
        
        if not hasattr(self, 'device_positions') or not self.device_positions:
            messagebox.showinfo("No Devices", "No devices to assign. Run Calculate Estimate first.", parent=self)
            return
        
        dialog = tk.Toplevel(self)
        dialog.title("Assign Devices to Pads")
        dialog.transient(self)
        # No grab_set() — allow pan/zoom on canvas behind dialog
        
        self.assigning_devices = True
        self.draw()
        
        # Size based on device count
        num_devices = len(self.device_positions)
        dialog_height = min(600, 120 + num_devices * 28)
        dialog.geometry(f"500x{dialog_height}")
        dialog.minsize(400, 200)
        

        # Instructions
        ttk.Label(dialog, text="Assign each device to a collection pad:",
                  font=('Helvetica', 10)).pack(anchor='w', padx=10, pady=(10, 0))
        ttk.Label(dialog, text="Tip: Drag the blue handle to fill multiple rows with the same pad",
                  font=('Helvetica', 8), foreground='gray').pack(anchor='w', padx=10, pady=(0, 5))
        
        # Build pad label list for dropdowns
        pad_labels = [pad.get('label', f'Pad {i+1}') for i, pad in enumerate(self.pads)]
        
        # Build reverse lookup: device_idx -> pad_idx
        device_to_pad = {}
        for pad_idx, pad in enumerate(self.pads):
            for dev_idx in pad.get('assigned_devices', []):
                device_to_pad[dev_idx] = pad_idx
        
        # Scrollable frame
        container = ttk.Frame(dialog)
        container.pack(fill='both', expand=True, padx=10, pady=5)
        
        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient='vertical', command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        
        scroll_frame.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=scroll_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')
        canvas.bind('<Enter>', lambda e: canvas.bind_all('<MouseWheel>', _on_mousewheel))
        canvas.bind('<Leave>', lambda e: canvas.unbind_all('<MouseWheel>'))
        
        # Headers
        header = ttk.Frame(scroll_frame)
        header.pack(fill='x', pady=(0, 5))
        ttk.Label(header, text="Device", font=('Helvetica', 9, 'bold'), width=12).pack(side='left', padx=5)
        ttk.Label(header, text="Strings", font=('Helvetica', 9, 'bold'), width=8).pack(side='left', padx=5)
        ttk.Label(header, text="Group", font=('Helvetica', 9, 'bold'), width=12).pack(side='left', padx=5)
        ttk.Label(header, text="Pad", font=('Helvetica', 9, 'bold'), width=12).pack(side='left', padx=5)
        ttk.Label(header, text="", width=1).pack(side='left', padx=2)
        
        ttk.Separator(scroll_frame, orient='horizontal').pack(fill='x', pady=2)
        
        # One row per device
        pad_vars = []
        pad_combos = []
        
        # Drag-fill handle state
        _fill = {'active': False, 'source_idx': None, 'value': None}
        _fill_handles = []
        
        def _handle_press(event, idx):
            """Start fill-drag from this row's handle"""
            _fill['active'] = True
            _fill['source_idx'] = idx
            _fill['value'] = pad_vars[idx].get()
            event.widget.config(cursor='sb_v_double_arrow')
        
        def _handle_motion(event):
            """Fill combos as drag passes over handles"""
            if not _fill['active']:
                return
            my = event.y_root
            src = _fill['source_idx']
            for i, handle in enumerate(_fill_handles):
                try:
                    hy = handle.winfo_rooty()
                    hh = handle.winfo_height()
                    row_center = hy + hh / 2
                    src_center = _fill_handles[src].winfo_rooty() + _fill_handles[src].winfo_height() / 2
                    in_range = min(my, src_center) <= row_center <= max(my, src_center)
                    if in_range:
                        pad_vars[i].set(_fill['value'])
                        pad_combos[i].focus_set()
                except tk.TclError:
                    pass
        
        def _handle_release(event):
            """End fill-drag"""
            if _fill['active']:
                event.widget.config(cursor='')
            _fill['active'] = False
            _fill['source_idx'] = None
            _fill['value'] = None

        for dev_idx, dev in enumerate(self.device_positions):
            row = ttk.Frame(scroll_frame)
            row.pack(fill='x', pady=1)
            
            # Device label
            ttk.Label(row, text=dev['label'], width=12).pack(side='left', padx=5)
            
            # String count
            num_strings = sum(len(v) for v in dev.get('assigned_strings', {}).values())
            ttk.Label(row, text=str(num_strings), width=8).pack(side='left', padx=5)
            
            # Group name
            grp_idx = dev.get('group_idx', 0)
            grp_name = self.groups[grp_idx]['name'] if grp_idx < len(self.groups) else '?'
            ttk.Label(row, text=grp_name, width=12).pack(side='left', padx=5)
            
            # Pad dropdown
            current_pad = device_to_pad.get(dev_idx, 0)
            if current_pad >= len(pad_labels):
                current_pad = 0
            var = tk.StringVar(value=pad_labels[current_pad])
            combo = ttk.Combobox(row, textvariable=var, values=pad_labels,
                                 state='readonly', width=12)
            combo.pack(side='left', padx=5)
            pad_vars.append(var)
            pad_combos.append(combo)
            
            # Fill handle — small draggable square
            handle = tk.Frame(row, width=10, height=18, bg='#4A90D9', cursor='sb_v_double_arrow')
            handle.pack(side='left', padx=(2, 0))
            handle.pack_propagate(False)
            handle.bind('<ButtonPress-1>', lambda e, i=dev_idx: _handle_press(e, i))
            handle.bind('<B1-Motion>', _handle_motion)
            handle.bind('<ButtonRelease-1>', _handle_release)
            _fill_handles.append(handle)
        
        # Buttons
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill='x', padx=10, pady=10)
        
        def _assign_all_to(pad_label):
            for var in pad_vars:
                var.set(pad_label)
        
        if len(self.pads) > 1:
            ttk.Label(btn_frame, text="Quick assign all:").pack(side='left', padx=(0, 5))
            for label in pad_labels:
                ttk.Button(btn_frame, text=label,
                           command=lambda l=label: _assign_all_to(l)).pack(side='left', padx=2)
        
        def _apply():
            # Rebuild pad assignments from dropdown values
            for pad in self.pads:
                pad['assigned_devices'] = []
            
            for dev_idx, var in enumerate(pad_vars):
                selected_label = var.get()
                for pad_idx, pad in enumerate(self.pads):
                    if pad.get('label', f'Pad {pad_idx+1}') == selected_label:
                        pad['assigned_devices'].append(dev_idx)
                        break
            
            self.assigning_devices = False
            self.draw()
            dialog.destroy()
        
        def _cancel():
            self.assigning_devices = False
            self.draw()
            dialog.destroy()
        
        dialog.protocol("WM_DELETE_WINDOW", _cancel)
        
        ttk.Button(btn_frame, text="Apply", command=_apply).pack(side='right', padx=(5, 0))
        ttk.Button(btn_frame, text="Cancel", command=_cancel).pack(side='right')
        
        # Center on parent
        dialog.update_idletasks()
        px = self.winfo_rootx()
        py = self.winfo_rooty()
        pw = self.winfo_width()
        ph = self.winfo_height()
        dw = dialog.winfo_width()
        dh = dialog.winfo_height()
        dialog.geometry(f"+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")

    def _normalize_to_device_strings(self):
        """Convert allocation/combiner data into per-string device format with physical positions.
        
        Returns:
            (device_data, metadata) or (None, None) if no data.
            
            device_data: list of dicts:
                {'name': str, 'strings': [(tracker_idx, physical_pos), ...]}
                where physical_pos is 0-based north-to-south position on the tracker.
            metadata: dict with 'module_isc', 'nec_factor', 'tracker_spt', 'source'
        """
        parent_qe = self.master
        
        module_isc = 0.0
        nec_factor = 1.56
        if hasattr(parent_qe, 'selected_module') and parent_qe.selected_module:
            module_isc = parent_qe.selected_module.isc
        
        tracker_spt = {}  # tracker_idx -> total strings on that tracker
        
        # We need to assign physical positions to strings.
        # Convention: within each tracker, devices are assigned north-to-south
        # in the order they appear in the allocation.
        # tracker_cursor tracks the next unassigned physical position per tracker.
        tracker_cursor = {}  # tracker_idx -> next physical position to assign
        
        # For CB topologies, prefer last_combiner_assignments
        assignments = getattr(parent_qe, 'last_combiner_assignments', [])
        
        if self.topology in ('Centralized String', 'Central Inverter') and assignments:
            # First pass: determine SPT for each tracker
            inv_summary = getattr(parent_qe, 'last_totals', {}).get('inverter_summary', {})
            alloc = inv_summary.get('allocation_result', {})
            for inv in alloc.get('inverters', []):
                for entry in inv.get('harness_map', []):
                    tracker_spt.setdefault(entry['tracker_idx'], entry['strings_per_tracker'])
            
            # Fallback: count from assignments
            if not tracker_spt:
                all_counts = {}
                for cb in assignments:
                    for conn in cb.get('connections', []):
                        tidx = conn['tracker_idx']
                        all_counts[tidx] = all_counts.get(tidx, 0) + conn['num_strings']
                tracker_spt = dict(all_counts)
            
            device_data = []
            tracker_cursor = {}
            for cb_idx, cb in enumerate(assignments):
                strings = []
                for conn in cb.get('connections', []):
                    tidx = conn['tracker_idx']
                    n = conn['num_strings']
                    # Use stored position if available
                    stored_start = conn.get('start_string_pos', None)
                    if stored_start is not None:
                        start = stored_start
                    else:
                        start = tracker_cursor.get(tidx, 0)
                    for p in range(start, start + n):
                        strings.append((tidx, p))
                    tracker_cursor[tidx] = max(tracker_cursor.get(tidx, 0), start + n)
                
                device_data.append({
                    'name': cb.get('combiner_name', f'CB-{cb_idx+1:02d}'),
                    'strings': strings,
                })
                if cb.get('module_isc'):
                    module_isc = cb['module_isc']
                if cb.get('nec_factor'):
                    nec_factor = cb['nec_factor']
            
            return device_data, {
                'module_isc': module_isc,
                'nec_factor': nec_factor,
                'tracker_spt': tracker_spt,
                'source': 'combiner',
            }
        
        # For Distributed String (or fallback): use allocation_result
        inv_summary = getattr(parent_qe, 'last_totals', {}).get('inverter_summary', {})
        alloc = inv_summary.get('allocation_result')
        
        if alloc and alloc.get('inverters'):
            device_data = []
            device_prefix = 'INV' if self.topology == 'Distributed String' else 'CB'
            tracker_cursor = {}
            
            for inv_idx, inv in enumerate(alloc['inverters']):
                strings = []
                for entry in inv.get('harness_map', []):
                    tidx = entry['tracker_idx']
                    tracker_spt.setdefault(tidx, entry['strings_per_tracker'])
                    if 'start_physical_pos' in entry:
                        start = entry['start_physical_pos']
                    else:
                        start = tracker_cursor.get(tidx, 0)
                    for p in range(start, start + entry['strings_taken']):
                        strings.append((tidx, p))
                    tracker_cursor[tidx] = max(tracker_cursor.get(tidx, 0), start + entry['strings_taken'])
                
                device_data.append({
                    'name': f'{device_prefix}-{inv_idx+1:02d}',
                    'strings': strings,
                })
            
            return device_data, {
                'module_isc': module_isc,
                'nec_factor': nec_factor,
                'tracker_spt': tracker_spt,
                'source': 'allocation',
            }
        
        return None, None

    def _rebuild_from_device_strings(self, device_data, metadata):
        """Rebuild allocation_result (and combiner assignments for CB topologies)
        from per-string device data with physical positions. Also locks the allocation.
        """
        parent_qe = self.master
        tracker_spt = metadata['tracker_spt']
        module_isc = metadata['module_isc']
        nec_factor = metadata['nec_factor']
        
        # --- Build allocation_result ---
        # Sort each device's strings by (tracker_idx, physical_pos)
        for dev in device_data:
            dev['strings'].sort(key=lambda s: (s[0], s[1]))
        
        inverters = []
        for dev_idx, dev in enumerate(device_data):
            harness_map = []
            tracker_indices = []
            pattern = []
            
            if not dev['strings']:
                inverters.append({
                    'pattern': [], 'tracker_indices': [], 'total_strings': 0,
                    'target_strings': 0, 'full_trackers': 0, 'split_trackers': 0,
                    'harness_map': [],
                })
                continue
            
            # Group consecutive same-tracker strings
            current_tidx = dev['strings'][0][0]
            current_start_pos = dev['strings'][0][1]
            current_count = 0
            
            for tidx, phys_pos in dev['strings']:
                if tidx == current_tidx:
                    current_count += 1
                else:
                    spt = tracker_spt.get(current_tidx, current_count)
                    harness_map.append({
                        'tracker_idx': current_tidx,
                        'strings_per_tracker': spt,
                        'strings_taken': current_count,
                        'is_split': current_count < spt,
                        'split_position': 'full',
                        'start_physical_pos': current_start_pos,
                    })
                    tracker_indices.append((current_tidx, current_count))
                    pattern.append(current_count)
                    current_tidx = tidx
                    current_start_pos = phys_pos
                    current_count = 1
            
            # Flush last group
            spt = tracker_spt.get(current_tidx, current_count)
            harness_map.append({
                'tracker_idx': current_tidx,
                'strings_per_tracker': spt,
                'strings_taken': current_count,
                'is_split': current_count < spt,
                'split_position': 'full',
                'start_physical_pos': current_start_pos,
            })
            tracker_indices.append((current_tidx, current_count))
            pattern.append(current_count)
            
            total = sum(pattern)
            inverters.append({
                'pattern': pattern,
                'tracker_indices': tracker_indices,
                'total_strings': total,
                'target_strings': total,
                'full_trackers': sum(1 for e in harness_map if not e['is_split']),
                'split_trackers': sum(1 for e in harness_map if e['is_split']),
                'harness_map': harness_map,
            })
        
        # Fix split_position across all inverters
        tracker_appearances = {}
        for inv_idx, inv in enumerate(inverters):
            for e_idx, entry in enumerate(inv['harness_map']):
                tidx = entry['tracker_idx']
                if tidx not in tracker_appearances:
                    tracker_appearances[tidx] = []
                tracker_appearances[tidx].append((inv_idx, e_idx))
        
        for tidx, appearances in tracker_appearances.items():
            if len(appearances) <= 1:
                continue
            for pos, (inv_idx, e_idx) in enumerate(appearances):
                entry = inverters[inv_idx]['harness_map'][e_idx]
                entry['is_split'] = True
                if pos == 0:
                    entry['split_position'] = 'head'
                elif pos == len(appearances) - 1:
                    entry['split_position'] = 'tail'
                else:
                    entry['split_position'] = 'middle'
        
        # Build summary
        total_strings = sum(inv['total_strings'] for inv in inverters)
        total_split = sum(1 for apps in tracker_appearances.values() if len(apps) > 1)
        inv_sizes = [inv['total_strings'] for inv in inverters if inv['total_strings'] > 0]
        max_spi = max(inv_sizes) if inv_sizes else 0
        min_spi = min(inv_sizes) if inv_sizes else 0
        
        tracker_type_counts = {}
        for spt in tracker_spt.values():
            tracker_type_counts[spt] = tracker_type_counts.get(spt, 0) + 1
        
        allocation_result = {
            'inverters': inverters,
            'summary': {
                'total_inverters': len(inv_sizes),
                'total_strings': total_strings,
                'total_trackers': len(tracker_spt),
                'total_split_trackers': total_split,
                'max_strings_per_inverter': max_spi,
                'min_strings_per_inverter': min_spi,
                'num_larger_inverters': sum(1 for s in inv_sizes if s == max_spi) if max_spi != min_spi else 0,
                'num_smaller_inverters': sum(1 for s in inv_sizes if s == min_spi) if max_spi != min_spi else 0,
                'tracker_type_counts': tracker_type_counts,
            },
        }
        
        # Write allocation to parent's last_totals
        if hasattr(parent_qe, 'last_totals') and parent_qe.last_totals:
            inv_summary = parent_qe.last_totals.get('inverter_summary', {})
            inv_summary['allocation_result'] = allocation_result
            inv_summary['total_inverters'] = allocation_result['summary']['total_inverters']
            inv_summary['total_strings'] = total_strings
            inv_summary['total_split_trackers'] = total_split
        
        # Lock the allocation
        parent_qe.allocation_locked = True
        parent_qe.locked_allocation_result = copy.deepcopy(allocation_result)
        self.allocation_locked = True
        
        # Store physical position ordering for draw code.
        # For each tracker, list device contributions sorted north-to-south
        # by minimum physical position that device owns.
        tracker_physical_order = {}
        for tidx in tracker_spt:
            dev_entries = []
            for dev_idx, dev in enumerate(device_data):
                positions = [p for t, p in dev['strings'] if t == tidx]
                if positions:
                    dev_entries.append((min(positions), dev_idx, len(positions)))
            dev_entries.sort()  # Sort by min physical position (northernmost first)
            tracker_physical_order[tidx] = [(dev_idx, count) for _, dev_idx, count in dev_entries]
        self._tracker_physical_order = tracker_physical_order
        
        # For CB topologies, also rebuild last_combiner_assignments
        if self.topology in ('Centralized String', 'Central Inverter'):
            BREAKER_SIZES = [100, 125, 150, 175, 200, 225, 250, 300, 350, 400, 450, 500, 600, 700, 800]
            new_assignments = []
            
            # Get harness config lookup from parent
            parent_qe = self.master
            
            for dev_idx, dev in enumerate(device_data):
                connections = []
                if dev['strings']:
                    # Group strings by tracker, preserving order
                    tracker_groups = []
                    current_tidx = dev['strings'][0][0]
                    current_positions = []
                    
                    for tidx, phys_pos in dev['strings']:
                        if tidx == current_tidx:
                            current_positions.append(phys_pos)
                        else:
                            tracker_groups.append((current_tidx, current_positions))
                            current_tidx = tidx
                            current_positions = [phys_pos]
                    tracker_groups.append((current_tidx, current_positions))
                    
                    # Build connections per harness (not per tracker)
                    for tidx, positions in tracker_groups:
                        spt = tracker_spt.get(tidx, len(positions))
                        start_pos = min(positions)
                        num_strings_here = len(positions)
                        
                        # Get the harness config for this tracker type
                        if hasattr(parent_qe, '_get_harness_config_for_tracker_type'):
                            full_harness_config = parent_qe._get_harness_config_for_tracker_type(spt)
                        else:
                            full_harness_config = [spt]
                        
                        # If this device owns ALL strings, use harness config directly
                        if num_strings_here == spt:
                            h_num = 1
                            pos_cursor = 0
                            for h_size in full_harness_config:
                                connections.append({
                                    'tracker_idx': tidx,
                                    'tracker_label': f'T{tidx+1:02d}',
                                    'harness_label': f'H{h_num:02d}',
                                    'num_strings': h_size,
                                    'module_isc': module_isc,
                                    'nec_factor': nec_factor,
                                    'wire_gauge': '10 AWG',
                                    'start_string_pos': pos_cursor,
                                })
                                pos_cursor += h_size
                                h_num += 1
                        else:
                            # Partial tracker — derive harnesses by walking config
                            # and finding overlap with our positions
                            harness_ranges = []
                            pos_cursor = 0
                            for h_size in full_harness_config:
                                harness_ranges.append((pos_cursor, pos_cursor + h_size - 1))
                                pos_cursor += h_size
                            
                            position_set = set(positions)
                            h_num = 1
                            for h_start, h_end in harness_ranges:
                                overlap = [p for p in range(h_start, h_end + 1) if p in position_set]
                                if overlap:
                                    connections.append({
                                        'tracker_idx': tidx,
                                        'tracker_label': f'T{tidx+1:02d}',
                                        'harness_label': f'H{h_num:02d}',
                                        'num_strings': len(overlap),
                                        'module_isc': module_isc,
                                        'nec_factor': nec_factor,
                                        'wire_gauge': '10 AWG',
                                        'start_string_pos': min(overlap),
                                    })
                                    h_num += 1
                
                total_current = sum(c['num_strings'] * module_isc * nec_factor for c in connections)
                calc_breaker = BREAKER_SIZES[-1]
                for bs in BREAKER_SIZES:
                    if bs >= total_current:
                        calc_breaker = bs
                        break
                
                new_assignments.append({
                    'combiner_name': dev['name'],
                    'device_idx': dev_idx,
                    'breaker_size': calc_breaker,
                    'module_isc': module_isc,
                    'nec_factor': nec_factor,
                    'connections': connections,
                })
            
            parent_qe.last_combiner_assignments = new_assignments

    def _show_edit_devices_dialog(self):
        """Show dialog to reassign string-level connections between devices."""
        device_data, metadata = self._normalize_to_device_strings()
        if not device_data:
            messagebox.showinfo("No Data", "Run Calculate Estimate first.", parent=self)
            return

        # Apply custom device names from canvas renaming
        for dev_idx, custom_name in self.device_names.items():
            if dev_idx < len(device_data):
                device_data[dev_idx]['name'] = custom_name

        original = copy.deepcopy(device_data)
        original_locked = self.allocation_locked
        original_parent_locked = getattr(self.master, 'allocation_locked', False)
        original_locked_result = copy.deepcopy(getattr(self.master, 'locked_allocation_result', None))
        # Snapshot full parent state for cancel restoration
        original_last_totals = copy.deepcopy(getattr(self.master, 'last_totals', None))
        original_combiner_assignments = copy.deepcopy(getattr(self.master, 'last_combiner_assignments', []))
        original_inv_summary = copy.deepcopy(self.inv_summary) if hasattr(self, 'inv_summary') else None
        original_physical_order = copy.deepcopy(getattr(self, '_tracker_physical_order', None))

        dialog = tk.Toplevel(self)
        dialog.title("Edit Device Assignments")
        dialog.geometry("400x650")
        dialog.transient(self)
        # No grab_set — allow pan/zoom on canvas behind dialog

        self._highlighted_strings = set()

        # Lookup: tree item iid -> tracker_idx
        string_tracker = {}

        def _regroup(dev):
            """Sort strings by (tracker_idx, physical_pos) to keep contiguous."""
            dev['strings'].sort(key=lambda s: (s[0], s[1]))

        def _sort_device_data():
            """Sort device_data in-place by natural name order (e.g. CB-02 before CB-10)."""
            import re
            def _natural_key(dev):
                return [int(c) if c.isdigit() else c.lower()
                        for c in re.split(r'(\d+)', dev['name'])]
            device_data.sort(key=_natural_key)

        def _validate_move(source_dev_idx, positions_to_move, target_dev_idx):
            """Check that moving strings preserves contiguity on both sides.
            
            Returns (ok, error_message).
            """
            # Group positions being moved by tracker
            moving_by_tracker = {}
            for pos_idx in positions_to_move:
                if pos_idx < len(device_data[source_dev_idx]['strings']):
                    tidx, phys = device_data[source_dev_idx]['strings'][pos_idx]
                    moving_by_tracker.setdefault(tidx, set()).add(phys)
            
            for tidx, moving_positions in moving_by_tracker.items():
                # Current positions this tracker has on source device
                source_positions = {p for t, p in device_data[source_dev_idx]['strings'] if t == tidx}
                remaining = source_positions - moving_positions
                
                # Check source contiguity after removal
                if len(remaining) > 1:
                    if max(remaining) - min(remaining) + 1 != len(remaining):
                        return False, (f"Cannot move: would leave a gap in T{tidx+1:02d} "
                                       f"on {device_data[source_dev_idx]['name']}")
                
                # Current positions this tracker has on target device
                target_positions = {p for t, p in device_data[target_dev_idx]['strings'] if t == tidx}
                combined = target_positions | moving_positions
                
                # Check target contiguity after addition
                if len(combined) > 1:
                    if max(combined) - min(combined) + 1 != len(combined):
                        return False, (f"Cannot move: would create a gap in T{tidx+1:02d} "
                                       f"on {device_data[target_dev_idx]['name']}")
            
            return True, ""

        # --- Tree ---
        tree_frame = ttk.Frame(dialog, padding="10")
        tree_frame.pack(fill='both', expand=True)

        tree = ttk.Treeview(tree_frame, selectmode='extended', show='tree')
        tree_scroll = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
        tree.config(yscrollcommand=tree_scroll.set)
        tree.pack(side='left', fill='both', expand=True)
        tree_scroll.pack(side='right', fill='y')

        # Drop target feedback
        drop_label_var = tk.StringVar(value="")
        drop_label = ttk.Label(dialog, textvariable=drop_label_var,
                               font=('Helvetica', 10, 'bold'), foreground='#1565C0')
        drop_label.pack(fill='x', padx=10)

        # Summary label
        summary_var = tk.StringVar()
        ttk.Label(dialog, textvariable=summary_var, foreground='gray',
                  font=('Helvetica', 9)).pack(fill='x', padx=10)

        def _update_summary():
            total_devs = len(device_data)
            total_strings = sum(len(d['strings']) for d in device_data)
            tracker_dev_map = {}
            for dev_idx, dev in enumerate(device_data):
                for tidx, _pos in dev['strings']:
                    tracker_dev_map.setdefault(tidx, set()).add(dev_idx)
            splits = sum(1 for devs in tracker_dev_map.values() if len(devs) > 1)
            summary_var.set(f"{total_devs} devices  |  {total_strings} strings  |  {splits} split trackers")

        # Holds a reference to any open inline-rename Entry widget so refresh_tree can clean it up
        inline_entry_ref = [None]

        def refresh_tree(select_device=None):
            # Destroy any open inline-rename entry before rebuilding
            if inline_entry_ref[0] is not None:
                try:
                    inline_entry_ref[0].destroy()
                except Exception:
                    pass
                inline_entry_ref[0] = None

            # Sort device_data in-place by natural name order before every rebuild
            import re as _sort_re
            # Resolve select_device to a name BEFORE sorting (indices shift after sort)
            select_name = None
            if select_device is not None and select_device < len(device_data):
                select_name = device_data[select_device]['name']

            device_data.sort(key=lambda d: [int(c) if c.isdigit() else c.lower()
                                            for c in _sort_re.split(r'(\d+)', d['name'])])

            # If we had a name to select, find its new index
            if select_name is not None:
                select_device = next((i for i, d in enumerate(device_data) if d['name'] == select_name), select_device)

            # Capture current open/closed state before rebuilding
            open_state = {}
            for child in tree.get_children():
                dev_idx_str = child.replace('dev_', '')
                if dev_idx_str.isdigit():
                    open_state[int(dev_idx_str)] = tree.item(child, 'open')

            tree.delete(*tree.get_children())
            string_tracker.clear()
            for dev_idx, dev in enumerate(device_data):
                total = len(dev['strings'])
                dev_iid = f'dev_{dev_idx}'
                # Preserve previous open state; default to open=(total <= 40) for new nodes
                is_open = open_state.get(dev_idx, False)
                tree.insert('', 'end', 
                            iid=dev_iid,
                            text=f"{dev['name']}  ({total} strings)",
                            open=is_open)

                for s_pos, (tidx, phys_pos) in enumerate(dev['strings']):
                    str_iid = f'dev_{dev_idx}_s_{s_pos}'
                    tree.insert(dev_iid, 'end', iid=str_iid,
                                text=f"  T{tidx+1:02d} - S{phys_pos+1:02d}")
                    string_tracker[str_iid] = (tidx, phys_pos)

            if select_device is not None:
                dev_iid = f'dev_{select_device}'
                if dev_iid in tree.get_children():
                    tree.selection_set(dev_iid)
                    tree.see(dev_iid)

            _update_summary()

        # --- Selection → highlight on canvas ---
        def on_tree_select(event=None):
            if drag_state.get('suppressing'):
                return
            selected = [iid for iid in tree.selection() if iid in string_tracker]
            self._highlighted_strings = set()

            for str_iid in selected:
                tidx, phys_pos = string_tracker[str_iid]
                self._highlighted_strings.add((tidx, phys_pos))

            self.draw()

        tree.bind('<<TreeviewSelect>>', on_tree_select)

        # --- Drag-and-drop state ---
        drag_state = {
            'active': False,
            'start_x': 0,
            'start_y': 0,
            'threshold': 5,
            'items': [],
            'suppressing': False,
            'last_target_dev': None,
        }

        # Tag for highlighting the drop target device row
        tree.tag_configure('drop_target', background='#BBDEFB')

        def _resolve_device_iid(y):
            """Given a y coordinate, resolve to the device iid the user is hovering over."""
            iid = tree.identify_row(y)
            if not iid:
                return None
            # If it's a string child, go to parent device
            if iid in string_tracker:
                iid = tree.parent(iid)
            if iid and iid.startswith('dev_'):
                return iid
            return None

        def _on_press(event):
            drag_state['active'] = False
            drag_state['suppressing'] = False
            drag_state['last_target_dev'] = None
            drag_state['items'] = []          # ← Clear stale items immediately
            drag_state['start_x'] = event.x
            drag_state['start_y'] = event.y
            # Capture selection NOW, before Tkinter's default handler changes it
            clicked_iid = tree.identify_row(event.y)
            current_sel = [iid for iid in tree.selection() if iid in string_tracker]
            if clicked_iid in current_sel:
                # Clicked on an already-selected item — drag the whole selection
                drag_state['items'] = list(current_sel)
            else:
                # Clicked on a new item — let default handler update selection, then capture
                def _capture():
                    drag_state['items'] = [iid for iid in tree.selection() if iid in string_tracker]
                tree.after_idle(_capture)

        def _on_motion(event):
            if not drag_state['active']:
                dx = abs(event.x - drag_state['start_x'])
                dy = abs(event.y - drag_state['start_y'])
                if dx > drag_state['threshold'] or dy > drag_state['threshold']:
                    if drag_state['items']:
                        drag_state['active'] = True
                        drag_state['suppressing'] = True
                        tree.config(cursor='hand2')
                        n = len(drag_state['items'])
                        drop_label_var.set(f"Dragging {n} string{'s' if n != 1 else ''}...")

            if drag_state['active']:
                dev_iid = _resolve_device_iid(event.y)

                # Clear previous highlight
                if drag_state['last_target_dev'] and drag_state['last_target_dev'] != dev_iid:
                    try:
                        old_tags = list(tree.item(drag_state['last_target_dev'], 'tags'))
                        if 'drop_target' in old_tags:
                            old_tags.remove('drop_target')
                        tree.item(drag_state['last_target_dev'], tags=old_tags)
                    except Exception:
                        pass

                if dev_iid:
                    dev_idx = int(dev_iid.split('_')[1])
                    dev_name = device_data[dev_idx]['name']

                    # Check if any source is different from target
                    source_devs = set()
                    for s_iid in drag_state['items']:
                        source_devs.add(tree.parent(s_iid))

                    if dev_iid in source_devs and len(source_devs) == 1:
                        drop_label_var.set(f"⚠ Cannot drop on source device ({dev_name})")
                    else:
                        drop_label_var.set(f"▶ Drop onto: {dev_name}")
                        try:
                            cur_tags = list(tree.item(dev_iid, 'tags'))
                            if 'drop_target' not in cur_tags:
                                cur_tags.append('drop_target')
                            tree.item(dev_iid, tags=cur_tags)
                        except Exception:
                            pass

                    drag_state['last_target_dev'] = dev_iid
                else:
                    drop_label_var.set("Dragging... (hover over a device to drop)")
                    drag_state['last_target_dev'] = None

        def _on_release(event):
            was_active = drag_state['active']
            drag_state['active'] = False
            drag_state['suppressing'] = False
            tree.config(cursor='')
            drop_label_var.set("")

            # Clear drop target highlight
            if drag_state['last_target_dev']:
                try:
                    old_tags = list(tree.item(drag_state['last_target_dev'], 'tags'))
                    if 'drop_target' in old_tags:
                        old_tags.remove('drop_target')
                    tree.item(drag_state['last_target_dev'], tags=old_tags)
                except Exception:
                    pass
            drag_state['last_target_dev'] = None

            if not was_active:
                return

            dev_iid = _resolve_device_iid(event.y)
            if not dev_iid:
                return

            target_dev = int(dev_iid.split('_')[1])

            selected = drag_state['items']
            if not selected:
                return

            # Group by source device
            by_source = {}
            for s_iid in selected:
                parent_iid = tree.parent(s_iid)
                if not parent_iid:
                    continue
                dev_idx = int(parent_iid.split('_')[1])
                s_pos = int(s_iid.rsplit('_', 1)[1])
                by_source.setdefault(dev_idx, []).append(s_pos)

            # Validate contiguity for each source
            for src_dev, positions in by_source.items():
                if src_dev == target_dev:
                    continue
                ok, err = _validate_move(src_dev, positions, target_dev)
                if not ok:
                    messagebox.showwarning("Invalid Move", err, parent=dialog)
                    drag_state['items'] = []
                    return

            moved = []
            for src_dev, positions in by_source.items():
                if src_dev == target_dev:
                    continue
                for pos in sorted(positions, reverse=True):
                    if pos < len(device_data[src_dev]['strings']):
                        entry = device_data[src_dev]['strings'].pop(pos)
                        moved.append(entry)
            if not moved:
                return

            device_data[target_dev]['strings'].extend(reversed(moved))
            _regroup(device_data[target_dev])

            refresh_tree(select_device=target_dev)
            _update_live_preview()
            drag_state['items'] = []
        tree.bind('<ButtonPress-1>', _on_press, add='+')
        tree.bind('<B1-Motion>', _on_motion)
        tree.bind('<ButtonRelease-1>', _on_release, add='+')

        # --- Inline rename on double-click ---
        def _begin_inline_rename(iid):
            if inline_entry_ref[0] is not None:
                try:
                    inline_entry_ref[0].destroy()
                except Exception:
                    pass
                inline_entry_ref[0] = None

            dev_idx = int(iid.split('_')[1])
            tree.see(iid)
            tree.update_idletasks()
            bbox = tree.bbox(iid, column='#0')
            if not bbox:
                return
            x, y, w, h = bbox
            entry_var = tk.StringVar(value=device_data[dev_idx]['name'])
            entry = ttk.Entry(tree, textvariable=entry_var)
            entry.place(x=x + 4, y=y, width=max(w - 8, 120), height=h)
            entry.select_range(0, 'end')
            entry.focus_set()
            inline_entry_ref[0] = entry

            def _commit(event=None):
                name = entry_var.get().strip()
                if name:
                    device_data[dev_idx]['name'] = name
                if inline_entry_ref[0] is entry:
                    inline_entry_ref[0] = None
                try:
                    entry.destroy()
                except Exception:
                    pass
                refresh_tree(select_device=dev_idx)
                _update_live_preview()

            def _cancel(event=None):
                if inline_entry_ref[0] is entry:
                    inline_entry_ref[0] = None
                try:
                    entry.destroy()
                except Exception:
                    pass

            entry.bind('<Return>', _commit)
            entry.bind('<FocusOut>', _commit)
            entry.bind('<Escape>', _cancel)

        def _on_double_click(event):
            iid = tree.identify_row(event.y)
            if iid and iid.startswith('dev_') and iid not in string_tracker:
                _begin_inline_rename(iid)
                return 'break'

        tree.bind('<Double-Button-1>', _on_double_click)

        # --- Action buttons ---
        action_frame = ttk.Frame(dialog, padding="5 0")
        action_frame.pack(fill='x', padx=10)

        def add_device():
            import re as _re
            default_prefix = 'INV-' if self.topology == 'Distributed String' else 'CB-'

            # Find the next available sequential number for the default prefix
            existing_nums = []
            for d in device_data:
                m = _re.match(rf'^{_re.escape(default_prefix)}(\d+)$', d['name'])
                if m:
                    existing_nums.append(int(m.group(1)))
            next_num = (max(existing_nums) + 1) if existing_nums else (len(device_data) + 1)

            add_dlg = tk.Toplevel(dialog)
            add_dlg.title("Add Device(s)")
            add_dlg.transient(dialog)
            add_dlg.resizable(False, False)

            frm = ttk.Frame(add_dlg, padding=12)
            frm.pack(fill='both', expand=True)

            ttk.Label(frm, text="Count:").grid(row=0, column=0, sticky='w', padx=(0, 8), pady=4)
            count_var = tk.IntVar(value=1)
            ttk.Spinbox(frm, from_=1, to=99, textvariable=count_var, width=6).grid(row=0, column=1, sticky='w')

            ttk.Label(frm, text="Prefix:").grid(row=1, column=0, sticky='w', padx=(0, 8), pady=4)
            prefix_var = tk.StringVar(value=default_prefix)
            ttk.Entry(frm, textvariable=prefix_var, width=12).grid(row=1, column=1, sticky='w')

            ttk.Label(frm, text="Start #:").grid(row=2, column=0, sticky='w', padx=(0, 8), pady=4)
            start_var = tk.IntVar(value=next_num)
            ttk.Spinbox(frm, from_=1, to=999, textvariable=start_var, width=6).grid(row=2, column=1, sticky='w')

            btn_frame = ttk.Frame(frm)
            btn_frame.grid(row=3, column=0, columnspan=2, pady=(10, 0))

            def _do_add():
                try:
                    count = int(count_var.get())
                    start = int(start_var.get())
                except (ValueError, tk.TclError):
                    return
                prefix = prefix_var.get()
                last_name = None
                for n in range(start, start + count):
                    last_name = f'{prefix}{n:02d}'
                    device_data.append({'name': last_name, 'strings': []})
                _sort_device_data()
                add_dlg.destroy()
                new_idx = next((i for i, d in enumerate(device_data) if d['name'] == last_name), len(device_data) - 1)
                refresh_tree(select_device=new_idx)
                _update_live_preview()

            ttk.Button(btn_frame, text="Add", command=_do_add).pack(side='left', padx=4)
            ttk.Button(btn_frame, text="Cancel", command=add_dlg.destroy).pack(side='left', padx=4)

            add_dlg.bind('<Return>', lambda e: _do_add())
            add_dlg.bind('<Escape>', lambda e: add_dlg.destroy())

            add_dlg.update_idletasks()
            x = dialog.winfo_rootx() + (dialog.winfo_width() - add_dlg.winfo_width()) // 2
            y = dialog.winfo_rooty() + (dialog.winfo_height() - add_dlg.winfo_height()) // 2
            add_dlg.geometry(f'+{x}+{y}')
            add_dlg.grab_set()
            add_dlg.focus_set()

        def remove_device():
            sel = tree.selection()
            dev_indices = set()
            for iid in sel:
                if iid.startswith('dev_') and iid not in string_tracker:
                    dev_indices.add(int(iid.split('_')[1]))
            if not dev_indices:
                messagebox.showinfo("Select Device",
                                    "Select a device node to remove.", parent=dialog)
                return
            for idx in sorted(dev_indices, reverse=True):
                if device_data[idx]['strings']:
                    if not messagebox.askyesno(
                        "Remove Device",
                        f"'{device_data[idx]['name']}' has "
                        f"{len(device_data[idx]['strings'])} strings.\n"
                        "They will be lost. Continue?", parent=dialog):
                        continue
                device_data.pop(idx)
            refresh_tree()
            _update_live_preview()

        def rename_device():
            import re as _re
            sel = [iid for iid in tree.selection()
                   if iid.startswith('dev_') and iid not in string_tracker]
            if not sel:
                messagebox.showinfo("Select Device",
                                    "Select a device node to rename.", parent=dialog)
                return

            if len(sel) == 1:
                dev_idx = int(sel[0].split('_')[1])
                new_name = simpledialog.askstring(
                    "Rename", f"New name for {device_data[dev_idx]['name']}:",
                    initialvalue=device_data[dev_idx]['name'], parent=dialog)
                if new_name and new_name.strip():
                    import re as _re2
                    clean_name = new_name.strip()
                    device_data[dev_idx]['name'] = clean_name
                    device_data.sort(key=lambda d: [int(c) if c.isdigit() else c.lower() for c in _re2.split(r'(\d+)', d['name'])])
                    sorted_idx = next(i for i, d in enumerate(device_data) if d['name'] == clean_name)
                    refresh_tree(select_device=sorted_idx)
                    _update_live_preview()
            else:
                # Bulk rename — put devices in tree display order
                tree_order = [iid for iid in tree.get_children() if iid in sel]
                dev_indices = [int(iid.split('_')[1]) for iid in tree_order]

                # Guess a common prefix by stripping trailing digits from shared prefix
                names = [device_data[i]['name'] for i in dev_indices]
                common = names[0]
                for name in names[1:]:
                    common = common[:sum(1 for a, b in zip(common, name) if a == b)]
                common = _re.sub(r'\d+$', '', common)

                ren_dlg = tk.Toplevel(dialog)
                ren_dlg.title(f"Bulk Rename {len(sel)} Devices")
                ren_dlg.transient(dialog)
                ren_dlg.resizable(False, False)

                frm = ttk.Frame(ren_dlg, padding=12)
                frm.pack(fill='both', expand=True)

                ttk.Label(frm, text=f"Rename {len(sel)} devices sequentially:").grid(
                    row=0, column=0, columnspan=2, sticky='w', pady=(0, 8))

                ttk.Label(frm, text="Prefix:").grid(row=1, column=0, sticky='w', padx=(0, 8), pady=4)
                prefix_var = tk.StringVar(value=common or 'CB-')
                ttk.Entry(frm, textvariable=prefix_var, width=14).grid(row=1, column=1, sticky='w')

                ttk.Label(frm, text="Start #:").grid(row=2, column=0, sticky='w', padx=(0, 8), pady=4)
                start_var = tk.IntVar(value=1)
                ttk.Spinbox(frm, from_=1, to=999, textvariable=start_var, width=6).grid(row=2, column=1, sticky='w')

                btn_frame = ttk.Frame(frm)
                btn_frame.grid(row=3, column=0, columnspan=2, pady=(10, 0))

                def _do_rename():
                    try:
                        start = int(start_var.get())
                    except (ValueError, tk.TclError):
                        return
                    prefix = prefix_var.get()
                    for n, dev_idx in enumerate(dev_indices):
                        device_data[dev_idx]['name'] = f'{prefix}{start + n:02d}'
                    _sort_device_data()
                    ren_dlg.destroy()
                    refresh_tree()
                    _update_live_preview()

                ttk.Button(btn_frame, text="Rename", command=_do_rename).pack(side='left', padx=4)
                ttk.Button(btn_frame, text="Cancel", command=ren_dlg.destroy).pack(side='left', padx=4)

                ren_dlg.bind('<Return>', lambda e: _do_rename())
                ren_dlg.bind('<Escape>', lambda e: ren_dlg.destroy())

                ren_dlg.update_idletasks()
                x = dialog.winfo_rootx() + (dialog.winfo_width() - ren_dlg.winfo_width()) // 2
                y = dialog.winfo_rooty() + (dialog.winfo_height() - ren_dlg.winfo_height()) // 2
                ren_dlg.geometry(f'+{x}+{y}')
                ren_dlg.grab_set()
                ren_dlg.focus_set()

        ttk.Label(action_frame, text="Drag strings to move  |",
                  foreground='gray', font=('Helvetica', 9)).pack(side='left', padx=(0, 8))
        ttk.Button(action_frame, text="Add Device(s)", command=add_device).pack(side='left', padx=2)
        ttk.Button(action_frame, text="Remove", command=remove_device).pack(side='left', padx=2)
        ttk.Button(action_frame, text="Rename", command=rename_device).pack(side='left', padx=2)

        # --- Live preview update ---
        def _update_live_preview():
            # Sync device count, names, and rebuild
            self.num_devices = len(device_data)
            self.device_names = {i: dev['name'] for i, dev in enumerate(device_data)}

            self._rebuild_from_device_strings(device_data, metadata)

            inv_summary = getattr(self.master, 'last_totals', {}).get('inverter_summary', {})
            if inv_summary and inv_summary.get('allocation_result'):
                self.inv_summary = inv_summary
                self.build_layout_data()
                self._recolor_from_cb_assignments()
                self._build_legend()
                self.draw()

        # --- Apply / Cancel ---
        bottom_frame = ttk.Frame(dialog, padding="10")
        bottom_frame.pack(fill='x')

        def apply_changes():
            self.num_devices = len(device_data)
            self.device_names = {i: dev['name'] for i, dev in enumerate(device_data)}
            self._rebuild_from_device_strings(device_data, metadata)
            self._update_lock_button()
            self._highlighted_strings = set()

            inv_summary = getattr(self.master, 'last_totals', {}).get('inverter_summary', {})
            if inv_summary and inv_summary.get('allocation_result'):
                self.inv_summary = inv_summary

            self.build_layout_data()
            self._recolor_from_cb_assignments()
            self._build_legend()

            # Update top bar summary
            alloc = inv_summary.get('allocation_result', {})
            num_inv = alloc.get('summary', {}).get('total_inverters', 0)
            total_str = alloc.get('summary', {}).get('total_strings', 0)
            split = alloc.get('summary', {}).get('total_split_trackers', 0)
            spatial_runs = alloc.get('spatial_runs', 1)
            actual_ratio = inv_summary.get('actual_dc_ac', 0)
            self.summary_label.config(
                text=f"{num_inv} Devices  |  {total_str} Strings  |  DC:AC: {actual_ratio:.2f}  |  "
                     f"{split} Split Trackers  |  {spatial_runs} Run(s)  |  {self.topology}  |  🔒 LOCKED"
            )

            self.draw()
            dialog.destroy()

        def cancel():
            # Restore full parent state from snapshot — no rebuild needed
            self.allocation_locked = original_locked
            self.master.allocation_locked = original_parent_locked
            self.master.locked_allocation_result = original_locked_result
            self.master.last_totals = original_last_totals
            self.master.last_combiner_assignments = original_combiner_assignments
            self._tracker_physical_order = original_physical_order
            self._update_lock_button()

            self._highlighted_strings = set()

            if original_inv_summary is not None:
                self.inv_summary = original_inv_summary
            self.build_layout_data()
            self._recolor_from_cb_assignments()
            self._build_legend()
            self.draw()
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", cancel)

        ttk.Button(bottom_frame, text="Apply", command=apply_changes).pack(side='right', padx=2)
        ttk.Button(bottom_frame, text="Cancel", command=cancel).pack(side='right', padx=2)
        ttk.Button(bottom_frame, text="Collapse All",
                   command=lambda: [tree.item(c, open=False) for c in tree.get_children()]).pack(side='left', padx=2)
        ttk.Button(bottom_frame, text="Expand All",
                   command=lambda: [tree.item(c, open=True) for c in tree.get_children()]).pack(side='left', padx=2)

        # Initial tree
        refresh_tree()

        # Center on parent
        dialog.update_idletasks()
        px = self.winfo_rootx()
        py = self.winfo_rooty()
        pw = self.winfo_width()
        ph = self.winfo_height()
        dw = dialog.winfo_width()
        dh = dialog.winfo_height()
        dialog.geometry(f"+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")

    def _on_pad_right_click(self, event):
        """Show context menu for pads."""
        hit = self.hit_test_pad(event.x, event.y)
        if hit is None:
            return
        
        self.selected_pad_idx = hit
        self.draw()
        
        menu = tk.Menu(self, tearoff=0)
        
        def _rename():
            from tkinter import simpledialog
            current = self.pads[hit].get('label', f'Pad {hit+1}')
            new_name = simpledialog.askstring("Rename Pad", "New label:", 
                                              initialvalue=current, parent=self)
            if new_name and new_name.strip():
                self.pads[hit]['label'] = new_name.strip()
                self.draw()
        
        def _delete():
            label = self.pads[hit].get('label', f'Pad {hit+1}')
            if not messagebox.askyesno("Delete Pad", f"Delete '{label}'?", parent=self):
                return
            
            # Reassign devices to first remaining pad if any
            orphaned = self.pads[hit].get('assigned_devices', [])
            del self.pads[hit]
            
            if self.pads and orphaned:
                self.pads[0]['assigned_devices'] = list(
                    set(self.pads[0].get('assigned_devices', []) + orphaned)
                )
            
            # Fix device indices in remaining pads (indices > hit shift down)
            # Not needed — pad indices don't change, only the list position
            
            self.selected_pad_idx = None
            self.draw()
        
        menu.add_command(label="Rename", command=_rename)
        menu.add_command(label="Delete", command=_delete)
        menu.tk_popup(event.x_root, event.y_root)

    def _draw_routes(self):
        """Draw L-shaped Manhattan routes from each device to its assigned pad."""
        if not self.show_routes_var.get():
            return
        if not self.pads or not hasattr(self, 'device_positions') or not self.device_positions:
            return
        
        PAD_COLORS = ['#C62828', '#1565C0', '#2E7D32', '#E65100', '#6A1B9A',
                      '#00838F', '#AD1457', '#4E342E']
        
        # Build device -> pad lookup
        device_to_pad = {}
        for pad_idx, pad in enumerate(self.pads):
            for dev_idx in pad.get('assigned_devices', []):
                device_to_pad[dev_idx] = pad_idx
        
        for dev_idx, dev in enumerate(self.device_positions):
            pad_idx = device_to_pad.get(dev_idx)
            if pad_idx is None or pad_idx >= len(self.pads):
                continue
            
            pad = self.pads[pad_idx]
            
            # Device center
            dev_cx = dev['x'] + dev['width_ft'] / 2
            dev_cy = dev['y'] + dev['height_ft'] / 2
            
            # Pad center
            pad_cx = pad['x'] + pad.get('width_ft', 10.0) / 2
            pad_cy = pad['y'] + pad.get('height_ft', 8.0) / 2
            
            # L-shaped route: go E-W first, then N-S
            corner_x = pad_cx
            corner_y = dev_cy
            
            # Convert to canvas coords
            cx1, cy1 = self.world_to_canvas(dev_cx, dev_cy)
            cx_corner, cy_corner = self.world_to_canvas(corner_x, corner_y)
            cx2, cy2 = self.world_to_canvas(pad_cx, pad_cy)
            
            color = PAD_COLORS[pad_idx % len(PAD_COLORS)]
            
            # Determine line style based on topology
            if self.topology == 'Distributed String':
                dash_pattern = (4, 4)  # Dashed for AC
            else:
                dash_pattern = ()  # Solid for DC
            
            line_width = 1
            
            # Draw E-W leg
            self.canvas.create_line(
                cx1, cy1, cx_corner, cy_corner,
                fill=color, width=line_width, dash=dash_pattern
            )
            
            # Draw N-S leg
            self.canvas.create_line(
                cx_corner, cy_corner, cx2, cy2,
                fill=color, width=line_width, dash=dash_pattern
            )

            # Show distance label if this device is selected in inspect mode
            if self.inspect_mode and self.selected_device_idx == dev_idx:
                ew_dist = abs(dev_cx - pad_cx)
                ns_dist = abs(dev_cy - pad_cy)
                total_dist = ew_dist + ns_dist  # Manhattan
                
                # Place label at the corner of the L
                font_size = max(6, min(10, int(8 * self.scale)))
                self.canvas.create_text(
                    cx_corner, cy_corner - 8,
                    text=f"{total_dist:.0f} ft",
                    font=('Helvetica', font_size, 'bold'),
                    fill=color
                )

    def _recolor_from_cb_assignments(self):
        """Recolor tracker assignments from parent's last_combiner_assignments.
        
        Called after build_layout_data() to override the default inverter-based
        coloring with CB-based coloring when manual CB edits exist.
        
        If _tracker_physical_order is available (from Edit Devices), uses
        physical position ordering so that the device owning the northernmost
        strings on a tracker gets painted first.
        """
        parent_qe = self.master
        assignments = getattr(parent_qe, 'last_combiner_assignments', [])
        physical_order = getattr(self, '_tracker_physical_order', None)
        
        # If no cached physical order, compute it from device string data
        if physical_order is None:
            device_data, metadata = self._normalize_to_device_strings()
            if device_data and metadata:
                tracker_spt = metadata.get('tracker_spt', {})
                physical_order = {}
                for tidx in tracker_spt:
                    dev_entries = []
                    for dev_idx, dev in enumerate(device_data):
                        positions = [p for t, p in dev['strings'] if t == tidx]
                        if positions:
                            dev_entries.append((min(positions), dev_idx, len(positions)))
                    dev_entries.sort()
                    physical_order[tidx] = [(dev_idx, count) for _, dev_idx, count in dev_entries]
                self._tracker_physical_order = physical_order
        
        # Physical ordering takes priority — works for all topologies
        if physical_order:
            # Use physical ordering — most accurate after manual edits
            global_idx = 0
            for group_data in self.group_layout:
                for tracker in group_data['trackers']:
                    if global_idx in physical_order:
                        new_assignments = []
                        for dev_idx, count in physical_order[global_idx]:
                            color = self.colors[dev_idx % len(self.colors)]
                            new_assignments.append({
                                'color': color,
                                'strings': count,
                                'inv_idx': dev_idx,
                            })
                        tracker['assignments'] = new_assignments
                    global_idx += 1
        else:
            if not assignments:
                return
            # Fallback: use start_string_pos if available, else device-index ordering
            tracker_cb_map = {}
            for cb_idx, cb in enumerate(assignments):
                for conn in cb.get('connections', []):
                    tidx = conn['tracker_idx']
                    if tidx not in tracker_cb_map:
                        tracker_cb_map[tidx] = []
                    start_pos = conn.get('start_string_pos', None)
                    tracker_cb_map[tidx].append((start_pos, cb_idx, conn['num_strings']))
            
            if not tracker_cb_map:
                return
            
            # Sort each tracker's contributions by start_string_pos (north first)
            for tidx in tracker_cb_map:
                entries = tracker_cb_map[tidx]
                if all(e[0] is not None for e in entries):
                    entries.sort(key=lambda e: e[0])
                # else: leave in original device order
            
            global_idx = 0
            for group_data in self.group_layout:
                for tracker in group_data['trackers']:
                    if global_idx in tracker_cb_map:
                        new_assignments = []
                        for _start_pos, cb_idx, strings_taken in tracker_cb_map[global_idx]:
                            color = self.colors[cb_idx % len(self.colors)]
                            new_assignments.append({
                                'color': color,
                                'strings': strings_taken,
                                'inv_idx': cb_idx,
                            })
                        tracker['assignments'] = new_assignments
                    global_idx += 1

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