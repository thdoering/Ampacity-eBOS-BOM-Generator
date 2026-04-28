import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import math
import copy


class SitePreviewWindow(tk.Toplevel):
    """Pop-out window for site layout preview with zoom and pan"""
    
    def __init__(self, parent, inv_summary, topology, colors, groups, enabled_templates, row_spacing_ft,
                 num_devices=0, device_label='CB', initial_inspect=False, pads=None, device_names=None,
                 device_feeder_sizes=None, device_feeder_parallel_counts=None, measurements=None):
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
        self.device_feeder_sizes = dict(device_feeder_sizes) if device_feeder_sizes else {}  # {device_idx: "cable_size"}
        self.device_feeder_parallel_counts = dict(device_feeder_parallel_counts) if device_feeder_parallel_counts else {}  # {device_idx: int parallel sets per pole}
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
        self._drag_group_starts = {}    # {idx: (start_x, start_y)} for all selected groups
        self._drag_leader_idx = None    # group idx that initiated the current drag
        self.selected_group_indices = set()  # set of selected group indices (layout mode)
        self.align_on_motor = True
        self._layout_box_selecting = False   # True during rubber-band box select (layout mode)
        self._layout_box_select_start = (0, 0)
        
        # Read lock state from parent (QuickEstimate)
        self.allocation_locked = getattr(self.master, 'allocation_locked', False)

        # Shared device-string state — used by both the Edit Devices dialog and canvas edits.
        # None means "needs to be (re)built from allocation data". Set by _show_edit_devices_dialog
        # and canvas drag operations; cleared by _invalidate_device_data().
        self._device_data = None
        self._device_metadata = None
        self._highlighted_strings = set()

        # Canvas string selection state (inspect mode only)
        self._last_clicked_string = None   # (tracker_idx, s_idx) of last plain/shift click
        self._string_rects = []            # cache populated by draw(); used by hit_test_string
        self._box_selecting = False        # True during Shift+drag rubber-band box select
        self._box_select_start = (0, 0)   # canvas coords where Shift+drag began

        # Dialog↔canvas sync references (set while Edit Devices dialog is open)
        self._edit_dialog_tree = None
        self._edit_dialog_string_tracker = None
        self._edit_dialog_refresh_tree = None
        self._canvas_syncing_tree = False  # guard flag — prevents <<TreeviewSelect>> bounce

        # String drag state (inspect mode only)
        self._pending_string_drag = False     # press on string, waiting to see if drag starts
        self._dragging_strings = False        # drag threshold exceeded, drag is active
        self._drag_string_payload = set()     # snapshot of _highlighted_strings at drag start
        self._drag_string_start_xy = (0, 0)  # canvas coords where drag began
        self._drag_hover_device_idx = None    # device currently under cursor during drag

        # Measurement tool state
        self.measurements = list(measurements) if measurements else []
        self.current_measure_pts = []   # world-coord points for the in-progress measurement
        self.measure_mode = False
        self.measure_mouse_pos = None   # canvas coords of cursor (for rubber-band)

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
        # --- Row 1: navigation and view controls ---
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

        self.zoom_label = ttk.Label(top_bar, text="100%")
        self.zoom_label.pack(side='left', padx=10)

        # Summary info (right-aligned in row 1)
        num_inv = self.inv_summary.get('total_inverters', 0)
        total_str = self.inv_summary.get('total_strings', 0)
        actual_ratio = self.inv_summary.get('actual_dc_ac', 0)
        split = self.inv_summary.get('total_split_trackers', 0)

        summary_text = self._format_summary(num_inv, total_str, actual_ratio, split)
        self.summary_label = ttk.Label(top_bar, text=summary_text, foreground='#333333')
        self.summary_label.pack(side='right', padx=10)

        # --- Row 2: layout editing and measurement tools ---
        top_bar2 = ttk.Frame(self, padding=(5, 0, 5, 4))
        top_bar2.pack(fill='x')

        self.add_pad_btn = ttk.Button(top_bar2, text="+ Add Pad", command=self._add_pad)
        self.add_pad_btn.pack(side='left', padx=2)

        self.assign_btn = ttk.Button(top_bar2, text="Assign Devices", command=self._show_assignment_dialog)
        self.assign_btn.pack(side='left', padx=2)

        self.edit_devices_btn = ttk.Button(top_bar2, text="Edit Devices", command=self._show_edit_devices_dialog)
        self.edit_devices_btn.pack(side='left', padx=2)

        self.show_routes_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            top_bar2, text="Show Routes",
            variable=self.show_routes_var,
            command=self.draw
        ).pack(side='left', padx=4)

        ttk.Separator(top_bar2, orient='vertical').pack(side='left', fill='y', padx=8, pady=2)

        self.measure_btn = ttk.Button(top_bar2, text="Measure", command=self._toggle_measure_mode)
        self.measure_btn.pack(side='left', padx=2)

        self.show_measurements_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(top_bar2, text="Show Dims",
                        variable=self.show_measurements_var,
                        command=self.draw).pack(side='left', padx=4)

        ttk.Button(top_bar2, text="Clear Dims", command=self._measure_clear).pack(side='left', padx=2)
        
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
        self.canvas.bind('<Motion>', self._on_measure_motion)
        self.bind('<Escape>', lambda e: self._measure_cancel())
        
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

    def _format_summary(self, num_inv, total_str, actual_ratio, split,
                        spatial_runs=1, locked=False):
        """Format the top-bar summary string, topology-aware."""
        lock_str = "  |  🔒 LOCKED" if locked else ""
        runs_str = f"  |  {spatial_runs} Run(s)" if spatial_runs > 1 else ""
        
        if self.topology == 'Central Inverter':
            num_cbs = self.num_devices
            central_count = self.inv_summary.get('central_inverter_count', 1)
            return (f"{num_cbs} CBs  |  {central_count} Central Inv  |  {total_str} Strings  |  "
                    f"DC:AC: {actual_ratio:.2f}  |  {split} Split Trackers{runs_str}  |  "
                    f"{self.topology}{lock_str}")
        elif self.topology == 'Centralized String':
            return (f"{num_inv} CBs/Inv  |  {total_str} Strings  |  "
                    f"DC:AC: {actual_ratio:.2f}  |  {split} Split Trackers{runs_str}  |  "
                    f"{self.topology}{lock_str}")
        else:
            return (f"{num_inv} SIs  |  {total_str} Strings  |  "
                    f"DC:AC: {actual_ratio:.2f}  |  {split} Split Trackers{runs_str}  |  "
                    f"{self.topology}{lock_str}")

    def build_layout_data(self):
        """Build a group-based layout of trackers with physical dimensions from templates.
        
        World units are in feet. Trackers run N-S (Y axis), spaced E-W (X axis).
        Each group has an (x, y) position in world-space representing its top-left corner.
        """
        self.group_layout = []
        self.selected_group_indices = set()

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
            grp_pitch = group_data.get('row_spacing_ft', self.row_spacing_ft)
            
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
                    else:
                        # Tracker added after allocation was locked — render as unallocated (gray)
                        tracker = {
                            'strings_per_tracker': 0,
                            'assignments': [],
                            'template_ref': ref,
                        }
                        if dims:
                            tracker['width_ft'] = dims[0]
                            tracker['length_ft'] = dims[1]
                        else:
                            tracker['width_ft'] = fallback_width_ft
                            tracker['length_ft'] = fallback_length_ft
                        motor_y, motor_gap, has_motor = self.get_motor_position_in_tracker(ref)
                        tracker['motor_y_ft'] = motor_y
                        tracker['motor_gap_ft'] = motor_gap
                        tracker['has_motor'] = has_motor
                        tracker['partial_module_count'] = 0
                        tracker['partial_string_side'] = 'north'
                        tracker['full_string_count'] = 0
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
                group_width = group_max_width + (num_trackers - 1) * grp_pitch
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
            driveline_tan = math.tan(driveline_angle_rad) if driveline_angle_deg != 0 else 0.0
            
            visual_min_y_base = 0.0
            visual_max_y_base = 0.0
            
            tracker_alignment = group_data.get('tracker_alignment', 'motor')
            if group_trackers:
                for t_i, t in enumerate(group_trackers):
                    t_length_val = t.get('length_ft', group_length)
                    if tracker_alignment == 'top':
                        y_offset = 0.0
                    elif tracker_alignment == 'bottom':
                        y_offset = group_length - t_length_val
                    else:  # 'motor'
                        t_motor = t.get('motor_y_ft', 0)
                        y_offset = (ref_motor or 0) - t_motor
                    angle_y = t_i * grp_pitch * driveline_tan
                    # Base bounds (no angle) — for parallelogram overlap checking
                    visual_min_y_base = min(visual_min_y_base, y_offset)
                    visual_max_y_base = max(visual_max_y_base, y_offset + t_length_val)
                    # Full bounds (with angle) — for bounding box and selection highlight
                    visual_min_y_offset = min(visual_min_y_offset, y_offset + angle_y)
                    visual_max_y_offset = max(visual_max_y_offset, y_offset + angle_y + t_length_val)
            
            azimuth_deg = group_data.get('azimuth', 180)
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
                'row_spacing_ft': grp_pitch,
                'azimuth': azimuth_deg,
                'rotation_deg': azimuth_deg - 180,
                'tracker_alignment': tracker_alignment,
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
        
        if inverters:
            self._compute_devices_from_allocation(
                inverters, device_width_ft, device_height_ft, offset_ft
            )
        elif self.topology == 'Central Inverter':
            # Fallback if no allocation available
            self._compute_devices_proportional(
                device_width_ft, device_height_ft, offset_ft
            )
    
    def _apply_middle_x_bias(self, device_x, device_y, center_local, local_indices,
                              strings_per_tracker_map, pitch, group_x, group_num_trackers):
        """Shift device_x into the row-spacing gap for 'middle' placement.

        Biases east if more strings are east of center_local, west if more are west.
        Tie-breaks toward the nearest pad (defaults east when no pads exist).
        Falls back to the opposite direction if the bias would leave the group bounding rect.
        """
        # If center_local is already at a gap position (half-integer like 0.5, 1.5 ...),
        # the device_x is already centered in the row gap — no bias needed.
        if abs(center_local % 1 - 0.5) < 0.01:
            return device_x

        half_pitch = pitch / 2.0
        east_strings = sum(
            strings_per_tracker_map.get(i, 1) for i in local_indices if i > center_local
        )
        west_strings = sum(
            strings_per_tracker_map.get(i, 1) for i in local_indices if i < center_local
        )

        if east_strings > west_strings:
            bias = half_pitch
        elif west_strings > east_strings:
            bias = -half_pitch
        else:
            bias = half_pitch  # default east
            if self.pads:
                nearest_pad = min(
                    self.pads,
                    key=lambda p: (device_x - (p['x'] + p.get('width_ft', 10.0) / 2)) ** 2
                                  + (device_y - (p['y'] + p.get('height_ft', 8.0) / 2)) ** 2
                )
                pad_cx = nearest_pad['x'] + nearest_pad.get('width_ft', 10.0) / 2
                bias = half_pitch if pad_cx >= device_x else -half_pitch

        x_min = group_x
        x_max = group_x + max(group_num_trackers - 1, 0) * pitch + self.max_tracker_width_ft

        new_x = device_x + bias
        if new_x < x_min or new_x > x_max:
            # Bias pushed outside group bounds — try opposite side
            new_x = max(x_min, min(x_max, device_x - bias))

        return new_x

    def _compute_devices_from_allocation(self, inverters, device_width_ft, device_height_ft, offset_ft):
        """Place one device per inverter, positioned at the center of that inverter's trackers."""
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
            
            pitch = group_data.get('row_spacing_ft', self.tracker_pitch_ft)
            if local_indices:
                center_local = (min(local_indices) + max(local_indices)) / 2.0
                device_x = gx + center_local * pitch + (max_width - device_width_ft) / 2
            else:
                center_local = 0
                device_x = gx
            
            # Driveline angle Y offset based on device's X position in group
            angle_y_offset = center_local * pitch * group_data.get('driveline_tan', 0.0)
            
            # Compute Y based on position setting.
            # For 'north' and 'south', find the tracker among THIS device's own trackers
            # whose X position is closest to the device's X, and anchor to that tracker's
            # edge. This places the combiner just off the tracker nearest to it, even
            # when the device's trackers have varying lengths.
            if device_position in ('north', 'south'):
                group_trackers_list = group_data.get('trackers', [])
                group_motor_y_ref = group_data.get('motor_y_ft', None)
                group_length = group_data.get('length_ft', 0)
                driveline_tan = group_data.get('driveline_tan', 0.0)
                align_on_motor = getattr(self, 'align_on_motor', False)
                
                # Find the tracker closest to the device's X position (center_local)
                closest_local_idx = None
                closest_dist = float('inf')
                for local_idx in local_indices:
                    if local_idx >= len(group_trackers_list):
                        continue
                    dist = abs(local_idx - center_local)
                    if dist < closest_dist:
                        closest_dist = dist
                        closest_local_idx = local_idx
                
                if closest_local_idx is not None:
                    t = group_trackers_list[closest_local_idx]
                    t_length = t.get('length_ft', group_length)
                    t_angle_y = closest_local_idx * pitch * driveline_tan
                    
                    if align_on_motor and t.get('has_motor', False) and group_motor_y_ref is not None:
                        ty = gy + (group_motor_y_ref - t.get('motor_y_ft', 0)) + t_angle_y
                    else:
                        ty = gy + (group_length - t_length) / 2 + t_angle_y
                    
                    if device_position == 'north':
                        device_y = ty - offset_ft - device_height_ft
                    else:  # 'south'
                        device_y = ty + t_length + offset_ft
                else:
                    # Fallback to group bounds if no trackers found
                    if device_position == 'north':
                        vis_min = group_data.get('visual_min_y', 0)
                        device_y = gy + vis_min - offset_ft - device_height_ft + angle_y_offset
                    else:
                        vis_max = group_data.get('visual_max_y', group_data['length_ft'])
                        device_y = gy + vis_max + offset_ft + angle_y_offset
            else:  # 'middle'
                motor_y = group_data.get('motor_y_ft', group_data['length_ft'] / 2)
                device_y = gy + motor_y - device_height_ft / 2 + angle_y_offset
                if local_indices:
                    spt_map = {
                        li: group_data['trackers'][li].get('strings_per_tracker', 1)
                        for li in local_indices if li < len(group_data.get('trackers', []))
                    }
                    device_x = self._apply_middle_x_bias(
                        device_x, device_y, center_local, local_indices, spt_map,
                        pitch, gx, len(group_data.get('trackers', []))
                    )

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
                
                pitch = group_data.get('row_spacing_ft', self.tracker_pitch_ft)
                center_tracker = tracker_start + sub_size / 2.0 - 0.5
                device_x = gx + center_tracker * pitch + (max_width - device_width_ft) / 2
                device_y = base_device_y + center_tracker * pitch * driveline_tan

                if device_position == 'middle':
                    sub_local_indices = list(range(tracker_start, tracker_start + sub_size))
                    spt_map = {
                        li: group_trackers[li].get('strings_per_tracker', 1)
                        for li in sub_local_indices if li < len(group_trackers)
                    }
                    device_x = self._apply_middle_x_bias(
                        device_x, device_y, center_tracker, sub_local_indices, spt_map,
                        pitch, gx, num_trackers_in_group
                    )

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
        
        all_xs, all_ys = [], []
        for g in self.group_layout:
            gx, gy = g['x'], g['y']
            vis_min = g.get('visual_min_y', 0)
            vis_max = g.get('visual_max_y', g['length_ft'])
            rd = g.get('rotation_deg', 0.0)
            rcx = gx + g['width_ft'] / 2
            rcy = gy + (vis_min + vis_max) / 2
            corners = [
                (gx, gy + vis_min), (gx + g['width_ft'], gy + vis_min),
                (gx + g['width_ft'], gy + vis_max), (gx, gy + vis_max),
            ]
            for cx2, cy2 in corners:
                if rd != 0:
                    cx2, cy2 = self._rotate_point(rcx, rcy, cx2, cy2, rd)
                all_xs.append(cx2)
                all_ys.append(cy2)
        min_x, max_x = min(all_xs), max(all_xs)
        min_y, max_y = min(all_ys), max(all_ys)
        
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
    
    def _rotate_point(self, cx, cy, x, y, angle_deg):
        """Rotate (x, y) around (cx, cy) by angle_deg degrees (positive = clockwise geographic)."""
        rad = math.radians(angle_deg)
        dx, dy = x - cx, y - cy
        rx = dx * math.cos(rad) - dy * math.sin(rad) + cx
        ry = dx * math.sin(rad) + dy * math.cos(rad) + cy
        return rx, ry

    def _device_rotation_info(self, dev):
        """Return (rot_cx, rot_cy, rotation_deg) for the group a device belongs to.
        Returns (None, None, 0.0) when there is no rotation."""
        grp_idx = dev.get('group_idx')
        if grp_idx is None or grp_idx >= len(self.group_layout):
            return None, None, 0.0
        g = self.group_layout[grp_idx]
        rd = g.get('rotation_deg', 0.0)
        if not rd:
            return None, None, 0.0
        vis_min = g.get('visual_min_y', 0)
        vis_max = g.get('visual_max_y', g['length_ft'])
        rcx = g['x'] + g['width_ft'] / 2
        rcy = g['y'] + (vis_min + vis_max) / 2
        return rcx, rcy, rd

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
        self.canvas.scale('all', center_x, center_y, factor, factor)
        self._schedule_redraw()
    
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
            rd = g.get('rotation_deg', 0.0)
            if rd != 0:
                # Un-rotate click point into group's local (unrotated) space
                rcx = g['x'] + g['width_ft'] / 2
                rcy = g['y'] + (vis_min + vis_max) / 2
                lwx, lwy = self._rotate_point(rcx, rcy, wx, wy, -rd)
            else:
                lwx, lwy = wx, wy
            if (g['x'] <= lwx <= g['x'] + g['width_ft'] and
                    g['y'] + vis_min <= lwy <= g['y'] + vis_max):
                return i
        return None
    
    def on_press(self, event):
        """Handle left mouse press — place pad, select/drag group or pad, or select device."""
        self.drag_start_x = event.x
        self.drag_start_y = event.y
        self._drag_moved = False
        self._dragging_pad = False

        # Measurement mode — each click plants a vertex
        if self.measure_mode:
            wx, wy = self.canvas_to_world(event.x, event.y)
            self.current_measure_pts.append([wx, wy])
            self.draw()
            return

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
            
            # Check strings (before devices so individual strings are selectable)
            hit_str = self.hit_test_string(event.x, event.y)
            if hit_str is not None:
                tracker_idx, s_idx = hit_str
                ctrl = event.state & 0x4
                shift = event.state & 0x1
                if ctrl:
                    key = (tracker_idx, s_idx)
                    if key in self._highlighted_strings:
                        self._highlighted_strings.discard(key)
                    else:
                        self._highlighted_strings.add(key)
                elif shift and self._last_clicked_string is not None:
                    last_tidx, last_sidx = self._last_clicked_string
                    if last_tidx == tracker_idx:
                        lo, hi = min(last_sidx, s_idx), max(last_sidx, s_idx)
                        for i in range(lo, hi + 1):
                            self._highlighted_strings.add((tracker_idx, i))
                    else:
                        self._highlighted_strings = {(tracker_idx, s_idx)}
                else:
                    key = (tracker_idx, s_idx)
                    if key not in self._highlighted_strings:
                        # Clicking an unselected string — replace selection
                        self._highlighted_strings = {key}
                    # Clicking an already-selected string — keep current selection so
                    # a subsequent drag carries all selected strings, not just this one
                self._last_clicked_string = (tracker_idx, s_idx)
                self.selected_device_idx = None
                self.selected_pad_inspect_idx = None
                self.dragging_canvas = False
                self.dragging_group = False
                self._pending_string_drag = True
                self._drag_string_start_xy = (event.x, event.y)
                self._sync_tree_from_canvas()
                self.draw()
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
                # Empty space — always start a potential box select.
                # Whether it's a click (clears selection) or a drag (box select)
                # is decided at release time based on how far the mouse moved.
                self.selected_device_idx = None
                self.selected_pad_inspect_idx = None
                self._box_selecting = True
                self._box_select_start = (event.x, event.y)
                self.dragging_canvas = False
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
            self.selected_group_indices = set()
            self.dragging_group = False
            self.draw()
            return

        self.selected_pad_idx = None

        hit = self.hit_test_group(event.x, event.y)
        ctrl = event.state & 0x4
        if hit is not None:
            if ctrl:
                # Ctrl+click: toggle membership, no drag
                if hit in self.selected_group_indices:
                    self.selected_group_indices.discard(hit)
                else:
                    self.selected_group_indices.add(hit)
                self.dragging_group = False
            else:
                if hit not in self.selected_group_indices:
                    # Plain click on unselected group — replace selection
                    self.selected_group_indices = {hit}
                # Start drag carrying all currently selected groups
                self._drag_leader_idx = hit
                self._drag_group_starts = {
                    idx: (self.group_layout[idx]['x'], self.group_layout[idx]['y'])
                    for idx in self.selected_group_indices
                }
                self.dragging_group = True
            self.draw()
        else:
            # Empty space — clear selection and start rubber-band box select
            self.selected_group_indices = set()
            self._layout_box_selecting = True
            self._layout_box_select_start = (event.x, event.y)
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
        if self.measure_mode:
            return

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
        if getattr(self, '_box_selecting', False):
            self.canvas.delete('selection_box')
            sx, sy = self._box_select_start
            self.canvas.create_rectangle(
                min(sx, event.x), min(sy, event.y),
                max(sx, event.x), max(sy, event.y),
                outline='#4A90D9', width=1, dash=(4, 3),
                tags='selection_box'
            )
            return

        if getattr(self, '_layout_box_selecting', False):
            self.canvas.delete('layout_selection_box')
            sx, sy = self._layout_box_select_start
            self.canvas.create_rectangle(
                min(sx, event.x), min(sy, event.y),
                max(sx, event.x), max(sy, event.y),
                outline='#4A90D9', width=1, dash=(4, 3),
                tags='layout_selection_box'
            )
            return

        # Promote pending string drag to active once threshold is exceeded
        if getattr(self, '_pending_string_drag', False):
            if (abs(event.x - self.drag_start_x) > 3 or
                    abs(event.y - self.drag_start_y) > 3):
                self._pending_string_drag = False
                self._dragging_strings = True
                self._drag_string_payload = set(self._highlighted_strings)

        # String drag in progress — draw ghost badge and device hover highlight
        if getattr(self, '_dragging_strings', False):
            self.canvas.delete('string_drag_ghost')
            self.canvas.delete('device_drop_target')

            n = len(self._drag_string_payload)
            badge_text = f"{n} string{'s' if n != 1 else ''}"
            tid = self.canvas.create_text(
                event.x + 14, event.y - 10,
                text=badge_text, font=('Helvetica', 9, 'bold'),
                fill='white', anchor='w', tags='string_drag_ghost'
            )
            bbox = self.canvas.bbox(tid)
            if bbox:
                self.canvas.create_rectangle(
                    bbox[0] - 4, bbox[1] - 3, bbox[2] + 4, bbox[3] + 3,
                    fill='#4A90D9', outline='#2266AA', width=1,
                    tags='string_drag_ghost'
                )
                self.canvas.tag_raise(tid)

            hit = self.hit_test_device_loose(event.x, event.y)
            self._drag_hover_device_idx = hit
            if hit is not None and hasattr(self, 'device_positions') and hit < len(self.device_positions):
                dev = self.device_positions[hit]
                rcx, rcy, rd = self._device_rotation_info(dev)
                corners = [
                    (dev['x'], dev['y']),
                    (dev['x'] + dev['width_ft'], dev['y']),
                    (dev['x'] + dev['width_ft'], dev['y'] + dev['height_ft']),
                    (dev['x'], dev['y'] + dev['height_ft']),
                ]
                poly_pts = []
                for wx2, wy2 in corners:
                    if rd:
                        wx2, wy2 = self._rotate_point(rcx, rcy, wx2, wy2, rd)
                    px2, py2 = self.world_to_canvas(wx2, wy2)
                    poly_pts.extend([px2, py2])
                self.canvas.create_polygon(
                    *poly_pts, fill='', outline='#00CC44', width=3,
                    tags='device_drop_target'
                )
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
        elif getattr(self, 'dragging_group', False) and self.selected_group_indices:
            dx_world = dx_px / self.scale if self.scale != 0 else 0
            dy_world = dy_px / self.scale if self.scale != 0 else 0

            leader_idx = self._drag_leader_idx
            if leader_idx is None or leader_idx not in self._drag_group_starts:
                leader_idx = next(iter(self.selected_group_indices))
            leader_sx, leader_sy = self._drag_group_starts.get(leader_idx, (0, 0))

            raw_x = leader_sx + dx_world
            raw_y = leader_sy + dy_world

            shift_held = event.state & 0x1
            if shift_held:
                # Shift held — constrain to N/S movement only (lock X)
                snapped_dx = 0
                snapped_dy = raw_y - leader_sy
            else:
                # Normal drag — snap leader, apply resulting delta to all followers
                new_x, new_y = self._snap_group_position(leader_idx, raw_x, raw_y)
                snapped_dx = new_x - leader_sx
                snapped_dy = new_y - leader_sy

            for idx in self.selected_group_indices:
                sx, sy = self._drag_group_starts.get(idx, (self.group_layout[idx]['x'], self.group_layout[idx]['y']))
                self.group_layout[idx]['x'] = sx + snapped_dx
                self.group_layout[idx]['y'] = sy + snapped_dy

            self.draw()
    
    def on_release(self, event):
        """Handle mouse release — finalize group or pad position."""
        if getattr(self, '_dragging_panel', False):
            self._dragging_panel = False
            self._panel_drag_start = None
            return

        if getattr(self, '_layout_box_selecting', False):
            self._layout_box_selecting = False
            self.canvas.delete('layout_selection_box')
            sx, sy = self._layout_box_select_start
            if abs(event.x - sx) > 5 or abs(event.y - sy) > 5:
                # Real drag — hit-test group bounding boxes against the rubber-band rect
                bx1, bx2 = min(sx, event.x), max(sx, event.x)
                by1, by2 = min(sy, event.y), max(sy, event.y)
                wx1, wy1 = self.canvas_to_world(bx1, by1)
                wx2, wy2 = self.canvas_to_world(bx2, by2)
                wxmin, wxmax = min(wx1, wx2), max(wx1, wx2)
                wymin, wymax = min(wy1, wy2), max(wy1, wy2)
                for i, g in enumerate(self.group_layout):
                    vis_min = g.get('visual_min_y', 0)
                    vis_max = g.get('visual_max_y', g['length_ft'])
                    gx1, gy1 = g['x'], g['y'] + vis_min
                    gx2, gy2 = g['x'] + g['width_ft'], g['y'] + vis_max
                    if gx2 >= wxmin and gx1 <= wxmax and gy2 >= wymin and gy1 <= wymax:
                        self.selected_group_indices.add(i)
            # else: plain click in empty space — selection already cleared in on_press
            self.draw()
            return

        if getattr(self, '_box_selecting', False):
            self._box_selecting = False
            self.canvas.delete('selection_box')
            sx, sy = self._box_select_start
            if abs(event.x - sx) < 5 and abs(event.y - sy) < 5:
                # Treat as a plain click — clear all string selection
                self._highlighted_strings = set()
                self._last_clicked_string = None
            else:
                # Real drag — add strings inside the box to the current selection
                hits = self.hit_test_strings_in_box(sx, sy, event.x, event.y)
                for tidx, sidx in hits:
                    self._highlighted_strings.add((tidx, sidx))
                self._sync_tree_from_canvas()
            self.draw()
            return

        if getattr(self, '_pending_string_drag', False):
            # Press on string, no drag motion — plain click already handled in on_press.
            self._pending_string_drag = False
            return

        if getattr(self, '_dragging_strings', False):
            self._dragging_strings = False
            self._pending_string_drag = False
            self.canvas.delete('string_drag_ghost')
            self.canvas.delete('device_drop_target')

            target_dev_idx = self.hit_test_device_loose(event.x, event.y)
            if target_dev_idx is not None and self._drag_string_payload:
                ok, err = self._canvas_move_strings(self._drag_string_payload, target_dev_idx)
                if ok:
                    self._highlighted_strings = set()
                    self._last_clicked_string = None
                    self._apply_canvas_string_edit()
                    self._sync_tree_from_canvas()
                elif err:
                    messagebox.showwarning("Invalid Move", err, parent=self)
            self.draw()
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
            self.canvas.move('all', dx_px, dy_px)

    def on_pan_release(self, event):
        """Handle middle mouse release — stop panning."""
        self.dragging_canvas = False
        self.draw()
    
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
    
    def refresh_wire_gauges(self):
        """Redraw the canvas so the device info panel reflects updated wire gauges."""
        self.draw()

    def _schedule_redraw(self, delay_ms=40):
        """Coalesce rapid redraw requests into a single deferred draw."""
        existing = getattr(self, '_redraw_after_id', None)
        if existing is not None:
            try:
                self.after_cancel(existing)
            except Exception:
                pass
        self._redraw_after_id = self.after(delay_ms, self._do_scheduled_redraw)

    def _do_scheduled_redraw(self):
        self._redraw_after_id = None
        self.draw()

    def draw(self):
        """Draw the site layout with to-scale trackers at their group positions.
        
        X = E-W (tracker width + row spacing gaps)
        Y = N-S (tracker length, north at top)
        World units = feet.
        """
        self.canvas.delete('all')
        self._string_rects = []

        if not self.group_layout:
            return

        # Viewport bounds in world space for per-group culling
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        view_x_min, view_y_min = self.canvas_to_world(0, 0)
        view_x_max, view_y_max = self.canvas_to_world(canvas_w, canvas_h)
        if view_x_min > view_x_max:
            view_x_min, view_x_max = view_x_max, view_x_min
        if view_y_min > view_y_max:
            view_y_min, view_y_max = view_y_max, view_y_min

        max_width = getattr(self, 'max_tracker_width_ft', 6)

        for group_idx, group_data in enumerate(self.group_layout):
            pitch = group_data.get('row_spacing_ft', getattr(self, 'tracker_pitch_ft', 20))
            gx = group_data['x']
            gy = group_data['y']
            is_selected = (group_idx in self.selected_group_indices)
            rotation_deg = group_data.get('rotation_deg', 0.0)
            vis_min = group_data.get('visual_min_y', 0)
            vis_max = group_data.get('visual_max_y', group_data['length_ft'])
            rot_cx = gx + group_data['width_ft'] / 2
            rot_cy = gy + (vis_min + vis_max) / 2

            def _wc(wx, wy, _rcx=rot_cx, _rcy=rot_cy, _rd=rotation_deg):
                """World-to-canvas with optional group rotation applied."""
                if _rd != 0:
                    wx, wy = self._rotate_point(_rcx, _rcy, wx, wy, _rd)
                return self.world_to_canvas(wx, wy)

            def _rect_as_poly(x1, y1, x2, y2, _rcx=rot_cx, _rcy=rot_cy, _rd=rotation_deg):
                """Return flat canvas coord list for a rotated rectangle polygon."""
                corners = [(x1, y1), (x2, y1), (x2, y2), (x1, y2)]
                pts = []
                for wx, wy in corners:
                    if _rd != 0:
                        wx, wy = self._rotate_point(_rcx, _rcy, wx, wy, _rd)
                    cx2, cy2 = self.world_to_canvas(wx, wy)
                    pts.extend([cx2, cy2])
                return pts

            # Viewport culling — skip groups entirely outside the visible canvas
            margin = max(self.max_tracker_width_ft, 10.0) * 2.0
            g_x_min = gx
            g_x_max = gx + group_data['width_ft']
            g_y_min = gy + vis_min
            g_y_max = gy + vis_max
            if rotation_deg != 0:
                half_w = (g_x_max - g_x_min) / 2
                half_h = (g_y_max - g_y_min) / 2
                diag = (half_w ** 2 + half_h ** 2) ** 0.5
                cx_g = (g_x_min + g_x_max) / 2
                cy_g = (g_y_min + g_y_max) / 2
                g_x_min, g_x_max = cx_g - diag, cx_g + diag
                g_y_min, g_y_max = cy_g - diag, cy_g + diag
            if (g_x_max + margin < view_x_min or
                    g_x_min - margin > view_x_max or
                    g_y_max + margin < view_y_min or
                    g_y_min - margin > view_y_max):
                continue

            # Draw selection highlight behind group (using visual bounds)
            if is_selected:
                pad = max_width * 0.3
                h_poly = _rect_as_poly(
                    gx - pad, gy + vis_min - pad,
                    gx + group_data['width_ft'] + pad, gy + vis_max + pad
                )
                self.canvas.create_polygon(
                    *h_poly, fill='', outline='#4A90D9', width=2
                )
            
            # Draw group label
            label_x, label_y = _wc(
                gx - max_width * 0.5,
                gy + group_data['length_ft'] / 2
            )
            font_size = max(6, min(11, int(9 * self.scale)))
            self._draw_text_with_bg(
                label_x, label_y,
                text=group_data['name'], font=('Helvetica', font_size),
                fill='#4A90D9' if is_selected else '#333333', anchor='e',
                bg_required=False
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
                _talign = group_data.get('tracker_alignment', 'motor')
                if _talign == 'top':
                    ty = gy + angle_y_offset
                elif _talign == 'bottom':
                    ty = gy + (group_data['length_ft'] - t_length) + angle_y_offset
                elif tracker.get('has_motor', False) and group_data.get('motor_y_ft', None) is not None:
                    # Motor alignment: offset so this tracker's motor Y matches group's reference motor Y
                    ty = gy + (group_data['motor_y_ft'] - tracker['motor_y_ft']) + angle_y_offset
                else:
                    # Center alignment fallback
                    ty_offset = (group_data['length_ft'] - t_length) / 2
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
                _hl = getattr(self, '_highlighted_strings', set())

                # First pass: compute render attributes per string slot
                _str_attrs = []  # (color, outline_color, outline_width, is_unowned_partial)
                for s_idx in range(draw_count):
                    is_unowned_partial = (partial_mods > 0 and spt <= full_str_count and
                                         ((partial_side == 'north' and s_idx == 0) or
                                          (partial_side == 'south' and s_idx == draw_count - 1)))
                    if is_unowned_partial:
                        color = '#D4C878'
                    else:
                        if partial_mods > 0 and partial_side == 'north':
                            color_idx = s_idx - 1
                            if spt > full_str_count and s_idx == 0:
                                color_idx = 0
                        elif partial_mods > 0 and partial_side == 'south':
                            if spt > full_str_count and s_idx == draw_count - 1:
                                color_idx = spt - 1
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
                            color = '#E0E0E0'
                            outline_color = '#CCCCCC'
                            outline_width = 1
                    elif self.assigning_devices:
                        color = '#E0E0E0'
                        outline_color = '#CCCCCC'
                        outline_width = 1
                    else:
                        outline_color = '#555555'
                        outline_width = 1

                    if (not is_unowned_partial and _hl and
                            (global_tracker_idx, s_idx) in _hl):
                        color = '#FFFF00'
                        outline_color = '#DAA520'
                        outline_width = 2

                    _str_attrs.append((color, outline_color, outline_width, is_unowned_partial))

                # Second pass: emit polygons.
                # When individual strings are tall enough on screen to interact with,
                # draw one polygon per string so the borders are visible. Below that
                # threshold, coalesce consecutive same-key strings to reduce item count.
                min_str_h_px = min(string_heights) * self.scale if string_heights else 0
                if min_str_h_px >= 8:
                    # Detail mode — individual string polygons
                    sy = ty
                    for s_idx in range(draw_count):
                        color, outline_color, outline_width, _ = _str_attrs[s_idx]
                        sh = string_heights[s_idx]
                        poly = _rect_as_poly(tx + tx_offset, sy, tx + tx_offset + t_width, sy + sh)
                        self.canvas.create_polygon(*poly, fill=color, outline=outline_color, width=outline_width)
                        sy += sh
                else:
                    # Performance mode — one polygon per run of consecutive same-key strings
                    sy_cursor = ty
                    run_start = 0
                    while run_start < draw_count:
                        run_color, run_outline, run_width, _ = _str_attrs[run_start]
                        run_key = (run_color, run_outline, run_width)
                        run_end = run_start
                        while (run_end + 1 < draw_count and
                               _str_attrs[run_end + 1][:3] == run_key):
                            run_end += 1
                        run_sy = sy_cursor
                        for i in range(run_start, run_end + 1):
                            sy_cursor += string_heights[i]
                        poly = _rect_as_poly(
                            tx + tx_offset, run_sy,
                            tx + tx_offset + t_width, sy_cursor
                        )
                        self.canvas.create_polygon(*poly, fill=run_color, outline=run_outline, width=run_width)
                        run_start = run_end + 1

                # Third pass: populate _string_rects per individual string (no polygon emission)
                sy = ty
                for s_idx in range(draw_count):
                    sh = string_heights[s_idx] if s_idx < len(string_heights) else string_heights[-1]
                    poly = _rect_as_poly(tx + tx_offset, sy, tx + tx_offset + t_width, sy + sh)
                    _, _, _, is_unowned_partial = _str_attrs[s_idx]
                    self._string_rects.append({
                        'tracker_idx': global_tracker_idx,
                        's_idx': s_idx,
                        'poly_canvas': poly,
                        'is_unowned_partial': is_unowned_partial,
                    })
                    sy += sh

                # Tracker outline
                out_poly = _rect_as_poly(
                    tx + tx_offset - 0.5, ty - 0.5,
                    tx + tx_offset + t_width + 0.5, ty + t_length + 0.5
                )
                self.canvas.create_polygon(
                    *out_poly, fill='', outline='#222222', width=1
                )
                # Keep ox1/oy1 for pixel_width calculation (use first two corners)
                ox1, oy1 = out_poly[0], out_poly[1]
                ox2, oy2 = out_poly[2], out_poly[3]

                # Motor indicator
                if tracker.get('has_motor', False):
                    motor_y = tracker['motor_y_ft']
                    motor_gap = tracker['motor_gap_ft']

                    motor_world_y = ty + motor_y
                    motor_x1 = tx + tx_offset - 0.3
                    motor_x2 = tx + tx_offset + t_width + 0.3

                    m_poly = _rect_as_poly(motor_x1, motor_world_y, motor_x2, motor_world_y + motor_gap)
                    self.canvas.create_polygon(
                        *m_poly, fill='#666666', outline='#444444', width=1
                    )

                    motor_cx = sum(m_poly[0::2]) / 4
                    motor_cy = sum(m_poly[1::2]) / 4
                    dot_r = max(2, min(4, 3 * self.scale))
                    self.canvas.create_oval(
                        motor_cx - dot_r, motor_cy - dot_r,
                        motor_cx + dot_r, motor_cy + dot_r,
                        fill='#FF8800', outline='#CC6600', width=1
                    )

                # Tracker label — use global tracker index to match info panel / assignments
                label_cx, label_cy = _wc(tx + tx_offset + t_width / 2, ty + t_length + 2)
                pixel_width = abs(ox2 - ox1)
                if pixel_width > 14:
                    lbl_size = max(6, min(9, int(8 * self.scale)))
                    self._draw_text_with_bg(
                        label_cx, label_cy,
                        text=f"T{global_tracker_idx+1}", font=('Helvetica', lbl_size), fill='#555555',
                        bg_required=False
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

        # Measurement annotations
        self._draw_measurements()

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

    def _draw_text_with_bg(self, x, y, text, font, fill='#333333', anchor='center', bg='white', pad=2, bg_required=True):
        """Draw canvas text, optionally with a white background rectangle for readability."""
        tid = self.canvas.create_text(x, y, text=text, font=font, fill=fill, anchor=anchor)
        if not bg_required:
            return tid
        bbox = self.canvas.bbox(tid)
        if bbox:
            self.canvas.create_rectangle(
                bbox[0] - pad, bbox[1] - pad,
                bbox[2] + pad, bbox[3] + pad,
                fill=bg, outline='', width=0
            )
            self.canvas.tag_raise(tid)
        return tid
    
    def _draw_motor_alignment_lines(self):
        """Draw a driveline across each group at its motor Y position,
        following the driveline angle if set. Applies group azimuth rotation."""
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
            
            # Apply group azimuth rotation around the group's rotation center
            rotation_deg = group_data.get('rotation_deg', 0.0)
            if rotation_deg != 0:
                vis_min = group_data.get('visual_min_y', 0)
                vis_max = group_data.get('visual_max_y', group_data['length_ft'])
                rcx = group_data['x'] + group_data['width_ft'] / 2
                rcy = group_data['y'] + (vis_min + vis_max) / 2
                left_x, left_y = self._rotate_point(rcx, rcy, left_x, left_y, rotation_deg)
                right_x, right_y = self._rotate_point(rcx, rcy, right_x, right_y, rotation_deg)
            
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
            self.selected_group_indices = set()
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
            rcx, rcy, rd = self._device_rotation_info(dev)
            if rd:
                lwx, lwy = self._rotate_point(rcx, rcy, wx, wy, -rd)
            else:
                lwx, lwy = wx, wy
            if (dev['x'] <= lwx <= dev['x'] + dev['width_ft'] and
                    dev['y'] <= lwy <= dev['y'] + dev['height_ft']):
                return i
        return None
    
    def hit_test_device_loose(self, cx, cy, padding_px=20):
        """Like hit_test_device but with an enlarged hitbox — used during string drag.

        padding_px is in canvas pixels so the hitbox stays consistently grabbable
        regardless of zoom level.
        """
        if not hasattr(self, 'device_positions') or not self.device_positions:
            return None
        wx, wy = self.canvas_to_world(cx, cy)
        pad = padding_px / self.scale if self.scale > 0 else 0
        for i, dev in enumerate(self.device_positions):
            rcx, rcy, rd = self._device_rotation_info(dev)
            if rd:
                lwx, lwy = self._rotate_point(rcx, rcy, wx, wy, -rd)
            else:
                lwx, lwy = wx, wy
            if (dev['x'] - pad <= lwx <= dev['x'] + dev['width_ft'] + pad and
                    dev['y'] - pad <= lwy <= dev['y'] + dev['height_ft'] + pad):
                return i
        return None

    def _canvas_move_strings(self, payload, target_dev_idx):
        """Move strings in payload (set of (tracker_idx, phys_pos)) to target device.

        Validates contiguity on source and target before executing. Returns (ok, error_msg).
        """
        if not self._ensure_device_data():
            return False, "No device data available."
        device_data = self._device_data

        # device_positions never contains the Unallocated entry, so target_dev_idx
        # is an index into the real-device subset of device_data. Map it through.
        real_dev_indices = [i for i, d in enumerate(device_data) if not d.get('is_unallocated')]
        if target_dev_idx >= len(real_dev_indices):
            return False, "No device data available."
        actual_target_idx = real_dev_indices[target_dev_idx]
        target_dev = device_data[actual_target_idx]

        # Build lookup: (tidx, phys_pos) -> source device index
        string_to_dev = {}
        for dev_idx, dev in enumerate(device_data):
            for key in dev['strings']:
                string_to_dev[key] = dev_idx

        # Group payload by source device
        by_source = {}
        for key in payload:
            src_idx = string_to_dev.get(key)
            if src_idx is not None:
                by_source.setdefault(src_idx, set()).add(key)

        if not by_source:
            return False, "Selected strings not found in any device."

        # Validate contiguity for each source→target pair. Positions that sit in
        # the Unallocated device don't count as gaps — they're just not assigned
        # yet, so a temporary "gap" there is acceptable during reallocation.
        unalloc_positions = {}  # tracker_idx -> set of unallocated phys_pos
        for dev in device_data:
            if dev.get('is_unallocated'):
                for t, p in dev['strings']:
                    unalloc_positions.setdefault(t, set()).add(p)

        for src_idx, src_strings in by_source.items():
            if src_idx == actual_target_idx:
                continue
            src_dev = device_data[src_idx]
            if src_dev.get('is_unallocated') or target_dev.get('is_unallocated'):
                continue

            by_tracker = {}
            for tidx, phys_pos in src_strings:
                by_tracker.setdefault(tidx, set()).add(phys_pos)

            for tidx, moving in by_tracker.items():
                unalloc_t = unalloc_positions.get(tidx, set())

                src_all = {p for t, p in src_dev['strings'] if t == tidx}
                remaining = src_all - moving
                if len(remaining) > 1 and max(remaining) - min(remaining) + 1 != len(remaining):
                    missing = set(range(min(remaining), max(remaining) + 1)) - remaining
                    if not missing.issubset(unalloc_t):
                        return False, (f"Cannot move: would leave a gap in T{tidx+1:02d} "
                                       f"on {src_dev['name']}")

                tgt_all = {p for t, p in target_dev['strings'] if t == tidx}
                combined = tgt_all | moving
                if len(combined) > 1 and max(combined) - min(combined) + 1 != len(combined):
                    missing = set(range(min(combined), max(combined) + 1)) - combined
                    if not missing.issubset(unalloc_t):
                        return False, (f"Cannot move: would create a gap in T{tidx+1:02d} "
                                       f"on {target_dev['name']}")

        # Execute moves
        any_moved = False
        for src_idx, src_strings in by_source.items():
            if src_idx == actual_target_idx:
                continue
            src_dev = device_data[src_idx]
            src_dev['strings'] = [(t, p) for t, p in src_dev['strings'] if (t, p) not in src_strings]
            target_dev['strings'].extend(src_strings)
            any_moved = True

        if not any_moved:
            return False, ""  # no-op: strings already on target

        target_dev['strings'].sort(key=lambda s: (s[0], s[1]))
        return True, ""

    def _apply_canvas_string_edit(self):
        """Rebuild allocation and redraw after a canvas-initiated string move.

        Mirrors _update_live_preview from the Edit Devices dialog.
        """
        if not self._device_data or not self._device_metadata:
            return
        real_devs = [d for d in self._device_data if not d.get('is_unallocated')]
        self._rebuild_from_device_strings(real_devs, self._device_metadata)
        self._update_lock_button()
        if hasattr(self.master, 'manually_edited'):
            self.master.manually_edited = True

        inv_summary = getattr(self.master, 'last_totals', {}).get('inverter_summary', {})
        if inv_summary and inv_summary.get('allocation_result'):
            self.inv_summary = inv_summary
            alloc = inv_summary.get('allocation_result', {})
            num_inv = alloc.get('summary', {}).get('total_inverters', 0)
            total_str = alloc.get('summary', {}).get('total_strings', 0)
            split = alloc.get('summary', {}).get('total_split_trackers', 0)
            spatial_runs = alloc.get('spatial_runs', 1)
            actual_ratio = inv_summary.get('actual_dc_ac', 0)
            self.summary_label.config(
                text=self._format_summary(num_inv, total_str, actual_ratio, split,
                                          spatial_runs=spatial_runs, locked=True)
            )

        self.build_layout_data()
        self._recolor_from_cb_assignments()
        self._build_legend()

        # Refresh dialog tree if open
        refresh_fn = self._edit_dialog_refresh_tree
        if refresh_fn is not None:
            try:
                refresh_fn()
            except Exception:
                pass

    def _point_in_polygon(self, px, py, flat_poly):
        """Ray-casting point-in-polygon. flat_poly is [x0,y0,x1,y1,...] in canvas coords."""
        n = len(flat_poly) // 2
        inside = False
        j = n - 1
        for i in range(n):
            xi, yi = flat_poly[i * 2], flat_poly[i * 2 + 1]
            xj, yj = flat_poly[j * 2], flat_poly[j * 2 + 1]
            if ((yi > py) != (yj > py)) and (px < (xj - xi) * (py - yi) / (yj - yi) + xi):
                inside = not inside
            j = i
        return inside

    def hit_test_string(self, cx, cy):
        """Return (tracker_idx, s_idx) for the string under canvas coords, or None.

        Iterates _string_rects in reverse (topmost drawn wins). Unowned partial
        bands are excluded — they have no device owner to move.
        """
        for rect in reversed(self._string_rects):
            if rect['is_unowned_partial']:
                continue
            if self._point_in_polygon(cx, cy, rect['poly_canvas']):
                return rect['tracker_idx'], rect['s_idx']
        return None

    def hit_test_strings_in_box(self, cx1, cy1, cx2, cy2):
        """Return list of (tracker_idx, s_idx) whose polygon overlaps the canvas box.

        Uses the polygon's axis-aligned bounding box, so any contact with the
        selection rectangle is enough to select the string.
        """
        x_lo, x_hi = min(cx1, cx2), max(cx1, cx2)
        y_lo, y_hi = min(cy1, cy2), max(cy1, cy2)
        results = []
        for rect in self._string_rects:
            if rect['is_unowned_partial']:
                continue
            poly = rect['poly_canvas']
            n = len(poly) // 2
            px_vals = [poly[i * 2] for i in range(n)]
            py_vals = [poly[i * 2 + 1] for i in range(n)]
            # Overlap: neither box is fully outside the other on either axis
            if (max(px_vals) >= x_lo and min(px_vals) <= x_hi and
                    max(py_vals) >= y_lo and min(py_vals) <= y_hi):
                results.append((rect['tracker_idx'], rect['s_idx']))
        return results

    def _sync_tree_from_canvas(self):
        """Push current canvas string selection into the Edit Devices dialog tree (if open)."""
        tree = self._edit_dialog_tree
        if tree is None:
            return
        try:
            if not tree.winfo_exists():
                self._edit_dialog_tree = None
                self._edit_dialog_string_tracker = None
                return
        except Exception:
            self._edit_dialog_tree = None
            self._edit_dialog_string_tracker = None
            return

        string_tracker = self._edit_dialog_string_tracker or {}
        # Build reverse map: (tracker_idx, phys_pos) -> list of iids
        reverse = {}
        for iid, key in string_tracker.items():
            reverse.setdefault(key, []).append(iid)

        target_iids = []
        for key in self._highlighted_strings:
            target_iids.extend(reverse.get(key, []))

        self._canvas_syncing_tree = True
        try:
            tree.selection_set(target_iids)
        except Exception:
            pass
        finally:
            self._canvas_syncing_tree = False

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
        """Re-run string allocation using current group positions, then refresh preview.
        
        If allocation is locked, skips string reallocation and only updates
        device (CB/inverter) positions based on the current group layout.
        """
        self._save_group_positions()
        
        parent = self.master
        
        if self.allocation_locked:
            manually_edited = getattr(parent, 'manually_edited', False)
            detail = " (including manual device assignments)" if manually_edited else ""
            confirmed = messagebox.askyesno(
                "Refresh Allocation",
                f"Re-run allocation from scratch? All current device assignments{detail} will be replaced.\n\n"
                "Click No to keep the current lock and just refresh the display.",
                parent=self
            )
            if not confirmed:
                # Keep lock — just refresh positions and display
                self.build_layout_data()
                self._recolor_from_cb_assignments()
                self._build_legend()
                inv_summary = getattr(parent, 'last_totals', {}).get('inverter_summary', {})
                num_inv = inv_summary.get('total_inverters', 0)
                total_str = inv_summary.get('total_strings', 0)
                actual_ratio = inv_summary.get('actual_dc_ac', 0)
                split = inv_summary.get('total_split_trackers', 0)
                spatial_runs = inv_summary.get('allocation_result', {}).get('spatial_runs', 1)
                self.summary_label.config(
                    text=self._format_summary(num_inv, total_str, actual_ratio, split,
                                              spatial_runs=spatial_runs, locked=True)
                )
                self.draw()
                return
            # Confirmed — fall through to full reallocation below

        if hasattr(parent, 'calculate_estimate'):
            # Clear stale combiner assignments so calculate_estimate rebuilds them fresh
            parent.last_combiner_assignments = []
            parent.allocation_locked = False
            parent.locked_allocation_result = None
            parent.manually_edited = False
            self.allocation_locked = False
            self._tracker_physical_order = None
            self._invalidate_device_data()
            parent.calculate_estimate()
            
            inv_summary = getattr(parent, 'last_totals', {}).get('inverter_summary', {})
            if inv_summary and inv_summary.get('allocation_result'):
                self.inv_summary = inv_summary
                self.build_layout_data()
                self._recolor_from_cb_assignments()
                self._build_legend()
                
                # Update top bar summary
                num_inv = inv_summary.get('total_inverters', 0)
                total_str = inv_summary.get('total_strings', 0)
                actual_ratio = inv_summary.get('actual_dc_ac', 0)
                split = inv_summary.get('total_split_trackers', 0)
                spatial_runs = inv_summary.get('allocation_result', {}).get('spatial_runs', 1)
                self.summary_label.config(
                    text=self._format_summary(num_inv, total_str, actual_ratio, split,
                                              spatial_runs=spatial_runs, locked=self.allocation_locked)
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
            self._invalidate_device_data()
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

    def _invalidate_device_data(self):
        """Clear the shared device-string state so it is rebuilt on next dialog open or canvas edit."""
        self._device_data = None
        self._device_metadata = None

    def _check_overlaps(self):
        """Check for overlapping groups and return list of overlapping pair indices."""

        def _poly(g):
            gx = g['x']
            gy = g['y']
            w = g['width_ft']
            dt = g.get('driveline_tan', 0.0)
            vis_min_base = g.get('visual_min_y_base', g.get('visual_min_y', 0))
            vis_max_base = g.get('visual_max_y_base', g.get('visual_max_y', g['length_ft']))
            rd = g.get('rotation_deg', 0.0)
            vis_min = g.get('visual_min_y', vis_min_base)
            vis_max = g.get('visual_max_y', vis_max_base)
            rcx = gx + w / 2
            rcy = gy + (vis_min + vis_max) / 2
            # Parallelogram corners: left side uses base Y bounds; right side shifts by w*dt
            corners = [
                (gx,     gy + vis_min_base),
                (gx,     gy + vis_max_base),
                (gx + w, gy + vis_max_base + w * dt),
                (gx + w, gy + vis_min_base + w * dt),
            ]
            if rd:
                corners = [self._rotate_point(rcx, rcy, px, py, rd) for px, py in corners]
            return corners

        def _sat(a, b):
            for poly in (a, b):
                n = len(poly)
                for k in range(n):
                    x1, y1 = poly[k]
                    x2, y2 = poly[(k + 1) % n]
                    nx, ny = -(y2 - y1), (x2 - x1)
                    pa = [nx * px + ny * py for px, py in a]
                    pb = [nx * px + ny * py for px, py in b]
                    if max(pa) < min(pb) or max(pb) < min(pa):
                        return False
            return True

        polys = [_poly(g) for g in self.group_layout]
        overlaps = []
        for i in range(len(polys)):
            for j in range(i + 1, len(polys)):
                if _sat(polys[i], polys[j]):
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
            rotation_deg = g.get('rotation_deg', 0.0)
            rcx = g['x'] + g['width_ft'] / 2
            rcy = g['y'] + (vis_min + vis_max) / 2
            
            # Four corners of the padded outline in world coords, then rotate
            corners_world = [
                (g['x'] - pad,                 g['y'] + vis_min - pad),
                (g['x'] + g['width_ft'] + pad, g['y'] + vis_min - pad),
                (g['x'] + g['width_ft'] + pad, g['y'] + vis_max + pad),
                (g['x'] - pad,                 g['y'] + vis_max + pad),
            ]
            poly_canvas = []
            for wx, wy in corners_world:
                if rotation_deg != 0:
                    wx, wy = self._rotate_point(rcx, rcy, wx, wy, rotation_deg)
                cx, cy = self.world_to_canvas(wx, wy)
                poly_canvas.extend([cx, cy])
            
            self.canvas.create_polygon(
                *poly_canvas,
                fill='', outline='#FF0000', width=3, dash=(8, 4),
                tags='overlap_warning'
            )
            
            # Warning label — anchor to the top-center of the rotated outline
            label_wx = g['x'] + g['width_ft'] / 2
            label_wy = g['y'] + vis_min - pad
            if rotation_deg != 0:
                label_wx, label_wy = self._rotate_point(rcx, rcy, label_wx, label_wy, rotation_deg)
            label_x, label_y = self.world_to_canvas(label_wx, label_wy)
            label_y -= 8
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

            rcx, rcy, rd = self._device_rotation_info(dev)

            def _rot_dev(wx, wy, _rcx=rcx, _rcy=rcy, _rd=rd):
                if _rd:
                    wx, wy = self._rotate_point(_rcx, _rcy, wx, wy, _rd)
                return self.world_to_canvas(wx, wy)

            corners = [(dx, dy), (dx + dw, dy), (dx + dw, dy + dh), (dx, dy + dh)]
            poly_pts = []
            for wx, wy in corners:
                px2, py2 = _rot_dev(wx, wy)
                poly_pts.extend([px2, py2])

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

            self.canvas.create_polygon(
                *poly_pts, fill=fill_color, outline=outline_color, width=outline_width
            )

            # Label — above the rotated device center-top
            cx, cy_top = _rot_dev(dx + dw / 2, dy)
            font_size = max(7, min(14, int(10 * self.scale)))
            self._draw_text_with_bg(
                cx, cy_top - font_size - 2,
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
            if self.topology == 'Distributed String':
                inv_summary = getattr(parent_qe, 'last_totals', {}).get('inverter_summary', {})
                alloc = inv_summary.get('allocation_result')
                if not alloc or self.selected_device_idx >= len(alloc.get('inverters', [])):
                    return
                inv = alloc['inverters'][self.selected_device_idx]
                module_isc = 0.0
                nec_factor = 1.56
                if hasattr(parent_qe, 'selected_module') and parent_qe.selected_module:
                    module_isc = parent_qe.selected_module.isc
                if hasattr(parent_qe, 'current_project') and parent_qe.current_project:
                    nec_factor = getattr(parent_qe.current_project, 'nec_safety_factor', 1.56)
                connections = []
                for entry in inv.get('harness_map', []):
                    tidx = entry['tracker_idx']
                    seg_info = tracker_seg_map[tidx] if tidx < len(tracker_seg_map) else {}
                    wire_gauge = seg_info.get('wire_gauge', '10 AWG') if seg_info else '10 AWG'
                    connections.append({
                        'tracker_idx': tidx,
                        'tracker_label': f'T{tidx+1:02d}',
                        'harness_label': 'H01',
                        'num_strings': entry['strings_taken'],
                        'module_isc': module_isc,
                        'nec_factor': nec_factor,
                        'wire_gauge': wire_gauge,
                        'start_string_pos': entry.get('start_physical_pos', 0),
                    })
                cb_data = {
                    'combiner_name': dev_label,
                    'device_idx': self.selected_device_idx,
                    'breaker_size': None,
                    'module_isc': module_isc,
                    'nec_factor': nec_factor,
                    'connections': connections,
                }
            else:
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
                        t_x = grp['x'] + t_i * grp.get('row_spacing_ft', self.tracker_pitch_ft)
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

        # Anchor point on device (center-top), rotated with group
        anchor_wx = dev['x'] + dev['width_ft'] / 2
        anchor_wy = dev['y']
        rcx, rcy, rd = self._device_rotation_info(dev)
        if rd:
            anchor_wx, anchor_wy = self._rotate_point(rcx, rcy, anchor_wx, anchor_wy, rd)
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

    # ---------------------------------------------------------------------------
    # Measurement tool
    # ---------------------------------------------------------------------------

    def _toggle_measure_mode(self):
        """Activate or deactivate the measurement drawing tool."""
        self.measure_mode = not self.measure_mode
        if self.measure_mode:
            self.measure_btn.config(text="Stop Measuring")
            self.canvas.config(cursor='crosshair')
            self.current_measure_pts = []
            self.measure_mouse_pos = None
        else:
            self._measure_finish()
            self.measure_btn.config(text="Measure")
            self.canvas.config(cursor='')

    def _on_measure_motion(self, event):
        """Redraw the rubber-band line from the last placed vertex to the cursor."""
        if not self.measure_mode or not self.current_measure_pts:
            self.canvas.delete('measure_rubber')
            return
        self.measure_mouse_pos = (event.x, event.y)
        self.canvas.delete('measure_rubber')

        wx_last, wy_last = self.current_measure_pts[-1]
        cx_last, cy_last = self.world_to_canvas(wx_last, wy_last)
        wmx, wmy = self.canvas_to_world(event.x, event.y)
        seg_dist = math.hypot(wmx - wx_last, wmy - wy_last)

        self.canvas.create_line(
            cx_last, cy_last, event.x, event.y,
            fill='#E65100', width=2, dash=(4, 4), tags='measure_rubber'
        )

        prior_total = sum(
            math.hypot(self.current_measure_pts[i + 1][0] - self.current_measure_pts[i][0],
                       self.current_measure_pts[i + 1][1] - self.current_measure_pts[i][1])
            for i in range(len(self.current_measure_pts) - 1)
        )
        label = f"{seg_dist:.1f} ft"
        if prior_total > 0:
            label = f"{seg_dist:.1f} ft  (total: {prior_total + seg_dist:.1f} ft)"

        lx, ly = event.x + 12, event.y - 14
        tid = self.canvas.create_text(
            lx, ly, text=label, anchor='w',
            font=('Helvetica', 8, 'bold'), fill='#E65100',
            tags='measure_rubber'
        )
        bbox = self.canvas.bbox(tid)
        if bbox:
            self.canvas.create_rectangle(
                bbox[0] - 2, bbox[1] - 2, bbox[2] + 2, bbox[3] + 2,
                fill='white', outline='', width=0, tags='measure_rubber'
            )
            self.canvas.tag_raise(tid)

    def _measure_finish(self):
        """Save the current in-progress measurement (≥2 points) and clear the working set."""
        if len(self.current_measure_pts) >= 2:
            self.measurements.append(list(self.current_measure_pts))
        self.current_measure_pts = []
        self.measure_mouse_pos = None
        self.canvas.delete('measure_rubber')
        self.draw()

    def _measure_cancel(self):
        """Discard the current in-progress measurement without saving (Escape)."""
        if not self.current_measure_pts:
            return
        self.current_measure_pts = []
        self.measure_mouse_pos = None
        self.canvas.delete('measure_rubber')
        self.draw()

    def _measure_clear(self):
        """Remove all saved measurements."""
        self.measurements.clear()
        self.current_measure_pts = []
        self.measure_mouse_pos = None
        self.canvas.delete('measure_rubber')
        self.draw()

    def _draw_measurements(self):
        """Draw all saved measurements and the current in-progress segments."""
        if not self.show_measurements_var.get():
            return

        MEAS_COLOR = '#E65100'
        FONT = ('Helvetica', 8, 'bold')

        def _draw_polyline(pts, in_progress=False):
            if len(pts) < 2:
                return
            total_dist = 0.0
            cpx = [self.world_to_canvas(wx, wy) for wx, wy in pts]
            for i in range(len(pts) - 1):
                cx1, cy1 = cpx[i]
                cx2, cy2 = cpx[i + 1]
                self.canvas.create_line(
                    cx1, cy1, cx2, cy2,
                    fill=MEAS_COLOR, width=2,
                    dash=(4, 4) if in_progress else (),
                    tags='measurement'
                )
                dist = math.hypot(pts[i + 1][0] - pts[i][0], pts[i + 1][1] - pts[i][1])
                total_dist += dist
                mx, my = (cx1 + cx2) / 2, (cy1 + cy2) / 2
                self._draw_text_with_bg(
                    mx, my - 10, f"{dist:.1f} ft",
                    font=FONT, fill=MEAS_COLOR, bg='white'
                )
            for cx, cy in cpx:
                self.canvas.create_oval(
                    cx - 3, cy - 3, cx + 3, cy + 3,
                    fill=MEAS_COLOR, outline='#8B2500', tags='measurement'
                )
            if len(pts) > 2 and not in_progress:
                cx_last, cy_last = cpx[-1]
                self._draw_text_with_bg(
                    cx_last + 6, cy_last - 16,
                    f"Total: {total_dist:.1f} ft",
                    font=FONT, fill='#8B2500', bg='#FFF9C4', anchor='w'
                )

        for meas in self.measurements:
            _draw_polyline(meas)

        if self.current_measure_pts:
            _draw_polyline(self.current_measure_pts, in_progress=True)

    def _add_pad(self):
        """Enter pad placement mode — next click on canvas places a new pad."""
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
            self._draw_text_with_bg(
                cx, cy,
                text=label, font=('Helvetica', font_size, 'bold'),
                fill='white', bg='black'
            )
            
            # Device count subtitle
            num_assigned = len(pad.get('assigned_devices', []))
            if num_assigned > 0:
                sub_size = max(5, min(8, int(6 * self.scale)))
                self._draw_text_with_bg(
                    cx, cy + font_size + 2,
                    text=f"({num_assigned} devices)", font=('Helvetica', sub_size),
                    fill='#CCCCCC', bg='black'
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
            rcx, rcy, rd = self._device_rotation_info(dev)
            if rd:
                lwx, lwy = self._rotate_point(rcx, rcy, wx, wy, -rd)
            else:
                lwx, lwy = wx, wy
            if (dev['x'] <= lwx <= dev['x'] + dev['width_ft'] and
                    dev['y'] <= lwy <= dev['y'] + dev['height_ft']):
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
        self.after_idle(self.draw)

        num_devices = len(self.device_positions)
        dialog_height = max(320, min(640, 160 + num_devices * 22))
        dialog.geometry(f"720x{dialog_height}")
        dialog.minsize(600, 300)

        # Build pad label list
        pad_labels = [pad.get('label', f'Pad {i+1}') for i, pad in enumerate(self.pads)]

        # Reverse lookup: device_idx -> pad_idx (mutable, updated on each cell commit)
        device_to_pad = {}
        for pad_idx, pad in enumerate(self.pads):
            for dev_idx in pad.get('assigned_devices', []):
                device_to_pad[dev_idx] = pad_idx

        # Snapshot for "Undo All Changes"
        _snap_assigned = [list(pad.get('assigned_devices', [])) for pad in self.pads]
        _snap_feeder_sizes = dict(self.device_feeder_sizes)
        _snap_parallel_counts = dict(self.device_feeder_parallel_counts)

        # Topology-driven feeder column label, default value, and blanket flag
        from src.utils.cable_sizing import get_available_sizes
        parent_qe = self.master
        material = 'aluminum'
        if hasattr(parent_qe, 'wire_sizing'):
            material = parent_qe.wire_sizing.get('feeder_material', 'aluminum')
        feeder_sizes_list = get_available_sizes(material)

        if self.topology == 'Distributed String':
            feeder_col_label = "AC HR Size"
            default_feeder = getattr(parent_qe, 'wire_sizing', {}).get('ac_homerun', '')
            blanket_on = getattr(parent_qe, 'wire_sizing', {}).get('ac_homerun_blanket_enabled', False)
        else:
            feeder_col_label = "DC Fdr Size"
            default_feeder = getattr(parent_qe, 'wire_sizing', {}).get('dc_feeder', '')
            blanket_on = getattr(parent_qe, 'wire_sizing', {}).get('dc_feeder_blanket_enabled', False)

        if not default_feeder and feeder_sizes_list:
            default_feeder = feeder_sizes_list[0]

        # Instructions
        ttk.Label(dialog, text="Assign each device to a collection pad:",
                  font=('Helvetica', 10)).pack(anchor='w', padx=10, pady=(10, 0))
        ttk.Label(dialog, text="Tip: Double-click a cell to edit.  Shift+click to select a range, then right-click any cell to set all selected rows to a chosen value.",
                  font=('Helvetica', 8), foreground='gray').pack(anchor='w', padx=10, pady=(0, 5))

        # Button row packed first at bottom so it never gets squeezed by the expanding Treeview
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill='x', padx=10, pady=10, side='bottom')

        # Treeview
        tree_frame = ttk.Frame(dialog)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=(0, 0))

        _COL_NAMES = ('device', 'strings', 'group', 'feeder', 'parallel', 'pad')
        tree = ttk.Treeview(tree_frame, columns=_COL_NAMES, show='headings',
                            selectmode='extended', height=min(num_devices, 20))
        tree.heading('device',   text='Device')
        tree.heading('strings',  text='Strings')
        tree.heading('group',    text='Group')
        tree.heading('feeder',   text=feeder_col_label)
        tree.heading('parallel', text='Parallel')
        tree.heading('pad',      text='Pad')

        tree.column('device',   width=100, anchor='w',      minwidth=70)
        tree.column('strings',  width=60,  anchor='center', minwidth=40)
        tree.column('group',    width=110, anchor='w',      minwidth=70)
        tree.column('feeder',   width=110, anchor='center', minwidth=70)
        tree.column('parallel', width=70,  anchor='center', minwidth=50)
        tree.column('pad',      width=140, anchor='w',      minwidth=80)

        vsb = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        def _row_values(dev_idx):
            dev = self.device_positions[dev_idx]
            num_strings = sum(len(v) for v in dev.get('assigned_strings', {}).values())
            grp_idx = dev.get('group_idx', 0)
            grp_name = self.groups[grp_idx]['name'] if grp_idx < len(self.groups) else '?'
            feeder_val = default_feeder if blanket_on else self.device_feeder_sizes.get(dev_idx, default_feeder)
            parallel_val = self.device_feeder_parallel_counts.get(dev_idx, 1)
            pad_idx = device_to_pad.get(dev_idx, 0)
            if pad_idx >= len(pad_labels):
                pad_idx = 0
            return (dev['label'], str(num_strings), grp_name, feeder_val, str(parallel_val), pad_labels[pad_idx])

        for dev_idx in range(num_devices):
            tree.insert('', 'end', iid=f"dev_{dev_idx}", values=_row_values(dev_idx))
        if num_devices > 0:
            tree.selection_set('dev_0')
            tree.focus('dev_0')

        # ---- Inline cell editor ----
        _active_editor = [None]

        def _destroy_editor():
            if _active_editor[0] is not None:
                try:
                    _active_editor[0].destroy()
                except tk.TclError:
                    pass
                _active_editor[0] = None

        def _on_double_click(event):
            _destroy_editor()
            if tree.identify_region(event.x, event.y) != 'cell':
                return
            col = tree.identify_column(event.x)
            row_iid = tree.identify_row(event.y)
            if not row_iid:
                return
            col_index = int(col.replace('#', '')) - 1
            if col_index in (0, 1, 2):
                return  # device / strings / group are read-only
            dev_idx = int(row_iid.replace('dev_', ''))
            bbox = tree.bbox(row_iid, col)
            if not bbox:
                return
            bx, by, bw, bh = bbox

            if col_index == 3:
                if blanket_on:
                    return
                var = tk.StringVar(value=self.device_feeder_sizes.get(dev_idx, default_feeder))
                editor = ttk.Combobox(tree, textvariable=var, values=feeder_sizes_list, state='readonly')
                editor.place(x=bx, y=by, width=bw, height=bh)
                editor.focus_set()
                _active_editor[0] = editor

                def _commit_feeder(e=None, _idx=dev_idx, _var=var, _ed=editor):
                    if _active_editor[0] is not _ed:
                        return
                    val = _var.get()
                    self.device_feeder_sizes[_idx] = val
                    tree.set(f"dev_{_idx}", 'feeder', val)
                    _destroy_editor()
                    self._schedule_redraw()

                editor.bind('<<ComboboxSelected>>', _commit_feeder)
                editor.bind('<Return>',   _commit_feeder)
                editor.bind('<Tab>',      _commit_feeder)
                editor.bind('<FocusOut>', _commit_feeder)
                editor.bind('<Escape>',   lambda e: _destroy_editor())

            elif col_index == 4:
                var = tk.StringVar(value=str(self.device_feeder_parallel_counts.get(dev_idx, 1)))
                editor = ttk.Spinbox(tree, textvariable=var, from_=1, to=10, increment=1)
                editor.place(x=bx, y=by, width=bw, height=bh)
                editor.focus_set()
                _active_editor[0] = editor

                def _commit_parallel(e=None, _idx=dev_idx, _var=var, _ed=editor):
                    if _active_editor[0] is not _ed:
                        return
                    try:
                        pval = max(1, min(10, int(_var.get())))
                    except (ValueError, TypeError):
                        pval = 1
                    self.device_feeder_parallel_counts[_idx] = pval
                    tree.set(f"dev_{_idx}", 'parallel', str(pval))
                    _destroy_editor()
                    self._schedule_redraw()

                editor.bind('<Return>',   _commit_parallel)
                editor.bind('<Tab>',      _commit_parallel)
                editor.bind('<FocusOut>', _commit_parallel)
                editor.bind('<Escape>',   lambda e: _destroy_editor())

            else:  # col_index == 5  (pad)
                old_pad_idx = device_to_pad.get(dev_idx, 0)
                if old_pad_idx >= len(pad_labels):
                    old_pad_idx = 0
                var = tk.StringVar(value=pad_labels[old_pad_idx])
                editor = ttk.Combobox(tree, textvariable=var, values=pad_labels, state='readonly')
                editor.place(x=bx, y=by, width=bw, height=bh)
                editor.focus_set()
                _active_editor[0] = editor

                def _commit_pad(e=None, _idx=dev_idx, _var=var, _ed=editor, _old=old_pad_idx):
                    if _active_editor[0] is not _ed:
                        return
                    new_label = _var.get()
                    new_pad_idx = next((i for i, lbl in enumerate(pad_labels) if lbl == new_label), 0)
                    # Incremental update — remove from old pad, add to new
                    old_assigned = self.pads[_old].get('assigned_devices', [])
                    if _idx in old_assigned:
                        old_assigned.remove(_idx)
                    if _idx not in self.pads[new_pad_idx].get('assigned_devices', []):
                        self.pads[new_pad_idx].setdefault('assigned_devices', []).append(_idx)
                    device_to_pad[_idx] = new_pad_idx
                    tree.set(f"dev_{_idx}", 'pad', new_label)
                    _destroy_editor()
                    self._schedule_redraw()

                editor.bind('<<ComboboxSelected>>', _commit_pad)
                editor.bind('<Return>',   _commit_pad)
                editor.bind('<Tab>',      _commit_pad)
                editor.bind('<FocusOut>', _commit_pad)
                editor.bind('<Escape>',   lambda e: _destroy_editor())

        tree.bind('<Double-1>', _on_double_click)

        # ---- Right-click context menu: set selection to a chosen value ----
        # Workflow: click a row (or Shift+click for a range), right-click any cell
        # in an editable column, pick the value to apply to all selected rows.
        context_menu = tk.Menu(dialog, tearoff=0)
        _ctx = {'col_index': None}

        def _apply_to_selection(value):
            col_index = _ctx['col_index']
            if col_index is None:
                return
            col_name = _COL_NAMES[col_index]
            for iid in tree.selection():
                i = int(iid.replace('dev_', ''))
                if col_index == 3:
                    self.device_feeder_sizes[i] = value
                    tree.set(iid, col_name, value)
                elif col_index == 4:
                    try:
                        pval = max(1, min(10, int(value)))
                    except (ValueError, TypeError):
                        pval = 1
                    self.device_feeder_parallel_counts[i] = pval
                    tree.set(iid, col_name, str(pval))
                elif col_index == 5:
                    new_pad_idx = next((pi for pi, lbl in enumerate(pad_labels) if lbl == value), 0)
                    old_pad_idx = device_to_pad.get(i, 0)
                    old_assigned = self.pads[old_pad_idx].get('assigned_devices', [])
                    if i in old_assigned:
                        old_assigned.remove(i)
                    if i not in self.pads[new_pad_idx].get('assigned_devices', []):
                        self.pads[new_pad_idx].setdefault('assigned_devices', []).append(i)
                    device_to_pad[i] = new_pad_idx
                    tree.set(iid, col_name, value)
            self._schedule_redraw()

        def _on_right_click(event):
            _destroy_editor()
            if tree.identify_region(event.x, event.y) != 'cell':
                return
            col = tree.identify_column(event.x)
            row_iid = tree.identify_row(event.y)
            if not row_iid:
                return
            col_index = int(col.replace('#', '')) - 1
            if col_index in (0, 1, 2):
                return
            if blanket_on and col_index == 3:
                return
            if row_iid not in tree.selection():
                tree.selection_set(row_iid)
            _ctx['col_index'] = col_index
            n = len(tree.selection())

            if col_index == 3:
                choices = feeder_sizes_list
            elif col_index == 4:
                choices = [str(i) for i in range(1, 11)]
            else:
                choices = pad_labels

            context_menu.delete(0, 'end')
            context_menu.add_command(
                label=f"Set {n} row{'s' if n != 1 else ''} to:",
                state='disabled'
            )
            context_menu.add_separator()
            for val in choices:
                context_menu.add_command(
                    label=val,
                    command=lambda v=val: _apply_to_selection(v)
                )
            try:
                context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                context_menu.grab_release()

        tree.bind('<Button-3>', _on_right_click)

        # ---- Undo All Changes ----
        def _undo_all():
            _destroy_editor()
            for pad_idx, pad in enumerate(self.pads):
                pad['assigned_devices'] = list(_snap_assigned[pad_idx])
            self.device_feeder_sizes.clear()
            self.device_feeder_sizes.update(_snap_feeder_sizes)
            self.device_feeder_parallel_counts.clear()
            self.device_feeder_parallel_counts.update(_snap_parallel_counts)
            device_to_pad.clear()
            for pad_idx, assigned in enumerate(_snap_assigned):
                for dev_idx in assigned:
                    device_to_pad[dev_idx] = pad_idx
            for dev_idx in range(num_devices):
                tree.item(f"dev_{dev_idx}", values=_row_values(dev_idx))
            self.draw()

        # ---- Close handler ----
        def _on_close():
            self.assigning_devices = False
            self.draw()
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", _on_close)

        ttk.Button(btn_frame, text="Undo All Changes", command=_undo_all).pack(side='right')

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

    def _ensure_device_data(self):
        """Build and cache device data if not already present.

        Returns True if data is available (either pre-existing or freshly built),
        False if there is no allocation data to work from.
        """
        if self._device_data is not None:
            return bool(self._device_data)

        raw_data, metadata = self._normalize_to_device_strings()
        if not raw_data:
            return False

        device_data = raw_data

        for dev_idx, custom_name in self.device_names.items():
            if dev_idx < len(device_data):
                device_data[dev_idx]['name'] = custom_name

        all_tracker_spt = {}
        _g_idx = 0
        for _grp in self.groups:
            for _seg in _grp.get('segments', []):
                _ref = _seg.get('template_ref')
                _raw = 1
                if _ref and _ref in self.enabled_templates:
                    _raw = self.enabled_templates[_ref].get('strings_per_tracker', 1)
                _spt = int(_raw) + (1 if _raw != int(_raw) else 0)
                for _ in range(_seg.get('quantity', 0)):
                    all_tracker_spt[_g_idx] = _spt
                    _g_idx += 1
        metadata['tracker_spt'].update(
            {k: v for k, v in all_tracker_spt.items() if k not in metadata['tracker_spt']}
        )

        covered_trackers = {tidx for dev in device_data for tidx, _ in dev['strings']}
        unallocated_strings = []
        for tidx, spt in sorted(all_tracker_spt.items()):
            if tidx not in covered_trackers:
                for phys_pos in range(spt):
                    unallocated_strings.append((tidx, phys_pos))

        if unallocated_strings:
            device_data.insert(0, {
                'name': 'Unallocated',
                'strings': unallocated_strings,
                'is_unallocated': True,
            })

        self._device_data = device_data
        self._device_metadata = metadata
        return True

    def _show_edit_devices_dialog(self):
        """Show dialog to reassign string-level connections between devices."""
        if not self._ensure_device_data():
            messagebox.showinfo("No Data", "Run Calculate Estimate first.", parent=self)
            return

        device_data = self._device_data
        metadata = self._device_metadata

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

        # Lookup: tree item iid -> (tracker_idx, phys_pos)
        string_tracker = {}
        self._edit_dialog_string_tracker = string_tracker  # kept alive for canvas sync

        def _regroup(dev):
            """Sort strings by (tracker_idx, physical_pos) to keep contiguous."""
            dev['strings'].sort(key=lambda s: (s[0], s[1]))

        def _sort_device_data():
            """Sort device_data in-place by natural name order; keeps unallocated device at top."""
            import re
            def _natural_key(dev):
                return [int(c) if c.isdigit() else c.lower()
                        for c in re.split(r'(\d+)', dev['name'])]
            has_ua = device_data and device_data[0].get('is_unallocated')
            if has_ua:
                rest = device_data[1:]
                rest.sort(key=_natural_key)
                device_data[1:] = rest
            else:
                device_data.sort(key=_natural_key)

        def _validate_move(source_dev_idx, positions_to_move, target_dev_idx):
            """Check that moving strings preserves contiguity on both sides.
            Gaps covered entirely by the Unallocated device are tolerated.

            Returns (ok, error_message).
            """
            # Build a map of positions currently held by the Unallocated device
            unalloc_positions = {}
            for dev in device_data:
                if dev.get('is_unallocated'):
                    for t, p in dev['strings']:
                        unalloc_positions.setdefault(t, set()).add(p)

            # Group positions being moved by tracker
            moving_by_tracker = {}
            for pos_idx in positions_to_move:
                if pos_idx < len(device_data[source_dev_idx]['strings']):
                    tidx, phys = device_data[source_dev_idx]['strings'][pos_idx]
                    moving_by_tracker.setdefault(tidx, set()).add(phys)

            for tidx, moving_positions in moving_by_tracker.items():
                unalloc_t = unalloc_positions.get(tidx, set())

                # Current positions this tracker has on source device
                source_positions = {p for t, p in device_data[source_dev_idx]['strings'] if t == tidx}
                remaining = source_positions - moving_positions

                # Check source contiguity after removal
                if len(remaining) > 1:
                    missing = set(range(min(remaining), max(remaining) + 1)) - remaining
                    if not missing.issubset(unalloc_t):
                        return False, (f"Cannot move: would leave a gap in T{tidx+1:02d} "
                                       f"on {device_data[source_dev_idx]['name']}")

                # Current positions this tracker has on target device
                target_positions = {p for t, p in device_data[target_dev_idx]['strings'] if t == tidx}
                combined = target_positions | moving_positions

                # Check target contiguity after addition
                if len(combined) > 1:
                    missing = set(range(min(combined), max(combined) + 1)) - combined
                    if not missing.issubset(unalloc_t):
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
        self._edit_dialog_tree = tree  # kept alive for canvas sync

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
            ua_dev = next((d for d in device_data if d.get('is_unallocated')), None)
            ua_count = len(ua_dev['strings']) if ua_dev else 0
            real_devs = [d for d in device_data if not d.get('is_unallocated')]
            total_devs = len(real_devs)
            total_strings = sum(len(d['strings']) for d in real_devs)
            tracker_dev_map = {}
            for dev_idx, dev in enumerate(real_devs):
                for tidx, _pos in dev['strings']:
                    tracker_dev_map.setdefault(tidx, set()).add(dev_idx)
            splits = sum(1 for devs in tracker_dev_map.values() if len(devs) > 1)
            if ua_count:
                summary_var.set(
                    f"{total_devs} devices  |  {total_strings} strings allocated"
                    f"  |  {ua_count} unallocated"
                )
            else:
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

            _has_ua = device_data and device_data[0].get('is_unallocated')
            _sort_key = lambda d: [int(c) if c.isdigit() else c.lower()
                                   for c in _sort_re.split(r'(\d+)', d['name'])]
            if _has_ua:
                _rest = device_data[1:]
                _rest.sort(key=_sort_key)
                device_data[1:] = _rest
            else:
                device_data.sort(key=_sort_key)

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
                is_ua = dev.get('is_unallocated', False)
                is_open = open_state.get(dev_idx, True if is_ua else False)
                label = f"{dev['name']}  ({total} strings)"
                node_tags = ('unallocated',) if is_ua else ()
                tree.insert('', 'end',
                            iid=dev_iid,
                            text=label,
                            open=is_open,
                            tags=node_tags)

                for s_pos, (tidx, phys_pos) in enumerate(dev['strings']):
                    str_iid = f'dev_{dev_idx}_s_{s_pos}'
                    tree.insert(dev_iid, 'end', iid=str_iid,
                                text=f"  T{tidx+1:02d} - S{phys_pos+1:02d}",
                                tags=node_tags)
                    string_tracker[str_iid] = (tidx, phys_pos)

            if select_device is not None:
                dev_iid = f'dev_{select_device}'
                if dev_iid in tree.get_children():
                    tree.selection_set(dev_iid)
                    tree.see(dev_iid)

            _update_summary()

        self._edit_dialog_refresh_tree = refresh_tree  # canvas drag calls this after moves

        # --- Selection → highlight on canvas ---
        def on_tree_select(event=None):
            if drag_state.get('suppressing') or self._canvas_syncing_tree:
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
        tree.tag_configure('unallocated', foreground='#B71C1C')

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

            # Validate contiguity for each source (skip when moving to/from unallocated pool)
            for src_dev, positions in by_source.items():
                if src_dev == target_dev:
                    continue
                if (device_data[src_dev].get('is_unallocated') or
                        device_data[target_dev].get('is_unallocated')):
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
                dev_idx = int(iid.split('_')[1])
                if device_data[dev_idx].get('is_unallocated'):
                    return 'break'
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
                if device_data[idx].get('is_unallocated'):
                    continue
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

        def auto_number():
            import re as _re
            from statistics import median as _median

            # Build spatial list for non-unallocated devices only
            real_devs = [(i, d) for i, d in enumerate(device_data) if not d.get('is_unallocated')]
            if not real_devs:
                return

            # Map device_data index → device_positions index (real devs in order)
            # device_positions is in the same order as real_devs (unallocated excluded)
            pos_list = getattr(self, 'device_positions', [])
            entries = []
            for pos_order, (data_idx, _) in enumerate(real_devs):
                if pos_order < len(pos_list):
                    dev_pos = pos_list[pos_order]
                    entries.append((data_idx, dev_pos['x'], dev_pos['y'], dev_pos['height_ft']))

            if not entries:
                return

            # Row tolerance = half the device height
            tol = _median(e[3] for e in entries) / 2

            # Sort into rows: group by Y within tol, then sort each row by X
            entries.sort(key=lambda e: (e[2], e[1]))
            rows = []
            for entry in entries:
                if rows and abs(entry[2] - rows[-1][0][2]) <= tol:
                    rows[-1].append(entry)
                else:
                    rows.append([entry])
            for row in rows:
                row.sort(key=lambda e: e[1])
            sorted_data_indices = [e[0] for row in rows for e in row]

            # Guess most common existing prefix (strip trailing digits, take mode)
            default_prefix = 'INV-' if self.topology == 'Distributed String' else 'CB-'
            prefix_counts = {}
            for _, d in real_devs:
                p = _re.sub(r'\d+$', '', d['name'])
                if p:
                    prefix_counts[p] = prefix_counts.get(p, 0) + 1
            if prefix_counts:
                default_prefix = max(prefix_counts, key=prefix_counts.get)

            # Open naming dialog
            num_dlg = tk.Toplevel(dialog)
            num_dlg.title("Auto-Number Devices")
            num_dlg.transient(dialog)
            num_dlg.resizable(False, False)

            frm = ttk.Frame(num_dlg, padding=12)
            frm.pack(fill='both', expand=True)

            ttk.Label(frm, text=f"Number {len(sorted_data_indices)} devices top-left → bottom-right:").grid(
                row=0, column=0, columnspan=2, sticky='w', pady=(0, 8))

            ttk.Label(frm, text="Prefix:").grid(row=1, column=0, sticky='w', padx=(0, 8), pady=4)
            prefix_var = tk.StringVar(value=default_prefix)
            ttk.Entry(frm, textvariable=prefix_var, width=12).grid(row=1, column=1, sticky='w')

            ttk.Label(frm, text="Start #:").grid(row=2, column=0, sticky='w', padx=(0, 8), pady=4)
            start_var = tk.IntVar(value=1)
            ttk.Spinbox(frm, from_=1, to=999, textvariable=start_var, width=6).grid(row=2, column=1, sticky='w')

            btn_frame = ttk.Frame(frm)
            btn_frame.grid(row=3, column=0, columnspan=2, pady=(10, 0))

            def _do_autonumber():
                try:
                    start = int(start_var.get())
                except (ValueError, tk.TclError):
                    return
                prefix = prefix_var.get()

                # Snapshot device-object → pad mapping before any reordering
                real_devs_before = [d for d in device_data if not d.get('is_unallocated')]
                obj_to_pad = {}
                for pad_idx, pad in enumerate(getattr(self, 'pads', [])):
                    for pos_idx in pad.get('assigned_devices', []):
                        if pos_idx < len(real_devs_before):
                            obj_to_pad[id(real_devs_before[pos_idx])] = pad_idx

                for n, data_idx in enumerate(sorted_data_indices):
                    device_data[data_idx]['name'] = f'{prefix}{start + n:02d}'
                _sort_device_data()

                # Rebuild pad assignments using object identity after sort
                if obj_to_pad and hasattr(self, 'pads'):
                    new_real_devs = [d for d in device_data if not d.get('is_unallocated')]
                    for pad in self.pads:
                        pad['assigned_devices'] = []
                    for new_pos_idx, d in enumerate(new_real_devs):
                        pad_idx = obj_to_pad.get(id(d))
                        if pad_idx is not None and pad_idx < len(self.pads):
                            self.pads[pad_idx]['assigned_devices'].append(new_pos_idx)
                    for pad in self.pads:
                        pad['assigned_devices'].sort()

                num_dlg.destroy()
                refresh_tree()
                _update_live_preview()

            ttk.Button(btn_frame, text="Apply", command=_do_autonumber).pack(side='left', padx=4)
            ttk.Button(btn_frame, text="Cancel", command=num_dlg.destroy).pack(side='left', padx=4)

            num_dlg.bind('<Return>', lambda e: _do_autonumber())
            num_dlg.bind('<Escape>', lambda e: num_dlg.destroy())

            num_dlg.update_idletasks()
            x = dialog.winfo_rootx() + (dialog.winfo_width() - num_dlg.winfo_width()) // 2
            y = dialog.winfo_rooty() + (dialog.winfo_height() - num_dlg.winfo_height()) // 2
            num_dlg.geometry(f'+{x}+{y}')
            num_dlg.grab_set()
            num_dlg.focus_set()

        ttk.Label(action_frame, text="Drag strings to move  |",
                  foreground='gray', font=('Helvetica', 9)).pack(side='left', padx=(0, 8))
        ttk.Button(action_frame, text="Add Device(s)", command=add_device).pack(side='left', padx=2)
        ttk.Button(action_frame, text="Remove", command=remove_device).pack(side='left', padx=2)
        ttk.Button(action_frame, text="Rename", command=rename_device).pack(side='left', padx=2)

        # --- Live preview update / autosave ---
        def _update_live_preview():
            # Commit changes immediately — every mutation autosaves to parent state
            real_devs = [d for d in device_data if not d.get('is_unallocated')]
            self.num_devices = len(real_devs)
            self.device_names = {i: dev['name'] for i, dev in enumerate(real_devs)}

            self._rebuild_from_device_strings(real_devs, metadata)
            self._update_lock_button()
            self._highlighted_strings = set()
            if hasattr(self.master, 'manually_edited'):
                self.master.manually_edited = True

            inv_summary = getattr(self.master, 'last_totals', {}).get('inverter_summary', {})
            if inv_summary and inv_summary.get('allocation_result'):
                self.inv_summary = inv_summary
                alloc = inv_summary.get('allocation_result', {})
                num_inv = alloc.get('summary', {}).get('total_inverters', 0)
                total_str = alloc.get('summary', {}).get('total_strings', 0)
                split = alloc.get('summary', {}).get('total_split_trackers', 0)
                spatial_runs = alloc.get('spatial_runs', 1)
                actual_ratio = inv_summary.get('actual_dc_ac', 0)
                self.summary_label.config(
                    text=self._format_summary(num_inv, total_str, actual_ratio, split,
                                              spatial_runs=spatial_runs, locked=True)
                )
            self.build_layout_data()
            self._recolor_from_cb_assignments()
            self._build_legend()
            self.draw()

        # --- Undo All Changes ---
        bottom_frame = ttk.Frame(dialog, padding="10")
        bottom_frame.pack(fill='x')

        def undo_changes():
            # Restore full parent state from the snapshot taken at dialog open
            self.allocation_locked = original_locked
            self.master.allocation_locked = original_parent_locked
            self.master.locked_allocation_result = original_locked_result
            self.master.last_totals = original_last_totals
            self.master.last_combiner_assignments = original_combiner_assignments
            self._tracker_physical_order = original_physical_order
            self._update_lock_button()
            self._highlighted_strings = set()

            # Restore device_data in-place so the tree can refresh from it
            device_data.clear()
            device_data.extend(copy.deepcopy(original))

            if original_inv_summary is not None:
                self.inv_summary = original_inv_summary
            self.build_layout_data()
            self._recolor_from_cb_assignments()
            self._build_legend()
            self.draw()
            refresh_tree()

        def _on_dialog_close():
            self._edit_dialog_tree = None
            self._edit_dialog_string_tracker = None
            self._edit_dialog_refresh_tree = None
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", _on_dialog_close)

        ttk.Button(bottom_frame, text="Undo All Changes", command=undo_changes).pack(side='right', padx=2)
        ttk.Button(bottom_frame, text="Collapse All",
                   command=lambda: [tree.item(c, open=False) for c in tree.get_children()]).pack(side='left', padx=2)
        ttk.Button(bottom_frame, text="Expand All",
                   command=lambda: [tree.item(c, open=True) for c in tree.get_children()]).pack(side='left', padx=2)
        ttk.Button(bottom_frame, text="Auto-Number", command=auto_number).pack(side='left', padx=2)

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
        if self.measure_mode:
            self._measure_finish()
            return
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
            
            # Device center — apply group rotation
            dev_cx = dev['x'] + dev['width_ft'] / 2
            dev_cy = dev['y'] + dev['height_ft'] / 2
            rcx, rcy, rd = self._device_rotation_info(dev)
            if rd:
                dev_cx, dev_cy = self._rotate_point(rcx, rcy, dev_cx, dev_cy, rd)

            # Pad center
            pad_cx = pad['x'] + pad.get('width_ft', 10.0) / 2
            pad_cy = pad['y'] + pad.get('height_ft', 8.0) / 2

            # Row direction: E-W tilted by driveline_angle, then rotated by azimuth
            grp_idx = dev.get('group_idx')
            grp = self.group_layout[grp_idx] if (grp_idx is not None and grp_idx < len(self.group_layout)) else {}
            driveline_tan = grp.get('driveline_tan', 0.0)
            rotation_deg = grp.get('rotation_deg', 0.0)
            mag = math.sqrt(1.0 + driveline_tan ** 2)
            dx_u, dy_u = 1.0 / mag, driveline_tan / mag
            if rotation_deg:
                cos_r = math.cos(math.radians(rotation_deg))
                sin_r = math.sin(math.radians(rotation_deg))
                row_dx = dx_u * cos_r - dy_u * sin_r
                row_dy = dx_u * sin_r + dy_u * cos_r
            else:
                row_dx, row_dy = dx_u, dy_u

            # Corner: projection of pad onto the line through device in row direction
            t = (pad_cx - dev_cx) * row_dx + (pad_cy - dev_cy) * row_dy
            corner_x = dev_cx + t * row_dx
            corner_y = dev_cy + t * row_dy

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

            # Leg 1: along the row/driveline direction
            self.canvas.create_line(
                cx1, cy1, cx_corner, cy_corner,
                fill=color, width=line_width, dash=dash_pattern
            )

            # Leg 2: from turn point to pad
            self.canvas.create_line(
                cx_corner, cy_corner, cx2, cy2,
                fill=color, width=line_width, dash=dash_pattern
            )

            # Show distance label if this device is selected in inspect mode
            if self.inspect_mode and self.selected_device_idx == dev_idx:
                leg1 = abs(t)
                leg2 = math.sqrt((pad_cx - corner_x) ** 2 + (pad_cy - corner_y) ** 2)
                total_dist = leg1 + leg2

                # Place label at the corner of the L
                font_size = max(6, min(10, int(8 * self.scale)))
                self._draw_text_with_bg(
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