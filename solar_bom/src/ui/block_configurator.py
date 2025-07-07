import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Dict, List
from ..models.block import BlockConfig, WiringType, TrackerPosition, DeviceType, CollectionPoint, WiringConfig, HarnessGroup
from ..models.tracker import TrackerTemplate, TrackerPosition, ModuleOrientation
from ..models.inverter import InverterSpec
from .inverter_manager import InverterManager
from pathlib import Path
import json
from ..models.module import ModuleSpec, ModuleType, ModuleOrientation
from ..utils.undo_manager import UndoManager
from copy import deepcopy

class BlockConfigurator(ttk.Frame):
    def __init__(self, parent, current_project=None, on_autosave=None):
        super().__init__(parent)
        self.parent = parent
        self.current_project = current_project

        self.updating_ui = False  # Flag to prevent recursive updates
        
        # State management
        self.device_type_var = None
        self.inverter_frame = None
        self._current_module = None
        self.available_templates = {}
        self.blocks: Dict[str, BlockConfig] = {}  # Store block configurations
        self.current_block: Optional[str] = None  # Currently selected block ID
        self.most_recent_block = None  # Track most recently created/copied block
        self.available_templates: Dict[str, TrackerTemplate] = {}  # Available tracker templates
        self.selected_inverter = None
        self.tracker_templates: Dict[str, TrackerTemplate] = {}
        self.dragging = False
        self.drag_template = None
        self.drag_start = None
        self.selected_tracker = None  # Store currently selected tracker
        self.grid_lines = []  # Store grid line IDs for cleanup 
        self.scale_factor = 10.0  # Starting scale (10 pixels per meter)
        self.pan_x = 0  # Pan offset in pixels
        self.pan_y = 0
        self.panning = False
        self.pan_start_x = 0
        self.pan_start_y = 0
        self.inverters = {}  # Store inverter configurations
        self.on_blocks_changed = None  # Callback for when blocks change
        self.on_autosave = on_autosave
        self.device_placement_mode = tk.StringVar(value="row_center")  # Default to row center

        # Initialize undo manager
        self.undo_manager = UndoManager()
        self.undo_manager.set_callbacks(
            get_state=self._get_current_state,
            set_state=self._restore_from_state
        )
        
        # First set up the UI
        self.setup_ui()
        
        # Then load and update templates
        self.load_templates()
        self.update_template_list()

    @property
    def current_module(self):
        return self._current_module
    
    @current_module.setter
    def current_module(self, module):
        self._current_module = module
        self.update_template_list()  # Refresh templates with new module


    def update_current_labels(self):
        """Update the string current and NEC current display labels"""
        if self.drag_template and self.drag_template.module_spec:
            module_spec = self.drag_template.module_spec
            
            # Update string current (Imp)
            self.string_current = module_spec.imp
            self.string_current_label.config(text=f"{self.string_current:.2f} A")
            
            # Update NEC current (Isc × 1.25)
            self.nec_current = module_spec.isc * 1.25
            self.nec_current_label.config(text=f"{self.nec_current:.2f} A")
        else:
            # No template selected - show dashes
            self.string_current = 0.0
            self.nec_current = 0.0
            self.string_current_label.config(text="-- A")
            self.nec_current_label.config(text="-- A")

    def validate_float_input(self, var, default_value):
        """Validate float input, returning default value if invalid"""
        try:
            value = var.get().strip()
            return float(value) if value else default_value
        except ValueError:
            return default_value

    def setup_ui(self):
        """Create and arrange UI components"""
        # Main container with padding
        main_container = ttk.Frame(self, padding="10")
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights for expansion
        main_container.grid_columnconfigure(2, weight=1)  # Make column 2 (canvas) expand
        main_container.grid_rowconfigure(0, weight=1)     # Make rows expand
        
        # Left side - Block List and Controls
        list_frame = ttk.LabelFrame(main_container, text="Blocks", padding="5")
        list_frame.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Block listbox
        self.block_listbox = tk.Listbox(list_frame, width=30, height=10)
        self.block_listbox.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.block_listbox.bind('<<ListboxSelect>>', self.on_block_select)
        
        # Block control buttons
        btn_frame = ttk.Frame(list_frame)
        btn_frame.grid(row=1, column=0, padx=5, pady=5)
        
        ttk.Button(btn_frame, text="New Block", command=self.create_new_block).grid(row=0, column=0, padx=2)
        ttk.Button(btn_frame, text="Delete Block", command=self.delete_block).grid(row=0, column=1, padx=2)
        ttk.Button(btn_frame, text="Copy Block", command=self.copy_block).grid(row=0, column=2, padx=2)
        
        # Right side - Block Configuration
        config_frame = ttk.LabelFrame(main_container, text="Block Configuration", padding="5")
        config_frame.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Block ID
        ttk.Label(config_frame, text="Block ID:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.block_id_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.block_id_var).grid(row=0, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        ttk.Button(config_frame, text="Rename", command=self.rename_block).grid(row=0, column=2, padx=5, pady=2)

        # Spacing Configuration
        spacing_frame = ttk.LabelFrame(config_frame, text="Spacing Configuration", padding="5")
        spacing_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))

        # Row Spacing
        ttk.Label(spacing_frame, text="Row Spacing (ft):").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.row_spacing_var = tk.StringVar(value="19.7")  # 6m in feet
        row_spacing_entry = ttk.Entry(spacing_frame, textvariable=self.row_spacing_var)
        row_spacing_entry.grid(row=0, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        row_spacing_entry.bind('<FocusOut>', self.update_gcr_from_row_spacing)
        row_spacing_entry.bind('<Return>', self.update_gcr_from_row_spacing)

        # GCR (Ground Coverage Ratio) - input field
        ttk.Label(spacing_frame, text="GCR:").grid(row=0, column=2, padx=5, pady=2, sticky=tk.W)
        self.gcr_var = tk.StringVar(value="--")
        gcr_entry = ttk.Entry(spacing_frame, textvariable=self.gcr_var, width=10)
        gcr_entry.grid(row=0, column=3, padx=5, pady=2, sticky=(tk.W, tk.E))
        gcr_entry.bind('<FocusOut>', self.update_row_spacing_from_gcr)
        gcr_entry.bind('<Return>', self.update_row_spacing_from_gcr)

        # N/S Tracker Spacing
        ttk.Label(spacing_frame, text="N/S Tracker Spacing (m):").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.ns_spacing_var = tk.StringVar(value="1.0")
        ns_spacing_entry = ttk.Entry(spacing_frame, textvariable=self.ns_spacing_var)
        ns_spacing_entry.grid(row=1, column=1, columnspan=3, padx=5, pady=2, sticky=(tk.W, tk.E))

        # Device Configuration
        device_frame = ttk.LabelFrame(config_frame, text="Downstream Device", padding="5")
        device_frame.grid(row=5, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))

        # Device Type Selection
        ttk.Label(device_frame, text="Device Type:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.device_type_var = tk.StringVar(value=DeviceType.STRING_INVERTER.value)
        device_type_combo = ttk.Combobox(device_frame, textvariable=self.device_type_var, state='readonly')
        device_type_combo['values'] = [t.value for t in DeviceType]
        device_type_combo.grid(row=0, column=1, columnspan=2, padx=5, pady=2, sticky=(tk.W, tk.E))

        # Number of Inputs
        ttk.Label(device_frame, text="Number of Inputs:").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.num_inputs_var = tk.StringVar(value="20")
        num_inputs_spinbox = ttk.Spinbox(
            device_frame,
            from_=8,
            to=40,
            textvariable=self.num_inputs_var,
            increment=1,
            width=10
        )
        num_inputs_spinbox.grid(row=1, column=1, padx=5, pady=2, sticky=tk.W)

        # Max Current per Input
        ttk.Label(device_frame, text="Max Current per Input (A):").grid(row=2, column=0, padx=5, pady=2, sticky=tk.W)
        self.max_current_per_input_var = tk.StringVar(value="20")
        max_current_spinbox = ttk.Spinbox(
            device_frame,
            from_=15,
            to=60,
            textvariable=self.max_current_per_input_var,
            increment=1,
            width=10
        )
        max_current_spinbox.grid(row=2, column=1, padx=5, pady=2, sticky=tk.W)

        # Device Max Current (calculated)
        ttk.Label(device_frame, text="Device Max Current (A):").grid(row=3, column=0, padx=5, pady=2, sticky=tk.W)
        self.device_max_current_label = ttk.Label(device_frame, text="--")
        self.device_max_current_label.grid(row=3, column=1, padx=5, pady=2, sticky=tk.W)

        # Device Spacing
        ttk.Label(device_frame, text="Device Spacing (ft):").grid(row=5, column=0, padx=5, pady=2, sticky=tk.W)
        self.device_spacing_var = tk.StringVar(value="6.0")
        device_spacing_entry = ttk.Entry(device_frame, textvariable=self.device_spacing_var)
        device_spacing_entry.grid(row=5, column=1, padx=5, pady=2, sticky=tk.W)
        self.device_spacing_meters_label = ttk.Label(device_frame, text="(1.83m)")
        self.device_spacing_meters_label.grid(row=5, column=2, padx=5, pady=2, sticky=tk.W)

        # Selected Inverter
        ttk.Label(device_frame, text="Selected Inverter:").grid(row=6, column=0, padx=5, pady=2, sticky=tk.W)
        self.inverter_label = ttk.Label(device_frame, text="None")
        self.inverter_label.grid(row=6, column=1, padx=5, pady=2, sticky=tk.W)
        # Change this line in the UI setup:
        self.inverter_select_button = ttk.Button(device_frame, text="Select Inverter", command=self.select_inverter)
        self.inverter_select_button.grid(row=6, column=2, padx=5, pady=2)
        # Add trace to update both display and block
        self.device_spacing_var.trace('w', lambda *args: (
            self.update_device_spacing_display(),
            self.draw_block()
        ))

        # Add placement mode radio buttons
        placement_frame = ttk.LabelFrame(device_frame, text="Device Placement Mode")
        placement_frame.grid(row=7, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))

        ttk.Radiobutton(placement_frame, text="Center Between Rows", 
                    variable=self.device_placement_mode, value="row_center",
                    command=self.update_device_placement).grid(
                    row=0, column=0, padx=5, pady=2, sticky=tk.W)

        ttk.Radiobutton(placement_frame, text="Align With Tracker", 
                    variable=self.device_placement_mode, value="tracker_align",
                    command=self.update_device_placement).grid(
                    row=1, column=0, padx=5, pady=2, sticky=tk.W)

        # Add spacing after device frame
        ttk.Label(config_frame, text="").grid(row=6, column=0, pady=5)  # Spacer

        # Configure Wiring button
        ttk.Button(config_frame, text="Configure Wiring", 
                command=self.configure_wiring).grid(row=7, column=0, columnspan=2, padx=5, pady=5)

        # Templates List Frame
        templates_frame = ttk.LabelFrame(config_frame, text="Tracker Templates", padding="5")
        templates_frame.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))

        # Create Treeview for hierarchical template display
        self.template_tree = ttk.Treeview(templates_frame, height=8)
        self.template_tree.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure tree columns
        self.template_tree.heading('#0', text='Templates')
        self.template_tree.column('#0', width=250)

        # Add scrollbar for tree
        template_scrollbar = ttk.Scrollbar(templates_frame, orient=tk.VERTICAL, command=self.template_tree.yview)
        template_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.template_tree.configure(yscrollcommand=template_scrollbar.set)

        # Bind selection event
        self.template_tree.bind('<<TreeviewSelect>>', self.on_template_select)

        # Add current values display
        self.current_frame = ttk.Frame(templates_frame)
        self.current_frame.grid(row=1, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))

        ttk.Label(self.current_frame, text="String Current (Imp):").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.string_current_label = ttk.Label(self.current_frame, text="-- A")
        self.string_current_label.grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)

        ttk.Label(self.current_frame, text="NEC Current (Isc×1.25):").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.nec_current_label = ttk.Label(self.current_frame, text="-- A")
        self.nec_current_label.grid(row=1, column=1, padx=5, pady=2, sticky=tk.W)

        # Add string_current and nec_current properties to the class
        self.string_current = 0.0
        self.nec_current = 0.0

        # Canvas frame for block layout - on the right side
        canvas_frame = ttk.LabelFrame(main_container, text="Block Layout", padding="5")
        canvas_frame.grid(row=0, rowspan=2, column=2, padx=5, pady=5)

        # Fixed size canvas
        self.canvas = tk.Canvas(canvas_frame, width=1000, height=800, bg='white')
        self.canvas.grid(row=0, column=0, padx=5, pady=5)

        # Add undo/redo buttons
        button_frame = ttk.Frame(canvas_frame)
        button_frame.grid(row=1, column=0, pady=5)
        
        ttk.Button(button_frame, text="Undo", command=self.undo).grid(row=0, column=0, padx=2)
        ttk.Button(button_frame, text="Redo", command=self.redo).grid(row=0, column=1, padx=2)

        # Bind keyboard shortcuts
        self.canvas.bind('<Control-z>', self.undo)
        self.canvas.bind('<Control-y>', self.redo)

        # Add mouse wheel binding for zoom
        self.canvas.bind('<MouseWheel>', self.on_mouse_wheel)  # Windows
        self.canvas.bind('<Button-4>', self.on_mouse_wheel)    # Linux scroll up
        self.canvas.bind('<Button-5>', self.on_mouse_wheel)    # Linux scroll down
        
        # Canvas bindings for clicking and dragging trackers
        self.canvas.bind('<Button-1>', self.on_canvas_click)
        self.canvas.bind('<B1-Motion>', self.on_canvas_drag)
        self.canvas.bind('<ButtonRelease-1>', self.on_canvas_release)

        # Keyboard bindings for deleting trackers
        self.canvas.bind('<Delete>', self.delete_selected_tracker)
        self.canvas.bind('<BackSpace>', self.delete_selected_tracker)

        # Pan bindings
        self.canvas.bind('<Button-2>', self.start_pan)  # Middle mouse button
        self.canvas.bind('<B2-Motion>', self.update_pan)
        self.canvas.bind('<ButtonRelease-2>', self.end_pan)
        # Alternative right-click pan
        self.canvas.bind('<Button-3>', self.start_pan)  
        self.canvas.bind('<B3-Motion>', self.update_pan)
        self.canvas.bind('<ButtonRelease-3>', self.end_pan)
        
        # Make canvas and frames expandable
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        main_container.grid_columnconfigure(1, weight=1)
        main_container.grid_rowconfigure(1, weight=1)
        
        # Bind canvas events for drag and drop implementation
        self.canvas.bind('<Button-1>', self.on_canvas_click)
        self.canvas.bind('<B1-Motion>', self.on_canvas_drag)
        self.canvas.bind('<ButtonRelease-1>', self.on_canvas_release)

        # Add traces for max current calculation
        self.num_inputs_var.trace('w', lambda *args: self.update_device_max_current())
        self.max_current_per_input_var.trace('w', lambda *args: self.update_device_max_current())

        # Add traces for device configuration changes
        self.device_type_var.trace('w', self.on_device_config_change)
        self.num_inputs_var.trace('w', self.on_device_config_change)  
        self.max_current_per_input_var.trace('w', self.on_device_config_change)

        # Add trace for device type to show/hide inverter frame
        self.device_type_var.trace('w', lambda *args: self.toggle_inverter_frame())

        # Initialize device max current calculation
        self.update_device_max_current()

        # Initialize GCR calculation

    def toggle_inverter_frame(self, *args):
        """Show/hide inverter selection elements based on device type"""
        if self.device_type_var.get() == DeviceType.STRING_INVERTER.value:
            # Show inverter widgets
            self.inverter_label.grid(row=6, column=1, padx=5, pady=2, sticky=tk.W)
            for widget in [
                self.inverter_label,
                self.inverter_select_button
            ]:
                widget.grid()
        else:
            # Hide inverter widgets
            for widget in [
                self.inverter_label,
                self.inverter_select_button
            ]:
                widget.grid_remove()

    def set_project(self, project):
        """Set the current project and update UI accordingly"""
        self.current_project = project
        
        # If there's a selected block, update the UI to show its device configuration
        if self.current_block and self.current_block in self.blocks:
            block = self.blocks[self.current_block]
            
            # Update device configuration UI from block
            self.updating_ui = True
            self.device_type_var.set(block.device_type.value)
            self.num_inputs_var.set(str(block.num_inputs))
            self.max_current_per_input_var.set(str(block.max_current_per_input))
            self.update_device_max_current()  # Recalculate total current
            self.updating_ui = False

    def on_inverter_selected(self, inverter):
        """Handle inverter selection"""
        self.selected_inverter = inverter
        if self.current_block:
            self.blocks[self.current_block].inverter = inverter

    def select_inverter(self):
        """Open inverter selection dialog"""
        dialog = tk.Toplevel(self)
        dialog.title("Select Inverter")
        dialog.transient(self)
        dialog.grab_set()
        
        inverter_selector = InverterManager(
            dialog,
            on_inverter_selected=lambda inv: self.handle_inverter_selection(inv, dialog)
        )
        inverter_selector.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Position dialog relative to parent
        x = self.winfo_rootx() + 50
        y = self.winfo_rooty() + 50
        dialog.geometry(f"+{x}+{y}")

    def handle_inverter_selection(self, inverter, dialog):
        """Handle inverter selection from dialog"""
        self.selected_inverter = inverter
        self.inverter_label.config(text=f"{inverter.manufacturer} {inverter.model}")
        # Add this line to store the inverter
        self.inverters[f"{inverter.manufacturer} {inverter.model}"] = inverter
        if self.current_block:
            self.blocks[self.current_block].inverter = inverter
        dialog.destroy()
        
    def create_new_block(self):
        """Create a new block configuration"""
        # Determine the block ID based on existing blocks
        if not self.blocks:
            # No existing blocks, start with default
            block_id = "Block_1"
        else:
            # Follow the pattern of the most recently created block
            if self.most_recent_block and self.most_recent_block in self.blocks:
                pattern_source = self.most_recent_block
            else:
                # No recent block, use any existing block as fallback
                pattern_source = list(self.blocks.keys())[0]
            
            import re
            match = re.match(r'(.*?)(\d+)$', pattern_source)
            
            if match:
                base_name = match.group(1)
                
                # Find the highest number with this prefix to avoid conflicts
                highest_for_prefix = 0
                for existing_id in self.blocks.keys():
                    existing_match = re.match(r'(.*?)(\d+)$', existing_id)
                    if existing_match and existing_match.group(1) == base_name:
                        existing_number = int(existing_match.group(2))
                        highest_for_prefix = max(highest_for_prefix, existing_number)
                
                next_number = highest_for_prefix + 1
                
                # Format with leading zeros if the source had them
                source_number_str = match.group(2)
                if len(source_number_str) > 1 and source_number_str.startswith('0'):
                    block_id = f"{base_name}{next_number:0{len(source_number_str)}d}"
                else:
                    block_id = f"{base_name}{next_number}"
            else:
                # Fallback if recent block has no number
                block_id = f"Block_{len(self.blocks) + 1:02d}"
        
        try:
            # Get selected template if any
            selection = self.template_tree.selection()
            selected_template = None
            if selection:
                item = selection[0]
                values = self.template_tree.item(item, 'values')
                if not values:
                    messagebox.showwarning("Warning", "Please select a template, not a manufacturer folder")
                    return
                template_key = values[0]
                selected_template = self.tracker_templates.get(template_key)
            
            # Get row spacing from project if available
            row_spacing_m = 6.0  # Default fallback
            if self.current_project and hasattr(self.current_project, 'default_row_spacing_m'):
                row_spacing_m = self.current_project.default_row_spacing_m
                    
            # Calculate device position based on row spacing    
            device_width_m = 0.91  # 3ft in meters
            # Calculate device position based on placement mode
            if self.device_placement_mode.get() == "row_center":
                # Center between rows
                initial_device_x = row_spacing_m / 2 + (device_width_m / 2)
            else:
                # Will be updated when a tracker is placed
                initial_device_x = row_spacing_m / 2 + (device_width_m / 2)  # Default until tracker placed
            
            block = BlockConfig(
                block_id=block_id,
                inverter=self.selected_inverter,
                tracker_template=selected_template,
                width_m=20,  # Initial minimum width
                height_m=20,  # Initial minimum height
                row_spacing_m=row_spacing_m,
                ns_spacing_m=float(self.ns_spacing_var.get()),
                gcr=0.0,
                description=f"New block {block_id}",
                device_spacing_m=self.ft_to_m(float(self.device_spacing_var.get())),
                device_x=initial_device_x,
                device_y=0.0,
                device_type=DeviceType(self.device_type_var.get()),
                num_inputs=int(self.num_inputs_var.get()),
                max_current_per_input=float(self.max_current_per_input_var.get())
            )
            
            # Add to blocks dictionary
            self.blocks[block_id] = block
            
            # Update listbox with sorted blocks
            self.update_block_listbox()
            
            # Update the row spacing variable in the UI without triggering callbacks
            self.updating_ui = True
            feet_value = self.m_to_ft(row_spacing_m)
            self.row_spacing_var.set(str(round(feet_value, 1)))
            self.updating_ui = False
            
            # Select new block
            self.block_listbox.selection_clear(0, tk.END)
            self.block_listbox.selection_set(tk.END)
            self.on_block_select()
            
            # Save initial empty state
            self._push_state("Create block")

            # Track this as the most recent block
            self.most_recent_block = block_id

            # Notify blocks changed
            self._notify_blocks_changed()
            
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid input: {str(e)}")
        
    def delete_block(self):
        """Delete currently selected block"""
        if not self.current_block:
            return
            
        # Check if block exists
        if self.current_block not in self.blocks:
            messagebox.showwarning("Warning", f"Block '{self.current_block}' not found. It may have been deleted already.")
            
            # Update listbox to remove this item
            for i in range(self.block_listbox.size()):
                if self.block_listbox.get(i) == self.current_block:
                    self.block_listbox.delete(i)
                    break
                    
            self.current_block = None
            return
            
        if messagebox.askyesno("Confirm", f"Delete block {self.current_block}?"):
            # Remove from dictionary
            del self.blocks[self.current_block]
            
            # Update listbox with sorted blocks
            self.update_block_listbox()
                
            # Clear current block
            self.current_block = None
            self.clear_config_display()

            # Notify blocks changed
            self._notify_blocks_changed()
            
    def on_block_select(self, event=None):
        """Handle block selection from listbox"""
        selection = self.block_listbox.curselection()
        if not selection:
            return
                
        block_id = self.block_listbox.get(selection[0])
        self.current_block = block_id

        # Add this check to preserve realistic routes if they exist
        if block_id in self.blocks and hasattr(self.blocks[block_id], 'wiring_config') and self.blocks[block_id].wiring_config:
            # If the block has existing realistic routes, keep a copy
            cable_routes = getattr(self.blocks[block_id], 'block_realistic_routes', None)
            wiring_routes = self.blocks[block_id].wiring_config.cable_routes
        
        # Add error handling to check if block exists
        if block_id not in self.blocks:
            print(f"Warning: Block '{block_id}' not found in blocks dictionary")
            messagebox.showwarning("Warning", f"Block '{block_id}' not found. It may have been deleted or corrupted.")
            
            # Remove this item from the listbox
            self.block_listbox.delete(selection[0])
            self.current_block = None
            return
            
        block = self.blocks[block_id]

        print(f"[BlockConfigurator.on_block_select] Selected block {block_id}:")
        print(f"  - device_type: {block.device_type.value}")
        print(f"  - num_inputs: {block.num_inputs}")
        print(f"  - max_current_per_input: {block.max_current_per_input}")
        
        # Update device configuration UI from block
        self.updating_ui = True
        self.device_type_var.set(block.device_type.value)
        self.num_inputs_var.set(str(block.num_inputs))
        self.max_current_per_input_var.set(str(block.max_current_per_input))
        self.update_device_max_current()  # Recalculate total current

        # Update UI with block data (convert meters to feet)
        self.updating_ui = True
        self.block_id_var.set(block.block_id)
        self.row_spacing_var.set(str(self.m_to_ft(block.row_spacing_m)))
        self.updating_ui = False
        self.calculate_gcr()  # Update the GCR label
        
        # After all other initialization, just before draw_block(), restore the routes if needed
        if block_id in self.blocks and hasattr(self.blocks[block_id], 'wiring_config') and self.blocks[block_id].wiring_config:
            if cable_routes and not hasattr(self.blocks[block_id], 'block_realistic_routes'):
                self.blocks[block_id].block_realistic_routes = cable_routes
            if wiring_routes and not self.blocks[block_id].wiring_config.cable_routes:
                self.blocks[block_id].wiring_config.cable_routes = wiring_routes

        # Update canvas
        self.draw_block()
        
        # Auto-select a template if available
        if self.template_tree.get_children():
            # If the block has a template, try to select it in the tree
            if block.tracker_template and block.tracker_template.template_name:
                template_name = block.tracker_template.template_name
                # Search through the tree to find and select the template
                for manufacturer_item in self.template_tree.get_children():
                    for template_item in self.template_tree.get_children(manufacturer_item):
                        values = self.template_tree.item(template_item, 'values')
                        if values and values[0] == template_name:
                            # Expand the manufacturer and select the template
                            self.template_tree.item(manufacturer_item, open=True)
                            self.template_tree.selection_set(template_item)
                            self.template_tree.see(template_item)
                            # Set the drag template
                            self.drag_template = self.tracker_templates.get(template_name)
                            # Manually call the template selection handler
                            self.on_template_select()
                            break
                
            # If no template was matched but there are templates, select the first available template
            if not self.drag_template and self.template_tree.get_children():
                # Find the first actual template (not just a manufacturer)
                for manufacturer_item in self.template_tree.get_children():
                    templates = self.template_tree.get_children(manufacturer_item)
                    if templates:
                        self.template_tree.item(manufacturer_item, open=True)
                        self.template_tree.selection_set(templates[0])
                        self.template_tree.see(templates[0])
                        break
                # Manually call the template selection handler
                self.on_template_select()
        
    def on_device_config_change(self, *args):
        """Update block configuration when device settings change"""
        if not self.current_block or self.updating_ui:
            return
            
        try:
            block = self.blocks[self.current_block]
            
            print(f"[BlockConfigurator.on_device_config_change] Updating block {self.current_block}")
            print(f"  - Old values: type={block.device_type.value}, inputs={block.num_inputs}, current={block.max_current_per_input}")
            
            # Update device type
            new_device_type = DeviceType(self.device_type_var.get())
            
            # Check if device type change is allowed
            if not self.check_device_type_consistency(new_device_type):
                # Revert to current block's device type
                self.device_type_var.set(block.device_type.value)
                return
            
            # Update block with new values
            block.device_type = new_device_type
            block.num_inputs = int(self.num_inputs_var.get())
            block.max_current_per_input = float(self.max_current_per_input_var.get())
            
            print(f"  - New values: type={block.device_type.value}, inputs={block.num_inputs}, current={block.max_current_per_input}")
            
            # Save state for undo
            self._push_state("Update device configuration")
            
        except ValueError:
            # Revert to current values if invalid input
            self.device_type_var.set(block.device_type.value)
            self.num_inputs_var.set(str(block.num_inputs))
            self.max_current_per_input_var.set(str(block.max_current_per_input))
    
    def check_device_type_consistency(self, new_device_type):
        """Check if all blocks have the same device type"""
        existing_types = set()
        for block_id, block in self.blocks.items():
            if block_id != self.current_block:  # Skip current block
                existing_types.add(block.device_type)
        
        if existing_types and new_device_type not in existing_types:
            # Different device type found
            existing_type = next(iter(existing_types))
            result = messagebox.askyesno(
                "Device Type Consistency",
                f"Other blocks use {existing_type.value}.\n"
                f"Do you want to change ALL blocks to {new_device_type.value}?"
            )
            
            if result:
                # Change all blocks to new device type
                for block in self.blocks.values():
                    block.device_type = new_device_type
                return True
            else:
                return False
        
        return True

    def clear_config_display(self):
        """Clear block configuration display"""
        self.block_id_var.set("")
        # self.width_var.set("200")
        # self.height_var.set("165")
        self.row_spacing_var.set("18.5")
        # self.gcr_var.set("0.4")
        self.canvas.delete("all")
        
    def draw_block(self):
        """Draw the current block on the canvas"""
        if not self.current_block:
            return
                
        try:
            # Validate current block exists
            if self.current_block and self.current_block not in self.blocks:
                print(f"Warning: Current block '{self.current_block}' not found")
                self.current_block = None
                self.clear_config_display()
                return
            if not self.validate_device_clearances():
                return

            block = self.blocks[self.current_block]

            # Clear canvas and grid lines list
            self.canvas.delete("all")
            self.grid_lines = []

            # Validate current block exists
            if self.current_block and self.current_block not in self.blocks:
                print(f"Warning: Current block '{self.current_block}' not found")
                self.current_block = None
                self.clear_config_display()
                return            
            if not self.validate_device_clearances():
                return

            block = self.blocks[self.current_block]

            # Clear canvas and grid lines list
            self.canvas.delete("all")
            self.grid_lines = []
            
            # Calculate block dimensions first
            block_width_m, block_height_m = self.calculate_block_dimensions()
            
            scale = self.get_canvas_scale()
            
            # Draw grid lines across entire canvas
            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()

            # Calculate grid extent - ensure we cover entire visible area plus padding
            grid_padding = 1000  # Add extra padding for panning
            start_x = -grid_padding
            start_y = -grid_padding
            end_x = canvas_width + grid_padding
            end_y = canvas_height + grid_padding

            # Draw vertical grid lines
            x = start_x
            while x <= end_x:
                scaled_x = x + self.pan_x
                line_id = self.canvas.create_line(
                    scaled_x, start_y,
                    scaled_x, end_y,
                    fill='gray', dash=(2, 4)
                )
                self.grid_lines.append(line_id)
                x += block.row_spacing_m * scale

            # Draw horizontal grid lines
            y = start_y
            while y <= end_y:
                scaled_y = y + self.pan_y
                line_id = self.canvas.create_line(
                    start_x, scaled_y,
                    end_x, scaled_y,
                    fill='gray', dash=(2, 4)
                )
                self.grid_lines.append(line_id)
                y += float(self.ns_spacing_var.get()) * scale

            # Draw device zones
            self.draw_device_zones(block.height_m)

            # Draw device
            if self.current_block:
                block = self.blocks[self.current_block]
                device_x, device_y, device_w, device_h = self.get_device_coordinates(block.height_m)
                scale = self.get_canvas_scale()
                
                # Convert to canvas coordinates
                x1 = 10 + self.pan_x + device_x * scale
                y1 = 10 + self.pan_y + device_y * scale
                x2 = x1 + device_w * scale
                y2 = y1 + device_h * scale
                
                # Draw device
                self.canvas.create_rectangle(x1, y1, x2, y2, fill='red', tags='device')    
            
            # Draw existing trackers with pan offset
            for pos in block.tracker_positions:
                x = 10 + self.pan_x + pos.x * scale
                y = 10 + self.pan_y + pos.y * scale
                self.draw_tracker(x, y, pos.template)

            # Draw wiring routes from wiring configuration
            self.draw_wiring_from_config()
        
        except (ValueError, TypeError):
            # Silently ignore transient errors during editing
            pass

    def draw_wiring_from_config(self):
        """Draw wiring routes from the wiring configuration"""
        if not self.current_block:
            return
            
        block = self.blocks[self.current_block]
        
        # Check if block has wiring configuration
        if not block.wiring_config:
            return
            
        scale = self.get_canvas_scale()
        
        # Draw all routes from the wiring configuration
        # Use realistic routes if available, otherwise fall back to conceptual routes
        routes_to_use = getattr(block.wiring_config, 'realistic_cable_routes', {})
        if not routes_to_use:
            routes_to_use = block.wiring_config.cable_routes

        for route_id, route_points in routes_to_use.items():
            if len(route_points) < 2:
                continue
                
            # Convert route points to canvas coordinates
            canvas_points = []
            for x, y in route_points:
                canvas_x = 10 + self.pan_x + x * scale
                canvas_y = 10 + self.pan_y + y * scale
                canvas_points.extend([canvas_x, canvas_y])
            
            # Determine wire properties based on route type
            is_positive = 'pos_' in route_id
            if 'src_' in route_id or 'node_' in route_id:
                wire_gauge = block.wiring_config.string_cable_size
                line_thickness = self.get_line_thickness_for_wire_gauge(wire_gauge)
            elif 'harness_' in route_id:
                wire_gauge = block.wiring_config.harness_cable_size
                line_thickness = self.get_line_thickness_for_wire_gauge(wire_gauge)
            elif 'extender_' in route_id:
                wire_gauge = getattr(block.wiring_config, 'extender_cable_size', '8 AWG')
                line_thickness = self.get_line_thickness_for_wire_gauge(wire_gauge)
            else:  # whip routes (dev_, main_)
                wire_gauge = getattr(block.wiring_config, 'whip_cable_size', '8 AWG')
                line_thickness = self.get_line_thickness_for_wire_gauge(wire_gauge)
            
            # Draw the route
            # Use the same color scheme as wiring configurator
            if 'src_' in route_id or 'node_' in route_id:
                color = '#FFB6C1' if is_positive else '#B0FFFF'  # Light Pink/Cyan (string)
            elif 'harness_' in route_id:
                color = '#FF0000' if is_positive else '#0000FF'  # Red/Blue (harness)
            elif 'extender_' in route_id:
                color = '#8B0000' if is_positive else '#800080'  # Dark Red/Purple (extender)
            else:  # whip routes
                color = '#FFA500' if is_positive else '#40E0D0'  # Orange/Turquoise (whip)
            self.canvas.create_line(canvas_points, fill=color, width=line_thickness, tags='wiring')

            # Draw visual connection points
            self.draw_wiring_points()

    def draw_wiring_points(self):
        """Draw collection points, whip points, and extender points like wiring configurator"""
        if not self.current_block:
            return
            
        block = self.blocks[self.current_block]
        if not block.wiring_config:
            return
            
        scale = self.get_canvas_scale()
        
        # Draw collection points (red/blue circles)
        for point in block.wiring_config.positive_collection_points:
            x = 10 + self.pan_x + point.x * scale
            y = 10 + self.pan_y + point.y * scale
            self.canvas.create_oval(x-3, y-3, x+3, y+3, fill='red', outline='darkred', tags='wiring')
            
        for point in block.wiring_config.negative_collection_points:
            x = 10 + self.pan_x + point.x * scale
            y = 10 + self.pan_y + point.y * scale
            self.canvas.create_oval(x-3, y-3, x+3, y+3, fill='blue', outline='darkblue', tags='wiring')

    def get_line_thickness_for_wire_gauge(self, wire_gauge: str) -> float:
        """Convert wire gauge to line thickness for display"""
        thickness_map = {
            "4 AWG": 5.0,
            "6 AWG": 4.0,
            "8 AWG": 3.0, 
            "10 AWG": 2.0
        }
        return thickness_map.get(wire_gauge, 2.0)

    def draw_tracker(self, x, y, template, tag=''):
        """Draw a single tracker at given position
        x and y are in canvas coordinates (already scaled and with padding)"""
        if not template:
            return
                    
        dims = template.get_physical_dimensions()
        scale = self.get_canvas_scale()
        
        # Get module dimensions
        if template.module_orientation == ModuleOrientation.PORTRAIT:
            module_height = template.module_spec.width_mm / 1000
            module_width = template.module_spec.length_mm / 1000
        else:
            module_height = template.module_spec.length_mm / 1000
            module_width = template.module_spec.width_mm / 1000

        # Create group tag for all elements
        group_tag = tag if tag else f'tracker_{x}_{y}'
        
        # Draw torque tube through center
        self.canvas.create_line(
            x + module_width * scale/2, y,
            x + module_width * scale/2, y + dims[1] * scale,
            width=3, fill='gray', tags=group_tag
        )

        # Handle different motor placement types
        if template.motor_placement_type == "middle_of_string":
            # Motor is in the middle of a specific string
            current_y = y
            
            for string_idx in range(template.strings_per_tracker):
                if string_idx + 1 == template.motor_string_index:  # This string has the motor (1-based index)
                    # Draw north modules
                    for i in range(template.motor_split_north):
                        self.canvas.create_rectangle(
                            x, current_y,
                            x + module_width * scale, current_y + module_height * scale,
                            fill='lightblue', outline='blue', tags=group_tag
                        )
                        current_y += (module_height + template.module_spacing_m) * scale
                    
                    # Draw motor gap with red circle
                    motor_y = current_y
                    gap_height = template.motor_gap_m * scale
                    circle_radius = min(gap_height / 3, module_width * scale / 4)
                    circle_center_x = x + module_width * scale / 2
                    circle_center_y = motor_y + gap_height / 2
                    
                    self.canvas.create_oval(
                        circle_center_x - circle_radius, circle_center_y - circle_radius,
                        circle_center_x + circle_radius, circle_center_y + circle_radius,
                        fill='red', outline='darkred', width=2, tags=group_tag
                    )
                    current_y += gap_height
                    
                    # Draw south modules
                    for i in range(template.motor_split_south):
                        self.canvas.create_rectangle(
                            x, current_y,
                            x + module_width * scale, current_y + module_height * scale,
                            fill='lightblue', outline='blue', tags=group_tag
                        )
                        current_y += (module_height + template.module_spacing_m) * scale
                else:
                    # Draw normal string without motor
                    for i in range(template.modules_per_string):
                        self.canvas.create_rectangle(
                            x, current_y,
                            x + module_width * scale, current_y + module_height * scale,
                            fill='lightblue', outline='blue', tags=group_tag
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
            y_pos = y
            
            # Draw modules above motor
            for i in range(modules_above_motor):
                self.canvas.create_rectangle(
                    x, y_pos,
                    x + module_width * scale, y_pos + module_height * scale,
                    fill='lightblue', outline='blue', tags=group_tag
                )
                y_pos += (module_height + template.module_spacing_m) * scale

            # Draw motor (only if there are strings below)
            if strings_below_motor > 0:
                motor_y = y_pos
                gap_height = template.motor_gap_m * scale
                circle_radius = min(gap_height / 3, module_width * scale / 4)
                circle_center_x = x + module_width * scale / 2
                circle_center_y = motor_y + gap_height / 2
                
                self.canvas.create_oval(
                    circle_center_x - circle_radius, circle_center_y - circle_radius,
                    circle_center_x + circle_radius, circle_center_y + circle_radius,
                    fill='red', outline='darkred', width=2, tags=group_tag
                )
                y_pos += gap_height

            # Draw modules below motor
            for i in range(modules_below_motor):
                self.canvas.create_rectangle(
                    x, y_pos,
                    x + module_width * scale, y_pos + module_height * scale,
                    fill='lightblue', outline='blue', tags=group_tag
                )
                y_pos += (module_height + template.module_spacing_m) * scale

    def get_canvas_scale(self):
        """Return current scale factor (pixels per meter)"""
        return self.scale_factor

    def on_canvas_click(self, event):
        """Handle canvas click for tracker placement"""
        if not self.current_block or not self.drag_template:
            return
            
        # Give canvas keyboard focus when clicked
        self.canvas.focus_set()

        # First check if we're clicking on an existing tracker
        if self.select_tracker(event.x, event.y):
            return
            
        self.dragging = True
        self.drag_start = (event.x, event.y)
        
        # Calculate snapped position - account for pan offset
        block = self.blocks[self.current_block]
        scale = self.get_canvas_scale()
        x_m = (event.x - 10 - self.pan_x) / scale  # Subtract pan offset
        y_m = (event.y - 10 - self.pan_y) / scale  # Subtract pan offset
        
        # Snap to grid
        x_m = round(x_m / block.row_spacing_m) * block.row_spacing_m
        y_m = round(y_m / float(self.ns_spacing_var.get())) * float(self.ns_spacing_var.get())
        
        # Draw preview - add pan offset back for canvas coordinates
        x = x_m * scale + 10 + self.pan_x  # Add pan offset
        y = y_m * scale + 10 + self.pan_y  # Add pan offset
        self.draw_tracker(x, y, self.drag_template, 'drag_preview')

    def on_canvas_drag(self, event):
        """Handle canvas drag for tracker movement"""
        if not self.dragging:
            return
        
        # Delete old preview
        self.canvas.delete('drag_preview')
        
        # Calculate snapped position - account for pan offset
        block = self.blocks[self.current_block]
        scale = self.get_canvas_scale()
        x_m = (event.x - 10 - self.pan_x) / scale  # Subtract pan offset
        y_m = (event.y - 10 - self.pan_y) / scale  # Subtract pan offset
        
        # Snap to grid
        x_m = round(x_m / block.row_spacing_m) * block.row_spacing_m
        y_m = round(y_m / float(self.ns_spacing_var.get())) * float(self.ns_spacing_var.get())
        
        # Draw new preview - add pan offset back for canvas coordinates
        x = x_m * scale + 10 + self.pan_x  # Add pan offset
        y = y_m * scale + 10 + self.pan_y  # Add pan offset
        self.draw_tracker(x, y, self.drag_template, 'drag_preview')

    def on_canvas_release(self, event):
        """Handle canvas release for tracker placement"""
        if not self.dragging or not self.current_block or not self.drag_template:
            return
                                
        self.dragging = False
        self.canvas.delete('drag_preview')
        
        # Get block reference
        block = self.blocks[self.current_block]
        
        # Convert canvas coordinates to meters - account for pan offset
        scale = self.get_canvas_scale()
        x_m = (event.x - 10 - self.pan_x) / scale  # Subtract pan offset
        y_m = (event.y - 10 - self.pan_y) / scale  # Subtract pan offset
        
        # Snap to grid
        x_m = round(x_m / block.row_spacing_m) * block.row_spacing_m
        y_m = round(y_m / float(self.ns_spacing_var.get())) * float(self.ns_spacing_var.get())
        
        # Check for collisions with existing trackers
        new_width, new_height = self.calculate_tracker_dimensions(self.drag_template)
        for pos in block.tracker_positions:
            existing_width, existing_height = self.calculate_tracker_dimensions(pos.template)
            if (x_m < pos.x + existing_width and 
                x_m + new_width > pos.x and 
                y_m < pos.y + existing_height and 
                y_m + new_height > pos.y):
                messagebox.showwarning("Invalid Position", "Cannot place tracker here - overlaps with existing tracker")
                return

        # Check if position is valid for device placement
        if not self.is_valid_tracker_position(x_m, y_m, new_height):
            messagebox.showwarning("Invalid Position", 
                "Cannot place tracker here due to device placement restrictions")
            return

        # Check module consistency before placement
        can_place, error_msg = self.check_module_consistency_before_placement(self.drag_template)
        if not can_place:
            messagebox.showwarning("Module Inconsistency", error_msg)
            return

        # Save state before adding tracker
        self._push_state("Before place tracker")
        
        # Create new TrackerPosition
        pos = TrackerPosition(x=x_m, y=y_m, rotation=0.0, template=self.drag_template)
        pos.calculate_string_positions() # Calculate string and collection point positions
        block.tracker_positions.append(pos)
        
        # Update device position if in tracker align mode
        if self.device_placement_mode.get() == "tracker_align" and len(block.tracker_positions) == 1:
            # This is the first tracker, align device with it
            if self.drag_template:
                module_width = self.drag_template.module_spec.length_mm / 1000
                device_width_m = 0.91
                block.device_x = x_m + (module_width / 2) - (device_width_m / 2)
        
        # Update block display
        self.draw_block()

        # Notify blocks changed
        self._notify_blocks_changed()


    def load_templates(self):
        """Load tracker templates from file, filtered by enabled templates"""
        template_path = Path('data/tracker_templates.json')
        if template_path.exists():
            try:
                with open(template_path, 'r') as f:
                    data = json.load(f)

                # Load all templates first
                all_templates = {}
                
                # Handle both flat and hierarchical formats
                if data:
                    first_value = next(iter(data.values()))
                    if isinstance(first_value, dict) and not any(key in first_value for key in ['module_orientation', 'modules_per_string']):
                        # New hierarchical format
                        for manufacturer, template_group in data.items():
                            for template_name, template_data in template_group.items():
                                unique_name = f"{manufacturer} - {template_name}"
                                all_templates[unique_name] = template_data
                    else:
                        # Old flat format
                        all_templates = data
                
                # Filter templates based on enabled status in current project
                self.tracker_templates = {}
                if self.current_project and hasattr(self.current_project, 'enabled_templates'):
                    # Only include enabled templates
                    for template_key, template_data in all_templates.items():
                        if template_key in self.current_project.enabled_templates:
                            try:
                                # Create TrackerTemplate object from data
                                module_data = template_data.get('module_spec', {})
                                module_spec = ModuleSpec(
                                    manufacturer=module_data.get('manufacturer', 'Unknown'),
                                    model=module_data.get('model', 'Unknown'),
                                    type=ModuleType(module_data.get('type', ModuleType.MONO_PERC.value)),
                                    length_mm=module_data.get('length_mm', 2000),
                                    width_mm=module_data.get('width_mm', 1000),
                                    depth_mm=module_data.get('depth_mm', 40),
                                    weight_kg=module_data.get('weight_kg', 25),
                                    wattage=module_data.get('wattage', 400),
                                    vmp=module_data.get('vmp', 40),
                                    imp=module_data.get('imp', 10),
                                    voc=module_data.get('voc', 48),
                                    isc=module_data.get('isc', 10.5),
                                    max_system_voltage=module_data.get('max_system_voltage', 1500)
                                )
                                
                                template = TrackerTemplate(
                                template_name=template_key,
                                module_spec=module_spec,
                                module_orientation=ModuleOrientation(template_data.get('module_orientation', ModuleOrientation.PORTRAIT.value)),
                                modules_per_string=template_data.get('modules_per_string', 28),
                                strings_per_tracker=template_data.get('strings_per_tracker', 2),
                                module_spacing_m=template_data.get('module_spacing_m', 0.01),
                                motor_gap_m=template_data.get('motor_gap_m', 1.0),
                                motor_position_after_string=template_data.get('motor_position_after_string', 0),
                                # New motor placement fields with backward compatibility
                                motor_placement_type=template_data.get('motor_placement_type', 'between_strings'),
                                motor_string_index=template_data.get('motor_string_index', 1),
                                motor_split_north=template_data.get('motor_split_north', 14),
                                motor_split_south=template_data.get('motor_split_south', 14)
                            )
                                
                                self.tracker_templates[template_key] = template
                            except Exception as e:
                                print(f"Error loading template {template_key}: {str(e)}")
                else:
                    # No project or no enabled_templates list - show all templates (fallback)
                    for template_key, template_data in all_templates.items():
                        try:
                            # Create TrackerTemplate object from data (same as above)
                            module_data = template_data.get('module_spec', {})
                            module_spec = ModuleSpec(
                                manufacturer=module_data.get('manufacturer', 'Unknown'),
                                model=module_data.get('model', 'Unknown'),
                                type=ModuleType(module_data.get('type', ModuleType.MONO_PERC.value)),
                                length_mm=module_data.get('length_mm', 2000),
                                width_mm=module_data.get('width_mm', 1000),
                                depth_mm=module_data.get('depth_mm', 40),
                                weight_kg=module_data.get('weight_kg', 25),
                                wattage=module_data.get('wattage', 400),
                                vmp=module_data.get('vmp', 40),
                                imp=module_data.get('imp', 10),
                                voc=module_data.get('voc', 48),
                                isc=module_data.get('isc', 10.5),
                                max_system_voltage=module_data.get('max_system_voltage', 1500)
                            )
                            
                            template = TrackerTemplate(
                                template_name=template_key,
                                module_spec=module_spec,
                                module_orientation=ModuleOrientation(template_data.get('module_orientation', ModuleOrientation.PORTRAIT.value)),
                                modules_per_string=template_data.get('modules_per_string', 28),
                                strings_per_tracker=template_data.get('strings_per_tracker', 2),
                                module_spacing_m=template_data.get('module_spacing_m', 0.01),
                                motor_gap_m=template_data.get('motor_gap_m', 1.0),
                                motor_position_after_string=template_data.get('motor_position_after_string', 0),
                                # New motor placement fields with backward compatibility
                                motor_placement_type=template_data.get('motor_placement_type', 'between_strings'),
                                motor_string_index=template_data.get('motor_string_index', 1),
                                motor_split_north=template_data.get('motor_split_north', 14),
                                motor_split_south=template_data.get('motor_split_south', 14)
                            )
                            
                            self.tracker_templates[template_key] = template
                        except Exception as e:
                            print(f"Error loading template {template_key}: {str(e)}")
                            
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load templates: {str(e)}")
                self.tracker_templates = {}
        else:
            self.tracker_templates = {}
                
    def on_template_select(self, event=None):
        """Handle template selection"""
        selection = self.template_tree.selection()
        if not selection:
            return
            
        item = selection[0]
        
        # Check if this is a template (has values) or manufacturer (no values)
        values = self.template_tree.item(item, 'values')
        if not values:
            # This is a manufacturer node, not a template
            return
            
        template_key = values[0]
        self.drag_template = self.tracker_templates.get(template_key)

        # Update the current labels
        self.update_current_labels()

        # Update current block's template if one is selected and there are no existing trackers
        if self.current_block and self.blocks[self.current_block] and len(self.blocks[self.current_block].tracker_positions) == 0:
            self.blocks[self.current_block].tracker_template = self.drag_template
            self._notify_blocks_changed()

    def update_template_list(self):
        """Update the template tree view with enabled templates only"""
        # Clear existing items
        for item in self.template_tree.get_children():
            self.template_tree.delete(item)
        
        # Check if we have any enabled templates
        if not self.tracker_templates:
            # Show message when no templates are enabled
            message = "No templates enabled for this project.\nGo to Tracker Creator to enable templates."
            self.template_tree.insert('', 'end', text=message, values=(), tags=('message',))
            # Configure message style
            self.template_tree.tag_configure('message', foreground='gray', font=('TkDefaultFont', 9, 'italic'))
            return
        
        # Group templates by manufacturer
        manufacturers = {}
        for template_key, template in self.tracker_templates.items():
            # Extract manufacturer from module_spec
            manufacturer = template.module_spec.manufacturer
            
            # Skip duplicate entries (we store both old and new names for compatibility)
            if ' - ' in template_key and not template_key.startswith(manufacturer):
                continue
                
            # Extract template name (remove manufacturer prefix if present)
            if ' - ' in template_key and template_key.startswith(manufacturer):
                template_name = template_key.split(' - ', 1)[1]
            else:
                template_name = template_key
                
            if manufacturer not in manufacturers:
                manufacturers[manufacturer] = []
            manufacturers[manufacturer].append((template_name, template_key, template))
        
        # Add manufacturers and their templates to tree
        for manufacturer, template_list in sorted(manufacturers.items()):
            # Add manufacturer node
            manufacturer_node = self.template_tree.insert('', 'end', text=manufacturer, open=True)
            
            # Add templates under manufacturer
            for template_name, template_key, template in sorted(template_list, key=lambda x: x[0]):
                # Show module info in template display
                model = template.module_spec.model
                wattage = template.module_spec.wattage
                template_text = f"{template_name} ({model} - {wattage}W)"
                self.template_tree.insert(manufacturer_node, 'end', text=template_text, values=(template_key,))

    def m_to_ft(self, meters):
        return meters * 3.28084

    def ft_to_m(self, feet):
        return feet / 3.28084
    
    def calculate_gcr(self):
        """Calculate and display GCR based on current row spacing and module length"""
        try:
            if not self.current_block or self.current_block not in self.blocks:
                self.updating_ui = True
                self.gcr_var.set("--")
                self.updating_ui = False
                return
                
            # Get module length from templates
            module_length_m, error = self.get_module_length_from_templates()
            if error:
                self.updating_ui = True
                self.gcr_var.set("--")
                self.updating_ui = False
                return
                
            # Get current row spacing
            row_spacing_m = self.blocks[self.current_block].row_spacing_m
            
            # Calculate GCR
            gcr = module_length_m / row_spacing_m
            
            # Update GCR display
            self.updating_ui = True
            self.gcr_var.set(f"{gcr:.3f}")
            self.updating_ui = False
            
            # Update block's GCR value
            self.blocks[self.current_block].gcr = gcr
            
        except (ValueError, ZeroDivisionError, AttributeError):
            self.updating_ui = True
            self.gcr_var.set("--")
            self.updating_ui = False

    def delete_selected_tracker(self, event=None):
        """Delete the currently selected tracker"""
        if not self.current_block or not self.selected_tracker:
            return
        
        block = self.blocks[self.current_block]
        
        # Save state before deletion
        self._push_state("Before delete tracker")
        
        # Find and remove the selected tracker
        positions_to_remove = []
        for i, pos in enumerate(block.tracker_positions):
            if (abs(pos.x - self.selected_tracker[0]) < 0.01 and 
                abs(pos.y - self.selected_tracker[1]) < 0.01):
                positions_to_remove.append(i)
        
        # Remove from highest index to lowest to avoid shifting issues
        for i in sorted(positions_to_remove, reverse=True):
            block.tracker_positions.pop(i)
        
        self.selected_tracker = None
        self.draw_block()

        # Notify blocks changed
        self._notify_blocks_changed()

    def select_tracker(self, x, y):
        """Select tracker at given coordinates"""
        if not self.current_block:
            return False

        block = self.blocks[self.current_block]
        scale = self.get_canvas_scale()
        
        # Convert canvas coordinates to meters, accounting for pan offset
        x_m = (x - 10 - self.pan_x) / scale
        y_m = (y - 10 - self.pan_y) / scale
        
        # Check if click is within any tracker
        for pos in block.tracker_positions:
            tracker_width, tracker_height = self.calculate_tracker_dimensions(pos.template)
            # Add some padding to make selection easier (0.2m padding)
            if (pos.x - 0.2 <= x_m <= pos.x + tracker_width + 0.2 and 
                pos.y - 0.2 <= y_m <= pos.y + tracker_height + 0.2):
                self.selected_tracker = (pos.x, pos.y)
                self.draw_block()
                    
                # Draw highlight rectangle with calculated dimensions
                x_canvas = 10 + pos.x * scale + self.pan_x
                y_canvas = 10 + pos.y * scale + self.pan_y
                # Create filled selection rectangle with transparency
                self.canvas.create_rectangle(
                    x_canvas - 2, y_canvas - 2,
                    x_canvas + tracker_width * scale + 2,
                    y_canvas + tracker_height * scale + 2,
                    outline='red', width=2,
                    fill='red', stipple='gray50',  # This creates a semi-transparent effect
                    tags='selection'
                )
                return True
        
        self.selected_tracker = None
        self.draw_block()
        return False
    
    def calculate_block_dimensions(self):
        """Calculate block dimensions based on placed trackers or template"""
        if not self.current_block:
            return (50, 50)  # Default minimum size
            
        block = self.blocks[self.current_block]
        
        # If no trackers but have a template, use template dimensions
        if not block.tracker_positions:
            if self.drag_template:
                dims = self.drag_template.get_physical_dimensions()
                # Add padding for initial placement
                width = max(50, dims[0] * 2)  # Allow for at least 2 trackers wide
                height = max(50, dims[1])  # Account for full tracker length
                return (width, height)
            return (50, 50)  # Default minimum size
                
        # Find max x and y coordinates including tracker dimensions
        max_x = 0
        max_y = 0
        for pos in block.tracker_positions:
            dims = pos.template.get_physical_dimensions()
            max_x = max(max_x, pos.x + dims[0])
            max_y = max(max_y, pos.y + dims[1])  # Explicitly add tracker length
            
        # Add 20% padding
        return (max_x * 1.2, max_y * 1.2)
    
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
            self.draw_block()

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
        self.draw_block()

    def end_pan(self, event):
        """End canvas panning"""
        self.panning = False
        self.canvas.config(cursor="")  # Reset cursor

    def calculate_tracker_dimensions(self, template):
        """Calculate actual tracker dimensions including all modules and gaps"""
        module_height = (template.module_spec.width_mm / 1000 if template.module_orientation == ModuleOrientation.PORTRAIT 
                        else template.module_spec.length_mm / 1000)
        module_width = (template.module_spec.length_mm / 1000 if template.module_orientation == ModuleOrientation.PORTRAIT 
                    else template.module_spec.width_mm / 1000)

        total_height = (
            # Height for modules above motor
            (template.modules_per_string * (template.strings_per_tracker - 1)) * 
            (module_height + template.module_spacing_m) +
            # Motor gap
            template.motor_gap_m +
            # Height for modules below motor
            template.modules_per_string * (module_height + template.module_spacing_m)
        )

        return module_width, total_height  # width, height
    
    def _restore_state(self, state):
        """Restore blocks from state"""
        if self.current_block:
            self._restore_from_state(state)
            self.draw_block()

    def _push_state(self, description: str):
        """Push current state to undo manager"""
        if self.current_block:
            self.undo_manager.push_state(description)

    def undo(self, event=None):
        """Handle undo command"""
        try:
            description = self.undo_manager.undo()
            if description:
                # Validate that current_block still exists
                if self.current_block and self.current_block not in self.blocks:
                    print(f"Warning: Current block '{self.current_block}' not found after undo")
                    self.current_block = None
                    self.clear_config_display()
                    return
                self.draw_block()
        except Exception as e:
            print(f"Error during undo: {str(e)}")
            messagebox.showerror("Undo Error", f"Failed to undo: {str(e)}")
            # Try to recover by clearing current selection
            self.current_block = None
            self.clear_config_display()

    def redo(self, event=None):
        """Handle redo command"""
        try:
            description = self.undo_manager.redo()
            if description:
                # Validate that current_block still exists
                if self.current_block and self.current_block not in self.blocks:
                    print(f"Warning: Current block '{self.current_block}' not found after redo")
                    self.current_block = None
                    self.clear_config_display()
                    return
                self.draw_block()
        except Exception as e:
            print(f"Error during redo: {str(e)}")
            messagebox.showerror("Redo Error", f"Failed to redo: {str(e)}")
            # Try to recover by clearing current selection
            self.current_block = None
            self.clear_config_display()

    def _get_current_state(self):
        """Get a deep copy of current block state"""
        if not self.current_block:
            return {}
        
        current_state = {}
        for id, block in self.blocks.items():
            positions = []
            for pos in block.tracker_positions:
                positions.append({
                    'x': pos.x,
                    'y': pos.y,
                    'rotation': pos.rotation,
                    'template_name': pos.template.template_name if pos.template else None,
                    'strings': [
                        {
                            'index': string.index,
                            'positive_source_x': string.positive_source_x,
                            'positive_source_y': string.positive_source_y,
                            'negative_source_x': string.negative_source_x,
                            'negative_source_y': string.negative_source_y,
                            'num_modules': string.num_modules
                        }
                        for string in pos.strings
                    ]
                })
                
            current_state[id] = {
                'block_id': block.block_id,
                'width_m': block.width_m,
                'height_m': block.height_m,
                'row_spacing_m': block.row_spacing_m,
                'ns_spacing_m': block.ns_spacing_m,
                'gcr': block.gcr,
                'description': block.description,
                'tracker_positions': positions,
                'inverter_name': f"{block.inverter.manufacturer} {block.inverter.model}" if block.inverter else None,
                'template_name': block.tracker_template.template_name if block.tracker_template else None,
                'device_spacing_m': block.device_spacing_m,
                'device_type': block.device_type.value,
                'num_inputs': block.num_inputs,
                'max_current_per_input': block.max_current_per_input
            }
            
            # Add wiring configuration if exists
            if block.wiring_config:
                # Serialize the wiring config
                wiring_config_data = {
                    'wiring_type': block.wiring_config.wiring_type.value,
                    'positive_collection_points': [
                        {
                            'x': point.x,
                            'y': point.y,
                            'connected_strings': point.connected_strings,
                            'current_rating': point.current_rating
                        }
                        for point in block.wiring_config.positive_collection_points
                    ],
                    'negative_collection_points': [
                        {
                            'x': point.x,
                            'y': point.y,
                            'connected_strings': point.connected_strings,
                            'current_rating': point.current_rating
                        }
                        for point in block.wiring_config.negative_collection_points
                    ],
                    'strings_per_collection': block.wiring_config.strings_per_collection,
                    'cable_routes': block.wiring_config.cable_routes,
                    'string_cable_size': getattr(block.wiring_config, 'string_cable_size', "10 AWG"),
                    'harness_cable_size': getattr(block.wiring_config, 'harness_cable_size', "8 AWG"),
                    'whip_cable_size': getattr(block.wiring_config, 'whip_cable_size', "8 AWG")
                }
                current_state[id]['wiring_config_data'] = wiring_config_data
                
        return current_state

    def _restore_from_state(self, state):
        """Restore block state from saved state"""
        if not state:
            self.blocks = {}
            return
            
        self.blocks = {}
        for id, block_data in state.items():
            # Get template and inverter from saved names
            template = next((t for t in self.tracker_templates.values() 
                            if t.template_name == block_data['template_name']), None)
            inverter = next((inv for name, inv in self.inverters.items() 
                            if f"{inv.manufacturer} {inv.model}" == block_data['inverter_name']), None)
            
            # Get row spacing value and convert feet to meters
            try:
                row_spacing_ft = float(self.row_spacing_var.get())
                row_spacing_m = self.ft_to_m(row_spacing_ft)
                initial_device_x = row_spacing_m / 2
            except ValueError:
                row_spacing_m = 6.0  # Default to 6m if invalid input
                initial_device_x = row_spacing_m / 2

            block = BlockConfig(
                block_id=block_data['block_id'],
                inverter=inverter,
                tracker_template=template,
                width_m=block_data['width_m'],
                height_m=block_data['height_m'],
                row_spacing_m=block_data['row_spacing_m'],
                ns_spacing_m=block_data['ns_spacing_m'],
                gcr=block_data['gcr'],
                description=block_data['description'],
                device_spacing_m=block_data.get('device_spacing_m', 1.83),  # 6ft default
                device_type=DeviceType(block_data.get('device_type', DeviceType.STRING_INVERTER.value)),
                num_inputs=block_data.get('num_inputs', 20),
                max_current_per_input=block_data.get('max_current_per_input', 20.0),
                device_x=initial_device_x,
                device_y=0.0
            )
            
            # Clear existing tracker positions
            block.tracker_positions = []
            
            # Import necessary classes
            from ..models.tracker import TrackerPosition, StringPosition
            
            # Restore tracker positions
            for pos_data in block_data['tracker_positions']:
                template = next((t for t in self.tracker_templates.values() 
                            if t.template_name == pos_data['template_name']), None)
                if template:
                    pos = TrackerPosition(
                        x=pos_data['x'],
                        y=pos_data['y'],
                        rotation=pos_data['rotation'],
                        template=template
                    )
                    
                    # Restore string data if available
                    if 'strings' in pos_data:
                        for string_data in pos_data['strings']:
                            string = StringPosition(
                                index=string_data['index'],
                                positive_source_x=string_data['positive_source_x'],
                                positive_source_y=string_data['positive_source_y'],
                                negative_source_x=string_data['negative_source_x'],
                                negative_source_y=string_data['negative_source_y'],
                                num_modules=string_data['num_modules']
                            )
                            pos.strings.append(string)
                    else:
                        # Fallback to calculation if string data not available
                        pos.calculate_string_positions()
                        
                    block.tracker_positions.append(pos)
            
            # Restore wiring configuration if available
            if 'wiring_config_data' in block_data:
                wiring_data = block_data['wiring_config_data']
                
                # Create collection points
                positive_points = [
                    CollectionPoint(
                        x=point['x'],
                        y=point['y'],
                        connected_strings=point['connected_strings'],
                        current_rating=point['current_rating']
                    )
                    for point in wiring_data['positive_collection_points']
                ]
                
                negative_points = [
                    CollectionPoint(
                        x=point['x'],
                        y=point['y'],
                        connected_strings=point['connected_strings'],
                        current_rating=point['current_rating']
                    )
                    for point in wiring_data['negative_collection_points']
                ]
                
                # Create wiring config with cable sizes
                block.wiring_config = WiringConfig(
                    wiring_type=WiringType(wiring_data['wiring_type']),
                    positive_collection_points=positive_points,
                    negative_collection_points=negative_points,
                    strings_per_collection=wiring_data['strings_per_collection'],
                    cable_routes=wiring_data['cable_routes'],
                    string_cable_size=wiring_data.get('string_cable_size', "10 AWG"),
                    harness_cable_size=wiring_data.get('harness_cable_size', "8 AWG"),
                    whip_cable_size=wiring_data.get('whip_cable_size', "8 AWG")
                )
            
            self.blocks[id] = block
            
        self.draw_block()

    def get_device_coordinates(self, block_height_m):
        """Calculate device coordinates based on placement"""
        device_height_m = 0.91  # 3ft in meters
        device_width_m = 0.91   # 3ft in meters
        
        if not self.current_block:
            return (0, 0, device_width_m, device_height_m)
            
        # Return actual device position from block config
        return (self.blocks[self.current_block].device_x, 
                self.blocks[self.current_block].device_y, 
                device_width_m, device_height_m)

    def is_valid_tracker_position(self, x, y, tracker_height):
        """Check if tracker position is valid based on device placement"""
        if not self.current_block:
            return False
            
        block = self.blocks[self.current_block]
        device_x, device_y, device_w, device_h = self.get_device_coordinates(block.height_m)
        
        # For now, all positions are valid as long as we have a block
        # You can add specific validation rules here later if needed
        return True
        
    def validate_device_clearances(self):
        """Validate device clearances against block dimensions"""
        if not self.current_block:
            return False
            
        block = self.blocks[self.current_block]
        spacing_m = self.ft_to_m(self.validate_float_input(self.device_spacing_var, 6.0))
        
        # Allow placement as long as block exists
        return True
    
    def draw_device_zones(self, block_height_m):
        """Draw device zones with semi-transparent shading"""
        if not self.current_block:
            return
            
        scale = self.get_canvas_scale()
        device_x, device_y, device_w, device_h = self.get_device_coordinates(block_height_m)
        spacing_m = self.ft_to_m(float(self.device_spacing_var.get()))
        
        # Convert device coordinates to canvas coordinates (include padding and pan)
        x1 = 10 + self.pan_x + device_x * scale
        x2 = x1 + device_w * scale
        
        # Add spacing zone width
        x1_zone = x1 - spacing_m * scale
        x2_zone = x2 + spacing_m * scale
        
        # Draw device safety zone (semi-transparent red)
        y1 = 10 + self.pan_y + (device_y - spacing_m) * scale
        y2 = 10 + self.pan_y + (device_y + device_h + spacing_m) * scale
        
        # Create safety zone with stipple pattern for transparency
        self.canvas.create_rectangle(
            x1_zone, y1, x2_zone, y2,
            fill='#ffcccc', stipple='gray50', tags='zones'
        )

        # Draw the device itself as a solid rectangle
        y1_device = 10 + self.pan_y + device_y * scale
        y2_device = y1_device + device_h * scale
        self.canvas.create_rectangle(
            x1, y1_device, x2, y2_device,
            fill='red', tags='device'
        )

    def update_device_max_current(self, *args):
        """Update the calculated device max current"""
        try:
            num_inputs = int(self.num_inputs_var.get())
            max_current_per_input = int(self.max_current_per_input_var.get())
            total_current = num_inputs * max_current_per_input
            self.device_max_current_label.config(text=f"{total_current}")
        except ValueError:
            self.device_max_current_label.config(text="--")

    def update_device_spacing_display(self):
        """Update the meters display when feet value changes"""
        try:
            feet = float(self.device_spacing_var.get())
            meters = self.ft_to_m(feet)
            self.device_spacing_meters_label.config(text=f"({meters:.2f}m)")
        except ValueError:
            self.device_spacing_meters_label.config(text="(invalid)")

    def configure_wiring(self):
        """Open wiring configuration window"""
            
        if not self.current_block:
            messagebox.showwarning("Warning", "Please select a block first")
            return
            
        from .wiring_configurator import WiringConfigurator
        wiring_config = WiringConfigurator(self, self.blocks[self.current_block])

        # Update block display when wiring configurator closes
        def on_wiring_closed():
            self.draw_block()  # Redraw to show new wiring routes
            
        # The wiring configurator will call this when it's destroyed/closed
        wiring_config.protocol("WM_DELETE_WINDOW", lambda: (on_wiring_closed(), wiring_config.destroy()))

    def copy_block(self):
        """Create a copy of the currently selected block with incremented name"""
        if not self.current_block:
            messagebox.showwarning("Warning", "Please select a block to copy")
            return
            
        # Get the source block
        source_block = self.blocks[self.current_block]
        source_id = source_block.block_id
        
        # Try to parse the source ID to find a numerical suffix
        import re
        match = re.match(r'(.*?)(\d+)$', source_id)
        
        if match:
            # If source ID ends with a number, increment it
            base_name = match.group(1)  # The part before the number
            number = int(match.group(2))  # The number part
            
            # Try incrementing the number until we find an unused ID
            new_id = f"{base_name}{number + 1:02d}"
            while new_id in self.blocks:
                number += 1
                new_id = f"{base_name}{number:02d}"
        else:
            # If no numeric suffix, use the old method
            base_id = f"{source_id}_copy"
            new_id = base_id
            counter = 1
            while new_id in self.blocks:
                new_id = f"{base_id}_{counter}"
                counter += 1
        
        # Create a deep copy of the block to avoid shared references
        from copy import deepcopy
        new_block = deepcopy(source_block)
        
        # Update the ID
        new_block.block_id = new_id
        
        # Add to blocks dictionary
        self.blocks[new_id] = new_block
        
        # Save state for undo
        self._push_state("Copy block")
        
        # Update listbox with sorted blocks
        self.update_block_listbox()
        
        # Select the new block
        self.block_listbox.selection_clear(0, tk.END)
        self.block_listbox.selection_set(tk.END)
        self.block_listbox.see(tk.END)  # Ensure it's visible
        self.on_block_select()

        # Track this as the most recent block
        self.most_recent_block = new_id

        # Notify blocks changed
        self._notify_blocks_changed()
        
        messagebox.showinfo("Success", f"Block copied as '{new_id}'")
    
    def _notify_blocks_changed(self):
        """Notify listeners that blocks have changed"""
        if self.on_blocks_changed:
            self.on_blocks_changed()
        
        # Trigger autosave
        if self.on_autosave:
            try:
                self.on_autosave()
            except Exception as e:
                print(f"Autosave failed: {str(e)}")

    def rename_block(self):
        """Rename the currently selected block"""
        if not self.current_block:
            messagebox.showwarning("Warning", "Please select a block to rename")
            return
            
        # Get the new ID from the entry field
        new_id = self.block_id_var.get().strip()
        
        # Validate the new ID
        if not new_id:
            messagebox.showerror("Error", "Block ID cannot be empty")
            return
            
        if new_id in self.blocks and new_id != self.current_block:
            messagebox.showerror("Error", f"Block ID '{new_id}' already exists")
            return
        
        # Save state for undo
        self._push_state("Before rename block")
        
        # Get the current block
        block = self.blocks[self.current_block]
        old_id = self.current_block
        
        # Remove the old entry from the dictionary
        del self.blocks[self.current_block]
        
        # Update the block ID
        block.block_id = new_id
        
        # Add back to dictionary with new ID
        self.blocks[new_id] = block
        
        # Update listbox with sorted blocks
        self.update_block_listbox()
        # Select the renamed block
        for i in range(self.block_listbox.size()):
            if self.block_listbox.get(i) == new_id:
                self.block_listbox.selection_set(i)
                break
        
        # Update current_block reference to the new ID
        self.current_block = new_id
        
        # Update UI display
        self.block_id_var.set(new_id)

        # Notify blocks changed
        self._notify_blocks_changed()
        
        messagebox.showinfo("Success", f"Block renamed to '{new_id}'")

    def update_block_row_spacing(self, event=None, *args):
        """Update the current block's row spacing from UI value"""
        # Return early if we're programmatically updating the UI
        if hasattr(self, 'updating_ui') and self.updating_ui:
            return
            
        if not self.current_block:
            return
                
        try:
            # Get row spacing value and convert feet to meters
            row_spacing_ft = float(self.row_spacing_var.get())
            row_spacing_m = self.ft_to_m(row_spacing_ft)
            
            # Check if value has actually changed
            if abs(self.blocks[self.current_block].row_spacing_m - row_spacing_m) < 0.001:
                return
                
            # Ask if change should apply to all blocks
            response = messagebox.askyesnocancel(
                "Update Row Spacing",
                "Do you want to apply this row spacing to all blocks?\n\n"
                "Yes - Apply to all blocks\n"
                "No - Apply only to the current block\n"
                "Cancel - Don't change row spacing"
            )
            
            if response is None:  # Cancel
                # Reset the row spacing input to match the current block
                self.updating_ui = True
                self.row_spacing_var.set(str(self.m_to_ft(self.blocks[self.current_block].row_spacing_m)))
                self.updating_ui = False
                return
                
            if response:  # Yes - apply to all blocks
                # Update all blocks
                for block_id, block in self.blocks.items():
                    block.row_spacing_m = row_spacing_m
                    
                # Update project default if possible
                if self.current_project:
                    # Update this instance's project reference
                    self.current_project.default_row_spacing_m = row_spacing_m
                    self.current_project.update_modified_date()
                    
                    # Also update the main app's project reference to keep them in sync
                    main_app = self.winfo_toplevel()
                    if hasattr(main_app, 'current_project') and main_app.current_project:
                        main_app.current_project.default_row_spacing_m = row_spacing_m
                        main_app.current_project.update_modified_date()
                        
                        # Auto-save the project
                        if hasattr(main_app, 'save_project'):
                            main_app.save_project()

            else:  # No - apply only to current block
                # Update current block only
                self.blocks[self.current_block].row_spacing_m = row_spacing_m
                
            # Update GCR after changing row spacing
            self.update_gcr_from_row_spacing()
            
            # Redraw the block
            self.draw_block()
            
        except (ValueError, KeyError):
            # Ignore conversion errors or invalid block references
            pass

    def reload_templates(self):
        """Reload tracker templates from disk and update UI"""
        self.load_templates()
        self.update_template_list()
        # If there's a current block, try to restore its template selection
        if self.current_block and self.blocks[self.current_block].tracker_template:
            template_name = self.blocks[self.current_block].tracker_template.template_name
            
            # Search through the tree to find and select the template
            for manufacturer_item in self.template_tree.get_children():
                for template_item in self.template_tree.get_children(manufacturer_item):
                    values = self.template_tree.item(template_item, 'values')
                    if values and values[0] == template_name:
                        # Expand the manufacturer and select the template
                        self.template_tree.item(manufacturer_item, open=True)
                        self.template_tree.selection_set(template_item)
                        self.template_tree.see(template_item)
                        # Set the drag template
                        self.drag_template = self.tracker_templates.get(template_name)
                        return

    def check_block_wiring_status(self):
        """Check if any blocks are missing wiring configuration and notify user"""
        blocks_without_wiring = []
        for block_id, block in self.blocks.items():
            if not block.wiring_config:
                blocks_without_wiring.append(block_id)
        
        if blocks_without_wiring:
            message = "The following blocks have no wiring configuration:\n"
            message += "\n".join(sorted(blocks_without_wiring))
            message += "\n\nBlocks without wiring configuration will have limited BOM output."
            messagebox.showwarning("Missing Wiring Configurations", message)
            return False
        return True
    
    def update_device_placement(self):
        """Update device position based on selected placement mode"""
        if not self.current_block:
            return
            
        block = self.blocks[self.current_block]
        mode = self.device_placement_mode.get()
        
        if mode == "row_center":
            # Place device at the center of row spacing
            block.device_x = block.row_spacing_m / 2
        elif mode == "tracker_align":
            # Find the first tracker or use default position
            if block.tracker_positions:
                # Align with the first tracker's center
                pos = block.tracker_positions[0]
                if pos.template:
                    # Get tracker dimensions to find center
                    module_width = pos.template.module_spec.length_mm / 1000
                    device_width_m = 0.91  # 3ft in meters
                    # Adjust device position so centers align
                    block.device_x = pos.x + (module_width / 2) - (device_width_m / 2)
                else:
                    # Fallback if no template
                    block.device_x = pos.x + 1.0  # 1m offset as fallback
            else:
                # No trackers, default to row center
                block.device_x = block.row_spacing_m / 2
        
        # Redraw with new device position
        self.draw_block()

    def update_block_listbox(self):
        """Update the block listbox with sorted block IDs"""
        self.block_listbox.delete(0, tk.END)
        
        # Sort blocks by numerical value
        def get_block_sort_key(block_id):
            import re
            # Try to extract prefix and number for proper sorting
            match = re.match(r'(.*?)(\d+)$', block_id)
            if match:
                prefix = match.group(1)
                number = int(match.group(2))
                return (prefix, number)
            else:
                # No number found, sort alphabetically
                return (block_id, 0)
        
        sorted_block_ids = sorted(self.blocks.keys(), key=get_block_sort_key)
        
        for block_id in sorted_block_ids:
            self.block_listbox.insert(tk.END, block_id)

    def get_module_length_from_templates(self):
        """Get module length from available templates and check for consistency"""
        if not self.tracker_templates:
            return None, "No tracker templates available"
        
        # Get all unique module lengths from available templates
        module_lengths = set()
        for template in self.tracker_templates.values():
            if template.module_spec and template.module_spec.length_mm:
                module_lengths.add(template.module_spec.length_mm)
        
        if not module_lengths:
            return None, "No module specifications found in templates"
        
        if len(module_lengths) > 1:
            # Multiple different module lengths - warn user
            lengths_ft = [f"{length/304.8:.2f}ft" for length in module_lengths]
            warning_msg = f"Warning: Templates have different module lengths: {', '.join(lengths_ft)}"
            messagebox.showwarning("Module Length Inconsistency", warning_msg)
            return None, "Inconsistent module lengths in templates"
        
        # All templates have the same module length
        module_length_mm = list(module_lengths)[0]
        return module_length_mm / 1000, None  # Convert to meters
    
    def update_gcr_from_row_spacing(self, event=None, *args):
        """Update GCR field when row spacing changes"""
        # Return early if we're programmatically updating the UI
        if hasattr(self, 'updating_ui') and self.updating_ui:
            return
        
        try:
            # Get row spacing value and convert feet to meters
            row_spacing_ft = float(self.row_spacing_var.get())
            if row_spacing_ft <= 0:
                messagebox.showerror("Error", "Row spacing must be a positive number")
                return
                
            row_spacing_m = self.ft_to_m(row_spacing_ft)
            
            # Get module length from templates
            module_length_m, error = self.get_module_length_from_templates()
            if error:
                # Clear GCR field if we can't calculate
                self.updating_ui = True
                self.gcr_var.set("--")
                self.updating_ui = False
                return
                
            # Calculate GCR
            gcr = module_length_m / row_spacing_m
            
            # Update GCR field
            self.updating_ui = True
            self.gcr_var.set(f"{gcr:.3f}")
            self.updating_ui = False
            
            # Update current block if selected
            if self.current_block:
                self.blocks[self.current_block].row_spacing_m = row_spacing_m
                self.blocks[self.current_block].gcr = gcr
                
                # Check if this is the first block and set project default
                if len(self.blocks) == 1:
                    self.set_project_default_row_spacing(row_spacing_m)
                
                # Redraw the block
                self.draw_block()
                
        except ValueError:
            messagebox.showerror("Error", "Row spacing must be a valid number")

    def update_row_spacing_from_gcr(self, event=None, *args):
        """Update row spacing field when GCR changes"""
        # Return early if we're programmatically updating the UI
        if hasattr(self, 'updating_ui') and self.updating_ui:
            return
        
        try:
            # Get GCR value
            gcr_str = self.gcr_var.get().strip()
            if gcr_str == "--":
                return
                
            gcr = float(gcr_str)
            if gcr <= 0 or gcr > 1.0:
                messagebox.showerror("Error", "GCR must be between 0 and 1.0")
                return
                
            # Get module length from templates
            module_length_m, error = self.get_module_length_from_templates()
            if error:
                messagebox.showerror("Error", f"Cannot calculate row spacing: {error}")
                return
                
            # Calculate row spacing
            row_spacing_m = module_length_m / gcr
            row_spacing_ft = self.m_to_ft(row_spacing_m)
            
            # Update row spacing field
            self.updating_ui = True
            self.row_spacing_var.set(f"{row_spacing_ft:.1f}")
            self.updating_ui = False
            
            # Update current block if selected
            if self.current_block:
                self.blocks[self.current_block].row_spacing_m = row_spacing_m
                self.blocks[self.current_block].gcr = gcr
                
                # Check if this is the first block and set project default
                if len(self.blocks) == 1:
                    self.set_project_default_row_spacing(row_spacing_m)
                
                # Redraw the block
                self.draw_block()
                
        except ValueError:
            messagebox.showerror("Error", "GCR must be a valid number")

    def set_project_default_row_spacing(self, row_spacing_m):
        """Set the project default row spacing and auto-save"""
        if self.current_project:
            # Update this instance's project reference
            self.current_project.default_row_spacing_m = row_spacing_m
            self.current_project.update_modified_date()
            
            # Also update the main app's project reference to keep them in sync
            main_app = self.winfo_toplevel()
            if hasattr(main_app, 'current_project') and main_app.current_project:
                main_app.current_project.default_row_spacing_m = row_spacing_m
                main_app.current_project.update_modified_date()
                
                # Auto-save the project
                if hasattr(main_app, 'save_project'):
                    main_app.save_project()

    def check_module_consistency_before_placement(self, template):
        """Check if placing this template would create module inconsistency"""
        if not template or not template.module_spec:
            return False, "Template has no module specification"
        
        # Get current module length from existing templates
        current_module_length_m, error = self.get_module_length_from_templates()
        
        # If no existing templates, this is fine
        if error:
            return True, None
        
        # Check if new template matches existing ones
        new_module_length_m = template.module_spec.length_mm / 1000
        
        if abs(current_module_length_m - new_module_length_m) > 0.001:  # Small tolerance for floating point
            current_ft = current_module_length_m * 3.28084
            new_ft = new_module_length_m * 3.28084
            warning_msg = (f"This template has a different module length ({new_ft:.2f}ft) "
                        f"than existing templates ({current_ft:.2f}ft). "
                        f"This will cause inconsistent GCR calculations.")
            return False, warning_msg
        
        return True, None