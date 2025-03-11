import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Dict, List
from ..models.block import BlockConfig, WiringType, TrackerPosition, DeviceType, WiringConfig, CollectionPoint
from ..models.tracker import TrackerTemplate, TrackerPosition, ModuleOrientation
from ..models.inverter import InverterSpec
from .inverter_manager import InverterManager
from pathlib import Path
import json
from ..models.module import ModuleSpec, ModuleType, ModuleOrientation
from ..utils.undo_manager import UndoManager
from copy import deepcopy

class BlockConfigurator(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        
        # State management
        self.device_type_var = None
        self.inverter_frame = None
        self._current_module = None
        self.available_templates = {}
        self.blocks: Dict[str, BlockConfig] = {}  # Store block configurations
        self.current_block: Optional[str] = None  # Currently selected block ID
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
        self.row_spacing_var.trace('w', lambda *args: self.calculate_gcr())

        # GCR (Ground Coverage Ratio) - calculated
        ttk.Label(spacing_frame, text="GCR:").grid(row=0, column=2, padx=5, pady=2, sticky=tk.W)
        self.gcr_label = ttk.Label(spacing_frame, text="--")
        self.gcr_label.grid(row=0, column=3, padx=5, pady=2, sticky=tk.W)

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

        # Add spacing after device frame
        ttk.Label(config_frame, text="").grid(row=6, column=0, pady=5)  # Spacer

        # Configure Wiring button
        ttk.Button(config_frame, text="Configure Wiring", 
                command=self.configure_wiring).grid(row=7, column=0, columnspan=2, padx=5, pady=5)

        # Templates List Frame
        templates_frame = ttk.LabelFrame(config_frame, text="Tracker Templates", padding="5")
        templates_frame.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))

        self.template_listbox = tk.Listbox(templates_frame, height=5)
        self.template_listbox.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))
        self.template_listbox.bind('<<ListboxSelect>>', self.on_template_select)

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
        # Generate unique block ID
        block_id = f"Block_{len(self.blocks) + 1}"
        
        try:
            # Get selected template if any
            selection = self.template_listbox.curselection()
            selected_template = None
            if selection:
                template_name = self.template_listbox.get(selection[0])
                selected_template = self.tracker_templates.get(template_name)
            
            # Get row spacing value and convert feet to meters
            try:
                row_spacing_ft = float(self.row_spacing_var.get())
                row_spacing_m = self.ft_to_m(row_spacing_ft)
                device_width_m = 0.91  # 3ft in meters
                initial_device_x = row_spacing_m / 2 + (device_width_m / 2)
            except ValueError:
                row_spacing_m = 6.0  # Default to 6m if invalid input
                device_width_m = 0.91  # 3ft in meters
                initial_device_x = row_spacing_m / 2 + (device_width_m / 2)
            
            block = BlockConfig(
                block_id=block_id,
                inverter=self.selected_inverter,
                tracker_template=selected_template,
                width_m=20,  # Initial minimum width
                height_m=20,  # Initial minimum height
                row_spacing_m=row_spacing_m,
                ns_spacing_m=float(self.ns_spacing_var.get()),
                gcr=0.0,  # This will be calculated when a tracker template is assigned
                description=f"New block {block_id}",
                device_spacing_m=self.ft_to_m(float(self.device_spacing_var.get())),
                device_x=initial_device_x,  # Set initial X position
                device_y=0.0  # Start at top
            )
            
            # Add to blocks dictionary
            self.blocks[block_id] = block
            
            # Update listbox
            self.block_listbox.insert(tk.END, block_id)
            
            # Select new block
            self.block_listbox.selection_clear(0, tk.END)
            self.block_listbox.selection_set(tk.END)
            self.on_block_select()
            
            # Save initial empty state
            self._push_state("Create block")
            
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid input: {str(e)}")
        
    def delete_block(self):
        """Delete currently selected block"""
        if not self.current_block:
            return
            
        if messagebox.askyesno("Confirm", f"Delete block {self.current_block}?"):
            # Remove from dictionary
            del self.blocks[self.current_block]
            
            # Update listbox
            selection = self.block_listbox.curselection()
            if selection:
                self.block_listbox.delete(selection[0])
                
            # Clear current block
            self.current_block = None
            self.clear_config_display()
            
    def on_block_select(self, event=None):
        """Handle block selection from listbox"""
        selection = self.block_listbox.curselection()
        if not selection:
            return
                
        block_id = self.block_listbox.get(selection[0])
        self.current_block = block_id
        block = self.blocks[block_id]
        
        # Update UI with block data (convert meters to feet)
        self.block_id_var.set(block.block_id)
        self.row_spacing_var.set(str(self.m_to_ft(block.row_spacing_m)))
        self.calculate_gcr()  # Update the GCR label
        
        # Update canvas
        self.draw_block()
        
    def clear_config_display(self):
        """Clear block configuration display"""
        self.block_id_var.set("")
        # self.width_var.set("200")
        # self.height_var.set("165")
        self.row_spacing_var.set("18.5")
        # self.gcr_var.set("0.4")
        self.canvas.delete("all")
        
    def draw_block(self):
        """Draw current block layout on canvas"""
        if not self.current_block:
            return
                
        try:
            if not self.validate_device_clearances():
                return

            block = self.blocks[self.current_block]

            # Clear canvas and grid lines list
            self.canvas.delete("all")
            self.grid_lines = []
                        
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
        
        except (ValueError, TypeError):
            # Silently ignore transient errors during editing
            pass

    def draw_tracker(self, x, y, template, tag=None):
        """Draw a tracker on the canvas with detailed module layout.
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
        
        # Calculate number of modules
        total_modules = template.modules_per_string * template.strings_per_tracker
        modules_per_string = template.modules_per_string
        strings_above_motor = template.strings_per_tracker - 1
        modules_above_motor = modules_per_string * strings_above_motor
        modules_below_motor = modules_per_string

        # Draw torque tube through center
        self.canvas.create_line(
            x + module_width * scale/2, y,
            x + module_width * scale/2, y + dims[1] * scale,
            width=3, fill='gray', tags=group_tag
        )

        # Draw all modules
        y_pos = y
        modules_drawn = 0
        
        # Draw modules above motor
        for i in range(modules_above_motor):
            self.canvas.create_rectangle(
                x, y_pos,
                x + module_width * scale, y_pos + module_height * scale,
                fill='lightblue', outline='blue', tags=group_tag
            )
            modules_drawn += 1
            y_pos += (module_height + template.module_spacing_m) * scale

        # Draw motor
        motor_y = y_pos
        self.canvas.create_oval(
            x + module_width * scale/2 - 5, motor_y - 5,
            x + module_width * scale/2 + 5, motor_y + 5,
            fill='red', tags=group_tag
        )
        y_pos += template.motor_gap_m * scale

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
        
        # Save state before adding tracker
        self._push_state("Before place tracker")
        
        # Create new TrackerPosition
        pos = TrackerPosition(x=x_m, y=y_m, rotation=0.0, template=self.drag_template)
        pos.calculate_string_positions() # Calculate string and collection point positions
        block.tracker_positions.append(pos)
        
        # Update block display
        self.draw_block()

    def load_templates(self):
        """Load tracker templates from file"""
        template_path = Path('data/tracker_templates.json')
        if template_path.exists():
            try:
                with open(template_path, 'r') as f:
                    data = json.load(f)
                    # Create ModuleSpec objects from the stored module_spec data in templates
                    self.tracker_templates = {}
                    for name, template in data.items():
                        # Extract the module spec data
                        module_data = template.get('module_spec', {})
                        
                        # Create proper ModuleSpec object with correct values
                        module_spec = ModuleSpec(
                            manufacturer=module_data.get('manufacturer', 'Default'),
                            model=module_data.get('model', 'Default'),
                            type=ModuleType.MONO_PERC,  # Default type
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
                        
                        # Create TrackerTemplate with the correct module_spec
                        self.tracker_templates[name] = TrackerTemplate(
                            template_name=name,
                            module_spec=module_spec,
                            module_orientation=ModuleOrientation(template.get('module_orientation', 'Portrait')),
                            modules_per_string=template.get('modules_per_string', 28),
                            strings_per_tracker=template.get('strings_per_tracker', 2),
                            module_spacing_m=template.get('module_spacing_m', 0.01),
                            motor_gap_m=template.get('motor_gap_m', 1.0)
                        )
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load templates: {str(e)}")
                self.tracker_templates = {}
        else:
            self.tracker_templates = {}
                
    def on_template_select(self, event=None):
        """Handle template selection"""
        selection = self.template_listbox.curselection()
        if selection:
            template_name = self.template_listbox.get(selection[0])
            self.drag_template = self.tracker_templates.get(template_name)
            # Update current block's template if one is selected
            if self.current_block and self.drag_template:
                self.blocks[self.current_block].tracker_template = self.drag_template
                self.calculate_gcr()  # Recalculate GCR with new template
                
                # Update string current values
                if self.drag_template.module_spec:
                    module = self.drag_template.module_spec
                    self.string_current = module.imp
                    self.nec_current = module.isc * 1.25
                    self.string_current_label.config(text=f"{self.string_current:.2f} A")
                    self.nec_current_label.config(text=f"{self.nec_current:.2f} A")
                else:
                    self.string_current = 0.0
                    self.nec_current = 0.0
                    self.string_current_label.config(text="-- A")
                    self.nec_current_label.config(text="-- A")

    def update_template_list(self):
        """Update template listbox with available templates"""
        self.template_listbox.delete(0, tk.END)
        for name in self.tracker_templates.keys():
            self.template_listbox.insert(tk.END, name)

    def m_to_ft(self, meters):
        return meters * 3.28084

    def ft_to_m(self, feet):
        return feet / 3.28084
    
    def calculate_gcr(self):
        """Calculate Ground Coverage Ratio"""
        if not self.current_block:
            return
            
        if not self.blocks[self.current_block].tracker_template:
            return
            
        if not self.blocks[self.current_block].tracker_template.module_spec:
            return
            
        try:
            template = self.blocks[self.current_block].tracker_template
            module_length = template.module_spec.length_mm / 1000  # convert to meters
            row_spacing = float(self.row_spacing_var.get())
            row_spacing_m = self.ft_to_m(row_spacing)
            gcr = module_length / row_spacing_m
            self.gcr_label.config(text=f"{gcr:.3f}")
        except (ValueError, ZeroDivisionError, AttributeError) as e:
            self.gcr_label.config(text="--")

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
        description = self.undo_manager.undo()
        if description:
            self.draw_block()

    def redo(self, event=None):
        """Handle redo command"""
        description = self.undo_manager.redo()
        if description:
            self.draw_block()

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
                    'template_name': pos.template.template_name if pos.template else None
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
                'device_spacing_m': block.device_spacing_m
            }
            
            # ADD THIS CODE HERE
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
                    'harness_cable_size': getattr(block.wiring_config, 'harness_cable_size', "8 AWG")
                }
                current_state[id]['wiring_config_data'] = wiring_config_data
            # END OF ADDED CODE
                
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
                device_x=initial_device_x,
                device_y=0.0
            )
            
            # Clear existing tracker positions
            block.tracker_positions = []
            
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
                    harness_cable_size=wiring_data.get('harness_cable_size', "8 AWG")
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

        # At the beginning of the method
        print("Opening wiring configurator")
        print("Current block has wiring config:", self.blocks[self.current_block].wiring_config is not None)
        if self.blocks[self.current_block].wiring_config:
            print("Wiring type:", self.blocks[self.current_block].wiring_config.wiring_type)
            print("Num positive collection points:", len(self.blocks[self.current_block].wiring_config.positive_collection_points))
            print("Num cable routes:", len(self.blocks[self.current_block].wiring_config.cable_routes))
            
        if not self.current_block:
            messagebox.showwarning("Warning", "Please select a block first")
            return
            
        from .wiring_configurator import WiringConfigurator
        WiringConfigurator(self, self.blocks[self.current_block])

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
            new_id = f"{base_name}{number + 1}"
            while new_id in self.blocks:
                number += 1
                new_id = f"{base_name}{number}"
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
        
        # Update listbox
        self.block_listbox.insert(tk.END, new_id)
        
        # Select the new block
        self.block_listbox.selection_clear(0, tk.END)
        self.block_listbox.selection_set(tk.END)
        self.block_listbox.see(tk.END)  # Ensure it's visible
        self.on_block_select()
        
        messagebox.showinfo("Success", f"Block copied as '{new_id}'")

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
        
        # Update the listbox - find the exact item that matches the old ID
        for i in range(self.block_listbox.size()):
            if self.block_listbox.get(i) == old_id:
                self.block_listbox.delete(i)
                self.block_listbox.insert(i, new_id)
                self.block_listbox.selection_set(i)
                break
        
        # Update current_block reference to the new ID
        self.current_block = new_id
        
        # Update UI display
        self.block_id_var.set(new_id)
        
        messagebox.showinfo("Success", f"Block renamed to '{new_id}'")