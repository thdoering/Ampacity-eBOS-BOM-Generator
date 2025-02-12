import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Dict, List
from ..models.block import BlockConfig
from ..models.tracker import TrackerTemplate
from ..models.inverter import InverterSpec
from .inverter_manager import InverterManager
from pathlib import Path
import json
from ..models.tracker import ModuleOrientation
from ..models.module import ModuleSpec, ModuleType, ModuleOrientation
from ..models.tracker import TrackerPosition

class BlockConfigurator(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        
        # State management
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
        
        # Right side - Block Configuration
        config_frame = ttk.LabelFrame(main_container, text="Block Configuration", padding="5")
        config_frame.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Block ID
        ttk.Label(config_frame, text="Block ID:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.block_id_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.block_id_var).grid(row=0, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
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
        
        # Inverter Selection
        inverter_frame = ttk.LabelFrame(config_frame, text="Inverter", padding="5")
        inverter_frame.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))

        ttk.Label(inverter_frame, text="Selected Inverter:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.inverter_label = ttk.Label(inverter_frame, text="None")
        self.inverter_label.grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)
        ttk.Button(inverter_frame, text="Select Inverter", command=self.select_inverter).grid(row=0, column=2, padx=5, pady=2)

        # Templates List Frame
        templates_frame = ttk.LabelFrame(config_frame, text="Tracker Templates", padding="5")
        templates_frame.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))

        self.template_listbox = tk.Listbox(templates_frame, height=5)
        self.template_listbox.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))
        self.template_listbox.bind('<<ListboxSelect>>', self.on_template_select)

        # Canvas frame for block layout - on the right side
        canvas_frame = ttk.LabelFrame(main_container, text="Block Layout", padding="5")
        canvas_frame.grid(row=0, rowspan=2, column=2, padx=5, pady=5)

        # Fixed size canvas
        self.canvas = tk.Canvas(canvas_frame, width=1000, height=800, bg='white')
        self.canvas.grid(row=0, column=0, padx=5, pady=5)

        # Canvas bindings for clicking and dragging trackers
        self.canvas.bind('<Button-1>', self.on_canvas_click)
        self.canvas.bind('<B1-Motion>', self.on_canvas_drag)
        self.canvas.bind('<ButtonRelease-1>', self.on_canvas_release)

        # Keyboard bindings for deleting trackers
        self.canvas.bind('<Delete>', self.delete_selected_tracker)
        self.canvas.bind('<BackSpace>', self.delete_selected_tracker)
        
        # Make canvas and frames expandable
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        main_container.grid_columnconfigure(1, weight=1)
        main_container.grid_rowconfigure(1, weight=1)
        
        # Bind canvas events for future drag and drop implementation
        self.canvas.bind('<Button-1>', self.on_canvas_click)
        self.canvas.bind('<B1-Motion>', self.on_canvas_drag)
        self.canvas.bind('<ButtonRelease-1>', self.on_canvas_release)

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
        if self.current_block:
            self.blocks[self.current_block].inverter = inverter
        dialog.destroy()
        
    def create_new_block(self):
        """Create a new block configuration"""
        # Generate unique block ID
        block_id = f"Block_{len(self.blocks) + 1}"
        
        try:
            # Convert feet to meters for storage
            row_spacing_ft = float(self.row_spacing_var.get())
            
            # Create new block config
            block = BlockConfig(
                block_id=block_id,
                inverter=self.selected_inverter,
                tracker_template=None,
                width_m=20,  # Initial minimum width
                height_m=20,  # Initial minimum height
                row_spacing_m=self.ft_to_m(row_spacing_ft),
                ns_spacing_m=float(self.ns_spacing_var.get()),
                gcr=0.0,  # This will be calculated when a tracker template is assigned
                description=f"New block {block_id}"
            )
            
            # Add to blocks dictionary
            self.blocks[block_id] = block
            
            # Update listbox
            self.block_listbox.insert(tk.END, block_id)
            
            # Select new block
            self.block_listbox.selection_clear(0, tk.END)
            self.block_listbox.selection_set(tk.END)
            self.on_block_select()
            
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
        self.width_var.set("200")
        self.height_var.set("165")
        self.row_spacing_var.set("18.5")
        self.gcr_var.set("0.4")
        self.canvas.delete("all")
        
    def draw_block(self):
        """Draw current block layout on canvas"""
        if not self.current_block:
            return
                
        block = self.blocks[self.current_block]
        
        # Clear canvas and grid lines list
        self.canvas.delete("all")
        self.grid_lines = []
        
        # Calculate block dimensions
        block_width_m, block_height_m = self.calculate_block_dimensions()
        
        # Update block dimensions
        block.width_m = block_width_m
        block.height_m = block_height_m
        
        # Get scale factor
        scale = self.get_canvas_scale()
        
        # Draw grid lines
        x = 10
        while x < block_width_m * scale + 10:
            line_id = self.canvas.create_line(
                x, 10,
                x, block_height_m * scale + 10,
                fill='gray', dash=(2, 4)
            )
            self.grid_lines.append(line_id)
            x += block.row_spacing_m * scale
        
        y = 10
        while y < block_height_m * scale + 10:
            line_id = self.canvas.create_line(
                10, y,
                block_width_m * scale + 10, y,
                fill='gray', dash=(2, 4)
            )
            self.grid_lines.append(line_id)
            y += float(self.ns_spacing_var.get()) * scale
        
        # Draw existing trackers
        for pos in block.tracker_positions:
            x = 10 + pos.x * scale
            y = 10 + pos.y * scale
            self.draw_tracker(x, y, pos.template)

    def draw_tracker(self, x, y, template, tag=None):
        """Draw a tracker on the canvas with detailed module layout"""
        if not template:
            return
                
        dims = template.get_physical_dimensions()
        scale = self.get_canvas_scale()
        
        # Get module dimensions
        if template.module_orientation == ModuleOrientation.PORTRAIT:
            module_height = template.module_spec.width_mm / 1000 * scale
            module_width = template.module_spec.length_mm / 1000 * scale
        else:
            module_height = template.module_spec.length_mm / 1000 * scale
            module_width = template.module_spec.width_mm / 1000 * scale

        # Create group tag for all elements
        group_tag = tag if tag else f'tracker_{x}_{y}'
        
        # Calculate number of modules above and below motor
        total_modules = template.modules_per_string * template.strings_per_tracker
        modules_per_string = template.modules_per_string
        strings_above_motor = template.strings_per_tracker - 1
        modules_above_motor = modules_per_string * strings_above_motor
        modules_below_motor = modules_per_string

        # Calculate total tracker height for torque tube
        total_height = (
            (total_modules * module_height) +  # All modules
            ((total_modules - 1) * template.module_spacing_m * scale) +  # Module spacing
            template.motor_gap_m * scale  # Motor gap
        )
        
        # Draw torque tube through center
        self.canvas.create_line(
            x + module_width/2, y,
            x + module_width/2, y + total_height,
            width=3, fill='gray', tags=group_tag
        )
        
        # Draw all modules
        y_pos = y
        modules_drawn = 0
        
        # Draw modules above motor
        for i in range(modules_above_motor):
            self.canvas.create_rectangle(
                x, y_pos,
                x + module_width, y_pos + module_height,
                fill='lightblue', outline='blue', tags=group_tag
            )
            modules_drawn += 1
            y_pos += module_height + template.module_spacing_m * scale

        # Draw motor
        motor_y = y_pos
        self.canvas.create_oval(
            x + module_width/2 - 5, motor_y - 5,
            x + module_width/2 + 5, motor_y + 5,
            fill='red', tags=group_tag
        )
        y_pos += template.motor_gap_m * scale

        # Draw modules below motor
        for i in range(modules_below_motor):
            self.canvas.create_rectangle(
                x, y_pos,
                x + module_width, y_pos + module_height,
                fill='lightblue', outline='blue', tags=group_tag
            )
            y_pos += module_height + template.module_spacing_m * scale

    def get_canvas_scale(self):
        """Calculate scale factor (pixels per meter)"""
        if not self.current_block:
            return 1.0
            
        # Get block dimensions
        block_width_m, block_height_m = self.calculate_block_dimensions()
        
        # Get canvas dimensions (subtract padding)
        canvas_width = self.canvas.winfo_width() - 20
        canvas_height = self.canvas.winfo_height() - 20
        
        # Calculate scale factors
        scale_x = canvas_width / block_width_m
        scale_y = canvas_height / block_height_m
        
        # Use minimum scale factor that ensures everything is visible
        scale = min(scale_x, scale_y)
        
        # Don't let scale get too small (prevents objects from becoming too tiny)
        MIN_SCALE = 0.5  # 0.5 pixels per meter minimum
        return max(scale, MIN_SCALE)

    def on_canvas_click(self, event):
        """Handle canvas click for tracker placement"""
        if not self.current_block or not self.drag_template:
            return
        
        # First check if we're clicking on an existing tracker
        if self.select_tracker(event.x, event.y):
            return
        
        self.dragging = True
        self.drag_start = (event.x, event.y)
        
        # Calculate snapped position
        block = self.blocks[self.current_block]
        scale = self.get_canvas_scale()
        x_m = (event.x - 10) / scale
        y_m = (event.y - 10) / scale
        
        # Snap to grid
        x_m = round(x_m / block.row_spacing_m) * block.row_spacing_m
        y_m = round(y_m / float(self.ns_spacing_var.get())) * float(self.ns_spacing_var.get())
        
        # Draw preview
        x = x_m * scale + 10
        y = y_m * scale + 10
        self.draw_tracker(x, y, self.drag_template, 'drag_preview')

    def on_canvas_drag(self, event):
        """Handle canvas drag for tracker movement"""
        if not self.dragging:
            return
        
        # Delete old preview
        self.canvas.delete('drag_preview')
        
        # Calculate snapped position
        block = self.blocks[self.current_block]
        scale = self.get_canvas_scale()
        x_m = (event.x - 10) / scale
        y_m = (event.y - 10) / scale
        
        # Snap to grid
        x_m = round(x_m / block.row_spacing_m) * block.row_spacing_m
        y_m = round(y_m / float(self.ns_spacing_var.get())) * float(self.ns_spacing_var.get())
        
        # Draw new preview
        x = x_m * scale + 10
        y = y_m * scale + 10
        self.draw_tracker(x, y, self.drag_template, 'drag_preview')

    def on_canvas_release(self, event):
        """Handle canvas release for tracker placement"""
        if not self.dragging or not self.current_block or not self.drag_template:
            return
                
        self.dragging = False
        self.canvas.delete('drag_preview')
        
        # Convert canvas coordinates to meters
        scale = self.get_canvas_scale()
        block = self.blocks[self.current_block]
        x_m = (event.x - 10) / scale
        y_m = (event.y - 10) / scale
        
        # Snap to grid based on row spacing
        x_m = round(x_m / block.row_spacing_m) * block.row_spacing_m
        y_m = round(y_m / float(self.ns_spacing_var.get())) * float(self.ns_spacing_var.get())
        
        # Add tracker if within bounds
        dims = self.drag_template.get_physical_dimensions()
        
        # Create new TrackerPosition with template
        pos = TrackerPosition(x=x_m, y=y_m, rotation=0.0, template=self.drag_template)
        block.tracker_positions.append(pos)
        
        # Update block display and resize canvas
        self.draw_block()

    def load_templates(self):
        """Load tracker templates from file"""
        template_path = Path('data/tracker_templates.json')
        if template_path.exists():
            try:
                with open(template_path, 'r') as f:
                    data = json.load(f)
                    # Use current module if available, otherwise use default
                    module_spec = self._current_module if self._current_module else ModuleSpec(
                        manufacturer="Default",
                        model="Default",
                        type=ModuleType.MONO_PERC,
                        length_mm=2000,
                        width_mm=1000,
                        depth_mm=40,
                        weight_kg=25,
                        wattage=400,
                        vmp=40,
                        imp=10,
                        voc=48,
                        isc=10.5,
                        max_system_voltage=1500
                    )
                    
                    self.tracker_templates = {
                        name: TrackerTemplate(
                            template_name=name,
                            module_spec=module_spec,
                            module_orientation=ModuleOrientation(template.get('module_orientation', 'Portrait')),
                            modules_per_string=template.get('modules_per_string', 28),
                            strings_per_tracker=template.get('strings_per_tracker', 2),
                            module_spacing_m=template.get('module_spacing_m', 0.01),
                            motor_gap_m=template.get('motor_gap_m', 1.0)
                        ) 
                        for name, template in data.items()
                    }
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
        if (self.current_block and 
            self.blocks[self.current_block].tracker_template and 
            self.blocks[self.current_block].tracker_template.module_spec):
            try:
                template = self.blocks[self.current_block].tracker_template
                module_length = template.module_spec.length_mm / 1000  # convert to meters
                row_spacing = float(self.row_spacing_var.get())
                row_spacing_m = self.ft_to_m(row_spacing)
                gcr = module_length / row_spacing_m
                self.gcr_label.config(text=f"{gcr:.3f}")
            except (ValueError, ZeroDivisionError):
                self.gcr_label.config(text="--")
        else:
            self.gcr_label.config(text="--")

    def delete_selected_tracker(self, event=None):
        """Delete the currently selected tracker"""
        if not self.current_block or not self.selected_tracker:
            return
        
        # Find and remove the selected tracker
        block = self.blocks[self.current_block]
        for i, pos in enumerate(block.tracker_positions):
            if (pos.x, pos.y) == self.selected_tracker:
                block.tracker_positions.pop(i)
                break
        
        self.selected_tracker = None
        self.draw_block()

    def select_tracker(self, x, y):
        """Select tracker at given coordinates"""
        if not self.current_block:
            return

        block = self.blocks[self.current_block]
        scale = self.get_canvas_scale()
        
        # Convert canvas coordinates to meters
        x_m = (x - 10) / scale
        y_m = (y - 10) / scale
        
        # Check if click is within any tracker
        for pos in block.tracker_positions:
            dims = pos.template.get_physical_dimensions()
            if (pos.x <= x_m <= pos.x + dims[0] and 
                pos.y <= y_m <= pos.y + dims[1]):
                self.selected_tracker = (pos.x, pos.y)
                self.draw_block()
                # Highlight selected tracker
                x_canvas = 10 + pos.x * scale
                y_canvas = 10 + pos.y * scale
                self.canvas.create_rectangle(
                    x_canvas - 2, y_canvas - 2,
                    x_canvas + dims[0] * scale + 2,
                    y_canvas + dims[1] * scale + 2,
                    outline='red', width=2
                )
                return True
        
        self.selected_tracker = None
        self.draw_block()
        return False
    
    def calculate_block_dimensions(self):
        """Calculate block dimensions based on placed trackers"""
        if not self.current_block or not self.blocks[self.current_block].tracker_positions:
            # If no trackers, use dimensions that would fit one tracker with room to spare
            if self.drag_template:
                # Use dimensions of selected template
                dims = self.drag_template.get_physical_dimensions()
                initial_width = dims[0] * 3  # Room for 3 trackers wide
                initial_height = dims[1] * 3  # Room for 3 trackers tall
                return (max(initial_width, 50), max(initial_height, 50))  # Min 50m in each direction
            return (50, 50)  # Default minimum size in meters if no template selected
            
        block = self.blocks[self.current_block]
        
        # Find max x and y coordinates including tracker dimensions
        max_x = 0
        max_y = 0
        for pos in block.tracker_positions:
            dims = pos.template.get_physical_dimensions()
            max_x = max(max_x, pos.x + dims[0])
            max_y = max(max_y, pos.y + dims[1])
        
        # Add padding (40% extra space)
        # This ensures room for additional trackers in any direction
        padding_x = max(50, max_x * 0.4)  # At least 50m padding or 40% of current width
        padding_y = max(50, max_y * 0.4)  # At least 50m padding or 40% of current height
        
        return (max_x + padding_x, max_y + padding_y)

    def resize_canvas(self):
        """Resize canvas to fit all trackers plus padding"""
        if not self.current_block:
            return
            
        # Calculate required dimensions
        block_width_m, block_height_m = self.calculate_block_dimensions()
        
        # Get current window dimensions
        window_width = self.winfo_width()
        window_height = self.winfo_height()
        
        # Calculate required canvas size to maintain scale
        # Use 80% of window size as maximum canvas dimension
        max_canvas_width = int(window_width * 0.8)
        max_canvas_height = int(window_height * 0.8)
        
        # Calculate scale factors
        scale_x = max_canvas_width / block_width_m
        scale_y = max_canvas_height / block_height_m
        scale = min(scale_x, scale_y)
        
        # Calculate new canvas dimensions
        canvas_width = int(block_width_m * scale) + 20  # Add padding
        canvas_height = int(block_height_m * scale) + 20
        
        # Resize canvas
        self.canvas.config(width=canvas_width, height=canvas_height)