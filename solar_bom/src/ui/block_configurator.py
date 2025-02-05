import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Dict, List
from ..models.block import BlockConfig
from ..models.tracker import TrackerTemplate
from ..models.inverter import InverterSpec
from .inverter_manager import InverterManager

class BlockConfigurator(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        
        # State management
        self.blocks: Dict[str, BlockConfig] = {}  # Store block configurations
        self.current_block: Optional[str] = None  # Currently selected block ID
        self.available_templates: Dict[str, TrackerTemplate] = {}  # Available tracker templates
        self.selected_inverter = None
        self.tracker_templates: Dict[str, TrackerTemplate] = {}
        self.dragging = False
        self.drag_template = None
        self.drag_start = None
        self.setup_ui()
        
    def setup_ui(self):
        """Create and arrange UI components"""
        # Main container with padding
        main_container = ttk.Frame(self, padding="10")
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
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
        
        templates_frame = ttk.LabelFrame(config_frame, text="Tracker Templates", padding="5")
        templates_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))

        self.template_listbox = tk.Listbox(templates_frame, height=5)
        self.template_listbox.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))
        self.template_listbox.bind('<<ListboxSelect>>', self.on_template_select)

        # Block Dimensions
        dims_frame = ttk.LabelFrame(config_frame, text="Dimensions", padding="5")
        dims_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        ttk.Label(dims_frame, text="Width (m):").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.width_var = tk.StringVar(value="100")
        ttk.Entry(dims_frame, textvariable=self.width_var).grid(row=0, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(dims_frame, text="Height (m):").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.height_var = tk.StringVar(value="50")
        ttk.Entry(dims_frame, textvariable=self.height_var).grid(row=1, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(dims_frame, text="Row Spacing (m):").grid(row=2, column=0, padx=5, pady=2, sticky=tk.W)
        self.row_spacing_var = tk.StringVar(value="6")
        ttk.Entry(dims_frame, textvariable=self.row_spacing_var).grid(row=2, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        # GCR (Ground Coverage Ratio)
        ttk.Label(dims_frame, text="GCR:").grid(row=3, column=0, padx=5, pady=2, sticky=tk.W)
        self.gcr_var = tk.StringVar(value="0.4")
        ttk.Entry(dims_frame, textvariable=self.gcr_var).grid(row=3, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        # Inverter Selection
        inverter_frame = ttk.LabelFrame(config_frame, text="Inverter", padding="5")
        inverter_frame.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))

        ttk.Label(inverter_frame, text="Selected Inverter:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.inverter_label = ttk.Label(inverter_frame, text="None")
        self.inverter_label.grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)
        ttk.Button(inverter_frame, text="Select Inverter", command=self.select_inverter).grid(row=0, column=2, padx=5, pady=2)

        # Canvas for block visualization
        canvas_frame = ttk.LabelFrame(main_container, text="Block Layout", padding="5")
        canvas_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.canvas = tk.Canvas(canvas_frame, width=800, height=400, bg='white')
        self.canvas.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
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
            # Create new block config
            block = BlockConfig(
                block_id=block_id,
                inverter=self.selected_inverter,  # Will need to add inverter selection
                tracker_template=None,  # Will need to add template selection
                width_m=float(self.width_var.get()),
                height_m=float(self.height_var.get()),
                row_spacing_m=float(self.row_spacing_var.get()),
                gcr=float(self.gcr_var.get()),
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
        
        # Update UI with block data
        self.block_id_var.set(block.block_id)
        self.width_var.set(str(block.width_m))
        self.height_var.set(str(block.height_m))
        self.row_spacing_var.set(str(block.row_spacing_m))
        self.gcr_var.set(str(block.gcr))
        
        # Update canvas
        self.draw_block()
        
    def clear_config_display(self):
        """Clear block configuration display"""
        self.block_id_var.set("")
        self.width_var.set("100")
        self.height_var.set("50")
        self.row_spacing_var.set("6")
        self.gcr_var.set("0.4")
        self.canvas.delete("all")
        
    def draw_block(self):
        """Draw current block layout on canvas"""
        if not self.current_block:
            return
            
        block = self.blocks[self.current_block]
        
        # Clear canvas
        self.canvas.delete("all")
        
        # Calculate scale factor to fit block in canvas
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        
        scale_x = (canvas_width - 20) / block.width_m
        scale_y = (canvas_height - 20) / block.height_m
        scale = min(scale_x, scale_y)
        
        # Draw block outline
        self.canvas.create_rectangle(
            10, 10,
            10 + block.width_m * scale,
            10 + block.height_m * scale,
            outline='black'
        )
        
        # Draw row lines based on row spacing
        y = 10
        while y < 10 + block.height_m * scale:
            self.canvas.create_line(
                10, y,
                10 + block.width_m * scale, y,
                fill='gray', dash=(2, 4)
            )
            y += block.row_spacing_m * scale
            
        # Draw trackers if any are placed
        for pos in block.tracker_positions:
            # TODO: Implement tracker drawing
            pass

    def draw_tracker(self, x, y, template, tag=None):
        """Draw a tracker on the canvas"""
        if not template:
            return
            
        dims = template.get_physical_dimensions()
        scale = self.get_canvas_scale()
        
        # Convert meters to pixels
        width = dims[0] * scale
        height = dims[1] * scale
        
        # Draw rectangle
        self.canvas.create_rectangle(
            x, y, x + width, y + height,
            fill='lightblue', outline='blue',
            tags=tag if tag else 'tracker'
        )

    def get_canvas_scale(self):
        """Calculate scale factor (pixels per meter)"""
        if not self.current_block:
            return 1.0
        block = self.blocks[self.current_block]
        canvas_width = self.canvas.winfo_width() - 20
        canvas_height = self.canvas.winfo_height() - 20
        scale_x = canvas_width / block.width_m
        scale_y = canvas_height / block.height_m
        return min(scale_x, scale_y)

    def on_canvas_click(self, event):
        """Handle canvas click for tracker placement"""
        if not self.current_block or not self.drag_template:
            return
        self.dragging = True
        self.drag_start = (event.x, event.y)

    def on_canvas_drag(self, event):
        """Handle canvas drag for tracker movement"""
        if not self.dragging:
            return
        self.canvas.delete('drag_preview')
        self.draw_tracker(event.x, event.y, self.drag_template, 'drag_preview')

    def on_canvas_release(self, event):
        """Handle canvas release for tracker placement"""
        if not self.dragging:
            return
            
        self.dragging = False
        self.canvas.delete('drag_preview')
        
        # Convert canvas coordinates to meters
        scale = self.get_canvas_scale()
        block = self.blocks[self.current_block]
        x_m = (event.x - 10) / scale
        y_m = (event.y - 10) / scale
        
        # Snap to grid based on row spacing
        y_m = round(y_m / block.row_spacing_m) * block.row_spacing_m
        
        # Add tracker if within bounds
        if 0 <= x_m <= block.width_m and 0 <= y_m <= block.height_m:
            block.tracker_positions.append(TrackerPosition(x=x_m, y=y_m, rotation=0.0))
            self.draw_block()

    def load_templates(self):
        """Load tracker templates from file"""
        template_path = Path('data/tracker_templates.json')
        if template_path.exists():
            with open(template_path, 'r') as f:
                data = json.load(f)
                self.tracker_templates = {name: TrackerTemplate(**template) 
                                    for name, template in data.items()}
                
    def on_template_select(self, event=None):
        """Handle template selection"""
        selection = self.template_listbox.curselection()
        if selection:
            template_name = self.template_listbox.get(selection[0])
            self.drag_template = self.tracker_templates[template_name]