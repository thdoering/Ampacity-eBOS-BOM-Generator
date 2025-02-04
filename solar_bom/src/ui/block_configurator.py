import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Dict, List
from ..models.block import BlockConfig
from ..models.tracker import TrackerTemplate
from ..models.inverter import InverterSpec

class BlockConfigurator(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        
        # State management
        self.blocks: Dict[str, BlockConfig] = {}  # Store block configurations
        self.current_block: Optional[str] = None  # Currently selected block ID
        self.available_templates: Dict[str, TrackerTemplate] = {}  # Available tracker templates
        
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
        
    def create_new_block(self):
        """Create a new block configuration"""
        # Generate unique block ID
        block_id = f"Block_{len(self.blocks) + 1}"
        
        try:
            # Create new block config
            block = BlockConfig(
                block_id=block_id,
                inverter=None,  # Will need to add inverter selection
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
            
    def on_canvas_click(self, event):
        """Handle canvas click for tracker placement"""
        # TODO: Implement tracker placement
        pass
        
    def on_canvas_drag(self, event):
        """Handle canvas drag for tracker movement"""
        # TODO: Implement tracker dragging
        pass
        
    def on_canvas_release(self, event):
        """Handle canvas release for tracker placement"""
        # TODO: Implement tracker placement completion
        pass