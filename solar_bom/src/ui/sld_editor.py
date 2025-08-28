"""
Single Line Diagram Editor Window
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Dict, Optional, Tuple
import math

from ..models.block import BlockConfig
from ..models.project import Project
from ..models.sld import (
    SLDDiagram, SLDElement, SLDConnection, SLDAnnotation,
    SLDElementType, ConnectionPortType, ConnectionPort
)
from ..utils.sld_symbols import ANSISymbols


class SLDEditor(tk.Toplevel):
    """Single Line Diagram Editor Window"""
    
    def __init__(self, parent, blocks: Dict[str, BlockConfig], project: Optional[Project] = None):
        super().__init__(parent)
        
        self.parent = parent
        self.blocks = blocks
        self.project = project
        
        # Window setup
        self.title(f"Single Line Diagram Editor - {project.metadata.name if project else 'New Diagram'}")
        self.geometry("1400x900")
        self.minsize(1000, 600)
        
        # Center the window
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (1400 // 2)
        y = (self.winfo_screenheight() // 2) - (900 // 2)
        self.geometry(f'+{x}+{y}')
        
        # SLD properties
        self.sld_diagram = None
        self.sld_elements = {}  # element_id -> canvas item IDs
        self.selected_element = None
        self.selected_items = []
        
        # Canvas properties
        self.canvas = None
        self.grid_visible = True
        self.grid_size = 10
        self.zoom_level = 1.0
        
        # Interaction properties
        self.dragging = False
        self.drag_data = {"x": 0, "y": 0, "item": None, "element_id": None}
        self.connecting = False
        self.connection_start = None
        self.temp_connection_line = None
        
        # Set up the UI
        self.setup_ui()
        
        # Generate initial SLD from blocks
        self.generate_sld_from_blocks()
        
        # Make window modal
        self.transient(parent)
        self.grab_set()
        
        # Handle window closing
        self.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def setup_ui(self):
        """Set up the user interface"""
        # Main container
        main_frame = ttk.Frame(self, padding="5")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)
        
        # Create menu bar
        self.create_menu()
        
        # Create toolbar
        self.create_toolbar(main_frame)
        
        # Create main content area with canvas and side panel
        content_frame = ttk.Frame(main_frame)
        content_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(5, 0))
        content_frame.grid_columnconfigure(0, weight=1)
        content_frame.grid_rowconfigure(0, weight=1)
        
        # Create canvas area
        self.create_canvas_area(content_frame)
        
        # Create status bar
        self.create_status_bar(main_frame)
    
    def create_menu(self):
        """Create the menu bar"""
        menubar = tk.Menu(self)
        self.configure(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Export PNG...", command=self.export_png, accelerator="Ctrl+P")
        file_menu.add_command(label="Export PDF...", command=self.export_pdf, accelerator="Ctrl+D")
        file_menu.add_separator()
        file_menu.add_command(label="Close", command=self.on_close, accelerator="Esc")
        
        # Edit menu
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="Select All", command=self.select_all, accelerator="Ctrl+A")
        edit_menu.add_command(label="Delete Selected", command=self.delete_selected, accelerator="Del")
        edit_menu.add_separator()
        edit_menu.add_command(label="Auto Layout", command=self.auto_layout)
        edit_menu.add_command(label="Clear All", command=self.clear_all)
        
        # View menu
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Zoom In", command=lambda: self.zoom(1.25), accelerator="Ctrl++")
        view_menu.add_command(label="Zoom Out", command=lambda: self.zoom(0.8), accelerator="Ctrl+-")
        view_menu.add_command(label="Fit to Window", command=self.fit_to_window, accelerator="Ctrl+0")
        view_menu.add_separator()
        view_menu.add_checkbutton(label="Show Grid", variable=tk.BooleanVar(value=True), 
                                  command=self.toggle_grid)
        
        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Refresh from Blocks", command=self.generate_sld_from_blocks)
        tools_menu.add_command(label="Validate Connections", command=self.validate_connections)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Keyboard Shortcuts", command=self.show_shortcuts)
        help_menu.add_command(label="About SLD Editor", command=self.show_about)
    
    def create_toolbar(self, parent):
        """Create the toolbar"""
        toolbar_frame = ttk.Frame(parent)
        toolbar_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # Layout section
        ttk.Label(toolbar_frame, text="Layout:").grid(row=0, column=0, padx=(5, 2))
        
        ttk.Button(
            toolbar_frame,
            text="Auto Layout",
            command=self.auto_layout
        ).grid(row=0, column=1, padx=2)
        
        ttk.Button(
            toolbar_frame,
            text="Align Left",
            command=lambda: self.align_selected('left')
        ).grid(row=0, column=2, padx=2)
        
        ttk.Button(
            toolbar_frame,
            text="Align Top",
            command=lambda: self.align_selected('top')
        ).grid(row=0, column=3, padx=2)
        
        ttk.Separator(toolbar_frame, orient='vertical').grid(row=0, column=4, padx=10, sticky='ns')
        
        # Zoom section
        ttk.Label(toolbar_frame, text="Zoom:").grid(row=0, column=5, padx=(5, 2))
        
        ttk.Button(
            toolbar_frame,
            text="-",
            width=3,
            command=lambda: self.zoom(0.8)
        ).grid(row=0, column=6, padx=1)
        
        self.zoom_label = ttk.Label(toolbar_frame, text="100%", width=6)
        self.zoom_label.grid(row=0, column=7, padx=2)
        
        ttk.Button(
            toolbar_frame,
            text="+",
            width=3,
            command=lambda: self.zoom(1.25)
        ).grid(row=0, column=8, padx=1)
        
        ttk.Button(
            toolbar_frame,
            text="Fit",
            command=self.fit_to_window
        ).grid(row=0, column=9, padx=2)
        
        ttk.Separator(toolbar_frame, orient='vertical').grid(row=0, column=10, padx=10, sticky='ns')
        
        # Export section
        ttk.Label(toolbar_frame, text="Export:").grid(row=0, column=11, padx=(5, 2))
        
        ttk.Button(
            toolbar_frame,
            text="Export PNG",
            command=self.export_png
        ).grid(row=0, column=12, padx=2)
        
        ttk.Button(
            toolbar_frame,
            text="Export PDF",
            command=self.export_pdf
        ).grid(row=0, column=13, padx=2)
        
        ttk.Separator(toolbar_frame, orient='vertical').grid(row=0, column=14, padx=10, sticky='ns')
        
        # Connection mode toggle
        self.connection_mode_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            toolbar_frame,
            text="Connection Mode",
            variable=self.connection_mode_var,
            command=self.toggle_connection_mode
        ).grid(row=0, column=15, padx=5)
        
        # Grid toggle
        self.grid_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            toolbar_frame,
            text="Show Grid",
            variable=self.grid_var,
            command=self.toggle_grid
        ).grid(row=0, column=16, padx=5)
    
    def create_canvas_area(self, parent):
        """Create the main canvas area"""
        # Canvas frame with border
        canvas_frame = ttk.Frame(parent, relief=tk.SUNKEN, borderwidth=2)
        canvas_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Create canvas
        self.canvas = tk.Canvas(
            canvas_frame,
            bg='white',
            width=1200,
            height=800,
            scrollregion=(0, 0, 2400, 1600)
        )
        self.canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbars
        h_scroll = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        h_scroll.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        v_scroll = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        v_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        self.canvas.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)
        
        # Configure grid weights
        canvas_frame.grid_columnconfigure(0, weight=1)
        canvas_frame.grid_rowconfigure(0, weight=1)
        
        # Bind canvas events
        self.canvas.bind("<Button-1>", self.on_canvas_click)
        self.canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        self.canvas.bind("<Button-3>", self.on_canvas_right_click)
        self.canvas.bind("<Double-1>", self.on_canvas_double_click)
        
        # Mouse wheel for zoom
        self.canvas.bind("<Control-MouseWheel>", self.on_canvas_zoom)  # Windows
        self.canvas.bind("<Control-Button-4>", self.on_canvas_zoom)  # Linux up
        self.canvas.bind("<Control-Button-5>", self.on_canvas_zoom)  # Linux down
        
        # Regular scroll without Ctrl
        self.canvas.bind("<MouseWheel>", self.on_canvas_scroll)
        self.canvas.bind("<Button-4>", self.on_canvas_scroll)
        self.canvas.bind("<Button-5>", self.on_canvas_scroll)
        
        # Middle button for panning
        self.canvas.bind("<Button-2>", self.on_pan_start)
        self.canvas.bind("<B2-Motion>", self.on_pan_motion)
        
        # Keyboard shortcuts
        self.bind("<Delete>", lambda e: self.delete_selected())
        self.bind("<Control-a>", lambda e: self.select_all())
        self.bind("<Control-plus>", lambda e: self.zoom(1.25))
        self.bind("<Control-minus>", lambda e: self.zoom(0.8))
        self.bind("<Control-0>", lambda e: self.fit_to_window())
        self.bind("<Escape>", lambda e: self.on_close())
        
        # Draw initial grid
        self.draw_grid()
    
    def create_status_bar(self, parent):
        """Create the status bar"""
        status_frame = ttk.Frame(parent)
        status_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # Status label
        self.status_label = ttk.Label(status_frame, text="Ready", relief=tk.SUNKEN)
        self.status_label.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 2))
        
        # Coordinates label
        self.coord_label = ttk.Label(status_frame, text="X: 0, Y: 0", relief=tk.SUNKEN, width=20)
        self.coord_label.grid(row=0, column=1, sticky=tk.E, padx=2)
        
        # Element count label
        self.element_count_label = ttk.Label(status_frame, text="Elements: 0", relief=tk.SUNKEN, width=15)
        self.element_count_label.grid(row=0, column=2, sticky=tk.E, padx=2)
        
        # Connection count label
        self.connection_count_label = ttk.Label(status_frame, text="Connections: 0", relief=tk.SUNKEN, width=15)
        self.connection_count_label.grid(row=0, column=3, sticky=tk.E)
        
        status_frame.grid_columnconfigure(0, weight=1)
        
        # Bind mouse motion for coordinate display
        self.canvas.bind("<Motion>", self.on_canvas_motion)
    
    def draw_grid(self):
        """Draw grid lines on canvas"""
        self.canvas.delete("grid")
        
        if not self.grid_visible:
            return
        
        width = 2400
        height = 1600
        
        # Apply zoom to grid
        grid_spacing = self.grid_size * self.zoom_level
        
        # Draw vertical lines
        for x in range(0, int(width * self.zoom_level), int(grid_spacing)):
            line_width = 1
            color = "#E0E0E0"
            
            # Every 5th line is thicker
            if (x / grid_spacing) % 5 == 0:
                line_width = 2
                color = "#D0D0D0"
            
            # Every 10th line is even thicker
            if (x / grid_spacing) % 10 == 0:
                line_width = 2
                color = "#C0C0C0"
            
            self.canvas.create_line(
                x, 0, x, height * self.zoom_level,
                fill=color,
                width=line_width,
                tags="grid"
            )
        
        # Draw horizontal lines
        for y in range(0, int(height * self.zoom_level), int(grid_spacing)):
            line_width = 1
            color = "#E0E0E0"
            
            # Every 5th line is thicker
            if (y / grid_spacing) % 5 == 0:
                line_width = 2
                color = "#D0D0D0"
            
            # Every 10th line is even thicker
            if (y / grid_spacing) % 10 == 0:
                line_width = 2
                color = "#C0C0C0"
            
            self.canvas.create_line(
                0, y, width * self.zoom_level, y,
                fill=color,
                width=line_width,
                tags="grid"
            )
        
        # Keep grid behind other elements
        self.canvas.tag_lower("grid")
    
    def generate_sld_from_blocks(self):
        """Generate SLD diagram from blocks"""
        if not self.blocks:
            self.status_label.configure(text="No blocks available")
            return
        
        # Clear canvas
        self.canvas.delete("all")
        self.draw_grid()
        
        # For now, create sample elements using the new symbol library
        self.status_label.configure(text="Drawing ANSI symbols - Parser in Card #4")
        
        # Draw sample PV block using ANSI symbols
        pv_result = ANSISymbols.draw_symbol(
            self.canvas,
            symbol_type='pv_array',
            x=100,
            y=200,
            width=150,
            height=100,
            label="Block A\n2.5 MW",
            element_id="Block_A",
            show_ports=True  # Show connection points for testing
        )
        
        # Store element reference
        self.sld_elements["Block_A"] = pv_result
        
        # Draw sample inverter
        inv_result = ANSISymbols.draw_symbol(
            self.canvas,
            symbol_type='inverter',
            x=800,
            y=200,
            width=150,
            height=100,
            label="INV-01\n2500 kVA",
            element_id="INV_01",
            show_ports=True
        )
        
        self.sld_elements["INV_01"] = inv_result
        
        # Draw sample combiner box
        cb_result = ANSISymbols.draw_symbol(
            self.canvas,
            symbol_type='combiner',
            x=450,
            y=220,
            width=80,
            height=80,
            label="CB-01",
            element_id="CB_01",
            show_ports=True
        )
        
        self.sld_elements["CB_01"] = cb_result
        
        # Draw a test connection line between PV and combiner
        if "Block_A" in self.sld_elements and "CB_01" in self.sld_elements:
            # Get connection points
            pv_pos_point = self.sld_elements["Block_A"]['connection_points']['dc_positive']
            cb_input_point = self.sld_elements["CB_01"]['connection_points']['input_1']
            
            # Draw connection line (will be improved in Card #8)
            self.canvas.create_line(
                pv_pos_point[0], pv_pos_point[1],
                cb_input_point[0], cb_input_point[1],
                fill='red',
                width=3,
                tags=('connection', 'dc_positive'),
                smooth=False
            )
        
        self.update_status_counts()
        self.status_label.configure(text="ANSI symbols drawn - Ready")

    def create_symbol_palette(self, parent):
        """Create a palette of available symbols"""
        palette_frame = ttk.LabelFrame(parent, text="Symbols", padding="5")
        palette_frame.grid(row=0, column=1, sticky=(tk.N, tk.S), padx=(5, 0))
        
        # Create small canvas for each symbol type
        symbols = ANSISymbols.get_available_symbols()
        
        for i, symbol_type in enumerate(symbols):
            # Create label
            ttk.Label(palette_frame, text=symbol_type.replace('_', ' ').title()).grid(
                row=i*2, column=0, pady=(5, 0)
            )
            
            # Create mini canvas for preview
            mini_canvas = tk.Canvas(
                palette_frame,
                width=80,
                height=60,
                bg='white',
                relief=tk.RAISED,
                borderwidth=1
            )
            mini_canvas.grid(row=i*2+1, column=0, padx=5, pady=2)
            
            # Draw mini symbol
            ANSISymbols.draw_symbol(
                mini_canvas,
                symbol_type=symbol_type,
                x=5,
                y=5,
                width=70,
                height=50,
                label=""
            )
            
            # Make draggable (implement in future cards)
            mini_canvas.bind("<Button-1>", lambda e, st=symbol_type: self.start_symbol_drag(st))
    
    def update_status_counts(self):
        """Update the status bar counts"""
        # Count elements
        elements = self.canvas.find_withtag("element")
        self.element_count_label.configure(text=f"Elements: {len(elements)}")
        
        # Count connections
        connections = self.canvas.find_withtag("connection")
        self.connection_count_label.configure(text=f"Connections: {len(connections)}")
    
    # Event handlers
    def on_canvas_click(self, event):
        """Handle canvas click"""
        if self.connection_mode_var.get():
            # Connection mode - will be implemented in Card #8
            pass
        else:
            # Selection/drag mode
            canvas_x = self.canvas.canvasx(event.x)
            canvas_y = self.canvas.canvasy(event.y)
            
            # Find clicked item
            item = self.canvas.find_closest(canvas_x, canvas_y)
            if item:
                tags = self.canvas.gettags(item)
                if "element" in tags:
                    # Start dragging
                    self.dragging = True
                    self.drag_data["x"] = canvas_x
                    self.drag_data["y"] = canvas_y
                    self.drag_data["item"] = item[0]
                    
                    # Highlight selection
                    self.canvas.itemconfig(item[0], width=3)
                    self.selected_element = item[0]
    
    def on_canvas_drag(self, event):
        """Handle canvas drag"""
        if self.dragging and self.drag_data["item"]:
            canvas_x = self.canvas.canvasx(event.x)
            canvas_y = self.canvas.canvasy(event.y)
            
            # Calculate movement
            dx = canvas_x - self.drag_data["x"]
            dy = canvas_y - self.drag_data["y"]
            
            # Move the element and its label
            item_tags = self.canvas.gettags(self.drag_data["item"])
            for tag in item_tags:
                if tag not in ["element", "current"]:
                    # Move element
                    self.canvas.move(tag, dx, dy)
                    # Move associated label
                    self.canvas.move(f"{tag}_label", dx, dy)
            
            # Update drag position
            self.drag_data["x"] = canvas_x
            self.drag_data["y"] = canvas_y
    
    def on_canvas_release(self, event):
        """Handle canvas mouse release"""
        if self.dragging:
            # Snap to grid if enabled
            if self.grid_visible and self.drag_data["item"]:
                coords = self.canvas.bbox(self.drag_data["item"])
                if coords:
                    # Calculate snap position
                    x1, y1, x2, y2 = coords
                    grid_snap = self.grid_size * self.zoom_level
                    
                    snapped_x = round(x1 / grid_snap) * grid_snap
                    snapped_y = round(y1 / grid_snap) * grid_snap
                    
                    # Calculate offset
                    dx = snapped_x - x1
                    dy = snapped_y - y1
                    
                    # Apply snap
                    item_tags = self.canvas.gettags(self.drag_data["item"])
                    for tag in item_tags:
                        if tag not in ["element", "current"]:
                            self.canvas.move(tag, dx, dy)
                            self.canvas.move(f"{tag}_label", dx, dy)
            
            self.dragging = False
            self.drag_data = {"x": 0, "y": 0, "item": None, "element_id": None}
    
    def on_canvas_right_click(self, event):
        """Handle right-click for context menu"""
        # This will be implemented in Card #10
        pass
    
    def on_canvas_double_click(self, event):
        """Handle double-click for editing"""
        # This will be implemented in Card #11
        pass
    
    def on_canvas_motion(self, event):
        """Update coordinate display on mouse motion"""
        canvas_x = int(self.canvas.canvasx(event.x) / self.zoom_level)
        canvas_y = int(self.canvas.canvasy(event.y) / self.zoom_level)
        self.coord_label.configure(text=f"X: {canvas_x}, Y: {canvas_y}")
    
    def on_canvas_zoom(self, event):
        """Handle zoom with Ctrl+scroll"""
        if event.delta > 0 or event.num == 4:
            self.zoom(1.1)
        else:
            self.zoom(0.9)
    
    def on_canvas_scroll(self, event):
        """Handle regular scrolling"""
        if event.delta:
            # Windows
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        else:
            # Linux
            if event.num == 4:
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.canvas.yview_scroll(1, "units")
    
    def on_pan_start(self, event):
        """Start panning with middle mouse"""
        self.canvas.scan_mark(event.x, event.y)
    
    def on_pan_motion(self, event):
        """Pan with middle mouse"""
        self.canvas.scan_dragto(event.x, event.y, gain=1)
    
    # Toolbar actions
    def auto_layout(self):
        """Auto-layout elements"""
        if self.sld_diagram:
            self.sld_diagram.auto_layout()
            # Refresh canvas - will be implemented in Card #5
        self.status_label.configure(text="Auto-layout will be implemented in Card #7")
    
    def align_selected(self, direction):
        """Align selected elements"""
        # This will be implemented with multi-select in Card #6
        self.status_label.configure(text=f"Align {direction} will be implemented in Card #6")
    
    def zoom(self, factor):
        """Zoom the canvas"""
        # Update zoom level
        self.zoom_level *= factor
        self.zoom_level = max(0.1, min(5.0, self.zoom_level))  # Limit zoom range
        
        # Update zoom label
        self.zoom_label.configure(text=f"{int(self.zoom_level * 100)}%")
        
        # Scale all canvas items
        self.canvas.scale("all", 0, 0, factor, factor)
        
        # Update scroll region
        bbox = self.canvas.bbox("all")
        if bbox:
            self.canvas.configure(scrollregion=bbox)
        
        # Redraw grid at new scale
        self.draw_grid()
        
        self.status_label.configure(text=f"Zoom: {int(self.zoom_level * 100)}%")
    
    def fit_to_window(self):
        """Fit diagram to window"""
        bbox = self.canvas.bbox("all")
        if not bbox:
            return
        
        # Get canvas dimensions
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        
        # Calculate required scale
        diagram_width = bbox[2] - bbox[0]
        diagram_height = bbox[3] - bbox[1]
        
        if diagram_width > 0 and diagram_height > 0:
            scale_x = canvas_width / diagram_width * 0.9
            scale_y = canvas_height / diagram_height * 0.9
            scale = min(scale_x, scale_y)
            
            # Reset zoom and apply new scale
            reset_factor = 1.0 / self.zoom_level
            self.canvas.scale("all", 0, 0, reset_factor, reset_factor)
            self.zoom_level = 1.0
            
            # Apply fit scale
            self.zoom(scale)
    
    def toggle_grid(self):
        """Toggle grid visibility"""
        self.grid_visible = self.grid_var.get()
        if self.grid_visible:
            self.draw_grid()
        else:
            self.canvas.delete("grid")
    
    def toggle_connection_mode(self):
        """Toggle connection drawing mode"""
        if self.connection_mode_var.get():
            self.status_label.configure(text="Connection mode - Click elements to connect")
            self.canvas.configure(cursor="cross")
        else:
            self.status_label.configure(text="Ready")
            self.canvas.configure(cursor="")
            # Cancel any in-progress connection
            if self.temp_connection_line:
                self.canvas.delete(self.temp_connection_line)
                self.temp_connection_line = None
    
    def clear_all(self):
        """Clear all elements from canvas"""
        response = messagebox.askyesno("Clear All", 
                                       "Are you sure you want to clear all elements?")
        if response:
            self.canvas.delete("all")
            self.draw_grid()
            self.sld_diagram = None
            self.sld_elements.clear()
            self.update_status_counts()
            self.status_label.configure(text="Canvas cleared")
    
    def select_all(self):
        """Select all elements"""
        elements = self.canvas.find_withtag("element")
        for elem in elements:
            self.canvas.itemconfig(elem, width=3)
        self.selected_items = list(elements)
        self.status_label.configure(text=f"Selected {len(elements)} elements")
    
    def delete_selected(self):
        """Delete selected elements"""
        if self.selected_element:
            # Get tags of selected element
            tags = self.canvas.gettags(self.selected_element)
            for tag in tags:
                if tag not in ["element", "current"]:
                    # Delete element and its label
                    self.canvas.delete(tag)
                    self.canvas.delete(f"{tag}_label")
            
            self.selected_element = None
            self.update_status_counts()
            self.status_label.configure(text="Deleted selected element")
    
    def validate_connections(self):
        """Validate all connections"""
        # This will be implemented with connection system in Card #8
        messagebox.showinfo("Validate", 
                          "Connection validation will be implemented in Card #8")
    
    # Export functions
    def export_png(self):
        """Export diagram as PNG"""
        # This will be implemented in Card #12
        messagebox.showinfo("Export PNG", 
                          "PNG export will be implemented in Card #12")
    
    def export_pdf(self):
        """Export diagram as PDF"""
        # This will be implemented in Card #13
        messagebox.showinfo("Export PDF", 
                          "PDF export will be implemented in Card #13")
    
    # Help functions
    def show_shortcuts(self):
        """Show keyboard shortcuts"""
        shortcuts = """
        Keyboard Shortcuts:
        
        Ctrl+A      - Select All
        Delete      - Delete Selected
        Ctrl++      - Zoom In
        Ctrl+-      - Zoom Out
        Ctrl+0      - Fit to Window
        Ctrl+P      - Export PNG
        Ctrl+D      - Export PDF
        
        Middle Mouse - Pan
        Ctrl+Scroll  - Zoom
        """
        messagebox.showinfo("Keyboard Shortcuts", shortcuts)
    
    def show_about(self):
        """Show about dialog"""
        about_text = """
        Single Line Diagram Editor
        Version 1.0
        
        Part of Solar eBOS BOM Generator
        
        Create professional electrical single line
        diagrams for solar PV systems.
        """
        messagebox.showinfo("About SLD Editor", about_text)
    
    def on_close(self):
        """Handle window closing"""
        # Save diagram to project if modified
        if self.sld_diagram and self.project:
            # This will be implemented with save functionality
            pass
        
        self.destroy()