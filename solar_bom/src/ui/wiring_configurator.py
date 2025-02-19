import tkinter as tk
from tkinter import ttk
from typing import Optional
from ..models.block import BlockConfig, WiringType
from ..models.module import ModuleOrientation

class WiringConfigurator(tk.Toplevel):
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
        self.wiring_type_var = tk.StringVar(value=WiringType.HOMERUN.value)
        wiring_type_combo = ttk.Combobox(controls_frame, textvariable=self.wiring_type_var, state='readonly')
        wiring_type_combo['values'] = [t.value for t in WiringType]
        wiring_type_combo.grid(row=0, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        wiring_type_combo.bind('<<ComboboxSelected>>', self.on_wiring_type_change)
        
        # Cable Specifications
        cable_frame = ttk.LabelFrame(controls_frame, text="Cable Specifications", padding="5")
        cable_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=10, sticky=(tk.W, tk.E))

        # String Cable Size
        ttk.Label(cable_frame, text="String Cable Size (mm²):").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.string_cable_size_var = tk.StringVar(value="4")
        ttk.Entry(cable_frame, textvariable=self.string_cable_size_var, width=10).grid(row=0, column=1, padx=5, pady=2)

        # Wire Harness Size
        self.harness_frame = ttk.Frame(cable_frame)
        self.harness_frame.grid(row=1, column=0, columnspan=2, padx=0, pady=5, sticky=(tk.W, tk.E))
        ttk.Label(self.harness_frame, text="Harness Cable Size (mm²):").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.harness_cable_size_var = tk.StringVar(value="35")
        ttk.Entry(self.harness_frame, textvariable=self.harness_cable_size_var, width=10).grid(row=0, column=1, padx=5, pady=2)
        
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
        self.draw_block()

    def draw_block(self):
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
        
        # Draw inverter/combiner
        device_x = 20 + self.pan_x + self.block.device_x * scale
        device_y = 20 + self.pan_y + self.block.device_y * scale
        device_size = 0.91 * scale  # 3ft = 0.91m
        self.canvas.create_rectangle(
            device_x, device_y,
            device_x + device_size,
            device_y + device_size,
            fill='red', outline='darkred'
        )

    def on_wiring_type_change(self, event=None):
        """Handle wiring type selection change"""
        self.update_ui_for_wiring_type()
        
    def on_canvas_resize(self, event):
        """Handle canvas resize event"""
        if event.width > 1 and event.height > 1:  # Ensure valid dimensions
            self.draw_block()
        
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
            self.draw_block()

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
        self.draw_block()

    def end_pan(self, event):
        """End canvas panning"""
        self.panning = False
        self.canvas.config(cursor="")  # Reset cursor