import tkinter as tk
from tkinter import ttk, messagebox
import json
from pathlib import Path
from typing import Optional, Callable
from ..models.tracker import TrackerTemplate, ModuleOrientation
from ..models.module import ModuleSpec, ModuleType

class TrackerTemplateCreator(ttk.Frame):
    def __init__(self, parent, module_spec: Optional[ModuleSpec] = None, 
             on_template_saved: Optional[Callable[[TrackerTemplate], None]] = None):
        super().__init__(parent)
        self.parent = parent
        self.on_template_saved = on_template_saved
        self._module_spec = None
        self.templates = self.load_templates()
        self.setup_ui()  # Creates current_module_label
        self.module_spec = module_spec  # Now safe to call

    @property 
    def module_spec(self):
        return self._module_spec
    
    @module_spec.setter
    def module_spec(self, value):
        self._module_spec = value
        self.update_module_display()
        # Force an immediate preview update with the new module dimensions
        self.canvas.after_idle(self.update_preview)
        
    def setup_ui(self):
        """Create and arrange UI components"""
        # Main container with padding
        main_container = ttk.Frame(self, padding="10")
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Left column
        left_column = ttk.Frame(main_container)
        left_column.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.N, tk.S))
        
        # Template List section in left column
        template_frame = ttk.LabelFrame(left_column, text="Saved Templates", padding="2")
        template_frame.grid(row=0, column=0, padx=2, pady=5, sticky=(tk.W, tk.E, tk.N))
        
        self.template_listbox = tk.Listbox(template_frame, width=30, height=15)
        self.template_listbox.grid(row=0, column=0, padx=2, pady=2)
        self.template_listbox.bind('<<ListboxSelect>>', self.on_template_select)
        
        template_buttons = ttk.Frame(template_frame)
        template_buttons.grid(row=1, column=0, padx=2, pady=2)
        
        ttk.Button(template_buttons, text="Load", command=self.load_template).grid(row=0, column=0, padx=2)
        ttk.Button(template_buttons, text="Delete", command=self.delete_template).grid(row=0, column=1, padx=2)
        
        # Editor section in left column
        editor_frame = ttk.LabelFrame(left_column, text="Template Editor", padding="5")
        editor_frame.grid(row=1, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Current Module
        ttk.Label(editor_frame, text="Current Module:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.current_module_label = ttk.Label(editor_frame, text="No module selected")
        self.current_module_label.grid(row=0, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))

        # Template Name
        ttk.Label(editor_frame, text="Template Name:").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.name_var = tk.StringVar()
        ttk.Entry(editor_frame, textvariable=self.name_var).grid(row=1, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        # Module Orientation
        ttk.Label(editor_frame, text="Module Orientation:").grid(row=2, column=0, padx=5, pady=2, sticky=tk.W)
        self.orientation_var = tk.StringVar(value=ModuleOrientation.PORTRAIT.value)
        orientation_combo = ttk.Combobox(editor_frame, textvariable=self.orientation_var)
        orientation_combo['values'] = [o.value for o in ModuleOrientation]
        orientation_combo.grid(row=2, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        # Modules per String
        ttk.Label(editor_frame, text="Modules per String:").grid(row=3, column=0, padx=5, pady=2, sticky=tk.W)
        self.modules_string_var = tk.StringVar(value="28")
        ttk.Spinbox(editor_frame, from_=1, to=100, textvariable=self.modules_string_var, 
            increment=1, validate='all', validatecommand=(self.register(lambda val: val.isdigit() or val == ""), '%P')
            ).grid(row=3, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        # Strings per Tracker
        ttk.Label(editor_frame, text="Strings per Tracker:").grid(row=4, column=0, padx=5, pady=2, sticky=tk.W)
        self.strings_tracker_var = tk.StringVar(value="2")
        ttk.Spinbox(editor_frame, from_=1, to=10, textvariable=self.strings_tracker_var,
            increment=1, validate='all', validatecommand=(self.register(lambda val: val.isdigit() or val == ""), '%P')
            ).grid(row=4, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        # Module Spacing
        ttk.Label(editor_frame, text="Module Spacing (m):").grid(row=5, column=0, padx=5, pady=2, sticky=tk.W)
        self.spacing_var = tk.StringVar(value="0.01")
        ttk.Entry(editor_frame, textvariable=self.spacing_var).grid(row=5, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        # Motor Gap
        ttk.Label(editor_frame, text="Motor Gap (m):").grid(row=6, column=0, padx=5, pady=2, sticky=tk.W)
        self.motor_gap_var = tk.StringVar(value="1.0")
        ttk.Entry(editor_frame, textvariable=self.motor_gap_var).grid(row=6, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))

        # Total Modules Display
        ttk.Label(editor_frame, text="Modules per Tracker:").grid(row=7, column=0, padx=5, pady=2, sticky=tk.W)
        self.total_modules_label = ttk.Label(editor_frame, text="--")
        self.total_modules_label.grid(row=7, column=1, padx=5, pady=2, sticky=tk.W)
        
        # Calculated Dimensions Display
        dims_frame = ttk.LabelFrame(editor_frame, text="Tracker Dimensions", padding="5")
        dims_frame.grid(row=8, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        self.length_label = ttk.Label(dims_frame, text="Length: --")
        self.length_label.grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        
        self.width_label = ttk.Label(dims_frame, text="Width: --")
        self.width_label.grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)

        # Update total modules calculation when inputs change
        def update_total_modules(*args):
            try:
                modules = int(self.modules_string_var.get()) * int(self.strings_tracker_var.get())
                self.total_modules_label.config(text=str(modules))
            except ValueError:
                self.total_modules_label.config(text="--")

        self.modules_string_var.trace('w', update_total_modules)
        self.strings_tracker_var.trace('w', update_total_modules)
        
        # Save Button
        ttk.Button(editor_frame, text="Save Template", command=self.save_template).grid(row=9, column=0, columnspan=2, pady=10)
        
        # Preview Canvas - Right column
        preview_frame = ttk.LabelFrame(main_container, text="Preview", padding="5")
        preview_frame.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.N, tk.S))

        self.canvas = tk.Canvas(preview_frame, width=300, height=600, bg='white')
        self.canvas.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.N, tk.S))
        
        # Bind events for real-time preview updates
        for var in [self.modules_string_var, self.strings_tracker_var, 
                self.spacing_var, self.motor_gap_var, self.orientation_var]:
            var.trace('w', lambda *args: self.update_preview())

    def update_module_display(self):
        """Update the current module display and related calculations"""
        if self._module_spec:
            self.current_module_label.config(
                text=f"{self._module_spec.manufacturer} {self._module_spec.model} ({self._module_spec.wattage}W)"
            )
            # Trigger preview update with new module dimensions
            self.update_preview()
        else:
            self.current_module_label.config(text="No module selected")
            # Clear preview when no module is selected
            self.canvas.delete("all")
            self.length_label.config(text="Length: --")
            self.width_label.config(text="Width: --")
            self.total_modules_label.config(text="--")
            
    def load_templates(self) -> dict:
        """Load saved templates from JSON file"""
        template_path = Path('data/tracker_templates.json')
        try:
            template_path.parent.mkdir(parents=True, exist_ok=True)
            if not template_path.exists():
                with open(template_path, 'w') as f:
                    json.dump({}, f)
                return {}
            with open(template_path, 'r') as f:
                content = f.read().strip()
                return json.loads(content) if content else {}
        except (json.JSONDecodeError, IOError) as e:
            messagebox.showerror("Error", f"Failed to load templates file: {str(e)}")
            return {}
            
    def save_templates(self):
        """Save templates to JSON file"""
        template_path = Path('data/tracker_templates.json')
        template_path.parent.mkdir(exist_ok=True)
        with open(template_path, 'w') as f:
            json.dump(self.templates, f, indent=2)
            
    def update_template_list(self):
        """Update the template listbox"""
        self.template_listbox.delete(0, tk.END)
        for name in self.templates.keys():
            self.template_listbox.insert(tk.END, name)
            
    def on_template_select(self, event=None):
        """Handle template selection event"""
        selection = self.template_listbox.curselection()
        if selection:
            self.load_template()
            
    def create_template(self) -> Optional[TrackerTemplate]:
        """Create a TrackerTemplate from current UI values"""
        try:
            template = TrackerTemplate(
                template_name=self.name_var.get(),
                module_spec=self.module_spec or ModuleSpec(
                    manufacturer="Sample",
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
                ),
                module_orientation=ModuleOrientation(self.orientation_var.get()),
                modules_per_string=int(self.modules_string_var.get()),
                strings_per_tracker=int(self.strings_tracker_var.get()),
                module_spacing_m=float(self.spacing_var.get()),
                motor_gap_m=float(self.motor_gap_var.get())
            )
            template.validate()
            return template
        except (ValueError, TypeError) as e:
            messagebox.showerror("Error", str(e))
            return None
            
    def save_template(self):
        """Save current template"""
        template = self.create_template()
        if template:
            name = template.template_name
            if not name:
                messagebox.showerror("Error", "Template name is required")
                return
                
            if name in self.templates:
                if not messagebox.askyesno("Confirm", f"Template '{name}' already exists. Overwrite?"):
                    return
                    
            # Save template data
            self.templates[name] = {
                "module_orientation": template.module_orientation.value,
                "modules_per_string": template.modules_per_string,
                "strings_per_tracker": template.strings_per_tracker,
                "module_spacing_m": template.module_spacing_m,
                "motor_gap_m": template.motor_gap_m
            }
            
            self.save_templates()
            self.update_template_list()
            
            if self.on_template_saved:
                self.on_template_saved(template)
                
            messagebox.showinfo("Success", f"Template '{name}' saved successfully")
            
    def load_template(self):
        """Load selected template"""
        selection = self.template_listbox.curselection()
        if not selection:
            return
            
        name = self.template_listbox.get(selection[0])
        template_data = self.templates[name]
        
        self.name_var.set(name)
        self.orientation_var.set(template_data["module_orientation"])
        self.modules_string_var.set(str(template_data["modules_per_string"]))
        self.strings_tracker_var.set(str(template_data["strings_per_tracker"]))
        self.spacing_var.set(str(template_data["module_spacing_m"]))
        self.motor_gap_var.set(str(template_data["motor_gap_m"]))
        
        self.update_preview()
        
    def delete_template(self):
        """Delete selected template"""
        selection = self.template_listbox.curselection()
        if not selection:
            return
            
        name = self.template_listbox.get(selection[0])
        if messagebox.askyesno("Confirm", f"Delete template '{name}'?"):
            del self.templates[name]
            self.save_templates()
            self.update_template_list()
            
    def update_preview(self):
        """Update the preview canvas with current template layout"""
        template = self.create_template()
        if not template:
            return
            
        self.canvas.delete("all")
        
        # Calculate module dimensions - flipped for vertical orientation
        if template.module_orientation == ModuleOrientation.PORTRAIT:
            module_height = template.module_spec.width_mm / 1000
            module_width = template.module_spec.length_mm / 1000
        else:
            module_height = template.module_spec.length_mm / 1000
            module_width = template.module_spec.width_mm / 1000

        # Rest of the code remains same as previous version
        total_strings_above_motor = template.strings_per_tracker - 1
        modules_above_motor = total_strings_above_motor * template.modules_per_string
        modules_below_motor = template.modules_per_string
        
        total_height = ((modules_above_motor + modules_below_motor) * module_height) + \
                        ((modules_above_motor + modules_below_motor - 1) * template.module_spacing_m) + \
                        template.motor_gap_m

        scale = min(280 / module_width, 580 / total_height)
        x_center = (300 - module_width * scale) / 2
        
        # Draw torque tube
        self.canvas.create_line(
            x_center + module_width * scale / 2, 10,
            x_center + module_width * scale / 2, 10 + total_height * scale,
            width=3, fill='gray'
        )
        
        # Draw modules above motor
        y_pos = 10
        for i in range(modules_above_motor):
            self.canvas.create_rectangle(
                x_center, y_pos,
                x_center + module_width * scale, y_pos + module_height * scale,
                fill='lightblue', outline='blue'
            )
            y_pos += (module_height + template.module_spacing_m) * scale

        # Draw motor
        motor_y = y_pos
        y_pos += template.motor_gap_m * scale
        self.canvas.create_oval(
            x_center + module_width * scale / 2 - 5, motor_y - 5,
            x_center + module_width * scale / 2 + 5, motor_y + 5,
            fill='red'
        )

        # Draw modules below motor
        for i in range(modules_below_motor):
            self.canvas.create_rectangle(
                x_center, y_pos,
                x_center + module_width * scale, y_pos + module_height * scale,
                fill='lightblue', outline='blue'
            )
            y_pos += (module_height + template.module_spacing_m) * scale

        # Update dimension labels
        self.length_label.config(text=f"Length: {module_width:.2f}m")
        self.width_label.config(text=f"Height: {total_height:.2f}m")