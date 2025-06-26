import tkinter as tk
from tkinter import ttk, messagebox
import json
from pathlib import Path
from typing import Optional, Callable
from ..models.tracker import TrackerTemplate, ModuleOrientation
from ..models.module import ModuleSpec, ModuleType

class TrackerTemplateCreator(ttk.Frame):
    def __init__(self, parent, module_spec: Optional[ModuleSpec] = None, 
             on_template_saved: Optional[Callable[[TrackerTemplate], None]] = None,
             on_template_deleted: Optional[Callable[[str], None]] = None):
        super().__init__(parent)
        self.parent = parent
        self.on_template_saved = on_template_saved
        self.on_template_deleted = on_template_deleted
        self._module_spec = None
        self.templates = self.load_templates()
        self.setup_ui()  # Creates current_module_label
        self.module_spec = module_spec  # Now safe to call
        self.update_template_list()

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
        
        # Create Treeview for hierarchical template display
        self.template_tree = ttk.Treeview(template_frame, height=15)
        self.template_tree.grid(row=0, column=0, padx=2, pady=2, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure tree columns
        self.template_tree.heading('#0', text='Templates')
        self.template_tree.column('#0', width=300)

        # Add scrollbar for tree
        tree_scrollbar = ttk.Scrollbar(template_frame, orient=tk.VERTICAL, command=self.template_tree.yview)
        tree_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.template_tree.configure(yscrollcommand=tree_scrollbar.set)

        # Bind selection event
        self.template_tree.bind('<<TreeviewSelect>>', self.on_template_select)
        
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
        ttk.Spinbox(editor_frame, from_=1, to=20, textvariable=self.strings_tracker_var,
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

        # Motor Position
        ttk.Label(editor_frame, text="Motor After String:").grid(row=7, column=0, padx=5, pady=2, sticky=tk.W)
        self.motor_position_var = tk.StringVar(value="2")  # Default to middle-ish
        self.motor_position_spinbox = ttk.Spinbox(editor_frame, from_=0, to=1, textvariable=self.motor_position_var,
            increment=1, validate='all', validatecommand=(self.register(lambda val: val.isdigit() or val == ""), '%P'))
        self.motor_position_spinbox.grid(row=7, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))

        # Total Modules Display
        ttk.Label(editor_frame, text="Modules per Tracker:").grid(row=8, column=0, padx=5, pady=2, sticky=tk.W)
        self.total_modules_label = ttk.Label(editor_frame, text="--")
        try:
            modules = int(self.modules_string_var.get()) * int(self.strings_tracker_var.get())
            self.total_modules_label.config(text=str(modules))
        except ValueError:
            self.total_modules_label.config(text="--")
        self.total_modules_label.grid(row=8, column=1, padx=5, pady=2, sticky=tk.W)
        
        # Calculated Dimensions Display
        dims_frame = ttk.LabelFrame(editor_frame, text="Tracker Dimensions", padding="5")
        dims_frame.grid(row=9, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
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
        ttk.Button(editor_frame, text="Save Template", command=self.save_template).grid(row=10, column=0, columnspan=2, pady=10)
        
        # Preview Canvas - Right column
        preview_frame = ttk.LabelFrame(main_container, text="Preview", padding="5")
        preview_frame.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.N, tk.S))

        self.canvas = tk.Canvas(preview_frame, width=300, height=600, bg='white')
        self.canvas.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.N, tk.S))
        
        # Bind events for real-time preview updates
        for var in [self.modules_string_var, self.strings_tracker_var, 
                self.spacing_var, self.motor_gap_var, self.orientation_var, self.motor_position_var]:
            var.trace('w', lambda *args: self.update_preview())

        # Add special trace for strings_tracker_var to update motor position range
        self.strings_tracker_var.trace('w', self.update_motor_position_range)

    def update_motor_position_range(self, *args):
        """Update the motor position spinbox range based on strings per tracker"""
        try:
            strings_count = int(self.strings_tracker_var.get())
            # Update the spinbox range: 0 to strings_count (inclusive)
            self.motor_position_spinbox.config(to=strings_count)
            
            # Set sensible default: middle for even, middle rounded up for odd
            current_val = int(self.motor_position_var.get())
            if current_val > strings_count:
                if strings_count % 2 == 0:
                    default_pos = strings_count // 2
                else:
                    default_pos = (strings_count + 1) // 2
                self.motor_position_var.set(str(default_pos))
        except ValueError:
            pass

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
                data = json.loads(content) if content else {}
                
            # Handle both old flat structure and new hierarchical structure
            templates = {}
            
            if data:
                # Check if this is the new hierarchical format
                first_value = next(iter(data.values()))
                if isinstance(first_value, dict) and not any(key in first_value for key in ['module_orientation', 'modules_per_string']):
                    # New hierarchical format: Manufacturer -> Template Name -> template_data
                    for manufacturer, template_group in data.items():
                        for template_name, template_data in template_group.items():
                            # Use manufacturer prefix to make template names unique
                            unique_name = f"{manufacturer} - {template_name}"
                            templates[unique_name] = template_data
                else:
                    # Old flat format
                    templates = data
                    
            return templates
        except (json.JSONDecodeError, IOError) as e:
            messagebox.showerror("Error", f"Failed to load templates file: {str(e)}")
            return {}
            
    def save_templates(self):
        """Save templates to JSON file in hierarchical format"""
        template_path = Path('data/tracker_templates.json')
        template_path.parent.mkdir(exist_ok=True)
        
        # Organize templates by manufacturer
        hierarchical_data = {}
        
        for template_key, template_data in self.templates.items():
            # Extract manufacturer from module_spec
            module_spec = template_data.get('module_spec', {})
            manufacturer = module_spec.get('manufacturer', 'Unknown')
            
            # Extract template name (remove manufacturer prefix if present)
            if ' - ' in template_key and template_key.startswith(manufacturer):
                template_name = template_key.split(' - ', 1)[1]
            else:
                template_name = template_key
            
            if manufacturer not in hierarchical_data:
                hierarchical_data[manufacturer] = {}
                
            hierarchical_data[manufacturer][template_name] = template_data
        
        with open(template_path, 'w') as f:
            json.dump(hierarchical_data, f, indent=2)
            
    def update_template_list(self):
        """Update the template tree view"""
        # Clear existing items
        for item in self.template_tree.get_children():
            self.template_tree.delete(item)
        
        # Group templates by manufacturer
        manufacturers = {}
        for template_key, template_data in self.templates.items():
            # Extract manufacturer from module_spec
            module_spec = template_data.get('module_spec', {})
            manufacturer = module_spec.get('manufacturer', 'Unknown')
            
            # Extract template name (remove manufacturer prefix if present)
            if ' - ' in template_key and template_key.startswith(manufacturer):
                template_name = template_key.split(' - ', 1)[1]
            else:
                template_name = template_key
                
            if manufacturer not in manufacturers:
                manufacturers[manufacturer] = []
            manufacturers[manufacturer].append((template_name, template_key, template_data))
        
        # Add manufacturers and their templates to tree
        for manufacturer, templates_list in sorted(manufacturers.items()):
            # Add manufacturer node
            manufacturer_node = self.template_tree.insert('', 'end', text=manufacturer, open=False)
            
            # Add templates under manufacturer
            for template_name, template_key, template_data in sorted(templates_list, key=lambda x: x[0]):
                # Show module info in template display
                module_spec = template_data.get('module_spec', {})
                model = module_spec.get('model', 'Unknown')
                wattage = module_spec.get('wattage', 0)
                template_text = f"{template_name} ({model} - {wattage}W)"
                self.template_tree.insert(manufacturer_node, 'end', text=template_text, values=(template_key,))
            
    def on_template_select(self, event=None):
        """Handle template selection event"""
        selection = self.template_tree.selection()
        if not selection:
            return
            
        item = selection[0]
        
        # Check if this is a template (has values) or manufacturer (no values)
        values = self.template_tree.item(item, 'values')
        if not values:
            # This is a manufacturer node, not a template
            return
            
        # This is a template, load it
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
                motor_gap_m=float(self.motor_gap_var.get()),
                motor_position_after_string=int(self.motor_position_var.get())
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
                "motor_gap_m": template.motor_gap_m,
                "motor_position_after_string": template.motor_position_after_string,
                "module_spec": {
                    "manufacturer": template.module_spec.manufacturer,
                    "model": template.module_spec.model,
                    "type": template.module_spec.type.value,
                    "length_mm": template.module_spec.length_mm,
                    "width_mm": template.module_spec.width_mm,
                    "depth_mm": template.module_spec.depth_mm,
                    "weight_kg": template.module_spec.weight_kg,
                    "wattage": template.module_spec.wattage,
                    "vmp": template.module_spec.vmp,
                    "imp": template.module_spec.imp,
                    "voc": template.module_spec.voc,
                    "isc": template.module_spec.isc,
                    "max_system_voltage": template.module_spec.max_system_voltage
                }
            }
            
            self.save_templates()
            self.update_template_list()
            
            if self.on_template_saved:
                self.on_template_saved(template)
                
            messagebox.showinfo("Success", f"Template '{name}' saved successfully")
            
    def load_template(self):
        """Load selected template"""
        selection = self.template_tree.selection()
        if not selection:
            return
            
        item = selection[0]
        values = self.template_tree.item(item, 'values')
        if not values:
            return  # This is a manufacturer node
            
        template_key = values[0]
        template_data = self.templates[template_key]
        
        self.name_var.set(template_key.split(' - ', 1)[-1] if ' - ' in template_key else template_key)
        self.orientation_var.set(template_data["module_orientation"])
        self.modules_string_var.set(str(template_data["modules_per_string"]))
        self.strings_tracker_var.set(str(template_data["strings_per_tracker"]))
        self.spacing_var.set(str(template_data["module_spacing_m"]))
        self.motor_gap_var.set(str(template_data["motor_gap_m"]))
        self.motor_position_var.set(str(template_data.get("motor_position_after_string", 2)))
        
        # If we have module spec data, use it
        if "module_spec" in template_data:
            module_data = template_data["module_spec"]
            module_spec = ModuleSpec(
                manufacturer=module_data["manufacturer"],
                model=module_data["model"],
                type=ModuleType(module_data["type"]),
                length_mm=module_data["length_mm"],
                width_mm=module_data["width_mm"],
                depth_mm=module_data["depth_mm"],
                weight_kg=module_data["weight_kg"],
                wattage=module_data["wattage"],
                vmp=module_data["vmp"],
                imp=module_data["imp"],
                voc=module_data["voc"],
                isc=module_data["isc"],
                max_system_voltage=module_data["max_system_voltage"]
            )
            self.module_spec = module_spec
        
        self.update_preview()
        
    def delete_template(self):
        """Delete selected template"""
        selection = self.template_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a template to delete")
            return
            
        item = selection[0]
        
        # Check if this is a template (has values) or manufacturer (no values)
        values = self.template_tree.item(item, 'values')
        if not values:
            # This is a manufacturer node, ask if they want to delete all templates
            manufacturer = self.template_tree.item(item, 'text')
            templates_to_delete = [key for key, template_data in self.templates.items() 
                                if template_data.get('module_spec', {}).get('manufacturer') == manufacturer]
            
            if not templates_to_delete:
                return
                
            if messagebox.askyesno("Confirm", 
                                f"Delete all {len(templates_to_delete)} templates from {manufacturer}?"):
                for template_key in templates_to_delete:
                    del self.templates[template_key]
                self.save_templates()
                self.update_template_list()
                
                # Call the deletion callback if provided
                if self.on_template_deleted:
                    for template_key in templates_to_delete:
                        self.on_template_deleted(template_key)
            return
            
        # Delete individual template
        template_key = values[0]
        template_text = self.template_tree.item(item, 'text')
        
        if messagebox.askyesno("Confirm", f"Delete template '{template_text}'?"):
            if template_key in self.templates:
                del self.templates[template_key]
                self.save_templates()
                self.update_template_list()
                
                # Call the deletion callback if provided
                if self.on_template_deleted:
                    self.on_template_deleted(template_key)
            
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

        # Calculate motor position and modules above/below
        motor_position = template.get_motor_position()
        strings_above_motor = motor_position
        strings_below_motor = template.strings_per_tracker - motor_position
        modules_above_motor = strings_above_motor * template.modules_per_string
        modules_below_motor = strings_below_motor * template.modules_per_string
        
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

        # Draw motor (only if there are strings above motor)
        if modules_above_motor > 0:
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