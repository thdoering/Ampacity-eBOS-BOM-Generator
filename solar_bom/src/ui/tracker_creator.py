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
        on_template_deleted: Optional[Callable[[str], None]] = None,
        current_project=None,
        on_template_enabled_changed: Optional[Callable[[], None]] = None):
        super().__init__(parent)
        self.current_project = current_project
        self.parent = parent
        self.on_template_saved = on_template_saved
        self.on_template_deleted = on_template_deleted
        self.on_template_enabled_changed = on_template_enabled_changed
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
        
        # Create Treeview for hierarchical template display with checkbox column
        self.template_tree = ttk.Treeview(template_frame, columns=('enabled',), height=15)
        self.template_tree.grid(row=0, column=0, padx=2, pady=2, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure tree columns
        self.template_tree.heading('#0', text='Templates')
        self.template_tree.heading('enabled', text='Enabled')
        self.template_tree.column('#0', width=250)
        self.template_tree.column('enabled', width=60, anchor='center')

        # Configure tags for visual feedback
        self.template_tree.tag_configure('checked', foreground='black')
        self.template_tree.tag_configure('unchecked', foreground='gray60')

        # Add scrollbar for tree
        tree_scrollbar = ttk.Scrollbar(template_frame, orient=tk.VERTICAL, command=self.template_tree.yview)
        tree_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.template_tree.configure(yscrollcommand=tree_scrollbar.set)

        # Bind click events for checkbox functionality
        self.template_tree.bind('<Button-1>', self.on_tree_click)
        self.template_tree.bind('<Double-1>', self.toggle_template_enabled)

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

        # Motor Placement Type
        ttk.Label(editor_frame, text="Motor Placement:").grid(row=7, column=0, padx=5, pady=2, sticky=tk.W)
        self.motor_placement_var = tk.StringVar(value="between_strings")
        placement_combo = ttk.Combobox(editor_frame, textvariable=self.motor_placement_var, 
                                     values=["between_strings", "middle_of_string"], state="readonly")
        placement_combo.grid(row=7, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))

        # Motor Position (Between Strings)
        ttk.Label(editor_frame, text="Motor After String:").grid(row=8, column=0, padx=5, pady=2, sticky=tk.W)
        self.motor_position_var = tk.StringVar(value="2")  # Default to middle-ish
        self.motor_position_spinbox = ttk.Spinbox(editor_frame, from_=0, to=1, textvariable=self.motor_position_var,
            increment=1, validate='all', validatecommand=(self.register(lambda val: val.isdigit() or val == ""), '%P'))
        self.motor_position_spinbox.grid(row=8, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        # Motor String Index (Middle of String)
        ttk.Label(editor_frame, text="Motor in String:").grid(row=9, column=0, padx=5, pady=2, sticky=tk.W)
        self.motor_string_var = tk.StringVar(value="2")
        self.motor_string_spinbox = ttk.Spinbox(editor_frame, from_=1, to=2, textvariable=self.motor_string_var,
            increment=1, validate='all', validatecommand=(self.register(lambda val: val.isdigit() or val == ""), '%P'))
        self.motor_string_spinbox.grid(row=9, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        # Motor Split Controls (Middle of String)
        self.split_frame = ttk.Frame(editor_frame)
        self.split_frame.grid(row=10, column=0, columnspan=2, padx=5, pady=2, sticky=(tk.W, tk.E))
        ttk.Label(self.split_frame, text="Split (North/South):").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        
        self.motor_split_north_var = tk.StringVar(value="14")
        self.motor_split_south_var = tk.StringVar(value="14")
        
        split_entry_frame = ttk.Frame(self.split_frame)
        split_entry_frame.grid(row=0, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        self.north_spinbox = ttk.Spinbox(split_entry_frame, from_=0, to=100, width=6, 
                                   textvariable=self.motor_split_north_var,
                                   increment=1, validate='all', 
                                   validatecommand=(self.register(lambda val: val.isdigit() or val == ""), '%P'))
        self.north_spinbox.grid(row=0, column=0, padx=2)
        
        ttk.Label(split_entry_frame, text="/").grid(row=0, column=1, padx=2)
        
        self.south_label = ttk.Label(split_entry_frame, text="14")
        self.south_label.grid(row=0, column=2, padx=2)

        # Total Modules Display
        ttk.Label(editor_frame, text="Modules per Tracker:").grid(row=11, column=0, padx=5, pady=2, sticky=tk.W)
        self.total_modules_label = ttk.Label(editor_frame, text="--")
        try:
            modules = int(self.modules_string_var.get()) * int(self.strings_tracker_var.get())
            self.total_modules_label.config(text=str(modules))
        except ValueError:
            self.total_modules_label.config(text="--")
        self.total_modules_label.grid(row=11, column=1, padx=5, pady=2, sticky=tk.W)
        
        # Calculated Dimensions Display
        dims_frame = ttk.LabelFrame(editor_frame, text="Tracker Dimensions", padding="5")
        dims_frame.grid(row=12, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
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
        ttk.Button(editor_frame, text="Save Template", command=self.save_template).grid(row=13, column=0, columnspan=2, pady=10)
        
        # Preview Canvas - Right column
        preview_frame = ttk.LabelFrame(main_container, text="Preview", padding="5")
        preview_frame.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.N, tk.S))

        self.canvas = tk.Canvas(preview_frame, width=300, height=600, bg='white')
        self.canvas.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.N, tk.S))
        
        # Bind events for real-time preview updates
        for var in [self.modules_string_var, self.strings_tracker_var, 
                self.spacing_var, self.motor_gap_var, self.orientation_var, self.motor_position_var,
                self.motor_placement_var, self.motor_string_var, self.motor_split_north_var]:
            var.trace('w', lambda *args: self.update_preview())

        # Add special traces for motor controls
        self.strings_tracker_var.trace('w', self.update_motor_position_range)
        self.strings_tracker_var.trace('w', self.update_motor_string_range)
        self.motor_placement_var.trace('w', self.update_motor_placement_visibility)
        self.modules_string_var.trace('w', self.update_motor_split_calculation)
        self.motor_split_north_var.trace('w', self.update_motor_split_calculation)
        
        # Initialize visibility
        self.update_motor_placement_visibility()
        self.update_motor_split_calculation()

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

    def update_motor_placement_visibility(self, *args):
        """Show/hide motor placement controls based on selected type"""
        placement_type = self.motor_placement_var.get()
        
        if placement_type == "between_strings":
            # Show between strings controls
            self.motor_position_spinbox.grid()
            # Hide middle of string controls
            self.motor_string_spinbox.grid_remove()
            self.split_frame.grid_remove()
        else:  # middle_of_string
            # Hide between strings controls  
            self.motor_position_spinbox.grid_remove()
            # Show middle of string controls
            self.motor_string_spinbox.grid()
            self.split_frame.grid()

    def update_motor_string_range(self, *args):
        """Update the motor string spinbox range based on strings per tracker"""
        try:
            strings_count = int(self.strings_tracker_var.get())
            self.motor_string_spinbox.config(to=strings_count)
            
            # Ensure current value is within new range
            current_val = int(self.motor_string_var.get())
            if current_val > strings_count:
                # Default to middle string
                default_pos = (strings_count + 1) // 2
                self.motor_string_var.set(str(default_pos))
        except ValueError:
            pass

    def update_motor_split_calculation(self, *args):
        """Auto-calculate south split when north split or modules per string changes"""
        try:
            modules_per_string = int(self.modules_string_var.get())
            north_split = int(self.motor_split_north_var.get())
            south_split = modules_per_string - north_split
            
            # Update the south label
            self.south_label.config(text=str(max(0, south_split)))
            self.motor_split_south_var.set(str(max(0, south_split)))
            
            # Update north spinbox range
            self.north_spinbox.config(to=modules_per_string)
        except ValueError:
            self.south_label.config(text="--")  

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
        """Update the template tree view with checkbox functionality"""
        # Save expanded state of all nodes
        expanded_nodes = set()
        
        def save_expanded_state(item, path=""):
            """Recursively save expanded state of tree"""
            if self.template_tree.item(item, 'open'):
                item_text = self.template_tree.item(item, 'text')
                full_path = f"{path}/{item_text}" if path else item_text
                expanded_nodes.add(full_path)
            
            for child in self.template_tree.get_children(item):
                item_text = self.template_tree.item(item, 'text')
                child_path = f"{path}/{item_text}" if path else item_text
                save_expanded_state(child, child_path)
        
        # Save expanded state of all root items
        for item in self.template_tree.get_children():
            save_expanded_state(item)
        
        # Clear existing items
        for item in self.template_tree.get_children():
            self.template_tree.delete(item)
        
        # Group templates hierarchically
        hierarchy = {}
        
        for template_key, template_data in self.templates.items():
            # Extract grouping information
            module_spec = template_data.get('module_spec', {})
            manufacturer = module_spec.get('manufacturer', 'Unknown')
            model = module_spec.get('model', 'Unknown')
            modules_per_string = template_data.get('modules_per_string', 0)
            
            # Build hierarchy: Manufacturer -> Model -> String Size -> Templates
            if manufacturer not in hierarchy:
                hierarchy[manufacturer] = {}
            if model not in hierarchy[manufacturer]:
                hierarchy[manufacturer][model] = {}
            if modules_per_string not in hierarchy[manufacturer][model]:
                hierarchy[manufacturer][model][modules_per_string] = []
            
            # Extract template name (remove manufacturer prefix if present)
            if ' - ' in template_key and template_key.startswith(manufacturer):
                template_name = template_key.split(' - ', 1)[1]
            else:
                template_name = template_key
                
            hierarchy[manufacturer][model][modules_per_string].append((template_name, template_key, template_data))
        
        # Store template_key mapping for quick lookup
        self.tree_item_to_template = {}
        
        # Add items to tree hierarchically
        for manufacturer in sorted(hierarchy.keys()):
            # Add manufacturer node
            manufacturer_path = manufacturer
            manufacturer_node = self.template_tree.insert('', 'end', text=manufacturer, values=('',), 
                                                        open=(manufacturer_path in expanded_nodes))
            
            for model in sorted(hierarchy[manufacturer].keys()):
                # Add model node
                model_path = f"{manufacturer}/{model}"
                model_node = self.template_tree.insert(manufacturer_node, 'end', text=model, values=('',),
                                                    open=(model_path in expanded_nodes))
                
                for string_size in sorted(hierarchy[manufacturer][model].keys()):
                    # Add string size node
                    string_size_text = f"{string_size} modules per string"
                    string_size_path = f"{manufacturer}/{model}/{string_size_text}"
                    string_size_node = self.template_tree.insert(model_node, 'end', text=string_size_text, 
                                                            values=('',),
                                                            open=(string_size_path in expanded_nodes))
                    
                    # Add templates under string size
                    templates_list = hierarchy[manufacturer][model][string_size]
                    for template_name, template_key, template_data in sorted(templates_list, key=lambda x: x[0]):
                        # Check if template is enabled
                        is_enabled = self._is_template_enabled(template_key)
                        checkbox = '☑' if is_enabled else '☐'
                        tag = 'checked' if is_enabled else 'unchecked'
                        
                        # Insert template with checkbox
                        template_item = self.template_tree.insert(string_size_node, 'end', 
                                                text=template_name,
                                                values=(checkbox,), 
                                                tags=(tag,))
                        
                        # Store mapping from tree item to template key
                        self.tree_item_to_template[template_item] = template_key
            
    def on_template_select(self, event=None):
        """Handle template selection event"""
        selection = self.template_tree.selection()
        if not selection:
            return
        
        item = selection[0]
        values = self.template_tree.item(item, 'values')
        if not values or values[0] == '':
            return  # This is a parent node (manufacturer, model, or string size)
            
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
                motor_position_after_string=int(self.motor_position_var.get()),
                # New motor placement fields
                motor_placement_type=self.motor_placement_var.get(),
                motor_string_index=int(self.motor_string_var.get()),
                motor_split_north=int(self.motor_split_north_var.get()),
                motor_split_south=int(self.motor_split_south_var.get())
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
                # New motor placement fields
                "motor_placement_type": template.motor_placement_type,
                "motor_string_index": template.motor_string_index,
                "motor_split_north": template.motor_split_north,
                "motor_split_south": template.motor_split_south,
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

            # Auto-enable newly created template with full key including manufacturer
            if self.current_project:
                # Get the manufacturer from the template's module spec
                manufacturer = template.module_spec.manufacturer
                full_template_key = f"{manufacturer} - {name}"
                self._add_enabled_template(full_template_key)

            # Update the UI to show the template as enabled
            for manufacturer_item in self.template_tree.get_children():
                for template_item in self.template_tree.get_children(manufacturer_item):
                    if template_item in self.tree_item_to_template:
                        if self.tree_item_to_template[template_item] == full_template_key:
                            values = list(self.template_tree.item(template_item, 'values'))
                            values[0] = '☑'
                            self.template_tree.item(template_item, values=values, tags=('checked',))
                            break

            # Expand the parent nodes to show the new template
            for manufacturer_item in self.template_tree.get_children():
                manufacturer_text = self.template_tree.item(manufacturer_item, 'text')
                if manufacturer_text == manufacturer:
                    self.template_tree.item(manufacturer_item, open=True)
                    
                    # Find and expand the model node
                    model = template.module_spec.model
                    for model_item in self.template_tree.get_children(manufacturer_item):
                        model_text = self.template_tree.item(model_item, 'text')
                        if model_text == model:
                            self.template_tree.item(model_item, open=True)
                            
                            # Find and expand the string size node
                            string_size = template.modules_per_string
                            string_size_text = f"{string_size} modules per string"
                            for string_item in self.template_tree.get_children(model_item):
                                string_text = self.template_tree.item(string_item, 'text')
                                if string_text == string_size_text:
                                    self.template_tree.item(string_item, open=True)
                                    break
                            break
                    break
            
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
            
        # Get template key from mapping (similar to toggle_item_enabled)
        if item not in self.tree_item_to_template:
            return
        template_key = self.tree_item_to_template[item]
        template_data = self.templates[template_key]
        
        self.name_var.set(template_key.split(' - ', 1)[-1] if ' - ' in template_key else template_key)
        self.orientation_var.set(template_data["module_orientation"])
        self.modules_string_var.set(str(template_data["modules_per_string"]))
        self.strings_tracker_var.set(str(template_data["strings_per_tracker"]))
        self.spacing_var.set(str(template_data["module_spacing_m"]))
        self.motor_gap_var.set(str(template_data["motor_gap_m"]))
        self.motor_position_var.set(str(template_data.get("motor_position_after_string", 2)))
        
        # Load new motor placement fields with defaults for backward compatibility
        self.motor_placement_var.set(template_data.get("motor_placement_type", "between_strings"))
        self.motor_string_var.set(str(template_data.get("motor_string_index", 2)))
        self.motor_split_north_var.set(str(template_data.get("motor_split_north", 14)))
        self.motor_split_south_var.set(str(template_data.get("motor_split_south", 14)))
        
        # Update UI visibility and calculations
        self.update_motor_placement_visibility()
        self.update_motor_split_calculation()
        
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
        
        # Check if this is a template or a parent node
        values = self.template_tree.item(item, 'values')
        if not values or values[0] == '':
            # This is a parent node - determine which level
            parent = self.template_tree.parent(item)
            grandparent = self.template_tree.parent(parent) if parent else None
            
            if not parent:
                # This is a manufacturer node
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
                            
            elif not grandparent:
                # This is a model node
                manufacturer = self.template_tree.item(parent, 'text')
                model = self.template_tree.item(item, 'text')
                templates_to_delete = [key for key, template_data in self.templates.items() 
                                    if template_data.get('module_spec', {}).get('manufacturer') == manufacturer
                                    and template_data.get('module_spec', {}).get('model') == model]
                
                if not templates_to_delete:
                    return
                    
                if messagebox.askyesno("Confirm", 
                                    f"Delete all {len(templates_to_delete)} templates for {manufacturer} {model}?"):
                    for template_key in templates_to_delete:
                        del self.templates[template_key]
                    self.save_templates()
                    self.update_template_list()
                    
                    # Call the deletion callback if provided
                    if self.on_template_deleted:
                        for template_key in templates_to_delete:
                            self.on_template_deleted(template_key)
                            
            else:
                # This is a string size node
                manufacturer = self.template_tree.item(grandparent, 'text')
                model = self.template_tree.item(parent, 'text')
                string_text = self.template_tree.item(item, 'text')
                # Extract number from "X modules per string"
                string_size = int(string_text.split()[0])
                
                templates_to_delete = [key for key, template_data in self.templates.items() 
                                    if template_data.get('module_spec', {}).get('manufacturer') == manufacturer
                                    and template_data.get('module_spec', {}).get('model') == model
                                    and template_data.get('modules_per_string') == string_size]
                
                if not templates_to_delete:
                    return
                    
                if messagebox.askyesno("Confirm", 
                                    f"Delete all {len(templates_to_delete)} templates for {manufacturer} {model} with {string_size} modules per string?"):
                    for template_key in templates_to_delete:
                        del self.templates[template_key]
                    self.save_templates()
                    self.update_template_list()
                    
                    # Call the deletion callback if provided
                    if self.on_template_deleted:
                        for template_key in templates_to_delete:
                            self.on_template_deleted(template_key)
            return
            
        # This is a template node - existing code continues
        # Get template key from mapping
        if item not in self.tree_item_to_template:
            messagebox.showwarning("Warning", "Template mapping not found")
            return
            
        template_key = self.tree_item_to_template[item]
        template_text = self.template_tree.item(item, 'text')

        if messagebox.askyesno("Confirm", f"Delete template '{template_text}'?"):
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

        # Calculate layout based on motor placement type
        if template.motor_placement_type == "middle_of_string":
            # Calculate total height for middle_of_string placement
            total_modules = template.strings_per_tracker * template.modules_per_string
            total_height = (total_modules * module_height) + \
                        ((total_modules - 1) * template.module_spacing_m) + \
                        template.motor_gap_m
            
            scale = min(280 / module_width, 580 / total_height)
            x_center = (300 - module_width * scale) / 2
            
            # Draw torque tube
            self.canvas.create_line(
                x_center + module_width * scale / 2, 10,
                x_center + module_width * scale / 2, 10 + total_height * scale,
                fill="brown", width=3, tags="torque_tube"
            )
            
            # Draw strings
            current_y = 10
            for string_idx in range(template.strings_per_tracker):
                if string_idx + 1 == template.motor_string_index:  # This string has the motor
                    # Draw north modules
                    for i in range(template.motor_split_north):
                        y_pos = current_y + (i * (module_height + template.module_spacing_m)) * scale
                        self.canvas.create_rectangle(
                            x_center, y_pos,
                            x_center + module_width * scale, y_pos + module_height * scale,
                            fill="lightblue", outline="blue", tags=f"module_string_{string_idx}"
                        )
                    
                    # Draw motor gap
                    motor_y = current_y + (template.motor_split_north * (module_height + template.module_spacing_m)) * scale
                    gap_height = template.motor_gap_m * scale
                    
                    # Draw motor as a red circle
                    circle_radius = min(gap_height / 3, module_width * scale / 4)  # Scale with gap and width
                    circle_center_x = x_center + module_width * scale / 2
                    circle_center_y = motor_y + gap_height / 2
                    
                    self.canvas.create_oval(
                        circle_center_x - circle_radius, circle_center_y - circle_radius,
                        circle_center_x + circle_radius, circle_center_y + circle_radius,
                        fill="red", outline="darkred", width=3, tags="motor"
                    )
                    
                    # Draw south modules
                    south_start_y = motor_y + template.motor_gap_m * scale
                    for i in range(template.motor_split_south):
                        y_pos = south_start_y + (i * (module_height + template.module_spacing_m)) * scale
                        self.canvas.create_rectangle(
                            x_center, y_pos,
                            x_center + module_width * scale, y_pos + module_height * scale,
                            fill="lightblue", outline="blue", tags=f"module_string_{string_idx}"
                        )
                    
                    current_y = south_start_y + (template.motor_split_south * (module_height + template.module_spacing_m)) * scale
                else:
                    # Draw normal string
                    for i in range(template.modules_per_string):
                        y_pos = current_y + (i * (module_height + template.module_spacing_m)) * scale
                        self.canvas.create_rectangle(
                            x_center, y_pos,
                            x_center + module_width * scale, y_pos + module_height * scale,
                            fill="lightblue", outline="blue", tags=f"module_string_{string_idx}"
                        )
                    
                    current_y += (template.modules_per_string * (module_height + template.module_spacing_m)) * scale
        else:
            # Original between_strings logic
            motor_position = template.get_motor_position()
            strings_above_motor = motor_position
            strings_below_motor = template.strings_per_tracker - motor_position
            modules_above_motor = strings_above_motor * template.modules_per_string
            modules_below_motor = strings_below_motor * template.modules_per_string
            
            total_height = ((modules_above_motor + modules_below_motor) * module_height) + \
                            ((modules_above_motor + modules_below_motor - 1) * template.module_spacing_m) + \
                            (template.motor_gap_m if strings_below_motor > 0 else 0)

            scale = min(280 / module_width, 580 / total_height)
            x_center = (300 - module_width * scale) / 2
            
            # Draw torque tube
            self.canvas.create_line(
                x_center + module_width * scale / 2, 10,
                x_center + module_width * scale / 2, 10 + total_height * scale,
                fill="brown", width=3, tags="torque_tube"
            )
            
            # Draw modules above motor
            for string_idx in range(strings_above_motor):
                for module_idx in range(template.modules_per_string):
                    y_pos = 10 + ((string_idx * template.modules_per_string + module_idx) * 
                                (module_height + template.module_spacing_m)) * scale
                    self.canvas.create_rectangle(
                        x_center, y_pos,
                        x_center + module_width * scale, y_pos + module_height * scale,
                        fill="lightblue", outline="blue", tags=f"module_string_{string_idx}"
                    )
            
            # Draw motor gap
            if strings_below_motor > 0:
                motor_y = 10 + (modules_above_motor * (module_height + template.module_spacing_m)) * scale
                gap_height = template.motor_gap_m * scale
                
                # Draw motor as a red circle
                circle_radius = min(gap_height / 3, module_width * scale / 4)  # Scale with gap and width
                circle_center_x = x_center + module_width * scale / 2
                circle_center_y = motor_y + gap_height / 2
                
                self.canvas.create_oval(
                    circle_center_x - circle_radius, circle_center_y - circle_radius,
                    circle_center_x + circle_radius, circle_center_y + circle_radius,
                    fill="red", outline="darkred", width=3, tags="motor"
                )
            
            # Draw modules below motor
            for string_idx in range(strings_below_motor):
                for module_idx in range(template.modules_per_string):
                    y_offset = template.motor_gap_m if strings_below_motor > 0 else 0
                    y_pos = 10 + ((modules_above_motor + string_idx * template.modules_per_string + module_idx) * 
                                (module_height + template.module_spacing_m) + y_offset) * scale
                    self.canvas.create_rectangle(
                        x_center, y_pos,
                        x_center + module_width * scale, y_pos + module_height * scale,
                        fill="lightblue", outline="blue", tags=f"module_string_{strings_above_motor + string_idx}"
                    )

        # Update dimension labels
        dims = template.get_physical_dimensions()
        self.length_label.config(text=f"Length: {dims[0]:.2f}m")
        self.width_label.config(text=f"Width: {dims[1]:.2f}m")

    def on_tree_click(self, event):
        """Handle tree click events for checkbox functionality"""
        item = self.template_tree.identify_row(event.y)
        column = self.template_tree.identify_column(event.x)
        
        # If clicked on the enabled column, toggle selection
        if item and column == '#1':  # Second column is enabled
            self.toggle_item_enabled(item)

    def toggle_template_enabled(self, event):
        """Toggle enabled state on double-click"""
        item = self.template_tree.focus()
        if item:
            self.toggle_item_enabled(item)

    def toggle_item_enabled(self, item):
        """Toggle the enabled state of a template"""
        # Check if this is a template (has checkbox) or parent node
        values = list(self.template_tree.item(item, 'values'))
        if not values or len(values) == 0 or values[0] == '':
            # This is a parent node, not a template
            return

    def _add_enabled_template(self, template_key):
        """Add template to enabled list"""
        if self.current_project and template_key not in self.current_project.enabled_templates:
            self.current_project.enabled_templates.append(template_key)
            if self.on_template_enabled_changed:
                self.on_template_enabled_changed()

    def _remove_enabled_template(self, template_key):
        """Remove template from enabled list"""
        if self.current_project and template_key in self.current_project.enabled_templates:
            self.current_project.enabled_templates.remove(template_key)
            if self.on_template_enabled_changed:
                self.on_template_enabled_changed()

    def _is_template_enabled(self, template_key):
        """Check if template is enabled for current project"""
        if not self.current_project:
            return True  # Default to enabled if no project
        return template_key in self.current_project.enabled_templates