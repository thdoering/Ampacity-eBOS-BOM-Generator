import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
from pathlib import Path
from typing import Optional, Callable, Dict
from ..models.module import ModuleSpec, ModuleType
from ..utils.pan_parser import parse_pan_file
from ..models.module import ModuleSpec, ModuleType, ModuleOrientation

class ModuleManager(ttk.Frame):
    def __init__(self, parent, on_module_selected: Optional[Callable[[ModuleSpec], None]] = None):
        super().__init__(parent)
        self.parent = parent
        self.on_module_selected = on_module_selected
        self.modules: Dict[str, ModuleSpec] = {}
        
        self.setup_ui()
        self.load_modules()
        
    def setup_ui(self):
        """Create and arrange UI components"""
        # Main container with padding
        main_container = ttk.Frame(self, padding="10")
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Left side - Module List
        list_frame = ttk.LabelFrame(main_container, text="Module Library", padding="5")
        list_frame.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Create Treeview for hierarchical module display
        self.module_tree = ttk.Treeview(list_frame, height=15)
        self.module_tree.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure tree columns
        self.module_tree.heading('#0', text='Modules')
        self.module_tree.column('#0', width=350)

        # Add scrollbar for tree
        tree_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.module_tree.yview)
        tree_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.module_tree.configure(yscrollcommand=tree_scrollbar.set)

        # Bind selection event
        self.module_tree.bind('<<TreeviewSelect>>', self.on_module_select)
        
        button_frame = ttk.Frame(list_frame)
        button_frame.grid(row=1, column=0, padx=5, pady=5)
        
        ttk.Button(button_frame, text="Import PAN", command=self.import_pan).grid(row=0, column=0, padx=2)
        ttk.Button(button_frame, text="Delete", command=self.delete_module).grid(row=0, column=1, padx=2)
        
        # Right side - Module Editor
        editor_frame = ttk.LabelFrame(main_container, text="Module Details", padding="5")
        editor_frame.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Basic Info
        ttk.Label(editor_frame, text="Manufacturer:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.manufacturer_var = tk.StringVar()
        ttk.Entry(editor_frame, textvariable=self.manufacturer_var).grid(row=0, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(editor_frame, text="Model:").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.model_var = tk.StringVar()
        ttk.Entry(editor_frame, textvariable=self.model_var).grid(row=1, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(editor_frame, text="Type:").grid(row=2, column=0, padx=5, pady=2, sticky=tk.W)
        self.type_var = tk.StringVar(value=ModuleType.MONO_PERC.value)
        type_combo = ttk.Combobox(editor_frame, textvariable=self.type_var)
        type_combo['values'] = [t.value for t in ModuleType]
        type_combo.grid(row=2, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        # Physical specs
        phys_frame = ttk.LabelFrame(editor_frame, text="Physical Specifications", padding="5")
        phys_frame.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        ttk.Label(phys_frame, text="Length (mm):").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.length_var = tk.StringVar()
        ttk.Entry(phys_frame, textvariable=self.length_var).grid(row=0, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(phys_frame, text="Width (mm):").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.width_var = tk.StringVar()
        ttk.Entry(phys_frame, textvariable=self.width_var).grid(row=1, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        # Electrical specs
        elec_frame = ttk.LabelFrame(editor_frame, text="Electrical Specifications", padding="5")
        elec_frame.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        ttk.Label(elec_frame, text="Wattage:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.wattage_var = tk.StringVar()
        ttk.Entry(elec_frame, textvariable=self.wattage_var).grid(row=0, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(elec_frame, text="Vmp:").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.vmp_var = tk.StringVar()
        ttk.Entry(elec_frame, textvariable=self.vmp_var).grid(row=1, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(elec_frame, text="Imp:").grid(row=2, column=0, padx=5, pady=2, sticky=tk.W)
        self.imp_var = tk.StringVar()
        ttk.Entry(elec_frame, textvariable=self.imp_var).grid(row=2, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(elec_frame, text="Voc:").grid(row=3, column=0, padx=5, pady=2, sticky=tk.W)
        self.voc_var = tk.StringVar()
        ttk.Entry(elec_frame, textvariable=self.voc_var).grid(row=3, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(elec_frame, text="Isc:").grid(row=4, column=0, padx=5, pady=2, sticky=tk.W)
        self.isc_var = tk.StringVar()
        ttk.Entry(elec_frame, textvariable=self.isc_var).grid(row=4, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        # Temperature Coefficients
        temp_frame = ttk.LabelFrame(editor_frame, text="Temperature Coefficients (%/°C)", padding="5")
        temp_frame.grid(row=5, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        ttk.Label(temp_frame, text="Pmax (%/°C):").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.temp_coeff_pmax_var = tk.StringVar()
        ttk.Entry(temp_frame, textvariable=self.temp_coeff_pmax_var).grid(row=0, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(temp_frame, text="Voc (%/°C):").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.temp_coeff_voc_var = tk.StringVar()
        ttk.Entry(temp_frame, textvariable=self.temp_coeff_voc_var).grid(row=1, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(temp_frame, text="Isc (%/°C):").grid(row=2, column=0, padx=5, pady=2, sticky=tk.W)
        self.temp_coeff_isc_var = tk.StringVar()
        ttk.Entry(temp_frame, textvariable=self.temp_coeff_isc_var).grid(row=2, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        # Save button
        ttk.Button(editor_frame, text="Save Module", command=self.save_module).grid(row=6, column=0, columnspan=2, pady=10)
        
    def create_module_spec(self) -> Optional[ModuleSpec]:
        """Create ModuleSpec from current UI values"""
        try:
            # Get temperature coefficients if provided
            temp_coeff_pmax = None
            temp_coeff_voc = None
            temp_coeff_isc = None
            
            if self.temp_coeff_pmax_var.get():
                try:
                    temp_coeff_pmax = float(self.temp_coeff_pmax_var.get())
                except ValueError:
                    pass
                    
            if self.temp_coeff_voc_var.get():
                try:
                    temp_coeff_voc = float(self.temp_coeff_voc_var.get())
                except ValueError:
                    pass
                    
            if self.temp_coeff_isc_var.get():
                try:
                    temp_coeff_isc = float(self.temp_coeff_isc_var.get())
                except ValueError:
                    pass
            
            return ModuleSpec(
                manufacturer=self.manufacturer_var.get(),
                model=self.model_var.get(),
                type=ModuleType(self.type_var.get()),
                length_mm=float(self.length_var.get()),
                width_mm=float(self.width_var.get()),
                depth_mm=40,  # Default
                weight_kg=25,  # Default
                wattage=float(self.wattage_var.get()),
                vmp=float(self.vmp_var.get()),
                imp=float(self.imp_var.get()),
                voc=float(self.voc_var.get()),
                isc=float(self.isc_var.get()),
                max_system_voltage=1500,  # Default
                temperature_coefficient_pmax=temp_coeff_pmax,
                temperature_coefficient_voc=temp_coeff_voc,
                temperature_coefficient_isc=temp_coeff_isc
            )
        except (ValueError, TypeError) as e:
            messagebox.showerror("Error", str(e))
            return None
            
    def load_modules(self):
        """Load saved modules from JSON file"""
        module_path = Path('data/module_templates.json')
        if not module_path.exists():
            return
            
        try:
            with open(module_path, 'r') as f:
                data = json.load(f)
                
            # Handle both old flat structure and new hierarchical structure
            self.modules = {}
            
            if data:
                # Check if this is the new hierarchical format
                first_value = next(iter(data.values()))
                if isinstance(first_value, dict) and not any(key in first_value for key in ['manufacturer', 'model', 'type']):
                    # New hierarchical format: Manufacturer -> Model -> module_data
                    for manufacturer, models in data.items():
                        for model, module_data in models.items():
                            module_key = f"{manufacturer} {model}"
                            # Handle backward compatibility for temperature coefficients
                            module_params = {k: v for k, v in module_data.items() if k != 'type' and k != 'default_orientation' and k != 'temperature_coefficient'}
                            
                            # If old temperature_coefficient exists and new ones don't, use it for Pmax
                            if 'temperature_coefficient' in module_data and 'temperature_coefficient_pmax' not in module_data:
                                module_params['temperature_coefficient_pmax'] = module_data['temperature_coefficient']
                            
                            self.modules[module_key] = ModuleSpec(
                                **module_params,
                                type=ModuleType(module_data['type']),
                                default_orientation=ModuleOrientation(module_data.get('default_orientation', ModuleOrientation.PORTRAIT.value))
                            )
                else:
                    # Old flat format: convert on the fly
                    self.modules = {
                        name: ModuleSpec(
                            **{k: v for k, v in specs.items() if k != 'type' and k != 'default_orientation'},
                            type=ModuleType(specs['type']),
                            default_orientation=ModuleOrientation(specs.get('default_orientation', ModuleOrientation.PORTRAIT.value))
                        ) 
                        for name, specs in data.items()
                    }
                    
            self.update_module_list()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load modules: {str(e)}")
                
    def save_modules(self):
        """Save modules to JSON file in hierarchical format"""
        module_path = Path('data/module_templates.json')
        module_path.parent.mkdir(exist_ok=True)
        
        # Organize modules by manufacturer
        hierarchical_data = {}
        
        for module in self.modules.values():
            manufacturer = module.manufacturer
            model = module.model
            
            if manufacturer not in hierarchical_data:
                hierarchical_data[manufacturer] = {}
                
            hierarchical_data[manufacturer][model] = {
                **module.__dict__,
                'type': module.type.value,  # Convert enum to string
                'default_orientation': module.default_orientation.value
            }
        
        with open(module_path, 'w') as f:
            json.dump(hierarchical_data, f, indent=2)
            
    def update_module_list(self):
        """Update the module tree view"""
        # Clear existing items
        for item in self.module_tree.get_children():
            self.module_tree.delete(item)
        
        # Group modules by manufacturer
        manufacturers = {}
        for module_key, module in self.modules.items():
            manufacturer = module.manufacturer
            if manufacturer not in manufacturers:
                manufacturers[manufacturer] = []
            manufacturers[manufacturer].append((module.model, module_key, module))
        
        # Add manufacturers and their modules to tree
        for manufacturer, modules_list in sorted(manufacturers.items()):
            # Add manufacturer node
            manufacturer_node = self.module_tree.insert('', 'end', text=manufacturer, open=False)
            
            # Add modules under manufacturer
            for model, module_key, module in sorted(modules_list, key=lambda x: x[0]):
                module_text = f"{model} ({module.wattage}W)"
                self.module_tree.insert(manufacturer_node, 'end', text=module_text, values=(module_key,))
            
    def import_pan(self):
        """Import module from PAN file"""
        file_path = filedialog.askopenfilename(
            filetypes=[("PAN files", "*.pan"), ("All files", "*.*")]
        )
        if not file_path:
            return
            
        try:
            with open(file_path, 'r') as f:
                content = f.read()
            
            params = parse_pan_file(content)
            module = ModuleSpec(
                type=ModuleType.MONO_PERC,  # Default type
                **params
            )
            
            name = f"{module.manufacturer} {module.model}"
            self.modules[name] = module
            self.save_modules()
            self.update_module_list()
            messagebox.showinfo("Success", f"Imported module: {name}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import PAN file: {str(e)}")
            
    def save_module(self):
        """Save current module"""
        module = self.create_module_spec()
        if not module:
            return
            
        name = f"{module.manufacturer} {module.model}"
        if name in self.modules:
            if not messagebox.askyesno("Confirm", f"Module '{name}' already exists. Overwrite?"):
                return
                
        self.modules[name] = module
        self.save_modules()
        self.update_module_list()
        
        if self.on_module_selected:
            self.on_module_selected(module)
            
        messagebox.showinfo("Success", f"Module '{name}' saved successfully")
        
    def delete_module(self):
        """Delete selected module"""
        selection = self.module_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a module to delete")
            return
            
        item = selection[0]
        
        # Check if this is a module (has values) or manufacturer (no values)
        values = self.module_tree.item(item, 'values')
        if not values:
            # This is a manufacturer node, ask if they want to delete all modules
            manufacturer = self.module_tree.item(item, 'text')
            modules_to_delete = [key for key, module in self.modules.items() 
                            if module.manufacturer == manufacturer]
            
            if not modules_to_delete:
                return
                
            if messagebox.askyesno("Confirm", 
                                f"Delete all {len(modules_to_delete)} modules from {manufacturer}?"):
                for module_key in modules_to_delete:
                    del self.modules[module_key]
                self.save_modules()
                self.update_module_list()
            return
            
        # Delete individual module
        module_key = values[0]
        module_text = self.module_tree.item(item, 'text')
        
        if messagebox.askyesno("Confirm", f"Delete module '{module_text}'?"):
            if module_key in self.modules:
                del self.modules[module_key]
                self.save_modules()
                self.update_module_list()
            
    def on_module_select(self, event=None):
        """Handle module selection from tree"""
        selection = self.module_tree.selection()
        if not selection:
            return
            
        item = selection[0]
        
        # Check if this is a module (has values) or manufacturer (no values)
        values = self.module_tree.item(item, 'values')
        if not values:
            # This is a manufacturer node, not a module
            return
            
        module_key = values[0]
        if module_key in self.modules:
            module = self.modules[module_key]
            
            # Update UI with selected module
            self.manufacturer_var.set(module.manufacturer)
            self.model_var.set(module.model)
            self.type_var.set(module.type.value)
            self.length_var.set(str(module.length_mm))
            self.width_var.set(str(module.width_mm))
            self.wattage_var.set(str(module.wattage))
            self.vmp_var.set(str(module.vmp))
            self.imp_var.set(str(module.imp))
            self.voc_var.set(str(module.voc))
            self.isc_var.set(str(module.isc))
            
            # Set temperature coefficients if available
            if hasattr(module, 'temperature_coefficient_pmax') and module.temperature_coefficient_pmax is not None:
                self.temp_coeff_pmax_var.set(str(module.temperature_coefficient_pmax))
            else:
                self.temp_coeff_pmax_var.set("")
                
            if hasattr(module, 'temperature_coefficient_voc') and module.temperature_coefficient_voc is not None:
                self.temp_coeff_voc_var.set(str(module.temperature_coefficient_voc))
            else:
                self.temp_coeff_voc_var.set("")
                
            if hasattr(module, 'temperature_coefficient_isc') and module.temperature_coefficient_isc is not None:
                self.temp_coeff_isc_var.set(str(module.temperature_coefficient_isc))
            else:
                self.temp_coeff_isc_var.set("")

            # Call the callback if provided
            if self.on_module_selected:
                self.on_module_selected(module)