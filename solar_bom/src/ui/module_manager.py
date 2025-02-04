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
        
        self.module_listbox = tk.Listbox(list_frame, width=40, height=15)
        self.module_listbox.grid(row=0, column=0, padx=5, pady=5)
        self.module_listbox.bind('<<ListboxSelect>>', self.on_module_select)
        
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
        
        # Save button
        ttk.Button(editor_frame, text="Save Module", command=self.save_module).grid(row=5, column=0, columnspan=2, pady=10)
        
    def create_module_spec(self) -> Optional[ModuleSpec]:
        """Create ModuleSpec from current UI values"""
        try:
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
                max_system_voltage=1500  # Default
            )
        except (ValueError, TypeError) as e:
            messagebox.showerror("Error", str(e))
            return None
            
    def load_modules(self):
        """Load saved modules from JSON file"""
        module_path = Path('data/modules.json')
        if not module_path.exists():
            return
            
        try:
            with open(module_path, 'r') as f:
                data = json.load(f)
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
        """Save modules to JSON file"""
        module_path = Path('data/modules.json')
        module_path.parent.mkdir(exist_ok=True)
        
        data = {
            f"{module.manufacturer} {module.model}": {
                **module.__dict__,
                'type': module.type.value,  # Convert enum to string
                'default_orientation': module.default_orientation.value
            }
            for module in self.modules.values()
        }
        
        with open(module_path, 'w') as f:
            json.dump(data, f, indent=2)
            
    def update_module_list(self):
        """Update the module listbox"""
        self.module_listbox.delete(0, tk.END)
        for module in self.modules.values():
            self.module_listbox.insert(tk.END, f"{module.manufacturer} {module.model}")
            
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
        selection = self.module_listbox.curselection()
        if not selection:
            return
            
        name = self.module_listbox.get(selection[0])
        if messagebox.askyesno("Confirm", f"Delete module '{name}'?"):
            del self.modules[name]
            self.save_modules()
            self.update_module_list()
            
    def on_module_select(self, event=None):
        """Handle module selection"""
        selection = self.module_listbox.curselection()
        if not selection:
            return
            
        name = self.module_listbox.get(selection[0])
        module = self.modules[name]

        print(f"ModuleManager: Module selected: {module}")  # Debug print
        
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

        # Call the callback if provided
        if self.on_module_selected:
            print("ModuleManager: Calling on_module_selected callback")  # Debug print
            self.on_module_selected(module)