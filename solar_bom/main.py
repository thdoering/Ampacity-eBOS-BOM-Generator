import tkinter as tk
from tkinter import ttk, messagebox, Menu
from src.models.module import ModuleSpec, ModuleType
from src.models.project import Project
from src.ui.tracker_creator import TrackerTemplateCreator
from src.ui.module_manager import ModuleManager
from src.ui.block_configurator import BlockConfigurator
from src.ui.inverter_manager import InverterManager
from src.ui.bom_manager import BOMManager
from src.ui.project_dashboard import ProjectDashboard
from version import get_version, get_version_info

class SolarBOMApplication:
    """Main application class for Solar eBOS BOM Generator"""
    
    def __init__(self, root):
        self.root = root
        version = get_version()
        self.root.title("Solar eBOS BOM Generator v{version}")
        
        # Current project context
        self.current_project = None
        
        # UI components
        self.main_frame = ttk.Frame(root)
        self.main_frame.pack(fill='both', expand=True)
        
        # Start with the dashboard
        self.show_dashboard()
        
        # Create menu
        self.create_menu()
        
        # Configure window size
        root.state('zoomed')
        root.minsize(800, 600)
        root.resizable(True, True)
        
    def create_menu(self):
        """Create the application menu"""
        menu = Menu(self.root)
        self.root.config(menu=menu)
        
        # File menu
        file_menu = Menu(menu, tearoff=0)
        menu.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Dashboard", command=self.show_dashboard)
        file_menu.add_command(label="New Project", command=self.new_project)
        file_menu.add_command(label="Save Project", command=self.save_project)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Help menu
        help_menu = Menu(menu, tearoff=0)
        menu.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)
    
    def show_dashboard(self):
        """Show the project dashboard"""
        # Clear main frame
        for widget in self.main_frame.winfo_children():
            widget.destroy()
        
        # Create dashboard
        dashboard = ProjectDashboard(
            self.main_frame,
            on_project_selected=self.load_project
        )
        dashboard.pack(fill='both', expand=True)
    
    def load_project(self, project):
        """Load a project and show the main application interface"""
        self.current_project = project
        
        # Update window title
        self.root.title(f"Solar eBOS BOM Generator - {project.metadata.name}")
        
        # Clear main frame
        for widget in self.main_frame.winfo_children():
            widget.destroy()
        
        # Create main notebook
        notebook = ttk.Notebook(self.main_frame)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create frames for each tab
        module_frame = ttk.Frame(notebook)
        tracker_frame = ttk.Frame(notebook)
        block_frame = ttk.Frame(notebook)
        bom_frame = ttk.Frame(notebook)
        
        # Create BOM manager tab
        bom_manager = BOMManager(bom_frame)
        bom_manager.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Create block configurator tab first (so we can reference it later)
        block_configurator = BlockConfigurator(block_frame, current_project=self.current_project)
        block_configurator.pack(fill='both', expand=True, padx=5, pady=5)

        # Function to update BOM manager with current blocks
        def update_bom_blocks():
            """Update BOM manager with current blocks and save to project"""
            # First convert blocks to a serializable format 
            serialized_blocks = {}
            
            for block_id, block in block_configurator.blocks.items():
                serialized_blocks[block_id] = block.to_dict()
            
            # Check if any blocks are missing from serialized version
            if len(serialized_blocks) != len(block_configurator.blocks):
                print("WARNING: Some blocks were not properly serialized!")
            
            # Update the BOM manager with the current blocks
            bom_manager.set_blocks(block_configurator.blocks)
            
            # Update project with current blocks
            if self.current_project:
                self.current_project.blocks = serialized_blocks
        
        # Define callback for tracker template creation
        def on_template_saved(template):
            print(f"Template saved: {template}")
            # Refresh templates in block configurator
            if hasattr(block_configurator, 'reload_templates'):
                block_configurator.reload_templates()

        # Define callback for tracker template deletion
        def on_template_deleted(template_name):
            print(f"Template deleted: {template_name}")
            # Refresh templates in block configurator
            if hasattr(block_configurator, 'reload_templates'):
                block_configurator.reload_templates()
        
        # Create tracker template creator with callback to block configurator
        tracker_creator = TrackerTemplateCreator(
            tracker_frame,
            module_spec=None,
            on_template_saved=on_template_saved,
            on_template_deleted=on_template_deleted
        )
        tracker_creator.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Define module selection callback
        def on_module_selected(module):
            # Update tracker creator
            tracker_creator.module_spec = module
            
            # Add module to project's selected modules
            if module and module.model:
                module_id = f"{module.manufacturer} {module.model}"
                if module_id not in self.current_project.selected_modules:
                    self.current_project.selected_modules.append(module_id)
            
            # Update block configurator
            block_configurator.current_module = module
        
        # Create module manager with callback
        module_manager = ModuleManager(
            module_frame,
            on_module_selected=on_module_selected
        )
        module_manager.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Add tabs to notebook
        notebook.add(module_frame, text='Modules')
        notebook.add(tracker_frame, text='Tracker Templates')
        notebook.add(block_frame, text='Block Layout')
        notebook.add(bom_frame, text='BOM Generator')
        
        # Load blocks from project if available
        if self.current_project.blocks:
            # Get templates and inverters for block reconstruction
            from src.utils.file_handlers import load_json_file
            
            # Load tracker templates
            tracker_templates = {}
            try:
                templates_data = load_json_file('data/tracker_templates.json')
                for name, template_data in templates_data.items():
                    # Extract the module spec data
                    module_data = template_data.get('module_spec', {})
                    
                    # Create proper ModuleSpec object with correct values
                    module_spec = ModuleSpec(
                        manufacturer=module_data.get('manufacturer', 'Default'),
                        model=module_data.get('model', 'Default'),
                        type=ModuleType.MONO_PERC,  # Default type
                        length_mm=module_data.get('length_mm', 2000),
                        width_mm=module_data.get('width_mm', 1000),
                        depth_mm=module_data.get('depth_mm', 40),
                        weight_kg=module_data.get('weight_kg', 25),
                        wattage=module_data.get('wattage', 400),
                        vmp=module_data.get('vmp', 40),
                        imp=module_data.get('imp', 10),
                        voc=module_data.get('voc', 48),
                        isc=module_data.get('isc', 10.5),
                        max_system_voltage=module_data.get('max_system_voltage', 1500)
                    )
                    
                    # Create TrackerTemplate with the correct module_spec
                    from src.models.tracker import TrackerTemplate, ModuleOrientation
                    tracker_templates[name] = TrackerTemplate(
                        template_name=name,
                        module_spec=module_spec,
                        module_orientation=ModuleOrientation(template_data.get('module_orientation', 'Portrait')),
                        modules_per_string=template_data.get('modules_per_string', 28),
                        strings_per_tracker=template_data.get('strings_per_tracker', 2),
                        module_spacing_m=template_data.get('module_spacing_m', 0.01),
                        motor_gap_m=template_data.get('motor_gap_m', 1.0),
                        motor_position_after_string=template_data.get('motor_position_after_string', 0)
                    )
            except Exception as e:
                print(f"Error loading templates: {str(e)}")
            
            # Load stored inverters
            inverters = {}
            try:
                inverters_data = load_json_file('data/inverters.json')
                for name, inverter_data in inverters_data.items():
                    # Create proper inverter object here (simplified)
                    from src.models.inverter import InverterSpec, MPPTChannel, MPPTConfig
                    
                    # Create MPPT channels
                    channels = []
                    for ch_data in inverter_data.get('mppt_channels', []):
                        channel = MPPTChannel(
                            max_input_current=float(ch_data.get('max_input_current', 10)),
                            min_voltage=float(ch_data.get('min_voltage', 200)),
                            max_voltage=float(ch_data.get('max_voltage', 1000)),
                            max_power=float(ch_data.get('max_power', 5000)),
                            num_string_inputs=int(ch_data.get('num_string_inputs', 2))
                        )
                        channels.append(channel)
                    
                    # Create inverter
                    inverters[name] = InverterSpec(
                        manufacturer=inverter_data.get('manufacturer', 'Unknown'),
                        model=inverter_data.get('model', 'Unknown'),
                        rated_power=float(inverter_data.get('rated_power', 10.0)),
                        max_efficiency=float(inverter_data.get('max_efficiency', 98.0)),
                        mppt_channels=channels,
                        mppt_configuration=MPPTConfig(inverter_data.get('mppt_configuration', 'Independent')),
                        max_dc_voltage=float(inverter_data.get('max_dc_voltage', 1000)),
                        startup_voltage=float(inverter_data.get('startup_voltage', 200)),
                        nominal_ac_voltage=float(inverter_data.get('nominal_ac_voltage', 400.0)),
                        max_ac_current=float(inverter_data.get('max_ac_current', 40.0)),
                        power_factor=float(inverter_data.get('power_factor', 0.99)),
                        dimensions_mm=inverter_data.get('dimensions_mm', (1000, 600, 300)),
                        weight_kg=float(inverter_data.get('weight_kg', 75.0)),
                        ip_rating=inverter_data.get('ip_rating', "IP65")
                    )
            except Exception as e:
                print(f"Error loading inverters: {str(e)}")
            
            # Reconstruct blocks
            from src.models.block import BlockConfig
            
            reconstructed_blocks = {}
            for block_id, block_data in self.current_project.blocks.items():
                try:
                    block = BlockConfig.from_dict(block_data, tracker_templates, inverters)
                    reconstructed_blocks[block_id] = block
                except Exception as e:
                    print(f"Error reconstructing block {block_id}: {str(e)}")
            
            # Set blocks in configurator
            block_configurator.blocks = reconstructed_blocks
            
            # Update block listbox
            if hasattr(block_configurator, 'block_listbox'):
                block_configurator.block_listbox.delete(0, tk.END)
                for block_id in reconstructed_blocks:
                    block_configurator.block_listbox.insert(tk.END, block_id)

            update_bom_blocks()
        
        # Add status bar with project info
        status_frame = ttk.Frame(self.main_frame, relief=tk.SUNKEN, borderwidth=1)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        project_label = ttk.Label(
            status_frame, 
            text=f"Project: {project.metadata.name} | Client: {project.metadata.client or 'N/A'}"
        )
        project_label.pack(side=tk.LEFT, padx=5, pady=2)
        
        dashboard_btn = ttk.Button(
            status_frame, text="Dashboard", command=self.show_dashboard, width=10
        )
        dashboard_btn.pack(side=tk.RIGHT, padx=5, pady=2)
        
        save_btn = ttk.Button(
            status_frame, text="Save Project", command=self.save_project, width=12
        )
        save_btn.pack(side=tk.RIGHT, padx=5, pady=2)
    
    def new_project(self):
        """Create a new project"""
        from src.ui.project_dashboard import ProjectDialog
        
        dialog = ProjectDialog(self.root, title="Create New Project")
        
        if dialog.result:
            from src.utils.project_manager import ProjectManager
            project_manager = ProjectManager()
            
            # Unpack the result tuple with the new row spacing
            if len(dialog.result) >= 6:  # Check if row spacing was included
                name, description, client, location, notes, row_spacing_m = dialog.result
            else:
                # Fallback for backward compatibility
                name, description, client, location, notes = dialog.result
                row_spacing_m = 6.0  # Default to 6m
            
            # Create the project
            project = project_manager.create_project(
                name=name,
                description=description,
                client=client,
                location=location,
                notes=notes
            )
            
            # Set default row spacing
            project.default_row_spacing_m = row_spacing_m
            
            if project_manager.save_project(project):
                messagebox.showinfo("Success", f"Project '{name}' created successfully")
                self.load_project(project)
            else:
                messagebox.showerror("Error", f"Failed to create project '{name}'")

    def show_about(self):
        """Show about dialog with version info"""
        from version import get_version_info
        messagebox.showinfo(
            "About Solar eBOS BOM Generator",
            f"Solar eBOS BOM Generator\n\n"
            f"{get_version_info()}\n\n"
            f"A tool for designing solar project layouts and generating accurate "
            f"bills of material for electrical balance of system components."
        )

    def save_project(self):
        """Save the current project"""
        if not self.current_project:
            messagebox.showinfo("No Project", "No project is currently open")
            return
            
        # Update project blocks from UI before saving
        self.update_project_blocks()
            
        from src.utils.project_manager import ProjectManager
        project_manager = ProjectManager()
        
        # Update project metadata
        self.current_project.update_modified_date()
        
        if project_manager.save_project(self.current_project):
            messagebox.showinfo("Success", "Project saved successfully")
        else:
            messagebox.showerror("Error", "Failed to save project")

    def update_project_blocks(self):
        """Update project with current blocks from UI before saving"""
        if not hasattr(self, 'current_project') or not self.current_project:
            return
            
        # Find the BlockConfigurator instance in the UI
        block_configurator = None
        for widget in self.main_frame.winfo_children():
            if isinstance(widget, ttk.Notebook):
                for tab_id in widget.tabs():
                    tab_frame = widget.nametowidget(tab_id)
                    for child in tab_frame.winfo_children():
                        if hasattr(child, 'blocks') and hasattr(child, 'draw_block'):
                            block_configurator = child
                            break
                            
        if block_configurator and hasattr(block_configurator, 'blocks'):
            # Convert blocks to serializable format
            serialized_blocks = {}
            for block_id, block in block_configurator.blocks.items():
                serialized_blocks[block_id] = block.to_dict()
                
            # Update the project's blocks
            self.current_project.blocks = serialized_blocks
    
    def show_about(self):
        """Show about dialog"""
        messagebox.showinfo(
            "About Solar eBOS BOM Generator",
            "Solar eBOS BOM Generator\n\n"
            "A tool for designing solar project layouts and generating accurate "
            "bills of material for electrical balance of system components."
        )


def main():
    root = tk.Tk()
    app = SolarBOMApplication(root)
    root.mainloop()

if __name__ == '__main__':
    main()