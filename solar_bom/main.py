import tkinter as tk
from tkinter import ttk, messagebox, Menu, filedialog
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
        file_menu.add_command(label="Apply Recommended Cable Sizes (All Blocks)", command=self.apply_recommended_sizes_all_blocks)
        file_menu.add_separator()
        file_menu.add_command(label="Export Project...", command=self.export_project)
        file_menu.add_command(label="Import Project...", command=self.import_project)
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

    def refresh_all_template_views(self):
        """Refresh template views in all components when project changes"""
        # This method can be called when enabled_templates list changes
        # Update any UI components that show templates
        pass
    
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
        device_frame = ttk.Frame(notebook)
        bom_frame = ttk.Frame(notebook)
        
        # Create Device Configurator tab
        from src.ui.device_configurator import DeviceConfigurator
        device_configurator = DeviceConfigurator(device_frame, project_manager=None)
        device_configurator.pack(fill='both', expand=True, padx=5, pady=5)
        # Add reference to main app
        device_configurator.main_app = self
        # Store reference for other components to access
        self.device_configurator = device_configurator
        
        # Create BOM manager tab - pass reference to main app
        bom_manager = BOMManager(bom_frame, main_app=self)
        bom_manager.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Create block configurator tab first (so we can reference it later)
        block_configurator = BlockConfigurator(block_frame, current_project=self.current_project, on_autosave=self.autosave_project)
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
            
            # UPDATE THE BOM MANAGER WITH THE ACTUAL BLOCKCONFIG OBJECTS
            bom_manager.set_blocks(block_configurator.blocks)
            
            # Update device configurator
            if self.current_project:
                self.current_project.blocks = block_configurator.blocks
                device_configurator.load_project(self.current_project)
            
            # Store serialized blocks in project for saving
            if hasattr(self, 'current_project') and self.current_project:
                self.current_project.blocks = serialized_blocks
        
        # Define callback for tracker template creation
        def on_template_saved(template):
            print(f"Template saved: {template}")
            # Refresh templates in block configurator (this will reload filtered templates)
            if hasattr(block_configurator, 'reload_templates'):
                block_configurator.reload_templates()

        # Define callback for tracker template deletion
        def on_template_deleted(template_name):
            print(f"Template deleted: {template_name}")
            # Remove from enabled templates if present
            if self.current_project and template_name in self.current_project.enabled_templates:
                self.current_project.enabled_templates.remove(template_name)
            # Refresh templates in block configurator
            if hasattr(block_configurator, 'reload_templates'):
                block_configurator.reload_templates()

        def on_template_enabled_changed():
            """Called when template enabled status changes in tracker creator"""
            # Refresh templates in block configurator to show only enabled ones
            if hasattr(block_configurator, 'reload_templates'):
                block_configurator.reload_templates()
            # Mark project as modified
            if self.current_project:
                self.current_project.update_modified_date()
        
        # Create tracker template creator with callback to block configurator
        tracker_creator = TrackerTemplateCreator(
            tracker_frame,
            module_spec=None,
            on_template_saved=on_template_saved,
            on_template_deleted=on_template_deleted,
            current_project=self.current_project,
            on_template_enabled_changed=on_template_enabled_changed
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
        notebook.add(device_frame, text='Configure Device')
        notebook.add(bom_frame, text='BOM Generator')
        
        # Load blocks from project if available
        if self.current_project.blocks:
            # Get templates and inverters for block reconstruction
            from src.utils.file_handlers import load_json_file
            
            # Load tracker templates
            tracker_templates = {}
            try:
                templates_data = load_json_file('data/tracker_templates.json')
                
                if templates_data:
                    # Check if this is the new hierarchical format
                    first_value = next(iter(templates_data.values()))
                    if isinstance(first_value, dict) and not any(key in first_value for key in ['module_orientation', 'modules_per_string']):
                        # New hierarchical format: Manufacturer -> Template Name -> template_data
                        for manufacturer, template_group in templates_data.items():
                            for template_name, template_data in template_group.items():
                                # Use manufacturer prefix to make template names unique
                                unique_name = f"{manufacturer} - {template_name}"
                                
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
                                
                                # Calculate appropriate default split values
                                modules_per_string = template_data.get('modules_per_string', 28)
                                default_split_north = modules_per_string // 2
                                default_split_south = modules_per_string - default_split_north
                                
                                tracker_templates[unique_name] = TrackerTemplate(
                                    template_name=unique_name,
                                    module_spec=module_spec,
                                    module_orientation=ModuleOrientation(template_data.get('module_orientation', ModuleOrientation.PORTRAIT.value)),
                                    modules_per_string=modules_per_string,
                                    strings_per_tracker=template_data.get('strings_per_tracker', 2),
                                    module_spacing_m=template_data.get('module_spacing_m', 0.01),
                                    motor_gap_m=template_data.get('motor_gap_m', 1.0),
                                    motor_position_after_string=template_data.get('motor_position_after_string', 0),
                                    # New motor placement fields with calculated defaults
                                    motor_placement_type=template_data.get('motor_placement_type', 'between_strings'),
                                    motor_string_index=template_data.get('motor_string_index', 1),
                                    motor_split_north=template_data.get('motor_split_north', default_split_north),
                                    motor_split_south=template_data.get('motor_split_south', default_split_south),
                                    # Multi-module-high configuration
                                    modules_high=template_data.get('modules_high', 1)
                                )
                                
                                # Also store with the old name for backwards compatibility
                                tracker_templates[template_name] = tracker_templates[unique_name]
                    else:
                        # Old flat format
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
                            
                            # Calculate appropriate default split values
                            modules_per_string = template_data.get('modules_per_string', 28)
                            default_split_north = modules_per_string // 2
                            default_split_south = modules_per_string - default_split_north
                            
                            tracker_templates[name] = TrackerTemplate(
                                template_name=name,
                                module_spec=module_spec,
                                module_orientation=ModuleOrientation(template_data.get('module_orientation', ModuleOrientation.PORTRAIT.value)),
                                modules_per_string=modules_per_string,
                                strings_per_tracker=template_data.get('strings_per_tracker', 2),
                                module_spacing_m=template_data.get('module_spacing_m', 0.01),
                                motor_gap_m=template_data.get('motor_gap_m', 1.0),
                                motor_position_after_string=template_data.get('motor_position_after_string', 0),
                                # New motor placement fields with calculated defaults
                                motor_placement_type=template_data.get('motor_placement_type', 'between_strings'),
                                motor_string_index=template_data.get('motor_string_index', 1),
                                motor_split_north=template_data.get('motor_split_north', default_split_north),
                                motor_split_south=template_data.get('motor_split_south', default_split_south),
                                # Multi-module-high configuration
                                modules_high=template_data.get('modules_high', 1)
                            )

            except Exception as e:
                print(f"Error loading templates: {str(e)}")
                        
            # Load stored inverters
            inverters = {}
            try:
                inverters_data = load_json_file('data/inverters.json')
                # Add this check to make sure we actually got data
                if inverters_data:
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
                else:
                    print("No inverters.json file found or empty - using empty inverter list")
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

            # Update device configurator with the loaded project and blocks
            self.current_project.blocks = reconstructed_blocks
            device_configurator.load_project(self.current_project)
            
            # Update block listbox
            if hasattr(block_configurator, 'block_listbox'):
                block_configurator.block_listbox.delete(0, tk.END)
                for block_id in reconstructed_blocks:
                    block_configurator.block_listbox.insert(tk.END, block_id)
            
            # Update the project reference and UI
            if hasattr(block_configurator, 'set_project'):
                block_configurator.set_project(self.current_project)

            update_bom_blocks()
        
        # Set up the callback to keep blocks synchronized with project
        # ADD THIS LINE HERE (outside and after the function definition):
        block_configurator.on_blocks_changed = update_bom_blocks

        # Store reference for menu actions
        self.block_configurator = block_configurator
        
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
            
            # Unpack the result tuple (no row spacing)
            name, description, client, location, notes = dialog.result
            
            # Create the project
            project = project_manager.create_project(
                name=name,
                description=description,
                client=client,
                location=location,
                notes=notes
            )
            
            # Don't set default row spacing here - it will be set by the first block
            
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

    def apply_recommended_sizes_all_blocks(self):
        """Apply recommended cable sizes to all harnesses in all blocks"""
        if not self.current_project:
            messagebox.showwarning("No Project", "Please load a project first.")
            return
        
        if not hasattr(self, 'block_configurator') or not self.block_configurator:
            messagebox.showwarning("Error", "Block configurator not available. Please reload the project.")
            return
        
        from src.utils.cable_sizing import calculate_all_cable_sizes
        from src.models.block import WiringType
        
        # Get NEC factor from project
        nec_factor = getattr(self.current_project, 'nec_safety_factor', 1.56)
        
        blocks_updated = 0
        harnesses_updated = 0
        
        # Update the live BlockConfig objects in block_configurator
        for block_id, block in self.block_configurator.blocks.items():
            if not hasattr(block, 'wiring_config') or not block.wiring_config:
                continue
            
            if block.wiring_config.wiring_type != WiringType.HARNESS:
                continue
            
            if not hasattr(block.wiring_config, 'harness_groupings') or not block.wiring_config.harness_groupings:
                continue
            
            # Get module Isc from block's tracker template
            module_isc = 10.0  # Default fallback
            if (block.tracker_template and 
                hasattr(block.tracker_template, 'module_spec') and 
                block.tracker_template.module_spec):
                module_isc = block.tracker_template.module_spec.isc
            
            block_changed = False
            
            for string_count, harness_list in block.wiring_config.harness_groupings.items():
                for harness in harness_list:
                    num_strings = len(harness.string_indices)
                    
                    # Calculate recommended sizes
                    recommended = calculate_all_cable_sizes(num_strings, module_isc, nec_factor)
                    
                    # Apply recommended sizes
                    harness.string_cable_size = recommended['string']
                    harness.cable_size = recommended['harness']
                    harness.extender_cable_size = recommended['extender']
                    harness.whip_cable_size = recommended['whip']
                    
                    harnesses_updated += 1
                    block_changed = True
            
            if block_changed:
                blocks_updated += 1
        
        # Update the serialized blocks in project for saving
        serialized_blocks = {}
        for block_id, block in self.block_configurator.blocks.items():
            serialized_blocks[block_id] = block.to_dict()
        self.current_project.blocks = serialized_blocks
        
        # Save the project
        if blocks_updated > 0:
            self.save_project()
            messagebox.showinfo(
                "Cable Sizes Updated", 
                f"Applied recommended cable sizes to {harnesses_updated} harness(es) in {blocks_updated} block(s).\n\nProject saved.\n\nSelect a block to see the updated values."
            )
        else:
            messagebox.showinfo("No Changes", "No harness configurations found to update.\n\nMake sure blocks have Wire Harness wiring type configured with harness groupings.")


    def export_project(self):
        """Export current project to a .ebom file for sharing"""
        if not self.current_project:
            messagebox.showwarning("No Project", "No project is currently open to export.")
            return
        
        # Suggest filename based on project name
        suggested_name = self.current_project.metadata.name.replace(' ', '_')
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".ebom",
            filetypes=[("eBOM Project Files", "*.ebom"), ("All Files", "*.*")],
            title="Export Project",
            initialfile=suggested_name
        )
        
        if not filepath:
            return  # User cancelled
        
        try:
            import json
            
            # Update modified date before export
            self.current_project.update_modified_date()
            
            with open(filepath, 'w') as f:
                json.dump(self.current_project.to_dict(), f, indent=2)
            
            messagebox.showinfo("Success", f"Project exported successfully to:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export project:\n{str(e)}")
    
    def import_project(self):
        """Import a .ebom project file and copy it to local projects folder"""
        filepath = filedialog.askopenfilename(
            filetypes=[("eBOM Project Files", "*.ebom"), ("All Files", "*.*")],
            title="Import Project"
        )
        
        if not filepath:
            return  # User cancelled
        
        try:
            import json
            import os
            from src.models.project import Project
            from src.utils.project_manager import ProjectManager
            
            # Load the project from the external file
            with open(filepath, 'r') as f:
                data = json.load(f)
            
            project = Project.from_dict(data)
            
            # Check if a project with this name already exists
            project_manager = ProjectManager()
            original_name = project.metadata.name
            
            if project_manager.project_name_exists(original_name):
                # Ask user what to do
                result = messagebox.askyesnocancel(
                    "Project Exists",
                    f"A project named '{original_name}' already exists.\n\n"
                    "Yes = Overwrite existing project\n"
                    "No = Import as copy with new name\n"
                    "Cancel = Abort import"
                )
                
                if result is None:  # Cancel
                    return
                elif result is False:  # No - create with new name
                    # Generate a unique name
                    counter = 1
                    new_name = f"{original_name} (Imported)"
                    while project_manager.project_name_exists(new_name):
                        counter += 1
                        new_name = f"{original_name} (Imported {counter})"
                    project.metadata.name = new_name
                # If result is True (Yes), we'll overwrite
            
            # Save to local projects folder
            if project_manager.save_project(project):
                messagebox.showinfo(
                    "Success", 
                    f"Project '{project.metadata.name}' imported successfully!"
                )
                # Open the imported project
                self.load_project(project)
            else:
                messagebox.showerror("Error", "Failed to save imported project to local folder.")
                
        except json.JSONDecodeError:
            messagebox.showerror("Import Error", "The selected file is not a valid project file.")
        except KeyError as e:
            messagebox.showerror("Import Error", f"The project file is missing required data: {e}")
        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import project:\n{str(e)}")


    def autosave_project(self):
        """Save the current project silently (for autosave functionality)"""
        if not self.current_project:
            return
            
        # Update project blocks from UI before saving
        self.update_project_blocks()
            
        from src.utils.project_manager import ProjectManager
        project_manager = ProjectManager()
        
        # Update project metadata
        self.current_project.update_modified_date()
        
        # Save without showing any messages
        project_manager.save_project(self.current_project)

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
                block_dict = block.to_dict()
                serialized_blocks[block_id] = block_dict
                
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