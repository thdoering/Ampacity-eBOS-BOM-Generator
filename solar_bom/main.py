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


class SolarBOMApplication:
    """Main application class for Solar eBOS BOM Generator"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Solar eBOS BOM Generator")
        
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
        
        # Create module manager tab
        module_frame = ttk.Frame(notebook)
        notebook.add(module_frame, text='Modules')
        
        # Create tracker template creator tab
        tracker_frame = ttk.Frame(notebook)
        tracker_creator = TrackerTemplateCreator(
            tracker_frame,
            module_spec=None,
            on_template_saved=lambda template: print(f"Template saved: {template}")
        )
        tracker_creator.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Create module manager with reference to tracker creator
        def on_module_selected(module):
            tracker_creator.module_spec = module
            # Add module to project's selected modules
            if module and module.model:
                module_id = f"{module.manufacturer} {module.model}"
                if module_id not in self.current_project.selected_modules:
                    self.current_project.selected_modules.append(module_id)
        
        module_manager = ModuleManager(
            module_frame,
            on_module_selected=on_module_selected
        )
        module_manager.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Add tracker frame to notebook
        notebook.add(tracker_frame, text='Tracker Templates')
        
        # Create block configurator tab
        block_frame = ttk.Frame(notebook)
        notebook.add(block_frame, text='Block Layout')
        
        block_configurator = BlockConfigurator(block_frame)
        block_configurator.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Load blocks from project if available
        if self.current_project.blocks:
            # TODO: Implement loading blocks from project
            pass
        
        # Create BOM manager tab
        bom_frame = ttk.Frame(notebook)
        notebook.add(bom_frame, text='BOM Generator')
        
        bom_manager = BOMManager(bom_frame)
        bom_manager.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Connect module manager to block configurator
        def on_module_selected_for_block(module):
            block_configurator.current_module = module
        
        # Function to update BOM manager with current blocks
        def update_bom_blocks():
            bom_manager.set_blocks(block_configurator.blocks)
            
            # Update project with current blocks
            if self.current_project:
                # Store block references in project
                self.current_project.blocks = {
                    block_id: block_id for block_id in block_configurator.blocks.keys()
                }
        
        module_manager.on_module_selected = lambda module: (
            on_module_selected(module),
            on_module_selected_for_block(module)
        )
        
        # Connect block configurator to BOM manager
        block_configurator.on_blocks_changed = update_bom_blocks
        
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
            
            name, description, client, location, notes = dialog.result
            
            # Create the project
            project = project_manager.create_project(
                name=name,
                description=description,
                client=client,
                location=location,
                notes=notes
            )
            
            if project_manager.save_project(project):
                messagebox.showinfo("Success", f"Project '{name}' created successfully")
                self.load_project(project)
            else:
                messagebox.showerror("Error", f"Failed to create project '{name}'")
    
    def save_project(self):
        """Save the current project"""
        if not self.current_project:
            messagebox.showinfo("No Project", "No project is currently open")
            return
            
        from src.utils.project_manager import ProjectManager
        project_manager = ProjectManager()
        
        # Update project metadata
        self.current_project.update_modified_date()
        
        if project_manager.save_project(self.current_project):
            messagebox.showinfo("Success", "Project saved successfully")
        else:
            messagebox.showerror("Error", "Failed to save project")
    
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