import tkinter as tk
from tkinter import ttk
from src.models.module import ModuleSpec, ModuleType
from src.ui.tracker_creator import TrackerTemplateCreator
from src.ui.module_manager import ModuleManager
from src.ui.block_configurator import BlockConfigurator
from src.ui.inverter_manager import InverterManager


def main():
    # Create root window
    root = tk.Tk()
    root.title("Solar eBOS BOM Generator")
    
    # Create main notebook for tabs
    notebook = ttk.Notebook(root)
    notebook.pack(fill='both', expand=True, padx=10, pady=10)

    # Create module manager tab
    module_frame = ttk.Frame(notebook)
    notebook.add(module_frame, text='Modules')

    # Create tracker template creator tab first (but don't add to notebook yet)
    tracker_frame = ttk.Frame(notebook)
    tracker_creator = TrackerTemplateCreator(
        tracker_frame,
        module_spec=None,  # Start with no module
        on_template_saved=lambda template: print(f"Template saved: {template}")
    )
    tracker_creator.pack(fill='both', expand=True, padx=5, pady=5)

    # Now create module manager with reference to tracker creator
    def on_module_selected(module):
        tracker_creator.module_spec = module

    module_manager = ModuleManager(
        module_frame,
        on_module_selected=on_module_selected
    )
    module_manager.pack(fill='both', expand=True, padx=5, pady=5)

    # Now add tracker frame to notebook
    notebook.add(tracker_frame, text='Tracker Templates')

    # Create block configurator tab
    block_frame = ttk.Frame(notebook)
    notebook.add(block_frame, text='Block Layout')

    block_configurator = BlockConfigurator(block_frame)
    block_configurator.pack(fill='both', expand=True, padx=5, pady=5)

    # Connect module manager to block configurator
    def on_module_selected_for_block(module):
        block_configurator.current_module = module  # We'll add this property

    module_manager.on_module_selected = lambda module: (
        on_module_selected(module),  # Original tracker creator connection
        on_module_selected_for_block(module)  # New block configurator connection
    )
    
    # Configure window size and center on screen
    root.state('zoomed')
    
    # Make window resizable
    root.minsize(800, 600)
    root.resizable(True, True)
    
    # Start the application
    root.mainloop()

if __name__ == '__main__':
    main()