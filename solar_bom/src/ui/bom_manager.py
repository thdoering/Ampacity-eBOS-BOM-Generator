import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, List, Any, Optional, Callable
import pandas as pd
import os

from ..models.block import BlockConfig
from ..utils.bom_generator import BOMGenerator

class BOMManager(ttk.Frame):
    """UI component for managing BOM generation"""
    
    def __init__(self, parent, blocks: Optional[Dict[str, BlockConfig]] = None):
        super().__init__(parent)
        self.parent = parent
        self.blocks = blocks or {}
        self.selected_blocks = []  # Store selected block IDs
        
        self.setup_ui()
        self.update_block_list()
        
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
        
        # Left side - Block List and Controls
        left_column = ttk.Frame(main_container)
        left_column.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.N, tk.S, tk.W))
        
        # Block selection frame
        block_frame = ttk.LabelFrame(left_column, text="Blocks", padding="5")
        block_frame.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Block selection listbox with checkboxes
        self.block_listbox = ttk.Treeview(block_frame, columns=('block',), show='tree')
        self.block_listbox.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.block_listbox.column('#0', width=30)
        self.block_listbox.column('block', width=150)
        self.block_listbox.heading('block', text="Block ID")
        
        # Add scrollbar to listbox
        scrollbar = ttk.Scrollbar(block_frame, orient="vertical", command=self.block_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.block_listbox.configure(yscrollcommand=scrollbar.set)
        
        # Block selection buttons
        block_button_frame = ttk.Frame(block_frame)
        block_button_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5)
        
        self.select_all_var = tk.BooleanVar(value=True)
        select_all_check = ttk.Checkbutton(
            block_button_frame, 
            text="Select All", 
            variable=self.select_all_var,
            command=self.toggle_select_all
        )
        select_all_check.grid(row=0, column=0, padx=5)
        
        # BOM Action Frame
        bom_frame = ttk.LabelFrame(left_column, text="BOM Generation", padding="5")
        bom_frame.grid(row=1, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Export button
        ttk.Button(
            bom_frame, 
            text="Export BOM to Excel", 
            command=self.export_bom
        ).grid(row=0, column=0, padx=5, pady=5)
        
        # Right side - BOM Preview
        right_column = ttk.Frame(main_container)
        right_column.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.N, tk.S, tk.E, tk.W))
        
        # Preview Frame
        preview_frame = ttk.LabelFrame(right_column, text="BOM Preview", padding="5")
        preview_frame.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Create treeview for BOM preview
        self.preview_tree = ttk.Treeview(
            preview_frame, 
            columns=('component', 'description', 'quantity', 'unit'),
            show='headings'
        )
        self.preview_tree.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure columns
        self.preview_tree.column('component', width=150)
        self.preview_tree.column('description', width=200)
        self.preview_tree.column('quantity', width=100)
        self.preview_tree.column('unit', width=80)
        
        # Add headings
        self.preview_tree.heading('component', text="Component Type")
        self.preview_tree.heading('description', text="Description")
        self.preview_tree.heading('quantity', text="Quantity")
        self.preview_tree.heading('unit', text="Unit")
        
        # Add scrollbar to preview
        preview_scrollbar = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_tree.yview)
        preview_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.preview_tree.configure(yscrollcommand=preview_scrollbar.set)
        
        # Make treeview expandable
        preview_frame.grid_columnconfigure(0, weight=1)
        preview_frame.grid_rowconfigure(0, weight=1)
        
        # Refresh button
        ttk.Button(
            preview_frame, 
            text="Refresh Preview", 
            command=self.update_preview
        ).grid(row=1, column=0, padx=5, pady=5)
    
    def set_blocks(self, blocks: Dict[str, BlockConfig]):
        """
        Set the blocks dictionary
        
        Args:
            blocks: Dictionary of block configurations (id -> BlockConfig)
        """
        self.blocks = blocks
        self.update_block_list()
    
    def update_block_list(self):
        """Update the block listbox with current blocks"""
        # Clear existing items
        for item in self.block_listbox.get_children():
            self.block_listbox.delete(item)
        
        self.selected_blocks = []
        
        # Sort block IDs for consistent order
        sorted_block_ids = sorted(self.blocks.keys())
        
        # Add blocks to listbox
        for block_id in sorted_block_ids:
            # Add block to listbox with checkbox
            self.block_listbox.insert(
                '', 'end', 
                values=(block_id,),
                tags=('checked',)
            )
            self.selected_blocks.append(block_id)
        
        # Update preview
        self.update_preview()
    
    def toggle_select_all(self):
        """Toggle selection of all blocks"""
        if self.select_all_var.get():
            # Select all blocks
            self.selected_blocks = list(self.blocks.keys())
            for item in self.block_listbox.get_children():
                self.block_listbox.item(item, tags=('checked',))
        else:
            # Deselect all blocks
            self.selected_blocks = []
            for item in self.block_listbox.get_children():
                self.block_listbox.item(item, tags=('unchecked',))
        
        # Update preview
        self.update_preview()
    
    def update_preview(self):
        """Update the BOM preview"""
        # Clear existing items
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        
        # Get selected blocks
        selected_blocks = {block_id: self.blocks[block_id] for block_id in self.selected_blocks if block_id in self.blocks}
        
        if not selected_blocks:
            return
        
        # Generate BOM
        bom_generator = BOMGenerator(selected_blocks)
        quantities = bom_generator.calculate_cable_quantities()
        summary_data = bom_generator.generate_summary_data(quantities)
        
        # Add summary data to preview
        for _, row in summary_data.iterrows():
            # Format quantity based on unit
            quantity = row['Quantity']
            if row['Unit'] == 'feet':
                quantity_str = f"{quantity:.1f}"
            else:
                quantity_str = f"{int(quantity)}"
                
            self.preview_tree.insert(
                '', 'end',
                values=(
                    row['Component Type'],
                    row['Description'],
                    quantity_str,
                    row['Unit']
                )
            )
    
    def export_bom(self):
        """Export BOM to Excel file"""
        # Get selected blocks
        selected_blocks = {block_id: self.blocks[block_id] for block_id in self.selected_blocks if block_id in self.blocks}
        
        if not selected_blocks:
            messagebox.showwarning("Warning", "No blocks selected for BOM export")
            return
        
        # Check for blocks without wiring configuration
        blocks_without_wiring = []
        for block_id, block in selected_blocks.items():
            if not hasattr(block, 'wiring_config') or not block.wiring_config:
                blocks_without_wiring.append(block_id)
        
        if blocks_without_wiring:
            message = "The following selected blocks have no wiring configuration:\n"
            message += "\n".join(sorted(blocks_without_wiring))
            message += "\n\nThese blocks will have limited BOM output. Do you want to continue?"
            if not messagebox.askyesno("Missing Wiring Configurations", message, icon='warning'):
                return
        
        # Ask for export location
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Export BOM to Excel"
        )
        
        if not filepath:
            return
        
        # Get project information if available
        project_info = None
        try:
            # Try different approaches to find the current project
            main_app = None
            
            # First, try to get from the root window
            root = self.winfo_toplevel()
            if hasattr(root, 'current_project') and root.current_project:
                main_app = root
            
            # Look up the widget hierarchy to find the main application
            if not main_app:
                widget = self
                while widget:
                    if hasattr(widget, 'current_project') and widget.current_project:
                        main_app = widget
                        break
                    widget = widget.master
            
            # If we found the main app with a project, get the info
            if main_app and main_app.current_project:
                project = main_app.current_project
                
                # Get system size and module counts from blocks
                system_size = 0
                total_modules = 0
                module_manufacturer = set()
                module_model = set()
                inverter_manufacturer = set()
                inverter_model = set()
                dc_collection_types = set()
                
                for block_id, block in selected_blocks.items():
                    # Count modules
                    if block.tracker_template and block.tracker_template.module_spec:
                        module_spec = block.tracker_template.module_spec
                        modules_per_tracker = block.tracker_template.get_total_modules()
                        tracker_count = len(block.tracker_positions)
                        block_modules = modules_per_tracker * tracker_count
                        
                        total_modules += block_modules
                        
                        # Add module info
                        module_manufacturer.add(module_spec.manufacturer)
                        module_model.add(module_spec.model)
                        
                        # Calculate system size
                        system_size += (block_modules * module_spec.wattage) / 1000
                    
                    # Add inverter info
                    if block.inverter:
                        inverter_manufacturer.add(block.inverter.manufacturer)
                        inverter_model.add(block.inverter.model)
                    
                    # Add DC collection type
                    if block.wiring_config:
                        dc_collection_types.add(block.wiring_config.wiring_type.value)
                
                # Create the project info dictionary
                project_info = {
                    'Project Name': project.metadata.name,
                    'Customer': project.metadata.client or 'Unknown',
                    'Location': project.metadata.location or 'Unknown',
                    'System Size (kW DC)': round(system_size, 2),
                    'Number of Modules': total_modules,
                    'Module Manufacturer': ', '.join(module_manufacturer) if module_manufacturer else 'Unknown',
                    'Module Model': ', '.join(module_model) if module_model else 'Unknown',
                    'Inverter Manufacturer': ', '.join(inverter_manufacturer) if inverter_manufacturer else 'Unknown',
                    'Inverter Model': ', '.join(inverter_model) if inverter_model else 'Unknown',
                    'DC Collection': ', '.join(dc_collection_types) if dc_collection_types else 'Unknown',
                    'Description': project.metadata.description or '',
                    'Notes': project.metadata.notes or '',
                    'Blocks Without Wiring': len(blocks_without_wiring) if blocks_without_wiring else 0
                }
                
                print("Project info for BOM:", project_info)
        except Exception as e:
            print(f"Error getting project info: {str(e)}")
        
        # Generate and export BOM
        try:
            bom_generator = BOMGenerator(selected_blocks)
            success = bom_generator.export_bom_to_excel(filepath, project_info)
            
            if success:
                messagebox.showinfo("Success", f"BOM exported successfully to {filepath}")
            else:
                messagebox.showerror("Error", "Failed to export BOM")
        except PermissionError:
            messagebox.showerror(
                "Permission Error", 
                f"Cannot write to {filepath}.\n\n"
                "This error usually occurs when:\n"
                "• The file is already open in Excel or another program\n"
                "• You don't have permission to write to this location\n\n"
                "Please close any programs that might be using this file, or choose a different filename/location."
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export BOM: {str(e)}")