import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, List, Any, Optional, Callable
import pandas as pd
import os

from ..models.block import BlockConfig
from ..utils.bom_generator import BOMGenerator
from .harness_catalog_dialog import HarnessCatalogDialog
from .harness_designer import HarnessDesigner

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
        main_container.grid_columnconfigure(0, weight=1)  # Left column gets some weight
        main_container.grid_columnconfigure(1, weight=4)  # Right column gets much more weight (4:1 ratio)
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
        
        # Harness designer button
        ttk.Button(
            bom_frame, 
            text="Harness Designer", 
            command=self.open_harness_designer
        ).grid(row=1, column=0, padx=5, pady=5)
        
        # Harness drawings button
        ttk.Button(
            bom_frame, 
            text="Generate Harness Drawings", 
            command=self.generate_harness_drawings
        ).grid(row=2, column=0, padx=5, pady=5)
        
        # Right side - BOM Preview
        right_column = ttk.Frame(main_container)
        right_column.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.N, tk.S, tk.E, tk.W))
        
        # Preview Frame
        preview_frame = ttk.LabelFrame(right_column, text="BOM Preview", padding="5")
        preview_frame.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Create treeview for BOM preview with checkbox and part number columns
        self.preview_tree = ttk.Treeview(
            preview_frame, 
            columns=('include', 'component', 'part_number', 'description', 'quantity', 'unit'),
            show='headings',
            height=25  # Make it much taller
        )
        self.preview_tree.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure tags for styling
        self.preview_tree.tag_configure('checked', foreground='black')
        self.preview_tree.tag_configure('unchecked', foreground='gray60')

        # Configure columns
        self.preview_tree.column('include', width=60, anchor='center')
        self.preview_tree.column('component', width=180)
        self.preview_tree.column('part_number', width=150, anchor='center')
        self.preview_tree.column('description', width=300)  # Wider for better readability
        self.preview_tree.column('quantity', width=100, anchor='center')
        self.preview_tree.column('unit', width=80, anchor='center')

        # Add headings
        self.preview_tree.heading('include', text="Include")
        self.preview_tree.heading('component', text="Component Type")
        self.preview_tree.heading('part_number', text="Part Number")
        self.preview_tree.heading('description', text="Description")
        self.preview_tree.heading('quantity', text="Quantity")
        self.preview_tree.heading('unit', text="Unit")

        # Bind click events for checkbox functionality
        self.preview_tree.bind('<Button-1>', self.on_tree_click)
        self.preview_tree.bind('<Double-1>', self.toggle_item_selection)

        # Track checked items
        self.checked_items = set()
        
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

        # Button frame for BOM actions
        button_frame = ttk.Frame(preview_frame)
        button_frame.grid(row=1, column=0, columnspan=2, pady=5)

        # Refresh button
        ttk.Button(
            button_frame, 
            text="Refresh Preview", 
            command=self.update_preview
        ).grid(row=0, column=0, padx=5)

        # Select All button
        ttk.Button(
            button_frame, 
            text="Select All", 
            command=self.select_all_items
        ).grid(row=0, column=1, padx=5)

        # Select None button
        ttk.Button(
            button_frame, 
            text="Select None", 
            command=self.select_none_items
        ).grid(row=0, column=2, padx=5)
    
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
        """Update BOM preview"""
        # Clear existing preview and checked items
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        
        self.checked_items.clear()
        
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
                
            # Get part number for harnesses and fuses
            part_number = self.get_part_number_for_component(row, selected_blocks)

            item = self.preview_tree.insert(
                '', 'end',
                values=(
                    '☑',  # Default to checked
                    row['Component Type'],
                    part_number,
                    row['Description'],
                    quantity_str,
                    row['Unit']
                ),
                tags=('checked',)
            )
            self.checked_items.add(item)
    
    def export_bom(self):
        """Export BOM to Excel file"""
        # Get selected blocks
        selected_blocks = {block_id: self.blocks[block_id] for block_id in self.selected_blocks if block_id in self.blocks}
        
        if not selected_blocks:
            messagebox.showwarning("Warning", "No blocks selected for BOM export")
            return
        
        # Check if any blocks are using conceptual routing
        conceptual_blocks = []
        for block_id, block in selected_blocks.items():
            if (hasattr(block, 'wiring_config') and 
                block.wiring_config and 
                hasattr(block.wiring_config, 'routing_mode') and
                getattr(block.wiring_config, 'routing_mode', 'realistic') == 'conceptual'):
                conceptual_blocks.append(block_id)

        if conceptual_blocks:
            message = "WARNING: The following blocks are using conceptual wiring routing:\n"
            message += "\n".join(sorted(conceptual_blocks))
            message += "\n\nConceptual routing may not accurately represent actual cable lengths."
            message += "\nFor accurate BOM calculations, use realistic routing in the wiring configurator."
            message += "\n\nDo you want to continue with BOM export?"
            if not messagebox.askyesno("Conceptual Routing Warning", message, icon='warning'):
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
        
        # Get checked components for filtering
        checked_components = self.get_checked_components()

        # Generate and export BOM
        try:
            bom_generator = BOMGenerator(selected_blocks)
            success = bom_generator.export_bom_to_excel(filepath, project_info, checked_components)
            
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

    def get_checked_components(self):
        """Get list of components that are checked for export"""
        checked_components = []
        
        for item in self.checked_items:
            values = self.preview_tree.item(item, 'values')
            if len(values) >= 6:  # Make sure we have all columns
                component_info = {
                    'component_type': values[1],
                    'part_number': values[2], 
                    'description': values[3],
                    'quantity': values[4],
                    'unit': values[5]
                }
                checked_components.append(component_info)
        
        return checked_components

    def on_tree_click(self, event):
        """Handle tree click events for checkbox functionality"""
        item = self.preview_tree.identify_row(event.y)
        column = self.preview_tree.identify_column(event.x)
        
        # If clicked on the include column, toggle selection
        if item and column == '#1':  # First column is include
            self.toggle_item_selection_by_item(item)

    def toggle_item_selection(self, event):
        """Toggle selection on double-click"""
        item = self.preview_tree.focus()
        if item:
            self.toggle_item_selection_by_item(item)

    def toggle_item_selection_by_item(self, item):
        """Toggle the selection state of a tree item"""
        current_values = list(self.preview_tree.item(item, 'values'))
        
        if current_values[0] == '☐':  # Currently unchecked
            current_values[0] = '☑'  # Check it
            self.preview_tree.item(item, values=current_values, tags=('checked',))
            self.checked_items.add(item)
        else:  # Currently checked
            current_values[0] = '☐'  # Uncheck it
            self.preview_tree.item(item, values=current_values, tags=('unchecked',))
            self.checked_items.discard(item)

    def select_all_items(self):
        """Select all items"""
        for item in self.preview_tree.get_children():
            current_values = list(self.preview_tree.item(item, 'values'))
            current_values[0] = '☑'  # Check it
            self.preview_tree.item(item, values=current_values, tags=('checked',))
            self.checked_items.add(item)

    def select_none_items(self):
        """Deselect all items"""
        for item in self.preview_tree.get_children():
            current_values = list(self.preview_tree.item(item, 'values'))
            current_values[0] = '☐'  # Uncheck it
            self.preview_tree.item(item, values=current_values, tags=('unchecked',))
            self.checked_items.discard(item)
        
        self.checked_items.clear()

    def get_part_number_for_component(self, row, selected_blocks):
        """Get part number for component based on type"""
        component_type = row['Component Type']
        
        # print(f"Getting part number for: {component_type}")  # Comment out for cleaner console
        
        # Handle harnesses
        if 'Harness' in component_type:
            part_num = self.get_harness_part_number(row, selected_blocks)
            # print(f"Harness part number: {part_num}")  # Comment out for cleaner console
            return part_num
        
        # Handle fuses
        elif 'Fuse' in component_type:
            part_num = self.get_fuse_part_number(row)
            # print(f"Fuse part number: {part_num}")  # Comment out for cleaner console
            return part_num

        # Handle whip cable segments
        elif 'Whip Cable Segment' in component_type:
            part_num = self.get_whip_segment_part_number(row, selected_blocks)
            return part_num

        # Handle extender cable segments  
        elif 'Extender Cable Segment' in component_type:
            part_num = self.get_extender_segment_part_number(row, selected_blocks)
            return part_num

        # Other components don't have part numbers yet
        return "N/A"

    def get_harness_part_number(self, row, selected_blocks):
        """Get matching harness part number from library"""
        try:
            from ..utils.bom_generator import BOMGenerator
            bom_generator = BOMGenerator(selected_blocks)
            
            # Extract harness info from description
            description = row['Description']
            
            # Determine polarity and string count from description
            polarity = 'positive' if 'Positive' in description else 'negative'
            
            # Extract string count (look for patterns like "2 String", "3 String", etc.)
            import re
            string_match = re.search(r'(\d+)-String', description)
            if not string_match:
                return "N/A"
            
            num_strings = int(string_match.group(1))
            
            # Get module specs from the first block to calculate spacing
            if not selected_blocks:
                return "N/A"
                
            first_block = next(iter(selected_blocks.values()))
            if not first_block.tracker_template or not first_block.tracker_template.module_spec:
                return "N/A"
            
            module_spec = first_block.tracker_template.module_spec
            modules_per_string = first_block.tracker_template.modules_per_string
            module_spacing_m = first_block.tracker_template.module_spacing_m
            
            # Calculate string spacing in feet
            string_spacing_ft = bom_generator.calculate_string_spacing_ft(
                modules_per_string, module_spec.width_mm, module_spacing_m
            )
            
            # Get trunk cable size from wiring config
            trunk_cable_size = getattr(first_block.wiring_config, 'harness_cable_size', '8 AWG')

            # Find matching harness
            return bom_generator.find_matching_harness_part_number(
                num_strings, polarity, string_spacing_ft, trunk_cable_size
            )
            
        except Exception as e:
            print(f"Error getting harness part number: {e}")
            return "N/A"

    def get_fuse_part_number(self, row):
        """Get fuse part number from fuse library"""
        try:
            # Load fuse library directly
            import json
            import os
            
            # Get fuse library path
            current_dir = os.path.dirname(os.path.abspath(__file__))  # src/ui/
            project_root = os.path.dirname(os.path.dirname(current_dir))  # Go up two levels
            library_path = os.path.join(project_root, 'data', 'fuse_library.json')
            
            with open(library_path, 'r') as f:
                fuse_library = json.load(f)
            
            # Extract fuse rating from description
            import re
            rating_match = re.search(r'(\d+)A', row['Description'])
            if not rating_match:
                print(f"No fuse rating found in description: {row['Description']}")
                return "N/A"
            
            fuse_rating_amps = int(rating_match.group(1))
            print(f"Looking for fuse with {fuse_rating_amps}A rating")
            
            # Find exact match first
            for part_number, spec in fuse_library.items():
                if spec.get('fuse_rating_amps') == fuse_rating_amps:
                    print(f"Found exact match: {part_number}")
                    return part_number
            
            # If no exact match, find the next higher rating
            available_ratings = []
            for part_number, spec in fuse_library.items():
                rating = spec.get('fuse_rating_amps')
                if rating and rating >= fuse_rating_amps:
                    available_ratings.append((rating, part_number))
            
            if available_ratings:
                # Sort by rating and return the lowest that meets the requirement
                available_ratings.sort(key=lambda x: x[0])
                result = available_ratings[0][1]
                print(f"Found next higher rating: {result} ({available_ratings[0][0]}A)")
                return result
            
            print(f"No suitable fuse found for {fuse_rating_amps}A")
            return "N/A"
            
        except Exception as e:
            print(f"Error getting fuse part number: {e}")
            import traceback
            traceback.print_exc()
            return "N/A"
        
    def get_whip_segment_part_number(self, row, selected_blocks):
        """Get whip cable segment part number from library"""
        try:
            from ..utils.bom_generator import BOMGenerator
            bom_generator = BOMGenerator(selected_blocks)
            
            component_type = row['Component Type']
            
            # Extract polarity from component type
            polarity = 'positive' if 'Positive' in component_type else 'negative'
            
            # Extract length from component type (e.g., "Positive Whip Cable Segment 25ft (8 AWG)")
            import re
            length_match = re.search(r'(\d+)ft', component_type)
            if not length_match:
                return "N/A"
            length_ft = int(length_match.group(1))
            
            # Extract wire gauge from component type
            gauge_match = re.search(r'\(([^)]+)\)', component_type)
            if not gauge_match:
                return "N/A"
            wire_gauge = gauge_match.group(1)
            
            # Find matching whip part number
            return bom_generator.find_matching_whip_part_number(wire_gauge, polarity, length_ft)
            
        except Exception as e:
            print(f"Error getting whip segment part number: {e}")
            return "N/A"

    def get_extender_segment_part_number(self, row, selected_blocks):
        """Get extender cable segment part number from library"""
        try:
            from ..utils.bom_generator import BOMGenerator
            bom_generator = BOMGenerator(selected_blocks)
            
            component_type = row['Component Type']
            
            # Extract polarity from component type
            polarity = 'positive' if 'Positive' in component_type else 'negative'
            
            # Extract length from component type (e.g., "Positive Extender Cable Segment 30ft (8 AWG)")
            import re
            length_match = re.search(r'(\d+)ft', component_type)
            if not length_match:
                return "N/A"
            length_ft = int(length_match.group(1))
            
            # Extract wire gauge from component type
            gauge_match = re.search(r'\(([^)]+)\)', component_type)
            if not gauge_match:
                return "N/A"
            wire_gauge = gauge_match.group(1)
            
            # Find matching extender part number
            return bom_generator.find_matching_extender_part_number(wire_gauge, polarity, length_ft)
            
        except Exception as e:
            print(f"Error getting extender segment part number: {e}")
            return "N/A"

    def open_harness_designer(self):
        """Open harness designer tool"""
        try:
            designer = HarnessDesigner(self)
            # Designer handles everything internally
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open harness designer: {str(e)}")

    def generate_harness_drawings(self):
        """Open harness catalog dialog for generating drawings"""
        try:
            dialog = HarnessCatalogDialog(self)
            # Dialog handles everything internally, no need to wait for result
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open harness catalog: {str(e)}")