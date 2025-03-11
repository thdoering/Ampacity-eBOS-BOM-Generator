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
        
        # Add blocks to listbox
        for block_id, block in self.blocks.items():
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
            self.preview_tree.insert(
                '', 'end',
                values=(
                    row['Component Type'],
                    row['Description'],
                    row['Quantity'],
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
        
        # Ask for export location
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Export BOM to Excel"
        )
        
        if not filepath:
            return
        
        # Generate and export BOM
        bom_generator = BOMGenerator(selected_blocks)
        success = bom_generator.export_bom_to_excel(filepath)
        
        if success:
            messagebox.showinfo("Success", f"BOM exported successfully to {filepath}")
        else:
            messagebox.showerror("Error", "Failed to export BOM")