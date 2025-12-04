"""
Pricing Manager Dialog for managing component pricing data
"""
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import json
import os
from typing import Dict, Any, Optional
import pandas as pd

class EditableTreeview(ttk.Treeview):
    """Treeview with editable cells"""
    
    def __init__(self, parent, on_cell_edit=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.on_cell_edit = on_cell_edit
        self._entry = None
        
        # Bind double-click to edit
        self.bind('<Double-1>', self._on_double_click)
    
    def _on_double_click(self, event):
        """Handle double-click to edit cell"""
        # Identify the cell
        region = self.identify_region(event.x, event.y)
        if region != 'cell':
            return
        
        column = self.identify_column(event.x)
        item = self.identify_row(event.y)
        
        if not item or not column:
            return
        
        # Don't allow editing the part number column (first column)
        col_idx = int(column.replace('#', '')) - 1
        if col_idx == 0:
            return
        
        # Get cell bbox
        bbox = self.bbox(item, column)
        if not bbox:
            return
        
        # Get current value
        values = self.item(item, 'values')
        current_value = values[col_idx] if col_idx < len(values) else ''
        
        # Create entry widget
        self._entry = ttk.Entry(self)
        self._entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        self._entry.insert(0, current_value)
        self._entry.select_range(0, tk.END)
        self._entry.focus_set()
        
        # Store edit info
        self._edit_item = item
        self._edit_column = column
        self._edit_col_idx = col_idx
        
        # Bind events
        self._entry.bind('<Return>', self._on_entry_return)
        self._entry.bind('<Escape>', self._on_entry_escape)
        self._entry.bind('<FocusOut>', self._on_entry_return)
    
    def _on_entry_return(self, event):
        """Handle entry confirmation"""
        if self._entry is None:
            return
        
        new_value = self._entry.get()
        
        # Validate it's a number
        try:
            if new_value.strip():
                float(new_value)
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter a valid number")
            self._entry.focus_set()
            return
        
        # Update the treeview
        values = list(self.item(self._edit_item, 'values'))
        old_value = values[self._edit_col_idx]
        values[self._edit_col_idx] = new_value
        self.item(self._edit_item, values=values)
        
        # Callback
        if self.on_cell_edit and new_value != old_value:
            self.on_cell_edit(self._edit_item, self._edit_col_idx, new_value)
        
        # Destroy entry
        self._entry.destroy()
        self._entry = None
    
    def _on_entry_escape(self, event):
        """Handle entry cancellation"""
        if self._entry:
            self._entry.destroy()
            self._entry = None

class PricingManager(tk.Toplevel):
    """Dialog for managing component pricing based on copper prices"""
    
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Pricing Manager")
        self.geometry("900x600")
        self.minsize(800, 500)
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
        # Pricing data
        self.pricing_data: Dict[str, Any] = {}
        self.modified = False
        
        # Load pricing data
        self.load_pricing_data()
        
        # Setup UI
        self.setup_ui()
        
        # Center on parent
        self.center_on_parent(parent)
    
    def center_on_parent(self, parent):
        """Center dialog on parent window"""
        self.update_idletasks()
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_w = parent.winfo_width()
        parent_h = parent.winfo_height()
        
        w = self.winfo_width()
        h = self.winfo_height()
        
        x = parent_x + (parent_w - w) // 2
        y = parent_y + (parent_h - h) // 2
        
        self.geometry(f"+{x}+{y}")
    
    def get_pricing_file_path(self) -> str:
        """Get the path to the pricing data file"""
        current_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(os.path.dirname(current_dir))
        return os.path.join(project_root, 'data', 'pricing_data.json')
    
    def load_pricing_data(self):
        """Load pricing data from JSON file"""
        filepath = self.get_pricing_file_path()
        
        if os.path.exists(filepath):
            try:
                with open(filepath, 'r') as f:
                    self.pricing_data = json.load(f)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load pricing data: {str(e)}")
                self.pricing_data = self.get_default_pricing_data()
        else:
            self.pricing_data = self.get_default_pricing_data()
    
    def get_default_pricing_data(self) -> Dict[str, Any]:
        """Return default pricing data structure"""
        return {
            "settings": {
                "current_copper_price": 4.5,
                "copper_price_tiers": [4.0, 4.5, 5.0, 5.5, 6.0]
            },
            "extenders": {"8_awg": {}, "10_awg": {}},
            "whips": {"8_awg": {}, "10_awg": {}},
            "harnesses": {"first_solar": {}, "standard": {}},
            "fuses": {},
            "combiner_boxes": {}
        }
    
    def save_pricing_data(self):
        """Save pricing data to JSON file"""
        filepath = self.get_pricing_file_path()
        
        try:
            # Ensure data directory exists
            os.makedirs(os.path.dirname(filepath), exist_ok=True)
            
            with open(filepath, 'w') as f:
                json.dump(self.pricing_data, f, indent=2)
            
            self.modified = False
            messagebox.showinfo("Success", "Pricing data saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save pricing data: {str(e)}")
    
    def setup_ui(self):
        """Setup the user interface"""
        # Main container
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Top section - Copper price settings
        settings_frame = ttk.LabelFrame(main_frame, text="Copper Price Settings", padding="10")
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Current copper price
        ttk.Label(settings_frame, text="Current Copper Price ($/lb):").grid(
            row=0, column=0, padx=5, pady=5, sticky=tk.W
        )
        
        self.copper_price_var = tk.StringVar(
            value=str(self.pricing_data.get('settings', {}).get('current_copper_price', 4.5))
        )
        self.copper_price_entry = ttk.Entry(settings_frame, textvariable=self.copper_price_var, width=10)
        self.copper_price_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # Active tier display
        ttk.Label(settings_frame, text="Active Price Tier:").grid(
            row=0, column=2, padx=(20, 5), pady=5, sticky=tk.W
        )
        
        self.active_tier_label = ttk.Label(settings_frame, text="", font=('TkDefaultFont', 10, 'bold'))
        self.active_tier_label.grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        
        # Update tier button
        ttk.Button(settings_frame, text="Update Tier", command=self.update_active_tier).grid(
            row=0, column=4, padx=5, pady=5
        )
        
        # Initialize active tier display
        self.update_active_tier()
        
        # Notebook for different component categories
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Create tabs for each category
        self.create_extenders_tab()
        self.create_whips_tab()
        self.create_harnesses_tab()
        self.create_fuses_tab()
        self.create_combiner_boxes_tab()
        
        # Bottom buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="Save", command=self.save_pricing_data).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Close", command=self.on_close).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Import from Excel", command=self.import_from_excel).pack(side=tk.LEFT, padx=5)
    
    def update_active_tier(self):
        """Update the active tier display based on current copper price"""
        try:
            current_price = float(self.copper_price_var.get())
            tiers = self.pricing_data.get('settings', {}).get('copper_price_tiers', [4.0, 4.5, 5.0, 5.5, 6.0])
            
            # Find the appropriate tier
            active_tier = tiers[0]
            for tier in tiers:
                if current_price >= tier:
                    active_tier = tier
                else:
                    break
            
            # Find the upper bound
            tier_idx = tiers.index(active_tier)
            if tier_idx < len(tiers) - 1:
                upper_bound = tiers[tier_idx + 1]
                tier_text = f"${active_tier:.2f} - ${upper_bound:.2f}"
            else:
                tier_text = f"${active_tier:.2f}+"
            
            self.active_tier_label.config(text=tier_text)
            
            # Update settings
            self.pricing_data['settings']['current_copper_price'] = current_price
            self.modified = True
            
        except ValueError:
            self.active_tier_label.config(text="Invalid price")
    
    def create_extenders_tab(self):
        """Create the extenders pricing tab"""
        frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(frame, text="Extenders")
        
        # Sub-notebook for 8 AWG and 10 AWG
        sub_notebook = ttk.Notebook(frame)
        sub_notebook.pack(fill=tk.BOTH, expand=True)
        
        # 8 AWG tab
        awg8_frame = ttk.Frame(sub_notebook, padding="5")
        sub_notebook.add(awg8_frame, text="8 AWG")
        self.create_pricing_table(awg8_frame, 'extenders', '8_awg')
        
        # 10 AWG tab
        awg10_frame = ttk.Frame(sub_notebook, padding="5")
        sub_notebook.add(awg10_frame, text="10 AWG")
        self.create_pricing_table(awg10_frame, 'extenders', '10_awg')
    
    def create_whips_tab(self):
        """Create the whips pricing tab"""
        frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(frame, text="Whips")
        
        # Sub-notebook for 8 AWG and 10 AWG
        sub_notebook = ttk.Notebook(frame)
        sub_notebook.pack(fill=tk.BOTH, expand=True)
        
        # 8 AWG tab
        awg8_frame = ttk.Frame(sub_notebook, padding="5")
        sub_notebook.add(awg8_frame, text="8 AWG")
        self.create_pricing_table(awg8_frame, 'whips', '8_awg')
        
        # 10 AWG tab
        awg10_frame = ttk.Frame(sub_notebook, padding="5")
        sub_notebook.add(awg10_frame, text="10 AWG")
        self.create_pricing_table(awg10_frame, 'whips', '10_awg')
    
    def create_harnesses_tab(self):
        """Create the harnesses pricing tab"""
        frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(frame, text="Harnesses")
        
        # Sub-notebook for First Solar and Standard
        sub_notebook = ttk.Notebook(frame)
        sub_notebook.pack(fill=tk.BOTH, expand=True)
        
        # First Solar tab
        fs_frame = ttk.Frame(sub_notebook, padding="5")
        sub_notebook.add(fs_frame, text="First Solar")
        self.create_pricing_table(fs_frame, 'harnesses', 'first_solar')
        
        # Standard tab
        std_frame = ttk.Frame(sub_notebook, padding="5")
        sub_notebook.add(std_frame, text="Standard")
        self.create_pricing_table(std_frame, 'harnesses', 'standard')
    
    def create_fuses_tab(self):
        """Create the fuses pricing tab (flat pricing)"""
        frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(frame, text="Fuses")
        
        # Info label
        ttk.Label(frame, text="Fuses have flat pricing (not copper-indexed). Double-click price to edit.").pack(pady=5)
        
        # Create editable treeview for flat pricing
        columns = ('part_number', 'price')
        
        def on_fuse_edit(item, col_idx, new_value):
            if col_idx == 1:  # Price column
                values = tree.item(item, 'values')
                part_number = values[0]
                try:
                    self.pricing_data['fuses'][part_number] = float(new_value) if new_value.strip() else 0
                    self.modified = True
                except ValueError:
                    pass
        
        tree = EditableTreeview(frame, on_cell_edit=on_fuse_edit, columns=columns, show='headings', height=15)
        
        tree.heading('part_number', text='Part Number')
        tree.heading('price', text='Unit Price ($)')
        
        tree.column('part_number', width=200)
        tree.column('price', width=100)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Populate with data
        fuses_data = self.pricing_data.get('fuses', {})
        for part_number, price in fuses_data.items():
            tree.insert('', 'end', values=(part_number, f"{price:.2f}"))
        
        # Count label
        ttk.Label(frame, text=f"Total items: {len(fuses_data)}").pack(pady=5)
    
    def create_combiner_boxes_tab(self):
        """Create the combiner boxes pricing tab (flat pricing)"""
        frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(frame, text="Combiner Boxes")
        
        # Info label
        ttk.Label(frame, text="Combiner boxes have flat pricing (not copper-indexed)").pack(pady=5)
        ttk.Label(frame, text="Coming soon - configure combiner box pricing here").pack(pady=20)
    
    def create_pricing_table(self, parent, category: str, subcategory: str):
        """Create a pricing table for copper-indexed components"""
        tiers = self.pricing_data.get('settings', {}).get('copper_price_tiers', [4.0, 4.5, 5.0, 5.5, 6.0])
        
        # Create columns: part_number + one column per tier
        columns = ['part_number'] + [f"tier_{t}" for t in tiers]
        
        # Use editable treeview with callback
        def on_edit(item, col_idx, new_value):
            self.on_pricing_cell_edit(tree, item, col_idx, category, subcategory, tiers)
        
        tree = EditableTreeview(parent, on_cell_edit=on_edit, columns=columns, show='headings', height=15)
        
        # Configure headings
        tree.heading('part_number', text='Part Number')
        tree.column('part_number', width=150)
        
        for tier in tiers:
            col_id = f"tier_{tier}"
            tree.heading(col_id, text=f"${tier:.1f}-${tier+0.5:.1f}")
            tree.column(col_id, width=100)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Populate with data
        category_data = self.pricing_data.get(category, {}).get(subcategory, {})
        for part_number, prices in category_data.items():
            values = [part_number]
            for tier in tiers:
                price = prices.get(str(tier), prices.get(tier, 0))
                values.append(f"{float(price):.2f}" if price else "")
            tree.insert('', 'end', values=values)
        
        # Bottom frame for count and buttons
        bottom_frame = ttk.Frame(parent)
        bottom_frame.pack(fill=tk.X, pady=5)
        
        # Show count
        count_label = ttk.Label(bottom_frame, text=f"Total items: {len(category_data)}")
        count_label.pack(side=tk.LEFT, padx=5)
        
        # Store reference for later
        tree._category = category
        tree._subcategory = subcategory
        tree._count_label = count_label

    def on_pricing_cell_edit(self, tree, item, col_idx, category: str, subcategory: str, tiers: list):
        """Handle cell edit in pricing table"""
        values = tree.item(item, 'values')
        part_number = values[0]
        
        # Get the tier for this column (col_idx 0 is part number, so tier index is col_idx - 1)
        tier_idx = col_idx - 1
        if tier_idx < 0 or tier_idx >= len(tiers):
            return
        
        tier = tiers[tier_idx]
        new_price = values[col_idx]
        
        # Update pricing data
        if category not in self.pricing_data:
            self.pricing_data[category] = {}
        if subcategory not in self.pricing_data[category]:
            self.pricing_data[category][subcategory] = {}
        if part_number not in self.pricing_data[category][subcategory]:
            self.pricing_data[category][subcategory][part_number] = {}
        
        try:
            self.pricing_data[category][subcategory][part_number][str(tier)] = float(new_price) if new_price.strip() else 0
            self.modified = True
        except ValueError:
            pass  # Invalid number, ignore
    
    def import_from_excel(self):
        """Import pricing data from Excel file"""
        filepath = filedialog.askopenfilename(
            title="Select Pricing Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if not filepath:
            return
        
        try:
            # Read all sheet names
            xlsx = pd.ExcelFile(filepath)
            sheet_names = xlsx.sheet_names
            
            # Show import dialog
            import_dialog = PricingImportDialog(self, filepath, sheet_names)
            self.wait_window(import_dialog)
            
            if import_dialog.result:
                # Merge imported data with existing data
                imported_data = import_dialog.result
                self._merge_imported_data(imported_data)
                
                # Refresh the UI
                self._refresh_all_tables()
                
                self.modified = True
                messagebox.showinfo("Import Complete", "Pricing data imported successfully!\n\nClick 'Save' to persist the changes.")
        
        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import pricing data:\n{str(e)}")
    
    def _merge_imported_data(self, imported_data: Dict[str, Any]):
        """Merge imported data into existing pricing data"""
        for category, subcategories in imported_data.items():
            if category == 'settings':
                continue  # Don't overwrite settings
            
            if category not in self.pricing_data:
                self.pricing_data[category] = {}
            
            if isinstance(subcategories, dict):
                # Check if it's a flat structure (like fuses) or nested (like extenders)
                first_value = next(iter(subcategories.values()), None) if subcategories else None
                
                if isinstance(first_value, dict):
                    # Nested structure (extenders, whips, harnesses)
                    for subcategory, items in subcategories.items():
                        if subcategory not in self.pricing_data[category]:
                            self.pricing_data[category][subcategory] = {}
                        self.pricing_data[category][subcategory].update(items)
                else:
                    # Flat structure (fuses)
                    self.pricing_data[category].update(subcategories)
    
    def _refresh_all_tables(self):
        """Refresh all pricing tables - recreate the notebook tabs"""
        # Store current tab
        current_tab = self.notebook.index(self.notebook.select())
        
        # Remove all tabs
        for tab in self.notebook.tabs():
            self.notebook.forget(tab)
        
        # Recreate tabs
        self.create_extenders_tab()
        self.create_whips_tab()
        self.create_harnesses_tab()
        self.create_fuses_tab()
        self.create_combiner_boxes_tab()
        
        # Restore tab selection
        try:
            self.notebook.select(current_tab)
        except:
            pass
    
    def on_close(self):
        """Handle dialog close"""
        if self.modified:
            result = messagebox.askyesnocancel(
                "Unsaved Changes",
                "You have unsaved changes. Do you want to save before closing?"
            )
            if result is True:  # Yes
                self.save_pricing_data()
                self.destroy()
            elif result is False:  # No
                self.destroy()
            # If None (Cancel), do nothing
        else:
            self.destroy()

class PricingImportDialog(tk.Toplevel):
    """Dialog for importing pricing data from Excel"""
    
    def __init__(self, parent, filepath: str, sheet_names: list):
        super().__init__(parent)
        self.title("Import Pricing Data")
        self.geometry("600x500")
        
        self.filepath = filepath
        self.sheet_names = sheet_names
        self.result = None
        
        # Make modal
        self.transient(parent)
        self.grab_set()
        
        self.setup_ui()
        self.analyze_sheets()
        
        # Center on parent
        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")
    
    def setup_ui(self):
        """Setup the UI"""
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # File info
        ttk.Label(main_frame, text=f"File: {os.path.basename(self.filepath)}", 
                  font=('TkDefaultFont', 10, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
        # Sheets found
        ttk.Label(main_frame, text="Sheets found and mapping:").pack(anchor=tk.W)
        
        # Treeview for sheet mapping
        columns = ('sheet', 'category', 'items')
        self.sheet_tree = ttk.Treeview(main_frame, columns=columns, show='headings', height=10)
        
        self.sheet_tree.heading('sheet', text='Sheet Name')
        self.sheet_tree.heading('category', text='Will Import As')
        self.sheet_tree.heading('items', text='Items Found')
        
        self.sheet_tree.column('sheet', width=150)
        self.sheet_tree.column('category', width=200)
        self.sheet_tree.column('items', width=100)
        
        self.sheet_tree.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Analyzing sheets...")
        self.status_label.pack(anchor=tk.W, pady=5)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(button_frame, text="Import Selected", command=self.do_import).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=5)
    
    def analyze_sheets(self):
        """Analyze the Excel sheets and show what can be imported"""
        # Mapping of sheet names to categories
        sheet_mapping = {
            '8EXT-Q': ('extenders', '8_awg', 'Extenders - 8 AWG'),
            '10EXT-Q': ('extenders', '10_awg', 'Extenders - 10 AWG'),
            '8WHI-Q': ('whips', '8_awg', 'Whips - 8 AWG'),
            '10WHI-Q': ('whips', '10_awg', 'Whips - 10 AWG'),
            'QUOTATION FS': ('harnesses', 'first_solar', 'Harnesses - First Solar'),
            'QUOTATION': ('harnesses', 'standard', 'Harnesses - Standard'),
            'QUOTATION Fuses': ('fuses', None, 'Fuses (flat pricing)'),
        }
        
        found_count = 0
        total_items = 0
        
        for sheet_name in self.sheet_names:
            if sheet_name in sheet_mapping:
                category, subcategory, display_name = sheet_mapping[sheet_name]
                
                # Count items in sheet
                try:
                    df = pd.read_excel(self.filepath, sheet_name=sheet_name)
                    # Count non-empty rows (approximate)
                    item_count = len(df) - 2  # Subtract header rows
                    if item_count < 0:
                        item_count = 0
                except:
                    item_count = "?"
                
                self.sheet_tree.insert('', 'end', values=(sheet_name, display_name, item_count))
                found_count += 1
                if isinstance(item_count, int):
                    total_items += item_count
        
        self.status_label.config(text=f"Found {found_count} importable sheets with ~{total_items} items")
    
    def do_import(self):
        """Perform the import"""
        self.status_label.config(text="Importing...")
        self.update()
        
        try:
            imported_data = self._parse_all_sheets()
            self.result = imported_data
            self.destroy()
        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to parse Excel data:\n{str(e)}")
    
    def _parse_all_sheets(self) -> Dict[str, Any]:
        """Parse all recognized sheets and return pricing data"""
        result = {
            'extenders': {'8_awg': {}, '10_awg': {}},
            'whips': {'8_awg': {}, '10_awg': {}},
            'harnesses': {'first_solar': {}, 'standard': {}},
            'fuses': {}
        }
        
        # Parse extenders
        if '8EXT-Q' in self.sheet_names:
            result['extenders']['8_awg'] = self._parse_copper_indexed_sheet('8EXT-Q')
        if '10EXT-Q' in self.sheet_names:
            result['extenders']['10_awg'] = self._parse_copper_indexed_sheet('10EXT-Q')
        
        # Parse whips
        if '8WHI-Q' in self.sheet_names:
            result['whips']['8_awg'] = self._parse_copper_indexed_sheet('8WHI-Q')
        if '10WHI-Q' in self.sheet_names:
            result['whips']['10_awg'] = self._parse_copper_indexed_sheet('10WHI-Q')
        
        # Parse harnesses
        if 'QUOTATION FS' in self.sheet_names:
            result['harnesses']['first_solar'] = self._parse_copper_indexed_sheet('QUOTATION FS')
        if 'QUOTATION' in self.sheet_names:
            result['harnesses']['standard'] = self._parse_copper_indexed_sheet('QUOTATION')
        
        # Parse fuses
        if 'QUOTATION Fuses' in self.sheet_names:
            result['fuses'] = self._parse_fuses_sheet()
        
        return result
    
    def _parse_copper_indexed_sheet(self, sheet_name: str) -> Dict[str, Dict[str, float]]:
        """Parse a copper-indexed pricing sheet"""
        df = pd.read_excel(self.filepath, sheet_name=sheet_name)
        
        result = {}
        
        # Get tier values from first row (row 0)
        # Tiers are in columns 1-5 (index 1 to 5)
        tiers = []
        for col_idx in range(1, min(6, len(df.columns))):
            try:
                tier_val = df.iloc[0, col_idx]
                if pd.notna(tier_val):
                    tiers.append(float(tier_val))
            except (ValueError, TypeError):
                continue
        
        if not tiers:
            tiers = [4.0, 4.5, 5.0, 5.5, 6.0]  # Default
        
        # Parse data rows (starting from row 2, index 2)
        for row_idx in range(2, len(df)):
            part_number = df.iloc[row_idx, 0]
            
            # Skip empty rows
            if pd.isna(part_number) or str(part_number).strip() == '':
                continue
            
            part_number = str(part_number).strip()
            
            # Skip if part number ends with -CUSTOM (per user request, remove suffix)
            if part_number.endswith('-CUSTOM'):
                part_number = part_number.replace('-CUSTOM', '')
            
            # Get prices for each tier
            prices = {}
            for tier_idx, tier in enumerate(tiers):
                col_idx = tier_idx + 1
                if col_idx < len(df.columns):
                    try:
                        price = df.iloc[row_idx, col_idx]
                        if pd.notna(price):
                            prices[str(tier)] = float(price)
                    except (ValueError, TypeError):
                        continue
            
            if prices:
                result[part_number] = prices
        
        return result
    
    def _parse_fuses_sheet(self) -> Dict[str, float]:
        """Parse the fuses pricing sheet (using 500 qty column)"""
        df = pd.read_excel(self.filepath, sheet_name='QUOTATION Fuses')
        
        result = {}
        
        # Data starts at row 1 (index 1), part number in col 0, 500 qty price in col 1
        for row_idx in range(1, len(df)):
            part_number = df.iloc[row_idx, 0]
            
            if pd.isna(part_number) or str(part_number).strip() == '':
                continue
            
            part_number = str(part_number).strip()
            
            # Skip header row if present
            if part_number.upper() == 'PART NO.':
                continue
            
            try:
                price = df.iloc[row_idx, 1]  # 500 qty column
                if pd.notna(price):
                    result[part_number] = float(price)
            except (ValueError, TypeError):
                continue
        
        return result