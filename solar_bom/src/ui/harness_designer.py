import tkinter as tk
from tkinter import ttk, messagebox
import json
from pathlib import Path
from typing import Optional, Dict, Any
from ..utils.harness_drawing_generator import HarnessDrawingGenerator

class HarnessDesigner(tk.Toplevel):
    """Harness design tool for creating custom harness templates"""
    
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        
        # Initialize harness library path
        self.harness_library_path = 'data/harness_library.json'
        self.harness_library = self.load_harness_library()
        
        # Set up window properties
        self.title("Harness Designer")
        self.geometry("1000x700")
        self.minsize(800, 600)
        
        # Make window modal
        self.transient(parent)
        self.grab_set()
        
        # Position window relative to parent
        x = parent.winfo_rootx() + 50
        y = parent.winfo_rooty() + 50
        self.geometry(f"+{x}+{y}")
        
        # Initialize UI
        self.setup_ui()
        self.update_template_list()
    
    def load_harness_library(self) -> Dict[str, Any]:
        """Load existing harness library"""
        try:
            # Ensure data directory exists
            Path('data').mkdir(exist_ok=True)
            
            if Path(self.harness_library_path).exists():
                with open(self.harness_library_path, 'r') as f:
                    return json.load(f)
            else:
                # Create empty library if file doesn't exist
                return {}
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load harness library: {str(e)}")
            return {}
    
    def save_harness_library(self):
        """Save harness library to JSON file"""
        try:
            with open(self.harness_library_path, 'w') as f:
                json.dump(self.harness_library, f, indent=2)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save harness library: {str(e)}")
            return False
    
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
        
        # Left column - Template List and Controls
        left_column = ttk.Frame(main_container)
        left_column.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.N, tk.S, tk.W))
        
        # Template List section
        template_frame = ttk.LabelFrame(left_column, text="Saved Templates", padding="5")
        template_frame.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N))
        
        self.template_listbox = tk.Listbox(template_frame, width=35, height=15)
        self.template_listbox.grid(row=0, column=0, padx=5, pady=5)
        self.template_listbox.bind('<<ListboxSelect>>', self.on_template_select)
        
        template_buttons = ttk.Frame(template_frame)
        template_buttons.grid(row=1, column=0, padx=5, pady=5)
        
        ttk.Button(template_buttons, text="Load", command=self.load_template).grid(row=0, column=0, padx=2)
        ttk.Button(template_buttons, text="Delete", command=self.delete_template).grid(row=0, column=1, padx=2)
        ttk.Button(template_buttons, text="Generate Drawing", command=self.generate_current_drawing).grid(row=0, column=2, padx=2)
        
        # Right column - Designer
        right_column = ttk.Frame(main_container)
        right_column.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Designer section
        self.setup_designer_form(right_column)
        
        # Bottom buttons
        self.setup_bottom_buttons(main_container)
    
    def setup_designer_form(self, parent):
        """Set up the harness designer form"""
        designer_frame = ttk.LabelFrame(parent, text="Harness Designer", padding="10")
        designer_frame.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Basic Information
        basic_frame = ttk.LabelFrame(designer_frame, text="Basic Information", padding="5")
        basic_frame.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Part Number
        ttk.Label(basic_frame, text="Part Number:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.part_number_var = tk.StringVar()
        ttk.Entry(basic_frame, textvariable=self.part_number_var, width=20).grid(row=0, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        # ATPI Part Number
        ttk.Label(basic_frame, text="ATPI Part Number:").grid(row=0, column=2, padx=5, pady=2, sticky=tk.W)
        self.atpi_part_number_var = tk.StringVar()
        ttk.Entry(basic_frame, textvariable=self.atpi_part_number_var, width=20).grid(row=0, column=3, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        # Description
        ttk.Label(basic_frame, text="Description:").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.description_var = tk.StringVar()
        ttk.Entry(basic_frame, textvariable=self.description_var, width=60).grid(row=1, column=1, columnspan=3, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        basic_frame.grid_columnconfigure(1, weight=1)
        basic_frame.grid_columnconfigure(3, weight=1)
        
        # Configuration
        config_frame = ttk.LabelFrame(designer_frame, text="Configuration", padding="5")
        config_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Number of Strings
        ttk.Label(config_frame, text="Number of Strings:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.num_strings_var = tk.StringVar(value="2")
        ttk.Spinbox(config_frame, from_=1, to=20, textvariable=self.num_strings_var, 
                   width=10, validate='all', validatecommand=(self.register(self.validate_integer), '%P')).grid(row=0, column=1, padx=5, pady=2)
        
        # Polarity
        ttk.Label(config_frame, text="Polarity:").grid(row=0, column=2, padx=5, pady=2, sticky=tk.W)
        self.polarity_var = tk.StringVar(value="positive")
        polarity_combo = ttk.Combobox(config_frame, textvariable=self.polarity_var, state='readonly', width=15)
        polarity_combo['values'] = ('positive', 'negative')
        polarity_combo.grid(row=0, column=3, padx=5, pady=2)
        
        # String Spacing
        ttk.Label(config_frame, text="String Spacing (ft):").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.string_spacing_var = tk.StringVar(value="102")
        ttk.Entry(config_frame, textvariable=self.string_spacing_var, width=10, 
                 validate='all', validatecommand=(self.register(self.validate_float), '%P')).grid(row=1, column=1, padx=5, pady=2)
        
        # Category
        ttk.Label(config_frame, text="Category:").grid(row=1, column=2, padx=5, pady=2, sticky=tk.W)
        self.category_var = tk.StringVar(value="Standard")
        category_combo = ttk.Combobox(config_frame, textvariable=self.category_var, state='readonly', width=15)
        category_combo['values'] = ('Standard', 'First Solar', 'Custom')
        category_combo.grid(row=1, column=3, padx=5, pady=2)
        
        # Wire Specifications
        wire_frame = ttk.LabelFrame(designer_frame, text="Wire Specifications", padding="5")
        wire_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Drop Wire Gauge
        ttk.Label(wire_frame, text="Drop Wire Gauge:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.drop_wire_gauge_var = tk.StringVar(value="10 AWG")
        drop_combo = ttk.Combobox(wire_frame, textvariable=self.drop_wire_gauge_var, state='readonly', width=15)
        drop_combo['values'] = ('4 AWG', '6 AWG', '8 AWG', '10 AWG', '12 AWG')
        drop_combo.grid(row=0, column=1, padx=5, pady=2)
        
        # Trunk Wire Gauge
        ttk.Label(wire_frame, text="Trunk Wire Gauge:").grid(row=0, column=2, padx=5, pady=2, sticky=tk.W)
        self.trunk_wire_gauge_var = tk.StringVar(value="8 AWG")
        trunk_combo = ttk.Combobox(wire_frame, textvariable=self.trunk_wire_gauge_var, state='readonly', width=15)
        trunk_combo['values'] = ('4 AWG', '6 AWG', '8 AWG', '10 AWG', '12 AWG')
        trunk_combo.grid(row=0, column=3, padx=5, pady=2)
        
        # Connector Type
        ttk.Label(wire_frame, text="Connector Type:").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.connector_type_var = tk.StringVar(value="MC4")
        connector_combo = ttk.Combobox(wire_frame, textvariable=self.connector_type_var, width=20)
        connector_combo['values'] = ('MC4', 'PV4S/MC4', 'H4', 'Custom')
        connector_combo.grid(row=1, column=1, columnspan=2, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        # Fuse Configuration
        fuse_frame = ttk.LabelFrame(designer_frame, text="Fuse Configuration", padding="5")
        fuse_frame.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Fused checkbox
        self.fused_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(fuse_frame, text="Include Fuses", variable=self.fused_var, 
                       command=self.on_fused_change).grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        
        # Fuse Rating
        ttk.Label(fuse_frame, text="Fuse Rating:").grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)
        self.fuse_rating_var = tk.StringVar(value="15A")
        self.fuse_rating_combo = ttk.Combobox(fuse_frame, textvariable=self.fuse_rating_var, state='readonly', width=10)
        self.fuse_rating_combo['values'] = ('5A', '10A', '15A', '20A', '25A', '30A', '32A', '35A', '40A', '45A')
        self.fuse_rating_combo.grid(row=0, column=2, padx=5, pady=2)
        self.fuse_rating_combo.configure(state='disabled')  # Initially disabled
        
        # Preview Section
        preview_frame = ttk.LabelFrame(designer_frame, text="Preview", padding="5")
        preview_frame.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Preview text
        self.preview_text = tk.Text(preview_frame, height=8, width=70, wrap=tk.WORD, state='disabled')
        self.preview_text.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        preview_scroll = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_text.yview)
        preview_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.preview_text.configure(yscrollcommand=preview_scroll.set)
        
        preview_frame.grid_columnconfigure(0, weight=1)
        preview_frame.grid_rowconfigure(0, weight=1)
        
        # Bind events for real-time preview updates
        for var in [self.part_number_var, self.atpi_part_number_var, self.description_var, 
                   self.num_strings_var, self.polarity_var, self.string_spacing_var, 
                   self.category_var, self.drop_wire_gauge_var, self.trunk_wire_gauge_var, 
                   self.connector_type_var, self.fuse_rating_var]:
            var.trace('w', lambda *args: self.update_preview())
        
        self.fused_var.trace('w', lambda *args: self.update_preview())
    
    def setup_bottom_buttons(self, parent):
        """Set up bottom action buttons"""
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=1, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="New Template", command=self.new_template).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Save Template", command=self.save_template).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame, text="Generate Drawing", command=self.generate_drawing).grid(row=0, column=2, padx=5)
        ttk.Button(button_frame, text="Generate All Drawings", command=self.generate_all_drawings).grid(row=0, column=3, padx=5)
        ttk.Button(button_frame, text="Close", command=self.destroy).grid(row=0, column=4, padx=5)
    
    def validate_integer(self, value):
        """Validate integer input"""
        if value == "":
            return True
        try:
            int(value)
            return True
        except ValueError:
            return False
    
    def validate_float(self, value):
        """Validate float input"""
        if value == "":
            return True
        try:
            float(value)
            return True
        except ValueError:
            return False
    
    def on_fused_change(self):
        """Handle fused checkbox change"""
        if self.fused_var.get():
            self.fuse_rating_combo.configure(state='readonly')
        else:
            self.fuse_rating_combo.configure(state='disabled')
        self.update_preview()
    
    def update_preview(self):
        """Update the preview text with current harness specification"""
        try:
            harness_spec = self.get_current_harness_spec()
            
            preview_text = f"""HARNESS SPECIFICATION PREVIEW

Part Number: {harness_spec.get('part_number', 'Not specified')}
ATPI Part Number: {harness_spec.get('atpi_part_number', 'Not specified')}
Description: {harness_spec.get('description', 'Not specified')}

Configuration:
  • Number of Strings: {harness_spec.get('num_strings', 0)}
  • Polarity: {harness_spec.get('polarity', '').title()}
  • String Spacing: {harness_spec.get('string_spacing_ft', 0)}'
  • Category: {harness_spec.get('category', '')}

Wire Specifications:
  • Drop Wire Gauge: {harness_spec.get('drop_wire_gauge', '')}
  • Trunk Wire Gauge: {harness_spec.get('trunk_wire_gauge', '')}
  • Connector Type: {harness_spec.get('connector_type', '')}

Fuse Configuration:
  • Fused: {'Yes' if harness_spec.get('fused', False) else 'No'}"""

            if harness_spec.get('fused', False):
                preview_text += f"\n  • Fuse Rating: {harness_spec.get('fuse_rating', 'Not specified')}"
            
            self.preview_text.configure(state='normal')
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(1.0, preview_text)
            self.preview_text.configure(state='disabled')
            
        except Exception as e:
            print(f"Error updating preview: {str(e)}")
    
    def get_current_harness_spec(self) -> Dict[str, Any]:
        """Get current harness specification from form"""
        return {
            'part_number': self.part_number_var.get().strip(),
            'atpi_part_number': self.atpi_part_number_var.get().strip(),
            'description': self.description_var.get().strip(),
            'num_strings': int(self.num_strings_var.get()) if self.num_strings_var.get() else 0,
            'polarity': self.polarity_var.get(),
            'string_spacing_ft': float(self.string_spacing_var.get()) if self.string_spacing_var.get() else 0,
            'category': self.category_var.get(),
            'drop_wire_gauge': self.drop_wire_gauge_var.get(),
            'trunk_wire_gauge': self.trunk_wire_gauge_var.get(),
            'connector_type': self.connector_type_var.get(),
            'fused': self.fused_var.get(),
            'fuse_rating': self.fuse_rating_var.get() if self.fused_var.get() else None
        }
    
    def new_template(self):
        """Clear form for new template"""
        self.part_number_var.set("")
        self.atpi_part_number_var.set("")
        self.description_var.set("")
        self.num_strings_var.set("2")
        self.polarity_var.set("positive")
        self.string_spacing_var.set("102")
        self.category_var.set("Standard")
        self.drop_wire_gauge_var.set("10 AWG")
        self.trunk_wire_gauge_var.set("8 AWG")
        self.connector_type_var.set("MC4")
        self.fused_var.set(False)
        self.fuse_rating_var.set("15A")
        self.on_fused_change()
    
    def save_template(self):
        """Save current template to library"""
        harness_spec = self.get_current_harness_spec()
        
        # Validate required fields
        if not harness_spec['part_number']:
            messagebox.showerror("Error", "Part number is required")
            return
        
        if not harness_spec['description']:
            messagebox.showerror("Error", "Description is required")
            return
        
        if harness_spec['num_strings'] <= 0:
            messagebox.showerror("Error", "Number of strings must be greater than 0")
            return
        
        if harness_spec['string_spacing_ft'] <= 0:
            messagebox.showerror("Error", "String spacing must be greater than 0")
            return
        
        part_number = harness_spec['part_number']
        
        # Check if part number already exists
        if part_number in self.harness_library:
            if not messagebox.askyesno("Confirm", f"Part number '{part_number}' already exists. Overwrite?"):
                return
        
        # Save to library
        self.harness_library[part_number] = harness_spec
        
        if self.save_harness_library():
            messagebox.showinfo("Success", f"Template '{part_number}' saved successfully")
            self.update_template_list()
        
    def update_template_list(self):
        """Update the template listbox"""
        self.template_listbox.delete(0, tk.END)
        
        # Sort templates by category then part number
        sorted_templates = sorted(self.harness_library.items(), 
                                key=lambda x: (x[1].get('category', ''), x[0]))
        
        for part_number, spec in sorted_templates:
            category = spec.get('category', 'Unknown')
            display_text = f"[{category}] {part_number}"
            self.template_listbox.insert(tk.END, display_text)
    
    def on_template_select(self, event=None):
        """Handle template selection event"""
        selection = self.template_listbox.curselection()
        if selection:
            self.load_template()
    
    def load_template(self):
        """Load selected template into form"""
        selection = self.template_listbox.curselection()
        if not selection:
            return
        
        # Extract part number from display text
        display_text = self.template_listbox.get(selection[0])
        # Format is "[Category] PartNumber"
        part_number = display_text.split('] ', 1)[1] if '] ' in display_text else display_text
        
        if part_number not in self.harness_library:
            messagebox.showerror("Error", f"Template '{part_number}' not found")
            return
        
        spec = self.harness_library[part_number]
        
        # Load values into form
        self.part_number_var.set(spec.get('part_number', ''))
        self.atpi_part_number_var.set(spec.get('atpi_part_number', ''))
        self.description_var.set(spec.get('description', ''))
        self.num_strings_var.set(str(spec.get('num_strings', 2)))
        self.polarity_var.set(spec.get('polarity', 'positive'))
        self.string_spacing_var.set(str(spec.get('string_spacing_ft', 102)))
        self.category_var.set(spec.get('category', 'Standard'))
        self.drop_wire_gauge_var.set(spec.get('drop_wire_gauge', '10 AWG'))
        self.trunk_wire_gauge_var.set(spec.get('trunk_wire_gauge', '8 AWG'))
        self.connector_type_var.set(spec.get('connector_type', 'MC4'))
        self.fused_var.set(spec.get('fused', False))
        self.fuse_rating_var.set(spec.get('fuse_rating', '15A') or '15A')
        
        self.on_fused_change()
        self.update_preview()
    
    def delete_template(self):
        """Delete selected template"""
        selection = self.template_listbox.curselection()
        if not selection:
            return
        
        # Extract part number from display text
        display_text = self.template_listbox.get(selection[0])
        part_number = display_text.split('] ', 1)[1] if '] ' in display_text else display_text
        
        if part_number not in self.harness_library:
            messagebox.showerror("Error", f"Template '{part_number}' not found")
            return
        
        if messagebox.askyesno("Confirm", f"Delete template '{part_number}'?"):
            del self.harness_library[part_number]
            if self.save_harness_library():
                messagebox.showinfo("Success", f"Template '{part_number}' deleted")
                self.update_template_list()
                self.new_template()  # Clear form
    
    def generate_drawing(self):
        """Generate drawing for current harness"""
        harness_spec = self.get_current_harness_spec()
        
        if not harness_spec['part_number']:
            messagebox.showerror("Error", "Part number is required to generate drawing")
            return
        
        try:
            # Temporarily add to library for drawing generation
            temp_part_number = harness_spec['part_number']
            generator = HarnessDrawingGenerator(self.harness_library_path)
            
            # Add current spec to generator's library
            generator.harness_library[temp_part_number] = harness_spec
            
            # Generate drawing
            if generator.generate_harness_drawing(temp_part_number):
                messagebox.showinfo("Success", f"Drawing generated for '{temp_part_number}'")
            else:
                messagebox.showerror("Error", f"Failed to generate drawing for '{temp_part_number}'")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate drawing: {str(e)}")
    
    def generate_current_drawing(self):
        """Generate drawing for currently selected template"""
        selection = self.template_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "No template selected")
            return
        
        # Load the template first
        self.load_template()
        
        # Then generate the drawing
        self.generate_drawing()
    
    def generate_all_drawings(self):
        """Generate drawings for all templates in library"""
        if not self.harness_library:
            messagebox.showwarning("Warning", "No templates available")
            return
        
        try:
            generator = HarnessDrawingGenerator(self.harness_library_path)
            count = generator.generate_all_harness_drawings()
            messagebox.showinfo("Success", f"Generated {count} harness drawings")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate drawings: {str(e)}")