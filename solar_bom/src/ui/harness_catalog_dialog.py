import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
from typing import Dict, List, Optional
from ..utils.harness_drawing_generator import HarnessDrawingGenerator

class HarnessCatalogDialog(tk.Toplevel):
    """Dialog for selecting and generating harness drawings"""
    
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.result = None
        
        # Initialize the drawing generator
        try:
            self.generator = HarnessDrawingGenerator()
            # The generator already filters comments, so just check if empty
            if not self.generator.harness_library:
                messagebox.showwarning("Empty Library", 
                                    "No harnesses found in library.\n\n"
                                    "Please use the Harness Designer to create harness templates.")
                self.destroy()
                return
            self.available_harnesses = self.generator.get_available_harnesses()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load harness library: {str(e)}")
            self.destroy()
            return
        
        if not self.available_harnesses:
            messagebox.showwarning("Warning", "No harnesses found in library")
            self.destroy()
            return
        
        # Set up window properties
        self.title("Generate Harness Drawings")
        self.geometry("800x600")
        self.minsize(600, 400)
        
        # Make window modal
        self.transient(parent)
        self.grab_set()
        
        # Position window relative to parent
        x = parent.winfo_rootx() + 50
        y = parent.winfo_rooty() + 50
        self.geometry(f"+{x}+{y}")
        
        # Initialize UI
        self.setup_ui()
        
        # Center on parent
        self.center_on_parent()
    
    def setup_ui(self):
        """Create and arrange UI components"""
        main_frame = ttk.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)
        
        # Title and instructions
        title_label = ttk.Label(main_frame, text="Harness Drawing Generator", 
                               font=('Arial', 14, 'bold'))
        title_label.grid(row=0, column=0, pady=(0, 10), sticky=tk.W)
        
        instruction_label = ttk.Label(main_frame, 
                                    text="Select harnesses to generate technical drawings:")
        instruction_label.grid(row=1, column=0, pady=(0, 10), sticky=tk.W)
        
        # Harness selection frame
        selection_frame = ttk.LabelFrame(main_frame, text="Available Harnesses", padding="5")
        selection_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        selection_frame.grid_columnconfigure(0, weight=1)
        selection_frame.grid_rowconfigure(0, weight=1)
        
        # Create treeview for harness selection
        self.setup_harness_treeview(selection_frame)
        
        # Output options frame
        output_frame = ttk.LabelFrame(main_frame, text="Output Options", padding="5")
        output_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Output directory selection
        ttk.Label(output_frame, text="Output Directory:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        
        self.output_dir_var = tk.StringVar(value="harness_drawings")
        output_entry = ttk.Entry(output_frame, textvariable=self.output_dir_var, width=40)
        output_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 5))
        
        browse_btn = ttk.Button(output_frame, text="Browse...", command=self.browse_output_dir)
        browse_btn.grid(row=0, column=2, padx=(5, 0))
        
        output_frame.grid_columnconfigure(1, weight=1)
        
        # Selection buttons frame
        selection_btn_frame = ttk.Frame(main_frame)
        selection_btn_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(selection_btn_frame, text="Select All", 
                  command=self.select_all).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(selection_btn_frame, text="Clear All", 
                  command=self.clear_all).grid(row=0, column=1, padx=(0, 5))
        ttk.Button(selection_btn_frame, text="First Solar Only", 
                  command=self.select_first_solar).grid(row=0, column=2, padx=(0, 5))
        ttk.Button(selection_btn_frame, text="Standard Only", 
                  command=self.select_standard).grid(row=0, column=3, padx=(0, 5))
        
        # Main action buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, sticky=tk.E)
        
        ttk.Button(button_frame, text="Generate Selected", 
                  command=self.generate_selected).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(button_frame, text="Generate All", 
                  command=self.generate_all).grid(row=0, column=1, padx=(0, 5))
        ttk.Button(button_frame, text="Close", 
                  command=self.close_dialog).grid(row=0, column=2)
    
    def setup_harness_treeview(self, parent_frame):
        """Set up the treeview for harness selection"""
        # Create treeview with scrollbars
        tree_frame = ttk.Frame(parent_frame)
        tree_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        
        # Treeview with checkboxes (simulated with tags)
        columns = ('selected', 'part_number', 'category', 'strings', 'polarity', 'spacing')
        self.harness_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=12)
        
        # Configure columns
        self.harness_tree.heading('selected', text='Select')
        self.harness_tree.heading('part_number', text='Part Number')
        self.harness_tree.heading('category', text='Category')
        self.harness_tree.heading('strings', text='Strings')
        self.harness_tree.heading('polarity', text='Polarity')
        self.harness_tree.heading('spacing', text='Spacing')
        
        self.harness_tree.column('selected', width=60, anchor='center')
        self.harness_tree.column('part_number', width=150)
        self.harness_tree.column('category', width=100)
        self.harness_tree.column('strings', width=70, anchor='center')
        self.harness_tree.column('polarity', width=80, anchor='center')
        self.harness_tree.column('spacing', width=80, anchor='center')
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.harness_tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.harness_tree.xview)
        self.harness_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid layout
        self.harness_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Populate with harness data
        self.populate_harness_tree()
        
        # Bind click events
        self.harness_tree.bind('<Button-1>', self.on_tree_click)
        self.harness_tree.bind('<Double-1>', self.toggle_selection)
    
    def populate_harness_tree(self):
        """Populate the treeview with harness data"""
        self.selected_harnesses = set()
        
        # Filter and sort harnesses by category then part number
        valid_harnesses = []
        for part_number, spec in self.generator.harness_library.items():
            # Skip comment entries and non-dictionary values
            if part_number.startswith('*comment*') or not isinstance(spec, dict):
                continue
            valid_harnesses.append((part_number, spec))
        
        # Sort by category then part number
        sorted_harnesses = sorted(valid_harnesses, 
                                key=lambda x: (x[1].get('category', ''), x[0]))
        
        for part_number, spec in sorted_harnesses:
            values = (
                '☐',  # Unchecked checkbox symbol
                part_number,
                spec.get('category', 'Unknown'),
                str(spec.get('num_strings', '')),
                spec.get('polarity', '').title(),
                f"{spec.get('string_spacing_ft', '')}'"
            )
    
    def on_tree_click(self, event):
        """Handle tree click events"""
        item = self.harness_tree.identify_row(event.y)
        column = self.harness_tree.identify_column(event.x)
        
        # If clicked on the select column, toggle selection
        if item and column == '#1':  # First column is select
            self.toggle_item_selection(item)
    
    def toggle_selection(self, event):
        """Toggle selection on double-click"""
        item = self.harness_tree.focus()
        if item:
            self.toggle_item_selection(item)
    
    def toggle_item_selection(self, item):
        """Toggle the selection state of a tree item"""
        current_values = list(self.harness_tree.item(item, 'values'))
        part_number = current_values[1]
        
        if current_values[0] == '☐':  # Currently unchecked
            current_values[0] = '☑'  # Check it
            self.harness_tree.item(item, values=current_values, tags=('checked',))
            self.selected_harnesses.add(part_number)
        else:  # Currently checked
            current_values[0] = '☐'  # Uncheck it
            self.harness_tree.item(item, values=current_values, tags=('unchecked',))
            self.selected_harnesses.discard(part_number)
        
        self.update_selection_count()
    
    def select_all(self):
        """Select all harnesses"""
        for item in self.harness_tree.get_children():
            current_values = list(self.harness_tree.item(item, 'values'))
            part_number = current_values[1]
            
            current_values[0] = '☑'
            self.harness_tree.item(item, values=current_values, tags=('checked',))
            self.selected_harnesses.add(part_number)
        
        self.update_selection_count()
    
    def clear_all(self):
        """Clear all selections"""
        for item in self.harness_tree.get_children():
            current_values = list(self.harness_tree.item(item, 'values'))
            current_values[0] = '☐'
            self.harness_tree.item(item, values=current_values, tags=('unchecked',))
        
        self.selected_harnesses.clear()
        self.update_selection_count()
    
    def select_first_solar(self):
        """Select only First Solar harnesses"""
        self.clear_all()
        
        for item in self.harness_tree.get_children():
            current_values = list(self.harness_tree.item(item, 'values'))
            category = current_values[2]
            part_number = current_values[1]
            
            if category == 'First Solar':
                current_values[0] = '☑'
                self.harness_tree.item(item, values=current_values, tags=('checked',))
                self.selected_harnesses.add(part_number)
        
        self.update_selection_count()
    
    def select_standard(self):
        """Select only Standard harnesses"""
        self.clear_all()
        
        for item in self.harness_tree.get_children():
            current_values = list(self.harness_tree.item(item, 'values'))
            category = current_values[2]
            part_number = current_values[1]
            
            if category == 'Standard':
                current_values[0] = '☑'
                self.harness_tree.item(item, values=current_values, tags=('checked',))
                self.selected_harnesses.add(part_number)
        
        self.update_selection_count()
    
    def update_selection_count(self):
        """Update the window title with selection count"""
        count = len(self.selected_harnesses)
        total = len(self.harness_tree.get_children())
        self.title(f"Generate Harness Drawings ({count}/{total} selected)")
    
    def browse_output_dir(self):
        """Browse for output directory"""
        directory = filedialog.askdirectory(
            title="Select Output Directory",
            initialdir=self.output_dir_var.get()
        )
        if directory:
            self.output_dir_var.set(directory)
    
    def generate_selected(self):
        """Generate drawings for selected harnesses"""
        if not self.selected_harnesses:
            messagebox.showwarning("Warning", "No harnesses selected")
            return
        
        output_dir = self.output_dir_var.get().strip()
        if not output_dir:
            messagebox.showerror("Error", "Please specify an output directory")
            return
        
        self.generate_drawings(list(self.selected_harnesses), output_dir)
    
    def generate_all(self):
        """Generate drawings for all harnesses"""
        output_dir = self.output_dir_var.get().strip()
        if not output_dir:
            messagebox.showerror("Error", "Please specify an output directory")
            return
        
        all_harnesses = list(self.available_harnesses.keys())
        self.generate_drawings(all_harnesses, output_dir)
    
    def generate_drawings(self, harness_list: List[str], output_dir: str):
        """Generate drawings for the specified harnesses"""
        try:
            # Create output directory if it doesn't exist
            os.makedirs(output_dir, exist_ok=True)
            
            # Show progress dialog
            progress_window = self.create_progress_window(len(harness_list))
            
            success_count = 0
            failed_harnesses = []
            
            for i, part_number in enumerate(harness_list):
                # Update progress
                progress_window.update_progress(i + 1, f"Generating {part_number}...")
                
                # Generate the drawing
                if self.generator.generate_harness_drawing(part_number, output_dir):
                    success_count += 1
                else:
                    failed_harnesses.append(part_number)
            
            # Close progress window
            progress_window.destroy()
            
            # Show results
            self.show_generation_results(success_count, failed_harnesses, output_dir)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate drawings: {str(e)}")
    
    def create_progress_window(self, total_items: int):
        """Create a progress window for generation"""
        progress_window = tk.Toplevel(self)
        progress_window.title("Generating Drawings...")
        progress_window.geometry("400x150")
        progress_window.transient(self)
        progress_window.grab_set()
        
        # Center on parent
        x = self.winfo_rootx() + 200
        y = self.winfo_rooty() + 200
        progress_window.geometry(f"+{x}+{y}")
        
        frame = ttk.Frame(progress_window, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Progress label
        progress_window.status_label = ttk.Label(frame, text="Preparing...")
        progress_window.status_label.pack(pady=(0, 10))
        
        # Progress bar
        progress_window.progress_bar = ttk.Progressbar(frame, length=300, mode='determinate')
        progress_window.progress_bar.pack(pady=(0, 10))
        progress_window.progress_bar['maximum'] = total_items
        
        # Progress text
        progress_window.progress_text = ttk.Label(frame, text="0 of 0")
        progress_window.progress_text.pack()
        
        def update_progress(current: int, status: str):
            progress_window.progress_bar['value'] = current
            progress_window.status_label.config(text=status)
            progress_window.progress_text.config(text=f"{current} of {total_items}")
            progress_window.update()
        
        progress_window.update_progress = update_progress
        
        return progress_window
    
    def show_generation_results(self, success_count: int, failed_harnesses: List[str], output_dir: str):
        """Show the results of drawing generation"""
        if failed_harnesses:
            failed_list = "\n".join(failed_harnesses)
            message = (f"Generated {success_count} drawings successfully.\n\n"
                      f"Failed to generate {len(failed_harnesses)} drawings:\n{failed_list}\n\n"
                      f"Output directory: {output_dir}")
            messagebox.showwarning("Generation Complete", message)
        else:
            message = (f"Successfully generated {success_count} harness drawings!\n\n"
                      f"Output directory: {output_dir}")
            messagebox.showinfo("Generation Complete", message)
        
        # Ask if user wants to open the output directory
        if messagebox.askyesno("Open Directory", "Would you like to open the output directory?"):
            self.open_output_directory(output_dir)
    
    def open_output_directory(self, output_dir: str):
        """Open the output directory in the system file manager"""
        try:
            import subprocess
            import platform
            
            if platform.system() == "Windows":
                os.startfile(output_dir)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", output_dir])
            else:  # Linux and others
                subprocess.run(["xdg-open", output_dir])
        except Exception as e:
            print(f"Could not open directory: {str(e)}")
    
    def close_dialog(self):
        """Close the dialog"""
        self.destroy()
    
    def center_on_parent(self):
        """Center the dialog on its parent window"""
        self.update_idletasks()
        
        # Get parent window position and size
        parent_x = self.parent.winfo_rootx()
        parent_y = self.parent.winfo_rooty()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()
        
        # Get dialog size
        dialog_width = self.winfo_width()
        dialog_height = self.winfo_height()
        
        # Calculate position to center on parent
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        self.geometry(f"+{x}+{y}")