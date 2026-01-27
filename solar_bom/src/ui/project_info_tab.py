import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Callable
from datetime import datetime
import uuid


class ProjectInfoTab(ttk.Frame):
    """UI component for displaying and editing project information"""
    
    def __init__(self, parent, current_project=None, on_project_changed: Optional[Callable] = None):
        super().__init__(parent)
        self.parent = parent
        self.current_project = current_project
        self.on_project_changed = on_project_changed
        
        # Track the currently selected estimate ID
        self.current_estimate_id = None
        
        self.setup_ui()
        self.load_project_data()
    
    def setup_ui(self):
        """Create and arrange UI components"""
        # Main container with padding
        main_container = ttk.Frame(self, padding="10")
        main_container.pack(fill='both', expand=True)
        
        # Title
        title_label = ttk.Label(main_container, text="Project Information", font=('Helvetica', 14, 'bold'))
        title_label.pack(anchor='w', pady=(0, 15))
        
        # Project details frame
        details_frame = ttk.LabelFrame(main_container, text="Project Details", padding="10")
        details_frame.pack(fill='x', pady=(0, 15))
        
        # Configure grid columns
        details_frame.columnconfigure(1, weight=1)
        
        # Project Name
        ttk.Label(details_frame, text="Project Name:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.name_var = tk.StringVar()
        self.name_entry = ttk.Entry(details_frame, textvariable=self.name_var, width=50)
        self.name_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        self.name_var.trace_add('write', self._on_field_changed)
        
        # Description
        ttk.Label(details_frame, text="Description:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.desc_var = tk.StringVar()
        self.desc_entry = ttk.Entry(details_frame, textvariable=self.desc_var, width=50)
        self.desc_entry.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        self.desc_var.trace_add('write', self._on_field_changed)
        
        # Client
        ttk.Label(details_frame, text="Client:").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        self.client_var = tk.StringVar()
        self.client_entry = ttk.Entry(details_frame, textvariable=self.client_var, width=50)
        self.client_entry.grid(row=2, column=1, padx=5, pady=5, sticky='ew')
        self.client_var.trace_add('write', self._on_field_changed)
        
        # Location
        ttk.Label(details_frame, text="Location:").grid(row=3, column=0, padx=5, pady=5, sticky='w')
        self.location_var = tk.StringVar()
        self.location_entry = ttk.Entry(details_frame, textvariable=self.location_var, width=50)
        self.location_entry.grid(row=3, column=1, padx=5, pady=5, sticky='ew')
        self.location_var.trace_add('write', self._on_field_changed)
        
        # Notes
        ttk.Label(details_frame, text="Notes:").grid(row=4, column=0, padx=5, pady=5, sticky='nw')
        self.notes_text = tk.Text(details_frame, height=4, width=50, wrap='word')
        self.notes_text.grid(row=4, column=1, padx=5, pady=5, sticky='ew')
        self.notes_text.bind('<KeyRelease>', self._on_notes_changed)
        
        # Dates frame (read-only)
        dates_frame = ttk.LabelFrame(main_container, text="Project Dates", padding="10")
        dates_frame.pack(fill='x', pady=(0, 15))
        
        # Created date
        ttk.Label(dates_frame, text="Created:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.created_label = ttk.Label(dates_frame, text="-", foreground='gray')
        self.created_label.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        
        # Modified date
        ttk.Label(dates_frame, text="Last Modified:").grid(row=0, column=2, padx=20, pady=5, sticky='w')
        self.modified_label = ttk.Label(dates_frame, text="-", foreground='gray')
        self.modified_label.grid(row=0, column=3, padx=5, pady=5, sticky='w')
        
        # Tools frame
        tools_frame = ttk.LabelFrame(main_container, text="Project Tools", padding="10")
        tools_frame.pack(fill='x', pady=(0, 15))
        
        # Quick Estimate section
        estimate_frame = ttk.Frame(tools_frame)
        estimate_frame.pack(fill='x', pady=5)
        
        ttk.Label(estimate_frame, text="Quick Estimate:").pack(side='left', padx=(0, 10))
        
        # Estimate dropdown (editable combobox)
        self.estimate_var = tk.StringVar()
        self.estimate_combo = ttk.Combobox(
            estimate_frame, 
            textvariable=self.estimate_var, 
            width=40,
            state='normal'  # Editable
        )
        self.estimate_combo.pack(side='left', padx=(0, 10))
        self.estimate_combo.bind('<<ComboboxSelected>>', self._on_estimate_selected)
        self.estimate_combo.bind('<Return>', self._on_estimate_name_changed)
        self.estimate_combo.bind('<FocusOut>', self._on_estimate_name_changed)
        
        # Open button
        self.open_estimate_btn = ttk.Button(
            estimate_frame, 
            text="Open", 
            command=self.open_quick_estimate,
            width=8
        )
        self.open_estimate_btn.pack(side='left', padx=(0, 5))
        
        # New button
        self.new_estimate_btn = ttk.Button(
            estimate_frame, 
            text="New", 
            command=self.new_quick_estimate,
            width=8
        )
        self.new_estimate_btn.pack(side='left', padx=(0, 5))
        
        # Delete button
        self.delete_estimate_btn = ttk.Button(
            estimate_frame, 
            text="Delete", 
            command=self.delete_quick_estimate,
            width=8
        )
        self.delete_estimate_btn.pack(side='left')
        
        # Description for Quick Estimate
        ttk.Label(
            tools_frame, 
            text="Early-stage BOM estimation for bid and preliminary designs",
            foreground='gray'
        ).pack(anchor='w', pady=(5, 0))
    
    def load_project_data(self):
        """Load project data into the form fields"""
        if not self.current_project:
            return
        
        metadata = self.current_project.metadata
        
        # Set field values (temporarily disable trace)
        self._loading = True
        
        self.name_var.set(metadata.name or "")
        self.desc_var.set(metadata.description or "")
        self.client_var.set(metadata.client or "")
        self.location_var.set(metadata.location or "")
        
        # Notes
        self.notes_text.delete('1.0', 'end')
        if metadata.notes:
            self.notes_text.insert('1.0', metadata.notes)
        
        # Dates
        if metadata.created_date:
            self.created_label.config(text=metadata.created_date.strftime("%Y-%m-%d %H:%M"))
        if metadata.modified_date:
            self.modified_label.config(text=metadata.modified_date.strftime("%Y-%m-%d %H:%M"))
        
        # Load quick estimates dropdown
        self._refresh_estimate_dropdown()
        
        self._loading = False
    
    def _refresh_estimate_dropdown(self):
        """Refresh the quick estimate dropdown with saved estimates"""
        if not self.current_project:
            return
        
        estimates = self.current_project.quick_estimates
        
        # Build list of estimate names for dropdown
        estimate_names = []
        self._estimate_id_map = {}  # Map display name to ID
        
        for est_id, est_data in estimates.items():
            name = est_data.get('name', 'Unnamed Estimate')
            estimate_names.append(name)
            self._estimate_id_map[name] = est_id
        
        self.estimate_combo['values'] = estimate_names
        
        # Select the most recently modified estimate, or first one
        if estimates:
            # Find most recently modified
            most_recent_id = None
            most_recent_date = None
            for est_id, est_data in estimates.items():
                mod_date = est_data.get('modified_date')
                if mod_date:
                    if most_recent_date is None or mod_date > most_recent_date:
                        most_recent_date = mod_date
                        most_recent_id = est_id
            
            if most_recent_id is None:
                most_recent_id = list(estimates.keys())[0]
            
            self.current_estimate_id = most_recent_id
            self.estimate_var.set(estimates[most_recent_id].get('name', 'Unnamed Estimate'))
        else:
            self.current_estimate_id = None
            self.estimate_var.set('')
    
    def _on_estimate_selected(self, event=None):
        """Handle selection of an estimate from the dropdown"""
        selected_name = self.estimate_var.get()
        if selected_name in self._estimate_id_map:
            self.current_estimate_id = self._estimate_id_map[selected_name]
    
    def _on_estimate_name_changed(self, event=None):
        """Handle renaming of the current estimate"""
        if not self.current_estimate_id or not self.current_project:
            return
        
        new_name = self.estimate_var.get().strip()
        if not new_name:
            return
        
        # Update the estimate name
        if self.current_estimate_id in self.current_project.quick_estimates:
            old_name = self.current_project.quick_estimates[self.current_estimate_id].get('name', '')
            if new_name != old_name:
                self.current_project.quick_estimates[self.current_estimate_id]['name'] = new_name
                self.current_project.quick_estimates[self.current_estimate_id]['modified_date'] = datetime.now().isoformat()
                self._refresh_estimate_dropdown()
                self.estimate_var.set(new_name)
                
                if self.on_project_changed:
                    self.on_project_changed()
    
    def _on_field_changed(self, *args):
        """Handle changes to text entry fields"""
        if getattr(self, '_loading', False):
            return
        self._save_to_project()
    
    def _on_notes_changed(self, event=None):
        """Handle changes to notes text widget"""
        if getattr(self, '_loading', False):
            return
        self._save_to_project()
    
    def _save_to_project(self):
        """Save current form values to the project"""
        if not self.current_project:
            return
        
        # Update metadata
        self.current_project.metadata.name = self.name_var.get()
        self.current_project.metadata.description = self.desc_var.get()
        self.current_project.metadata.client = self.client_var.get()
        self.current_project.metadata.location = self.location_var.get()
        self.current_project.metadata.notes = self.notes_text.get('1.0', 'end-1c')
        
        # Update modified date
        self.current_project.metadata.modified_date = datetime.now()
        self.modified_label.config(
            text=self.current_project.metadata.modified_date.strftime("%Y-%m-%d %H:%M")
        )
        
        # Notify parent of changes
        if self.on_project_changed:
            self.on_project_changed()
    
    def new_quick_estimate(self):
        """Create a new quick estimate"""
        if not self.current_project:
            return
        
        # Generate new ID and default name
        estimate_id = f"estimate_{uuid.uuid4().hex[:8]}"
        estimate_num = len(self.current_project.quick_estimates) + 1
        estimate_name = f"Estimate {estimate_num}"
        
        # Create new estimate with default structure
        new_estimate = {
            'name': estimate_name,
            'created_date': datetime.now().isoformat(),
            'modified_date': datetime.now().isoformat(),
            'module_width_mm': 1134,
            'modules_per_string': 28,
            'subarrays': {}
        }
        
        # Add to project
        self.current_project.quick_estimates[estimate_id] = new_estimate
        
        # Refresh dropdown and select new estimate
        self._refresh_estimate_dropdown()
        self.current_estimate_id = estimate_id
        self.estimate_var.set(estimate_name)
        
        # Save and open
        if self.on_project_changed:
            self.on_project_changed()
        
        self.open_quick_estimate()
    
    def open_quick_estimate(self):
        """Open the Quick Estimate dialog"""
        if not self.current_project:
            return
        
        # If no estimate selected, create one
        if not self.current_estimate_id:
            self.new_quick_estimate()
            return
        
        from .quick_estimate import QuickEstimateDialog
        
        dialog = QuickEstimateDialog(
            self.winfo_toplevel(), 
            self.current_project,
            self.current_estimate_id,
            on_save=self._on_estimate_saved
        )
    
    def _on_estimate_saved(self):
        """Callback when estimate is saved from dialog"""
        self._refresh_estimate_dropdown()
        if self.current_estimate_id and self.current_estimate_id in self.current_project.quick_estimates:
            self.estimate_var.set(self.current_project.quick_estimates[self.current_estimate_id].get('name', ''))
        
        if self.on_project_changed:
            self.on_project_changed()
    
    def delete_quick_estimate(self):
        """Delete the currently selected quick estimate"""
        if not self.current_project or not self.current_estimate_id:
            messagebox.showinfo("No Estimate", "No estimate selected to delete.")
            return
        
        estimate_name = self.current_project.quick_estimates.get(
            self.current_estimate_id, {}
        ).get('name', 'this estimate')
        
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{estimate_name}'?"):
            del self.current_project.quick_estimates[self.current_estimate_id]
            self.current_estimate_id = None
            self._refresh_estimate_dropdown()
            
            if self.on_project_changed:
                self.on_project_changed()
    
    def set_project(self, project):
        """Update the current project and refresh the display"""
        self.current_project = project
        self.load_project_data()