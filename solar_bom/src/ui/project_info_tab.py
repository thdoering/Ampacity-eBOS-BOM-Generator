import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Callable
from datetime import datetime


class ProjectInfoTab(ttk.Frame):
    """UI component for displaying and editing project information"""
    
    def __init__(self, parent, current_project=None, on_project_changed: Optional[Callable] = None):
        super().__init__(parent)
        self.parent = parent
        self.current_project = current_project
        self.on_project_changed = on_project_changed

        self._field_change_after_id = None

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
        
        self._loading = False
    
    def _on_field_changed(self, *args):
        """Handle changes to text entry fields (debounced)"""
        if getattr(self, '_loading', False):
            return
        # Cancel any pending save
        if hasattr(self, '_field_change_after_id') and self._field_change_after_id:
            self.after_cancel(self._field_change_after_id)
        # Schedule save after 1 second of no changes
        self._field_change_after_id = self.after(1000, self._debounced_save)
    
    def _debounced_save(self):
        """Execute the debounced save"""
        self._field_change_after_id = None
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
    
    def set_project(self, project):
        """Update the current project and refresh the display"""
        self.current_project = project
        self.load_project_data()