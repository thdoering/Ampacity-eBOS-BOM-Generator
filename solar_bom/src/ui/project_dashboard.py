import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import os
from typing import Optional, Callable, Dict, List, Tuple
from datetime import datetime
from ..utils.project_manager import ProjectManager
from ..models.project import Project, ProjectMetadata

class ProjectDashboard(ttk.Frame):
    """UI component for project dashboard and management"""
    
    def __init__(self, parent, on_project_selected: Optional[Callable[[Project], None]] = None):
        super().__init__(parent)
        self.parent = parent
        self.on_project_selected = on_project_selected
        
        # Initialize project manager
        self.project_manager = ProjectManager()
        
        # Set up the UI
        self.setup_ui()
        
        # Populate projects
        self.load_projects()
        
    def setup_ui(self):
        """Create and arrange UI components"""
        # Main container with padding
        main_container = ttk.Frame(self, padding="10")
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        main_container.grid_columnconfigure(0, weight=1)
        main_container.grid_rowconfigure(1, weight=1)
        
        # Header - Title and create button
        header_frame = ttk.Frame(main_container)
        header_frame.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        title_label = ttk.Label(header_frame, text="Solar eBOS Project Dashboard", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, sticky=tk.W)
        
        # Create project button
        create_btn = ttk.Button(header_frame, text="New Project", command=self.create_new_project)
        create_btn.grid(row=0, column=1, padx=10, sticky=tk.E)
        
        # Add spacer
        ttk.Separator(main_container, orient='horizontal').grid(
            row=1, column=0, pady=10, sticky=(tk.W, tk.E))
        
        # Recent projects section
        recent_frame = ttk.LabelFrame(main_container, text="Recent Projects", padding="5")
        recent_frame.grid(row=2, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Recent projects will be added dynamically in load_projects method
        self.recent_projects_frame = ttk.Frame(recent_frame)
        self.recent_projects_frame.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Add spacer
        ttk.Separator(main_container, orient='horizontal').grid(
            row=3, column=0, pady=10, sticky=(tk.W, tk.E))
        
        # Projects browser section
        browser_frame = ttk.LabelFrame(main_container, text="All Projects", padding="5")
        browser_frame.grid(row=4, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        browser_frame.grid_columnconfigure(0, weight=1)
        browser_frame.grid_rowconfigure(1, weight=1)
        
        # Search and sort controls
        controls_frame = ttk.Frame(browser_frame)
        controls_frame.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))
        controls_frame.grid_columnconfigure(1, weight=1)
        
        # Search
        ttk.Label(controls_frame, text="Search:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(controls_frame, textvariable=self.search_var)
        search_entry.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
        search_entry.bind('<Return>', lambda e: self.search_projects())
        ttk.Button(controls_frame, text="Search", command=self.search_projects).grid(
            row=0, column=2, padx=5, pady=5)
        
        # Sort options
        ttk.Label(controls_frame, text="Sort by:").grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        self.sort_var = tk.StringVar(value='modified')
        sort_combo = ttk.Combobox(controls_frame, textvariable=self.sort_var, state='readonly')
        sort_combo['values'] = ['Name', 'Last Modified', 'Date Created', 'Client']
        sort_combo.current(1)  # Default to Last Modified
        sort_combo.grid(row=0, column=4, padx=5, pady=5)
        sort_combo.bind('<<ComboboxSelected>>', lambda e: self.load_projects())
        
        # Projects treeview
        columns = ('name', 'client', 'modified', 'created')
        self.projects_tree = ttk.Treeview(browser_frame, columns=columns, show='headings')
        self.projects_tree.grid(row=1, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure treeview columns
        self.projects_tree.column('name', width=200, anchor='w')
        self.projects_tree.column('client', width=150, anchor='w')
        self.projects_tree.column('modified', width=150, anchor='w')
        self.projects_tree.column('created', width=150, anchor='w')
        
        # Set headings
        self.projects_tree.heading('name', text='Project Name')
        self.projects_tree.heading('client', text='Client')
        self.projects_tree.heading('modified', text='Last Modified')
        self.projects_tree.heading('created', text='Date Created')
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(browser_frame, orient="vertical", command=self.projects_tree.yview)
        scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.projects_tree.configure(yscrollcommand=scrollbar.set)
        
        # Double-click to open project
        self.projects_tree.bind('<Double-1>', self.on_project_double_click)
        
        # Project actions
        actions_frame = ttk.Frame(browser_frame)
        actions_frame.grid(row=2, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        ttk.Button(actions_frame, text="Open", command=self.open_selected_project).grid(
            row=0, column=0, padx=2)
        ttk.Button(actions_frame, text="Delete", command=self.delete_selected_project).grid(
            row=0, column=1, padx=2)
        ttk.Button(actions_frame, text="Copy", command=self.copy_selected_project).grid(
            row=0, column=2, padx=2)
        
    def load_projects(self):
        """Load and display projects in the UI"""
        # Clear existing projects
        for child in self.recent_projects_frame.winfo_children():
            child.destroy()
            
        # Clear treeview
        for item in self.projects_tree.get_children():
            self.projects_tree.delete(item)
            
        # Get projects sorted by the selected criteria
        sort_option = self.sort_var.get()
        sort_map = {
            'Name': 'name',
            'Last Modified': 'modified',
            'Date Created': 'created',
            'Client': 'client'
        }
        sort_by = sort_map.get(sort_option, 'modified')
        
        projects = self.project_manager.list_projects(sort_by=sort_by)
        
        # Load recent projects
        recent_projects = self.project_manager.get_recent_projects()
        
        if not recent_projects:
            ttk.Label(self.recent_projects_frame, 
                    text="No recent projects. Create a new project to get started.").grid(
                row=0, column=0, padx=5, pady=10)
        else:
            # Create a card for each recent project
            for i, (filepath, metadata) in enumerate(recent_projects):
                self._create_project_card(self.recent_projects_frame, filepath, metadata, i)
        
        # Load all projects into treeview
        for filepath, metadata in projects:
            # Format dates
            modified = metadata.modified_date.strftime("%Y-%m-%d %H:%M")
            created = metadata.created_date.strftime("%Y-%m-%d %H:%M")
            
            # Add to treeview with filepath as hidden id
            self.projects_tree.insert('', 'end', values=(
                metadata.name, 
                metadata.client or "",
                modified,
                created
            ), tags=(filepath,))
    
    def _create_project_card(self, parent, filepath, metadata, position):
        """Create a project card UI element"""
        card = ttk.Frame(parent, borderwidth=1, relief="solid", padding="5")
        card.grid(row=position // 3, column=position % 3, padx=10, pady=10, sticky=(tk.W, tk.E))
        
        # Project name with larger font
        name_label = ttk.Label(card, text=metadata.name, font=('Arial', 12, 'bold'))
        name_label.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        # Client
        if metadata.client:
            ttk.Label(card, text=f"Client: {metadata.client}").grid(
                row=1, column=0, columnspan=2, sticky=tk.W)
        
        # Last modified date
        modified = metadata.modified_date.strftime("%Y-%m-%d %H:%M")
        ttk.Label(card, text=f"Modified: {modified}").grid(
            row=2, column=0, columnspan=2, sticky=tk.W)
        
        # Open button
        open_btn = ttk.Button(card, text="Open", 
                             command=lambda f=filepath: self.open_project(f))
        open_btn.grid(row=3, column=0, pady=(5, 0), sticky=tk.W)
        
        # Delete button
        del_btn = ttk.Button(card, text="Delete", 
                            command=lambda f=filepath: self.delete_project(f))
        del_btn.grid(row=3, column=1, pady=(5, 0), sticky=tk.E)

        # Add copy button after delete button
        copy_btn = ttk.Button(card, text="Copy", 
                            command=lambda f=filepath: self.copy_project(f))
        copy_btn.grid(row=3, column=2, pady=(5, 0), sticky=tk.E)
    
    def search_projects(self):
        """Search projects based on search query"""
        query = self.search_var.get()
        if not query:
            self.load_projects()  # If empty, just reload all projects
            return
            
        # Clear treeview
        for item in self.projects_tree.get_children():
            self.projects_tree.delete(item)
            
        # Search projects
        results = self.project_manager.search_projects(query)
        
        # Load results into treeview
        for filepath, metadata in results:
            # Format dates
            modified = metadata.modified_date.strftime("%Y-%m-%d %H:%M")
            created = metadata.created_date.strftime("%Y-%m-%d %H:%M")
            
            # Add to treeview with filepath as hidden id
            self.projects_tree.insert('', 'end', values=(
                metadata.name, 
                metadata.client or "",
                modified,
                created
            ), tags=(filepath,))
    
    def create_new_project(self):
        """Open dialog to create a new project"""
        # Create a dialog for new project details
        dialog = ProjectDialog(self.parent, title="Create New Project")
        
        if dialog.result:
            name, description, client, location, notes = dialog.result
            
            # Create and save the project
            project = self.project_manager.create_project(
                name=name,
                description=description,
                client=client,
                location=location,
                notes=notes
            )
            
            if self.project_manager.save_project(project):
                messagebox.showinfo("Success", f"Project '{name}' created successfully")
                self.load_projects()  # Refresh project list
                
                # Open the new project if callback is set
                if self.on_project_selected:
                    self.on_project_selected(project)
            else:
                messagebox.showerror("Error", f"Failed to create project '{name}'")
    
    def open_project(self, filepath):
        """Open a project by filepath"""
        project = self.project_manager.load_project(filepath)
        
        if project:
            if self.on_project_selected:
                self.on_project_selected(project)
        else:
            messagebox.showerror("Error", "Failed to open project")
    
    def delete_project(self, filepath):
        """Delete a project by filepath"""
        # Get project name for confirmation
        project_name = "this project"
        for item in self.projects_tree.get_children():
            if self.projects_tree.item(item, 'tags')[0] == filepath:
                project_name = self.projects_tree.item(item, 'values')[0]
                break
        
        # Confirm deletion
        if messagebox.askyesno("Confirm Delete", 
                             f"Are you sure you want to delete '{project_name}'?"):
            if self.project_manager.delete_project(filepath):
                messagebox.showinfo("Success", f"Project '{project_name}' deleted successfully")
                self.load_projects()  # Refresh project list
            else:
                messagebox.showerror("Error", f"Failed to delete project '{project_name}'")
    
    def on_project_double_click(self, event):
        """Handle double-click on a project in the treeview"""
        item = self.projects_tree.identify_row(event.y)
        if item:
            # Get filepath from item tags
            filepath = self.projects_tree.item(item, 'tags')[0]
            self.open_project(filepath)
    
    def open_selected_project(self):
        """Open the currently selected project in the treeview"""
        selection = self.projects_tree.selection()
        if selection:
            item = selection[0]
            filepath = self.projects_tree.item(item, 'tags')[0]
            self.open_project(filepath)
        else:
            messagebox.showinfo("No Selection", "Please select a project to open")
    
    def delete_selected_project(self):
        """Delete the currently selected project in the treeview"""
        selection = self.projects_tree.selection()
        if selection:
            item = selection[0]
            filepath = self.projects_tree.item(item, 'tags')[0]
            self.delete_project(filepath)
        else:
            messagebox.showinfo("No Selection", "Please select a project to delete")

    def copy_selected_project(self):
        """Copy the currently selected project in the treeview"""
        selection = self.projects_tree.selection()
        if selection:
            item = selection[0]
            filepath = self.projects_tree.item(item, 'tags')[0]
            self.copy_project(filepath)
        else:
            messagebox.showinfo("No Selection", "Please select a project to copy")

    def copy_project(self, filepath):
        """Copy a project by filepath"""
        # Load the original project for the dialog
        original_project = self.project_manager.load_project(filepath)
        if not original_project:
            messagebox.showerror("Error", "Failed to load original project")
            return
        
        # Create copy dialog
        dialog = ProjectDialog(self.parent, title="Copy Project", 
                            copy_mode=True, original_project=original_project)
        
        if dialog.result:
            new_name = dialog.result[0]  # Get the name from dialog result
            
            # Validate name doesn't already exist
            if self.project_manager.project_name_exists(new_name):
                messagebox.showerror("Error", f"A project named '{new_name}' already exists. Please choose a different name.")
                return
            
            # Perform the copy
            if self.project_manager.copy_project(filepath, new_name):
                self.load_projects()  # Refresh project list
            else:
                messagebox.showerror("Error", f"Failed to copy project")


class ProjectDialog(tk.Toplevel):
    """Dialog for creating or editing projects"""
    
    def __init__(self, parent, title="Project Details", project=None, copy_mode=False, original_project=None):
        super().__init__(parent)
        self.parent = parent
        self.title(title)
        self.result = None
        
        # Store copy mode and original project
        self.copy_mode = copy_mode
        self.original_project = original_project
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
        # Create and place widgets - this creates all the variables
        self.setup_ui(project)
        
        # NOW set the values if in copy mode (after setup_ui has created the variables)
        if self.copy_mode and self.original_project:
            self.name_var.set(f"Copy of {self.original_project.metadata.name}")
            if self.original_project.metadata.description:
                self.desc_var.set(self.original_project.metadata.description)
            if self.original_project.metadata.client:
                self.client_var.set(self.original_project.metadata.client)
            if self.original_project.metadata.location:
                self.location_var.set(self.original_project.metadata.location)
            if self.original_project.metadata.notes:
                self.notes_var.set(self.original_project.metadata.notes)
        
        # Center dialog
        self.geometry("400x400")
        self.resizable(False, False)
        
        # Wait for window to be closed
        self.wait_window(self)
    
    # Add these conversion methods to the class
    def m_to_ft(self, meters):
        return meters * 3.28084
        
    def ft_to_m(self, feet):
        return feet / 3.28084
    
    def setup_ui(self, project=None):
        """Set up dialog UI"""
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Project name
        ttk.Label(main_frame, text="Project Name:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.name_var = tk.StringVar(value=project.metadata.name if project else "")
        ttk.Entry(main_frame, textvariable=self.name_var, width=40).grid(
            row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Description
        ttk.Label(main_frame, text="Description:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.desc_var = tk.StringVar(value=project.metadata.description if project else "")
        ttk.Entry(main_frame, textvariable=self.desc_var, width=40).grid(
            row=1, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Client
        ttk.Label(main_frame, text="Client:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.client_var = tk.StringVar(value=project.metadata.client if project else "")
        ttk.Entry(main_frame, textvariable=self.client_var, width=40).grid(
            row=2, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Location
        ttk.Label(main_frame, text="Location:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        self.location_var = tk.StringVar(value=project.metadata.location if project else "")
        ttk.Entry(main_frame, textvariable=self.location_var, width=40).grid(
            row=3, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Notes
        ttk.Label(main_frame, text="Notes:").grid(row=4, column=0, padx=5, pady=5, sticky=tk.NW)
        self.notes_var = tk.StringVar(value=project.metadata.notes if project else "")
        notes_entry = tk.Text(main_frame, width=40, height=5)
        notes_entry.grid(row=4, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
        if project and project.metadata.notes:
            notes_entry.insert('1.0', project.metadata.notes)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Cancel", command=self.cancel).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="OK", command=lambda: self.ok(notes_entry)).grid(
            row=0, column=1, padx=5)
        
        # Set focus to name entry
        if not self.name_var.get():
            self.after(100, lambda: self.winfo_children()[0].winfo_children()[1].focus())
    
    def ok(self, notes_entry):
        """Validate and store result"""
        name = self.name_var.get().strip()
        
        if not name:
            messagebox.showerror("Error", "Project name is required")
            return
            
        # Get values and store result
        self.result = (
            name,
            self.desc_var.get().strip(),
            self.client_var.get().strip(),
            self.location_var.get().strip(),
            notes_entry.get('1.0', 'end-1c').strip()
        )
        
        self.destroy()
    
    def cancel(self):
        """Cancel dialog"""
        self.result = None
        self.destroy()