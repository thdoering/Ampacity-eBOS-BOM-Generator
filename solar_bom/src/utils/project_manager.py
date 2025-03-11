import os
import json
from typing import List, Optional, Dict, Tuple
from datetime import datetime
from pathlib import Path
from ..models.project import Project, ProjectMetadata

class ProjectManager:
    """Utility class for managing solar project files"""
    
    def __init__(self, projects_dir: str = 'projects', max_recent: int = 5):
        """
        Initialize the project manager
        
        Args:
            projects_dir: Directory where projects are stored
            max_recent: Maximum number of recent projects to track
        """
        self.projects_dir = projects_dir
        self.max_recent = max_recent
        self.recent_projects_file = os.path.join(projects_dir, '.recent_projects')
        
        # Create projects directory if it doesn't exist
        os.makedirs(self.projects_dir, exist_ok=True)
        
        # Load recent projects
        self.recent_projects = self._load_recent_projects()
    
    def create_project(self, name: str, description: str = "", location: str = "", 
                     client: str = "", notes: str = "") -> Project:
        """
        Create a new project with given metadata
        
        Returns:
            Project: The newly created project
        """
        metadata = ProjectMetadata(
            name=name,
            description=description,
            location=location,
            client=client,
            notes=notes
        )
        return Project(metadata=metadata)
    
    def save_project(self, project: Project) -> bool:
        """
        Save a project to file and update recent projects
        
        Args:
            project: Project to save
            
        Returns:
            bool: True if successful, False otherwise
        """
        result = project.save(self.projects_dir)
        if result:
            self._add_to_recent(self._get_filepath(project.metadata.name))
        return result
    
    def _get_filepath(self, project_name: str) -> str:
        """Generate filepath for a project based on its name"""
        # Create a valid filename from project name
        filename = "".join(c for c in project_name if c.isalnum() or c in (' ', '_')).rstrip()
        filename = filename.replace(' ', '_') + '.json'
        return os.path.join(self.projects_dir, filename)
    
    def load_project(self, filepath: str) -> Optional[Project]:
        """
        Load a project from file and update recent projects
        
        Args:
            filepath: Path to project file
            
        Returns:
            Project or None: The loaded project or None if loading failed
        """
        project = Project.load(filepath)
        if project:
            self._add_to_recent(filepath)
        return project
    
    def delete_project(self, filepath: str) -> bool:
        """
        Delete a project file
        
        Args:
            filepath: Path to project file
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            if os.path.exists(filepath):
                os.remove(filepath)
                self._remove_from_recent(filepath)
                return True
            return False
        except Exception as e:
            print(f"Error deleting project: {str(e)}")
            return False
    
    def list_projects(self, sort_by: str = 'modified', reverse: bool = True) -> List[Tuple[str, ProjectMetadata]]:
        """
        List all projects in the projects directory
        
        Args:
            sort_by: Field to sort by ('name', 'modified', 'created', 'client')
            reverse: Whether to reverse sort order
            
        Returns:
            List of (filepath, metadata) tuples
        """
        projects = []
        
        for filename in os.listdir(self.projects_dir):
            if filename.endswith('.json') and not filename.startswith('.'):
                filepath = os.path.join(self.projects_dir, filename)
                try:
                    with open(filepath, 'r') as f:
                        data = json.load(f)
                        metadata = ProjectMetadata(
                            name=data['metadata']['name'],
                            description=data['metadata']['description'],
                            location=data['metadata']['location'],
                            client=data['metadata']['client'],
                            created_date=datetime.fromisoformat(data['metadata']['created_date']),
                            modified_date=datetime.fromisoformat(data['metadata']['modified_date']),
                            notes=data['metadata']['notes']
                        )
                        projects.append((filepath, metadata))
                except Exception as e:
                    print(f"Error loading project metadata for {filename}: {str(e)}")
        
        # Sort projects
        if sort_by == 'name':
            projects.sort(key=lambda p: p[1].name.lower(), reverse=reverse)
        elif sort_by == 'modified':
            projects.sort(key=lambda p: p[1].modified_date, reverse=reverse)
        elif sort_by == 'created':
            projects.sort(key=lambda p: p[1].created_date, reverse=reverse)
        elif sort_by == 'client':
            projects.sort(key=lambda p: (p[1].client or "").lower(), reverse=reverse)
        
        return projects
    
    def get_recent_projects(self) -> List[Tuple[str, ProjectMetadata]]:
        """
        Get list of recent projects with metadata
        
        Returns:
            List of (filepath, metadata) tuples for recent projects
        """
        recent_with_metadata = []
        
        for filepath in self.recent_projects:
            if os.path.exists(filepath):
                try:
                    with open(filepath, 'r') as f:
                        data = json.load(f)
                        metadata = ProjectMetadata(
                            name=data['metadata']['name'],
                            description=data['metadata']['description'],
                            location=data['metadata']['location'],
                            client=data['metadata']['client'],
                            created_date=datetime.fromisoformat(data['metadata']['created_date']),
                            modified_date=datetime.fromisoformat(data['metadata']['modified_date']),
                            notes=data['metadata']['notes']
                        )
                        recent_with_metadata.append((filepath, metadata))
                except Exception as e:
                    print(f"Error loading recent project metadata for {filepath}: {str(e)}")
        
        return recent_with_metadata
    
    def search_projects(self, query: str, case_sensitive: bool = False) -> List[Tuple[str, ProjectMetadata]]:
        """
        Search projects by name, description, client or location
        
        Args:
            query: Search query string
            case_sensitive: Whether search is case sensitive
            
        Returns:
            List of (filepath, metadata) tuples matching the query
        """
        if not case_sensitive:
            query = query.lower()
            
        results = []
        all_projects = self.list_projects()
        
        for filepath, metadata in all_projects:
            # Search in relevant fields
            searchable_text = " ".join([
                metadata.name,
                metadata.description or "",
                metadata.client or "",
                metadata.location or "",
                metadata.notes or ""
            ])
            
            if not case_sensitive:
                searchable_text = searchable_text.lower()
                
            if query in searchable_text:
                results.append((filepath, metadata))
                
        return results
    
    def _load_recent_projects(self) -> List[str]:
        """Load list of recent projects from file"""
        if os.path.exists(self.recent_projects_file):
            try:
                with open(self.recent_projects_file, 'r') as f:
                    return json.load(f)
            except Exception:
                return []
        return []
    
    def _save_recent_projects(self):
        """Save list of recent projects to file"""
        try:
            with open(self.recent_projects_file, 'w') as f:
                json.dump(self.recent_projects, f)
        except Exception as e:
            print(f"Error saving recent projects: {str(e)}")
    
    def _add_to_recent(self, filepath: str):
        """Add a project to the recent projects list"""
        # Remove if already in list
        if filepath in self.recent_projects:
            self.recent_projects.remove(filepath)
            
        # Add to front of list
        self.recent_projects.insert(0, filepath)
        
        # Trim to max size
        self.recent_projects = self.recent_projects[:self.max_recent]
        
        # Save updated list
        self._save_recent_projects()
        
    def _remove_from_recent(self, filepath: str):
        """Remove a project from the recent projects list"""
        if filepath in self.recent_projects:
            self.recent_projects.remove(filepath)
            self._save_recent_projects()