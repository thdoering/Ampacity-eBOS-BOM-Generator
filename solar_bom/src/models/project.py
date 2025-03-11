from dataclasses import dataclass, field
from typing import Dict, List, Optional
from datetime import datetime
import json
import os
from pathlib import Path

@dataclass
class ProjectMetadata:
    """Data class for project metadata"""
    name: str
    description: Optional[str] = None
    location: Optional[str] = None
    client: Optional[str] = None
    created_date: datetime = field(default_factory=datetime.now)
    modified_date: datetime = field(default_factory=datetime.now)
    notes: Optional[str] = None

@dataclass
class Project:
    """Data class representing a solar project configuration"""
    metadata: ProjectMetadata
    # Store references to blocks by their IDs
    blocks: Dict[str, str] = field(default_factory=dict)
    # Store selected module IDs for this project
    selected_modules: List[str] = field(default_factory=list)
    # Store selected inverter IDs for this project
    selected_inverters: List[str] = field(default_factory=list)
    
    def update_modified_date(self):
        """Update the last modified date"""
        self.metadata.modified_date = datetime.now()
    
    def to_dict(self) -> dict:
        """Convert project to dictionary for serialization"""
        return {
            'metadata': {
                'name': self.metadata.name,
                'description': self.metadata.description,
                'location': self.metadata.location,
                'client': self.metadata.client,
                'created_date': self.metadata.created_date.isoformat(),
                'modified_date': self.metadata.modified_date.isoformat(),
                'notes': self.metadata.notes
            },
            'blocks': self.blocks,
            'selected_modules': self.selected_modules,
            'selected_inverters': self.selected_inverters
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'Project':
        """Create project instance from dictionary"""
        metadata = ProjectMetadata(
            name=data['metadata']['name'],
            description=data['metadata']['description'],
            location=data['metadata']['location'],
            client=data['metadata']['client'],
            created_date=datetime.fromisoformat(data['metadata']['created_date']),
            modified_date=datetime.fromisoformat(data['metadata']['modified_date']),
            notes=data['metadata']['notes']
        )
        
        return cls(
            metadata=metadata,
            blocks=data.get('blocks', {}),
            selected_modules=data.get('selected_modules', []),
            selected_inverters=data.get('selected_inverters', [])
        )
    
    def save(self, projects_dir: str = 'projects') -> bool:
        """Save project to file"""
        try:
            # Create projects directory if it doesn't exist
            os.makedirs(projects_dir, exist_ok=True)
            
            # Create a valid filename from project name
            filename = "".join(c for c in self.metadata.name if c.isalnum() or c in (' ', '_')).rstrip()
            filename = filename.replace(' ', '_') + '.json'
            filepath = os.path.join(projects_dir, filename)
            
            # Update modified date
            self.update_modified_date()
            
            with open(filepath, 'w') as f:
                json.dump(self.to_dict(), f, indent=2)
                
            return True
        except Exception as e:
            print(f"Error saving project: {str(e)}")
            return False
    
    @classmethod
    def load(cls, filepath: str) -> Optional['Project']:
        """Load project from file"""
        try:
            with open(filepath, 'r') as f:
                data = json.load(f)
                
            return cls.from_dict(data)
        except Exception as e:
            print(f"Error loading project: {str(e)}")
            return None