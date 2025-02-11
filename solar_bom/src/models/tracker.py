from dataclasses import dataclass
from typing import List, Optional, Tuple
from .module import ModuleSpec, ModuleOrientation

@dataclass
class TrackerPosition:
    """Data class representing a tracker's position in a block"""
    x: float  # X coordinate in meters
    y: float  # Y coordinate in meters
    rotation: float  # Rotation angle in degrees
    template: 'TrackerTemplate'  # Forward reference to avoid circular import

@dataclass
class TrackerTemplate:
    """Data class representing a solar tracker template configuration"""
    # Required parameters
    template_name: str
    module_spec: ModuleSpec
    module_orientation: ModuleOrientation
    modules_per_string: int
    strings_per_tracker: int
    
    # Optional parameters with defaults
    description: Optional[str] = None
    module_spacing_m: float = 0.01  # Default gap between modules
    motor_gap_m: float = 1.0  # Default gap for motor/drive
    
    def validate(self) -> bool:
        """
        Validate tracker template configuration
        Returns True if valid, raises ValueError if invalid
        """
        if self.modules_per_string <= 0:
            raise ValueError("Modules per string must be positive")
            
        if self.strings_per_tracker <= 0:
            raise ValueError("Strings per tracker must be positive")
            
        if self.module_spacing_m < 0:
            raise ValueError("Module spacing cannot be negative")
            
        if self.motor_gap_m < 0:
            raise ValueError("Motor gap cannot be negative")
            
        return True
    
    def get_total_modules(self) -> int:
        """Calculate total number of modules on the tracker"""
        return self.modules_per_string * self.strings_per_tracker
    
    def get_physical_dimensions(self) -> Tuple[float, float]:
        """
        Calculate physical dimensions of tracker in meters
        Returns (length, width) tuple
        """
        module_length = self.module_spec.length_mm / 1000
        module_width = self.module_spec.width_mm / 1000
        
        if self.module_orientation == ModuleOrientation.PORTRAIT:
            module_length, module_width = module_width, module_length
            
        # Calculate total length including spacing and motor gap
        total_length = (module_length * self.modules_per_string) + \
                      (self.module_spacing_m * (self.modules_per_string - 1)) + \
                      self.motor_gap_m
                      
        # Calculate total width including module spacing
        total_width = (module_width * self.strings_per_tracker) + \
                     (self.module_spacing_m * (self.strings_per_tracker - 1))
                     
        return (total_length, total_width)
    
    def get_string_positions(self) -> List[List[TrackerPosition]]:
        """
        Calculate positions of all strings on the tracker
        Returns list of lists, where each inner list represents module positions for one string
        """
        module_length = self.module_spec.length_mm / 1000
        module_width = self.module_spec.width_mm / 1000
        
        if self.module_orientation == ModuleOrientation.PORTRAIT:
            module_length, module_width = module_width, module_length
            
        string_positions = []
        
        for string_idx in range(self.strings_per_tracker):
            string = []
            y_pos = string_idx * (module_width + self.module_spacing_m)
            
            for module_idx in range(self.modules_per_string):
                # Add motor gap after halfway point
                motor_offset = self.motor_gap_m if module_idx >= self.modules_per_string / 2 else 0
                x_pos = module_idx * (module_length + self.module_spacing_m) + motor_offset
                
                string.append(TrackerPosition(
                    x=x_pos,
                    y=y_pos,
                    rotation=0.0
                ))
            
            string_positions.append(string)
            
        return string_positions
    
    def __str__(self) -> str:
        dims = self.get_physical_dimensions()
        return (f"{self.template_name} - {self.get_total_modules()} modules "
                f"({dims[0]:.1f}m x {dims[1]:.1f}m)")