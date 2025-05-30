from dataclasses import dataclass, field
from typing import List, Optional, Tuple
from .module import ModuleSpec, ModuleOrientation

"""
Solar Tracker Terminology:

1. Tracker Orientation:
   - N/S Tracker: From bird's eye view, tracker appears taller than wide
   - Modules stack vertically (in N/S direction)
   - Torque tube runs along N/S axis

2. String Configuration:
   - String: A vertical sequence of modules on the same torque tube
   - Motor position: Configurable between any strings
   - Multiple strings stack vertically
   - Total modules = modules_per_string × strings_per_tracker

3. Physical Dimensions:
   Portrait Module:
   - total_length = (module_width × modules_per_string) + 
                    (module_spacing × (modules_per_string - 1)) + 
                    motor_gap
   - total_width = module_length

   Landscape Module:
   - total_length = (module_length × modules_per_string) + 
                    (module_spacing × (modules_per_string - 1)) + 
                    motor_gap
   - total_width = module_width
"""

@dataclass
class StringPosition:
    """Represents a single string on a tracker with its source points"""
    index: int  # Index of this string on the tracker
    positive_source_x: float  # X coordinate of positive source point relative to tracker
    positive_source_y: float  # Y coordinate of positive source point relative to tracker
    negative_source_x: float  # X coordinate of negative source point relative to tracker
    negative_source_y: float  # Y coordinate of negative source point relative to tracker
    num_modules: int  # Number of modules in this string

@dataclass
class TrackerPosition:
    """Data class representing a tracker's position in a block"""
    x: float  # X coordinate in meters
    y: float  # Y coordinate in meters
    rotation: float  # Rotation angle in degrees
    template: 'TrackerTemplate'  # Forward reference to avoid circular import
    strings: List[StringPosition] = field(default_factory=list)  # List of strings on this tracker

    def calculate_string_positions(self) -> None:
        """Calculate string positions and their source points"""
        if not self.template:
            return

        # Clear existing strings
        self.strings.clear()

        # Get module dimensions based on orientation
        if self.template.module_orientation == ModuleOrientation.PORTRAIT:
            module_height = self.template.module_spec.width_mm / 1000
            module_width = self.template.module_spec.length_mm / 1000
        else:
            module_height = self.template.module_spec.length_mm / 1000
            module_width = self.template.module_spec.width_mm / 1000

        # Calculate height for a single string of modules
        modules_per_string = self.template.modules_per_string
        single_string_height = (modules_per_string * module_height + 
                            (modules_per_string - 1) * self.template.module_spacing_m)

        # Create string positions
        for i in range(self.template.strings_per_tracker):
            # For strings above motor (all except last string)
            if i < self.template.strings_per_tracker - 1:
                y_start = i * single_string_height
                y_end = y_start + single_string_height
            else:
                # Last string goes below motor gap
                y_start = (i * single_string_height) + self.template.motor_gap_m
                y_end = y_start + single_string_height
                
            string = StringPosition(
                index=i,
                positive_source_x=0,  # Left side of torque tube
                positive_source_y=y_start,  # Top of string
                negative_source_x=module_width,  # Right side of torque tube
                negative_source_y=y_end,  # Bottom of string
                num_modules=modules_per_string
            )
            self.strings.append(string)

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
    motor_position_after_string: int = 0  # Motor position (0 means calculate default)
    
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
        
        if self.motor_position_after_string < 0 or self.motor_position_after_string > self.strings_per_tracker:
            raise ValueError("Motor position must be between 0 and strings_per_tracker")
            
        return True
    
    def get_motor_position(self) -> int:
        """Get the motor position"""
        return self.motor_position_after_string
    
    def get_total_modules(self) -> int:
        """Calculate total number of modules on the tracker"""
        return self.modules_per_string * self.strings_per_tracker
    
    def get_physical_dimensions(self) -> Tuple[float, float]:
        """
        Calculate physical dimensions of tracker in meters.
        For N/S trackers:
        - Length is the vertical dimension (N/S direction)
        - Width is the horizontal dimension (E/W direction)
        
        Length calculation accounts for:
        - All modules in all strings
        - Spacing between modules
        - Motor gap
        - Spacing between strings
        
        Width is simply the module dimension perpendicular to the torque tube.
        
        Returns:
            Tuple[float, float]: (length, width) in meters
        """
        # Convert module dimensions from mm to meters
        module_length = self.module_spec.length_mm / 1000
        module_width = self.module_spec.width_mm / 1000
        
        # Calculate length of a single string including module spacing
        if self.module_orientation == ModuleOrientation.PORTRAIT:
            # In portrait, module width runs along torque tube
            single_string_length = (module_width * self.modules_per_string) + \
                        (self.module_spacing_m * (self.modules_per_string - 1))
            # Tracker width is module length
            total_width = module_length
        else:  # LANDSCAPE
            # In landscape, module length runs along torque tube
            single_string_length = (module_length * self.modules_per_string) + \
                        (self.module_spacing_m * (self.modules_per_string - 1))
            # Tracker width is module width
            total_width = module_width
        
        # Calculate total length including all strings and motor gap
        motor_position = self.get_motor_position()
        strings_above_motor = motor_position
        strings_below_motor = self.strings_per_tracker - motor_position

        if strings_below_motor > 0:
            # Motor gap is only added when there are strings below motor
            total_length = (single_string_length * strings_above_motor) + \
                        self.motor_gap_m + \
                        (single_string_length * strings_below_motor)
        else:
            # No strings below motor, no gap needed
            total_length = single_string_length * strings_above_motor
                        
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