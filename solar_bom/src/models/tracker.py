from dataclasses import dataclass, field
from typing import List, Optional, Tuple
from .module import ModuleSpec, ModuleOrientation
from enum import Enum  # Add if not already imported

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

    def _apply_polarity_convention(self, polarity_convention_value: str, device_y: float = None):
        """
        Flip string polarity based on convention. Called after normal string position calculation.
        Default calculation assumes Negative South (positive at top, negative at bottom).
        
        Args:
            polarity_convention_value: String value of PolarityConvention enum
            device_y: Y coordinate of the device in the block (needed for toward-device modes)
        """
        for string in self.strings:
            should_flip = False
            
            if polarity_convention_value == "Negative Always North":
                # Flip all strings: negative goes to top (north), positive to bottom (south)
                should_flip = True
                
            elif polarity_convention_value == "Negative Toward Device":
                if device_y is not None:
                    # Calculate this string's center Y in block coordinates
                    string_center_y = self.y + (string.positive_source_y + string.negative_source_y) / 2
                    # If string center is south of device (higher Y), flip so negative faces north (toward device)
                    # If string center is north of device (lower Y), keep default (negative at south, toward device)
                    should_flip = string_center_y > device_y
                    
            elif polarity_convention_value == "Positive Toward Device":
                if device_y is not None:
                    # Calculate this string's center Y in block coordinates
                    string_center_y = self.y + (string.positive_source_y + string.negative_source_y) / 2
                    # If string center is north of device (lower Y), flip so positive faces south (toward device)
                    # If string center is south of device (higher Y), keep default (positive at north, toward device)
                    should_flip = string_center_y <= device_y
            
            # "Negative Always South" = default behavior, no flip needed
            
            if should_flip:
                # Swap positive and negative Y source positions
                string.positive_source_y, string.negative_source_y = (
                    string.negative_source_y, string.positive_source_y
                )

    def set_polarity_info(self, polarity_convention_value: str, device_y: float = None):
        """
        Store polarity convention info for use during string position calculation.
        
        Args:
            polarity_convention_value: String value of PolarityConvention enum
            device_y: Y coordinate of the device in the block
        """
        self._polarity_convention = polarity_convention_value
        self._device_y = device_y

    def calculate_string_positions(self) -> None:
        """Calculate string positions and their source points"""
        if not self.template:
            print("No template - returning")
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

        # Handle different motor placement types
        if self.template.motor_placement_type == "middle_of_string":
            # Motor is in the middle of a specific string
            current_y = 0
            
            for i in range(self.template.strings_per_tracker):
                if i + 1 == self.template.motor_string_index:  # This string has the motor (1-based index)
                    # Calculate split string dimensions
                    north_modules = self.template.motor_split_north
                    south_modules = self.template.motor_split_south
                    
                    north_height = (north_modules * module_height + 
                                (north_modules - 1) * self.template.module_spacing_m)
                    south_height = (south_modules * module_height + 
                                (south_modules - 1) * self.template.module_spacing_m)
                    
                    # This string spans from current_y to current_y + north_height + gap + south_height
                    y_start = current_y
                    y_end = current_y + north_height + self.template.motor_gap_m + south_height
                    
                    # Move current_y for next string
                    current_y = y_end
                else:
                    # Normal string without motor
                    y_start = current_y
                    y_end = current_y + single_string_height
                    current_y = y_end
                
                # Get wiring mode from project if available
                wiring_mode = 'daisy_chain'  # default
                if hasattr(self, '_project_ref') and hasattr(self._project_ref, 'wiring_mode'):
                    wiring_mode = self._project_ref.wiring_mode
                
                if wiring_mode == 'leapfrog':
                    # In leapfrog mode, both positive and negative connect at top
                    string = StringPosition(
                        index=i,
                        positive_source_x=0,  # Left side of torque tube
                        positive_source_y=y_start,  # Top of string
                        negative_source_x=module_width,  # Right side of torque tube
                        negative_source_y=y_start,  # Also at top of string
                        num_modules=modules_per_string
                    )
                else:
                    # Daisy-chain mode (default)
                    string = StringPosition(
                        index=i,
                        positive_source_x=0,  # Left side of torque tube
                        positive_source_y=y_start,  # Top of string
                        negative_source_x=module_width,  # Right side of torque tube
                        negative_source_y=y_end,  # Bottom of string (for daisy-chain)
                        num_modules=modules_per_string
                    )
                self.strings.append(string)
        else:
            # Original between_strings logic
            for i in range(self.template.strings_per_tracker):
                # For strings above motor (all except last string)
                if i < self.template.strings_per_tracker - 1:
                    y_start = i * single_string_height
                    y_end = y_start + single_string_height
                else:
                    # Last string goes below motor gap
                    y_start = (i * single_string_height) + self.template.motor_gap_m
                    y_end = y_start + single_string_height
                
                wiring_mode = 'daisy_chain'  # default
                if hasattr(self, '_project_ref') and hasattr(self._project_ref, 'wiring_mode'):
                    wiring_mode = self._project_ref.wiring_mode
                
                if wiring_mode == 'leapfrog':
                    # In leapfrog mode, both positive and negative connect at top
                    string = StringPosition(
                        index=i,
                        positive_source_x=0,  # Left side of torque tube
                        positive_source_y=y_start,  # Top of string
                        negative_source_x=module_width,  # Right side of torque tube
                        negative_source_y=y_start,  # Also at top of string (for leapfrog)
                        num_modules=modules_per_string
                    )
                else:
                    # Daisy-chain mode (default)
                    string = StringPosition(
                        index=i,
                        positive_source_x=0,  # Left side of torque tube
                        positive_source_y=y_start,  # Top of string
                        negative_source_x=module_width,  # Right side of torque tube
                        negative_source_y=y_end,  # Bottom of string (for daisy-chain)
                        num_modules=modules_per_string
                    )
                self.strings.append(string)

        # Apply polarity convention if set
        if hasattr(self, '_polarity_convention') and self._polarity_convention:
            self._apply_polarity_convention(
                self._polarity_convention,
                getattr(self, '_device_y', None)
            )

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
    
    # New motor placement options
    motor_placement_type: str = "between_strings"  # "between_strings" or "middle_of_string"
    motor_string_index: int = 1  # Which string (1-based) when middle_of_string
    motor_split_north: int = 0  # Modules north of motor when middle_of_string  
    motor_split_south: int = 0  # Modules south of motor when middle_of_string
    
    # Multi-module-high configuration (1p, 2p, 4p, 1l, 2l, 4l)
    # Each column of stacked modules is a separate string
    modules_high: int = 1  # Number of modules stacked E/W (1, 2, or 4)
    
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
            
        # Validate new motor placement fields
        if self.motor_placement_type not in ["between_strings", "middle_of_string"]:
            raise ValueError("Motor placement type must be 'between_strings' or 'middle_of_string'")
            
        if self.motor_placement_type == "middle_of_string":
            if self.motor_string_index < 1 or self.motor_string_index > self.strings_per_tracker:
                raise ValueError("Motor string index must be between 1 and strings_per_tracker")
            if self.motor_split_north < 0 or self.motor_split_south < 0:
                raise ValueError("Motor split values cannot be negative")
            if self.motor_split_north + self.motor_split_south != self.modules_per_string:
                raise ValueError("Motor split north + south must equal modules_per_string")
        
        # Validate modules_high
        if self.modules_high not in [1, 2, 4]:
            raise ValueError("Modules high must be 1, 2, or 4")
            
        return True
    
    def get_motor_position(self) -> int:
        """Get the motor position"""
        return self.motor_position_after_string
    
    def get_total_modules(self) -> int:
        """Calculate total number of modules on the tracker"""
        return self.modules_per_string * self.strings_per_tracker * self.modules_high
    
    def get_total_strings(self) -> int:
        """Calculate total number of strings on the tracker (each E/W column is a separate string)"""
        return self.strings_per_tracker * self.modules_high
    
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
            # Tracker width is module length × modules_high (stacked E/W)
            total_width = module_length * self.modules_high
        else:  # LANDSCAPE
            # In landscape, module length runs along torque tube
            single_string_length = (module_length * self.modules_per_string) + \
                        (self.module_spacing_m * (self.modules_per_string - 1))
            # Tracker width is module width × modules_high (stacked E/W)
            total_width = module_width * self.modules_high
        
        # Calculate total length based on motor placement type
        if self.motor_placement_type == "middle_of_string":
            # Motor is in the middle of a specific string
            # All strings have same length, but one string has a motor gap in it
            total_length = single_string_length * self.strings_per_tracker + self.motor_gap_m
        else:
            # Original between_strings logic
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
        config_suffix = f"{self.modules_high}{'P' if self.module_orientation == ModuleOrientation.PORTRAIT else 'L'}"
        return (f"{self.template_name} - {self.get_total_modules()} modules "
                f"({dims[0]:.1f}m x {dims[1]:.1f}m) [{config_suffix}]")