from dataclasses import dataclass, field
from typing import Dict, List, Optional
from enum import Enum
from .inverter import InverterSpec
from .tracker import TrackerTemplate, TrackerPosition

class WiringType(Enum):
    """Enumeration of wiring configuration types"""
    HOMERUN = "String Homerun"
    HARNESS = "Wire Harness"

class DeviceType(Enum):
    """Enumeration of downstream device types"""
    STRING_INVERTER = "String Inverter"
    COMBINER_BOX = "Combiner Box"

@dataclass
class CollectionPoint:
    """Data class representing a wiring collection point"""
    x: float
    y: float
    connected_strings: List[int]  # List of string IDs connected to this point
    current_rating: float

@dataclass
class DeviceInputPoint:
    """Represents an input connection point on a device"""
    index: int  # Input number
    x: float  # X coordinate relative to device corner
    y: float  # Y coordinate relative to device corner
    max_current: float  # Maximum current rating for this input
    
@dataclass
class WiringConfig:
    """Data class representing block wiring configuration"""
    wiring_type: WiringType
    positive_collection_points: List[CollectionPoint]
    negative_collection_points: List[CollectionPoint]
    strings_per_collection: Dict[int, int]  # Collection point ID -> number of strings
    cable_routes: Dict[str, List[tuple[float, float]]]  # Route ID -> list of coordinates
    string_cable_size: str = "10 AWG"  # Default value
    harness_cable_size: str = "8 AWG"  # Default value

@dataclass
class BlockConfig:
    """Data class representing a solar block configuration"""
    
    # Block identification and core components must come first (no defaults)
    block_id: str
    inverter: InverterSpec
    tracker_template: TrackerTemplate
    width_m: float
    height_m: float
    row_spacing_m: float  # Distance between tracker rows
    ns_spacing_m: float  # Distance between trackers in north/south direction
    gcr: float  # Ground Coverage Ratio
    device_x: float = 0.0  # X coordinate of device in meters
    device_y: float = 0.0  # Y coordinate of device in meters
    device_spacing_m: float = 1.83  # 6ft in meters default
    input_points: List[DeviceInputPoint] = field(default_factory=list)
    
    # Optional fields with defaults must come after
    description: Optional[str] = None
    tracker_positions: List[TrackerPosition] = field(default_factory=list)
    wiring_config: Optional[WiringConfig] = None
    
    def validate(self) -> bool:
        """
        Validate block configuration
        Returns True if valid, raises ValueError if invalid
        """
        if self.width_m <= 0 or self.height_m <= 0:
            raise ValueError("Block dimensions must be positive")
            
        if self.row_spacing_m <= 0:
            raise ValueError("Row spacing must be positive")
            
        if not (0 < self.gcr <= 1):
            raise ValueError("Ground Coverage Ratio must be between 0 and 1")
            
        # Calculate total number of strings in block
        total_strings = len(self.tracker_positions) * self.tracker_template.strings_per_tracker
        
        # Validate against inverter capacity
        if total_strings > self.inverter.get_total_string_capacity():
            raise ValueError("Total number of strings exceeds inverter capacity")
            
        # Validate tracker positions are within block boundaries
        tracker_dims = self.tracker_template.get_physical_dimensions()
        for pos in self.tracker_positions:
            if pos.x < 0 or pos.x + tracker_dims[0] > self.width_m:
                raise ValueError("Tracker position exceeds block width")
            if pos.y < 0 or pos.y + tracker_dims[1] > self.height_m:
                raise ValueError("Tracker position exceeds block height")
                
        return True
    
    def calculate_power(self) -> float:
        """Calculate total DC power capacity of the block"""
        modules_per_tracker = self.tracker_template.get_total_modules()
        total_modules = len(self.tracker_positions) * modules_per_tracker
        return total_modules * self.tracker_template.module_spec.wattage
    
    def get_tracker_coordinates(self) -> List[tuple[float, float, float]]:
        """
        Get list of all tracker coordinates and rotations
        Returns list of (x, y, rotation) tuples
        """
        return [(pos.x, pos.y, pos.rotation) for pos in self.tracker_positions]
    
    def calculate_cable_lengths(self) -> Dict[str, float]:
        """
        Calculate required cable lengths for the block
        Returns dictionary of cable types and their total lengths
        """
        if not self.wiring_config:
            return {}
            
        lengths = {}
        
        if self.wiring_config.wiring_type == WiringType.HOMERUN:
            # Calculate individual string cable lengths
            string_cable_length = 0
            for route in self.wiring_config.cable_routes.values():
                points = route
                for i in range(len(points) - 1):
                    dx = points[i+1][0] - points[i][0]
                    dy = points[i+1][1] - points[i][1]
                    string_cable_length += (dx**2 + dy**2)**0.5
            
            lengths["string_cable"] = string_cable_length
            
        else:  # Wire Harness configuration
            # Calculate harness lengths
            harness_length = 0
            for route in self.wiring_config.cable_routes.values():
                points = route
                for i in range(len(points) - 1):
                    dx = points[i+1][0] - points[i][0]
                    dy = points[i+1][1] - points[i][1]
                    harness_length += (dx**2 + dy**2)**0.5
            
            lengths["harness_cable"] = harness_length
            
        return lengths
    
    def __str__(self) -> str:
        return (f"Block {self.block_id} - {len(self.tracker_positions)} trackers, "
                f"{self.calculate_power()/1000:.1f}kW")