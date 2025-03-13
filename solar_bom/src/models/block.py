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
    whip_cable_size: str = "8 AWG"  # Default value for whips

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
            for route_id, route in self.wiring_config.cable_routes.items():
                points = route
                for i in range(len(points) - 1):
                    dx = points[i+1][0] - points[i][0]
                    dy = points[i+1][1] - points[i][1]
                    string_cable_length += (dx**2 + dy**2)**0.5
            
            lengths["string_cable"] = string_cable_length
            
        else:  # Wire Harness configuration
            # Calculate string cable lengths
            string_cable_length = 0
            harness_cable_length = 0
            
            for route_id, route in self.wiring_config.cable_routes.items():
                points = route
                route_length = 0
                for i in range(len(points) - 1):
                    dx = points[i+1][0] - points[i][0]
                    dy = points[i+1][1] - points[i][1]
                    route_length += (dx**2 + dy**2)**0.5
                
                # Determine if this is a string or harness route
                if "node" in route_id or "src" in route_id:
                    string_cable_length += route_length
                elif "harness" in route_id or "main" in route_id:
                    harness_cable_length += route_length
            
            lengths["string_cable"] = string_cable_length
            lengths["harness_cable"] = harness_cable_length
            
        return lengths
    
    def __str__(self) -> str:
        return (f"Block {self.block_id} - {len(self.tracker_positions)} trackers, "
                f"{self.calculate_power()/1000:.1f}kW")
    
    def to_dict(self) -> dict:
        """Convert block configuration to dictionary for serialization"""
        tracker_positions_data = []
        for pos in self.tracker_positions:
            pos_data = {
                'x': pos.x,
                'y': pos.y,
                'rotation': pos.rotation,
                'template_name': pos.template.template_name if pos.template else None,
                'strings': []
            }
            
            # Save strings data
            for string in pos.strings:
                string_data = {
                    'index': string.index,
                    'positive_source_x': string.positive_source_x,
                    'positive_source_y': string.positive_source_y,
                    'negative_source_x': string.negative_source_x,
                    'negative_source_y': string.negative_source_y,
                    'num_modules': string.num_modules
                }
                pos_data['strings'].append(string_data)
                
            tracker_positions_data.append(pos_data)
        
        # Basic block data
        data = {
            'block_id': self.block_id,
            'width_m': self.width_m,
            'height_m': self.height_m,
            'row_spacing_m': self.row_spacing_m,
            'ns_spacing_m': self.ns_spacing_m,
            'gcr': self.gcr,
            'description': self.description,
            'tracker_positions': tracker_positions_data,
            'device_x': self.device_x,
            'device_y': self.device_y,
            'device_spacing_m': self.device_spacing_m
        }
        
        # Add inverter reference if exists
        if self.inverter:
            data['inverter_id'] = f"{self.inverter.manufacturer} {self.inverter.model}"
        
        # Add tracker template reference if exists
        # Use the template from the first tracker position if available, otherwise use the block's template
        if self.tracker_positions and self.tracker_positions[0].template:
            data['tracker_template_name'] = self.tracker_positions[0].template.template_name
        elif self.tracker_template:
            data['tracker_template_name'] = self.tracker_template.template_name
        
        # Add wiring config if exists
        if self.wiring_config:
            wiring_data = {
                'wiring_type': self.wiring_config.wiring_type.value,
                'string_cable_size': self.wiring_config.string_cable_size,
                'harness_cable_size': self.wiring_config.harness_cable_size,
                'positive_collection_points': [],
                'negative_collection_points': [],
                'strings_per_collection': self.wiring_config.strings_per_collection,
                'cable_routes': self.wiring_config.cable_routes
            }
            
            # Serialize collection points
            for point in self.wiring_config.positive_collection_points:
                wiring_data['positive_collection_points'].append({
                    'x': point.x,
                    'y': point.y,
                    'connected_strings': point.connected_strings,
                    'current_rating': point.current_rating
                })
                
            for point in self.wiring_config.negative_collection_points:
                wiring_data['negative_collection_points'].append({
                    'x': point.x,
                    'y': point.y,
                    'connected_strings': point.connected_strings,
                    'current_rating': point.current_rating
                })
                
            data['wiring_config'] = wiring_data
        
        return data

    @classmethod
    def from_dict(cls, data: dict, tracker_templates: dict, inverters: dict):
        """Create a BlockConfig instance from dictionary data"""
        # Get referenced objects
        tracker_template = None
        if 'tracker_template_name' in data and data['tracker_template_name'] in tracker_templates:
            tracker_template = tracker_templates[data['tracker_template_name']]
        
        inverter = None
        if 'inverter_id' in data and data['inverter_id'] in inverters:
            inverter = inverters[data['inverter_id']]
        
        # Create block instance
        block = cls(
            block_id=data['block_id'],
            inverter=inverter,
            tracker_template=tracker_template,
            width_m=data['width_m'],
            height_m=data['height_m'],
            row_spacing_m=data['row_spacing_m'],
            ns_spacing_m=data['ns_spacing_m'],
            gcr=data['gcr'],
            description=data.get('description'),
            device_x=data.get('device_x', 0.0),
            device_y=data.get('device_y', 0.0),
            device_spacing_m=data.get('device_spacing_m', 1.83)
        )
        
        # Load tracker positions
        from .tracker import TrackerPosition, StringPosition
        
        for pos_data in data.get('tracker_positions', []):
            # Use the template_name from the position data to get the correct template
            position_template = None
            if 'template_name' in pos_data and pos_data['template_name'] in tracker_templates:
                position_template = tracker_templates[pos_data['template_name']]
            else:
                # Fall back to block's template if position doesn't have a valid template
                position_template = tracker_template
                
            if not position_template:
                continue
                
            pos = TrackerPosition(
                x=pos_data['x'],
                y=pos_data['y'],
                rotation=pos_data['rotation'],
                template=position_template
            )
            
            # Load strings data if available
            for string_data in pos_data.get('strings', []):
                string = StringPosition(
                    index=string_data['index'],
                    positive_source_x=string_data['positive_source_x'],
                    positive_source_y=string_data['positive_source_y'],
                    negative_source_x=string_data['negative_source_x'],
                    negative_source_y=string_data['negative_source_y'],
                    num_modules=string_data['num_modules']
                )
                pos.strings.append(string)
                
            block.tracker_positions.append(pos)
        
        # Load wiring configuration if exists
        if 'wiring_config' in data:
            wiring_data = data['wiring_config']
            
            # Create collection points
            positive_points = []
            for point_data in wiring_data.get('positive_collection_points', []):
                point = CollectionPoint(
                    x=point_data['x'],
                    y=point_data['y'],
                    connected_strings=point_data['connected_strings'],
                    current_rating=point_data['current_rating']
                )
                positive_points.append(point)
                
            negative_points = []
            for point_data in wiring_data.get('negative_collection_points', []):
                point = CollectionPoint(
                    x=point_data['x'],
                    y=point_data['y'],
                    connected_strings=point_data['connected_strings'],
                    current_rating=point_data['current_rating']
                )
                negative_points.append(point)
            
            # Create wiring config
            block.wiring_config = WiringConfig(
                wiring_type=WiringType(wiring_data['wiring_type']),
                positive_collection_points=positive_points,
                negative_collection_points=negative_points,
                strings_per_collection=wiring_data.get('strings_per_collection', {}),
                cable_routes=wiring_data.get('cable_routes', {}),
                string_cable_size=wiring_data.get('string_cable_size', "10 AWG"),
                harness_cable_size=wiring_data.get('harness_cable_size', "8 AWG")
            )
        
        return block
