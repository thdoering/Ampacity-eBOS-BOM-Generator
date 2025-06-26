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
class HarnessGroup:
    """Represents a group of strings combined into a harness"""
    string_indices: List[int]  # Indices of strings in this harness
    cable_size: str = "8 AWG"  # Cable size for this harness
    fuse_rating_amps: int = 15  # Fuse rating in amps
    use_fuse: bool = True  # Whether to use fuses for this harness (default to True for 2+ strings)
    
@dataclass
class WiringConfig:
    """Data class representing block wiring configuration"""
    wiring_type: WiringType
    positive_collection_points: List[CollectionPoint]
    negative_collection_points: List[CollectionPoint]
    strings_per_collection: Dict[int, int]  # Collection point ID -> number of strings
    cable_routes: Dict[str, List[tuple[float, float]]]  # Route ID -> list of coordinates
    realistic_cable_routes: Dict[str, List[tuple[float, float]]] = field(default_factory=dict)  # Realistic route ID -> list of coordinates (for BOM)
    string_cable_size: str = "10 AWG"  # Default value
    harness_cable_size: str = "8 AWG"  # Default value
    whip_cable_size: str = "8 AWG"  # Default value for whips
    extender_cable_size: str = "8 AWG"  # Default value for extenders
    custom_whip_points: Dict[str, Dict[str, tuple[float, float]]] = field(default_factory=dict)   # Format: {'tracker_id': {'positive': (x, y), 'negative': (x, y)}}
    harness_groupings: Dict[int, List[HarnessGroup]] = field(default_factory=dict)
    custom_harness_whip_points: Dict[str, Dict[int, Dict[str, tuple[float, float]]]] = field(default_factory=dict)  # Format: {'tracker_id': {harness_idx: {'positive': (x, y), 'negative': (x, y)}}}
    use_custom_positions_for_bom: bool = False  # New field - default to FALSE
    routing_mode: str = "realistic"

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
    block_realistic_routes: Dict[str, List[tuple[float, float]]] = field(default_factory=dict)
    
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
        Calculate required cable lengths for the block, separated by polarity
        Returns dictionary of cable types and their total lengths
        """
        if not self.wiring_config:
            return {}
                
        lengths = {}

        # Use routes from wiring configuration only
        cable_routes = getattr(self.wiring_config, 'cable_routes', {}) or {}
        
        # If still no routes, return empty dictionary
        if not cable_routes:
            return lengths

        if self.wiring_config.wiring_type == WiringType.HOMERUN:
            # For homerun, we track string cable length - split by polarity
            pos_string_cable_length = 0
            neg_string_cable_length = 0
            pos_whip_cable_length = 0
            neg_whip_cable_length = 0
            
            for route_id, route in cable_routes.items():
                points = route
                route_length = 0
                for i in range(len(points) - 1):
                    dx = points[i+1][0] - points[i][0]
                    dy = points[i+1][1] - points[i][1]
                    route_length += (dx**2 + dy**2)**0.5
                
                # Determine if this is a string or whip route and polarity
                if "pos_string" in route_id:
                    pos_string_cable_length += route_length
                elif "neg_string" in route_id:
                    neg_string_cable_length += route_length
                elif "pos_whip" in route_id:
                    pos_whip_cable_length += route_length
                elif "neg_whip" in route_id:
                    neg_whip_cable_length += route_length
            
            lengths["string_cable_positive"] = pos_string_cable_length
            lengths["string_cable_negative"] = neg_string_cable_length
            lengths["whip_cable_positive"] = pos_whip_cable_length
            lengths["whip_cable_negative"] = neg_whip_cable_length
            
        else:  # Wire Harness configuration
            # Calculate cable lengths by type and polarity
            pos_string_cable_length = 0
            neg_string_cable_length = 0
            pos_harness_cable_length = 0
            neg_harness_cable_length = 0
            pos_whip_cable_length = 0
            neg_whip_cable_length = 0
            
            for route_id, route in cable_routes.items():
                points = route
                route_length = 0
                for i in range(len(points) - 1):
                    dx = points[i+1][0] - points[i][0]
                    dy = points[i+1][1] - points[i][1]
                    route_length += (dx**2 + dy**2)**0.5
                
                # Determine route type and polarity - with more flexible matching
                if "pos_src" in route_id or "pos_node" in route_id:
                    pos_string_cable_length += route_length
                elif "neg_src" in route_id or "neg_node" in route_id:
                    neg_string_cable_length += route_length
                elif "pos_harness" in route_id:
                    pos_harness_cable_length += route_length
                elif "neg_harness" in route_id:
                    neg_harness_cable_length += route_length
                # More flexible whip route pattern matching
                elif "pos_main" in route_id or "pos_dev" in route_id or "whip_pos" in route_id or "pos_whip" in route_id:
                    pos_whip_cable_length += route_length
                elif "neg_main" in route_id or "neg_dev" in route_id or "whip_neg" in route_id or "neg_whip" in route_id:
                    neg_whip_cable_length += route_length
            
            lengths["string_cable_positive"] = pos_string_cable_length
            lengths["string_cable_negative"] = neg_string_cable_length
            lengths["harness_cable_positive"] = pos_harness_cable_length
            lengths["harness_cable_negative"] = neg_harness_cable_length
            lengths["whip_cable_positive"] = pos_whip_cable_length
            lengths["whip_cable_negative"] = neg_whip_cable_length

            # Calculate extender cable lengths
            pos_extender_cable_length = 0
            neg_extender_cable_length = 0

            for route_id, route in cable_routes.items():
                points = route
                route_length = 0
                for i in range(len(points) - 1):
                    dx = points[i+1][0] - points[i][0]
                    dy = points[i+1][1] - points[i][1]
                    route_length += (dx**2 + dy**2)**0.5
                
                # Determine if this is an extender route
                if "pos_extender" in route_id:
                    pos_extender_cable_length += route_length
                elif "neg_extender" in route_id:
                    neg_extender_cable_length += route_length

            lengths["extender_cable_positive"] = pos_extender_cable_length
            lengths["extender_cable_negative"] = neg_extender_cable_length
            
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
                'whip_cable_size': getattr(self.wiring_config, 'whip_cable_size', "8 AWG"),
                'extender_cable_size': getattr(self.wiring_config, 'extender_cable_size', "8 AWG"),
                'routing_mode': getattr(self.wiring_config, 'routing_mode', 'realistic'),
                'positive_collection_points': [],
                'negative_collection_points': [],
                'strings_per_collection': self.wiring_config.strings_per_collection,
                'cable_routes': self.wiring_config.cable_routes,
                'realistic_cable_routes': getattr(self.wiring_config, 'realistic_cable_routes', {}),
                'custom_whip_points': getattr(self.wiring_config, 'custom_whip_points', {})
            }

            # Line to include harness_groupings:
            if hasattr(self.wiring_config, 'harness_groupings'):
                # Need special handling since each harness contains objects
                harness_groups_data = {}
                for string_count, harness_list in self.wiring_config.harness_groupings.items():
                    harness_groups_data[string_count] = [
                        {
                            'string_indices': harness.string_indices,
                            'cable_size': getattr(harness, 'cable_size', "8 AWG"),
                            'fuse_rating_amps': getattr(harness, 'fuse_rating_amps', 15),
                            'use_fuse': getattr(harness, 'use_fuse', True)
                        }
                        for harness in harness_list
                    ]
                wiring_data['harness_groupings'] = harness_groups_data
            
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
            
            harness_groupings = {}
            if 'harness_groupings' in wiring_data:
                from ..models.block import HarnessGroup  # Import at the top of the function if not already there
                
                for string_count_str, harness_list_data in wiring_data['harness_groupings'].items():
                    string_count = int(string_count_str)
                    harness_list = []
                    
                    for harness_data in harness_list_data:
                        harness = HarnessGroup(
                            string_indices=harness_data['string_indices'],
                            cable_size=harness_data.get('cable_size', "8 AWG"),
                            fuse_rating_amps=harness_data.get('fuse_rating_amps', 15),
                            use_fuse=harness_data.get('use_fuse', True)
                        )
                        harness_list.append(harness)
                        
                    harness_groupings[string_count] = harness_list

            # Create wiring config
            block.wiring_config = WiringConfig(
                wiring_type=WiringType(wiring_data['wiring_type']),
                positive_collection_points=positive_points,
                negative_collection_points=negative_points,
                strings_per_collection=wiring_data.get('strings_per_collection', {}),
                cable_routes=wiring_data.get('cable_routes', {}),
                realistic_cable_routes=wiring_data.get('realistic_cable_routes', {}),
                string_cable_size=wiring_data.get('string_cable_size', "10 AWG"),
                harness_cable_size=wiring_data.get('harness_cable_size', "8 AWG"),
                whip_cable_size=wiring_data.get('whip_cable_size', "8 AWG"),
                extender_cable_size=wiring_data.get('extender_cable_size', "8 AWG"),
                custom_whip_points=wiring_data.get('custom_whip_points', {}),
                harness_groupings=harness_groupings,
                routing_mode=wiring_data.get('routing_mode', 'realistic')
            )
        
        # ADD THIS DEBUG CODE RIGHT BEFORE THE RETURN STATEMENT:
        print(f"=== BLOCK.PY from_dict DEBUG ===")
        print(f"Block {data['block_id']} - Original data had {len(data.get('tracker_positions', []))} positions")
        print(f"Block {data['block_id']} - Loaded {len(block.tracker_positions)} tracker positions")
        if 'tracker_template_name' in data:
            print(f"Block {data['block_id']} - Looking for template: '{data['tracker_template_name']}'")
            print(f"Block {data['block_id']} - Template found: {data['tracker_template_name'] in tracker_templates}")
        print("================================")

        return block
