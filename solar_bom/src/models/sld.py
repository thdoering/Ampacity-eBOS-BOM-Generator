"""
Single Line Diagram (SLD) data models for Solar eBOS BOM Generator
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple
from enum import Enum


class SLDElementType(Enum):
    """Types of elements that can appear in an SLD"""
    PV_BLOCK = "pv_block"
    INVERTER = "inverter"
    COMBINER_BOX = "combiner_box"
    TRANSFORMER = "transformer"
    SWITCHGEAR = "switchgear"
    METER = "meter"
    UTILITY_CONNECTION = "utility_connection"


class ConnectionPortType(Enum):
    """Types of connection ports on SLD elements"""
    DC_POSITIVE = "dc_positive"
    DC_NEGATIVE = "dc_negative"
    AC_L1 = "ac_l1"
    AC_L2 = "ac_l2"
    AC_L3 = "ac_l3"
    AC_NEUTRAL = "ac_neutral"
    GROUND = "ground"


@dataclass
class ConnectionPort:
    """Represents a connection point on an SLD element"""
    port_id: str  # Unique identifier for this port
    port_type: ConnectionPortType
    side: str  # 'top', 'bottom', 'left', 'right'
    offset: float  # Position along the side (0.0 to 1.0)
    max_current: Optional[float] = None  # Maximum current rating
    voltage_nominal: Optional[float] = None  # Nominal voltage
    
    def to_dict(self) -> dict:
        """Convert to dictionary for serialization"""
        return {
            'port_id': self.port_id,
            'port_type': self.port_type.value,
            'side': self.side,
            'offset': self.offset,
            'max_current': self.max_current,
            'voltage_nominal': self.voltage_nominal
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'ConnectionPort':
        """Create from dictionary"""
        return cls(
            port_id=data['port_id'],
            port_type=ConnectionPortType(data['port_type']),
            side=data['side'],
            offset=data['offset'],
            max_current=data.get('max_current'),
            voltage_nominal=data.get('voltage_nominal')
        )


@dataclass
class SLDElement:
    """Represents a single element in the SLD"""
    element_id: str  # Unique identifier
    element_type: SLDElementType
    x: float  # X position in canvas coordinates
    y: float  # Y position in canvas coordinates
    width: float = 80.0  # Width in canvas units
    height: float = 60.0  # Height in canvas units
    rotation: float = 0.0  # Rotation in degrees
    label: str = ""  # Display label
    
    # Reference to source data
    source_block_id: Optional[str] = None  # For PV blocks
    source_inverter_id: Optional[str] = None  # For inverters
    source_device_id: Optional[str] = None  # For other devices
    
    # Electrical properties
    power_kw: Optional[float] = None
    voltage_dc: Optional[float] = None
    voltage_ac: Optional[float] = None
    current_dc: Optional[float] = None
    current_ac: Optional[float] = None
    
    # Connection ports
    ports: List[ConnectionPort] = field(default_factory=list)
    
    # Visual properties
    color: str = "#4A90E2"  # Default blue
    stroke_color: str = "#2C5282"
    stroke_width: float = 2.0
    fill_opacity: float = 1.0
    
    # Additional properties stored as key-value pairs
    properties: Dict[str, any] = field(default_factory=dict)
    
    # Canvas item IDs for interaction
    canvas_items: List[int] = field(default_factory=list)
    
    def get_port(self, port_id: str) -> Optional[ConnectionPort]:
        """Get a specific port by ID"""
        for port in self.ports:
            if port.port_id == port_id:
                return port
        return None
    
    def get_port_position(self, port_id: str) -> Optional[Tuple[float, float]]:
        """Get the absolute position of a port in canvas coordinates"""
        port = self.get_port(port_id)
        if not port:
            return None
            
        # Calculate position based on side and offset
        if port.side == 'top':
            port_x = self.x + (self.width * port.offset)
            port_y = self.y
        elif port.side == 'bottom':
            port_x = self.x + (self.width * port.offset)
            port_y = self.y + self.height
        elif port.side == 'left':
            port_x = self.x
            port_y = self.y + (self.height * port.offset)
        elif port.side == 'right':
            port_x = self.x + self.width
            port_y = self.y + (self.height * port.offset)
        else:
            return None
            
        return (port_x, port_y)
    
    def to_dict(self) -> dict:
        """Convert to dictionary for serialization"""
        return {
            'element_id': self.element_id,
            'element_type': self.element_type.value,
            'x': self.x,
            'y': self.y,
            'width': self.width,
            'height': self.height,
            'rotation': self.rotation,
            'label': self.label,
            'source_block_id': self.source_block_id,
            'source_inverter_id': self.source_inverter_id,
            'source_device_id': self.source_device_id,
            'power_kw': self.power_kw,
            'voltage_dc': self.voltage_dc,
            'voltage_ac': self.voltage_ac,
            'current_dc': self.current_dc,
            'current_ac': self.current_ac,
            'ports': [port.to_dict() for port in self.ports],
            'color': self.color,
            'stroke_color': self.stroke_color,
            'stroke_width': self.stroke_width,
            'fill_opacity': self.fill_opacity,
            'properties': self.properties
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'SLDElement':
        """Create from dictionary"""
        element = cls(
            element_id=data['element_id'],
            element_type=SLDElementType(data['element_type']),
            x=data['x'],
            y=data['y'],
            width=data.get('width', 80.0),
            height=data.get('height', 60.0),
            rotation=data.get('rotation', 0.0),
            label=data.get('label', ''),
            source_block_id=data.get('source_block_id'),
            source_inverter_id=data.get('source_inverter_id'),
            source_device_id=data.get('source_device_id'),
            power_kw=data.get('power_kw'),
            voltage_dc=data.get('voltage_dc'),
            voltage_ac=data.get('voltage_ac'),
            current_dc=data.get('current_dc'),
            current_ac=data.get('current_ac'),
            color=data.get('color', '#4A90E2'),
            stroke_color=data.get('stroke_color', '#2C5282'),
            stroke_width=data.get('stroke_width', 2.0),
            fill_opacity=data.get('fill_opacity', 1.0),
            properties=data.get('properties', {})
        )
        
        # Load ports
        element.ports = [ConnectionPort.from_dict(port_data) 
                        for port_data in data.get('ports', [])]
        
        return element


@dataclass
class SLDConnection:
    """Represents a connection between two elements"""
    connection_id: str  # Unique identifier
    from_element: str  # Source element ID
    from_port: str  # Source port ID
    to_element: str  # Destination element ID
    to_port: str  # Destination port ID
    
    # Cable specifications
    cable_type: str = "DC"  # 'DC' or 'AC'
    cable_size: str = "10 AWG"  # Wire gauge
    cable_length_m: Optional[float] = None  # Length in meters
    num_conductors: int = 2  # Number of conductors
    
    # Electrical properties
    voltage: Optional[float] = None
    current: Optional[float] = None
    voltage_drop_percent: Optional[float] = None
    
    # Visual properties
    color: str = "#DC143C"  # Default red for positive
    stroke_width: float = 3.0
    stroke_style: str = "solid"  # 'solid', 'dashed', 'dotted'
    
    # Routing points for orthogonal path
    path_points: List[Tuple[float, float]] = field(default_factory=list)
    
    # Canvas item IDs
    canvas_items: List[int] = field(default_factory=list)
    
    def calculate_orthogonal_path(self, from_pos: Tuple[float, float], 
                                  to_pos: Tuple[float, float]) -> List[Tuple[float, float]]:
        """Calculate orthogonal (right-angle) path between two points"""
        path = [from_pos]
        
        # Simple orthogonal routing: go horizontal first, then vertical
        # Can be enhanced with obstacle avoidance later
        if from_pos[0] != to_pos[0] and from_pos[1] != to_pos[1]:
            # Need to make a turn
            mid_x = to_pos[0]
            mid_y = from_pos[1]
            path.append((mid_x, mid_y))
        
        path.append(to_pos)
        return path
    
    def to_dict(self) -> dict:
        """Convert to dictionary for serialization"""
        return {
            'connection_id': self.connection_id,
            'from_element': self.from_element,
            'from_port': self.from_port,
            'to_element': self.to_element,
            'to_port': self.to_port,
            'cable_type': self.cable_type,
            'cable_size': self.cable_size,
            'cable_length_m': self.cable_length_m,
            'num_conductors': self.num_conductors,
            'voltage': self.voltage,
            'current': self.current,
            'voltage_drop_percent': self.voltage_drop_percent,
            'color': self.color,
            'stroke_width': self.stroke_width,
            'stroke_style': self.stroke_style,
            'path_points': self.path_points
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'SLDConnection':
        """Create from dictionary"""
        return cls(
            connection_id=data['connection_id'],
            from_element=data['from_element'],
            from_port=data['from_port'],
            to_element=data['to_element'],
            to_port=data['to_port'],
            cable_type=data.get('cable_type', 'DC'),
            cable_size=data.get('cable_size', '10 AWG'),
            cable_length_m=data.get('cable_length_m'),
            num_conductors=data.get('num_conductors', 2),
            voltage=data.get('voltage'),
            current=data.get('current'),
            voltage_drop_percent=data.get('voltage_drop_percent'),
            color=data.get('color', '#DC143C'),
            stroke_width=data.get('stroke_width', 3.0),
            stroke_style=data.get('stroke_style', 'solid'),
            path_points=data.get('path_points', [])
        )


@dataclass
class SLDAnnotation:
    """Represents a text annotation on the diagram"""
    annotation_id: str
    text: str
    x: float
    y: float
    font_size: float = 12.0
    font_family: str = "Arial"
    color: str = "#000000"
    anchor: str = "center"  # 'center', 'left', 'right'
    rotation: float = 0.0
    visible: bool = True
    
    # Associated element (optional)
    element_id: Optional[str] = None
    
    # Canvas item ID
    canvas_item_id: Optional[int] = None
    
    def to_dict(self) -> dict:
        """Convert to dictionary for serialization"""
        return {
            'annotation_id': self.annotation_id,
            'text': self.text,
            'x': self.x,
            'y': self.y,
            'font_size': self.font_size,
            'font_family': self.font_family,
            'color': self.color,
            'anchor': self.anchor,
            'rotation': self.rotation,
            'visible': self.visible,
            'element_id': self.element_id
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'SLDAnnotation':
        """Create from dictionary"""
        return cls(
            annotation_id=data['annotation_id'],
            text=data['text'],
            x=data['x'],
            y=data['y'],
            font_size=data.get('font_size', 12.0),
            font_family=data.get('font_family', 'Arial'),
            color=data.get('color', '#000000'),
            anchor=data.get('anchor', 'center'),
            rotation=data.get('rotation', 0.0),
            visible=data.get('visible', True),
            element_id=data.get('element_id')
        )


@dataclass
class SLDDiagram:
    """Complete SLD for a project"""
    project_id: str
    diagram_name: str = "Block Level SLD"
    
    # Diagram components
    elements: List[SLDElement] = field(default_factory=list)
    connections: List[SLDConnection] = field(default_factory=list)
    annotations: List[SLDAnnotation] = field(default_factory=list)
    
    # Canvas properties
    canvas_width: int = 1200
    canvas_height: int = 800
    grid_size: int = 10
    grid_visible: bool = True
    
    # View settings
    zoom_level: float = 1.0
    pan_x: float = 0.0
    pan_y: float = 0.0
    
    # Layout zones for auto-layout
    pv_zone_x_start: float = 50
    pv_zone_x_end: float = 400
    device_zone_x_start: float = 400
    device_zone_x_end: float = 800
    inverter_zone_x_start: float = 800
    inverter_zone_x_end: float = 1150
    
    # Metadata
    created_date: Optional[str] = None
    modified_date: Optional[str] = None
    version: str = "1.0"
    
    def add_element(self, element: SLDElement) -> None:
        """Add an element to the diagram"""
        self.elements.append(element)
    
    def remove_element(self, element_id: str) -> bool:
        """Remove an element and its connections"""
        # Remove element
        element_removed = False
        self.elements = [e for e in self.elements 
                        if e.element_id != element_id or not (element_removed := True)]
        
        if element_removed:
            # Remove connections involving this element
            self.connections = [c for c in self.connections 
                               if c.from_element != element_id and c.to_element != element_id]
            # Remove annotations for this element
            self.annotations = [a for a in self.annotations 
                               if a.element_id != element_id]
        
        return element_removed
    
    def add_connection(self, connection: SLDConnection) -> None:
        """Add a connection to the diagram"""
        self.connections.append(connection)
    
    def remove_connection(self, connection_id: str) -> bool:
        """Remove a connection"""
        initial_count = len(self.connections)
        self.connections = [c for c in self.connections 
                           if c.connection_id != connection_id]
        return len(self.connections) < initial_count
    
    def get_element(self, element_id: str) -> Optional[SLDElement]:
        """Get an element by ID"""
        for element in self.elements:
            if element.element_id == element_id:
                return element
        return None
    
    def get_connections_for_element(self, element_id: str) -> List[SLDConnection]:
        """Get all connections involving an element"""
        return [c for c in self.connections 
                if c.from_element == element_id or c.to_element == element_id]
    
    def validate_connection(self, from_element_id: str, from_port_id: str,
                           to_element_id: str, to_port_id: str) -> Tuple[bool, str]:
        """Validate if a connection is allowed"""
        from_element = self.get_element(from_element_id)
        to_element = self.get_element(to_element_id)
        
        if not from_element or not to_element:
            return False, "Invalid element IDs"
        
        from_port = from_element.get_port(from_port_id)
        to_port = to_element.get_port(to_port_id)
        
        if not from_port or not to_port:
            return False, "Invalid port IDs"
        
        # Check port type compatibility
        if from_port.port_type == ConnectionPortType.DC_POSITIVE:
            if to_port.port_type not in [ConnectionPortType.DC_POSITIVE]:
                return False, "DC positive must connect to DC positive"
        elif from_port.port_type == ConnectionPortType.DC_NEGATIVE:
            if to_port.port_type not in [ConnectionPortType.DC_NEGATIVE]:
                return False, "DC negative must connect to DC negative"
        
        # Check for duplicate connections
        for conn in self.connections:
            if (conn.from_element == from_element_id and conn.from_port == from_port_id and
                conn.to_element == to_element_id and conn.to_port == to_port_id):
                return False, "Connection already exists"
        
        return True, "Connection valid"
    
    def auto_layout(self) -> None:
        """Automatically position elements in their respective zones"""
        # Separate elements by type
        pv_blocks = []
        inverters = []
        other_devices = []
        
        for element in self.elements:
            if element.element_type == SLDElementType.PV_BLOCK:
                pv_blocks.append(element)
            elif element.element_type == SLDElementType.INVERTER:
                inverters.append(element)
            else:
                other_devices.append(element)
        
        # Layout PV blocks in left zone
        y_spacing = 100
        current_y = 50
        
        for element in pv_blocks:
            element.x = self.pv_zone_x_start + 50
            element.y = current_y
            current_y += element.height + y_spacing
        
        # Layout inverters in right zone
        current_y = 50
        for element in inverters:
            element.x = self.inverter_zone_x_start + 50
            element.y = current_y
            current_y += element.height + y_spacing
        
        # Layout other devices in center zone
        current_y = 50
        for element in other_devices:
            element.x = self.device_zone_x_start + 50
            element.y = current_y
            current_y += element.height + y_spacing
    
    def to_dict(self) -> dict:
        """Convert to dictionary for serialization"""
        return {
            'project_id': self.project_id,
            'diagram_name': self.diagram_name,
            'elements': [e.to_dict() for e in self.elements],
            'connections': [c.to_dict() for c in self.connections],
            'annotations': [a.to_dict() for a in self.annotations],
            'canvas_width': self.canvas_width,
            'canvas_height': self.canvas_height,
            'grid_size': self.grid_size,
            'grid_visible': self.grid_visible,
            'zoom_level': self.zoom_level,
            'pan_x': self.pan_x,
            'pan_y': self.pan_y,
            'pv_zone_x_start': self.pv_zone_x_start,
            'pv_zone_x_end': self.pv_zone_x_end,
            'device_zone_x_start': self.device_zone_x_start,
            'device_zone_x_end': self.device_zone_x_end,
            'inverter_zone_x_start': self.inverter_zone_x_start,
            'inverter_zone_x_end': self.inverter_zone_x_end,
            'created_date': self.created_date,
            'modified_date': self.modified_date,
            'version': self.version
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'SLDDiagram':
        """Create from dictionary"""
        diagram = cls(
            project_id=data['project_id'],
            diagram_name=data.get('diagram_name', 'Block Level SLD'),
            canvas_width=data.get('canvas_width', 1200),
            canvas_height=data.get('canvas_height', 800),
            grid_size=data.get('grid_size', 10),
            grid_visible=data.get('grid_visible', True),
            zoom_level=data.get('zoom_level', 1.0),
            pan_x=data.get('pan_x', 0.0),
            pan_y=data.get('pan_y', 0.0),
            created_date=data.get('created_date'),
            modified_date=data.get('modified_date'),
            version=data.get('version', '1.0')
        )
        
        # Load layout zones
        diagram.pv_zone_x_start = data.get('pv_zone_x_start', 50)
        diagram.pv_zone_x_end = data.get('pv_zone_x_end', 400)
        diagram.device_zone_x_start = data.get('device_zone_x_start', 400)
        diagram.device_zone_x_end = data.get('device_zone_x_end', 800)
        diagram.inverter_zone_x_start = data.get('inverter_zone_x_start', 800)
        diagram.inverter_zone_x_end = data.get('inverter_zone_x_end', 1150)
        
        # Load components
        diagram.elements = [SLDElement.from_dict(e) for e in data.get('elements', [])]
        diagram.connections = [SLDConnection.from_dict(c) for c in data.get('connections', [])]
        diagram.annotations = [SLDAnnotation.from_dict(a) for a in data.get('annotations', [])]
        
        return diagram