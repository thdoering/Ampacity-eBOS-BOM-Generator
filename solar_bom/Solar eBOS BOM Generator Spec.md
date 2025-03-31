# Solar eBOS BOM Generator Software Specification (Updated)

## 1. System Overview

A Python-based application for creating detailed Bills of Material (BOM) for solar project electrical Balance of System (eBOS) components, focusing on string-to-inverter connections. The application provides a comprehensive GUI for designing solar tracker layouts, configuring electrical components, and generating accurate material lists with project management capabilities.

## 2. Core Components

### 2.1 Module Specification Management

#### Features

- Import module specifications via .pan file
- Manual entry option for module specifications
- Store and validate module parameters:
  - Model name/identifier
  - Manufacturer information
  - Physical dimensions (length, width, depth)
  - Weight specifications
  - Electrical characteristics:
    - Short circuit current (Isc)
    - Maximum power current (Imp)
    - Open circuit voltage (Voc)
    - Maximum power voltage (Vmp)
    - Power rating (Wattage)
  - Temperature coefficients
  - Maximum system voltage
  - Module efficiency (optional)
  - Bifaciality factor (optional)
  - Default orientation preference
  - Cell count per module
- Module library management
  - Save/load module specifications
  - Delete saved modules
  - List available modules

#### Module Data Model
```python
class ModuleSpec:
    # Basic module info
    manufacturer: str
    model: str
    type: ModuleType
    
    # Physical specifications
    length_mm: float
    width_mm: float
    depth_mm: float
    weight_kg: float
    
    # Electrical specifications
    wattage: float
    vmp: float  # Maximum power voltage
    imp: float  # Maximum power current
    voc: float  # Open circuit voltage
    isc: float  # Short circuit current
    max_system_voltage: float
    
    # Optional specifications
    efficiency: Optional[float] = None
    temperature_coefficient: Optional[float] = None
    bifaciality_factor: Optional[float] = None
    
    # Default mounting configuration
    default_orientation: ModuleOrientation = ModuleOrientation.PORTRAIT
    cells_per_module: int = 72
```

### 2.2 Tracker Template Creator

#### Features

- Define tracker physical characteristics:
  - Number of modules per tracker (calculated)
  - Module spacing (default 0.01m)
  - Motor/drive gaps (default 1.0m)
  - String size (modules per string)
  - Strings per tracker
- Save templates for reuse within project
- Real-time preview updates as parameters change
- Support for both portrait and landscape module orientations
- Calculation and visualization of total tracker dimensions
- Load saved templates for editing
- Delete templates from library

#### User Interface

- Form-based input for specifications
- Visual representation of tracker layout with accurate dimensions
- Template management (save, load, edit, delete)
- Real-time calculation updates and preview
- Module spacing visualization
- Motor gap representation
- Display of total tracker dimensions and module count

#### Tracker Template Data Model
```python
class TrackerTemplate:
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
```

### 2.3 Block Configuration Tool

#### Features

- Inverter specification input:
  - Import via .ond file
  - Manual entry option
  - Support for multiple MPPT channels
- Define row spacing/pitch (GCR)
- Configure multiple blocks using saved tracker templates
- Validate block configuration against inverter specifications
- Flexible device placement:
  - Place inverters/combiner boxes anywhere within block
  - Visual safety zones around devices with configurable clearance
  - Device dimensions and clearances in feet/meters
- Navigation and visualization controls:
  - Advanced pan and zoom functionality for large layouts
  - Dynamic scaling based on block size
  - Mouse wheel zoom with smooth scaling
  - Middle/right mouse button panning
  - Dynamic redrawing with scale preservation
  - Undo/redo system for all block modifications
- Block management features:
  - Create new blocks
  - Delete blocks
  - Copy existing blocks with incremental naming
  - Rename blocks
- Device configuration:
  - Set device type (String Inverter or Combiner Box)
  - Configure number of inputs
  - Set maximum current per input
  - Display total maximum current capacity

#### User Interface

- Block-based navigation system
- Drag-and-drop interface for placing tracker templates
- 2D layout visualization per block:
  - Customizable grid based on row spacing and N/S tracker spacing
  - Grid snapping for precise tracker placement
  - Real-time visualization of trackers and devices
  - Semi-transparent safety zones around devices
- Scale and navigation controls:
  - Mouse wheel zoom
  - Middle/right-click pan
  - Visual indicators for scale and dimensions
- "Add Block" functionality for multi-block projects
- Grid snapping functionality:
  - Automatic snapping to row spacing intervals
  - Automatic snapping to N/S tracker spacing intervals
  - Visual gridlines showing valid placement positions
- Real-time updates of the GCR (Ground Coverage Ratio) as row spacing changes
- Display of device spacing in multiple units (feet and meters)
- Tracker selection with visual highlighting:
  - Multi-select capability with shift/ctrl modifiers
  - Visual feedback for selected elements with color changes
  - Drag to select multiple elements with selection box
  - Keyboard shortcuts for selection manipulation (Ctrl+A, Delete, Esc)
- Dual-unit display (feet/meters) for measurements

#### Block Configuration Data Model
```python
class BlockConfig:
    # Block identification and core components
    block_id: str
    inverter: InverterSpec
    tracker_template: TrackerTemplate
    width_m: float
    height_m: float
    row_spacing_m: float  # Distance between tracker rows
    ns_spacing_m: float  # Distance between trackers in north/south direction
    gcr: float  # Ground Coverage Ratio
    description: Optional[str] = None
    
    # Device positioning and configuration
    device_x: float = 0.0  # X coordinate of device in meters
    device_y: float = 0.0  # Y coordinate of device in meters
    device_spacing_m: float = 1.83  # 6ft in meters default
    input_points: List[DeviceInputPoint] = field(default_factory=list)
    
    # Layout and wiring 
    tracker_positions: List[TrackerPosition] = field(default_factory=list)
    wiring_config: Optional[WiringConfig] = None
```

### 2.4 Inverter Management

#### Features

- Detailed inverter specification management:
  - Manufacturer and model information
  - Power rating and efficiency data
  - MPPT configuration (Independent, Parallel, Symmetric)
  - Input/output specifications
  - Physical dimensions
- Dynamic MPPT channel configuration:
  - Add/remove MPPT channels
  - Configure maximum current per channel
  - Set voltage ranges per channel
  - Define number of string inputs per channel
  - Set maximum power per channel
- Startup voltage and maximum DC voltage parameters
- Enhanced validation logic for inverter specifications
- Save and load inverter configurations
- Inverter library management:
  - Add, edit, delete inverters
  - Select from library for project use

#### Inverter Data Model
```python
class InverterSpec:
    # Basic inverter info
    manufacturer: str
    model: str
    rated_power: float
    max_efficiency: float
    
    # Input specifications
    mppt_channels: List[MPPTChannel]
    mppt_configuration: MPPTConfig
    max_dc_voltage: float
    startup_voltage: float
    
    # Output specifications
    nominal_ac_voltage: float
    max_ac_current: float
    power_factor: float
    
    # Physical specifications
    dimensions_mm: tuple[float, float, float]  # length, width, depth
    weight_kg: float
    ip_rating: str
    
    # Optional specifications
    temperature_range: Optional[tuple[float, float]] = None  # min, max °C
    altitude_limit_m: Optional[float] = None
    communication_protocol: Optional[str] = None
```

### 2.5 String-to-Inverter Wiring Configuration

#### Wire Routing Rules

- All cable routes follow a vertical-then-horizontal pattern
- Maintain visual separation between parallel cable runs
- Route cables along tracker edges with small offsets
- Positive cables route along left side of tracker, negative along right
- Routes determined by source point location relative to destination

#### Wire Harness Collection Points

- Node points placed near string source points with small offset
- Node points follow same left/right positioning as source points
- Each tracker maintains independent harness system
- No combining of harnesses between trackers

#### Visualization Standards

- Red indicates positive polarity cables/points
- Blue indicates negative polarity cables/points
- Line thickness corresponds to wire gauge size
- Visual offsets between parallel runs for clarity
- Source points, node points, and destination points visually distinct
- Current labels for visualizing electrical loads on wire segments
- Different line thicknesses based on wire gauge

#### Features

Support two wiring approaches, both with automatic cable routing:

1. String Homeruns:
   - Individual positive/negative cables per string
   - Direct connection to downstream device
   - Cable paths automatically generated from each string to device
   - Cable sizing based on single string current
   - More inputs used on downstream device
   - Longer total cable length

2. Wire Harness Solution:
   - Collection point at each tracker
   - Separate positive and negative wire harnesses
   - Combined strings at collection points
   - Two main cable runs to downstream device (pos/neg)
   - Cable sizing based on combined current
   - Fewer inputs on downstream device
   - Less total cable length but higher gauge wire
   - Advanced string grouping capabilities for creating custom harnesses
   - Custom harness configuration with individual string selection
   - Multiple harness support per tracker with different cable sizes
   - Support for individual whip point positioning for routing optimization
   - Realistic cable routing calculation options for accurate BOM generation

#### Harness Configuration Options
- Flexible string grouping within trackers
- Multiple independent harnesses per tracker
- Separate cable sizing for string, harness and whip cables
- Customizable whip points with visual interactive positioning
- Quick pattern templates for common harness configurations:
  - Split evenly in two harnesses
  - Separate furthest string from others
  - Default single harness configuration

#### Cable Specifications
- Support for different cable sizes:
  - 4 AWG (21.15 mm²)
  - 6 AWG (13.30 mm²)
  - 8 AWG (8.37 mm²)
  - 10 AWG (5.26 mm²)
- Independent configuration of string and harness cable sizes
- Visual representation of different cable sizes with line thickness
- Current calculations based on module specifications and number of combined strings
- Current labeling toggle for detailed electrical analysis

#### Automatic Routing Algorithm

- Routes calculated based on:
  - String collection point locations
  - Device location
  - Other tracker locations (for obstacle avoidance)
  - Row spacing
  - N/S spacing
- No manual waypoint placement or route editing
- Routes optimized for:
  - Minimum cable length
  - Following tracker rows where possible
  - Maintaining clearances from other equipment
- Detailed cable routing algorithms for both homerun and harness configurations
- Node-to-node connections for wire harness configurations

#### Wiring Configuration Data Model
```python
class WiringConfig:
    wiring_type: WiringType
    positive_collection_points: List[CollectionPoint]
    negative_collection_points: List[CollectionPoint]
    strings_per_collection: Dict[int, int]  # Collection point ID -> number of strings
    cable_routes: Dict[str, List[tuple[float, float]]]  # Route ID -> list of coordinates
    realistic_cable_routes: Dict[str, List[tuple[float, float]]]  # Realistic routes for BOM
    string_cable_size: str = "10 AWG"  # Default string cable size
    harness_cable_size: str = "8 AWG"  # Default harness cable size
    whip_cable_size: str = "8 AWG"  # Default whip cable size
    custom_whip_points: Dict[str, Dict[str, tuple[float, float]]]  # Custom whip positions
    harness_groupings: Dict[int, List[HarnessGroup]]  # Custom harness configurations
    use_custom_positions_for_bom: bool = False  # Use custom positions for BOM calculations
```

#### Validation Rules

- Maximum MPPT current limits
- Available input connections
- NEC current limits:
  - Individual string cables
  - Combined current in wire harnesses
- Maximum inputs per downstream device
- Collection point current ratings

### 2.6 Project Management System

#### Features

- Project metadata management:
  - Project name, description, and location
  - Client information
  - Creation and modification dates
  - Notes and additional documentation
- Project dashboard:
  - Recent projects with card-based presentation
  - Full project list with sorting and filtering
  - Search functionality across project metadata
  - One-click project access
- Project operations:
  - Create new projects
  - Open existing projects
  - Save project updates
  - Delete projects with confirmation
  - Rename projects
- Multi-project workflow:
  - Switch between projects
  - Copy data between projects
  - Status bar with current project info

#### Project Data Model
```python
class Project:
    metadata: ProjectMetadata
    blocks: Dict[str, dict] = field(default_factory=dict)
    selected_modules: List[str] = field(default_factory=list)
    selected_inverters: List[str] = field(default_factory=list)
    default_row_spacing_m: float = 6.0  # Default row spacing in meters
    
class ProjectMetadata:
    name: str
    description: Optional[str] = None
    location: Optional[str] = None
    client: Optional[str] = None
    created_date: datetime = field(default_factory=datetime.now)
    modified_date: datetime = field(default_factory=datetime.now)
    notes: Optional[str] = None
```

### 2.7 BOM Generation

#### Features

- Component calculation:
  - Automatic quantity calculation based on block configurations
  - Cable length calculations for different wiring types
  - Harness count based on number of strings
  - Support for multiple component categories
  - Segment-based cable length calculations for higher accuracy
  - Wire size-specific component categorization
  - Support for mixed cable sizes in complex harness configurations
  - Optional use of custom routing positions for BOM calculations
  - Automatic segment length rounding to standard increments
- BOM preview:
  - Real-time BOM generation
  - Component categorization and grouping
  - Block-specific and project-wide views
- Excel export:
  - Formatted Excel output with project information
  - Multiple sheets for summary and detailed views
  - Automatic column sizing and formatting
  - Project statistics summary
- Component categorization:
  - eBOS components
  - Structural elements
  - Interconnection equipment

#### BOM Export Format

- Project information sheet:
  - Project name, client, and location
  - System size and module specifications
  - Inverter and DC collection types
  - Project notes and description
- BOM summary sheet:
  - Component types and descriptions
  - Total quantities with appropriate units
  - Category grouping with visual separation
- Block details sheet:
  - Per-block component breakdown
  - Component types and quantities by block
  - Detailed component specifications
- Segment-based wire listings with specific lengths
- Detailed harness breakdown by string count and cable size
- Polarity-specific component listings (positive/negative separate)
- Warning indicators for sections with electrical concerns
- Enhanced project statistics and electrical configuration summary

#### Cable Segment Analysis
- Breakdown of cable runs into practical installation segments
- Length-specific segment counts for installation planning
- Standardized length increments with appropriate waste factors
- Separation of string, harness, and whip cable segments
- Support for mixed cable gauge systems within the same installation

### 2.8 Whip Point Management

#### Features
- Interactive positioning of connection points
  - Drag-and-drop interface for whip point placement
  - Visual feedback for selected points
  - Single and multi-point selection
  - Reset to default positions
- Harness-specific whip points
  - Vertical offset for multiple harnesses
  - Independent positioning for each harness group
  - Custom routing from collection nodes to whip points
- Position memory
  - Persistence of custom point positions between sessions
  - Optional use of custom positions for BOM calculations
  - Quick reset options for individual or all whip points
- Visual differentiation
  - Color-coding by polarity and selection state
  - Size changes for selected points
  - Distinct visualization for harness-specific points

## 3. Technical Requirements

### 3.1 Data Validation

- Inverter compatibility checks
- NEC compliance validation for current limits
- MPPT input validation
- String cable sizing validation
- Harness cable sizing validation
- No validation of string sizing required
- Real-time current calculation and visualization
- Wire gauge selection validation against current loads
- Visual warning system for approaching/exceeding ampacity limits
- Highlighting of problematic wire segments with overload indicators
- MPPT channel capacity validation against total string current
- Interactive warning panel with problem descriptions and locations

### 3.2 Calculations

- Cable length calculations based on layout
- Voltage drop calculations
- Power loss calculations
- Current calculations for wire harness solutions
- Conductor ampacity calculations
- String electrical characteristics with temperature effects
- Wire harness compatibility checking
- Ground Coverage Ratio (GCR) calculations
- Load percentage calculations relative to NEC requirements
- Ampacity verification for different wire gauges (4, 6, 8, 10 AWG)
- Harness compatibility checking with current load visualization
- Real-time calculation of accumulated current in wire harnesses
- Dynamic current labeling with visual indicators

### 3.3 File Handling

- Import .ond files for inverter specifications
- Import .pan files for module specifications
- Export BOM to Excel format
- Save/load complete project configurations
- Directory structure verification and creation
- Filename validation and cleanup
- Error handling for file operations
- Support for multiple file types and validation

### 3.4 State Management

- Undo/redo system for block configurations
- State preservation for project loading/saving
- Deep state copying to prevent reference issues
- Error recovery for file operations
- Consistent state validation

## 4. User Interface Requirements

- Python-based GUI framework (Tkinter)
- Tab-based interface for different functions
- Block-based navigation system
- Drag-and-drop functionality for tracker placement with grid snapping
- Form-based input for specifications
- Real-time validation feedback
- 2D layout visualization
- Error messaging for invalid configurations
- Undo/redo functionality with descriptive actions
- Advanced pan and zoom controls in layout visualizations
  - Mouse wheel zoom with smooth scaling
  - Middle/right mouse button panning
  - Dynamic redrawing with scale preservation
- Interactive element manipulation:
  - Selection of single or multiple whip points
  - Drag-and-drop repositioning of connection points
  - Rectangle selection for multiple elements
  - Context menus for quick actions
- Status and warning visualization:
  - Current load indicators with color coding
  - Warning panel for electrical issues
  - Visual highlights for overloaded segments
  - Interactive wire highlighting on warning selection
- Tracker selection with visual highlighting
- Real-time calculation updates
- Dual-unit display (feet/meters) for all dimensions
- Selection highlighting for interactive elements
- Multi-select capability with shift/ctrl modifiers
- Visual feedback for selected elements with color changes
- Drag to select multiple elements with selection box
- Keyboard shortcuts for selection manipulation (Ctrl+A, Delete, Esc)

## 5. Data Structures

### 5.1 Module Specification
```python
class ModuleSpec:
    # See detailed model in section 2.1
```

### 5.2 Tracker Template
```python
class TrackerTemplate:
    # See detailed model in section 2.2
```

### 5.3 Block Configuration
```python
class BlockConfig:
    # See detailed model in section 2.3
```

### 5.4 Inverter Specification
```python
class InverterSpec:
    # See detailed model in section 2.4
```

### 5.5 Wiring Configuration
```python
class WiringConfig:
    # See detailed model in section 2.5
```

### 5.6 Project and Metadata
```python
class Project:
    # See detailed model in section 2.6
    
class ProjectMetadata:
    # See detailed model in section 2.6
```

### 5.7 Tracker Position
```python
class TrackerPosition:
    x: float  # X coordinate in meters
    y: float  # Y coordinate in meters
    rotation: float  # Rotation angle in degrees
    template: TrackerTemplate
    strings: List[StringPosition] = field(default_factory=list)  # List of strings on this tracker
```

### 5.8 String Position
```python
class StringPosition:
    index: int  # Index of this string on the tracker
    positive_source_x: float  # X coordinate of positive source point relative to tracker
    positive_source_y: float  # Y coordinate of positive source point relative to tracker
    negative_source_x: float  # X coordinate of negative source point relative to tracker
    negative_source_y: float  # Y coordinate of negative source point relative to tracker
    num_modules: int  # Number of modules in this string
```

### 5.9 Collection Point
```python
class CollectionPoint:
    x: float
    y: float
    connected_strings: List[int]  # List of string IDs connected to this point
    current_rating: float
```

### 5.10 Device Input Point
```python
class DeviceInputPoint:
    index: int  # Input number
    x: float  # X coordinate relative to device corner
    y: float  # Y coordinate relative to device corner
    max_current: float  # Maximum current rating for this input
```

### 5.11 MPPT Channel
```python
class MPPTChannel:
    max_input_current: float
    min_voltage: float
    max_voltage: float
    max_power: float
    num_string_inputs: int
```

### 5.12 State Management
```python
class UndoState:
    state: Any
    description: str

class UndoManager:
    undo_stack: List[UndoState]
    redo_stack: List[UndoState]
    max_states: int
    
    # Methods for state management
    push_state(description: str)
    undo() -> Optional[str]
    redo() -> Optional[str]
    can_undo() -> bool
    can_redo() -> bool
```

### 5.13 Harness Group
```python
class HarnessGroup:
    string_indices: List[int]  # Indices of strings in this harness
    cable_size: str  # Cable size for this harness
```

## 6. Implementation Phases

### Phase 1: Core Framework (Complete)
1. Set up project structure
2. Implement data models (Module, Tracker, Inverter)
3. Create basic UI framework
4. Implement file parsing (.pan, .ond)

### Phase 2: Template System (Complete)
1. Implement tracker template creator
2. Add template management
3. Create template visualization
4. Implement validation rules

### Phase 3: Block Configuration (Complete)
1. Develop block configuration interface
2. Implement drag-and-drop functionality
3. Add grid system
4. Create block visualization

### Phase 4: Wiring System (Complete)
1. Implement string homerun configuration
2. Add wire harness functionality
3. Develop electrical validation
4. Create wiring visualization

### Phase 5: BOM Generation (Complete)
1. Implement quantity calculations
2. Add length calculations
3. Create Excel export functionality
4. Add configuration save/load features

### Phase 6: Project Management (Complete)
1. Implement project creation and management
2. Add project dashboard
3. Add project search and filtering
4. Create project save/load functionality

### Phase 7: Advanced Wiring Features (Complete)
1. Implement custom harness grouping
2. Add interactive whip point positioning
3. Develop electrical warning system
4. Create realistic routing calculations for BOM

## 7. Output Requirements

### 7.1 BOM Excel Format
- Project information sheet
- Component summary sheet:
  - Part numbers
  - Quantities
  - Lengths
  - Ratings
- Block detail sheet with per-block breakdown
- Components to include:
  - String cables (by polarity and size)
  - Wire harnesses (by string count and size)
  - Wire management
  - Whips (by polarity and size)
  - Above ground wire management
  - Standard segment lengths with counts

### 7.2 Configuration Saves
- JSON format for all configurations
- Separate files for templates and blocks
- Complete project state preservation
- Module specs include additional fields like efficiency and bifaciality
- Tracker templates store full module specifications
- Inverter configurations maintain detailed MPPT channel information
- Wiring configurations with cable sizes and current ratings
- Custom whip point positions and harness groupings

## 8. Testing Requirements
- Unit tests for core functionality
- Integration tests for UI components
- Validation testing for electrical rules
- User acceptance testing for workflow
- Performance testing for large layouts
- Electrical limit testing with various wire configurations

## 9. Directory Structure
```
solar_bom/
│
├── src/
│   ├── __init__.py
│   ├── models/
│   │   ├── __init__.py
│   │   ├── module.py      # ModuleSpec, ModuleType, ModuleOrientation
│   │   ├── inverter.py    # InverterSpec, MPPTChannel, MPPTConfig
│   │   ├── tracker.py     # TrackerTemplate, TrackerPosition
│   │   ├── block.py       # BlockConfig, WiringType, DeviceType
│   │   └── project.py     # Project, ProjectMetadata
│   │
│   ├── ui/
│   │   ├── __init__.py
│   │   ├── tracker_creator.py     # TrackerTemplateCreator UI
│   │   ├── block_configurator.py  # BlockConfigurator UI
│   │   ├── module_manager.py      # ModuleManager UI
│   │   ├── inverter_manager.py    # InverterManager UI
│   │   ├── wiring_configurator.py # WiringConfigurator UI
│   │   ├── bom_manager.py         # BOMManager UI
│   │   └── project_dashboard.py   # ProjectDashboard UI
│   │
│   └── utils/
│       ├── __init__.py
│       ├── pan_parser.py          # .pan file parsing
│       ├── file_handlers.py       # File operations
│       ├── undo_manager.py        # Undo/redo state management
│       ├── calculations.py        # Shared calculations/validations
│       ├── project_manager.py     # Project operations management
│       └── bom_generator.py       # BOM generation utilities
│
├── data/                          # Persistent storage directory
│   ├── tracker_templates.json     # Saved tracker templates
│   ├── module_templates.json      # Saved module specifications
│   └── inverter_templates.json    # Saved inverter specifications
│
├── projects/                      # Project storage directory
│   ├── .recent_projects           # Recent projects tracking
│   └── *_project.json            # Individual project files
│
├── tests/
│   ├── __init__.py
│   ├── test_models/               # Model unit tests
│   ├── test_ui/                   # UI component tests
│   └── test_utils/                # Utility function tests
│
└── main.py                        # Main application entry point
```

## 10. User Interface Guidelines

### 10.1 Block Configuration Interface
- Maintain 1:1 aspect ratio for block visualization
- Use semi-transparent colors for clearance zones
- Provide visual feedback for invalid placements
- Support both pixel and meter-based coordinate systems
- Implement smooth pan and zoom transitions
- Show device spacing in both feet and meters
- Display tracker selection with visual highlighting
- Support keyboard shortcuts for common operations

### 10.2 Grid System
- Draw gridlines based on configurable spacing values
- Extend grid beyond visible area for panning
- Use dashed lines for grid visualization
- Provide visual distinction between primary and secondary grid lines
- Update grid based on row spacing and N/S tracker spacing

### 10.3 Scale Management
- Default scale: 10 pixels per meter
- Minimum scale: 5 pixels per meter
- Maximum scale: 50 pixels per meter
- Maintain scale during window resizing
- Implement smooth zoom transitions

### 10.4 Device Visualization
- Standard device size: 3ft x 3ft (0.91m x 0.91m)
- Minimum clearance: 1ft (0.3m)
- Default clearance: 6ft (1.83m)
- Show clearance zones with semi-transparent fill
- Update clearance visualization in real-time when modified
- Display device input points
- Show device type indicator

### 10.5 Wiring Visualization
- Use red for positive cables and blue for negative cables
- Line thickness corresponds to wire gauge
- Show collection points with distinct colors
- Display whip points for cable connections
- Provide optional current labels on wire segments
- Implement real-time updating as configuration changes
- Highlight overloaded segments with warning indicators
- Show accumulated current at harness junction points

### 10.6 Project Dashboard
- Card-based layout for recent projects
- Table view for all projects
- Search and sort controls
- Visual feedback for active project
- Simple project creation flow
- Confirmation for destructive actions

### 10.7 Whip Point Interaction
- Highlight selected whip points with color changes
- Support multi-selection via drag box or Ctrl+click
- Display context menu for quick actions
- Provide visual feedback during dragging operations
- Animate highlights when clicking on warning messages
- Enable keyboard shortcuts for common operations

## 11. Implementation Best Practices

### 11.1 State Management
- Save state before any block modification
- Maintain separate undo/redo stacks per block
- Clear redo stack when new action is performed
- Limit undo history to prevent memory issues
- Provide clear descriptions for undo/redo actions
- Deep copy state to prevent reference issues
- Preserve wiring configuration state in undo/redo operations

### 11.2 Coordinate Systems
- Store all positions in meters
- Convert between pixels and meters using scale factor
- Account for pan offset in all coordinate calculations
- Apply grid snapping after coordinate conversion
- Validate positions against block boundaries
- Support both imperial and metric unit display

### 11.3 Performance Considerations
- Limit canvas redraws to necessary updates
- Use canvas tags for grouped elements
- Implement efficient pan and zoom operations
- Cache calculated values where appropriate
- Clear unused canvas elements
- Optimize large layout rendering

### 11.4 File Operations
- Validate file paths before operations
- Handle file exceptions gracefully
- Create directories as needed
- Clean up filenames to ensure validity
- Check file extensions before processing
- Use proper JSON formatting for configuration files

### 11.5 Electrical Calculations
- Implement voltage drop calculations based on industry standards
- Calculate conductor ampacity per NEC guidelines
- Apply temperature corrections to calculations
- Calculate power loss for system efficiency analysis
- Validate string configurations against inverter specifications
- Apply proper safety factors to current calculations
- Separate calculations for different wire gauges

## 12. Advanced Features

### 12.1 Electrical Calculations
- Voltage drop calculations for DC circuits
- Power loss calculations
- Conductor ampacity based on NEC guidelines
- String electrical characteristics with temperature effects
- Wire harness compatibility checking
- Segment-specific load calculations

### 12.2 Block Management Enhancements
- Copy existing blocks with incremental naming
- Rename blocks with validation
- Enhanced selection and manipulation of trackers
- Device positioning with real-time validation of clearances

### 12.3 Wiring Configuration Enhancements
- Support for different cable sizes with visual representation
- Current calculations and visualization
- Collection point current ratings
- Whip point identification and routing
- Node-to-node harness routing optimization
- Custom harness grouping with string-level control
- Interactive whip point repositioning
- Multiple independent harnesses per tracker

### 12.4 User Interface Improvements
- Real-time GCR calculations and display
- Dual-unit display for dimensions (feet/meters)
- Enhanced pan and zoom functionality
- Tracker selection with visual highlighting
- Device zone visualization with semi-transparency
- Tab-based application interface
- Status bar with project information
- Electrical warning panel with interactive highlighting
- Context menus for common operations