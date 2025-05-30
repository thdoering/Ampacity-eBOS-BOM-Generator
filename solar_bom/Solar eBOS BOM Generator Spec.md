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
  - Motor position configuration
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
    motor_position_after_string: int = 0  # Motor position (0 means calculate default)
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
  - Configurable placement modes (row center vs tracker aligned)
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
    block_realistic_routes: Dict[str, List[tuple[float, float]]] = field(default_factory=dict)
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

#### Routing Modes

The system supports two distinct routing calculation modes:

1. **Realistic Routing Mode**:
   - Cable routes follow actual installation patterns
   - Routes along tracker centerlines and torque tubes
   - Accounts for physical tracker structure
   - Provides accurate cable length calculations for BOM
   - Whip points align with harness cables for optimal installation
   - More accurate representation of field installation practices

2. **Conceptual Routing Mode**:
   - Simplified routing for visualization purposes
   - Direct point-to-point connections with minimal waypoints
   - Faster calculation and rendering
   - Suitable for preliminary design and understanding system topology
   - May not reflect actual installation cable lengths

#### Wire Routing Rules

- All cable routes follow a vertical-then-horizontal pattern in conceptual mode
- Realistic mode routes follow tracker structure and centerlines
- Maintain visual separation between parallel cable runs
- Route cables along tracker edges with appropriate offsets
- Positive cables route along left side of tracker, negative along right
- Routes determined by source point location relative to destination

#### Wire Harness Collection Points

- Node points placed near string source points with calculated offsets
- Node points follow positioning rules based on routing mode
- Each tracker maintains independent harness system with optional grouping
- Support for combining harnesses between trackers via extender system
- Collection points have current ratings and capacity validation

#### Visualization Standards

- Red indicates positive polarity cables/points
- Blue indicates negative polarity cables/points
- Line thickness corresponds to wire gauge size
- Visual offsets between parallel runs for clarity
- Source points, node points, and destination points visually distinct
- Current labels for visualizing electrical loads on wire segments
- Different line thicknesses based on wire gauge
- Color coding by cable importance (darker = higher current/importance)

#### Wiring Configuration Types

Support two primary wiring approaches with advanced configuration options:

1. **String Homeruns**:
   - Individual positive/negative cables per string
   - Direct connection to downstream device
   - Cable paths automatically generated from each string to device
   - Cable sizing based on single string current
   - More inputs used on downstream device
   - Longer total cable length
   - Simpler installation but higher material cost

2. **Wire Harness Solution**:
   - Collection point at each tracker with advanced grouping options
   - Separate positive and negative wire harnesses
   - Combined strings at collection points
   - Advanced string grouping capabilities for creating custom harnesses
   - Custom harness configuration with individual string selection
   - Multiple independent harnesses per tracker with different cable sizes
   - Support for individual whip point positioning for routing optimization
   - Realistic cable routing calculation options for accurate BOM generation

#### Advanced Harness Configuration Options

- **Flexible String Grouping**: Group any combination of strings within trackers
- **Multiple Independent Harnesses**: Create multiple harnesses per tracker, each with its own cable sizing and routing
- **Custom Cable Sizing**: Separate cable sizing for string, harness, whip, and extender cables
- **Interactive Whip Point Management**: 
  - Drag-and-drop whip point positioning
  - Visual feedback for selected points
  - Single and multi-point selection
  - Reset to default positions
  - Persistence of custom positions between sessions
- **Quick Pattern Templates**: Pre-configured harness patterns for common scenarios:
  - Split evenly in two harnesses
  - Separate furthest string from others
  - Default single harness configuration
- **Fuse Configuration**: 
  - Configurable fuse ratings per harness
  - Automatic NEC-compliant fuse sizing recommendations
  - Option to disable fuses for single-string harnesses
  - Fuse count calculations for BOM generation

#### Extender Cable System

Advanced extender cable management for complex installations:

- **Multi-Harness Extenders**: When trackers have multiple harnesses, secondary harnesses use extender cables to reach shared whip points
- **Stacked Tracker Extenders**: Trackers positioned further from devices automatically use extenders
- **Intelligent Extender Routing**: System determines optimal extender placement based on:
  - Tracker position relative to device
  - Harness configuration
  - Cable length optimization
- **Extender Point Visualization**: Dedicated visual indicators for extender connection points
- **Automatic Extender Calculation**: System automatically determines when extenders are needed

#### Cable Specifications

Support for comprehensive cable sizing with industry-standard gauges:
- **Available Sizes**:
  - 4 AWG (21.15 mm²)
  - 6 AWG (13.30 mm²)
  - 8 AWG (8.37 mm²)
  - 10 AWG (5.26 mm²)
- **Independent Configuration**: Separate sizing for string, harness, whip, and extender cables
- **Visual Representation**: Different cable sizes shown with appropriate line thickness
- **Current Calculations**: Based on module specifications and number of combined strings
- **Interactive Current Labeling**: Toggle for detailed electrical analysis with draggable labels

#### Interactive Whip Point Management

- **Drag-and-Drop Interface**: 
  - Click and drag individual or multiple whip points
  - Visual selection feedback with color changes
  - Rectangle selection for multiple points
  - Keyboard shortcuts (Ctrl+A, Delete, Esc)
- **Position Memory**: 
  - Custom positions persist between sessions
  - Optional use of custom positions for BOM calculations
  - Quick reset options for individual or all whip points
- **Visual Differentiation**: 
  - Color-coding by polarity and selection state
  - Size changes for selected points
  - Distinct visualization for harness-specific points
- **Context Menus**: Right-click access to common operations

#### Electrical Validation and Warning System

- **Real-Time Current Calculation**: Automatic calculation of current loads through all wire segments
- **NEC Compliance Checking**: Validation against National Electrical Code requirements
- **Ampacity Verification**: Check wire gauge capacity against calculated loads
- **Visual Warning System**: 
  - Color-coded warnings for different severity levels
  - Interactive warning panel with clickable items
  - Wire highlighting when warnings are selected
  - Overload indicators with percentage calculations
- **MPPT Capacity Validation**: Verify total string current against inverter MPPT capacity
- **Warning Categories**:
  - Caution (60-80% capacity)
  - Warning (80-100% capacity)
  - Overload (>100% capacity)

#### Automatic Routing Algorithm

Enhanced routing calculations based on selected mode:

- **Realistic Mode Routes**:
  - Follow tracker centerlines and torque tube structures
  - Account for module layout and physical constraints
  - Optimize for actual installation practices
  - Provide accurate cable length calculations
- **Conceptual Mode Routes**:
  - Simplified point-to-point connections
  - Maintain visual clarity and understanding
  - Faster calculation for large systems
- **Route Optimization**: 
  - Minimum cable length where possible
  - Clearance maintenance from other equipment
  - Obstacle avoidance algorithms
- **Node-to-Node Connections**: Sophisticated multi-point routing for wire harness configurations

#### Wiring Configuration Data Model
```python
class WiringConfig:
    wiring_type: WiringType
    positive_collection_points: List[CollectionPoint]
    negative_collection_points: List[CollectionPoint]
    strings_per_collection: Dict[int, int]  # Collection point ID -> number of strings
    cable_routes: Dict[str, List[tuple[float, float]]]  # Route ID -> list of coordinates
    realistic_cable_routes: Dict[str, List[tuple[float, float]]] = field(default_factory=dict)  # Realistic routes for BOM
    string_cable_size: str = "10 AWG"  # Default string cable size
    harness_cable_size: str = "8 AWG"  # Default harness cable size
    whip_cable_size: str = "8 AWG"  # Default whip cable size
    extender_cable_size: str = "8 AWG"  # Default extender cable size
    custom_whip_points: Dict[str, Dict[str, tuple[float, float]]] = field(default_factory=dict)   # Format: {'tracker_id': {'positive': (x, y), 'negative': (x, y)}}
    harness_groupings: Dict[int, List[HarnessGroup]] = field(default_factory=dict)
    custom_harness_whip_points: Dict[str, Dict[int, Dict[str, tuple[float, float]]]] = field(default_factory=dict)  # Format: {'tracker_id': {harness_idx: {'positive': (x, y), 'negative': (x, y)}}}
    use_custom_positions_for_bom: bool = False  # Use custom positions for BOM calculations
    routing_mode: str = "realistic"  # "realistic" or "conceptual"
```

#### Validation Rules

- Maximum MPPT current limits with real-time feedback
- Available input connections validation
- NEC current limits with visual indicators:
  - Individual string cables
  - Combined current in wire harnesses
  - Extender cable capacity
- Maximum inputs per downstream device
- Collection point current ratings
- Wire gauge ampacity verification
- Interactive warning system for electrical issues

### 2.6 Project Management System

#### Features

- Project metadata management:
  - Project name, description, and location
  - Client information
  - Creation and modification dates
  - Notes and additional documentation
  - Default row spacing configuration
- Project dashboard:
  - Recent projects with card-based presentation
  - Full project list with sorting and filtering
  - Search functionality across project metadata
  - One-click project access
- Project operations:
  - Create new projects with configurable defaults
  - Open existing projects
  - Save project updates with automatic modification date tracking
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

- **Advanced Component Calculation**:
  - Automatic quantity calculation based on block configurations
  - Intelligent cable length calculations using routing mode data
  - Harness count based on number of strings and custom groupings
  - Support for multiple component categories with polarity separation
  - Segment-based cable length calculations for higher accuracy
  - Wire size-specific component categorization
  - Support for mixed cable sizes in complex harness configurations
  - Configurable use of realistic vs conceptual routing for calculations
  - Automatic segment length rounding to standard increments with waste factors

- **Routing Mode Impact on BOM**:
  - **Realistic Mode**: Uses actual installation routing for accurate cable lengths
  - **Conceptual Mode**: Uses simplified routing (may not reflect actual installation)
  - **Warning System**: Alerts users when BOM uses conceptual routing
  - **Route Priority**: BOM calculations prioritize block configurator's realistic routes

- **Advanced Cable Segment Analysis**:
  - Breakdown of cable runs into practical installation segments
  - Length-specific segment counts for installation planning
  - Standardized length increments with appropriate waste factors
  - Separation of string, harness, whip, and extender cable segments
  - Support for mixed cable gauge systems within the same installation
  - Individual segment tracking with counts and lengths

- **BOM Preview and Validation**:
  - Real-time BOM generation with instant updates
  - Component categorization and grouping by type and polarity
  - Block-specific and project-wide views
  - Electrical validation warnings integrated into BOM
  - Missing configuration detection and warnings

- **Excel Export with Enhanced Features**:
  - Formatted Excel output with comprehensive project information
  - Multiple sheets for summary, detailed views, and project data
  - Automatic column sizing and professional formatting
  - Project statistics summary with electrical configuration details
  - Warning indicators for sections with electrical concerns
  - Enhanced project statistics and electrical configuration summary

#### Component Categorization

- **eBOS Components**:
  - String cables (by polarity and size)
  - Wire harnesses (by string count and size)
  - Whip cables (by polarity and size)
  - Extender cables (by polarity and size)
  - Above ground wire management
- **Electrical Components**:
  - DC fuses (by rating and quantity)
  - Connection hardware
  - Wire management accessories
- **Structural Elements**: 
  - Tracker assemblies (by string count)
  - Mounting hardware

#### BOM Export Format

- **Project Information Sheet**:
  - Project name, client, and location
  - System size and module specifications
  - Inverter and DC collection types
  - Project notes and description
  - Electrical configuration summary
  - Warning count and status indicators

- **BOM Summary Sheet**:
  - Component types and descriptions
  - Total quantities with appropriate units
  - Category grouping with visual separation
  - Polarity-specific component listings
  - Wire gauge and current rating information

- **Block Details Sheet**:
  - Per-block component breakdown
  - Component types and quantities by block
  - Detailed component specifications
  - Block-specific electrical loads

- **Segment Analysis Sheet**:
  - Segment-based wire listings with specific lengths
  - Detailed harness breakdown by string count and cable size
  - Installation-ready cable cut lists
  - Waste factor calculations

#### Cable Route Calculation Priority

- **Primary Source**: Block configurator's realistic routes (when available)
- **Fallback**: Wiring configurator routes for visualization
- **BOM Accuracy**: Realistic routing provides better estimates for:
  - String cables routed along torque tubes to whip points
  - Whip cables routed from whip points to device center
  - Routes considering tracker structure and physical installation constraints
- **Warning System**: Alerts when conceptual routing is used for BOM calculations

#### Advanced BOM Features

- **Electrical Load Analysis**: Integration of current calculations and NEC compliance
- **Multi-Configuration Support**: Handle projects with mixed wiring types
- **Segment Optimization**: Intelligent cable segment analysis for procurement optimization
- **Installation Planning**: Cable cut lists organized by installation sequence
- **Quality Assurance**: Validation checks for missing configurations and electrical issues

### 2.8 Whip Point Management

#### Features
- **Interactive Positioning**:
  - Drag-and-drop interface for whip point placement
  - Visual feedback for selected points with size and color changes
  - Single and multi-point selection with rectangle selection
  - Reset to default positions (individual or bulk)
- **Harness-Specific Management**:
  - Independent whip points for each harness group
  - Vertical offset handling for multiple harnesses
  - Custom routing from collection nodes to whip points
- **Position Persistence**:
  - Memory of custom point positions between sessions
  - Optional use of custom positions for BOM calculations
  - Quick reset options with confirmation
  - Context menus for common operations
- **Visual System**:
  - Color-coding by polarity and selection state
  - Size changes indicating selection status
  - Distinct visualization for harness-specific points
  - Leader lines for distant label positioning

## 3. Technical Requirements

### 3.1 Data Validation

- Inverter compatibility checks with real-time feedback
- NEC compliance validation for current limits with visual warnings
- MPPT input validation against total system current
- String cable sizing validation with ampacity checking
- Harness cable sizing validation with accumulated current analysis
- Wire gauge selection validation against current loads
- Visual warning system for approaching/exceeding ampacity limits
- Highlighting of problematic wire segments with overload indicators
- MPPT channel capacity validation against total string current
- Interactive warning panel with problem descriptions and locations
- Real-time electrical calculations with temperature effects

### 3.2 Calculations

- **Cable Length Calculations**: Based on realistic or conceptual routing modes
- **Voltage Drop Calculations**: Industry-standard formulas with temperature correction
- **Power Loss Calculations**: System efficiency analysis
- **Current Calculations**: For wire harness solutions with accumulated loads
- **Conductor Ampacity Calculations**: Per NEC guidelines with safety factors
- **String Electrical Characteristics**: With temperature effects and validation
- **Wire Harness Compatibility**: Checking with current load visualization
- **Ground Coverage Ratio (GCR)**: Real-time calculations
- **Load Percentage Calculations**: Relative to NEC requirements
- **Ampacity Verification**: For different wire gauges (4, 6, 8, 10 AWG)
- **Harness Compatibility**: Checking with current load visualization
- **Dynamic Current Labeling**: With visual indicators and drag capability

### 3.3 File Handling

- Import .ond files for inverter specifications
- Import .pan files for module specifications
- Export BOM to Excel format with multiple sheets
- Save/load complete project configurations with version tracking
- Directory structure verification and creation
- Filename validation and cleanup
- Error handling for file operations with user feedback
- Support for multiple file types and validation
- Backup and recovery mechanisms

### 3.4 State Management

- Undo/redo system for block configurations with descriptive actions
- State preservation for project loading/saving
- Deep state copying to prevent reference issues
- Error recovery for file operations
- Consistent state validation
- Configuration persistence across sessions
- Real-time state synchronization between UI components

## 4. User Interface Requirements

- Python-based GUI framework (Tkinter)
- Tab-based interface for different functions
- Block-based navigation system
- Drag-and-drop functionality for tracker placement with grid snapping
- Form-based input for specifications with real-time validation
- Real-time validation feedback with visual indicators
- 2D layout visualization with professional rendering
- Error messaging for invalid configurations
- Undo/redo functionality with descriptive actions
- Advanced pan and zoom controls in layout visualizations:
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
  - Warning panel for electrical issues with clickable items
  - Visual highlights for overloaded segments
  - Interactive wire highlighting on warning selection
- Real-time calculation updates with instant feedback
- Dual-unit display (feet/meters) for all dimensions
- Selection highlighting for interactive elements
- Multi-select capability with shift/ctrl modifiers
- Visual feedback for selected elements with color changes
- Keyboard shortcuts for selection manipulation (Ctrl+A, Delete, Esc)
- Professional color schemes and visual hierarchy

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
    fuse_rating_amps: int = 15  # Fuse rating in amps
    use_fuse: bool = True  # Whether to use fuses for this harness
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
5. Add routing mode selection
6. Implement extender cable system

### Phase 8: Enhanced BOM and Validation (Complete)
1. Advanced segment analysis
2. Routing mode impact on calculations
3. Electrical validation and warning systems
4. Enhanced Excel export with warnings

## 7. Output Requirements

### 7.1 BOM Excel Format
- **Project Information Sheet**: Comprehensive project data and electrical summary
- **Component Summary Sheet**:
  - Part numbers and descriptions
  - Quantities by polarity and cable type
  - Lengths with waste factors
  - Current ratings and wire gauge specifications
- **Block Detail Sheet**: Per-block breakdown with electrical loads
- **Segment Analysis Sheet**: Installation-ready cable cut lists
- **Components Include**:
  - String cables (by polarity and size)
  - Wire harnesses (by string count and size)
  - Whip cables (by polarity and size)
  - Extender cables (by polarity and size)
  - DC fuses (by rating and quantity)
  - Wire management accessories
  - Standard segment lengths with counts

### 7.2 Configuration Saves
- JSON format for all configurations with version tracking
- Separate files for templates and blocks
- Complete project state preservation including:
  - Module specs with efficiency and bifaciality data
  - Tracker templates with full module specifications
  - Inverter configurations with detailed MPPT channel information
  - Wiring configurations with cable sizes, current ratings, and routing modes
  - Custom whip point positions and harness groupings
  - Realistic route data for accurate BOM calculations

## 8. Testing Requirements
- Unit tests for core functionality
- Integration tests for UI components
- Validation testing for electrical rules
- User acceptance testing for workflow
- Performance testing for large layouts
- Electrical limit testing with various wire configurations
- BOM accuracy validation against field installations

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
│   │   ├── block.py       # BlockConfig, WiringType, DeviceType, HarnessGroup
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
- **Color Scheme**:
  - Red for positive cables and blue for negative cables
  - Color intensity indicates cable importance (darker = higher current)
  - Line thickness corresponds to wire gauge
- **Interactive Elements**:
  - Show collection points with distinct colors
  - Display whip points for cable connections
  - Provide optional current labels on wire segments with drag capability
  - Implement real-time updating as configuration changes
  - Highlight overloaded segments with warning indicators
  - Show accumulated current at harness junction points
- **Routing Mode Visualization**:
  - Realistic mode: Routes follow tracker centerlines and structure
  - Conceptual mode: Simplified direct routing
  - Visual indicators for current routing mode

### 10.6 Electrical Warning System
- **Warning Panel**: 
  - Floating panel with clickable warning items
  - Color-coded severity levels (caution, warning, overload)
  - Interactive highlighting of problematic wires
- **Visual Indicators**: 
  - Wire segment highlighting with pulsing effects
  - Color-coded current load indicators
  - Percentage loading displays

### 10.7 Project Dashboard
- Card-based layout for recent projects
- Table view for all projects
- Search and sort controls
- Visual feedback for active project
- Simple project creation flow
- Confirmation for destructive actions

### 10.8 Whip Point Interaction
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
- Implement progressive loading for complex projects

### 11.4 File Operations
- Validate file paths before operations
- Handle file exceptions gracefully
- Create directories as needed
- Clean up filenames to ensure validity
- Check file extensions before processing
- Use proper JSON formatting for configuration files
- Implement backup and recovery mechanisms

### 11.5 Electrical Calculations
- Implement voltage drop calculations based on industry standards
- Calculate conductor ampacity per NEC guidelines
- Apply temperature corrections to calculations
- Calculate power loss for system efficiency analysis
- Validate string configurations against inverter specifications
- Apply proper safety factors to current calculations
- Separate calculations for different wire gauges
- Real-time validation with visual feedback

## 12. Advanced Features

### 12.1 Electrical Calculations and Validation
- **Real-Time Analysis**: Continuous electrical validation during design
- **NEC Compliance**: Automatic checking against National Electrical Code
- **Temperature Effects**: Calculations include temperature coefficients
- **Safety Factors**: Proper application of electrical safety margins
- **Load Analysis**: Comprehensive current flow analysis through all segments

### 12.2 Block Management Enhancements
- **Advanced Copying**: Copy blocks with incremental naming
- **Renaming Validation**: Ensure unique block identifiers
- **Enhanced Selection**: Multi-tracker selection and manipulation
- **Device Positioning**: Real-time validation of clearances and placement modes

### 12.3 Wiring Configuration Enhancements
- **Multi-Size Support**: Different cable sizes with visual representation
- **Current Visualization**: Real-time current calculations and display
- **Collection Point Management**: Current ratings and capacity validation
- **Advanced Routing**: Both realistic and conceptual routing modes
- **Interactive Elements**: Drag-and-drop whip point repositioning
- **Multiple Harness Support**: Independent harnesses per tracker with custom configurations

### 12.4 User Interface Improvements
- **Real-Time GCR**: Calculations and display
- **Dual-Unit Display**: Dimensions in both feet and meters
- **Enhanced Navigation**: Pan and zoom functionality
- **Visual Selection**: Tracker selection with highlighting
- **Interactive Zones**: Device zone visualization with semi-transparency
- **Professional Interface**: Tab-based application with status bar
- **Warning Integration**: Electrical warning panel with interactive highlighting
- **Context Menus**: Right-click access to common operations

### 12.5 BOM Generation Advances
- **Routing Mode Integration**: Use of realistic routing for accurate calculations
- **Segment Analysis**: Detailed cable segment breakdown
- **Installation Planning**: Cable cut lists and installation sequences
- **Quality Assurance**: Comprehensive validation and warning systems
- **Export Enhancement**: Professional Excel output with multiple sheets and formatting