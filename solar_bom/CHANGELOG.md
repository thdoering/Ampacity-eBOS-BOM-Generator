# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.7.0] - 2026-02-07

### Added
- **Quick Estimate Tab**: New streamlined interface for rapid project estimates with dedicated Excel BOM export
- **DC Feeder Lengths**: Added input fields for DC feeder cable lengths per block for accurate material takeoffs
- **Revision Tracking**: Added revision number input for Excel BOM version control
- **Module Orientation**: Added module orientation input for improved layout specifications

### Changed
- **BOM Formatting**: Multiple Excel BOM formatting improvements for better readability
- **Removed Category Column**: Streamlined BOM by removing redundant category column

### Fixed
- Various bug fixes and stability improvements throughout the application

### Enhanced
- Faster estimate workflows with dedicated Quick Estimate interface
- More accurate DC cable calculations with block-specific feeder lengths
- Better project documentation with revision tracking

## [2.6.0] - 2026-01-13

### Added
- **Tracker Multi-Select and Move**: Select and move multiple trackers simultaneously for faster layout adjustments
- **2P Tracker Support**: Added preliminary support for 2-portrait tracker configurations (work in progress)
- **Batch Wire Sizing**: Added button to apply recommended wire sizes to all blocks at once
- **Whip Pricing**: Updated pricing data to include previously missing whip cables

### Changed
- **Default Wire Sizes**: All wires now default to recommended cable sizes automatically
- **Harness Geometry**: Reverted to identical positive/negative harness routing with differentiated extender cables

### Enhanced
- Faster project layout workflow with multi-tracker operations
- Improved cable sizing efficiency with bulk apply function
- More accurate cost estimates with complete whip pricing

## [2.5.0] - 2025-12-11

### Added
- **Automatic Pricing**: Integrated pricing calculation with importable pricing data for accurate cost estimates
- **Project Import/Export**: Added file sharing functionality to import and export projects for team collaboration
- **Multiple Module Types**: Enabled support for projects using different module types across blocks
- **User-Configurable NEC Factor**: Added ability to customize NEC safety factor in device configurator
- **Shift+Select Function**: Implemented shift+click multi-select for string selection in harness configurator
- **Module Wattages Display**: Added module wattage information to block allocation tab in Excel BOM

### Changed
- **Custom Part Labels**: BOM now displays "CUSTOM" instead of "N/A" for custom parts (improved clarity)

### Fixed
- **Tracker Auto-populate Bug**: Corrected issue with automatic tracker template population

### Enhanced
- More flexible project configuration with multiple module type support
- Improved cost estimation capabilities with automatic pricing
- Better team collaboration with project sharing features
- Enhanced harness configuration workflow with multi-select functionality

## [2.4.1] - 2025-11-17

### Added
- **README.txt File**: Included download instructions and "What's New" documentation with executable
- **Block Allocation Tab**: Added dedicated block allocation tab to Excel BOM export for better project organization
- **Combiner Box Information**: Enhanced BOM to include detailed combiner box specifications
- **Cable Sizing Note**: Added preliminary cable sizes disclaimer to Excel BOM exports

### Changed
- **Row Spacing Precision**: Increased row spacing precision to 3 decimal points (from 1) for more accurate layouts
- **Tracker Template Auto-population**: Template names now automatically populate when selected
- **Block Configurator Interface**: Removed irrelevant input fields for streamlined configuration

### Enhanced
- Improved Excel BOM organization with better tab structure
- Cleaner block configuration workflow with focused inputs
- More precise project layout capabilities

## [2.3.0] - 2025-10-06

### Added
- **SLD Generator**: Implemented Single Line Diagram data model and canvas integration in BOM Manager
- **ANSI Symbol Library**: Added comprehensive ANSI symbol library for professional electrical diagrams
- **Thermal Parameters**: Added thermal parameter inputs to module specification page for accurate performance calculations
- **Wiring Bulk Selection**: Added select all/deselect all buttons in wiring configurator for easier configuration

### Changed
- **NEC Safety Factor**: Updated NEC safety factor to 1.56 across all relevant electrical calculations for code compliance
- **Row Spacing Lock**: Row spacing is now locked once trackers are placed to prevent layout inconsistencies
- **Wiring Mode Restriction**: Removed conceptual routing option - only realistic routing is now available for more accurate installations
- **SLD Drawing Elements**: Enhanced drawing elements and symbols in SLD generator for better diagram quality

### Fixed
- **Harness Drawing Generator**: Resolved multiple bugs in harness drawing generation
- **SLD Dragging**: Fixed dragging behavior issues in SLD generator canvas
- **Cable Totals**: Corrected bug that was doubling cable length totals in BOM calculations

### Notes
- Projects with conceptual wiring configurations will need to be reconfigured using realistic routing
- NEC safety factor change may affect cable sizing in existing projects - review electrical calculations

## [2.2.0] - 2025-08-22

### Added
- **Block List Sorting**: Added sorting functionality for better organization of block lists
- **Harness-Specific Cable Sizes**: Enhanced wiring configurator with harness-specific cable sizing for more accurate material requirements
- **Clear All Trackers Button**: Added convenient button to clear all tracker placements from blocks at once

### Enhanced
- Improved user workflow with better block management controls
- More precise cable sizing calculations for procurement accuracy

## [2.1.0] - 2025-07-17

### Added
- **Updated Harness Library**: Enhanced wire harness specifications with long trunk ends for improved installation flexibility
- **Automatic Part Numbers**: BOM generator now automatically pulls in standard part numbers for all components
- **Motor Placement Control**: Added ability to specify motor placement within tracker strings
- **Underground Routing Controls**: Added controls for routing whip cables underground before connecting to devices
- **Leapfrog Wiring Input**: Added support for leapfrog wiring configuration as an alternative to daisy chain
- **Device Configuration Tab**: Created new tab for configuring combiner boxes and other electrical devices
- **Excel File Naming**: Automated suggestion of appropriate Excel file names based on project details
- **Collapsible Wiring Sections**: Made sections in wiring configurator collapsible for better organization
- **Combiner Boxes Tab**: Added dedicated combiner boxes tab to Excel BOM export

### Changed
- **GCR/Row Spacing**: Made GCR and row spacing inputs bidirectional - updating one automatically updates the other
- **Whip Point Manipulation**: Allow realistic manipulation of routing whip connection points
- **BOM Descriptions**: Updated all BOM descriptions to match standard part library nomenclature
- **BOM Project Info**: Enhanced project information section in BOM exports
- **Excel BOM Styling**: Improved Excel BOM formatting and styling for better readability
- **Tracker Template Organization**: Templates now categorized based on module model and string size
- **Wiring Configurator Access**: Allow wiring configurator to remain open while working in other tabs

### Fixed
- **String Homerun Configurations**: Fixed bugs related to string homerun cable calculations and configurations
- **Tracker Template Display**: Resolved issue where tracker templates weren't displaying correctly in selection dialog
- **Various Bug Fixes**: Multiple minor bug fixes and performance improvements throughout the application

## [2.0.0] - 2025-06-27

### BREAKING CHANGES
- **Template Organization**: Module and tracker templates are now organized hierarchically by manufacturer
- **Project Compatibility**: Projects created with previous versions will load but tracker placements will need to be recreated due to template structure changes

### Added
- **Part Number Column**: Added part number column to BOM export for better procurement tracking
- **Harness Drawings**: Created ability to provide harness drawings to customers
- **Copy Project Functionality**: Added ability to copy entire projects with all configurations
- **Hierarchical Template Browser**: Tree view interface for module and tracker templates organized by manufacturer
- **Project-Specific Template Filtering**: Added tracker template filtering with checkboxes for better project organization
- **Automatic Part Number Integration**: BOM generator now automatically pulls fuse and harness standard part numbers
- **Enhanced BOM Functionality**: Expanded BOM generation capabilities with more detailed component tracking

### Changed
- **Template File Format**: Module and tracker template files converted to hierarchical structure organized by manufacturer
- **Block Naming Convention**: Improved naming convention for new blocks with better sequential numbering
- **Template Selection Interface**: Replaced flat template lists with expandable tree views for better organization

### Fixed
- **Whip Calculation Bug**: Corrected whip cable length calculations for accurate BOM output
- **Undo/Redo System**: Fixed undo/redo functionality to properly maintain state history
- **Row Spacing Persistence**: Fixed issue where row spacing settings weren't properly saved/restored
- **Motor Placement Bug**: Corrected tracker motor positioning calculations
- **Wiring Current Calculation**: Fixed current calculation bug in wiring configurations
- **Template Loading**: Fixed template loading issues when switching between projects

### Enhanced
- **Autosave**: Automatic project saving whenever new blocks are created or copied
- **Template Backwards Compatibility**: Maintains compatibility with existing template references during migration
- **User Experience**: Improved template browsing with manufacturer-based organization

### Migration Notes
- Open existing projects and re-place trackers in blocks using the new hierarchical template browser
- Module and tracker template libraries will be automatically migrated to new format on first use
- Template references in saved projects maintain backwards compatibility
- All other project data (wiring configurations, BOM settings, etc.) will be preserved

### Technical Improvements
- Enhanced template loading with support for both legacy and new formats
- Improved error handling for template migration
- Better state management for block configurations
- Optimized template lookup performance with hierarchical structure

## [1.1.0] - 2025-05-30

### Added
- Custom harness configuration functionality for advanced string grouping
- Realistic and conceptual wiring routing modes in wiring configurator
- Extender wire support for complex installation scenarios
- Enhanced tracker motor placement capabilities
- Whip cable segments analysis in BOM generation
- Numerical ordering for block lists for better organization

### Changed
- Moved wiring calculations from block configurator to wiring configurator for improved accuracy
- Enhanced BOM generation with more detailed wire segment analysis

### Fixed
- Various bug fixes and stability improvements

### Notes
- Projects created with older versions should remain compatible
- New wiring features provide more accurate cable length calculations

## [1.0.0] - 2025-04-28

### Added
- Initial release of Solar eBOS BOM Generator
- Module management with import capabilities for .pan files
- Tracker template creation and management
- Block layout configuration with drag-and-drop interface
- Wiring configuration with string homerun and wire harness options
- BOM generation with Excel export
- Project management system with save/load functionality

### Known Issues
- [List any known issues or limitations here]