# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [3.3.0] - 2026-04-23

### Added
- **2P Tracker Support**: Completed full 2-portrait tracker configuration support
- **Parallel DC Feeders**: Added support for parallel DC feeder configurations
- **Azimuth Rotation**: Added azimuth rotation input to tracker groups in Quick Estimate
- **Canvas String Selection & Drag-to-Device**: Strings can now be selected and dragged to devices directly on the site preview canvas
- **Polyline Measuring Tool**: New measuring tool in site preview with rubber-band preview, per-segment and cumulative distance labels, and persistent save/load
- **Fuses in Quick Estimate BOM**: Quick Estimate BOM now includes inline harness fuse line items

### Changed
- **PDF System Summary Table**: Auto-detects the least-obstructed corner for placement; falls back to shrinking the drawing area if all corners are occupied
- **PDF Export Dialog Removed**: Placement dialog no longer appears before every PDF export
- **Inverter Library Tree View**: Flat listbox replaced with grouped ttk.Treeview organized by manufacturer → model (collapsed by default)
- **Site Preview Toolbar**: Split into two rows for better visibility on smaller screens
- **Manual Allocation Preservation**: Manual string assignments survive structural changes; new trackers default to unallocated
- **Tracker Segment Info Cleanup**: Cleaned up tracker segment display in Quick Estimate

### Fixed
- **Canvas String Move No-Op**: Fixed silent failure in `_canvas_move_strings` caused by the Unallocated device shifting real device indices
- **SAT Overlap Detection**: Replaced `_check_overlaps` with SAT-based polygon overlap detection — correctly handles driveline-sheared and azimuth-rotated groups
- **Edit Devices Validation**: `_validate_move` now tolerates contiguity gaps filled by the Unallocated device
- **Device Data Caching**: Extracted `_ensure_device_data()` for lazy build/cache — canvas string moves no longer require opening the Edit Devices dialog first
- **Combiner Box Placement**: Fixed bug where combiner boxes were placed far from the tracker edge when a group contained multiple-sized trackers and device position was North or South
- **Split Tracker 1-String Harness**: Fixed bug preventing configuration of 1-string harnesses on split trackers
- **Inverter Import/Export**: Fixed inverter data not round-tripping correctly through project import/export
- **Module Width Validation**: Module width is now prohibited from exceeding module length
- **Wire Sync, Inverter Naming, Tracker Split**: Various QoL refinements across device configurator, tracker creator, and Quick Estimate

## [3.2.0] - 2026-04-10

### Changed
- **Driveline Angle Range**: Expanded driveline angle range from 0–45° to -45–45° to support negative angles

### Fixed
- **Configure Device Sorting**: Fixed sorting bug in the Configure Device tab

### Enhanced
- **Import/Export Equipment Data**: Project import/export now includes module, inverter, and tracker template information so equipment libraries don't have to be recreated by the importer

## [3.1.0] - 2026-04-07

### Added
- **PDF Export — Site Preview**: New PDF deliverable of the string allocation site preview with professional titleblock and summary tables
- **Background Labels**: Added background labels to site preview for better orientation and readability

### Changed
- Updated default values across the application

## [3.0.0] - 2026-03-20

### Added
- **Quick Estimate Overhaul**: Completely rebuilt the Quick Estimate tab with site preview, tracker template integration, driveline angle support, half-string trackers, inverter pad placement, manual/automatic string allocation to combiner boxes, DC feeder/AC homerun cable inputs per device, and copy estimate functionality
- **Site Preview**: Interactive site preview with panning, device selection, info pop-ups for CBs/SIs, and motor alignment controls
- **Central Inverter Topology**: Added support for central inverter topology in Quick Estimate with dedicated BOM handling
- **Diagnostics File**: Created diagnostics output to verify calculations for different project configurations
- **Segment Rounding Dropdown**: Added option to consolidate line items in the BOM by rounding cable segments
- **Combiner Boxes Tab (QE)**: Added combiner boxes tab to Quick Estimate Excel BOM export
- **LV Collection Inputs**: Added string homerun, harness, and trunk bus inputs for low-voltage collection
- **Inverter Allocation**: Added inverter allocation to Quick BOM with DC/AC ratio display
- **Collapse/Expand All Button**: Added collapse/expand all in edit devices dialog
- **Per-Device Strings Input**: Strings per device is now a per-device input instead of global
- **Per-Device DC Feeder/AC Homerun**: DC feeder and AC homerun cable sizes assignable per device in the Assign Devices dialog
- **Copper Rate on BOM**: Module polarity orientation and copper rate now included on BOM exports
- **Bulk Copy DC Feeder Lengths**: Button to copy DC feeder lengths across blocks

### Changed
- **Row Spacing Moved to Group Level**: Row spacing and GCR are now per-group inputs instead of global, supporting projects with multiple modules/row spacings
- **String Allocation Algorithm**: Updated to prefer balanced site allocations
- **Terminology Update**: Changed internal terminology from "row" to "group" for clarity
- **Removed String Cable Segments Header**: Cleaned up BOM section headers
- **DC Feeder Inheritance**: New blocks inherit DC feeder size from previous block
- **Site Preview Refactored**: Moved site preview into its own file for maintainability
- **Configure Device Integration**: Configure device page now pulls from Quick Estimate data and vice versa
- **Inverter Manager**: Updated to handle datasheet info more robustly

### Fixed
- Fixed pricing update bug
- Fixed whip point dragging bug on string homerun configurations
- Fixed device configurator outputting wrong combiner boxes/fuses for string HR configs
- Fixed tracker motor drawing and string point bugs
- Fixed center device between rows bug
- Fixed tracker footprint bugs
- Fixed device placement bug when creating new blocks
- Fixed extender length calculation bugs (multiple fixes)
- Fixed combiner assignment bug for split trackers
- Fixed string reallocation bugs in site preview
- Fixed custom device label bug for inverters
- Fixed module selection warning
- Fixed wire gauge highlighting in device configurator
- Fixed whip/extender/harness calculation bugs
- Fixed DC/AC ratio cap and display bug
- Fixed button squishing bug in Quick Estimate
- Fixed Quick BOM length calculation bugs
- Fixed tracker drawing bug in Quick Estimate
- Fixed 2-string harness display bug
- Various Quick Estimate bug fixes and UI improvements
- Fixed device info panel not showing in distributed string inverter topology

### Enhanced
- Significantly expanded Quick Estimate capabilities from a basic estimator to a full-featured project planning tool
- Improved calculation accuracy with per-device cable sizing and diagnostics verification
- Better multi-module project support with group-level row spacing

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