================================================================================
Solar eBOS BOM Generator v3.4.0
================================================================================

INSTALLATION INSTRUCTIONS
-------------------------

1. Download the executable from SharePoint
   - Click on "Solar eBOS BOM Generator v3.4.0.exe"
   - When you see the warning "isn't commonly downloaded", click the three dots (...)
   - Click "Keep"
   - Confirm again if prompted

2. Run the Application
   - Navigate to your Downloads folder
   - Double-click "Solar eBOS BOM Generator v3.4.0.exe"
   - If Windows SmartScreen appears, click "More info" then "Run anyway"

This is our internal tool - these warnings are normal for unsigned applications.

WHAT'S NEW IN VERSION 3.4.0 (April 28, 2026)
----------------------------------------------

MAJOR NEW FEATURES:
- Block Details Sheet: New sheet in Quick Estimate Excel export with per-device
  part breakdowns (extenders, harnesses, fuses, whips, DC feeders/AC homeruns)
  and a summary table aggregating line items by part number
- String Inverter Support: Configure Device tab and BOM export now include a
  dedicated String Inverter sheet with MPPT Max Current and Max AC Output Current
- Multi-Group Selection & Drag: Ctrl+click to toggle group selection, rubber-band
  box select, and multi-group drag in Quick Estimate layout mode
- Tracker Alignment: New top/motor/bottom alignment dropdown in Quick Estimate
  group editor for mixed-length tracker groups
- Assign Devices Live Preview: Assignment changes commit to canvas immediately;
  single "Undo All Changes" button replaces Apply/Cancel
- Auto-Number Devices: New button in Edit Devices for top-left → bottom-right
  device numbering with configurable prefix and start number

IMPROVEMENTS:
- Auto-Calculate DC Feeder Size: Block configurator derives cable size from breaker
  rating automatically, with manual-set tracking and "Reset to recommended"
- AL/CU material indicators on all cable descriptions in both BOM exports
- Description column added to Quick Estimate BOM (results tree and Excel export)
- Collapsible sections in block configurator
- Wattage Spinbox in Module Manager (100–1000 W, 5 W steps; warns on Vmp×Imp mismatch)
- "Apply to All" toggle for DC Fdr and AC HR wire sizing in Quick Estimate
- Auto-enable newly created/duplicated tracker templates
- Multi-select delete for tracker templates with in-use guard

BUG FIXES:
- Fixed Copy Project bug where the original file was deleted instead of duplicated
- Fixed duplicate Site Preview windows stacking on repeat clicks
- Fixed Assign Devices dialog buttons clipped on small projects
- Fixed performance regression with many groups/trackers in Quick Estimate
- Fixed central inverter string and combiner box count bug
- Fixed stale wire sizing rows not clearing when LV collection method changes

PREVIOUS VERSION HIGHLIGHTS (v3.3.0):
- Canvas Drag-to-Device string allocation
- Polyline measuring tool with save/load
- 2P Tracker and Parallel DC Feeder support
- Azimuth rotation for tracker groups
- SAT-based overlap detection for rotated/sheared groups

PREVIOUS VERSION HIGHLIGHTS (v3.2.0):
- Driveline angle range expanded to -45 to 45 degrees
- Import/Export now transfers module, inverter, and tracker template libraries
- Fixed sorting bug in Configure Device tab

PREVIOUS VERSION HIGHLIGHTS (v3.1.0):
- PDF Export of Site Preview with professional titleblock
- Background labels in site preview
- Updated default values

PREVIOUS VERSION HIGHLIGHTS (v3.0.0):
- Quick Estimate Overhaul with site preview, tracker templates, inverter pads
- Central Inverter Topology support
- Diagnostics file for calculation verification
- Per-device cable sizing (DC feeder, AC homerun)
- Row spacing/GCR moved to group level
- 20+ bug fixes


SUPPORT
-------
For questions or issues, contact Tyler Doering (tdoering@ampacity.com).

Previous versions are archived in the "Archive" folder.


================================================================================
