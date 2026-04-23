================================================================================
Solar eBOS BOM Generator v3.3.0
================================================================================

INSTALLATION INSTRUCTIONS
-------------------------

1. Download the executable from SharePoint
   - Click on "Solar eBOS BOM Generator v3.3.0.exe"
   - When you see the warning "isn't commonly downloaded", click the three dots (...)
   - Click "Keep"
   - Confirm again if prompted

2. Run the Application
   - Navigate to your Downloads folder
   - Double-click "Solar eBOS BOM Generator v3.3.0.exe"
   - If Windows SmartScreen appears, click "More info" then "Run anyway"

This is our internal tool - these warnings are normal for unsigned applications.

WHAT'S NEW IN VERSION 3.3.0 (April 23, 2026)
----------------------------------------------

MAJOR NEW FEATURES:
- Canvas Drag-to-Device: Select strings on the site preview canvas and drag them
  to devices for fast manual string allocation
- Measuring Tool: Polyline measuring tool with per-segment and cumulative
  distances; measurements save and reload with the project
- 2P Tracker Support: Full two-portrait tracker configuration now complete
- Parallel DC Feeders: Support for parallel DC feeder configurations
- Azimuth Rotation: Tracker groups now support azimuth rotation input

IMPROVEMENTS:
- PDF summary table auto-selects least-obstructed corner; placement dialog
  no longer appears before every export
- Inverter library now shows as manufacturer/model tree (collapsed by default)
- Site preview toolbar split into two rows for better visibility on smaller screens
- Manual string allocations preserved when tracker structure changes; new
  trackers default to unallocated

BUG FIXES:
- Fixed silent no-op when moving strings on canvas (Unallocated device index shift)
- Fixed overlap detection for rotated/sheared tracker groups (now SAT-based)
- Fixed combiner box placement off-edge for mixed-size tracker groups
- Fixed 1-string harness config on split trackers
- Fixed inverter import/export bug
- Fixed module width > length validation

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
