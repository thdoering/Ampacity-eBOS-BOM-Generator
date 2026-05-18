================================================================================
Solar eBOS BOM Generator v3.5.0
================================================================================

INSTALLATION INSTRUCTIONS
-------------------------

1. Download the executable from SharePoint
   - Click on "Solar eBOS BOM Generator v3.5.0.exe"
   - When you see the warning "isn't commonly downloaded", click the three dots (...)
   - Click "Keep"
   - Confirm again if prompted

2. Run the Application
   - Navigate to your Downloads folder
   - Double-click "Solar eBOS BOM Generator v3.5.0.exe"
   - If Windows SmartScreen appears, click "More info" then "Run anyway"

This is our internal tool - these warnings are normal for unsigned applications.

WHAT'S NEW IN VERSION 3.5.0 (May 15, 2026)
--------------------------------------------

MAJOR NEW FEATURES:
- Factory Module & Inverter Libraries: Read-only factory libraries now ship with
  the app and merge at load time — user entries are never overwritten
- Cable Corridors: Draw polyline cable corridors in Site Preview with axis-snap,
  vertex editing, and device assignment; rendered in PDF exports; parallel-sets
  spinboxes multiply DC feeder and AC homerun footage per set in the BOM
- LV Collection Detail Drawings: PDF wiring diagrams now include full LV collection
  detail layouts
- Circuit Containers: Groups can be organized into named circuits in the group tree
  (visual/organizational only, no BOM impact)
- Device Sequencing Controls: Set the order/sequence of devices in Site Preview
  and Device Configurator
- Device E-W Nudge: Nudge individual devices east-west with arrow keys in
  Site Preview; persisted to project and resettable via Reset Positions

IMPROVEMENTS:
- Module and Inverter search bars for faster library navigation
- Sortable columns in the Site Preview device assignment table
- Unlink All Tracker Templates button for faster group editing
- Enabled template ● indicator on manufacturer/model nodes in Tracker Creator
- Group list replaced with hierarchical tree view supporting circuit containers
- New estimates default to Centralized String topology
- Inverter selection for Quick Estimate moved to the Inverter equipment page
- Row spacing enforcement: Add Group and Calculate Estimate blocked until all
  groups have a spacing value
- Copied groups append to end of list instead of inserting after source
- Strings/CB input cap removed; library max is now default only
- Removed global breaker size and avg DC feeder length inputs
- System Summary Total Strings correctly counts paired half-strings
- Pricing data updated (copper rate and harness library)

BUG FIXES:
- Fixed auto-numbering dropping string assignments
- Fixed overlap detection bug for new block placement in site preview
- Fixed Block Details DC feeder display bug
- Fixed Calculate Estimate always-stale bug (cached result now reused correctly)
- Fixed string coloring in Site Preview
- Fixed strings_per_device override being overwritten on project load
- Fixed rounding bug in estimate calculations

PREVIOUS VERSION HIGHLIGHTS (v3.4.0):
- Block Details Sheet in Quick Estimate Excel export
- String Inverter support in Configure Device and BOM export
- Multi-group selection and drag in layout mode
- Tracker alignment dropdown (top/motor/bottom)
- Assign Devices live preview with Undo All Changes
- Auto-Number Devices button

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
