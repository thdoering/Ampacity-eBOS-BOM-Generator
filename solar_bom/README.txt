================================================================================
Solar eBOS BOM Generator v3.6.0
================================================================================

INSTALLATION INSTRUCTIONS
-------------------------

1. Download the executable from SharePoint
   - Click on "Solar eBOS BOM Generator v3.6.0.exe"
   - When you see the warning "isn't commonly downloaded", click the three dots (...)
   - Click "Keep"
   - Confirm again if prompted

2. Run the Application
   - Navigate to your Downloads folder
   - Double-click "Solar eBOS BOM Generator v3.6.0.exe"
   - If Windows SmartScreen appears, click "More info" then "Run anyway"

This is our internal tool - these warnings are normal for unsigned applications.

WHAT'S NEW IN VERSION 3.6.0 (June 23, 2026)
--------------------------------------------

MAJOR NEW FEATURES:
- NEC 2023 Ampacity Engine: Wire sizing in Quick Estimate (and the block/wiring
  configurator workflow) now runs through a NEC 2023-based ampacity engine that
  selects the minimum gauge satisfying both ampacity (ambient + CCC derating,
  termination cap, 690.8 source-circuit factor) and a configurable voltage-drop
  target, across all five cable types
- Wire Sizing Settings Panel: New collapsible panel for per-cable control over
  insulation type, termination temp, install method, conductor material, VD%
  target, and circuits-sharing count; table now shows a Sizing Detail column
  with a per-row NEC breakdown, yellow highlighting for manual overrides, and a
  (VD↑) marker when voltage drop drove the gauge up
- Skids Field: Decouples AC-homerun quantity from inverter count (Central Inverter
  and Centralized String topologies); export is gated until placed pads match the
  skid target, with a live "N of M skids placed" hint in Site Preview
- Per-Estimate Revision Field: Quick Estimate gained a Revision field; PDF and Excel
  filenames and the titleblock REV cell now carry version + revision + date

IMPROVEMENTS:
- 32' FS harnesses added to the harness library
- Harness slack is now visible in the wiring workflow
- User-editable string-to-string length for harnesses
- Factory module and inverter libraries are clearly labeled (factory) with Edit/Delete
  disabled and an alert explaining factory inverters can't be modified
- AC homerun default length changed from 500 ft to 50 ft
- Combiner box east-west nudge step is now half the row spacing
- "Apply to all" button added for row spacing
- Calculate Estimate is blocked with a warning when no inverter is assigned
- Wiring geometry now references tracker physical center and the true driveline
- Pricing data updated

BUG FIXES:
- Extender lengths are now independent of combiner box north-south position
- Fixed split-tracker overlay harness bleeding onto strings of another combiner box
- Fixed the Block Details tab in the Excel BOM
- Fixed device alignment bugs
- Fixed a combiner box placement bug
- Fixed the Calculate button falsely turning red again after a successful run

PREVIOUS VERSION HIGHLIGHTS (v3.5.0):
- Factory Module & Inverter Libraries that ship with the app and merge at load time
- Cable Corridors in Site Preview with PDF rendering and parallel-sets footage
- LV Collection Detail drawings in PDF wiring diagrams
- Circuit Containers for organizing groups in the group tree
- Device Sequencing controls and Device E-W nudge in Site Preview
- Module and Inverter search bars

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


SUPPORT
-------
For questions or issues, contact Tyler Doering (tdoering@ampacity.com).

Previous versions are archived in the "Archive" folder.


================================================================================
