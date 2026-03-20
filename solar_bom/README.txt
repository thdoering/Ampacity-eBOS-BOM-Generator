================================================================================
Solar eBOS BOM Generator v3.0.0
================================================================================

INSTALLATION INSTRUCTIONS
-------------------------

1. Download the executable from SharePoint
   - Click on "Solar eBOS BOM Generator v3.0.0.exe"
   - When you see the warning "isn't commonly downloaded", click the three dots (...)
   - Click "Keep"
   - Confirm again if prompted

2. Run the Application
   - Navigate to your Downloads folder
   - Double-click "Solar eBOS BOM Generator v3.0.0.exe"
   - If Windows SmartScreen appears, click "More info" then "Run anyway"

This is our internal tool - these warnings are normal for unsigned applications.

WHAT'S NEW IN VERSION 3.0.0 (March 20, 2026)
----------------------------------------------

MAJOR NEW FEATURES:
- Quick Estimate Overhaul: Completely rebuilt with site preview, tracker
  template integration, driveline angle, half-string trackers, inverter pads,
  manual/auto string allocation, per-device cable inputs, and copy estimates
- Site Preview: Interactive preview with panning, device selection, info
  pop-ups, and motor alignment controls
- Central Inverter Topology: Full support in Quick Estimate with dedicated BOM
- Diagnostics File: Verify calculations across different project configurations

NEW FEATURES:
- Segment Rounding Dropdown: Consolidate BOM line items by rounding segments
- Combiner Boxes Tab in Quick Estimate Excel BOM
- LV Collection Inputs: String HR, harness, and trunk bus inputs
- Inverter Allocation in Quick BOM with DC/AC ratio display
- Per-Device Strings, DC Feeder, and AC Homerun inputs
- Copper Rate and Module Polarity Orientation on BOM
- Bulk Copy DC Feeder Lengths across blocks
- Collapse/Expand All in edit devices dialog

KEY CHANGES:
- Row Spacing and GCR moved to group level (supports multiple modules/spacings)
- Updated string allocation algorithm for balanced sites
- Configure Device page now integrates with Quick Estimate data
- New blocks inherit DC feeder size from previous block

BUG FIXES:
- Fixed 20+ bugs across Quick Estimate, device configurator, site preview,
  extender calculations, string reallocation, and BOM generation
- See CHANGELOG.md for full details

PREVIOUS VERSION HIGHLIGHTS (v2.7.0):
- Quick Estimate Tab for fast project estimates
- DC Feeder Lengths input per block
- Revision number tracking
- Module orientation input

WHAT'S NEW IN VERSION 2.7.0 (February 25, 2026)
----------------------------------------------

MAJOR NEW FEATURES:
- Quick Estimate Tab: Fast project estimates with dedicated Excel export
- DC Feeder Lengths: Input specific feeder cable lengths for each block

NEW FEATURES:
- Revision Number: Track BOM revisions with version numbering
- Module Orientation: Specify module orientation in project configuration

IMPROVEMENTS:
- Enhanced Excel BOM formatting for better clarity
- Streamlined BOM by removing redundant category column
- Various bug fixes and stability improvements

PREVIOUS VERSION HIGHLIGHTS (v2.6.0):
- Tracker multi-select and move functionality
- 2P tracker support (WIP)
- Batch apply recommended wire sizes
- Default recommended cable sizes for all wires


SUPPORT
-------
For questions or issues, contact Tyler Doering (tdoering@ampacity.com).

Previous versions are archived in the "Archive" folder.


================================================================================