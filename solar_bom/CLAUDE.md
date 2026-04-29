# Solar eBOS BOM Generator

## Project Overview

Solar eBOS (electrical Balance of System) BOM (Bill of Materials) Generator written in Python with a Tkinter UI. Lets the user lay out solar project electrical designs — modules, trackers, blocks, inverters, combiner boxes, wire harnesses, DC feeders — and generate a bill of materials for the balance-of-system components. Also supports a "Quick Estimate" workflow for early-stage sizing before a full block layout exists.

Current version: 3.2.0 (see `CHANGELOG.md`).

## Directory Structure

```
solar_bom/
├── main.py                    # Application entry point
├── build_app.py               # PyInstaller build entry
├── version.py                 # Version constants
├── solar_bom.spec             # PyInstaller spec
├── CHANGELOG.md
├── src/
│   ├── models/                # Core data classes
│   ├── ui/                    # Tkinter UI
│   └── utils/                 # Parsing, calculations, generators
├── data/                      # JSON libraries (ignored by git except committed seeds)
├── projects/                  # Saved projects (gitignored)
└── harness_drawings/          # Generated PNG harness drawings (gitignored)
```

## Models (`src/models/`)

- `module.py` — `ModuleSpec`, `ModuleType`, `ModuleOrientation`
- `inverter.py` — `InverterSpec`, `MPPTChannel`, `MPPTConfig`
- `tracker.py` — `TrackerTemplate`, `TrackerPosition`
- `block.py` — `BlockConfig`, `WiringType`, `DeviceType`, `HarnessGroup`, `WiringConfig`
- `device.py` — `HarnessConnection`, `CombinerBoxConfig` (combiner box / device-level config)
- `project.py` — `Project`, `ProjectMetadata`

## UI (`src/ui/`)

- `project_dashboard.py` — landing screen, recent + all projects
- `module_manager.py` — module library management
- `inverter_manager.py` — inverter library management
- `tracker_creator.py` — tracker template builder
- `block_configurator.py` — block-level layout editor
- `wiring_configurator.py` — wiring / eBOS layout for a block
- `device_configurator.py` — combiner box / device configuration
- `harness_designer.py` — harness template designer
- `harness_catalog_dialog.py` — harness drawing generation dialog
- `dc_feeder_dialog.py` — DC feeder sizing per block
- `quick_estimate.py` — early-stage sizing workflow (pre-block layout)
- `site_preview.py` — site preview window used by Quick Estimate
- `bom_manager.py` — BOM review and Excel export

## Utilities (`src/utils/`)

- `pan_parser.py` — parses `.pan` PV module spec files
- `file_handlers.py` — file I/O helpers (JSON load/save, path utilities)
- `calculations.py` — shared engineering calculations, fuse size tables
- `cable_sizing.py` — NEC ampacity lookups and cable size recommendations
- `project_manager.py` — project save/load, recent projects
- `undo_manager.py` — undo/redo support
- `bom_generator.py` — BOM assembly and Excel export
- `pricing_lookup.py` — component pricing / copper tier lookups
- `string_allocation.py` — algorithms for allocating strings across inverters
- `harness_drawing_generator.py` — PIL-based technical drawings for harnesses
- `site_pdf_generator.py` — matplotlib-based 11x17 site layout PDFs

## Data (`data/`)

Runtime JSON libraries. User-specific ones are gitignored; committed ones are the part/price reference data.

- `module_library_factory.json` — **read-only factory module library**, committed to git and shipped in the bundle. Never written by the running app. Hierarchical format identical to `module_templates.json`. At load time the two are merged (factory wins on conflict); user entries that collide with a factory key are silently shadowed but left intact in the user file.
- `module_templates.json`, `tracker_templates.json`, `inverter_templates.json` — user-saved templates
- `harness_library.json` — harness part catalog
- `fuse_library.json` — inline harness fuses
- `combiner_box_library.json` — combiner box catalog
- `combiner_box_fuse_library.json` — combiner box fuses
- `extender_library.json` — extender cable parts
- `whip_library.json` — whip cable parts
- `pricing_data.json` — pricing with copper-tier settings
- (and any NEC ampacity table JSON referenced by `cable_sizing.py`)

## Working Rules

Please follow these when making changes.

### Ask before acting on incomplete context
If a change depends on context I haven't given you — a design decision, how a method is used elsewhere, which of two interpretations I mean — ask before coding. I'd rather answer a question than undo a wrong assumption. Prioritize asking questions over making changes with incomplete information.

### Read the actual current file before editing
Always open the file and check its current state before proposing or making edits. Don't rely on what you remember from earlier in the session, what the file "probably" looks like, or what was true last time. The file on disk is the source of truth.

### Small, focused changes, one step at a time
I work best when changes land in small, testable pieces. Prefer making one change, letting me test it, and then moving to the next over stacking many edits into one round. If a feature naturally needs multiple changes, break it into a sequence and pause between steps so I can verify each one.

### Name the file and method when describing changes
When you explain what you're about to change (or just changed), reference the file and the specific method or class. Many methods in this project share names across different files (for example, `update_layout` or `save`), so naming the change without the file is ambiguous. Say the file every time, even when it seems obvious.

### Don't expand scope
If you notice something adjacent that looks wrong or suboptimal, mention it and let me decide — don't fix it as part of an unrelated change. Stay inside the change I asked for.

### Check for references before removing code
Before deleting a method, attribute, class, or widget, search the project for references to it. Tkinter event bindings, callbacks, string-based tag lookups, and dynamically-resolved imports make it easy to break things silently.

### Preserve existing patterns
When editing a file, match the style and widget patterns already in that file rather than introducing new ones. If you think a different pattern would be better, say so as a suggestion first — don't just switch styles inside an unrelated edit.

### Respect the two data paths
The app has two parallel paths to BOM-relevant data: the full block layout (`block_configurator` → `wiring_configurator` → `device_configurator` → `bom_manager`) and the Quick Estimate path (`quick_estimate` → `site_preview` → `site_pdf_generator`). Device Configurator can pull from either. Before changing shared downstream code, check whether both paths feed it.

## Separation of Concerns

- UI files shouldn't contain engineering calculations — those belong in `calculations.py`, `cable_sizing.py`, `string_allocation.py`, or the relevant model.
- Model files shouldn't build Tkinter widgets.
- Library lookups (pricing, NEC tables, part catalogs) belong in their dedicated utility modules, not inlined in UI or BOM generation code.
- If a change would cross any of these lines, flag it before doing it.

## Versioning

`version.py` is the single source of truth for the app version. `solar_bom.spec` and the main window title read from it. Bump it intentionally, not as a side effect of other changes.