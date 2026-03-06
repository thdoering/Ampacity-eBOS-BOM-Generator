import tkinter as tk
from tkinter import ttk
from typing import Optional, Dict, List, Any
from pathlib import Path
import json
import uuid
from datetime import datetime
from src.utils.string_allocation import allocate_strings


class QuickEstimate(ttk.Frame):
    """Quick estimation tool for early-stage project sizing with hierarchical structure"""
    
    def __init__(self, parent, current_project=None, on_save=None):
        super().__init__(parent)
        self.current_project = current_project
        self.estimate_id = None
        self.on_save = on_save
        self.pricing_data = self.load_pricing_data()
        self._estimate_id_map = {}  # Map display name to ID
        self.available_modules = self.load_module_library()  # {display_name: ModuleSpec}
        self.selected_module = None  # Currently selected ModuleSpec
        self.available_inverters = self.load_inverter_library()  # {display_name: InverterSpec}
        self.selected_inverter = None  # Currently selected InverterSpec
        
        # Data structure for subarrays and blocks
        self.subarrays: Dict[str, Dict[str, Any]] = {}
        
        # Global settings defaults
        self.module_width_default = 1134
        self.modules_per_string_default = 28
        self.row_spacing_default = 20.0
        self.wire_gauge_default = '10 AWG'
        
        # Track currently selected item
        self.selected_item_id = None
        self.selected_item_type = None  # 'subarray' or 'block'
        self.checked_items = set()  # Items checked for export
        self._results_stale = True
        self._calc_btn = None  # Reference to calculate button
        self._autosave_after_id = None
        
        self.setup_ui()
        
        # Load most recent estimate or show empty state
        self._refresh_estimate_dropdown(auto_select=True)

    def load_pricing_data(self):
        """Load pricing data from JSON file"""
        try:
            pricing_path = Path('data/pricing_data.json')
            if pricing_path.exists():
                with open(pricing_path, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading pricing data: {e}")
        return {}

    def load_module_library(self):
        """Load modules from the module templates JSON file"""
        modules = {}
        try:
            module_path = Path('data/module_templates.json')
            if module_path.exists():
                with open(module_path, 'r') as f:
                    data = json.load(f)
                
                from ..models.module import ModuleSpec, ModuleType, ModuleOrientation
                
                for manufacturer, models in data.items():
                    if not isinstance(models, dict):
                        continue
                    for model_name, spec_data in models.items():
                        try:
                            # Handle enum conversions
                            mod_type = ModuleType(spec_data.get('type', 'Mono PERC'))
                            orientation = ModuleOrientation(spec_data.get('default_orientation', 'Portrait'))
                            
                            module = ModuleSpec(
                                manufacturer=spec_data.get('manufacturer', manufacturer),
                                model=spec_data.get('model', model_name),
                                type=mod_type,
                                length_mm=float(spec_data.get('length_mm', 0)),
                                width_mm=float(spec_data.get('width_mm', 0)),
                                depth_mm=float(spec_data.get('depth_mm', 40)),
                                weight_kg=float(spec_data.get('weight_kg', 25)),
                                wattage=float(spec_data.get('wattage', 0)),
                                vmp=float(spec_data.get('vmp', 0)),
                                imp=float(spec_data.get('imp', 0)),
                                voc=float(spec_data.get('voc', 0)),
                                isc=float(spec_data.get('isc', 0)),
                                max_system_voltage=float(spec_data.get('max_system_voltage', 1500)),
                                temperature_coefficient_pmax=spec_data.get('temperature_coefficient_pmax'),
                                temperature_coefficient_voc=spec_data.get('temperature_coefficient_voc'),
                                temperature_coefficient_isc=spec_data.get('temperature_coefficient_isc'),
                                default_orientation=orientation,
                                cells_per_module=int(spec_data.get('cells_per_module', 72))
                            )
                            display_name = f"{module.manufacturer} {module.model} ({module.wattage}W)"
                            modules[display_name] = module
                        except (ValueError, TypeError) as e:
                            print(f"Error loading module {manufacturer}/{model_name}: {e}")
        except Exception as e:
            print(f"Error loading module library: {e}")
        return modules
    
    def load_inverter_library(self):
        """Load inverters from the inverters JSON file"""
        inverters = {}
        try:
            inverter_path = Path('data/inverters.json')
            if inverter_path.exists():
                with open(inverter_path, 'r') as f:
                    data = json.load(f)
                
                from ..models.inverter import InverterSpec, MPPTChannel, MPPTConfig, InverterType
                
                for name, specs in data.items():
                    try:
                        rated_power = specs.get('rated_power_kw', specs.get('rated_power', 10.0))
                        max_dc_power = specs.get('max_dc_power_kw', float(rated_power) * 1.5)
                        inverter_type_str = specs.get('inverter_type', 'String')
                        
                        channels = []
                        for ch in specs.get('mppt_channels', []):
                            channels.append(MPPTChannel(**ch))
                        
                        inverter = InverterSpec(
                            manufacturer=specs.get('manufacturer', 'Unknown'),
                            model=specs.get('model', 'Unknown'),
                            inverter_type=InverterType(inverter_type_str),
                            rated_power_kw=float(rated_power),
                            max_dc_power_kw=float(max_dc_power),
                            max_efficiency=float(specs.get('max_efficiency', 98.0)),
                            mppt_channels=channels,
                            mppt_configuration=MPPTConfig(specs.get('mppt_configuration', 'Independent')),
                            max_dc_voltage=float(specs.get('max_dc_voltage', 1500)),
                            startup_voltage=float(specs.get('startup_voltage', 150)),
                            nominal_ac_voltage=float(specs.get('nominal_ac_voltage', 400.0)),
                            max_ac_current=float(specs.get('max_ac_current', 40.0)),
                            power_factor=float(specs.get('power_factor', 0.99)),
                            dimensions_mm=tuple(specs.get('dimensions_mm', (1000, 600, 300))),
                            weight_kg=float(specs.get('weight_kg', 75.0)),
                            ip_rating=specs.get('ip_rating', 'IP65'),
                            max_short_circuit_current=specs.get('max_short_circuit_current')
                        )
                        display_name = f"{inverter.manufacturer} {inverter.model}"
                        inverters[display_name] = inverter
                    except Exception as e:
                        print(f"Warning: Failed to load inverter '{name}': {e}")
        except Exception as e:
            print(f"Error loading inverter library: {e}")
        return inverters

    def generate_id(self, prefix: str) -> str:
        """Generate a unique ID with a prefix"""
        return f"{prefix}_{uuid.uuid4().hex[:8]}"
    
    def get_next_block_name(self, subarray_id: str, source_name: str) -> str:
        """Generate the next logical block name based on the source name pattern"""
        import re
        
        # Try to find a trailing number (with optional leading zeros)
        match = re.match(r'^(.*?)(\d+)$', source_name)
        
        if not match:
            # No number found, just append (Copy)
            return f"{source_name} (Copy)"
        
        prefix = match.group(1)
        number_str = match.group(2)
        number = int(number_str)
        
        # Determine the padding (e.g., "01" is width 2, "001" is width 3)
        padding = len(number_str)
        
        # Get existing block names in this subarray
        existing_names = set()
        if subarray_id in self.subarrays:
            for block_data in self.subarrays[subarray_id]['blocks'].values():
                existing_names.add(block_data['name'])
        
        # Find the next available number
        next_number = number + 1
        while True:
            new_name = f"{prefix}{str(next_number).zfill(padding)}"
            if new_name not in existing_names:
                return new_name
            next_number += 1
            # Safety limit to prevent infinite loop
            if next_number > number + 1000:
                return f"{source_name} (Copy)"
            
    def disable_combobox_scroll(self, combobox):
        """Prevent combobox from responding to mouse wheel"""
        def _ignore_scroll(event):
            return "break"
        
        combobox.bind("<MouseWheel>", _ignore_scroll)
        # For Linux compatibility
        combobox.bind("<Button-4>", _ignore_scroll)
        combobox.bind("<Button-5>", _ignore_scroll)

    def update_string_count(self):
        """Update the string count label for the current block"""
        if not hasattr(self, 'string_count_label') or not hasattr(self, 'current_block'):
            return
        
        total_strings = 0
        total_trackers = 0
        
        for tracker in self.current_block.get('trackers', []):
            qty = tracker.get('quantity', 0)
            strings = tracker.get('strings', 0)
            total_trackers += qty
            total_strings += qty * strings
        
        self.string_count_label.config(text=f"Total Strings: {total_strings:,}  |  Total Trackers: {total_trackers:,}")

    # ==================== Data Management ====================
    
    def add_subarray(self, auto_add_block: bool = False) -> str:
        """Add a new subarray and return its ID"""
        subarray_id = self.generate_id("subarray")
        subarray_num = len(self.subarrays) + 1
        
        self.subarrays[subarray_id] = {
            'name': f"Subarray {subarray_num}",
            'transformer_mva': 4.0,
            'blocks': {}
        }
        
        # Add to tree view
        self.tree.insert('', 'end', subarray_id, text=f"Subarray {subarray_num}", open=True)
        
        # Optionally add a default block
        if auto_add_block:
            self.add_block(subarray_id)
        
        # Select the new subarray
        self.tree.selection_set(subarray_id)
        self.on_tree_select(None)
        
        return subarray_id

    def add_block(self, subarray_id: str) -> str:
        """Add a new block to a subarray and return its ID"""
        if subarray_id not in self.subarrays:
            return None
        
        block_id = self.generate_id("block")
        block_num = len(self.subarrays[subarray_id]['blocks']) + 1
        
        self.subarrays[subarray_id]['blocks'][block_id] = {
            'name': f"Block {block_num}",
            'type': 'combiner',  # 'combiner' or 'string_inverter'
            'num_combiners': 1,
            'breaker_size': 400,
            'dc_feeder_distance_ft': 0.0,
            'dc_feeder_cable_size': '4/0 AWG',
            'trackers': [
                {'strings': 2, 'quantity': 1, 'harness_config': '2'}
            ]
        }
        
        # Add to tree view under the subarray
        self.tree.insert(subarray_id, 'end', block_id, text=f"Block {block_num}")
        
        # Expand the parent subarray
        self.tree.item(subarray_id, open=True)
        
        # Select the new block
        self.tree.selection_set(block_id)
        self.on_tree_select(None)
        
        return block_id

    def delete_selected_item(self):
        """Delete the currently selected subarray or block"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item_id = selection[0]
        parent_id = self.tree.parent(item_id)
        
        if parent_id == '':
            # It's a subarray
            if item_id in self.subarrays:
                del self.subarrays[item_id]
                self.tree.delete(item_id)
        else:
            # It's a block
            if parent_id in self.subarrays:
                if item_id in self.subarrays[parent_id]['blocks']:
                    del self.subarrays[parent_id]['blocks'][item_id]
                    self.tree.delete(item_id)
        
        # Clear the details panel
        self.clear_details_panel()

    def get_subarray_for_block(self, block_id: str) -> Optional[str]:
        """Find which subarray a block belongs to"""
        for subarray_id, subarray_data in self.subarrays.items():
            if block_id in subarray_data['blocks']:
                return subarray_id
        return None

    def round_whip_length(self, raw_length_ft):
        """Apply 5% waste factor and round up to nearest 5ft increment (min 10ft)"""
        WASTE_FACTOR = 1.05
        length_with_waste = raw_length_ft * WASTE_FACTOR
        rounded = 5 * ((length_with_waste + 5 - 0.1) // 5 + 1)
        return max(10, int(rounded))
    
    def load_combiner_library(self):
        """Load combiner box library from JSON file"""
        try:
            lib_path = Path('data/combiner_box_library.json')
            if lib_path.exists():
                with open(lib_path, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading combiner box library: {e}")
        return {}

    def get_fuse_holder_category(self, fuse_current_amps):
        """Determine fuse holder rating category from required fuse current"""
        if fuse_current_amps <= 20:
            return "20A and Below"
        elif fuse_current_amps <= 32:
            return "25-32A"
        else:
            return "32A and Above"

    def find_combiner_box(self, strings_per_cb, breaker_size, fuse_holder_rating):
        """Find a matching combiner box from the library.
        
        Returns the part_number and description, or None if no match.
        Prefers smallest max_inputs that fits, with whips (whip_length > 0).
        """
        combiner_library = self.load_combiner_library()
        
        candidates = []
        for part_num, cb_data in combiner_library.items():
            if (cb_data.get('max_inputs', 0) >= strings_per_cb and
                cb_data.get('breaker_size', 0) == breaker_size and
                cb_data.get('fuse_holder_rating', '') == fuse_holder_rating):
                candidates.append(cb_data)
        
        if not candidates:
            return None
        
        # Sort by max_inputs (prefer smallest that fits), then prefer with whips
        candidates.sort(key=lambda c: (c['max_inputs'], -c.get('whip_length_ft', 0)))
        return candidates[0]

    def calculate_cb_whip_distances(self, total_rows, num_combiners, row_spacing_ft):
        """Calculate whip distances for each row based on evenly-spaced CB placement.
        
        Returns a list of (distance_ft, cb_index) for each row.
        
        CB placement logic:
        - Rows are divided into groups, first CB(s) get extra rows if uneven
        - Each CB is placed at the center of its group
        - Each row connects to its nearest CB
        """
        if total_rows <= 0 or num_combiners <= 0:
            return []
        
        # Divide rows into groups for each CB
        base_group_size = total_rows // num_combiners
        extra_rows = total_rows % num_combiners
        
        # Build groups - first CBs get the extra rows
        groups = []
        row_start = 1
        for cb_idx in range(num_combiners):
            group_size = base_group_size + (1 if cb_idx < extra_rows else 0)
            row_end = row_start + group_size - 1
            # CB position is center of its group
            cb_position = (row_start + row_end) / 2.0
            groups.append({
                'cb_idx': cb_idx,
                'row_start': row_start,
                'row_end': row_end,
                'cb_position': cb_position
            })
            row_start = row_end + 1
        
        # Calculate distance from each row to its assigned CB
        whip_distances = []
        for row_num in range(1, total_rows + 1):
            # Find which group this row belongs to
            for group in groups:
                if group['row_start'] <= row_num <= group['row_end']:
                    distance_ft = abs(row_num - group['cb_position']) * row_spacing_ft
                    whip_distances.append((distance_ft, group['cb_idx']))
                    break
        
        return whip_distances

    def lookup_part_and_price(self, item_type, **kwargs):
        """Look up part number and unit price for a BOM item.
        item_type: 'harness', 'whip', 'extender'
        Returns (part_number, unit_price_str, ext_price_str)
        """
        try:
            import os
            import json
            from src.utils.pricing_lookup import PricingLookup

            wire_gauge = getattr(self, 'wire_gauge_var', None)
            wire_gauge = wire_gauge.get() if wire_gauge else '10 AWG'

            pricing = PricingLookup()
            part_number = 'N/A'

            if item_type == 'harness':
                num_strings = kwargs.get('num_strings', 1)
                polarity = kwargs.get('polarity', 'positive')
                qty = kwargs.get('qty', 1)

                # Calculate string spacing from module width and modules per string
                try:
                    mps = int(self.modules_per_string_var.get())
                except ValueError:
                    mps = 28
                module_width_mm = self.selected_module.width_mm if self.selected_module else 1134
                # Use 0.02m (20mm) as default module gap spacing
                string_spacing_ft = (mps * module_width_mm + (mps - 1) * 20) / 1000 * 3.28084

                # Load harness library
                current_dir = os.path.dirname(os.path.abspath(__file__))
                project_root = os.path.dirname(os.path.dirname(current_dir))
                lib_path = os.path.join(project_root, 'data', 'harness_library.json')
                with open(lib_path, 'r') as f:
                    harness_library = json.load(f)

                available_spacings = [26.0, 102.0, 113.0, 122.0, 133.0]
                target_spacing = max(available_spacings)
                for sp in sorted(available_spacings):
                    if sp >= string_spacing_ft:
                        target_spacing = sp
                        break

                matches = []
                for pn, spec in harness_library.items():
                    if pn.startswith('_comment_'):
                        continue
                    if (spec.get('num_strings') == num_strings and
                            spec.get('polarity') == polarity and
                            abs(spec.get('string_spacing_ft', 0) - target_spacing) < 0.1):
                        spec_trunk = spec.get('trunk_cable_size', spec.get('trunk_wire_gauge', ''))
                        if spec_trunk == wire_gauge:
                            matches.append(pn)
                part_number = matches[0] if len(matches) == 1 else ('N/A' if not matches else ' or '.join(sorted(matches)))

            elif item_type in ('whip', 'extender'):
                polarity = kwargs.get('polarity', 'positive')
                length_ft = kwargs.get('length_ft', 10)
                qty = kwargs.get('qty', 1)

                target_length = ((length_ft - 1) // 5 + 1) * 5
                target_length = max(10, target_length)

                lib_name = 'whip_library.json' if item_type == 'whip' else 'extender_library.json'
                current_dir = os.path.dirname(os.path.abspath(__file__))
                project_root = os.path.dirname(os.path.dirname(current_dir))
                lib_path = os.path.join(project_root, 'data', lib_name)
                with open(lib_path, 'r') as f:
                    library = json.load(f)

                for pn, spec in library.items():
                    if pn.startswith('_comment_'):
                        continue
                    if (spec.get('wire_gauge') == wire_gauge and
                            spec.get('polarity') == polarity and
                            spec.get('length_ft') == target_length):
                        part_number = pn
                        break
            else:
                qty = kwargs.get('qty', 1)

            # Look up price
            unit_price = pricing.get_price(part_number) if part_number != 'N/A' else None
            if unit_price is not None:
                qty_val = kwargs.get('qty', 1)
                return part_number, f"${unit_price:,.2f}", f"${unit_price * qty_val:,.2f}"
            else:
                return part_number, '', ''

        except Exception as e:
            print(f"lookup_part_and_price error ({item_type}): {e}")
            return 'N/A', '', ''

    def _on_results_tree_click(self, event):
        """Handle click on results tree - toggle checkbox if include column clicked"""
        item = self.results_tree.identify_row(event.y)
        column = self.results_tree.identify_column(event.x)
        if item and column == '#1':  # include column
            self._toggle_result_item(item)

    def _toggle_result_item(self, item):
        """Toggle the checked state of a results tree item"""
        values = list(self.results_tree.item(item, 'values'))
        # Don't allow toggling section headers
        if str(values[1]).startswith('---'):
            return
        if values[0] == '☐':
            values[0] = '☑'
            self.checked_items.add(item)
            self.results_tree.item(item, values=values, tags=('checked',))
        else:
            values[0] = '☐'
            self.checked_items.discard(item)
            self.results_tree.item(item, values=values, tags=('unchecked',))

    def _results_select_all(self):
        """Check all non-section rows in results tree"""
        for item in self.results_tree.get_children():
            values = list(self.results_tree.item(item, 'values'))
            if not str(values[1]).startswith('---'):
                values[0] = '☑'
                self.checked_items.add(item)
                self.results_tree.item(item, values=values, tags=('checked',))

    def _results_select_none(self):
        """Uncheck all rows in results tree"""
        for item in self.results_tree.get_children():
            values = list(self.results_tree.item(item, 'values'))
            if not str(values[1]).startswith('---'):
                values[0] = '☐'
                self.checked_items.discard(item)
                self.results_tree.item(item, values=values, tags=('unchecked',))
        self.checked_items.clear()

    def get_default_combiner_price(self):
        """Get default combiner box price from pricing data"""
        try:
            combiner_data = self.pricing_data.get('combiner_boxes', {})
            for key, value in combiner_data.items():
                if isinstance(value, dict) and 'price' in value:
                    return float(value['price'])
                elif isinstance(value, (int, float)):
                    return float(value)
            return 0.0
        except:
            return 0.0

    def get_harness_options(self, num_strings):
        """Get valid harness configuration options for a given number of strings"""
        if num_strings == 1:
            return ["1"]
        elif num_strings == 2:
            return ["2", "1+1"]
        elif num_strings == 3:
            return ["3", "2+1", "1+1+1"]
        elif num_strings == 4:
            return ["4", "3+1", "2+2", "2+1+1", "1+1+1+1"]
        elif num_strings == 5:
            return ["5", "4+1", "3+2"]
        elif num_strings == 6:
            return ["6", "5+1", "4+2", "3+3"]
        elif num_strings == 7:
            return ["7", "6+1", "5+2", "4+3"]
        elif num_strings == 8:
            return ["8", "7+1", "6+2", "5+3", "4+4"]
        elif num_strings == 9:
            return ["9", "8+1", "7+2", "6+3", "5+4"]
        elif num_strings == 10:
            return ["10", "9+1", "8+2", "7+3", "6+4", "5+5"]
        elif num_strings == 11:
            return ["11", "10+1", "9+2", "8+3", "7+4", "6+5"]
        elif num_strings == 12:
            return ["12", "11+1", "10+2", "9+3", "8+4", "7+5", "6+6"]
        elif num_strings == 13:
            return ["13", "12+1", "11+2", "10+3", "9+4", "8+5", "7+6"]
        elif num_strings == 14:
            return ["14", "13+1", "12+2", "11+3", "10+4", "9+5", "8+6", "7+7"]
        elif num_strings == 15:
            return ["15", "14+1", "13+2", "12+3", "11+4", "10+5", "9+6", "8+7"]
        elif num_strings == 16:
            return ["16", "15+1", "14+2", "13+3", "12+4", "11+5", "10+6", "9+7", "8+8"]
        else:
            return [str(num_strings)]

    def parse_harness_config(self, config_str):
        """Parse harness config string like '2+1' into list of integers [2, 1]"""
        try:
            return [int(x) for x in config_str.split('+')]
        except:
            return []
        
    def load_estimate(self):
        """Load estimate data from the project"""
        if not self.current_project or not self.estimate_id:
            self.add_subarray(auto_add_block=True)
            return
        
        self._loading = True
        
        estimate_data = self.current_project.quick_estimates.get(self.estimate_id)
        if not estimate_data:
            self._loading = False
            self.add_subarray(auto_add_block=True)
            return
        
        # Load global settings
        self.modules_per_string_default = estimate_data.get('modules_per_string', 28)
        self.row_spacing_default = estimate_data.get('row_spacing_ft', 20.0)
        self.wire_gauge_default = estimate_data.get('wire_gauge', '10 AWG')
        
        # Restore module selection
        saved_module_name = estimate_data.get('module_name', '')
        if saved_module_name and hasattr(self, 'module_combo'):
            # Find the matching display name in available modules
            for display_name, module in self.available_modules.items():
                if f"{module.manufacturer} {module.model}" == saved_module_name:
                    self.module_select_var.set(display_name)
                    self._on_module_selected()
                    break
        
        # Update UI if already set up
        if hasattr(self, 'modules_per_string_var'):
            self.modules_per_string_var.set(str(self.modules_per_string_default))
        if hasattr(self, 'row_spacing_var'):
            self.row_spacing_var.set(str(self.row_spacing_default))
        if hasattr(self, 'wire_gauge_var'):
            self.wire_gauge_var.set(estimate_data.get('wire_gauge', '10 AWG'))
        
        # Restore inverter selection
        saved_inverter_name = estimate_data.get('inverter_name', '')
        if saved_inverter_name and hasattr(self, 'inverter_combo'):
            if saved_inverter_name in self.available_inverters:
                self.inverter_select_var.set(saved_inverter_name)
                self._on_inverter_selected()
        
        # Restore topology and DC:AC ratio
        if hasattr(self, 'topology_var'):
            self.topology_var.set(estimate_data.get('topology', 'Distributed String'))
        if hasattr(self, 'dc_ac_ratio_var'):
            self.dc_ac_ratio_var.set(str(estimate_data.get('dc_ac_ratio', 1.25)))
        if hasattr(self, 'dc_feeder_distance_var'):
            self.dc_feeder_distance_var.set(str(estimate_data.get('dc_feeder_distance_ft', 500)))
        if hasattr(self, 'ac_homerun_distance_var'):
            self.ac_homerun_distance_var.set(str(estimate_data.get('ac_homerun_distance_ft', 500)))
        
        # Load subarrays
        saved_subarrays = estimate_data.get('subarrays', {})
        
        if not saved_subarrays:
            # No saved data, create default
            self._loading = False
            self.add_subarray(auto_add_block=True)
            return
        
        # Clear existing tree items
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.subarrays.clear()
        
        # Reconstruct subarrays and blocks
        for subarray_id, subarray_data in saved_subarrays.items():
            # Add to internal data
            self.subarrays[subarray_id] = {
                'name': subarray_data.get('name', 'Subarray'),
                'transformer_mva': subarray_data.get('transformer_mva', 4.0),
                'blocks': {}
            }
            
            # Add to tree
            self.tree.insert('', 'end', subarray_id, text=subarray_data.get('name', 'Subarray'), open=True)
            
            # Load blocks
            for block_id, block_data in subarray_data.get('blocks', {}).items():
                self.subarrays[subarray_id]['blocks'][block_id] = {
                    'name': block_data.get('name', 'Block'),
                    'type': block_data.get('type', 'combiner'),
                    'num_combiners': block_data.get('num_combiners', 1),
                    'breaker_size': block_data.get('breaker_size', 400),
                    'cb_override': block_data.get('cb_override'),
                    'trackers': block_data.get('trackers', [{'strings': 2, 'quantity': 1, 'harness_config': '2'}])
                }
                
                # Add block to tree
                self.tree.insert(subarray_id, 'end', block_id, text=block_data.get('name', 'Block'))
        
        # Select first item
        first_items = self.tree.get_children()
        if first_items:
            self.tree.selection_set(first_items[0])
            self.on_tree_select(None)
        
        self._loading = False

        # Auto-calculate on load
        self.after(100, self.calculate_estimate)

    def save_estimate(self):
        """Save estimate data to the project"""
        if not self.current_project or not self.estimate_id:
            return
        
        # Get current estimate or create new one
        if self.estimate_id not in self.current_project.quick_estimates:
            self.current_project.quick_estimates[self.estimate_id] = {
                'name': 'Unnamed Estimate',
                'created_date': datetime.now().isoformat(),
            }
        
        estimate_data = self.current_project.quick_estimates[self.estimate_id]
        
        # Update global settings
        if self.selected_module:
            estimate_data['module_name'] = f"{self.selected_module.manufacturer} {self.selected_module.model}"
            estimate_data['module_width_mm'] = self.selected_module.width_mm
            estimate_data['module_isc'] = self.selected_module.isc
        
        try:
            estimate_data['modules_per_string'] = int(self.modules_per_string_var.get())
        except ValueError:
            estimate_data['modules_per_string'] = 28
        
        try:
            estimate_data['row_spacing_ft'] = float(self.row_spacing_var.get())
        except ValueError:
            estimate_data['row_spacing_ft'] = 20.0

        estimate_data['wire_gauge'] = self.wire_gauge_var.get()
        
        # Save inverter selection
        if self.selected_inverter:
            estimate_data['inverter_name'] = f"{self.selected_inverter.manufacturer} {self.selected_inverter.model}"
        
        # Save topology and DC:AC ratio
        estimate_data['topology'] = self.topology_var.get()
        try:
            estimate_data['dc_ac_ratio'] = float(self.dc_ac_ratio_var.get())
        except ValueError:
            estimate_data['dc_ac_ratio'] = 1.25
        
        # Save distance inputs
        try:
            estimate_data['dc_feeder_distance_ft'] = float(self.dc_feeder_distance_var.get())
        except ValueError:
            estimate_data['dc_feeder_distance_ft'] = 500.0
        try:
            estimate_data['ac_homerun_distance_ft'] = float(self.ac_homerun_distance_var.get())
        except ValueError:
            estimate_data['ac_homerun_distance_ft'] = 500.0
        
        # Update modified date
        estimate_data['modified_date'] = datetime.now().isoformat()
        
        # Save subarrays
        estimate_data['subarrays'] = {}
        
        for subarray_id, subarray_data in self.subarrays.items():
            estimate_data['subarrays'][subarray_id] = {
                'name': subarray_data['name'],
                'transformer_mva': subarray_data['transformer_mva'],
                'blocks': {}
            }
            
            for block_id, block_data in subarray_data['blocks'].items():
                estimate_data['subarrays'][subarray_id]['blocks'][block_id] = {
                    'name': block_data['name'],
                    'type': block_data['type'],
                    'num_combiners': block_data.get('num_combiners', 1),
                    'breaker_size': block_data.get('breaker_size', 400),
                    'cb_override': block_data.get('cb_override'),
                    'trackers': block_data['trackers']
                }
        
        # Notify callback
        if self.on_save:
            self.on_save()

    # ==================== Estimate Management ====================

    def _refresh_estimate_dropdown(self, auto_select=False):
        """Refresh the estimate dropdown with saved estimates"""
        if not self.current_project:
            return
        
        estimates = self.current_project.quick_estimates
        
        # Build list of estimate names for dropdown
        estimate_names = []
        self._estimate_id_map = {}
        
        for est_id, est_data in estimates.items():
            name = est_data.get('name', 'Unnamed Estimate')
            estimate_names.append(name)
            self._estimate_id_map[name] = est_id
        
        if hasattr(self, 'estimate_combo'):
            self.estimate_combo['values'] = estimate_names
        
        if auto_select and estimates:
            # Find most recently modified
            most_recent_id = None
            most_recent_date = None
            for est_id, est_data in estimates.items():
                mod_date = est_data.get('modified_date')
                if mod_date:
                    if most_recent_date is None or mod_date > most_recent_date:
                        most_recent_date = mod_date
                        most_recent_id = est_id
            
            if most_recent_id is None:
                most_recent_id = list(estimates.keys())[0]
            
            self.estimate_id = most_recent_id
            if hasattr(self, 'estimate_var'):
                self.estimate_var.set(estimates[most_recent_id].get('name', ''))
            self.load_estimate()
        elif auto_select and not estimates:
            # No estimates exist — create a default one
            self.new_estimate()

    def _on_estimate_selected(self, event=None):
        """Handle selection of an estimate from the dropdown"""
        # Save current estimate before switching
        if self.estimate_id:
            self.save_estimate()
        
        selected_name = self.estimate_var.get()
        if selected_name in self._estimate_id_map:
            self.estimate_id = self._estimate_id_map[selected_name]
            self.load_estimate()

    def _on_estimate_name_changed(self, event=None):
        """Handle renaming of the current estimate via the combobox"""
        if not self.estimate_id or not self.current_project:
            return
        
        new_name = self.estimate_var.get().strip()
        if not new_name:
            return
        
        if self.estimate_id in self.current_project.quick_estimates:
            self.current_project.quick_estimates[self.estimate_id]['name'] = new_name
            self._refresh_estimate_dropdown()
            self.estimate_var.set(new_name)

    def new_estimate(self):
        """Create a new quick estimate"""
        if not self.current_project:
            return
        
        # Save current estimate first
        if self.estimate_id:
            self.save_estimate()
        
        # Generate new ID and name
        estimate_id = f"estimate_{uuid.uuid4().hex[:8]}"
        estimate_num = len(self.current_project.quick_estimates) + 1
        estimate_name = f"Estimate {estimate_num}"
        
        new_estimate = {
            'name': estimate_name,
            'created_date': datetime.now().isoformat(),
            'modified_date': datetime.now().isoformat(),
            'module_width_mm': 1134,
            'modules_per_string': 28,
            'row_spacing_ft': 20.0,
            'topology': 'Distributed String',
            'dc_ac_ratio': 1.25,
            'subarrays': {}
        }
        
        self.current_project.quick_estimates[estimate_id] = new_estimate
        
        self.estimate_id = estimate_id
        self._refresh_estimate_dropdown()
        self.estimate_var.set(estimate_name)
        
        # Clear and set up fresh
        self._clear_estimate_ui()
        self.add_subarray(auto_add_block=True)
        
        if self.on_save:
            self.on_save()

    def delete_estimate(self):
        """Delete the currently selected estimate"""
        if not self.current_project or not self.estimate_id:
            from tkinter import messagebox
            messagebox.showinfo("No Estimate", "No estimate selected to delete.")
            return
        
        from tkinter import messagebox
        estimate_name = self.current_project.quick_estimates.get(
            self.estimate_id, {}
        ).get('name', 'this estimate')
        
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{estimate_name}'?"):
            del self.current_project.quick_estimates[self.estimate_id]
            self.estimate_id = None
            self._clear_estimate_ui()
            self._refresh_estimate_dropdown(auto_select=True)
            
            if self.on_save:
                self.on_save()

    def _clear_estimate_ui(self):
        """Clear the tree and details when switching/deleting estimates"""
        # Clear tree
        if hasattr(self, 'tree'):
            for item in self.tree.get_children():
                self.tree.delete(item)
        self.subarrays.clear()
        
        # Clear details panel
        if hasattr(self, 'details_container'):
            self.clear_details_panel()
        
        # Clear results
        if hasattr(self, 'results_tree'):
            for item in self.results_tree.get_children():
                self.results_tree.delete(item)
        
        # Reset inverter and topology fields
        if hasattr(self, 'inverter_select_var'):
            self.inverter_select_var.set('')
            self.selected_inverter = None
            self.inverter_info_label.config(text="No inverter selected", foreground='gray')
        if hasattr(self, 'topology_var'):
            self.topology_var.set('Distributed String')
        if hasattr(self, 'dc_ac_ratio_var'):
            self.dc_ac_ratio_var.set('1.25')
        if hasattr(self, 'strings_per_inverter_var'):
            self.strings_per_inverter_var.set('--')
        if hasattr(self, 'isc_warning_label'):
            self.isc_warning_label.config(text="")
        if hasattr(self, 'dc_feeder_distance_var'):
            self.dc_feeder_distance_var.set('500')
        if hasattr(self, 'ac_homerun_distance_var'):
            self.ac_homerun_distance_var.set('500')

    def _schedule_autosave(self):
        """Debounced auto-save — saves estimate after a brief pause"""
        if getattr(self, '_loading', False):
            return
        if hasattr(self, '_autosave_after_id') and self._autosave_after_id:
            self.after_cancel(self._autosave_after_id)
        self._autosave_after_id = self.after(1000, self._do_autosave)
    
    def _do_autosave(self):
        """Execute the debounced save"""
        self._autosave_after_id = None
        self.save_estimate()

    def _mark_stale(self):
        """Mark results as stale and re-enable the calculate button"""
        self._results_stale = True
        if self._calc_btn:
            self._calc_btn.config(state='normal')

    def _on_module_selected(self, event=None):
        """Handle module selection from dropdown"""
        selected_name = self.module_select_var.get()
        if selected_name in self.available_modules:
            self.selected_module = self.available_modules[selected_name]
            self.module_info_label.config(
                text=f"Isc: {self.selected_module.isc}A  |  Width: {self.selected_module.width_mm}mm  |  Voc: {self.selected_module.voc}V",
                foreground='black'
            )
            # Recalculate strings per inverter with new module
            self._update_strings_per_inverter()
            # Auto-save when module changes (but not during load)
            if not getattr(self, '_loading', False):
                self._mark_stale()
                self.save_estimate()
        else:
            self.selected_module = None
            self.module_info_label.config(text="No module selected", foreground='gray')

    def _on_inverter_selected(self, event=None):
        """Handle inverter selection from dropdown"""
        selected_name = self.inverter_select_var.get()
        if selected_name in self.available_inverters:
            self.selected_inverter = self.available_inverters[selected_name]
            inv = self.selected_inverter
            type_str = inv.inverter_type.value if hasattr(inv, 'inverter_type') else 'String'
            self.inverter_info_label.config(
                text=f"{inv.rated_power_kw}kW AC  |  {inv.max_dc_power_kw}kW DC  |  {inv.get_total_string_capacity()} inputs  |  {type_str}",
                foreground='black'
            )
            self._update_strings_per_inverter()
            # Auto-save when inverter changes (but not during load)
            if not getattr(self, '_loading', False):
                self._mark_stale()
                self.save_estimate()
        else:
            self.selected_inverter = None
            self.inverter_info_label.config(text="No inverter selected", foreground='gray')
            self.strings_per_inverter_var.set('--')
            self.isc_warning_label.config(text="")

    # Inverter group color palette
    INVERTER_COLORS = [
        '#4A90D9',  # Blue
        '#E67E22',  # Orange
        '#2ECC71',  # Green
        '#9B59B6',  # Purple
        '#E74C3C',  # Red
        '#1ABC9C',  # Teal
        '#F39C12',  # Yellow
        '#3498DB',  # Light Blue
        '#E91E63',  # Pink
        '#00BCD4',  # Cyan
        '#8BC34A',  # Light Green
        '#FF9800',  # Amber
        '#795548',  # Brown
        '#607D8B',  # Blue Gray
        '#CDDC39',  # Lime
    ]

    def show_site_preview(self):
        """Open the site preview in a pop-out window"""
        inv_summary = getattr(self, 'last_totals', {}).get('inverter_summary', {})
        
        if not inv_summary or not inv_summary.get('allocations'):
            from tkinter import messagebox
            messagebox.showinfo("No Data", "Run Calculate Estimate first to generate preview data.")
            return
        
        topology = self.topology_var.get()
        SitePreviewWindow(self, inv_summary, topology, self.INVERTER_COLORS)

    def _refresh_block_details(self):
        """Re-render the block detail panel if a block is currently selected."""
        if self.selected_item_type == 'block' and self.selected_item_id:
            parent_id = self.tree.parent(self.selected_item_id)
            if parent_id:
                for widget in self.details_container.winfo_children():
                    widget.destroy()
                self.show_block_details(parent_id, self.selected_item_id)

    def _on_strings_per_inverter_changed(self, *args):
        """When user manually edits strings/inverter, reverse-calculate DC:AC ratio."""
        if self._updating_spi:
            return
        if not self.selected_inverter or not self.selected_module:
            return
        
        try:
            spi = int(self.strings_per_inverter_var.get())
        except ValueError:
            return
        
        if spi <= 0:
            return
        
        try:
            modules_per_string = int(self.modules_per_string_var.get())
        except ValueError:
            modules_per_string = 28
        
        # Reverse-calculate DC:AC ratio
        module_wattage = self.selected_module.wattage
        actual_ratio = self.selected_inverter.dc_ac_ratio(spi, module_wattage, modules_per_string)
        
        if actual_ratio > 0:
            # Update DC:AC ratio without triggering forward calc
            self._updating_spi = True
            self.dc_ac_ratio_var.set(f"{actual_ratio:.2f}")
            self._updating_spi = False
        
        # Update Isc warning for new string count
        module_isc = self.selected_module.isc
        total_isc = spi * module_isc
        max_isc = getattr(self.selected_inverter, 'max_short_circuit_current', None)
        
        if max_isc and total_isc > max_isc:
            self.isc_warning_label.config(
                text=f"⚠️ Isc limit exceeded: {total_isc:.1f}A > {max_isc:.0f}A max"
            )
        else:
            self.isc_warning_label.config(text="")
        
        self._mark_stale()
        self._schedule_autosave()

    def _update_distance_hints(self):
        """Update the hint text next to distance inputs based on topology."""
        if not hasattr(self, 'distance_hint_label'):
            return
        topology = self.topology_var.get()
        if topology == 'Distributed String':
            self.distance_hint_label.config(text="(DC feeders N/A for distributed — AC homeruns are primary cable)")
        elif topology == 'Central Inverter':
            self.distance_hint_label.config(text="(Long DC feeders to central pad — short AC from inverter)")
        elif topology == 'Centralized String':
            self.distance_hint_label.config(text="(DC feeders to inverter bank — short AC from bank)")
        else:
            self.distance_hint_label.config(text="")

    def _get_harness_config_for_tracker_type(self, strings_per_tracker):
        """Find the harness config used for trackers with the given string count."""
        for subarray_id, subarray in self.subarrays.items():
            for block_id, block in subarray['blocks'].items():
                for tracker in block['trackers']:
                    if tracker['strings'] == strings_per_tracker and tracker['quantity'] > 0:
                        return self.parse_harness_config(tracker['harness_config'])
        # Fallback: single harness equal to string count
        return [strings_per_tracker]

    def _adjust_harnesses_for_splits(self, totals):
        """Adjust harness counts based on inverter allocation split trackers.
        
        When a tracker is split between two inverters, its original harness config
        is replaced by harnesses matching the split amounts.
        E.g., a 3-string tracker split 1/2 → remove one 3-string harness, 
              add one 1-string + one 2-string harness.
        """
        inv_summary = totals.get('inverter_summary', {})
        allocations = inv_summary.get('allocations', [])
        
        if not allocations:
            return
        
        for alloc_info in allocations:
            strings_per_tracker = alloc_info['strings_per_tracker']
            allocation = alloc_info['allocation']
            inverters = allocation.get('inverters', [])
            
            if not inverters:
                continue
            
            # Find the original harness config for this tracker type
            original_harness_sizes = self._get_harness_config_for_tracker_type(strings_per_tracker)
            
            split_count = 0
            
            # Check each inverter's pattern for split trackers
            for i, inv_info in enumerate(inverters):
                pattern = inv_info['pattern']
                tail = pattern[-1]
                
                if tail >= strings_per_tracker:
                    continue  # Full tracker at boundary, no split
                
                # Last inverter with partial load = end of allocation, not a split
                is_last = (i == len(inverters) - 1)
                if is_last and allocation['summary'].get('partial_inverter_strings', 0) > 0:
                    continue
                
                complement = strings_per_tracker - tail
                
                # Remove original harness(es) for one tracker
                for size in original_harness_sizes:
                    if size in totals['harnesses_by_size']:
                        totals['harnesses_by_size'][size] -= 1
                        if totals['harnesses_by_size'][size] <= 0:
                            del totals['harnesses_by_size'][size]
                
                # Add split harnesses (one per side of the split)
                for split_size in [tail, complement]:
                    if split_size not in totals['harnesses_by_size']:
                        totals['harnesses_by_size'][split_size] = 0
                    totals['harnesses_by_size'][split_size] += 1
                
                split_count += 1
            
            print(f"  Harness adjustment for {strings_per_tracker}-string trackers: {split_count} split trackers adjusted")

    def _update_strings_per_inverter(self):
        """Auto-calculate strings per inverter from DC:AC ratio and show Isc warning if needed"""
        if self._updating_spi:
            return
        if not self.selected_inverter or not self.selected_module:
            self.strings_per_inverter_var.set('--')
            self.isc_warning_label.config(text="")
            return
        
        try:
            target_ratio = float(self.dc_ac_ratio_var.get())
        except ValueError:
            self.strings_per_inverter_var.set('--')
            return
        
        try:
            modules_per_string = int(self.modules_per_string_var.get())
        except ValueError:
            modules_per_string = 28
        
        module_wattage = self.selected_module.wattage
        string_power_kw = (module_wattage * modules_per_string) / 1000
        
        if string_power_kw <= 0:
            self.strings_per_inverter_var.set('--')
            return
        
        topology = self.topology_var.get()
        
        # Calculate power-based string count (always applies)
        target_dc_kw = target_ratio * self.selected_inverter.rated_power_kw
        power_based_strings = round(target_dc_kw / string_power_kw)
        
        # Also check DC power limit
        dc_power_limited = int(self.selected_inverter.max_dc_power_kw / string_power_kw)
        
        if topology == 'Distributed String':
            # Physical input count matters — strings connect directly to inverter
            input_limited = self.selected_inverter.get_total_string_capacity()
            strings_per_inv = min(power_based_strings, dc_power_limited, input_limited)
        else:
            # Centralized String or Central Inverter — combiners aggregate strings
            # Only power limits apply, not physical input count
            strings_per_inv = min(power_based_strings, dc_power_limited)
        
        strings_per_inv = max(strings_per_inv, 1)  # At least 1
        self._updating_spi = True
        self.strings_per_inverter_var.set(str(strings_per_inv))
        self._updating_spi = False
        
        # Calculate actual DC:AC ratio achieved
        actual_ratio = self.selected_inverter.dc_ac_ratio(
            strings_per_inv, module_wattage, modules_per_string
        )
        
        # Show warning if DC power limit or input limit is capping the ratio
        if strings_per_inv < power_based_strings:
            max_achievable_ratio = self.selected_inverter.max_dc_power_kw / self.selected_inverter.rated_power_kw
            if topology == 'Distributed String' and self.selected_inverter.get_total_string_capacity() < dc_power_limited:
                self.inverter_info_label.config(
                    text=f"{self.selected_inverter.rated_power_kw}kW AC  |  Capped by string inputs ({self.selected_inverter.get_total_string_capacity()})  |  Max DC:AC ≈ {actual_ratio:.2f}",
                    foreground='orange'
                )
            else:
                self.inverter_info_label.config(
                    text=f"{self.selected_inverter.rated_power_kw}kW AC  |  Capped by DC power ({self.selected_inverter.max_dc_power_kw}kW)  |  Max DC:AC = {max_achievable_ratio:.2f}",
                    foreground='orange'
                )
        
        # Check Isc hard limit
        module_isc = self.selected_module.isc
        total_isc = strings_per_inv * module_isc
        max_isc = getattr(self.selected_inverter, 'max_short_circuit_current', None)
        
        if max_isc and total_isc > max_isc:
            self.isc_warning_label.config(
                text=f"⚠️ Isc limit exceeded: {total_isc:.1f}A > {max_isc:.0f}A max"
            )
        else:
            self.isc_warning_label.config(text="")


    # ==================== UI Setup ====================
    
    def setup_ui(self):
        """Create and arrange UI components"""
        # Main container with padding
        main_container = ttk.Frame(self, padding="10")
        main_container.pack(fill='both', expand=True)
        
        # Top bar: Title + Estimate selector
        top_bar = ttk.Frame(main_container)
        top_bar.pack(fill='x', pady=(0, 10))
        
        # Title
        title_label = ttk.Label(top_bar, text="Quick Estimate", font=('Helvetica', 14, 'bold'))
        title_label.pack(side='left', padx=(0, 20))
        
        # Estimate selector
        ttk.Label(top_bar, text="Estimate:").pack(side='left', padx=(0, 5))
        self.estimate_var = tk.StringVar()
        self.estimate_combo = ttk.Combobox(
            top_bar,
            textvariable=self.estimate_var,
            width=30,
            state='normal'
        )
        self.estimate_combo.pack(side='left', padx=(0, 5))
        self.estimate_combo.bind('<<ComboboxSelected>>', self._on_estimate_selected)
        self.estimate_combo.bind('<Return>', self._on_estimate_name_changed)
        self.estimate_combo.bind('<FocusOut>', self._on_estimate_name_changed)
        
        ttk.Button(top_bar, text="New", command=self.new_estimate, width=6).pack(side='left', padx=2)
        ttk.Button(top_bar, text="Delete", command=self.delete_estimate, width=6).pack(side='left', padx=2)
        
        # Description
        desc_label = ttk.Label(main_container, text="Early-stage BOM estimation for bid and preliminary designs", foreground='gray')
        desc_label.pack(anchor='w', pady=(0, 10))
        
        # Main content area - PanedWindow for resizable split
        paned = ttk.PanedWindow(main_container, orient='horizontal')
        paned.pack(fill='both', expand=True, pady=(0, 10))
        
        # Left panel - Tree view
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=1)
        
        # Right panel - Details
        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=3)
        
        self.setup_tree_panel(left_frame)
        self.setup_details_panel(right_frame)
        
        # Bottom section - Calculate button and results
        bottom_frame = ttk.Frame(main_container)
        bottom_frame.pack(fill='both', expand=True)
        
        # Global settings frame (applies to all calculations)
        settings_frame = ttk.LabelFrame(bottom_frame, text="Global Settings", padding="5")
        settings_frame.pack(fill='x', pady=(0, 10))
        
        # Row 1: Module selection
        module_row = ttk.Frame(settings_frame)
        module_row.pack(fill='x', pady=(0, 5))
        
        ttk.Label(module_row, text="Module:").pack(side='left', padx=(0, 5))
        self.module_select_var = tk.StringVar()
        self.module_combo = ttk.Combobox(
            module_row,
            textvariable=self.module_select_var,
            values=sorted(self.available_modules.keys()),
            state='readonly',
            width=50
        )
        self.module_combo.pack(side='left', padx=(0, 15))
        self.module_combo.bind('<<ComboboxSelected>>', self._on_module_selected)
        self.disable_combobox_scroll(self.module_combo)
        
        # Module info display
        self.module_info_label = ttk.Label(module_row, text="No module selected", foreground='gray')
        self.module_info_label.pack(side='left', padx=(5, 0))
        
        # Row 2: Inverter selection
        inverter_row = ttk.Frame(settings_frame)
        inverter_row.pack(fill='x', pady=(0, 5))
        
        ttk.Label(inverter_row, text="Inverter:").pack(side='left', padx=(0, 5))
        self.inverter_select_var = tk.StringVar()
        self.inverter_combo = ttk.Combobox(
            inverter_row,
            textvariable=self.inverter_select_var,
            values=sorted(self.available_inverters.keys()),
            state='readonly',
            width=50
        )
        self.inverter_combo.pack(side='left', padx=(0, 15))
        self.inverter_combo.bind('<<ComboboxSelected>>', self._on_inverter_selected)
        self.disable_combobox_scroll(self.inverter_combo)
        
        # Inverter info display
        self.inverter_info_label = ttk.Label(inverter_row, text="No inverter selected", foreground='gray')
        self.inverter_info_label.pack(side='left', padx=(5, 0))
        
        # Row 3: Topology and DC:AC ratio
        topology_row = ttk.Frame(settings_frame)
        topology_row.pack(fill='x', pady=(0, 5))
        
        ttk.Label(topology_row, text="Topology:").pack(side='left', padx=(0, 5))
        self.topology_var = tk.StringVar(value='Distributed String')
        topology_combo = ttk.Combobox(
            topology_row,
            textvariable=self.topology_var,
            values=['Distributed String', 'Centralized String', 'Central Inverter'],
            state='readonly',
            width=20
        )
        topology_combo.pack(side='left', padx=(0, 15))
        self.disable_combobox_scroll(topology_combo)
        self.topology_var.trace_add('write', lambda *args: (self._update_distance_hints(), self._refresh_block_details(), self._mark_stale(), self._schedule_autosave()))
        
        ttk.Label(topology_row, text="DC:AC Ratio:").pack(side='left', padx=(0, 5))
        self.dc_ac_ratio_var = tk.StringVar(value='1.25')
        ttk.Spinbox(
            topology_row, from_=1.0, to=2.0, increment=0.05,
            textvariable=self.dc_ac_ratio_var, width=6, format='%.2f'
        ).pack(side='left', padx=(0, 15))
        self.dc_ac_ratio_var.trace_add('write', lambda *args: (self._update_strings_per_inverter(), self._mark_stale(), self._schedule_autosave()))
        
        ttk.Label(topology_row, text="Strings/Inverter:").pack(side='left', padx=(0, 5))
        self.strings_per_inverter_var = tk.StringVar(value='--')
        self._updating_spi = False  # Guard to prevent infinite loop
        spi_spinbox = ttk.Spinbox(
            topology_row, from_=1, to=100,
            textvariable=self.strings_per_inverter_var, width=6,
            font=('Helvetica', 10, 'bold')
        )
        spi_spinbox.pack(side='left', padx=(0, 15))
        self.strings_per_inverter_var.trace_add('write', self._on_strings_per_inverter_changed)
        
        # Isc warning label (hidden by default)
        self.isc_warning_label = ttk.Label(topology_row, text="", foreground='red')
        self.isc_warning_label.pack(side='left', padx=(5, 0))
        
        # Row 4: Other settings
        settings_inner = ttk.Frame(settings_frame)
        settings_inner.pack(fill='x')
        
        ttk.Label(settings_inner, text="Modules per String:").pack(side='left', padx=(0, 5))
        self.modules_per_string_var = tk.StringVar(value=str(getattr(self, 'modules_per_string_default', 28)))
        ttk.Spinbox(settings_inner, from_=1, to=100, textvariable=self.modules_per_string_var, width=6).pack(side='left', padx=(0, 15))
        self.modules_per_string_var.trace_add('write', lambda *args: (self._update_strings_per_inverter(), self._mark_stale(), self._schedule_autosave()))
        
        ttk.Label(settings_inner, text="Row Spacing (ft):").pack(side='left', padx=(0, 5))
        self.row_spacing_var = tk.StringVar(value=str(getattr(self, 'row_spacing_default', 20.0)))
        ttk.Spinbox(settings_inner, from_=1, to=100, textvariable=self.row_spacing_var, width=6).pack(side='left', padx=(0, 15))
        self.row_spacing_var.trace_add('write', lambda *args: (self._mark_stale(), self._schedule_autosave()))

        ttk.Label(settings_inner, text="Wire Gauge:").pack(side='left', padx=(0, 5))
        self.wire_gauge_var = tk.StringVar(value=getattr(self, 'wire_gauge_default', '8 AWG'))
        wire_gauge_combo = ttk.Combobox(
            settings_inner,
            textvariable=self.wire_gauge_var,
            values=['8 AWG', '10 AWG'],
            state='readonly',
            width=8
        )
        wire_gauge_combo.pack(side='left')
        self.disable_combobox_scroll(wire_gauge_combo)
        self.wire_gauge_var.trace_add('write', lambda *args: (self._mark_stale(), self._schedule_autosave()))
        
        # Row 5: Topology-dependent distance inputs
        distance_row = ttk.Frame(settings_frame)
        distance_row.pack(fill='x', pady=(5, 0))
        
        ttk.Label(distance_row, text="Avg DC Feeder Run (ft):").pack(side='left', padx=(0, 5))
        self.dc_feeder_distance_var = tk.StringVar(value='500')
        ttk.Spinbox(distance_row, from_=0, to=5000, increment=50,
                     textvariable=self.dc_feeder_distance_var, width=8).pack(side='left', padx=(0, 15))
        self.dc_feeder_distance_var.trace_add('write', lambda *args: (self._mark_stale(), self._schedule_autosave()))
        
        ttk.Label(distance_row, text="Avg AC Homerun (ft):").pack(side='left', padx=(0, 5))
        self.ac_homerun_distance_var = tk.StringVar(value='500')
        ttk.Spinbox(distance_row, from_=0, to=5000, increment=50,
                     textvariable=self.ac_homerun_distance_var, width=8).pack(side='left', padx=(0, 15))
        self.ac_homerun_distance_var.trace_add('write', lambda *args: (self._mark_stale(), self._schedule_autosave()))
        
        # Topology hint label
        self.distance_hint_label = ttk.Label(distance_row, text="", foreground='gray')
        self.distance_hint_label.pack(side='left', padx=(5, 0))
        self._update_distance_hints()
        
        # Button row
        button_row = ttk.Frame(bottom_frame)
        button_row.pack(fill='x', pady=(0, 10))
        
        self._calc_btn = ttk.Button(button_row, text="Calculate Estimate", command=self.calculate_estimate, state='disabled')
        self._calc_btn.pack(side='left', padx=(0, 10))
        
        export_btn = ttk.Button(button_row, text="Export to Excel", command=self.export_to_excel)
        export_btn.pack(side='left')
        
        preview_btn = ttk.Button(button_row, text="Site Preview", command=self.show_site_preview)
        preview_btn.pack(side='left', padx=(10, 0))
        
        # Results frame (full width)
        results_frame = ttk.LabelFrame(bottom_frame, text="Estimated BOM (Rolled-Up Totals)", padding="10")
        results_frame.pack(fill='both', expand=True)
        
        # Results treeview
        columns = ('include', 'item', 'part_number', 'quantity', 'unit', 'unit_cost', 'ext_cost')
        self.results_tree = ttk.Treeview(results_frame, columns=columns, show='headings', height=8)
        self.results_tree.heading('include', text='')
        self.results_tree.heading('item', text='Item')
        self.results_tree.heading('part_number', text='Part Number')
        self.results_tree.heading('quantity', text='Quantity')
        self.results_tree.heading('unit', text='Unit')
        self.results_tree.heading('unit_cost', text='Unit Cost')
        self.results_tree.heading('ext_cost', text='Ext. Cost')
        self.results_tree.column('include', width=30, anchor='center')
        self.results_tree.column('item', width=230, anchor='w')
        self.results_tree.column('part_number', width=160, anchor='w')
        self.results_tree.column('quantity', width=80, anchor='center')
        self.results_tree.column('unit', width=50, anchor='center')
        self.results_tree.column('unit_cost', width=80, anchor='e')
        self.results_tree.column('ext_cost', width=90, anchor='e')

        # Tags for checked/unchecked styling
        self.results_tree.tag_configure('checked', foreground='black')
        self.results_tree.tag_configure('unchecked', foreground='gray60')
        self.results_tree.tag_configure('section', foreground='black')

        # Click handler for checkbox column
        self.results_tree.bind('<Button-1>', self._on_results_tree_click)
        
        # Scrollbar for results
        results_scroll = ttk.Scrollbar(results_frame, orient='vertical', command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=results_scroll.set)
        
        self.results_tree.pack(side='left', fill='both', expand=True)
        results_scroll.pack(side='right', fill='y')

        # Select all / none buttons
        check_btn_frame = ttk.Frame(results_frame)
        check_btn_frame.pack(fill='x', pady=(5, 0))
        ttk.Button(check_btn_frame, text="Select All", command=self._results_select_all).pack(side='left', padx=(0, 5))
        ttk.Button(check_btn_frame, text="Select None", command=self._results_select_none).pack(side='left')

    def setup_tree_panel(self, parent):
        """Setup the left tree view panel"""
        # Tree frame with label
        tree_label = ttk.Label(parent, text="Project Structure", font=('Helvetica', 10, 'bold'))
        tree_label.pack(anchor='w', pady=(0, 5))
        
        # Tree view
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill='both', expand=True)
        
        self.tree = ttk.Treeview(tree_frame, show='tree', selectmode='browse')
        self.tree.pack(side='left', fill='both', expand=True)
        
        # Scrollbar
        tree_scroll = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        tree_scroll.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=tree_scroll.set)
        
        # Bind selection event
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        
        # Buttons frame
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill='x', pady=(10, 0))
        
        add_subarray_btn = ttk.Button(btn_frame, text="+ Subarray", command=lambda: self.add_subarray(auto_add_block=True))
        add_subarray_btn.pack(side='left', padx=(0, 5))
        
        add_block_btn = ttk.Button(btn_frame, text="+ Block", command=self.add_block_to_selected)
        add_block_btn.pack(side='left', padx=(0, 5))
        
        copy_block_btn = ttk.Button(btn_frame, text="Copy Block", command=self.copy_selected_block)
        copy_block_btn.pack(side='left', padx=(0, 5))
        
        delete_btn = ttk.Button(btn_frame, text="Delete", command=self.delete_selected_item)
        delete_btn.pack(side='left')

    def add_block_to_selected(self):
        """Add a block to the currently selected subarray"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item_id = selection[0]
        parent_id = self.tree.parent(item_id)
        
        # Determine which subarray to add to
        if parent_id == '':
            # Selected item is a subarray
            subarray_id = item_id
        else:
            # Selected item is a block, get its parent subarray
            subarray_id = parent_id
        
        if subarray_id in self.subarrays:
            self.add_block(subarray_id)

    def copy_selected_block(self):
        """Copy the currently selected block with all its configurations"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item_id = selection[0]
        parent_id = self.tree.parent(item_id)
        
        # Only copy if a block is selected (not a subarray)
        if parent_id == '':
            # Selected item is a subarray, not a block
            return
        
        subarray_id = parent_id
        source_block_id = item_id
        
        if subarray_id not in self.subarrays:
            return
        if source_block_id not in self.subarrays[subarray_id]['blocks']:
            return
        
        # Get source block data
        source_block = self.subarrays[subarray_id]['blocks'][source_block_id]
        
        # Create new block ID
        new_block_id = self.generate_id("block")
        
        # Deep copy the tracker configurations
        import copy
        copied_trackers = copy.deepcopy(source_block['trackers'])
        
        # Create new block with copied data
        new_block_name = self.get_next_block_name(subarray_id, source_block['name'])
        
        self.subarrays[subarray_id]['blocks'][new_block_id] = {
            'name': new_block_name,
            'type': source_block['type'],
            'num_combiners': source_block.get('num_combiners', 1),
            'breaker_size': source_block.get('breaker_size', 400),
            'cb_override': source_block.get('cb_override'),
            'trackers': copied_trackers
        }
        
        # Add to tree view under the same subarray
        self.tree.insert(subarray_id, 'end', new_block_id, text=new_block_name)
        
        # Select the new block
        self.tree.selection_set(new_block_id)
        self.on_tree_select(None)

    def setup_details_panel(self, parent):
        """Setup the right details panel"""
        # Details label
        self.details_label = ttk.Label(parent, text="Details", font=('Helvetica', 10, 'bold'))
        self.details_label.pack(anchor='w', pady=(0, 5))
        
        # Container for dynamic content
        self.details_container = ttk.Frame(parent)
        self.details_container.pack(fill='both', expand=True)
        
        # Placeholder
        self.placeholder_label = ttk.Label(self.details_container, text="Select a subarray or block to view details", foreground='gray')
        self.placeholder_label.pack(pady=20)

    def clear_details_panel(self):
        """Clear the details panel"""
        for widget in self.details_container.winfo_children():
            widget.destroy()
        
        self.placeholder_label = ttk.Label(self.details_container, text="Select a subarray or block to view details", foreground='gray')
        self.placeholder_label.pack(pady=20)
        
        self.details_label.config(text="Details")

    def on_tree_select(self, event):
        """Handle tree selection changes"""
        selection = self.tree.selection()
        if not selection:
            self.clear_details_panel()
            return
        
        item_id = selection[0]
        parent_id = self.tree.parent(item_id)
        
        # Clear existing details
        for widget in self.details_container.winfo_children():
            widget.destroy()
        
        if parent_id == '':
            # It's a subarray
            self.selected_item_id = item_id
            self.selected_item_type = 'subarray'
            self.show_subarray_details(item_id)
        else:
            # It's a block
            self.selected_item_id = item_id
            self.selected_item_type = 'block'
            self.show_block_details(parent_id, item_id)

    def show_subarray_details(self, subarray_id: str):
        """Show details panel for a subarray"""
        if subarray_id not in self.subarrays:
            return
        
        subarray = self.subarrays[subarray_id]
        self.details_label.config(text=f"Subarray: {subarray['name']}")
        
        # Create form
        form_frame = ttk.Frame(self.details_container, padding="10")
        form_frame.pack(fill='x')
        
        # Subarray Name
        ttk.Label(form_frame, text="Name:").grid(row=0, column=0, sticky='w', pady=5)
        name_var = tk.StringVar(value=subarray['name'])
        name_entry = ttk.Entry(form_frame, textvariable=name_var, width=30)
        name_entry.grid(row=0, column=1, sticky='w', pady=5, padx=(10, 0))
        
        def update_name(*args):
            subarray['name'] = name_var.get()
            self.tree.item(subarray_id, text=name_var.get())
            self.details_label.config(text=f"Subarray: {name_var.get()}")
            self._schedule_autosave()
        name_var.trace_add('write', update_name)
        
        # Transformer MVA
        ttk.Label(form_frame, text="Transformer (MVA):").grid(row=1, column=0, sticky='w', pady=5)
        mva_var = tk.StringVar(value=str(subarray['transformer_mva']))
        mva_spinbox = ttk.Spinbox(form_frame, from_=0.5, to=100, increment=0.5, textvariable=mva_var, width=10)
        mva_spinbox.grid(row=1, column=1, sticky='w', pady=5, padx=(10, 0))
        
        def update_mva(*args):
            try:
                subarray['transformer_mva'] = float(mva_var.get())
            except ValueError:
                pass
            self._schedule_autosave()
        mva_var.trace_add('write', update_mva)
        
        # Summary info
        summary_frame = ttk.LabelFrame(self.details_container, text="Summary", padding="10")
        summary_frame.pack(fill='x', pady=(20, 0), padx=10)
        
        num_blocks = len(subarray['blocks'])
        total_trackers = sum(
            sum(t['quantity'] for t in block['trackers'])
            for block in subarray['blocks'].values()
        )
        
        ttk.Label(summary_frame, text=f"Blocks: {num_blocks}").pack(anchor='w')
        ttk.Label(summary_frame, text=f"Total Trackers: {total_trackers}").pack(anchor='w')

    def show_block_details(self, subarray_id: str, block_id: str):
        """Show details panel for a block"""
        if subarray_id not in self.subarrays:
            return
        if block_id not in self.subarrays[subarray_id]['blocks']:
            return
        
        block = self.subarrays[subarray_id]['blocks'][block_id]
        self.details_label.config(text=f"Block: {block['name']}")
        
        # Create scrollable frame for block details
        canvas = tk.Canvas(self.details_container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.details_container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        def _bind_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        def _unbind_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
        
        # Bind when mouse enters canvas, unbind when it leaves
        canvas.bind("<Enter>", _bind_mousewheel)
        canvas.bind("<Leave>", _unbind_mousewheel)
        scrollable_frame.bind("<Enter>", _bind_mousewheel)
        scrollable_frame.bind("<Leave>", _unbind_mousewheel)
        
        # Form frame
        form_frame = ttk.Frame(scrollable_frame, padding="10")
        form_frame.pack(fill='x')
        
        # Block Name
        ttk.Label(form_frame, text="Name:").grid(row=0, column=0, sticky='w', pady=5)
        name_var = tk.StringVar(value=block['name'])
        name_entry = ttk.Entry(form_frame, textvariable=name_var, width=25)
        name_entry.grid(row=0, column=1, sticky='w', pady=5, padx=(10, 0))
        
        def update_name(*args):
            block['name'] = name_var.get()
            self.tree.item(block_id, text=name_var.get())
            self.details_label.config(text=f"Block: {name_var.get()}")
            self._schedule_autosave()
        name_var.trace_add('write', update_name)
        
        # Topology-driven combiner controls
        topology = self.topology_var.get()
        current_row = 1
        
        if topology == 'Central Inverter':
            # Central Inverter: user specifies combiner count manually
            ttk.Label(form_frame, text="# Combiner Boxes:").grid(row=current_row, column=0, sticky='w', pady=5)
            num_cb_var = tk.StringVar(value=str(block.get('num_combiners', 1)))
            num_cb_spinbox = ttk.Spinbox(form_frame, from_=1, to=50, textvariable=num_cb_var, width=10)
            num_cb_spinbox.grid(row=current_row, column=1, sticky='w', pady=5, padx=(10, 0))
            ttk.Label(form_frame, text="(from project drawings)", font=('Helvetica', 8), foreground='gray').grid(row=current_row, column=2, sticky='w', padx=(5, 0))
            
            def update_num_cb(*args):
                try:
                    block['num_combiners'] = int(num_cb_var.get())
                except ValueError:
                    pass
                self._mark_stale()
                self._schedule_autosave()
            num_cb_var.trace_add('write', update_num_cb)
            current_row += 1
        
        if topology in ('Centralized String', 'Central Inverter'):
            # Breaker Size — relevant when combiners exist
            ttk.Label(form_frame, text="Breaker Size (A):").grid(row=current_row, column=0, sticky='w', pady=5)
            breaker_sizes = [100, 125, 150, 175, 200, 225, 250, 300, 350, 400, 450, 500, 600, 700, 800]
            breaker_var = tk.StringVar(value=str(block.get('breaker_size', 400)))
            breaker_combo = ttk.Combobox(form_frame, textvariable=breaker_var, values=breaker_sizes, state='readonly', width=10)
            breaker_combo.grid(row=current_row, column=1, sticky='w', pady=5, padx=(10, 0))
            self.disable_combobox_scroll(breaker_combo)
            
            def update_breaker(*args):
                try:
                    block['breaker_size'] = int(breaker_var.get())
                except ValueError:
                    pass
                self._mark_stale()
                self._schedule_autosave()
            breaker_var.trace_add('write', update_breaker)
            current_row += 1
            
            # Combiner Box Override
            ttk.Label(form_frame, text="Combiner Box:").grid(row=current_row, column=0, sticky='w', pady=5)
            combiner_library = self.load_combiner_library()
            cb_options = ['Auto'] + sorted(
                [pn for pn in combiner_library.keys() if not pn.startswith('_')],
                key=lambda pn: (
                    combiner_library[pn].get('max_inputs', 0),
                    combiner_library[pn].get('breaker_size', 0)
                )
            )
            cb_override_var = tk.StringVar(value=block.get('cb_override', 'Auto'))
            cb_override_combo = ttk.Combobox(form_frame, textvariable=cb_override_var, values=cb_options, state='readonly', width=25)
            cb_override_combo.grid(row=current_row, column=1, sticky='w', pady=5, padx=(10, 0))
            self.disable_combobox_scroll(cb_override_combo)

            def update_cb_override(*args):
                val = cb_override_var.get()
                block['cb_override'] = val if val != 'Auto' else None
                self._mark_stale()
                self._schedule_autosave()
            cb_override_var.trace_add('write', update_cb_override)
            current_row += 1
            
            if topology == 'Centralized String':
                ttk.Label(form_frame, text="", foreground='gray', font=('Helvetica', 8)).grid(row=current_row, column=0, columnspan=3, sticky='w', pady=2)
                # Update hint after allocation runs
                self._cb_hint_row = current_row
                self._cb_hint_frame = form_frame
        
        if topology == 'Distributed String':
            ttk.Label(form_frame, text="No combiner boxes — strings connect directly to inverters",
                      foreground='gray', font=('Helvetica', 9)).grid(row=current_row, column=0, columnspan=3, sticky='w', pady=5)

        # Tracker Configurations Section
        tracker_frame = ttk.LabelFrame(scrollable_frame, text="Tracker Configurations", padding="10")
        tracker_frame.pack(fill='x', pady=(15, 0), padx=10)
        
        # Store references for tracker rows
        self.current_block_tracker_frame = tracker_frame
        self.current_block_id = block_id
        self.current_subarray_id = subarray_id
        self.current_block = block
        
        # String count display
        count_frame = ttk.Frame(tracker_frame)
        count_frame.pack(fill='x', pady=(0, 10))
        self.string_count_label = ttk.Label(count_frame, text="Total Strings: 0", font=('Helvetica', 10, 'bold'))
        self.string_count_label.pack(side='left')
        
        # Headers
        header_frame = ttk.Frame(tracker_frame)
        header_frame.pack(fill='x')
        ttk.Label(header_frame, text="Strings/Tracker", width=14).pack(side='left', padx=2)
        ttk.Label(header_frame, text="Quantity", width=10).pack(side='left', padx=2)
        ttk.Label(header_frame, text="Harness Config", width=14).pack(side='left', padx=2)
        ttk.Label(header_frame, text="", width=8).pack(side='left', padx=2)
        
        # Container for tracker rows
        self.tracker_rows_container = ttk.Frame(tracker_frame)
        self.tracker_rows_container.pack(fill='x', pady=(5, 0))
        
        # Add existing tracker rows
        for i, tracker in enumerate(block['trackers']):
            self.add_tracker_row_ui(block, i, tracker)
        
        # Update initial string count
        self.update_string_count()
        
        # Add tracker button
        add_btn = ttk.Button(tracker_frame, text="+ Add Tracker Type", 
                            command=lambda: self.add_new_tracker_to_block(block))
        add_btn.pack(anchor='w', pady=(10, 0))

    def add_tracker_row_ui(self, block: dict, index: int, tracker: dict):
        """Add a tracker configuration row to the UI"""
        row_frame = ttk.Frame(self.tracker_rows_container)
        row_frame.pack(fill='x', pady=2)
        
        # Strings dropdown
        strings_var = tk.StringVar(value=str(tracker['strings']))
        strings_combo = ttk.Combobox(row_frame, textvariable=strings_var, 
                                     values=["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16"], 
                                     width=12, state='readonly')
        strings_combo.pack(side='left', padx=2)
        self.disable_combobox_scroll(strings_combo)
        
        # Quantity
        qty_var = tk.StringVar(value=str(tracker['quantity']))
        qty_spinbox = ttk.Spinbox(row_frame, from_=0, to=10000, textvariable=qty_var, width=8)
        qty_spinbox.pack(side='left', padx=2)
        
        # Harness config
        harness_var = tk.StringVar(value=tracker['harness_config'])
        harness_options = self.get_harness_options(tracker['strings'])
        harness_combo = ttk.Combobox(row_frame, textvariable=harness_var, 
                                     values=harness_options, width=12)
        harness_combo.pack(side='left', padx=2)
        self.disable_combobox_scroll(harness_combo)
        
        # Update harness options when strings change
        def on_strings_change(*args):
            try:
                new_strings = int(strings_var.get())
                tracker['strings'] = new_strings
                new_options = self.get_harness_options(new_strings)
                harness_combo['values'] = new_options
                if harness_var.get() not in new_options:
                    harness_var.set(new_options[0])
                    tracker['harness_config'] = new_options[0]
            except ValueError:
                pass
            self.update_string_count()
            self._mark_stale()
        strings_var.trace_add('write', on_strings_change)
        
        # Update data when values change
        def on_qty_change(*args):
            try:
                tracker['quantity'] = int(qty_var.get())
            except ValueError:
                pass
            self.update_string_count()
            self._mark_stale()
            self._schedule_autosave()
        qty_var.trace_add('write', on_qty_change)
        
        def on_harness_change(*args):
            tracker['harness_config'] = harness_var.get()
            self._mark_stale()
            self._schedule_autosave()
        harness_var.trace_add('write', on_harness_change)
        
        # Remove button
        def remove_tracker():
            if tracker in block['trackers']:
                block['trackers'].remove(tracker)
            row_frame.destroy()
            self.update_string_count()
        
        remove_btn = ttk.Button(row_frame, text="✕", width=3, command=remove_tracker)
        remove_btn.pack(side='left', padx=2)

    def add_new_tracker_to_block(self, block: dict):
        """Add a new tracker configuration to a block"""
        new_tracker = {'strings': 2, 'quantity': 0, 'harness_config': '2'}
        block['trackers'].append(new_tracker)
        self.add_tracker_row_ui(block, len(block['trackers']) - 1, new_tracker)

    # ==================== Calculation Methods ====================

    def calculate_estimate(self):
        """Calculate and display the rolled-up BOM estimate"""
        # Clear previous results
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        self.checked_items.clear()
        
        # Aggregated totals
        totals = {
            'combiners_by_breaker': {},  # {breaker_size: quantity}
            'combiner_details': [],  # list of matched CB info dicts
            'string_inverters': 0,
            'trackers_by_string': {},  # {num_strings: quantity}
            'harnesses_by_size': {},   # {num_strings: quantity}
            'whips_by_length': {},     # {length_ft: quantity}
            'extenders_short': 0,
            'extenders_long': 0,
            'total_whip_length': 0,
            'dc_feeder_total_ft': 0,
            'dc_feeder_count': 0,
            'ac_homerun_total_ft': 0,
            'ac_homerun_count': 0,
        }
        
        # Topology and strings-per-inverter (used throughout calculation)
        topology = self.topology_var.get()
        try:
            strings_per_inv = int(self.strings_per_inverter_var.get())
        except (ValueError, AttributeError):
            strings_per_inv = 0
        
        # Validate module selection
        if not self.selected_module:
            from tkinter import messagebox
            messagebox.showwarning("No Module Selected", "Please select a module in Global Settings before calculating.")
            return
        
        # Get module info from selected module
        module_isc = self.selected_module.isc
        module_width_mm = self.selected_module.width_mm
        try:
            modules_per_string = int(self.modules_per_string_var.get())
        except ValueError:
            modules_per_string = 28
        module_width_ft = module_width_mm / 304.8
        string_length_ft = module_width_ft * modules_per_string
        
        # Process each subarray and block
        for subarray_id, subarray in self.subarrays.items():
            for block_id, block in subarray['blocks'].items():
                breaker_size = block.get('breaker_size', 400)
                
                # Process trackers in this block
                block_total_strings = 0
                block_total_rows = 0
                block_total_harnesses = 0
                max_harness_strings = 0  # Largest harness group size in block
                
                for tracker in block['trackers']:
                    strings = tracker['strings']
                    qty = tracker['quantity']
                    harness_config = tracker['harness_config']
                    
                    if qty <= 0:
                        continue
                    
                    block_total_strings += qty * strings
                    block_total_rows += qty
                    
                    # Count trackers by string count
                    if strings not in totals['trackers_by_string']:
                        totals['trackers_by_string'][strings] = 0
                    totals['trackers_by_string'][strings] += qty
                    
                    # Count harnesses by size and track max harness group
                    harness_sizes = self.parse_harness_config(harness_config)
                    for size in harness_sizes:
                        if size > max_harness_strings:
                            max_harness_strings = size
                        if size not in totals['harnesses_by_size']:
                            totals['harnesses_by_size'][size] = 0
                        totals['harnesses_by_size'][size] += qty
                        block_total_harnesses += qty
                
                # --- Topology-driven device count and combiner sizing ---
                import math
                
                num_devices = 1
                num_combiners = 0
                strings_per_cb = 0
                
                if topology == 'Distributed String':
                    # Devices are inverters — no combiners
                    if strings_per_inv > 0:
                        num_devices = math.ceil(block_total_strings / strings_per_inv)
                
                elif topology == 'Centralized String':
                    # 1 combiner per inverter, auto-derived
                    if strings_per_inv > 0:
                        num_devices = math.ceil(block_total_strings / strings_per_inv)
                    num_combiners = num_devices
                    strings_per_cb = strings_per_inv
                
                elif topology == 'Central Inverter':
                    # Manual combiner count from block data
                    num_combiners = block.get('num_combiners', 1)
                    num_devices = num_combiners
                    if num_combiners > 0 and block_total_strings > 0:
                        strings_per_cb = math.ceil(block_total_strings / num_combiners)
                
                # Combiner box sizing (Centralized String and Central Inverter only)
                if num_combiners > 0 and block_total_strings > 0:
                    if breaker_size not in totals['combiners_by_breaker']:
                        totals['combiners_by_breaker'][breaker_size] = 0
                    totals['combiners_by_breaker'][breaker_size] += num_combiners
                    
                    fuse_current = module_isc * 1.56 * max(max_harness_strings, 1)
                    fuse_holder_rating = self.get_fuse_holder_category(fuse_current)
                    
                    cb_override = block.get('cb_override')
                    if cb_override:
                        combiner_library = self.load_combiner_library()
                        matched_cb = combiner_library.get(cb_override)
                    else:
                        matched_cb = self.find_combiner_box(strings_per_cb, breaker_size, fuse_holder_rating)
                    
                    if matched_cb:
                        totals['combiner_details'].append({
                            'part_number': matched_cb.get('part_number', ''),
                            'description': matched_cb.get('description', ''),
                            'max_inputs': matched_cb.get('max_inputs', 0),
                            'breaker_size': breaker_size,
                            'fuse_holder_rating': fuse_holder_rating,
                            'strings_per_cb': strings_per_cb,
                            'quantity': num_combiners,
                            'block_name': block['name']
                        })
                    else:
                        totals['combiner_details'].append({
                            'part_number': 'NO MATCH',
                            'description': f'No CB found: {strings_per_cb} inputs, {breaker_size}A, {fuse_holder_rating}',
                            'max_inputs': strings_per_cb,
                            'breaker_size': breaker_size,
                            'fuse_holder_rating': fuse_holder_rating,
                            'strings_per_cb': strings_per_cb,
                            'quantity': num_combiners,
                            'block_name': block['name']
                        })
                
                # --- Whip calculation (to nearest device — inverter or combiner) ---
                try:
                    row_spacing = float(self.row_spacing_var.get())
                except ValueError:
                    row_spacing = 20.0
                
                if block_total_rows > 0 and num_devices > 0:
                    whip_distances = self.calculate_cb_whip_distances(
                        block_total_rows, num_devices, row_spacing
                    )
                    
                    for distance_ft, device_idx in whip_distances:
                        whip_length = self.round_whip_length(distance_ft)
                        whips_at_length = 2  # pos + neg per row
                        
                        if whip_length not in totals['whips_by_length']:
                            totals['whips_by_length'][whip_length] = 0
                        totals['whips_by_length'][whip_length] += whips_at_length
                        totals['total_whip_length'] += whip_length * whips_at_length
                
                # Calculate extenders for this block (2 per harness - short and long)
                totals['extenders_short'] += block_total_harnesses
                totals['extenders_long'] += block_total_harnesses
        
        # Calculate extender lengths
        short_extender_length = self.round_whip_length(10)  # Standard 10ft short side
        long_extender_length = self.round_whip_length(string_length_ft)
        
        # ==================== Inverter Allocation ====================
        
        totals['inverter_allocation'] = None
        totals['inverter_summary'] = {}
        
        # DEBUG: Print all relevant design inputs
        print("\n" + "="*80)
        print("QUICK ESTIMATE DEBUG - INVERTER ALLOCATION INPUTS")
        print("="*80)
        
        print(f"\n--- MODULE ---")
        if self.selected_module:
            print(f"  Module: {self.selected_module.manufacturer} {self.selected_module.model}")
            print(f"  Wattage: {self.selected_module.wattage} W")
            print(f"  Isc: {self.selected_module.isc} A")
            print(f"  Imp: {self.selected_module.imp} A")
            print(f"  Voc: {self.selected_module.voc} V")
            print(f"  Vmp: {self.selected_module.vmp} V")
            print(f"  Width: {self.selected_module.width_mm} mm")
        else:
            print("  No module selected")
        
        print(f"\n--- INVERTER ---")
        if self.selected_inverter:
            inv = self.selected_inverter
            print(f"  Inverter: {inv.manufacturer} {inv.model}")
            print(f"  Type: {inv.inverter_type.value}")
            print(f"  Rated AC Power: {inv.rated_power_kw} kW")
            print(f"  Max DC Power: {inv.max_dc_power_kw} kW")
            print(f"  Max DC Voltage: {inv.max_dc_voltage} V")
            print(f"  Startup Voltage: {inv.startup_voltage} V")
            print(f"  Total String Inputs: {inv.get_total_string_capacity()}")
            print(f"  Max Short Circuit Current: {getattr(inv, 'max_short_circuit_current', None)}")
            print(f"  MPPT Channels: {len(inv.mppt_channels)}")
            for i, ch in enumerate(inv.mppt_channels):
                print(f"    Channel {i+1}: max_current={ch.max_input_current}A, voltage={ch.min_voltage}-{ch.max_voltage}V, max_power={ch.max_power}W, inputs={ch.num_string_inputs}")
        else:
            print("  No inverter selected")
        
        print(f"\n--- GLOBAL SETTINGS ---")
        print(f"  Modules per String: {modules_per_string}")
        print(f"  Topology: {self.topology_var.get()}")
        print(f"  DC:AC Ratio (target): {self.dc_ac_ratio_var.get()}")
        print(f"  Strings/Inverter (calculated): {self.strings_per_inverter_var.get()}")
        
        print(f"\n--- TRACKER TOTALS ---")
        for strings_count, qty in totals['trackers_by_string'].items():
            print(f"  {strings_count}-string trackers: {qty} units = {qty * strings_count} total strings")
        total_all_strings = sum(s * q for s, q in totals['trackers_by_string'].items())
        total_all_trackers = sum(totals['trackers_by_string'].values())
        print(f"  TOTAL: {total_all_trackers} trackers, {total_all_strings} strings")
        
        print(f"\n--- DERIVED VALUES ---")
        if self.selected_module and self.selected_inverter:
            string_power_kw = (self.selected_module.wattage * modules_per_string) / 1000
            print(f"  String Power: {string_power_kw:.2f} kW")
            print(f"  String Isc: {self.selected_module.isc} A")
            try:
                spi = int(self.strings_per_inverter_var.get())
                inv_dc_power = spi * string_power_kw
                print(f"  Inverter DC Power ({spi} strings): {inv_dc_power:.2f} kW")
                print(f"  Actual DC:AC Ratio: {inv_dc_power / inv.rated_power_kw:.3f}")
                print(f"  Total Isc ({spi} strings): {spi * self.selected_module.isc:.1f} A")
                if total_all_strings > 0:
                    num_inverters = total_all_strings / spi
                    print(f"  Inverters Needed: {total_all_strings} / {spi} = {num_inverters:.2f}")
            except ValueError:
                print(f"  Strings/Inverter is not a valid number: {self.strings_per_inverter_var.get()}")
        
        print("="*80 + "\n")
        
        if self.selected_inverter:
            try:
                strings_per_inv = int(self.strings_per_inverter_var.get())
            except (ValueError, AttributeError):
                strings_per_inv = 0
            
            if strings_per_inv > 0:
                # Run allocation per tracker type, then aggregate
                all_allocations = []
                total_inverters = 0
                total_full_inverters = 0
                total_split_trackers = 0
                total_strings_allocated = 0
                
                for strings_count, qty in totals['trackers_by_string'].items():
                    if qty > 0:
                        alloc = allocate_strings(strings_count, strings_per_inv, qty)
                        all_allocations.append({
                            'strings_per_tracker': strings_count,
                            'tracker_count': qty,
                            'allocation': alloc
                        })
                        total_inverters += alloc['summary']['total_inverters']
                        total_full_inverters += alloc['summary']['full_inverters']
                        total_split_trackers += alloc['summary']['total_split_trackers']
                        total_strings_allocated += alloc['summary']['total_strings']
                
                # Calculate actual DC:AC ratio
                module_wattage = self.selected_module.wattage
                actual_dc_ac = self.selected_inverter.dc_ac_ratio(
                    strings_per_inv, module_wattage, modules_per_string
                )
                
                totals['inverter_summary'] = {
                    'strings_per_inverter': strings_per_inv,
                    'total_inverters': total_inverters,
                    'full_inverters': total_full_inverters,
                    'total_split_trackers': total_split_trackers,
                    'total_strings': total_strings_allocated,
                    'actual_dc_ac': actual_dc_ac,
                    'allocations': all_allocations
                }
        
        # Adjust harness counts for split trackers
        self._adjust_harnesses_for_splits(totals)
        
        # --- DC Feeder and AC Homerun calculations (topology-driven) ---
        try:
            dc_feeder_avg_ft = float(self.dc_feeder_distance_var.get())
        except (ValueError, AttributeError):
            dc_feeder_avg_ft = 500.0
        try:
            ac_homerun_avg_ft = float(self.ac_homerun_distance_var.get())
        except (ValueError, AttributeError):
            ac_homerun_avg_ft = 500.0
        
        total_inverters = totals.get('inverter_summary', {}).get('total_inverters', 0)
        total_combiners = sum(totals['combiners_by_breaker'].values())
        
        if topology == 'Distributed String':
            # No DC feeders; AC homerun from each inverter
            totals['dc_feeder_count'] = 0
            totals['dc_feeder_total_ft'] = 0
            totals['ac_homerun_count'] = total_inverters
            totals['ac_homerun_total_ft'] = total_inverters * ac_homerun_avg_ft
        elif topology == 'Centralized String':
            # DC feeder from each combiner to inverter bank; short AC from bank
            totals['dc_feeder_count'] = total_combiners
            totals['dc_feeder_total_ft'] = total_combiners * dc_feeder_avg_ft
            totals['ac_homerun_count'] = total_inverters
            totals['ac_homerun_total_ft'] = total_inverters * ac_homerun_avg_ft
        elif topology == 'Central Inverter':
            # DC feeder from each combiner to central inverter; minimal AC
            totals['dc_feeder_count'] = total_combiners
            totals['dc_feeder_total_ft'] = total_combiners * dc_feeder_avg_ft
            # Central inverter: typically 1 AC connection per block/subarray
            num_subarrays = len(self.subarrays)
            totals['ac_homerun_count'] = max(num_subarrays, 1)
            totals['ac_homerun_total_ft'] = totals['ac_homerun_count'] * ac_homerun_avg_ft

        # ==================== Display Results ====================
        
        total_inverters = totals.get('inverter_summary', {}).get('total_inverters', 0)
        total_combiners = sum(totals['combiners_by_breaker'].values())

        # Store totals for Excel export
        self.last_totals = totals
        self._results_stale = False
        if self._calc_btn:
            self._calc_btn.config(state='disabled')
        
        def insert_section(label):
            self.results_tree.insert('', 'end', values=('', f'--- {label} ---', '', '', '', '', ''), tags=('section',))

        def insert_row(item, part_number, qty, unit, unit_cost='', ext_cost=''):
            iid = self.results_tree.insert('', 'end', values=('☑', item, part_number, qty, unit, unit_cost, ext_cost), tags=('checked',))
            self.checked_items.add(iid)

        # Combiner Boxes by breaker size
        if totals['combiners_by_breaker']:
            insert_section('COMBINER BOXES')
            total_cbs = 0
            for breaker_size in sorted(totals['combiners_by_breaker'].keys()):
                qty = totals['combiners_by_breaker'][breaker_size]
                total_cbs += qty
                insert_row(f"Combiner Box ({breaker_size}A breaker)", '', qty, 'ea')
            if len(totals['combiners_by_breaker']) > 1:
                insert_row('Total Combiner Boxes', '', total_cbs, 'ea')

            for detail in totals['combiner_details']:
                if detail['part_number'] != 'NO MATCH':
                    insert_row(
                        f"  └ {detail['block_name']}: ({detail['max_inputs']}-input, {detail['fuse_holder_rating']})",
                        detail['part_number'], detail['quantity'], 'ea'
                    )
                else:
                    insert_row(
                        f"  └ {detail['block_name']}: ⚠ {detail['description']}",
                        '', detail['quantity'], 'ea'
                    )

        if totals['string_inverters'] > 0:
            insert_row('String Inverters', '', totals['string_inverters'], 'ea')
            # Show allocation pattern per tracker type
            for alloc_data in inv_summary.get('allocations', []):
                cycle = alloc_data['allocation']['cycle']
                spt = alloc_data['strings_per_tracker']
                if cycle:
                    patterns = ['  '.join(f'{s}' for s in pattern) for pattern in cycle[:3]]
                    cycle_str = '  /  '.join(f"[{'-'.join(str(s) for s in p)}]" for p in cycle[:3])
                    insert_row(
                        f"  {spt}-String Trackers: {cycle_str}",
                        '', '', ''
                    )

        # Harnesses
        if totals['harnesses_by_size']:
            insert_section('HARNESSES')
            for size in sorted(totals['harnesses_by_size'].keys(), reverse=True):
                qty = totals['harnesses_by_size'][size]
                pos_pn, pos_unit, pos_ext = self.lookup_part_and_price('harness', num_strings=size, polarity='positive', qty=qty)
                neg_pn, neg_unit, neg_ext = self.lookup_part_and_price('harness', num_strings=size, polarity='negative', qty=qty)
                if size == 1:
                    iid = self.results_tree.insert('', 'end', values=('☐', f"{size}-String Harness (Pos)", pos_pn, qty, 'ea', pos_unit, pos_ext), tags=('unchecked',))
                    iid2 = self.results_tree.insert('', 'end', values=('☐', f"{size}-String Harness (Neg)", neg_pn, qty, 'ea', neg_unit, neg_ext), tags=('unchecked',))
                else:
                    insert_row(f"{size}-String Harness (Pos)", pos_pn, qty, 'ea', pos_unit, pos_ext)
                    insert_row(f"{size}-String Harness (Neg)", neg_pn, qty, 'ea', neg_unit, neg_ext)

        # Extenders
        total_extenders = totals['extenders_short'] + totals['extenders_long']
        if total_extenders > 0:
            insert_section('EXTENDERS')
            s_pn, s_unit, s_ext = self.lookup_part_and_price('extender', polarity='positive', length_ft=short_extender_length, qty=totals['extenders_short'])
            insert_row(f"Extender {short_extender_length}ft (short side, Pos)", s_pn, totals['extenders_short'], 'ea', s_unit, s_ext)
            l_pn, l_unit, l_ext = self.lookup_part_and_price('extender', polarity='negative', length_ft=long_extender_length, qty=totals['extenders_long'])
            insert_row(f"Extender {long_extender_length}ft (long side, Neg)", l_pn, totals['extenders_long'], 'ea', l_unit, l_ext)

        # Whips
        if totals['whips_by_length']:
            whip_label = 'WHIPS (to inverter)' if topology == 'Distributed String' else 'WHIPS (to combiner)'
            insert_section(whip_label)
            for length in sorted(totals['whips_by_length'].keys()):
                qty = totals['whips_by_length'][length]
                w_pn, w_unit, w_ext = self.lookup_part_and_price('whip', polarity='positive', length_ft=length, qty=qty)
                insert_row(f"Whip {length}ft", w_pn, qty, 'ea', w_unit, w_ext)
        
        # Inverters
        if total_inverters > 0:
            insert_section('INVERTERS')
            inv_name = f"{self.selected_inverter.manufacturer} {self.selected_inverter.model}" if self.selected_inverter else "Inverter"
            inv_summary = totals.get('inverter_summary', {})
            actual_ratio = inv_summary.get('actual_dc_ac', 0)
            insert_row(f"{inv_name} (DC:AC {actual_ratio:.2f})", '', total_inverters, 'ea')
        
        # DC Feeders (Centralized String and Central Inverter only)
        if totals['dc_feeder_count'] > 0:
            insert_section('DC FEEDERS')
            insert_row(f"DC Feeder — avg {dc_feeder_avg_ft:.0f}ft run × {totals['dc_feeder_count']} runs (pos)", '', f"{totals['dc_feeder_total_ft']:.0f}", 'ft')
            insert_row(f"DC Feeder — avg {dc_feeder_avg_ft:.0f}ft run × {totals['dc_feeder_count']} runs (neg)", '', f"{totals['dc_feeder_total_ft']:.0f}", 'ft')
        
        # AC Homeruns
        if totals['ac_homerun_count'] > 0:
            insert_section('AC HOMERUNS')
            insert_row(f"AC Homerun — avg {ac_homerun_avg_ft:.0f}ft run × {totals['ac_homerun_count']} runs", '', f"{totals['ac_homerun_total_ft']:.0f}", 'ft')

    def export_to_excel(self):
        """Export the quick estimate BOM to Excel"""
        from tkinter import filedialog, messagebox
        
        # Run calculation first to ensure results are current
        self.calculate_estimate()
        
        if not self.selected_module:
            messagebox.showwarning("No Module", "Please select a module before exporting.")
            return
        
        # Gather all the data we need
        try:
            modules_per_string = int(self.modules_per_string_var.get())
        except ValueError:
            modules_per_string = 28
        try:
            row_spacing = float(self.row_spacing_var.get())
        except ValueError:
            row_spacing = 20.0
        
        module_isc = self.selected_module.isc
        module_width_mm = self.selected_module.width_mm
        module_width_ft = module_width_mm / 304.8
        string_length_ft = module_width_ft * modules_per_string
        
        # Build suggested filename
        project_name = "Quick_Estimate"
        estimate_name = "Estimate"
        if self.current_project and self.current_project.metadata:
            project_name = self.current_project.metadata.name or "Quick_Estimate"
        if self.estimate_id and self.current_project:
            est_data = self.current_project.quick_estimates.get(self.estimate_id, {})
            estimate_name = est_data.get('name', 'Estimate')
        
        # Clean filename
        safe_project = "".join(c for c in project_name if c.isalnum() or c in (' ', '-', '_')).strip()
        safe_estimate = "".join(c for c in estimate_name if c.isalnum() or c in (' ', '-', '_')).strip()
        suggested_filename = f"{safe_project}_{safe_estimate}_Quick_BOM.xlsx"
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Export Quick Estimate BOM",
            initialfile=suggested_filename
        )
        
        if not filepath:
            return
        
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter
            
            wb = Workbook()
            ws = wb.active
            ws.title = "Quick Estimate BOM"
            
            # Define styles
            title_font = Font(bold=True, size=14)
            header_font = Font(bold=True, size=11, color="FFFFFF")
            header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
            section_font = Font(bold=True, size=11)
            section_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
            label_font = Font(bold=True)
            thin_border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
            center_align = Alignment(horizontal='center', vertical='center')
            wrap_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
            warning_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
            
            row = 1
            
            # ========== PROJECT INFO SECTION ==========
            ws.merge_cells(f'A{row}:E{row}')
            ws.cell(row=row, column=1, value="Quick Estimate BOM").font = title_font
            row += 2
            
            # Project info pairs
            info_items = []
            if self.current_project and self.current_project.metadata:
                meta = self.current_project.metadata
                if meta.name:
                    info_items.append(("Project Name:", meta.name))
                if meta.client:
                    info_items.append(("Customer:", meta.client))
                if meta.location:
                    info_items.append(("Location:", meta.location))
            
            info_items.append(("Module:", f"{self.selected_module.manufacturer} {self.selected_module.model} ({self.selected_module.wattage}W)"))
            info_items.append(("Module Isc:", f"{module_isc} A"))
            info_items.append(("Module Width:", f"{module_width_mm} mm"))
            info_items.append(("Modules per String:", str(modules_per_string)))
            info_items.append(("Row Spacing:", f"{row_spacing} ft"))
            
            if self.selected_inverter:
                inv = self.selected_inverter
                info_items.append(("Inverter:", f"{inv.manufacturer} {inv.model} ({inv.rated_power_kw}kW AC)"))
                info_items.append(("Topology:", self.topology_var.get()))
                info_items.append(("DC:AC Ratio (target):", self.dc_ac_ratio_var.get()))
                if hasattr(self, 'last_totals') and self.last_totals.get('inverter_summary'):
                    inv_sum = self.last_totals['inverter_summary']
                    info_items.append(("DC:AC Ratio (actual):", f"{inv_sum.get('actual_dc_ac', 0):.2f}"))
                    info_items.append(("Strings per Inverter:", str(inv_sum.get('strings_per_inverter', ''))))
                    info_items.append(("Total Inverters:", str(inv_sum.get('total_inverters', ''))))
                    info_items.append(("Split Trackers:", str(inv_sum.get('total_split_trackers', ''))))
            
            if self.estimate_id and self.current_project:
                est_data = self.current_project.quick_estimates.get(self.estimate_id, {})
                info_items.append(("Estimate:", est_data.get('name', '')))
            
            for label, value in info_items:
                ws.cell(row=row, column=1, value=label).font = label_font
                ws.cell(row=row, column=2, value=value)
                row += 1
            
            row += 1
            
            # ========== BLOCK SUMMARY SECTION ==========
            ws.merge_cells(f'A{row}:E{row}')
            cell = ws.cell(row=row, column=1, value="Block Configuration Summary")
            cell.font = title_font
            row += 1
            
            # Block summary headers (topology-aware)
            topology = self.topology_var.get()
            if topology == 'Distributed String':
                block_headers = ['Subarray', 'Block', 'Tracker Configs', 'Total Strings', 'Total Trackers']
            elif topology == 'Centralized String':
                block_headers = ['Subarray', 'Block', 'Breaker (A)', 'Tracker Configs', 'Total Strings', 'Total Trackers']
            else:  # Central Inverter
                block_headers = ['Subarray', 'Block', '# CBs', 'Breaker (A)', 'Tracker Configs', 'Total Strings', 'Total Trackers']
            for col, header in enumerate(block_headers, 1):
                cell = ws.cell(row=row, column=col, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = center_align
                cell.border = thin_border
            row += 1
            
            # Block data rows
            for subarray_id, subarray in self.subarrays.items():
                for block_id, block in subarray['blocks'].items():
                    block_strings = sum(t['quantity'] * t['strings'] for t in block['trackers'] if t['quantity'] > 0)
                    block_trackers = sum(t['quantity'] for t in block['trackers'] if t['quantity'] > 0)
                    
                    tracker_summary = ", ".join(
                        f"{t['quantity']}x {t['strings']}S ({t['harness_config']})"
                        for t in block['trackers'] if t['quantity'] > 0
                    )
                    
                    if topology == 'Distributed String':
                        row_data = [
                            subarray['name'],
                            block['name'],
                            tracker_summary,
                            block_strings,
                            block_trackers
                        ]
                    elif topology == 'Centralized String':
                        row_data = [
                            subarray['name'],
                            block['name'],
                            block.get('breaker_size', 400),
                            tracker_summary,
                            block_strings,
                            block_trackers
                        ]
                    else:  # Central Inverter
                        row_data = [
                            subarray['name'],
                            block['name'],
                            block.get('num_combiners', 1),
                            block.get('breaker_size', 400),
                            tracker_summary,
                            block_strings,
                            block_trackers
                        ]
                    for col, value in enumerate(row_data, 1):
                        cell = ws.cell(row=row, column=col, value=value)
                        cell.border = thin_border
                        cell.alignment = center_align
                    row += 1
            
            row += 1
            
            # ========== TRACKER SUMMARY SECTION ==========
            if hasattr(self, 'last_totals') and self.last_totals.get('trackers_by_string'):
                ws.merge_cells(f'A{row}:E{row}')
                ws.cell(row=row, column=1, value="Tracker Summary").font = title_font
                row += 1
                
                tracker_headers = ['Tracker Type', 'Quantity', 'Unit']
                for col, header in enumerate(tracker_headers, 1):
                    cell = ws.cell(row=row, column=col, value=header)
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = center_align
                    cell.border = thin_border
                row += 1
                
                total_trackers = 0
                for strings in sorted(self.last_totals['trackers_by_string'].keys()):
                    qty = self.last_totals['trackers_by_string'][strings]
                    total_trackers += qty
                    ws.cell(row=row, column=1, value=f"{strings}-String Trackers").border = thin_border
                    cell_qty = ws.cell(row=row, column=2, value=qty)
                    cell_qty.border = thin_border
                    cell_qty.alignment = center_align
                    cell_unit = ws.cell(row=row, column=3, value='ea')
                    cell_unit.border = thin_border
                    cell_unit.alignment = center_align
                    row += 1
                
                # Total row
                total_label = ws.cell(row=row, column=1, value="Total Trackers")
                total_label.font = label_font
                total_label.border = thin_border
                total_qty = ws.cell(row=row, column=2, value=total_trackers)
                total_qty.font = label_font
                total_qty.border = thin_border
                total_qty.alignment = center_align
                total_unit = ws.cell(row=row, column=3, value='ea')
                total_unit.border = thin_border
                total_unit.alignment = center_align
                row += 2
            
            # ========== INVERTER ALLOCATION SECTION ==========
            if hasattr(self, 'last_totals') and self.last_totals.get('inverter_summary'):
                inv_sum = self.last_totals['inverter_summary']
                
                ws.merge_cells(f'A{row}:E{row}')
                ws.cell(row=row, column=1, value="Inverter Allocation Summary").font = title_font
                row += 1
                
                alloc_headers = ['Inverter', 'Strings', 'Trackers', 'Pattern']
                for col, header in enumerate(alloc_headers, 1):
                    cell = ws.cell(row=row, column=col, value=header)
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = center_align
                    cell.border = thin_border
                row += 1
                
                global_inv_idx = 0
                for alloc_data in inv_sum.get('allocations', []):
                    allocation = alloc_data['allocation']
                    spt = alloc_data['strings_per_tracker']
                    
                    for inv in allocation['inverters']:
                        pattern_str = '-'.join(str(s) for s in inv['pattern'])
                        
                        inv_row = [
                            f"Inverter {global_inv_idx + 1}",
                            inv['total_strings'],
                            len(inv['tracker_indices']),
                            f"[{pattern_str}]"
                        ]
                        for col, value in enumerate(inv_row, 1):
                            cell = ws.cell(row=row, column=col, value=value)
                            cell.border = thin_border
                            cell.alignment = center_align
                        row += 1
                        global_inv_idx += 1
                
                row += 1

            # ========== BOM RESULTS SECTION ==========
            ws.merge_cells(f'A{row}:F{row}')
            ws.cell(row=row, column=1, value="Estimated Bill of Materials").font = title_font
            row += 1
            
            # BOM headers
            bom_headers = ['Item', 'Part Number', 'Quantity', 'Unit', 'Unit Cost', 'Ext. Cost']
            for col, header in enumerate(bom_headers, 1):
                cell = ws.cell(row=row, column=col, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = center_align
                cell.border = thin_border
            row += 1
            
            # Pull results from the results treeview - only checked items
            bom_first_data_row = None
            for item_id in self.results_tree.get_children():
                values = self.results_tree.item(item_id, 'values')
                if len(values) < 7:
                    continue

                include = values[0]   # ☑ or ☐ or ''
                item_name = values[1]
                part_number = values[2]
                qty = values[3]
                unit = values[4]
                unit_cost = values[5]

                is_section = str(item_name).startswith('---')
                is_warning = '⚠' in str(item_name)

                # Skip unchecked non-section rows
                if not is_section and include != '☑':
                    continue

                cell_item = ws.cell(row=row, column=1, value=str(item_name).replace('---', '').strip() if is_section else item_name)
                cell_pn = ws.cell(row=row, column=2, value=part_number if not is_section else '')
                cell_qty = ws.cell(row=row, column=3, value=qty if qty else '')
                cell_unit = ws.cell(row=row, column=4, value=unit if unit else '')
                cell_unit_cost = ws.cell(row=row, column=5, value=unit_cost if unit_cost else '')

                # Ext. Cost: formula for all non-section rows
                if not is_section:
                    cell_ext_cost = ws.cell(row=row, column=6, value=f'=IF(E{row}="","",C{row}*E{row})')
                    cell_ext_cost.number_format = '"$"#,##0.00'
                    if bom_first_data_row is None:
                        bom_first_data_row = row
                else:
                    cell_ext_cost = ws.cell(row=row, column=6, value='')

                if is_section:
                    for c in [cell_item, cell_pn, cell_qty, cell_unit, cell_unit_cost, cell_ext_cost]:
                        c.font = section_font
                        c.fill = section_fill
                elif is_warning:
                    for c in [cell_item, cell_pn, cell_qty, cell_unit, cell_unit_cost, cell_ext_cost]:
                        c.fill = warning_fill

                for c in [cell_item, cell_pn, cell_qty, cell_unit, cell_unit_cost, cell_ext_cost]:
                    c.border = thin_border
                    c.alignment = center_align
                cell_item.alignment = wrap_align

                row += 1

            # Total Cost row
            if bom_first_data_row:
                row += 1
                total_label = ws.cell(row=row, column=5, value='Total Cost:')
                total_label.font = Font(bold=True)
                total_label.alignment = Alignment(horizontal='right')
                total_label.border = thin_border
                total_cell = ws.cell(row=row, column=6, value=f'=SUM(F{bom_first_data_row}:F{row - 2})')
                total_cell.font = Font(bold=True)
                total_cell.number_format = '"$"#,##0.00'
                total_cell.border = thin_border
                total_cell.alignment = center_align
                row += 1
            
            # ========== AUTO-FIT COLUMNS ==========
            for col_idx in range(1, 9):
                max_length = 0
                col_letter = get_column_letter(col_idx)
                for cell in ws[col_letter]:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = min(max_length + 4, 55)
            
            # Save
            wb.save(filepath)
            
            # Try to open the file
            import os
            try:
                os.startfile(filepath)
            except Exception:
                pass
            
            messagebox.showinfo("Success", f"Quick Estimate BOM exported to:\n{filepath}")
            
        except PermissionError:
            messagebox.showerror(
                "Permission Error",
                f"Cannot write to {filepath}.\n\n"
                "The file may be open in Excel. Please close it and try again."
            )
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export BOM:\n{str(e)}")

class QuickEstimateDialog(tk.Toplevel):
    """Dialog wrapper for Quick Estimate tool"""
    
    def __init__(self, parent, current_project=None, estimate_id=None, on_save=None):
        super().__init__(parent)
        self.title("Quick Estimate")
        self.current_project = current_project
        self.estimate_id = estimate_id
        self.on_save = on_save
        
        # Set dialog size
        self.geometry("1100x850")
        self.minsize(900, 700)
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
        # Create the Quick Estimate frame inside the dialog
        self.quick_estimate = QuickEstimate(
            self, 
            current_project=current_project,
            estimate_id=estimate_id,
            on_save=self._handle_save
        )
        self.quick_estimate.pack(fill='both', expand=True)
        
        # Add button frame at bottom
        button_frame = ttk.Frame(self)
        button_frame.pack(fill='x', padx=10, pady=10)
        
        # Save button
        save_btn = ttk.Button(button_frame, text="Save", command=self.save_and_close)
        save_btn.pack(side='right', padx=(5, 0))
        
        # Close button (save on close)
        close_btn = ttk.Button(button_frame, text="Close", command=self.save_and_close)
        close_btn.pack(side='right')
        
        # Center the dialog on the parent window
        self.center_on_parent(parent)
        
        # Focus on the dialog
        self.focus_set()
        
        # Handle window close button (X)
        self.protocol("WM_DELETE_WINDOW", self.save_and_close)
        
        # Wait for window to close before returning
        self.wait_window(self)
    
    def _handle_save(self):
        """Internal save handler"""
        if self.on_save:
            self.on_save()
    
    def save_and_close(self):
        """Save the estimate and close the dialog"""
        self.quick_estimate.save_estimate()
        self.destroy()
    
    def center_on_parent(self, parent):
        """Center the dialog on the parent window"""
        self.update_idletasks()
        
        # Get parent geometry
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        
        # Get dialog size
        dialog_width = self.winfo_width()
        dialog_height = self.winfo_height()
        
        # Calculate position
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        # Ensure dialog is on screen
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        x = max(0, min(x, screen_width - dialog_width))
        y = max(0, min(y, screen_height - dialog_height))
        
        self.geometry(f"+{x}+{y}")

class SitePreviewWindow(tk.Toplevel):
    """Pop-out window for site layout preview with zoom and pan"""
    
    def __init__(self, parent, inv_summary, topology, colors):
        super().__init__(parent)
        self.title("Site Preview — Inverter Allocation")
        self.geometry("1100x750")
        self.minsize(600, 400)
        
        self.inv_summary = inv_summary
        self.topology = topology
        self.colors = colors
        
        # Zoom and pan state
        self.scale = 1.0
        self.pan_x = 0
        self.pan_y = 0
        self.dragging = False
        self.drag_start_x = 0
        self.drag_start_y = 0
        
        self.setup_ui()
        self.build_layout_data()
        self.after(50, self.fit_and_redraw)
    
    def setup_ui(self):
        """Create the preview window UI"""
        # Top bar with controls
        top_bar = ttk.Frame(self, padding="5")
        top_bar.pack(fill='x')
        
        ttk.Button(top_bar, text="Fit to Window", command=self.fit_and_redraw).pack(side='left', padx=2)
        ttk.Button(top_bar, text="Zoom In", command=lambda: self.zoom(1.3)).pack(side='left', padx=2)
        ttk.Button(top_bar, text="Zoom Out", command=lambda: self.zoom(0.7)).pack(side='left', padx=2)
        
        self.zoom_label = ttk.Label(top_bar, text="100%")
        self.zoom_label.pack(side='left', padx=10)
        
        # Summary info
        num_inv = self.inv_summary.get('total_inverters', 0)
        total_str = self.inv_summary.get('total_strings', 0)
        actual_ratio = self.inv_summary.get('actual_dc_ac', 0)
        split = self.inv_summary.get('total_split_trackers', 0)
        
        summary_text = f"{num_inv} Inverters  |  {total_str} Strings  |  DC:AC: {actual_ratio:.2f}  |  {split} Split Trackers  |  {self.topology}"
        ttk.Label(top_bar, text=summary_text, foreground='#333333').pack(side='right', padx=10)
        
        # Canvas
        canvas_frame = ttk.Frame(self)
        canvas_frame.pack(fill='both', expand=True)
        
        self.canvas = tk.Canvas(canvas_frame, bg='white', highlightthickness=0)
        self.canvas.pack(fill='both', expand=True)
        
        # Bind events
        self.canvas.bind('<MouseWheel>', self.on_mousewheel)
        self.canvas.bind('<Button-4>', lambda e: self.zoom(1.1))
        self.canvas.bind('<Button-5>', lambda e: self.zoom(0.9))
        self.canvas.bind('<ButtonPress-1>', self.on_drag_start)
        self.canvas.bind('<B1-Motion>', self.on_drag)
        self.canvas.bind('<ButtonRelease-1>', self.on_drag_end)
        self.canvas.bind('<Configure>', lambda e: self.draw())
        
        # Bottom legend
        legend_frame = ttk.Frame(self, padding="5")
        legend_frame.pack(fill='x')
        
        # Color swatches row
        swatch_frame = ttk.Frame(legend_frame)
        swatch_frame.pack(anchor='w')
        
        num_inv = self.inv_summary.get('total_inverters', 0)
        max_show = min(num_inv, 15)
        for i in range(max_show):
            color = self.colors[i % len(self.colors)]
            swatch = tk.Canvas(swatch_frame, width=12, height=12, highlightthickness=0)
            swatch.create_rectangle(0, 0, 12, 12, fill=color, outline='#333333')
            swatch.pack(side='left', padx=(0, 2))
            ttk.Label(swatch_frame, text=f"Inv {i+1}", font=('Helvetica', 8)).pack(side='left', padx=(0, 8))
        if num_inv > max_show:
            ttk.Label(swatch_frame, text=f"... +{num_inv - max_show} more",
                     font=('Helvetica', 8, 'italic'), foreground='gray').pack(side='left', padx=(5, 0))
        
        # Cycle patterns
        for alloc_data in self.inv_summary.get('allocations', []):
            cycle = alloc_data['allocation']['cycle']
            spt = alloc_data['strings_per_tracker']
            if cycle:
                cycle_str = '  →  '.join(f"[{'-'.join(str(s) for s in p)}]" for p in cycle)
                ttk.Label(legend_frame, text=f"{spt}-String Tracker Pattern: {cycle_str}",
                         font=('Helvetica', 9), foreground='#555555').pack(anchor='w')
    
    def build_layout_data(self):
        """Build a flat list of trackers in order with their inverter color assignments"""
        self.tracker_list = []
        
        for alloc_data in self.inv_summary['allocations']:
            allocation = alloc_data['allocation']
            spt = alloc_data['strings_per_tracker']
            
            # First pass: build a dict of tracker_idx -> list of assignments
            tracker_map = {}
            global_inv_idx = 0
            
            # Count previous allocations to get correct global inverter index
            for prev_alloc in self.inv_summary['allocations']:
                if prev_alloc is alloc_data:
                    break
                global_inv_idx += len(prev_alloc['allocation']['inverters'])
            
            for inv in allocation['inverters']:
                color = self.colors[global_inv_idx % len(self.colors)]
                
                for tracker_idx, strings_taken in inv['tracker_indices']:
                    if tracker_idx not in tracker_map:
                        tracker_map[tracker_idx] = {
                            'strings_per_tracker': spt,
                            'assignments': []
                        }
                    tracker_map[tracker_idx]['assignments'].append({
                        'color': color,
                        'strings': strings_taken,
                        'inv_idx': global_inv_idx
                    })
                
                global_inv_idx += 1
            
            # Convert to ordered list
            for t_idx in sorted(tracker_map.keys()):
                self.tracker_list.append(tracker_map[t_idx])
        
        # World-space dimensions
        # Each tracker is a vertical bar
        self.tracker_w = 8     # Width of each tracker column (E-W)
        self.row_gap = 6       # Gap between tracker columns (row spacing)
        self.string_h = 30     # Height per string (N-S)
        self.string_gap = 2    # Gap between strings within a tracker
        
        # Calculate world bounds
        total_trackers = len(self.tracker_list)
        if total_trackers == 0:
            self.world_width = 0
            self.world_height = 0
            return
        
        max_strings = max(t['strings_per_tracker'] for t in self.tracker_list)
        
        self.world_width = total_trackers * (self.tracker_w + self.row_gap) - self.row_gap
        self.world_height = max_strings * (self.string_h + self.string_gap) - self.string_gap
    
    def fit_to_canvas(self):
        """Calculate scale and pan to fit all content"""
        self.canvas.update_idletasks()
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        
        if cw < 10 or ch < 10:
            cw = 1100
            ch = 750
        
        margin = 40
        
        if self.world_width <= 0 or self.world_height <= 0:
            return
        
        scale_x = (cw - 2 * margin) / self.world_width
        scale_y = (ch - 2 * margin) / self.world_height
        self.scale = min(scale_x, scale_y)
        
        scaled_w = self.world_width * self.scale
        scaled_h = self.world_height * self.scale
        self.pan_x = (cw - scaled_w) / 2
        self.pan_y = (ch - scaled_h) / 2
    
    def fit_and_redraw(self):
        """Fit to window and redraw"""
        self.fit_to_canvas()
        self.draw()
    
    def world_to_canvas(self, wx, wy):
        """Convert world coordinates to canvas coordinates"""
        cx = self.pan_x + wx * self.scale
        cy = self.pan_y + wy * self.scale
        return cx, cy
    
    def zoom(self, factor):
        """Zoom in/out centered on the canvas"""
        self.canvas.update_idletasks()
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        
        center_x = cw / 2
        center_y = ch / 2
        
        self.pan_x = center_x - (center_x - self.pan_x) * factor
        self.pan_y = center_y - (center_y - self.pan_y) * factor
        self.scale *= factor
        
        self.zoom_label.config(text=f"{self.scale * 100:.0f}%")
        self.draw()
    
    def on_mousewheel(self, event):
        """Handle mouse wheel zoom"""
        if event.delta > 0:
            self.zoom(1.1)
        else:
            self.zoom(0.9)
    
    def on_drag_start(self, event):
        """Start panning"""
        self.dragging = True
        self.drag_start_x = event.x
        self.drag_start_y = event.y
    
    def on_drag(self, event):
        """Handle pan dragging"""
        if self.dragging:
            dx = event.x - self.drag_start_x
            dy = event.y - self.drag_start_y
            self.pan_x += dx
            self.pan_y += dy
            self.drag_start_x = event.x
            self.drag_start_y = event.y
            self.draw()
    
    def on_drag_end(self, event):
        """End panning"""
        self.dragging = False
    
    def draw(self):
        """Draw the site layout on the canvas"""
        self.canvas.delete('all')
        
        if not self.tracker_list:
            return
        
        # Draw each tracker as a vertical column of strings
        for i, tracker in enumerate(self.tracker_list):
            spt = tracker['strings_per_tracker']
            assignments = tracker['assignments']
            
            # X position for this tracker column
            wx = i * (self.tracker_w + self.row_gap)
            
            # Build a list of string colors from top to bottom
            string_colors = []
            for assignment in assignments:
                for _ in range(assignment['strings']):
                    string_colors.append(assignment['color'])
            
            # Draw each string as a segment of the vertical bar
            for s_idx in range(spt):
                if s_idx < len(string_colors):
                    color = string_colors[s_idx]
                else:
                    color = '#D0D0D0'
                
                wy = s_idx * (self.string_h + self.string_gap)
                
                sx1, sy1 = self.world_to_canvas(wx, wy)
                sx2, sy2 = self.world_to_canvas(wx + self.tracker_w, wy + self.string_h)
                
                self.canvas.create_rectangle(
                    sx1, sy1, sx2, sy2,
                    fill=color, outline='#444444', width=1
                )
            
            # Draw tracker outline around all strings
            ty1_world = 0
            ty2_world = spt * (self.string_h + self.string_gap) - self.string_gap
            
            ox1, oy1 = self.world_to_canvas(wx - 1, ty1_world - 1)
            ox2, oy2 = self.world_to_canvas(wx + self.tracker_w + 1, ty2_world + 1)
            
            self.canvas.create_rectangle(
                ox1, oy1, ox2, oy2,
                fill='', outline='#222222', width=1
            )
            
            # Tracker label below
            label_x, label_y = self.world_to_canvas(
                wx + self.tracker_w / 2,
                ty2_world + 5
            )
            
            bar_width = abs(ox2 - ox1)
            if bar_width > 12:
                font_size = max(6, min(9, int(8 * self.scale)))
                self.canvas.create_text(
                    label_x, label_y,
                    text=f"T{i+1}", font=('Helvetica', font_size), fill='#555555'
                )
        
        # Draw compass indicator in top-right
        self.canvas.update_idletasks()
        cw = self.canvas.winfo_width()
        compass_x = cw - 30
        compass_y = 30
        arrow_len = 18
        
        self.canvas.create_line(
            compass_x, compass_y + arrow_len,
            compass_x, compass_y - arrow_len,
            fill='#333333', width=2, arrow='last'
        )
        self.canvas.create_text(
            compass_x, compass_y - arrow_len - 8,
            text='N', font=('Helvetica', 9, 'bold'), fill='#333333'
        )