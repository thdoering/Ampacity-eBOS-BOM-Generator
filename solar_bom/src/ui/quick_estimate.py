import tkinter as tk
from tkinter import ttk
from typing import Optional, Dict, List, Any
from pathlib import Path
import json
import uuid
from datetime import datetime


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
        
        # Data structure for subarrays and blocks
        self.subarrays: Dict[str, Dict[str, Any]] = {}
        
        # Global settings defaults
        self.module_width_default = 1134
        self.modules_per_string_default = 28
        self.row_spacing_default = 20.0
        
        # Track currently selected item
        self.selected_item_id = None
        self.selected_item_type = None  # 'subarray' or 'block'
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
            return "20A and below"
        elif fuse_current_amps <= 32:
            return "25-32A"
        else:
            return "32A and above"

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
            return ["3", "2+1"]
        elif num_strings == 4:
            return ["4", "3+1", "2+2"]
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
        
        estimate_data = self.current_project.quick_estimates.get(self.estimate_id)
        if not estimate_data:
            self.add_subarray(auto_add_block=True)
            return
        
        # Load global settings
        self.modules_per_string_default = estimate_data.get('modules_per_string', 28)
        self.row_spacing_default = estimate_data.get('row_spacing_ft', 20.0)
        
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
        
        # Load subarrays
        saved_subarrays = estimate_data.get('subarrays', {})
        
        if not saved_subarrays:
            # No saved data, create default
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
                    'trackers': block_data.get('trackers', [{'strings': 2, 'quantity': 1, 'harness_config': '2'}])
                }
                
                # Add block to tree
                self.tree.insert(subarray_id, 'end', block_id, text=block_data.get('name', 'Block'))
        
        # Select first item
        first_items = self.tree.get_children()
        if first_items:
            self.tree.selection_set(first_items[0])
            self.on_tree_select(None)

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

    def _schedule_autosave(self):
        """Debounced auto-save — saves estimate after a brief pause"""
        if hasattr(self, '_autosave_after_id') and self._autosave_after_id:
            self.after_cancel(self._autosave_after_id)
        self._autosave_after_id = self.after(1000, self._do_autosave)
    
    def _do_autosave(self):
        """Execute the debounced save"""
        self._autosave_after_id = None
        self.save_estimate()

    def _on_module_selected(self, event=None):
        """Handle module selection from dropdown"""
        selected_name = self.module_select_var.get()
        if selected_name in self.available_modules:
            self.selected_module = self.available_modules[selected_name]
            self.module_info_label.config(
                text=f"Isc: {self.selected_module.isc}A  |  Width: {self.selected_module.width_mm}mm  |  Voc: {self.selected_module.voc}V",
                foreground='black'
            )
            # Auto-save when module changes
            self.save_estimate()
        else:
            self.selected_module = None
            self.module_info_label.config(text="No module selected", foreground='gray')

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
        
        # Row 2: Other settings
        settings_inner = ttk.Frame(settings_frame)
        settings_inner.pack(fill='x')
        
        ttk.Label(settings_inner, text="Modules per String:").pack(side='left', padx=(0, 5))
        self.modules_per_string_var = tk.StringVar(value=str(getattr(self, 'modules_per_string_default', 28)))
        ttk.Spinbox(settings_inner, from_=1, to=100, textvariable=self.modules_per_string_var, width=6).pack(side='left', padx=(0, 15))
        self.modules_per_string_var.trace_add('write', lambda *args: self._schedule_autosave())
        
        ttk.Label(settings_inner, text="Row Spacing (ft):").pack(side='left', padx=(0, 5))
        self.row_spacing_var = tk.StringVar(value=str(getattr(self, 'row_spacing_default', 20.0)))
        ttk.Spinbox(settings_inner, from_=1, to=100, textvariable=self.row_spacing_var, width=6).pack(side='left')
        self.row_spacing_var.trace_add('write', lambda *args: self._schedule_autosave())
        
        # Button row
        button_row = ttk.Frame(bottom_frame)
        button_row.pack(fill='x', pady=(0, 10))
        
        calc_btn = ttk.Button(button_row, text="Calculate Estimate", command=self.calculate_estimate)
        calc_btn.pack(side='left', padx=(0, 10))
        
        export_btn = ttk.Button(button_row, text="Export to Excel", command=self.export_to_excel)
        export_btn.pack(side='left')
        
        # Results frame
        results_frame = ttk.LabelFrame(bottom_frame, text="Estimated BOM (Rolled-Up Totals)", padding="10")
        results_frame.pack(fill='both', expand=True)
        
        # Results treeview
        columns = ('item', 'quantity', 'unit')
        self.results_tree = ttk.Treeview(results_frame, columns=columns, show='headings', height=8)
        self.results_tree.heading('item', text='Item')
        self.results_tree.heading('quantity', text='Quantity')
        self.results_tree.heading('unit', text='Unit')
        self.results_tree.column('item', width=250, anchor='w')
        self.results_tree.column('quantity', width=100, anchor='center')
        self.results_tree.column('unit', width=60, anchor='center')
        
        # Scrollbar for results
        results_scroll = ttk.Scrollbar(results_frame, orient='vertical', command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=results_scroll.set)
        
        self.results_tree.pack(side='left', fill='both', expand=True)
        results_scroll.pack(side='right', fill='y')

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
        
        # Block Type
        ttk.Label(form_frame, text="Type:").grid(row=1, column=0, sticky='w', pady=5)
        type_var = tk.StringVar(value=block['type'])
        type_combo = ttk.Combobox(form_frame, textvariable=type_var, values=['combiner', 'string_inverter'], state='readonly', width=15)
        type_combo.grid(row=1, column=1, sticky='w', pady=5, padx=(10, 0))
        self.disable_combobox_scroll(type_combo)
        
        def update_type(*args):
            block['type'] = type_var.get()
            self._schedule_autosave()
        type_var.trace_add('write', update_type)
        
        # Number of Combiner Boxes
        ttk.Label(form_frame, text="# Combiner Boxes:").grid(row=2, column=0, sticky='w', pady=5)
        num_cb_var = tk.StringVar(value=str(block.get('num_combiners', 1)))
        num_cb_spinbox = ttk.Spinbox(form_frame, from_=1, to=20, textvariable=num_cb_var, width=10)
        num_cb_spinbox.grid(row=2, column=1, sticky='w', pady=5, padx=(10, 0))
        ttk.Label(form_frame, text="(evenly spaced among tracker rows)", font=('Helvetica', 8), foreground='gray').grid(row=2, column=2, sticky='w', padx=(5, 0))
        
        def update_num_cb(*args):
            try:
                block['num_combiners'] = int(num_cb_var.get())
            except ValueError:
                pass
            self._schedule_autosave()
        num_cb_var.trace_add('write', update_num_cb)
        
        # Breaker Size
        ttk.Label(form_frame, text="Breaker Size (A):").grid(row=3, column=0, sticky='w', pady=5)
        breaker_sizes = [100, 125, 150, 175, 200, 225, 250, 300, 350, 400, 450, 500, 600, 700, 800]
        breaker_var = tk.StringVar(value=str(block.get('breaker_size', 400)))
        breaker_combo = ttk.Combobox(form_frame, textvariable=breaker_var, values=breaker_sizes, state='readonly', width=10)
        breaker_combo.grid(row=3, column=1, sticky='w', pady=5, padx=(10, 0))
        self.disable_combobox_scroll(breaker_combo)
        
        def update_breaker(*args):
            try:
                block['breaker_size'] = int(breaker_var.get())
            except ValueError:
                pass
            self._schedule_autosave()
        breaker_var.trace_add('write', update_breaker)
        
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
        strings_var.trace_add('write', on_strings_change)
        
        # Update data when values change
        def on_qty_change(*args):
            try:
                tracker['quantity'] = int(qty_var.get())
            except ValueError:
                pass
            self.update_string_count()
            self._schedule_autosave()
        qty_var.trace_add('write', on_qty_change)
        
        def on_harness_change(*args):
            tracker['harness_config'] = harness_var.get()
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
        }
        
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
                num_combiners = block.get('num_combiners', 1)
                breaker_size = block.get('breaker_size', 400)
                
                # Count block device types
                if block['type'] == 'combiner':
                    if breaker_size not in totals['combiners_by_breaker']:
                        totals['combiners_by_breaker'][breaker_size] = 0
                    totals['combiners_by_breaker'][breaker_size] += num_combiners
                else:
                    totals['string_inverters'] += 1
                
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
                
                # --- Combiner box sizing ---
                if block['type'] == 'combiner' and block_total_strings > 0:
                    import math
                    strings_per_cb = math.ceil(block_total_strings / num_combiners)
                    
                    # Calculate fuse current: Isc * 1.56 * max harness strings
                    fuse_current = module_isc * 1.56 * max(max_harness_strings, 1)
                    fuse_holder_rating = self.get_fuse_holder_category(fuse_current)
                    
                    # Try to find matching combiner from library
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
                
                # --- Whip calculation with CB placement ---
                try:
                    row_spacing = float(self.row_spacing_var.get())
                except ValueError:
                    row_spacing = 20.0
                
                if block_total_rows > 0:
                    cb_count = num_combiners if block['type'] == 'combiner' else 1
                    whip_distances = self.calculate_cb_whip_distances(
                        block_total_rows, cb_count, row_spacing
                    )
                    
                    for distance_ft, cb_idx in whip_distances:
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
        
        # ==================== Display Results ====================
        
        # Combiner Boxes by breaker size
        if totals['combiners_by_breaker']:
            self.results_tree.insert('', 'end', values=('--- COMBINER BOXES ---', '', ''))
            total_cbs = 0
            for breaker_size in sorted(totals['combiners_by_breaker'].keys()):
                qty = totals['combiners_by_breaker'][breaker_size]
                total_cbs += qty
                self.results_tree.insert('', 'end', values=(
                    f"Combiner Box ({breaker_size}A breaker)", qty, 'ea'
                ))
            if len(totals['combiners_by_breaker']) > 1:
                self.results_tree.insert('', 'end', values=('Total Combiner Boxes', total_cbs, 'ea'))
            
            # Show matched part numbers
            for detail in totals['combiner_details']:
                if detail['part_number'] != 'NO MATCH':
                    self.results_tree.insert('', 'end', values=(
                        f"  └ {detail['block_name']}: {detail['part_number']} ({detail['max_inputs']}-input, {detail['fuse_holder_rating']})",
                        detail['quantity'], 'ea'
                    ))
                else:
                    self.results_tree.insert('', 'end', values=(
                        f"  └ {detail['block_name']}: ⚠ {detail['description']}",
                        detail['quantity'], 'ea'
                    ))
        
        if totals['string_inverters'] > 0:
            self.results_tree.insert('', 'end', values=('String Inverters', totals['string_inverters'], 'ea'))
        
        # Trackers
        if totals['trackers_by_string']:
            self.results_tree.insert('', 'end', values=('--- TRACKERS ---', '', ''))
            total_trackers = 0
            for strings in sorted(totals['trackers_by_string'].keys()):
                qty = totals['trackers_by_string'][strings]
                total_trackers += qty
                self.results_tree.insert('', 'end', values=(f"{strings}-String Trackers", qty, 'ea'))
            self.results_tree.insert('', 'end', values=('Total Trackers', total_trackers, 'ea'))
        
        # Harnesses
        if totals['harnesses_by_size']:
            self.results_tree.insert('', 'end', values=('--- HARNESSES ---', '', ''))
            for size in sorted(totals['harnesses_by_size'].keys(), reverse=True):
                qty = totals['harnesses_by_size'][size]
                self.results_tree.insert('', 'end', values=(f"{size}-String Harness", qty, 'ea'))
        
        # Extenders
        total_extenders = totals['extenders_short'] + totals['extenders_long']
        if total_extenders > 0:
            self.results_tree.insert('', 'end', values=('--- EXTENDERS ---', '', ''))
            self.results_tree.insert('', 'end', values=(
                f"Extender {short_extender_length}ft (short side)", 
                totals['extenders_short'], 
                'ea'
            ))
            self.results_tree.insert('', 'end', values=(
                f"Extender {long_extender_length}ft (long side)", 
                totals['extenders_long'], 
                'ea'
            ))
        
        # Whips
        if totals['whips_by_length']:
            self.results_tree.insert('', 'end', values=('--- WHIPS ---', '', ''))
            for length in sorted(totals['whips_by_length'].keys()):
                qty = totals['whips_by_length'][length]
                self.results_tree.insert('', 'end', values=(f"Whip {length}ft", qty, 'ea'))
            self.results_tree.insert('', 'end', values=(
                'Total Whip Length', 
                f"{totals['total_whip_length']:,}", 
                'ft'
            ))

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
            
            # Block summary headers
            block_headers = ['Subarray', 'Block', 'Type', '# CBs', 'Breaker (A)', 
                           'Tracker Configs', 'Total Strings', 'Total Trackers']
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
                    
                    row_data = [
                        subarray['name'],
                        block['name'],
                        block['type'].title(),
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
            
            # ========== BOM RESULTS SECTION ==========
            ws.merge_cells(f'A{row}:E{row}')
            ws.cell(row=row, column=1, value="Estimated Bill of Materials").font = title_font
            row += 1
            
            # BOM headers
            bom_headers = ['Item', 'Quantity', 'Unit']
            for col, header in enumerate(bom_headers, 1):
                cell = ws.cell(row=row, column=col, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = center_align
                cell.border = thin_border
            row += 1
            
            # Pull results from the results treeview
            for item_id in self.results_tree.get_children():
                values = self.results_tree.item(item_id, 'values')
                if len(values) >= 3:
                    item_name = values[0]
                    qty = values[1]
                    unit = values[2]
                    
                    # Check if this is a section header
                    is_section = str(item_name).startswith('---')
                    is_warning = '⚠' in str(item_name)
                    
                    cell_item = ws.cell(row=row, column=1, value=str(item_name).replace('---', '').strip() if is_section else item_name)
                    cell_qty = ws.cell(row=row, column=2, value=qty if qty else '')
                    cell_unit = ws.cell(row=row, column=3, value=unit if unit else '')
                    
                    if is_section:
                        cell_item.font = section_font
                        cell_item.fill = section_fill
                        cell_qty.fill = section_fill
                        cell_unit.fill = section_fill
                    elif is_warning:
                        cell_item.fill = warning_fill
                        cell_qty.fill = warning_fill
                        cell_unit.fill = warning_fill
                    
                    for c in [cell_item, cell_qty, cell_unit]:
                        c.border = thin_border
                        c.alignment = center_align
                    cell_item.alignment = wrap_align
                    
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