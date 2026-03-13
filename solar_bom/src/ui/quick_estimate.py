import tkinter as tk
from tkinter import ttk
from typing import Optional, Dict, List, Any
from pathlib import Path
import json
import uuid
import math
import copy
from datetime import datetime
from src.utils.string_allocation import allocate_strings, allocate_strings_sequential, allocate_strings_spatial


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
        
        self.groups = []
        self.pads = []  # Collection points (inverter pads)
        self.device_names = {}  # {device_idx: "custom_name"} for CB/SI renaming
        self.last_combiner_assignments = []  # Structured CB data for Device Configurator
        self._harness_combos = []  # Track harness combo widgets for LV collection disabling
        self.selected_group_idx = None
        self._updating_listbox = False
        self.enabled_templates = self.load_enabled_templates()

        # Global settings defaults
        self.module_width_default = 1134
        self.modules_per_string_default = 28
        self.row_spacing_default = 20.0
        
        # Track currently selected item
        self.checked_items = set()  # Items checked for export
        self._results_stale = True
        self._calc_btn = None  # Reference to calculate button
        self._autosave_after_id = None
        
        # Allocation lock state
        self.allocation_locked = False
        self.locked_allocation_result = None
        
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

    def load_enabled_templates(self):
        """Load tracker templates that are enabled for the current project.
        Returns {template_key: template_data_dict} for enabled templates only."""
        templates = {}
        try:
            template_path = Path('data/tracker_templates.json')
            if not template_path.exists():
                return templates
            
            with open(template_path, 'r') as f:
                data = json.load(f)
            
            if not data:
                return templates
            
            # Parse hierarchical format
            all_templates = {}
            first_value = next(iter(data.values()))
            if isinstance(first_value, dict) and not any(key in first_value for key in ['module_orientation', 'modules_per_string']):
                for manufacturer, template_group in data.items():
                    for template_name, template_data in template_group.items():
                        unique_name = f"{manufacturer} - {template_name}"
                        all_templates[unique_name] = template_data
            else:
                all_templates = data
            
            # Filter to enabled only
            if self.current_project and hasattr(self.current_project, 'enabled_templates'):
                for key in self.current_project.enabled_templates:
                    if key in all_templates:
                        templates[key] = all_templates[key]
            else:
                templates = all_templates
                
        except Exception as e:
            print(f"Error loading enabled templates: {e}")
        
        return templates

    def get_template_display_name(self, template_key):
        """Build a short display name for a template key.
        E.g. 'JA Solar - JAM72S30-550/MR 2x28 3S' -> 'JAM72S30-550/MR 2x28 3S (3-string)'
        """
        template_data = self.enabled_templates.get(template_key)
        if not template_data:
            return template_key
        
        spt = template_data.get('strings_per_tracker', '?')
        # Strip manufacturer prefix for shorter display
        if ' - ' in template_key:
            short_name = template_key.split(' - ', 1)[1]
        else:
            short_name = template_key
        return f"{short_name} ({spt}S)"

    def get_template_module_id(self, template_key):
        """Get a module identity string from a template for consistency checking.
        Returns 'Manufacturer Model' or None."""
        template_data = self.enabled_templates.get(template_key)
        if not template_data:
            return None
        module_spec = template_data.get('module_spec', {})
        manufacturer = module_spec.get('manufacturer', '')
        model = module_spec.get('model', '')
        if manufacturer and model:
            return f"{manufacturer} {model}"
        return None

    def get_group_module_id(self, group):
        """Get the module identity for a group based on its first segment's template.
        Returns module_id string or None if no segments have templates."""
        for seg in group.get('segments', []):
            ref = seg.get('template_ref')
            if ref:
                return self.get_template_module_id(ref)
        return None

    def get_compatible_templates(self, group):
        """Get template keys compatible with the group's existing module.
        If group has no template-linked segments, all enabled templates are compatible."""
        group_module = self.get_group_module_id(group)
        
        if group_module is None:
            # No module constraint yet — all enabled templates are valid
            return list(self.enabled_templates.keys())
        
        # Filter to templates with the same module
        compatible = []
        for key in self.enabled_templates:
            if self.get_template_module_id(key) == group_module:
                compatible.append(key)
        return compatible
    
    def refresh_templates(self):
        """Reload enabled templates from disk. Call when templates may have changed."""
        self.enabled_templates = self.load_enabled_templates()

    def _derive_module_from_templates(self):
        """Derive the active module and modules_per_string from linked templates.
        Sets self.selected_module and updates modules_per_string_var.
        Returns True if a module was found."""
        from ..models.module import ModuleSpec, ModuleType
        
        # Scan all groups for the first template-linked segment
        for group in self.groups:
            for seg in group.get('segments', []):
                ref = seg.get('template_ref')
                if ref and ref in self.enabled_templates:
                    tdata = self.enabled_templates[ref]
                    module_data = tdata.get('module_spec', {})
                    mps = tdata.get('modules_per_string', 28)
                    
                    try:
                        self.selected_module = ModuleSpec(
                            manufacturer=module_data.get('manufacturer', 'Unknown'),
                            model=module_data.get('model', 'Unknown'),
                            type=ModuleType(module_data.get('type', 'Mono PERC')),
                            length_mm=float(module_data.get('length_mm', 2000)),
                            width_mm=float(module_data.get('width_mm', 1000)),
                            depth_mm=float(module_data.get('depth_mm', 40)),
                            weight_kg=float(module_data.get('weight_kg', 25)),
                            wattage=float(module_data.get('wattage', 400)),
                            vmp=float(module_data.get('vmp', 40)),
                            imp=float(module_data.get('imp', 10)),
                            voc=float(module_data.get('voc', 48)),
                            isc=float(module_data.get('isc', 10.5)),
                            max_system_voltage=float(module_data.get('max_system_voltage', 1500))
                        )
                    except Exception as e:
                        print(f"Error creating ModuleSpec from template: {e}")
                        return False
                    
                    # Update modules_per_string from template
                    if hasattr(self, 'modules_per_string_var'):
                        self.modules_per_string_var.set(str(mps))
                    
                    # Update module info label
                    if hasattr(self, 'module_info_label'):
                        self.module_info_label.config(
                            text=f"{self.selected_module.manufacturer} {self.selected_module.model} ({self.selected_module.wattage}W)  |  Isc: {self.selected_module.isc}A  |  Width: {self.selected_module.width_mm}mm",
                            foreground='black'
                        )
                    
                    self._update_strings_per_inverter()
                    return True
        
        # No templates linked — try fallback from saved estimate data
        if self.current_project and self.estimate_id:
            estimate_data = self.current_project.quick_estimates.get(self.estimate_id, {})
            saved_module_name = estimate_data.get('module_name', '')
            
            if saved_module_name and saved_module_name in self.available_modules:
                self.selected_module = self.available_modules[saved_module_name]
                saved_mps = estimate_data.get('modules_per_string', 28)
                if hasattr(self, 'modules_per_string_var'):
                    self.modules_per_string_var.set(str(saved_mps))
                if hasattr(self, 'module_info_label'):
                    self.module_info_label.config(
                        text=f"(Legacy) {self.selected_module.manufacturer} {self.selected_module.model} ({self.selected_module.wattage}W)  —  link templates to update",
                        foreground='orange'
                    )
                self._update_strings_per_inverter()
                return True
            
            # Try matching by partial name (old format might not match exactly)
            if saved_module_name:
                for display_name, mod in self.available_modules.items():
                    if (mod.manufacturer in saved_module_name and 
                        mod.model in saved_module_name):
                        self.selected_module = mod
                        saved_mps = estimate_data.get('modules_per_string', 28)
                        if hasattr(self, 'modules_per_string_var'):
                            self.modules_per_string_var.set(str(saved_mps))
                        if hasattr(self, 'module_info_label'):
                            self.module_info_label.config(
                                text=f"(Legacy) {mod.manufacturer} {mod.model} ({mod.wattage}W)  —  link templates to update",
                                foreground='orange'
                            )
                        self._update_strings_per_inverter()
                        return True
        
        # Truly no module available
        self.selected_module = None
        if hasattr(self, 'module_info_label'):
            self.module_info_label.config(
                text="No templates linked — link a tracker template to a segment",
                foreground='gray'
            )
        return False

    def get_group_summary_info(self, group):
        """Build summary info lines for a group based on its linked templates.
        Returns a list of (label, value) tuples."""
        info = []
        
        # Collect unique templates used in this group
        template_refs = set()
        for seg in group.get('segments', []):
            ref = seg.get('template_ref')
            if ref and ref in self.enabled_templates:
                template_refs.add(ref)
        
        if not template_refs:
            info.append(("Templates", "No templates linked (unlinked segments)"))
            return info
        
        # Module info (should be consistent within a group)
        first_ref = next(iter(template_refs))
        first_data = self.enabled_templates[first_ref]
        module_spec = first_data.get('module_spec', {})
        
        module_str = f"{module_spec.get('manufacturer', '?')} {module_spec.get('model', '?')} ({module_spec.get('wattage', '?')}W)"
        info.append(("Module", module_str))
        info.append(("Module Isc", f"{module_spec.get('isc', '?')} A"))
        info.append(("Module Width", f"{module_spec.get('width_mm', '?')} mm"))
        
        # Per-template info
        for ref in sorted(template_refs):
            tdata = self.enabled_templates[ref]
            spt = tdata.get('strings_per_tracker', '?')
            mps = tdata.get('modules_per_string', '?')
            modules_high = tdata.get('modules_high', 1)
            orientation = tdata.get('module_orientation', 'Portrait')
            has_motor = tdata.get('has_motor', True)
            motor_type = tdata.get('motor_placement_type', 'between_strings')
            
            # Calculate physical width
            width_mm = module_spec.get('width_mm', 1000)
            length_mm = module_spec.get('length_mm', 2000)
            spacing_m = tdata.get('module_spacing_m', 0.02)
            motor_gap_m = tdata.get('motor_gap_m', 1.0) if has_motor else 0
            
            if orientation == 'Portrait':
                module_dim_along_tracker = length_mm / 1000  # meters
            else:
                module_dim_along_tracker = width_mm / 1000
            
            try:
                full_spt_info = int(float(spt))
                mps_info = int(mps)
                partial_info = round((float(spt) - full_spt_info) * mps_info) if float(spt) != full_spt_info else 0
                total_modules = full_spt_info * mps_info + partial_info
                modules_in_row = total_modules
                tracker_length_m = (modules_in_row * module_dim_along_tracker + 
                                   (modules_in_row - 1) * spacing_m +
                                   (motor_gap_m if has_motor else 0))
            except (ValueError, TypeError):
                total_modules = '?'
                tracker_length_m = None
            tracker_length_ft = f"{tracker_length_m * 3.28084:.1f} ft" if tracker_length_m else '?'
            
            short_name = ref.split(' - ', 1)[1] if ' - ' in ref else ref
            info.append(("", ""))  # spacer
            info.append(("Template", short_name))
            info.append(("  Strings/Tracker", str(spt)))
            info.append(("  Modules/String", str(mps)))
            info.append(("  Modules High", str(modules_high)))
            info.append(("  Orientation", orientation))
            info.append(("  Total Modules", str(total_modules)))
            info.append(("  Tracker Length", tracker_length_ft))
            if has_motor:
                info.append(("  Motor", motor_type.replace('_', ' ').title()))
        
        return info

    def generate_id(self, prefix: str) -> str:
        """Generate a unique ID with a prefix"""
        return f"{prefix}_{uuid.uuid4().hex[:8]}"
            
    def disable_combobox_scroll(self, combobox):
        """Prevent combobox from responding to mouse wheel"""
        def _ignore_scroll(event):
            return "break"
        
        combobox.bind("<MouseWheel>", _ignore_scroll)
        # For Linux compatibility
        combobox.bind("<Button-4>", _ignore_scroll)
        combobox.bind("<Button-5>", _ignore_scroll)

    def update_string_count(self):
        """Legacy method — now handled by _update_group_string_count"""
        pass

    # ==================== Data Management ====================

    def add_group(self) -> int:
        """Add a new group and return its index"""
        group_num = len(self.groups) + 1
        
        # Start unconstrained — let the user pick a template
        default_ref = None
        default_spt = 3
        
        group = {
            'name': f"Group {group_num}",
            'device_position': 'middle',
            'driveline_angle': 0.0,
            'segments': [
                {'quantity': 1, 'strings_per_tracker': default_spt, 'harness_config': str(default_spt), 'template_ref': default_ref}
            ]
        }
        self.groups.append(group)
        self._refresh_group_listbox()
        
        # Select the new group
        idx = len(self.groups) - 1
        self.group_listbox.selection_clear(0, tk.END)
        self.group_listbox.selection_set(idx)
        self.group_listbox.see(idx)
        self.on_group_select(None)
        
        self._auto_unlock_allocation()
        self._mark_stale()
        self._schedule_autosave()
        return idx
    
    def copy_selected_group(self):
        """Copy the currently selected group"""
        sel = self.group_listbox.curselection()
        if not sel:
            return
        
        source = self.groups[sel[0]]
        new_group = copy.deepcopy(source)
        new_group['name'] = f"{source['name']} (Copy)"
        
        # Insert after the selected group
        insert_idx = sel[0] + 1
        self.groups.insert(insert_idx, new_group)
        self._refresh_group_listbox()
        
        self.group_listbox.selection_clear(0, tk.END)
        self.group_listbox.selection_set(insert_idx)
        self.group_listbox.see(insert_idx)
        self.on_group_select(None)
        
        self._auto_unlock_allocation()
        self._mark_stale()
        self._schedule_autosave()
    
    def delete_selected_group(self):
        """Delete the currently selected group"""
        sel = self.group_listbox.curselection()
        if not sel:
            return
        
        idx = sel[0]
        del self.groups[idx]
        self.selected_group_idx = None
        self._refresh_group_listbox(preserve_selection=False)
        
        # Select nearest group
        if self.groups:
            new_idx = min(idx, len(self.groups) - 1)
            self.group_listbox.selection_set(new_idx)
            self.on_group_select(None)
        else:
            self.clear_details_panel()
        
        self._mark_stale()
        self._schedule_autosave()
    
    def move_group_up(self):
        """Move selected group up"""
        sel = self.group_listbox.curselection()
        if not sel or sel[0] == 0:
            return
        
        idx = sel[0]
        self.groups[idx], self.groups[idx - 1] = self.groups[idx - 1], self.groups[idx]
        self._refresh_group_listbox()
        self.group_listbox.selection_set(idx - 1)
        self.on_group_select(None)
        
        self._mark_stale()
        self._schedule_autosave()
    
    def move_group_down(self):
        """Move selected group down"""
        sel = self.group_listbox.curselection()
        if not sel or sel[0] >= len(self.groups) - 1:
            return
        
        idx = sel[0]
        self.groups[idx], self.groups[idx + 1] = self.groups[idx + 1], self.groups[idx]
        self._refresh_group_listbox()
        self.group_listbox.selection_set(idx + 1)
        self.on_group_select(None)
        
        self._mark_stale()
        self._schedule_autosave()
    
    def _refresh_group_listbox(self, preserve_selection=True):
        """Refresh the listbox display from self.groups"""
        sel = self.group_listbox.curselection()
        old_idx = sel[0] if sel else None
        
        self._updating_listbox = True
        self.group_listbox.delete(0, tk.END)
        for group in self.groups:
            total_trackers = sum(seg['quantity'] for seg in group['segments'])
            total_strings = sum(int(seg['quantity'] * seg['strings_per_tracker']) for seg in group['segments'])
            display = f"{group['name']}  ({total_trackers}T / {total_strings}S)"
            self.group_listbox.insert(tk.END, display)
        
        if preserve_selection and old_idx is not None and old_idx < len(self.groups):
            self.group_listbox.selection_set(old_idx)
        self._updating_listbox = False

    def round_whip_length(self, raw_length_ft):
        """Apply 5% waste factor and round up to nearest 10ft increment (min 10ft)"""
        import math
        WASTE_FACTOR = 1.05
        INCREMENT = 10
        length_with_waste = raw_length_ft * WASTE_FACTOR
        rounded = INCREMENT * math.ceil(length_with_waste / INCREMENT)
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
        for group_num in range(1, total_rows + 1):
            # Find which group this row belongs to
            for group in groups:
                if group['row_start'] <= group_num <= group['row_end']:
                    distance_ft = abs(group_num - group['cb_position']) * row_spacing_ft
                    whip_distances.append((distance_ft, group['cb_idx']))
                    break
        
        return whip_distances

    def calculate_whip_distances_from_positions(self, allocation_result, topology, num_devices, row_spacing_ft):
        """Calculate whip distances using device positions derived from allocation.
        
        For Distributed String and Centralized String: uses allocation result
        to map each inverter's trackers to a device, computing real E-W + N-S distance.
        
        For Central Inverter: falls back to the abstract even-spacing method.
        
        Returns a flat list of distance_ft values, one per tracker.
        """
        if topology == 'Central Inverter' or not allocation_result:
            total_trackers = sum(
                sum(seg['quantity'] for seg in group['segments'])
                for group in self.groups
            )
            old_distances = self.calculate_cb_whip_distances(
                total_trackers, num_devices, row_spacing_ft
            )
            # Build flat spt list matching tracker order
            spt_list = []
            for group in self.groups:
                for seg in group['segments']:
                    for _ in range(seg['quantity']):
                        spt_list.append(seg['strings_per_tracker'])
            return [(d[0], spt_list[i] if i < len(spt_list) else 0, i, -1) for i, d in enumerate(old_distances)]
        
        # Build world X for every global tracker index
        # Uses saved group positions or auto-layout, same as site preview
        tracker_world_x = []
        tracker_world_y = []
        tracker_group = []
        auto_x_cursor = 0.0
        
        for grp_idx, group in enumerate(self.groups):
            saved_x = group.get('position_x')
            saved_y = group.get('position_y')
            if saved_x is not None:
                group_x = saved_x
            else:
                group_x = auto_x_cursor
            group_y = saved_y if saved_y is not None else 0.0
            
            driveline_angle_deg = group.get('driveline_angle', 0.0)
            driveline_tan = math.tan(math.radians(driveline_angle_deg)) if driveline_angle_deg > 0 else 0.0
            
            local_idx = 0
            for seg in group['segments']:
                for _ in range(seg['quantity']):
                    local_x_offset = local_idx * row_spacing_ft
                    tracker_world_x.append(group_x + local_x_offset)
                    tracker_world_y.append(group_y + local_x_offset * driveline_tan)
                    tracker_group.append(grp_idx)
                    local_idx += 1
            
            # Advance auto cursor
            group_tracker_count = sum(seg['quantity'] for seg in group['segments'])
            auto_x_cursor += group_tracker_count * row_spacing_ft + row_spacing_ft * 2
        
        # Compute device world X and metadata for each inverter
        inverters = allocation_result.get('inverters', [])
        device_info = []  # list of (world_x, device_position) per inverter
        
        for inv_idx, inv in enumerate(inverters):
            harness_map = inv.get('harness_map', [])
            if not harness_map:
                device_info.append(None)
                continue
            
            # Find majority group
            group_counts = {}
            for entry in harness_map:
                tidx = entry['tracker_idx']
                if tidx < len(tracker_group):
                    grp = tracker_group[tidx]
                    group_counts[grp] = group_counts.get(grp, 0) + 1
            
            if not group_counts:
                device_info.append(None)
                continue
            
            primary_grp = max(group_counts, key=group_counts.get)
            group_source = self.groups[primary_grp] if primary_grp < len(self.groups) else {}
            device_position = group_source.get('device_position', 'middle')
            
            # Device X = average of its trackers' world X positions in the primary group
            inv_tracker_xs = []
            for entry in harness_map:
                tidx = entry['tracker_idx']
                if tidx < len(tracker_world_x) and tracker_group[tidx] == primary_grp:
                    inv_tracker_xs.append(tracker_world_x[tidx])
            
            if inv_tracker_xs:
                device_x = (min(inv_tracker_xs) + max(inv_tracker_xs)) / 2.0
            else:
                device_x = 0
            
            # Compute device Y from average of its trackers' Y positions + device_position offset
            inv_tracker_ys = []
            for entry in harness_map:
                tidx = entry['tracker_idx']
                if tidx < len(tracker_world_y) and tracker_group[tidx] == primary_grp:
                    inv_tracker_ys.append(tracker_world_y[tidx])
            device_y = (min(inv_tracker_ys) + max(inv_tracker_ys)) / 2.0 if inv_tracker_ys else 0.0
            
            device_info.append((device_x, device_y, device_position))
        
        def get_ns_base_offset(device_position):
            """Fixed N-S offset from device_position setting (north/south/middle)."""
            if device_position == 'middle':
                return 0.0
            return 6.5  # 5ft offset + half device height
        
        # Compute distance for each tracker to its assigned device
        # Deduplicate: each physical tracker generates whips once, even if split across inverters
        whip_distances = []
        seen_trackers = set()
        
        for inv_idx, inv in enumerate(inverters):
            if inv_idx >= len(device_info) or device_info[inv_idx] is None:
                continue
            
            dev_x, dev_y, dev_position = device_info[inv_idx]
            ns_base = get_ns_base_offset(dev_position)
            
            for entry in inv['harness_map']:
                tidx = entry['tracker_idx']
                is_split = entry.get('is_split', False)
                
                # Allow split tracker portions through (they appear in multiple inverters)
                if tidx in seen_trackers and not is_split:
                    continue
                seen_trackers.add(tidx)
                
                if tidx >= len(tracker_world_x):
                    continue
                
                ew_distance = abs(tracker_world_x[tidx] - dev_x)
                # N-S distance = angle-based Y difference + fixed device_position offset
                ns_from_angle = abs(tracker_world_y[tidx] - dev_y)
                ns_total = ns_from_angle + ns_base
                total_distance = (ew_distance**2 + ns_total**2) ** 0.5
                
                spt = entry.get('strings_per_tracker', 0)
                
                whip_distances.append((total_distance, spt, tidx, inv_idx))
        
        return whip_distances
    
    def calculate_routed_feeder_distances(self, allocation_result, topology, row_spacing_ft):
        """Calculate Manhattan feeder distances from each device to its assigned pad.
        
        Returns a dict with:
            'feeder_distances': list of (device_label, distance_ft) tuples
            'feeder_total_ft': total feeder cable
            'feeder_count': number of feeder runs
            'homerun_distances': list of (pad_label, distance_ft) for AC homeruns (stub)
        """
        result = {
            'feeder_distances': [],
            'feeder_total_ft': 0,
            'feeder_count': 0,
        }
        
        if not self.pads or not allocation_result:
            return result
        
        inverters = allocation_result.get('inverters', [])
        if not inverters:
            return result
        
        # Build device world X positions (same as whip calc)
        device_world_x = []
        device_group = []
        auto_x_cursor = 0.0
        
        for grp_idx, group in enumerate(self.groups):
            saved_x = group.get('position_x')
            group_x = saved_x if saved_x is not None else auto_x_cursor
            
            local_idx = 0
            tracker_info_local = []
            for seg in group['segments']:
                for _ in range(seg['quantity']):
                    tracker_info_local.append((grp_idx, local_idx))
                    local_idx += 1
            
            group_tracker_count = sum(seg['quantity'] for seg in group['segments'])
            auto_x_cursor += group_tracker_count * row_spacing_ft + row_spacing_ft * 2
        
        # Recompute device centers (same logic as whip calc and site preview)
        tracker_world_x = []
        tracker_grp = []
        auto_x_cursor = 0.0
        
        for grp_idx, group in enumerate(self.groups):
            saved_x = group.get('position_x')
            group_x = saved_x if saved_x is not None else auto_x_cursor
            
            local_idx = 0
            for seg in group['segments']:
                for _ in range(seg['quantity']):
                    tracker_world_x.append(group_x + local_idx * row_spacing_ft)
                    tracker_grp.append(grp_idx)
                    local_idx += 1
            
            group_tracker_count = sum(seg['quantity'] for seg in group['segments'])
            auto_x_cursor += group_tracker_count * row_spacing_ft + row_spacing_ft * 2
        
        # Compute device positions (center X, Y based on device_position)
        device_positions = []
        for inv_idx, inv in enumerate(inverters):
            harness_map = inv.get('harness_map', [])
            if not harness_map:
                device_positions.append(None)
                continue
            
            group_counts = {}
            for entry in harness_map:
                tidx = entry['tracker_idx']
                if tidx < len(tracker_grp):
                    g = tracker_grp[tidx]
                    group_counts[g] = group_counts.get(g, 0) + 1
            
            if not group_counts:
                device_positions.append(None)
                continue
            
            primary_grp = max(group_counts, key=group_counts.get)
            group_source = self.groups[primary_grp] if primary_grp < len(self.groups) else {}
            device_position = group_source.get('device_position', 'middle')
            
            inv_xs = [tracker_world_x[e['tracker_idx']] for e in harness_map 
                      if e['tracker_idx'] < len(tracker_world_x) and tracker_grp[e['tracker_idx']] == primary_grp]
            
            dev_x = (min(inv_xs) + max(inv_xs)) / 2.0 if inv_xs else 0
            
            # Approximate device Y from group position
            saved_y = group_source.get('position_y', 0)
            group_y = saved_y if saved_y is not None else 0
            
            # Get tracker length for Y offset calculation
            first_ref = None
            for seg in group_source.get('segments', []):
                ref = seg.get('template_ref')
                if ref and ref in self.enabled_templates:
                    first_ref = ref
                    break
            
            tracker_length_ft = 180.0  # fallback
            if first_ref:
                dims = self._get_estimate_tracker_dims_ft(first_ref)
                if dims:
                    tracker_length_ft = dims[1]
            
            if device_position == 'north':
                dev_y = group_y - 5.0
            elif device_position == 'south':
                dev_y = group_y + tracker_length_ft + 5.0
            else:
                dev_y = group_y + tracker_length_ft / 2.0
            
            device_positions.append((dev_x, dev_y))
        
        # Build device -> pad lookup
        device_to_pad = {}
        for pad_idx, pad in enumerate(self.pads):
            for dev_idx in pad.get('assigned_devices', []):
                device_to_pad[dev_idx] = pad_idx
        
        # Compute Manhattan distance from each device to its pad
        for dev_idx, dev_pos in enumerate(device_positions):
            if dev_pos is None:
                continue
            
            pad_idx = device_to_pad.get(dev_idx)
            if pad_idx is None or pad_idx >= len(self.pads):
                continue
            
            pad = self.pads[pad_idx]
            pad_cx = pad['x'] + pad.get('width_ft', 10.0) / 2
            pad_cy = pad['y'] + pad.get('height_ft', 8.0) / 2
            
            manhattan = abs(dev_pos[0] - pad_cx) + abs(dev_pos[1] - pad_cy)
            
            label = f"Dev-{dev_idx+1:02d}"
            result['feeder_distances'].append((label, manhattan))
            result['feeder_total_ft'] += manhattan
            result['feeder_count'] += 1
        
        return result
    
    def _build_wire_sizing_frame(self, parent):
        """Build the Wire Sizing widget to the right of Global Settings."""
        from src.utils.cable_sizing import (
            CABLE_SIZE_ORDER, CABLE_SIZE_ORDER_EXTENDED, 
            get_available_sizes
        )
        
        ws_frame = ttk.LabelFrame(parent, text="Wire Sizing", padding="5")
        ws_frame.pack(side='left', fill='y', padx=(10, 0))
        self._ws_frame = ws_frame
        
        # Row 0: Temp rating, Material toggle, Reset button
        controls_row = ttk.Frame(ws_frame)
        controls_row.pack(fill='x', pady=(0, 5))
        
        ttk.Label(controls_row, text="Temp:").pack(side='left', padx=(0, 2))
        self._ws_temp_var = tk.StringVar(value=self.wire_sizing.get('temp_rating', '90C'))
        temp_combo = ttk.Combobox(
            controls_row, textvariable=self._ws_temp_var,
            values=['60C', '75C', '90C'], state='readonly', width=4
        )
        temp_combo.pack(side='left', padx=(0, 8))
        self.disable_combobox_scroll(temp_combo)
        
        ttk.Label(controls_row, text="Feeder:").pack(side='left', padx=(0, 2))
        self._ws_material_var = tk.StringVar(value=self.wire_sizing.get('feeder_material', 'aluminum'))
        material_combo = ttk.Combobox(
            controls_row, textvariable=self._ws_material_var,
            values=['aluminum', 'copper'], state='readonly', width=9
        )
        material_combo.pack(side='left', padx=(0, 8))
        self.disable_combobox_scroll(material_combo)
        
        ttk.Button(controls_row, text="Reset", width=5,
                    command=self._reset_wire_sizing_to_recommended).pack(side='left')
        
        # Column headers
        header_row = ttk.Frame(ws_frame)
        header_row.pack(fill='x', pady=(0, 2))
        
        header_font = ('TkDefaultFont', 8, 'bold')
        ttk.Label(header_row, text="Strings", font=header_font, width=7).pack(side='left')
        ttk.Label(header_row, text="Harness", font=header_font, width=10).pack(side='left', padx=2)
        ttk.Label(header_row, text="Extender", font=header_font, width=10).pack(side='left', padx=2)
        ttk.Label(header_row, text="Whip", font=header_font, width=10).pack(side='left', padx=2)
        
        # Dynamic rows container — cleared and rebuilt by refresh_wire_sizing_table
        self._ws_rows_frame = ttk.Frame(ws_frame)
        self._ws_rows_frame.pack(fill='x')
        
        # Storage for combo vars and override tracking
        self._ws_lv_combos = {}    # {(string_count, cable_type): tk.StringVar}
        self._ws_feeder_var = tk.StringVar(value=self.wire_sizing.get('dc_feeder', ''))
        self._ws_homerun_var = tk.StringVar(value=self.wire_sizing.get('ac_homerun', ''))
        
        # Available LV cable sizes (copper AWG only — pre-made assemblies)
        self._ws_lv_sizes = CABLE_SIZE_ORDER  # ['10 AWG' through '4/0 AWG']
        
        # Traces for temp and material changes
        self._ws_temp_var.trace_add('write', lambda *a: self._on_wire_sizing_setting_changed())
        self._ws_material_var.trace_add('write', lambda *a: self._on_wire_sizing_setting_changed())
        
    def _on_wire_sizing_setting_changed(self):
        """Handle change to temp rating or feeder material — recalc recommendations."""
        self.wire_sizing['temp_rating'] = self._ws_temp_var.get()
        self.wire_sizing['feeder_material'] = self._ws_material_var.get()
        self._reset_wire_sizing_to_recommended()
        self._mark_stale()
        self._schedule_autosave()

    def refresh_wire_sizing_table(self):
        """Refresh the wire sizing table rows based on current harness configs."""
        if not hasattr(self, '_ws_rows_frame'):
            return
        
        from src.utils.cable_sizing import CABLE_SIZE_ORDER, get_available_sizes
        
        # Clear existing rows
        for widget in self._ws_rows_frame.winfo_children():
            widget.destroy()
        self._ws_lv_combos.clear()
        
        active_counts = self._collect_active_string_counts()
        topology = self.topology_var.get() if hasattr(self, 'topology_var') else 'Distributed String'
        by_sc = self.wire_sizing.get('by_string_count', {})
        
        # Build LV rows (harness/extender/whip per string count)
        for sc in active_counts:
            row_frame = ttk.Frame(self._ws_rows_frame)
            row_frame.pack(fill='x', pady=1)
            
            ttk.Label(row_frame, text=f"{sc}-str", width=7).pack(side='left')
            
            # Get current sizes from wire_sizing dict (int or str key)
            entry = by_sc.get(sc) or by_sc.get(str(sc)) or {}
            
            for cable_type in ('harness', 'extender', 'whip'):
                current_val = entry.get(cable_type, '10 AWG')
                var = tk.StringVar(value=current_val)
                combo = ttk.Combobox(
                    row_frame, textvariable=var,
                    values=CABLE_SIZE_ORDER, state='readonly', width=8
                )
                combo.pack(side='left', padx=2)
                self.disable_combobox_scroll(combo)
                
                # Store reference
                self._ws_lv_combos[(sc, cable_type)] = var
                
                # Bind change handler with closure
                var.trace_add('write', lambda *a, s=sc, ct=cable_type: self._on_lv_size_changed(s, ct))
        
        # Separator before feeder rows
        if active_counts and (topology != 'Distributed String' or True):
            sep = ttk.Separator(self._ws_rows_frame, orient='horizontal')
            sep.pack(fill='x', pady=3)
        
        # DC Feeder row (only for Centralized String / Central Inverter)
        if topology in ('Centralized String', 'Central Inverter'):
            feeder_row = ttk.Frame(self._ws_rows_frame)
            feeder_row.pack(fill='x', pady=1)
            
            ttk.Label(feeder_row, text="DC Fdr", width=7).pack(side='left')
            material = self.wire_sizing.get('feeder_material', 'aluminum')
            feeder_sizes = get_available_sizes(material)
            current_feeder = self.wire_sizing.get('dc_feeder', '')
            self._ws_feeder_var.set(current_feeder)
            feeder_combo = ttk.Combobox(
                feeder_row, textvariable=self._ws_feeder_var,
                values=feeder_sizes, state='readonly', width=10
            )
            feeder_combo.pack(side='left', padx=2)
            self.disable_combobox_scroll(feeder_combo)
            self._ws_feeder_var.trace_add('write', lambda *a: self._on_feeder_size_changed('dc_feeder'))
        
        # AC Homerun row (all topologies)
        homerun_row = ttk.Frame(self._ws_rows_frame)
        homerun_row.pack(fill='x', pady=1)
        
        ttk.Label(homerun_row, text="AC HR", width=7).pack(side='left')
        material = self.wire_sizing.get('feeder_material', 'aluminum')
        homerun_sizes = get_available_sizes(material)
        current_homerun = self.wire_sizing.get('ac_homerun', '')
        self._ws_homerun_var.set(current_homerun)
        homerun_combo = ttk.Combobox(
            homerun_row, textvariable=self._ws_homerun_var,
            values=homerun_sizes, state='readonly', width=10
        )
        homerun_combo.pack(side='left', padx=2)
        self.disable_combobox_scroll(homerun_combo)
        self._ws_homerun_var.trace_add('write', lambda *a: self._on_feeder_size_changed('ac_homerun'))
    
    def _reset_wire_sizing_to_recommended(self):
        """Reset all wire sizes to calculated recommendations."""
        from src.utils.cable_sizing import (
            recommend_lv_cable_sizes, recommend_dc_feeder_size, 
            recommend_ac_homerun_size
        )
        
        temp = self._ws_temp_var.get() if hasattr(self, '_ws_temp_var') else '90C'
        material = self._ws_material_var.get() if hasattr(self, '_ws_material_var') else 'aluminum'
        
        self.wire_sizing['temp_rating'] = temp
        self.wire_sizing['feeder_material'] = material
        self.wire_sizing['user_overrides'] = {}  # Clear all overrides
        
        # Get module Isc
        module_isc = self.selected_module.isc if self.selected_module else 13.0
        
        # Recommend LV sizes for each active string count
        active_counts = self._collect_active_string_counts()
        by_sc = {}
        for sc in active_counts:
            sizes = recommend_lv_cable_sizes(sc, module_isc, nec_factor=1.56, temp_rating=temp)
            by_sc[sc] = sizes
        self.wire_sizing['by_string_count'] = by_sc
        
        # Recommend DC feeder
        topology = self.topology_var.get() if hasattr(self, 'topology_var') else 'Distributed String'
        if topology in ('Centralized String', 'Central Inverter'):
            try:
                breaker = float(self.breaker_size_var.get())
            except (ValueError, AttributeError):
                breaker = 400.0
            self.wire_sizing['dc_feeder'] = recommend_dc_feeder_size(breaker, material, temp)
        else:
            self.wire_sizing['dc_feeder'] = ''
        
        # Recommend AC homerun
        if self.selected_inverter and hasattr(self.selected_inverter, 'max_ac_current'):
            max_ac = self.selected_inverter.max_ac_current
            self.wire_sizing['ac_homerun'] = recommend_ac_homerun_size(max_ac, material, temp)
        else:
            self.wire_sizing['ac_homerun'] = ''
        
        # Rebuild the table UI with new values
        self.refresh_wire_sizing_table()
        self._mark_stale()
        self._schedule_autosave()

    def _on_lv_size_changed(self, string_count, cable_type):
        """Handle user changing a harness/extender/whip size in the wire sizing table."""
        from src.utils.cable_sizing import get_cable_size_index
        
        key = (string_count, cable_type)
        if key not in self._ws_lv_combos:
            return
        
        new_size = self._ws_lv_combos[key].get()
        
        # Ensure by_string_count dict has this entry
        by_sc = self.wire_sizing.setdefault('by_string_count', {})
        entry = by_sc.get(string_count) or by_sc.get(str(string_count))
        if entry is None:
            entry = {'harness': '10 AWG', 'extender': '10 AWG', 'whip': '10 AWG'}
            by_sc[string_count] = entry
        
        entry[cable_type] = new_size
        
        # Enforce floor constraint: extender and whip >= harness
        harness_idx = get_cable_size_index(entry['harness'])
        
        if cable_type == 'harness':
            # If harness got bigger, bump extender and whip up if they're smaller
            for dep_type in ('extender', 'whip'):
                dep_idx = get_cable_size_index(entry[dep_type])
                if dep_idx < harness_idx:
                    entry[dep_type] = entry['harness']
                    dep_key = (string_count, dep_type)
                    if dep_key in self._ws_lv_combos:
                        self._ws_lv_combos[dep_key].set(entry['harness'])
        else:
            # If extender or whip was set smaller than harness, snap it back
            my_idx = get_cable_size_index(new_size)
            if my_idx < harness_idx:
                entry[cable_type] = entry['harness']
                self._ws_lv_combos[key].set(entry['harness'])
        
        # Track as user override
        overrides = self.wire_sizing.setdefault('user_overrides', {})
        overrides[f"{string_count}_{cable_type}"] = True
        
        self._mark_stale()
        self._schedule_autosave()
    
    def _on_feeder_size_changed(self, feeder_type):
        """Handle user changing DC feeder or AC homerun size."""
        if feeder_type == 'dc_feeder':
            self.wire_sizing['dc_feeder'] = self._ws_feeder_var.get()
        elif feeder_type == 'ac_homerun':
            self.wire_sizing['ac_homerun'] = self._ws_homerun_var.get()
        
        overrides = self.wire_sizing.setdefault('user_overrides', {})
        overrides[feeder_type] = True
        
        self._mark_stale()
        self._schedule_autosave()

    def _refresh_wire_sizing_for_segments(self):
        """Refresh wire sizing when segments/harness configs change.
        Preserves user overrides, adds new string counts with recommendations,
        removes string counts no longer in use."""
        from src.utils.cable_sizing import recommend_lv_cable_sizes
        
        if not hasattr(self, '_ws_rows_frame'):
            return
        
        temp = self.wire_sizing.get('temp_rating', '90C')
        module_isc = self.selected_module.isc if self.selected_module else 13.0
        overrides = self.wire_sizing.get('user_overrides', {})
        
        active_counts = self._collect_active_string_counts()
        by_sc = self.wire_sizing.get('by_string_count', {})
        
        # Add recommendations for any new string counts
        for sc in active_counts:
            if sc not in by_sc and str(sc) not in by_sc:
                sizes = recommend_lv_cable_sizes(sc, module_isc, nec_factor=1.56, temp_rating=temp)
                by_sc[sc] = sizes
        
        # Remove string counts no longer in use (but only if not user-overridden)
        keys_to_remove = []
        for sc_key in list(by_sc.keys()):
            sc_int = int(sc_key) if isinstance(sc_key, str) else sc_key
            if sc_int not in active_counts:
                # Check if any overrides exist for this string count
                has_override = any(
                    overrides.get(f"{sc_int}_{ct}") 
                    for ct in ('harness', 'extender', 'whip')
                )
                if not has_override:
                    keys_to_remove.append(sc_key)
        
        for key in keys_to_remove:
            del by_sc[key]
        
        self.wire_sizing['by_string_count'] = by_sc
        self.refresh_wire_sizing_table()
    
    def _on_topology_changed_wire_sizing(self):
        """Handle topology change — show/hide DC feeder row, recalc if needed."""
        from src.utils.cable_sizing import recommend_dc_feeder_size
        
        topology = self.topology_var.get()
        material = self.wire_sizing.get('feeder_material', 'aluminum')
        temp = self.wire_sizing.get('temp_rating', '90C')
        overrides = self.wire_sizing.get('user_overrides', {})
        
        if topology in ('Centralized String', 'Central Inverter'):
            # Generate DC feeder recommendation if not already overridden
            if not overrides.get('dc_feeder') and not self.wire_sizing.get('dc_feeder'):
                try:
                    breaker = float(self.breaker_size_var.get())
                except (ValueError, AttributeError):
                    breaker = 400.0
                self.wire_sizing['dc_feeder'] = recommend_dc_feeder_size(breaker, material, temp)
        else:
            self.wire_sizing['dc_feeder'] = ''
        
        self.refresh_wire_sizing_table()
    
    def _on_breaker_changed_wire_sizing(self):
        """Handle breaker size change — update DC feeder recommendation."""
        from src.utils.cable_sizing import recommend_dc_feeder_size
        
        overrides = self.wire_sizing.get('user_overrides', {})
        if overrides.get('dc_feeder'):
            return  # User has overridden, don't auto-update
        
        topology = self.topology_var.get()
        if topology not in ('Centralized String', 'Central Inverter'):
            return
        
        material = self.wire_sizing.get('feeder_material', 'aluminum')
        temp = self.wire_sizing.get('temp_rating', '90C')
        
        try:
            breaker = float(self.breaker_size_var.get())
        except (ValueError, AttributeError):
            breaker = 400.0
        
        self.wire_sizing['dc_feeder'] = recommend_dc_feeder_size(breaker, material, temp)
        self.refresh_wire_sizing_table()
    
    def _on_inverter_changed_wire_sizing(self):
        """Handle inverter change — update AC homerun recommendation."""
        from src.utils.cable_sizing import recommend_ac_homerun_size
        
        overrides = self.wire_sizing.get('user_overrides', {})
        if overrides.get('ac_homerun'):
            return  # User has overridden, don't auto-update
        
        material = self.wire_sizing.get('feeder_material', 'aluminum')
        temp = self.wire_sizing.get('temp_rating', '90C')
        
        if self.selected_inverter and hasattr(self.selected_inverter, 'max_ac_current'):
            max_ac = self.selected_inverter.max_ac_current
            self.wire_sizing['ac_homerun'] = recommend_ac_homerun_size(max_ac, material, temp)
        else:
            self.wire_sizing['ac_homerun'] = ''
        
        self.refresh_wire_sizing_table()
    
    def _collect_active_string_counts(self):
        """Scan segments and return sorted list of harness string counts in use."""
        string_counts = set()
        lv_method = self.lv_collection_var.get() if hasattr(self, 'lv_collection_var') else 'Wire Harness'
        
        for group in self.groups:
            for segment in group.get('segments', []):
                if lv_method == 'String HR':
                    string_counts.add(1)
                else:
                    harness_config = segment.get('harness_config', '')
                    if harness_config:
                        for part in harness_config.split('+'):
                            try:
                                sc = int(part.strip())
                                if sc > 0:
                                    string_counts.add(sc)
                            except ValueError:
                                pass
                    else:
                        spt = segment.get('strings_per_tracker', 1)
                        if spt > 0:
                            string_counts.add(spt)
        
        result = sorted(string_counts)

        return result

    def get_wire_size_for(self, cable_type, num_strings=None):
        """
        Look up wire size from the wire_sizing dict.
        
        Args:
            cable_type: 'harness', 'extender', 'whip', 'dc_feeder', 'ac_homerun'
            num_strings: Required for harness/extender/whip — the harness string count
            
        Returns:
            str: Cable size string (e.g., '10 AWG', '500 kcmil')
        """
        if cable_type in ('dc_feeder', 'ac_homerun'):
            size = self.wire_sizing.get(cable_type, '')
            return size if size else '4/0 AWG'
        
        if num_strings is None:
            return '10 AWG'
        
        by_sc = self.wire_sizing.get('by_string_count', {})
        # Try int key first, then string key (JSON round-trip converts int keys to strings)
        entry = by_sc.get(num_strings) or by_sc.get(str(num_strings))
        if entry:
            return entry.get(cable_type, '10 AWG')
        
        return '10 AWG'
    
    def _gauge_to_string_count(self, cable_type, gauge):
        """Reverse-lookup: find a string count that maps to this gauge for a given cable type.
        Used to pass num_strings to lookup_part_and_price for correct part matching."""
        by_sc = self.wire_sizing.get('by_string_count', {})
        for sc_key, entry in by_sc.items():
            if entry.get(cable_type) == gauge:
                return int(sc_key) if isinstance(sc_key, str) else sc_key
        return None

    
    def _get_effective_wire_size(self, cable_type):
        """
        Get the effective wire size for a cable type across all active string counts.
        If all string counts agree, returns that size. Otherwise returns the largest.
        Used for whips/extenders which are bucketed by length without string count context.
        
        Args:
            cable_type: 'harness', 'extender', or 'whip'
        Returns:
            str: Cable size string
        """
        from src.utils.cable_sizing import get_cable_size_index
        
        by_sc = self.wire_sizing.get('by_string_count', {})
        if not by_sc:
            return '10 AWG'
        
        sizes = set()
        for sc_key, entry in by_sc.items():
            size = entry.get(cable_type, '10 AWG')
            sizes.add(size)
        
        if len(sizes) == 1:
            return sizes.pop()
        
        # Multiple sizes — return the largest
        best = None
        best_idx = -1
        for s in sizes:
            idx = get_cable_size_index(s)
            if idx > best_idx:
                best_idx = idx
                best = s
        
        result = best or '10 AWG'
        return result

    def lookup_part_and_price(self, item_type, **kwargs):
        """Look up part number and unit price for a BOM item.
        item_type: 'harness', 'whip', 'extender'
        Returns (part_number, unit_price_str, ext_price_str)
        """
        try:
            import os
            import json
            from src.utils.pricing_lookup import PricingLookup

            # Wire gauge is now looked up per cable type from wire_sizing dict
            wire_gauge = None  # Will be set per item_type below

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
                        harness_gauge = self.get_wire_size_for('harness', num_strings)
                        if spec_trunk == harness_gauge:
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

                # Look up wire gauge for this cable type
                # If num_strings is provided, use per-string-count size; otherwise use effective size
                item_num_strings = kwargs.get('num_strings', None)
                if item_num_strings is not None:
                    item_wire_gauge = self.get_wire_size_for(item_type, item_num_strings)
                else:
                    item_wire_gauge = self._get_effective_wire_size(item_type)

                for pn, spec in library.items():
                    if pn.startswith('_comment_'):
                        continue
                    if (spec.get('wire_gauge') == item_wire_gauge and
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
        
    def _get_effective_harness_config(self, seg):
        """Get the effective harness config for a segment, considering LV Collection Method.
        
        When 'String HR' is selected, override to all 1-string harnesses (1+1+...+1).
        Otherwise, return the segment's stored harness_config.
        """
        lv_method = self.lv_collection_var.get() if hasattr(self, 'lv_collection_var') else 'Wire Harness'
        if lv_method == 'String HR':
            spt = seg.get('strings_per_tracker', 1)
            return '+'.join(['1'] * int(spt))
        return seg.get('harness_config', str(seg.get('strings_per_tracker', 1)))
        
    def calculate_extender_lengths_per_segment(self, seg, device_position):
        """Calculate per-harness positive and negative extender lengths for a segment.
        
        Returns a list of (pos_length_ft, neg_length_ft) tuples, one per harness in the config.
        Multiply each by seg['quantity'] for total counts.
        """
        template_ref = seg.get('template_ref')
        harness_config = self._get_effective_harness_config(seg)
        harness_sizes = self.parse_harness_config(harness_config)
        spt = seg['strings_per_tracker']
        
        if not harness_sizes:
            return []
        
        m_to_ft = 3.28084
        
        # Get template geometry
        tracker_length_ft = None
        string_length_ft = None
        motor_y_ft = None
        motor_gap_ft = 0
        has_motor = False
        
        if template_ref and template_ref in self.enabled_templates:
            tdata = self.enabled_templates[template_ref]
            mod_spec = tdata.get('module_spec', {})
            orientation = tdata.get('module_orientation', 'Portrait')
            mps = tdata.get('modules_per_string', 28)
            spacing_m = tdata.get('module_spacing_m', 0.02)
            has_motor = tdata.get('has_motor', True)
            motor_gap_m = tdata.get('motor_gap_m', 1.0) if has_motor else 0
            
            if orientation == 'Portrait':
                mod_along_m = mod_spec.get('width_mm', 1000) / 1000
            else:
                mod_along_m = mod_spec.get('length_mm', 2000) / 1000
            
            string_length_ft = (mps * mod_along_m + (mps - 1) * spacing_m) * m_to_ft
            motor_gap_ft = motor_gap_m * m_to_ft
            
            # Total tracker length — use full strings + partial modules
            full_spt = int(spt)
            partial_mods = round((spt - full_spt) * mps) if spt != full_spt else 0
            total_modules = full_spt * mps + partial_mods
            tracker_length_m = (total_modules * mod_along_m + 
                               (total_modules - 1) * spacing_m + 
                               (motor_gap_m if has_motor else 0))
            tracker_length_ft = tracker_length_m * m_to_ft
            
            # Motor Y position from north end
            if has_motor:
                motor_placement = tdata.get('motor_placement_type', 'between_strings')
                motor_pos_after = tdata.get('motor_position_after_string', None)
                motor_string_idx = tdata.get('motor_string_index', None)
                motor_split_north = tdata.get('motor_split_north', mps // 2)
                
                # Partial string on north adds height before everything
                partial_north_mods = 0
                spt_val = tdata.get('strings_per_tracker', 1)
                if spt_val != int(spt_val) and tdata.get('partial_string_side', 'north') == 'north':
                    partial_north_mods = round((spt_val - int(spt_val)) * mps)
                partial_north_m = partial_north_mods * (mod_along_m + spacing_m) if partial_north_mods > 0 else 0
                
                # Partial string on north shifts motor further south
                partial_north_mods_ext = 0
                if spt != int(spt):
                    tdata_ext = self.enabled_templates.get(template_ref, {})
                    if tdata_ext.get('partial_string_side', 'north') == 'north':
                        partial_north_mods_ext = round((spt - int(spt)) * mps)
                partial_north_m_ext = partial_north_mods_ext * (mod_along_m + spacing_m)
                
                if motor_placement == 'between_strings':
                    if motor_pos_after is not None:
                        strings_before = motor_pos_after
                    else:
                        strings_before = 1  # default: after first string
                    modules_before = strings_before * mps
                    motor_y_m = partial_north_m + (modules_before * mod_along_m + 
                                (modules_before - 1) * spacing_m + spacing_m)
                    motor_y_ft = motor_y_m * m_to_ft
                elif motor_placement == 'middle_of_string':
                    if motor_string_idx is not None:
                        strings_before = motor_string_idx - 1  # 1-based
                    else:
                        strings_before = 0
                    modules_before = strings_before * mps + motor_split_north
                    motor_y_m = partial_north_m + (modules_before * mod_along_m + 
                                max(modules_before - 1, 0) * spacing_m + spacing_m)
                    motor_y_ft = motor_y_m * m_to_ft
                else:
                    motor_y_ft = tracker_length_ft / 2
            else:
                motor_y_ft = tracker_length_ft / 2
        
        # Fallback if template not found
        if tracker_length_ft is None:
            module_width_ft = (self.selected_module.width_mm / 304.8) if self.selected_module else 3.3
            try:
                mps = int(self.modules_per_string_var.get())
            except ValueError:
                mps = 28
            string_length_ft = module_width_ft * mps
            motor_gap_ft = 3.28
            tracker_length_ft = string_length_ft * spt + motor_gap_ft
            motor_y_ft = tracker_length_ft / 2
            has_motor = True
        
        # Determine extender target Y based on device position
        device_offset_ft = 5.0
        if device_position == 'north':
            target_y = -device_offset_ft
        elif device_position == 'south':
            target_y = tracker_length_ft + device_offset_ft
        else:  # middle
            target_y = motor_y_ft
        
        # Resolve polarity convention
        polarity = self.polarity_convention_var.get()
        
        # Determine which absolute string index the motor gap follows
        motor_after_string = None  # 0-based absolute string index
        if has_motor and template_ref and template_ref in self.enabled_templates:
            tdata_motor = self.enabled_templates[template_ref]
            motor_placement = tdata_motor.get('motor_placement_type', 'between_strings')
            if motor_placement == 'between_strings':
                pos_after = tdata_motor.get('motor_position_after_string', None)
                if pos_after is not None:
                    motor_after_string = pos_after - 1 if pos_after > 0 else 0
                else:
                    motor_after_string = 0
            elif motor_placement == 'middle_of_string':
                # Motor is inside a string — gap is effectively after (string_index - 1)
                # but for extender purposes, treat it as between string boundaries
                idx = tdata_motor.get('motor_string_index', 1)
                motor_after_string = idx - 1 if idx > 0 else 0
        elif has_motor:
            motor_after_string = 0  # default fallback
        
        # Build string boundary positions N→S
        string_positions = []  # list of (north_edge, south_edge, harness_idx)
        y_cursor = 0.0
        abs_string_idx = 0
        inter_string_gap = (spacing_m if (template_ref and template_ref in self.enabled_templates) else 0.02) * m_to_ft
        
        for h_idx, h_size in enumerate(harness_sizes):
            for s in range(h_size):
                north_edge = y_cursor
                south_edge = y_cursor + string_length_ft
                string_positions.append((north_edge, south_edge, h_idx))
                y_cursor = south_edge
                
                # Add motor gap after the correct absolute string
                if has_motor and motor_after_string is not None and abs_string_idx == motor_after_string:
                    y_cursor += motor_gap_ft
                elif abs_string_idx < spt - 1:
                    # Add inter-string spacing (not after last string)
                    y_cursor += inter_string_gap
                
                abs_string_idx += 1
        
        # Calculate per-harness extender lengths
        result = []
        for h_idx, h_size in enumerate(harness_sizes):
            harness_strings = [(n, s) for n, s, hi in string_positions if hi == h_idx]
            
            if not harness_strings:
                result.append((10.0, 10.0))
                continue
            
            harness_north = harness_strings[0][0]
            harness_south = harness_strings[-1][1]
            
            # Determine which end is positive and which is negative
            if polarity == 'Negative Always South':
                pos_y = harness_north
                neg_y = harness_south
            elif polarity == 'Negative Always North':
                pos_y = harness_south
                neg_y = harness_north
            elif polarity == 'Negative Toward Device':
                if device_position == 'north':
                    neg_y = harness_north
                    pos_y = harness_south
                elif device_position == 'south':
                    neg_y = harness_south
                    pos_y = harness_north
                else:
                    if harness_north < motor_y_ft:
                        neg_y = harness_south
                        pos_y = harness_north
                    else:
                        neg_y = harness_north
                        pos_y = harness_south
            elif polarity == 'Positive Toward Device':
                if device_position == 'north':
                    pos_y = harness_north
                    neg_y = harness_south
                elif device_position == 'south':
                    pos_y = harness_south
                    neg_y = harness_north
                else:
                    if harness_north < motor_y_ft:
                        pos_y = harness_south
                        neg_y = harness_north
                    else:
                        pos_y = harness_north
                        neg_y = harness_south
            else:
                pos_y = harness_north
                neg_y = harness_south
            
            pos_extender = max(abs(pos_y - target_y), 5.0)
            neg_extender = max(abs(neg_y - target_y), 5.0)
            
            result.append((pos_extender, neg_extender))

        return result
        
    def load_estimate(self):
        """Load estimate data from the project"""
        if not self.current_project or not self.estimate_id:
            return
        
        estimate_data = self.current_project.quick_estimates.get(self.estimate_id)
        if not estimate_data:
            return
        
        self._loading = True
        
        # Module is now derived from templates — skip old module dropdown restore
        # (module will be set by _derive_module_from_templates after groups load)
        
        # Restore numeric fields
        # modules_per_string: store saved value as fallback; templates will override if linked
        self._saved_modules_per_string = estimate_data.get('modules_per_string', 28)
        if hasattr(self, 'modules_per_string_var'):
            self.modules_per_string_var.set(str(self._saved_modules_per_string))
        if hasattr(self, 'row_spacing_var'):
            self.row_spacing_var.set(str(estimate_data.get('row_spacing_ft', 20.0)))
        # Load wire sizing (new format) or migrate from old wire_gauge
        saved_wire_sizing = estimate_data.get('wire_sizing')
        if saved_wire_sizing:
            self.wire_sizing = copy.deepcopy(saved_wire_sizing)
        else:
            # Backward compat: old estimates had a single wire_gauge field
            self.wire_sizing = {
                'temp_rating': '90C',
                'feeder_material': 'aluminum',
                'by_string_count': {},
                'dc_feeder': '',
                'ac_homerun': '',
                'user_overrides': {}
            }
        
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
        if hasattr(self, 'breaker_size_var'):
            self.breaker_size_var.set(estimate_data.get('breaker_size', '400'))
        if hasattr(self, 'dc_feeder_distance_var'):
            self.dc_feeder_distance_var.set(str(estimate_data.get('dc_feeder_distance', 500)))
        if hasattr(self, 'polarity_convention_var'):
            self.polarity_convention_var.set(estimate_data.get('polarity_convention', 'Negative Always South'))
        if hasattr(self, 'lv_collection_var'):
            self.lv_collection_var.set(estimate_data.get('lv_collection_method', 'Wire Harness'))
        if hasattr(self, 'ac_homerun_distance_var'):
            self.ac_homerun_distance_var.set(str(estimate_data.get('ac_homerun_distance', 500)))
        
        # Load groups (new format) or convert from old subarrays format
        saved_subarrays = estimate_data.get('subarrays', {})
        saved_groups = estimate_data.get('groups', estimate_data.get('rows', []))
        
        self.groups.clear()
        self._refresh_group_listbox()
        
        if saved_groups:
            # Ensure all segments have template_ref field (backward compat)
            for group in saved_groups:
                for seg in group.get('segments', []):
                    if 'template_ref' not in seg:
                        seg['template_ref'] = None
            self.groups = copy.deepcopy(saved_groups)
        elif saved_subarrays:
            # Backward compat: convert old subarray/block format to groups
            group_num = 0
            for subarray_id, subarray_data in saved_subarrays.items():
                for block_id, block_data in subarray_data.get('blocks', {}).items():
                    group_num += 1
                    segments = []
                    for tracker in block_data.get('trackers', []):
                        segments.append({
                            'quantity': tracker.get('quantity', 1),
                            'strings_per_tracker': tracker.get('strings', 3),
                            'harness_config': tracker.get('harness_config', '3'),
                            'template_ref': None
                        })
                    if not segments:
                        segments = [{'quantity': 1, 'strings_per_tracker': 3, 'harness_config': '3', 'template_ref': None}]
                    self.groups.append({
                        'name': block_data.get('name', f"Group {group_num}"),
                        'segments': segments
                    })
        else:
            self._loading = False
            self.add_group()
            return
        
        self._refresh_group_listbox()
        
        # Select first group
        if self.groups:
            self.group_listbox.selection_set(0)
            self.on_group_select(None)
        
        self._last_inspect_mode = estimate_data.get('inspect_mode', False)
        if hasattr(self, 'use_routed_var'):
                    self.use_routed_var.set(estimate_data.get('use_routed_distances', False))
        self.pads = copy.deepcopy(estimate_data.get('pads', []))
        
        # Load device names (convert str keys back to int)
        saved_names = estimate_data.get('device_names', {})
        self.device_names = {int(k): v for k, v in saved_names.items()}
        
        # Load allocation lock state
        self.allocation_locked = estimate_data.get('allocation_locked', False)
        self.locked_allocation_result = estimate_data.get('locked_allocation_result', None)

        # Derive module from templates
        self._derive_module_from_templates()

        # Refresh wire sizing table — always reconcile saved data with current segments
        # This adds missing string counts and removes stale ones
        self._refresh_wire_sizing_for_segments()

        # Re-enable autosave now that loading is complete
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
            estimate_data['module_name'] = f"{self.selected_module.manufacturer} {self.selected_module.model} ({self.selected_module.wattage}W)"
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

        estimate_data['wire_sizing'] = copy.deepcopy(self.wire_sizing)
        
        # Save inverter selection
        if self.selected_inverter:
            estimate_data['inverter_name'] = self.inverter_select_var.get()
        
        # Save topology and DC:AC ratio
        estimate_data['topology'] = self.topology_var.get()
        estimate_data['inspect_mode'] = getattr(self, '_last_inspect_mode', False)
        estimate_data['use_routed_distances'] = self.use_routed_var.get()
        estimate_data['breaker_size'] = self.breaker_size_var.get()
        estimate_data['polarity_convention'] = self.polarity_convention_var.get()
        estimate_data['lv_collection_method'] = self.lv_collection_var.get()
        try:
            estimate_data['dc_feeder_distance'] = float(self.dc_feeder_distance_var.get())
        except (ValueError, AttributeError):
            pass
        try:
            estimate_data['ac_homerun_distance'] = float(self.ac_homerun_distance_var.get())
        except (ValueError, AttributeError):
            pass
        try:
            estimate_data['dc_ac_ratio'] = float(self.dc_ac_ratio_var.get())
        except ValueError:
            estimate_data['dc_ac_ratio'] = 1.25
        
        # Update modified date
        estimate_data['modified_date'] = datetime.now().isoformat()
        
        # Save groups (new format) — deep copy to avoid reference aliasing
        estimate_data['groups'] = copy.deepcopy(self.groups)
        estimate_data['subarrays'] = {}
        
        # Save pads
        estimate_data['pads'] = copy.deepcopy(self.pads)
        
        # Save device names (convert int keys to str for JSON)
        estimate_data['device_names'] = {str(k): v for k, v in self.device_names.items()}
        
        # Save allocation lock state
        estimate_data['allocation_locked'] = self.allocation_locked
        if self.allocation_locked and self.locked_allocation_result is not None:
            estimate_data['locked_allocation_result'] = copy.deepcopy(self.locked_allocation_result)
        else:
            estimate_data['locked_allocation_result'] = None
        
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
            'row_spacing_ft': 20.0,
            'topology': 'Distributed String',
            'dc_ac_ratio': 1.25,
            'subarrays': {},
            'groups': [],
            'pads': []
        }
        
        self.current_project.quick_estimates[estimate_id] = new_estimate
        
        self.estimate_id = estimate_id
        self._refresh_estimate_dropdown()
        self.estimate_var.set(estimate_name)
        
        # Clear and set up fresh
        self._clear_estimate_ui()
        self.add_group()
        self._derive_module_from_templates()
        
        if self.on_save:
            self.on_save()

    def copy_estimate(self):
        """Duplicate the currently selected estimate"""
        if not self.current_project or not self.estimate_id:
            from tkinter import messagebox
            messagebox.showinfo("No Estimate", "No estimate selected to copy.")
            return
        
        # Save current estimate first so we capture latest state
        self.save_estimate()
        
        # Get current estimate data
        source_data = self.current_project.quick_estimates.get(self.estimate_id)
        if not source_data:
            return
        
        # Deep copy the estimate data
        new_data = copy.deepcopy(source_data)
        
        # Generate new ID and name
        new_id = f"estimate_{uuid.uuid4().hex[:8]}"
        source_name = source_data.get('name', 'Unnamed')
        new_data['name'] = f"{source_name} (Copy)"
        new_data['created_date'] = datetime.now().isoformat()
        new_data['modified_date'] = datetime.now().isoformat()
        
        # Store in project
        self.current_project.quick_estimates[new_id] = new_data
        
        # Switch to the new copy
        self.estimate_id = new_id
        self._refresh_estimate_dropdown()
        self.estimate_var.set(new_data['name'])
        self.load_estimate()
        
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

    def _auto_unlock_allocation(self):
        """Unlock allocation if locked, with a user notification.
        
        Called when structural changes (inverter, topology, segments, groups)
        invalidate a locked allocation.
        """
        if not self.allocation_locked:
            return
        
        self.allocation_locked = False
        self.locked_allocation_result = None
        
        from tkinter import messagebox
        messagebox.showinfo(
            "Allocation Unlocked",
            "The allocation lock has been released because the estimate structure changed.\n\n"
            "Re-run Calculate Estimate and lock again from the Site Preview if needed.",
            parent=self
        )

    def _clear_estimate_ui(self):
        """Clear the groups and details when switching/deleting estimates"""
        # Clear groups and pads
        self.groups.clear()
        self.pads.clear()
        self.device_names.clear()
        self.last_combiner_assignments = []
        self.allocation_locked = False
        self.locked_allocation_result = None
        if hasattr(self, 'group_listbox'):
            self._refresh_group_listbox()
        
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
        if hasattr(self, 'lv_collection_var'):
            self.lv_collection_var.set('Wire Harness')
        if hasattr(self, 'breaker_size_var'):
            self.breaker_size_var.set('400')
        if hasattr(self, 'dc_ac_ratio_var'):
            self.dc_ac_ratio_var.set('1.25')
        if hasattr(self, 'strings_per_inverter_var'):
            self.strings_per_inverter_var.set('--')
        if hasattr(self, 'isc_warning_label'):
            self.isc_warning_label.config(text="")

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

    def _on_lv_collection_changed(self, *args):
        """Handle LV Collection Method change — update harness combo states in details panel"""
        lv_method = self.lv_collection_var.get()
        
        # Update any currently displayed harness combos
        if hasattr(self, '_harness_combos'):
            for combo, harness_var, segment in self._harness_combos:
                try:
                    if lv_method == 'String HR':
                        spt = segment.get('strings_per_tracker', 1)
                        override_display = '+'.join(['1'] * int(spt))
                        original_config = segment['harness_config']
                        harness_var.set(override_display)
                        segment['harness_config'] = original_config  # restore saved config
                        combo.config(state='disabled')
                    elif lv_method == 'Trunk Bus':
                        combo.config(state='disabled')
                    else:
                        # Wire Harness — restore the real config and re-enable
                        harness_var.set(segment['harness_config'])
                        combo.config(state='readonly')
                except tk.TclError:
                    pass  # Widget may have been destroyed

        # Refresh wire sizing table since string counts may have changed
        self._refresh_wire_sizing_for_segments()

    def _on_module_selected(self, event=None):
        """Legacy — module is now derived from templates. Kept for backward compat."""
        pass

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
            self._on_inverter_changed_wire_sizing()
            # Auto-save when inverter changes (but not during load)
            if not getattr(self, '_loading', False):
                self._auto_unlock_allocation()
                self._mark_stale()
                self.save_estimate()
        else:
            self.selected_inverter = None
            self.inverter_info_label.config(text="No inverter selected", foreground='gray')
            self.strings_per_inverter_var.set('--')
            self.isc_warning_label.config(text="")

    INVERTER_COLORS = [
        '#4A90D9',  # Blue
        '#2ECC71',  # Green
        '#9B59B6',  # Purple
        '#E74C3C',  # Red
        '#1ABC9C',  # Teal
        '#3498DB',  # Light Blue
        '#E91E63',  # Pink
        '#00BCD4',  # Cyan
        '#8BC34A',  # Light Green
        '#795548',  # Brown
        '#607D8B',  # Blue Gray
        '#CDDC39',  # Lime
        '#5C6BC0',  # Indigo
        '#26A69A',  # Dark Teal
        '#78909C',  # Steel
    ]

    def show_site_preview(self):
        """Open the site preview in a pop-out window"""
        inv_summary = getattr(self, 'last_totals', {}).get('inverter_summary', {})
        
        if not inv_summary or not inv_summary.get('allocation_result'):
            from tkinter import messagebox
            messagebox.showinfo("No Data", "Run Calculate Estimate first to generate preview data.")
            return
        
        topology = self.topology_var.get()
        try:
            row_spacing_ft = float(self.row_spacing_var.get())
        except ValueError:
            row_spacing_ft = 20.0
        
        # Compute device info for preview
        totals = getattr(self, 'last_totals', {})
        total_combiners = sum(totals.get('combiners_by_breaker', {}).values())
        
        if topology == 'Distributed String':
            num_devices = totals.get('string_inverters', 0)
            device_label = 'SI'
        else:
            # Central Inverter and Centralized String show combiner boxes
            num_devices = total_combiners
            device_label = 'CB'
        
        # Restore inspect mode from previous session
        initial_inspect = getattr(self, '_last_inspect_mode', False)
        
        preview = SitePreviewWindow(
            self, inv_summary, topology, self.INVERTER_COLORS,
            self.groups, self.enabled_templates, row_spacing_ft,
            num_devices=num_devices, device_label=device_label,
            initial_inspect=initial_inspect, pads=self.pads,
            device_names=self.device_names
        )
        
        # When window closes, save state back
        def _on_preview_close():
            self._last_inspect_mode = preview.inspect_mode
            self.pads = preview.pads  # Save pad positions back
            self.device_names = dict(preview.device_names)  # Save renamed devices back
            
            # If CB assignments were edited, refresh the estimate results
            if hasattr(self, 'last_combiner_assignments') and self.last_combiner_assignments:
                self._refresh_combiner_results_from_assignments()
            
            self._schedule_autosave()
            preview.destroy()
        preview.protocol("WM_DELETE_WINDOW", _on_preview_close)

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
        
        use_routed = self.use_routed_var.get() if hasattr(self, 'use_routed_var') else False
        
        
        if use_routed:
            if not self.pads:
                self.distance_hint_label.config(
                    text="⚠ No pads placed — add pads in Site Preview first",
                    foreground='orange'
                )
            else:
                num_pads = len(self.pads)
                self.distance_hint_label.config(
                    text=f"✓ Using routed distances from {num_pads} pad{'s' if num_pads > 1 else ''} (avg inputs ignored)",
                    foreground='green'
                )
            return
        
        topology = self.topology_var.get()
        if topology == 'Distributed String':
            self.distance_hint_label.config(text="(DC feeders N/A for distributed — AC homeruns are primary cable)", foreground='gray')
        elif topology == 'Central Inverter':
            self.distance_hint_label.config(text="(Long DC feeders to central pad — short AC from inverter)", foreground='gray')
        elif topology == 'Centralized String':
            self.distance_hint_label.config(text="(DC feeders to inverter bank — short AC from bank)", foreground='gray')
        else:
            self.distance_hint_label.config(text="", foreground='gray')

    def _get_estimate_tracker_dims_ft(self, template_ref):
        """Get (width_ft, length_ft) for a tracker from its template reference.
        
        Width = E-W dimension (across tracker row).
        Length = N-S dimension (along tracker).
        
        Returns (width_ft, length_ft) or None if template not found.
        """
        if not template_ref or template_ref not in self.enabled_templates:
            return None
        
        tdata = self.enabled_templates[template_ref]
        ms = tdata.get('module_spec', {})
        mps = tdata.get('modules_per_string', 28)
        strings_per_tracker = tdata.get('strings_per_tracker', 1)
        orientation = tdata.get('module_orientation', 'Portrait')
        module_spacing_m = tdata.get('module_spacing_m', 0.02)
        
        mod_w_m = ms.get('width_mm', 1134) / 1000.0
        mod_l_m = ms.get('length_mm', 2278) / 1000.0
        
        if orientation == 'Portrait':
            mod_across = mod_l_m   # E-W (width of tracker)
            mod_along = mod_w_m    # N-S (length of tracker)
        else:
            mod_across = mod_w_m
            mod_along = mod_l_m
        
        # Width (E-W): module across dimension × modules_high (stacked columns)
        modules_high = tdata.get('modules_high', 1)
        width_m = mod_across * modules_high
        
        # Length (N-S): all modules laid end-to-end (full strings + partial) plus gaps and motor
        full_spt = int(strings_per_tracker)
        partial_mods = round((strings_per_tracker - full_spt) * mps) if strings_per_tracker != full_spt else 0
        modules_in_row = full_spt * mps + partial_mods
        
        motor_gap_m = tdata.get('motor_gap_m', 0)
        has_motor = tdata.get('has_motor', True)
        if not has_motor:
            motor_gap_m = 0
        
        length_m = modules_in_row * mod_along + max(modules_in_row - 1, 0) * module_spacing_m + motor_gap_m

        width_ft = width_m * 3.28084
        length_ft = length_m * 3.28084
        
        return (width_ft, length_ft)

    def _get_harness_config_for_tracker_type(self, strings_per_tracker):
        """Find the harness config used for trackers with the given string count."""
        for group in self.groups:
            for seg in group['segments']:
                if seg['strings_per_tracker'] == strings_per_tracker and seg['quantity'] > 0:
                    return self.parse_harness_config(self._get_effective_harness_config(seg))
                
        # Fallback: single harness equal to full string count
        return [int(strings_per_tracker)]

    def _build_combiner_assignments(self, totals, topology):
        """Build structured combiner box assignments for Device Configurator.
        
        Produces self.last_combiner_assignments — a list of dicts, one per CB,
        each containing the tracker/harness connections that feed it.
        
        For Centralized String: 1 CB per inverter, connections from allocation harness_map.
        For Central Inverter: distribute trackers proportionally across N combiners.
        For Distributed String: no CBs — returns empty list.
        """
        self.last_combiner_assignments = []
        
        if topology == 'Distributed String':
            return
        
        inv_summary = totals.get('inverter_summary', {})
        allocation_result = inv_summary.get('allocation_result')
        module_isc = self.selected_module.isc if self.selected_module else 0
        
        try:
            breaker_size = int(self.breaker_size_var.get())
        except (ValueError, AttributeError):
            breaker_size = 400
        
        # NEC factor — use project setting if available
        nec_factor = 1.56
        if self.current_project:
            nec_factor = getattr(self.current_project, 'nec_safety_factor', 1.56)
        
        # Standard breaker sizes for per-CB calculation
        BREAKER_SIZES = [100, 125, 150, 175, 200, 225, 250, 300, 350, 400, 450, 500, 600, 700, 800]
        
        # Build flat tracker list with segment metadata for harness config lookup
        tracker_segment_map = []  # flat list: one entry per tracker with its segment ref
        for group in self.groups:
            for seg in group['segments']:
                harness_config_str = self._get_effective_harness_config(seg)
                harness_sizes = self.parse_harness_config(harness_config_str)
                for _ in range(seg['quantity']):
                    tracker_segment_map.append({
                        'spt': seg['strings_per_tracker'],
                        'harness_sizes': list(harness_sizes),
                        'wire_gauge': self._get_wire_gauge_for_segment(seg, 'whip'),
                    })
        
        if topology == 'Centralized String' and allocation_result:
            # 1 CB per inverter — connections directly from allocation harness_map
            for inv_idx, inv in enumerate(allocation_result['inverters']):
                cb_name = self.device_names.get(inv_idx, f"CB-{inv_idx + 1:02d}")
                connections = self._build_connections_from_harness_map(
                    inv['harness_map'], tracker_segment_map, module_isc, nec_factor
                )
                # Calculate per-CB breaker from actual total current
                total_current = sum(
                    c['num_strings'] * c['module_isc'] * c['nec_factor']
                    for c in connections
                )
                calc_breaker = breaker_size  # fallback to global
                for bs in BREAKER_SIZES:
                    if bs >= total_current:
                        calc_breaker = bs
                        break
                
                self.last_combiner_assignments.append({
                    'combiner_name': cb_name,
                    'device_idx': inv_idx,
                    'breaker_size': calc_breaker,
                    'module_isc': module_isc,
                    'nec_factor': nec_factor,
                    'connections': connections,
                })
        
        elif topology == 'Central Inverter':
            total_combiners = sum(totals.get('combiners_by_breaker', {}).values())
            if total_combiners <= 0 or not tracker_segment_map:
                return
            
            # Distribute trackers across combiners proportionally
            # (same logic as _compute_devices_proportional but at the QE level)
            num_trackers = len(tracker_segment_map)
            
            # Find max inputs for the matched CB
            fuse_current = module_isc * nec_factor * max(
                (max(t['harness_sizes']) if t['harness_sizes'] else 1)
                for t in tracker_segment_map
            )
            fuse_holder_rating = self.get_fuse_holder_category(fuse_current)
            combiner_library = self.load_combiner_library()
            
            matching_cbs = [
                cb_data for cb_data in combiner_library.values()
                if (cb_data.get('breaker_size', 0) == breaker_size and
                    cb_data.get('fuse_holder_rating', '') == fuse_holder_rating)
            ]
            if matching_cbs:
                matching_cbs.sort(key=lambda c: c.get('max_inputs', 0), reverse=True)
                max_inputs = matching_cbs[0].get('max_inputs', 24)
            else:
                max_inputs = 24
            
            # Simple even distribution of trackers across CBs
            base_per_cb = num_trackers // total_combiners
            extra = num_trackers % total_combiners
            
            tracker_cursor = 0
            for cb_idx in range(total_combiners):
                count = base_per_cb + (1 if cb_idx < extra else 0)
                if count <= 0:
                    continue
                
                cb_name = self.device_names.get(cb_idx, f"CB-{cb_idx + 1:02d}")
                connections = []
                
                for local_i in range(count):
                    tidx = tracker_cursor + local_i
                    if tidx >= num_trackers:
                        break
                    t_info = tracker_segment_map[tidx]
                    
                    # Build one connection per harness in this tracker
                    for h_idx, h_size in enumerate(t_info['harness_sizes']):
                        connections.append({
                            'tracker_idx': tidx,
                            'tracker_label': f"T{tidx + 1:02d}",
                            'harness_label': f"H{h_idx + 1:02d}",
                            'num_strings': h_size,
                            'module_isc': module_isc,
                            'nec_factor': nec_factor,
                            'wire_gauge': t_info['wire_gauge'],
                        })
                
                # Calculate per-CB breaker from actual total current
                total_current = sum(
                    c['num_strings'] * c['module_isc'] * c['nec_factor']
                    for c in connections
                )
                calc_breaker = breaker_size  # fallback to global
                for bs in BREAKER_SIZES:
                    if bs >= total_current:
                        calc_breaker = bs
                        break
                
                self.last_combiner_assignments.append({
                    'combiner_name': cb_name,
                    'device_idx': cb_idx,
                    'breaker_size': calc_breaker,
                    'module_isc': module_isc,
                    'nec_factor': nec_factor,
                    'connections': connections,
                })
                
                tracker_cursor += count

    def _read_combiner_bom_from_device_config(self):
        """Read combiner BOM data from Device Configurator (single source of truth).
        
        Populates self.last_totals['combiners_by_breaker'] and ['combiner_details']
        from the Device Configurator's combiner_configs, which reflect the user's
        NEC multiplier, breaker overrides, and fuse sizes.
        
        Also syncs breaker sizes back into self.last_combiner_assignments.
        
        Returns True if Device Configurator had data, False if fallback is needed.
        """
        if not hasattr(self, 'last_totals') or not self.last_totals:
            return False
        
        # Get Device Configurator
        main_app = getattr(self, 'main_app', None)
        if not main_app or not hasattr(main_app, 'device_configurator'):
            return False
        
        dc = main_app.device_configurator
        if not hasattr(dc, 'combiner_configs') or not dc.combiner_configs:
            return False
        
        # Only use DC data if it's in QE mode
        if getattr(dc, 'data_source', 'blocks') != 'quick_estimate':
            return False
        
        totals = self.last_totals
        totals['combiners_by_breaker'] = {}
        totals['combiner_details'] = []
        
        # Sync breaker sizes back to last_combiner_assignments
        dc_breakers = {}
        for combiner_id, config in dc.combiner_configs.items():
            dc_breakers[combiner_id] = config.get_display_breaker_size()
        
        for cb in self.last_combiner_assignments:
            cb_name = cb.get('combiner_name', '')
            if cb_name in dc_breakers:
                cb['breaker_size'] = dc_breakers[cb_name]
        
        # Build totals from Device Configurator configs
        # Group by (breaker_size, fuse_holder_rating, max_inputs) for CB matching
        cb_groups = {}  # (breaker_size, fuse_holder_rating) -> {max_inputs, count}
        
        for combiner_id, config in dc.combiner_configs.items():
            if not config.connections:
                continue
            
            breaker_size = config.get_display_breaker_size()
            num_inputs = len(config.connections)
            
            # Get fuse holder rating from the DC's actual fuse sizes
            max_fuse = max(conn.get_display_fuse_size() for conn in config.connections)
            fuse_holder_rating = self.get_fuse_holder_category(max_fuse)
            
            # Count by breaker size
            if breaker_size not in totals['combiners_by_breaker']:
                totals['combiners_by_breaker'][breaker_size] = 0
            totals['combiners_by_breaker'][breaker_size] += 1
            
            # Group for CB part matching
            key = (breaker_size, fuse_holder_rating)
            if key not in cb_groups:
                cb_groups[key] = {'max_inputs': 0, 'count': 0}
            cb_groups[key]['max_inputs'] = max(cb_groups[key]['max_inputs'], num_inputs)
            cb_groups[key]['count'] += 1
        
        # Match CB parts for each group
        for (breaker_size, fuse_holder_rating), group_info in cb_groups.items():
            matched_cb = self.find_combiner_box(group_info['max_inputs'], breaker_size, fuse_holder_rating)
            if matched_cb:
                totals['combiner_details'].append({
                    'part_number': matched_cb.get('part_number', ''),
                    'description': matched_cb.get('description', ''),
                    'max_inputs': matched_cb.get('max_inputs', 0),
                    'breaker_size': breaker_size,
                    'fuse_holder_rating': fuse_holder_rating,
                    'strings_per_cb': group_info['max_inputs'],
                    'quantity': group_info['count'],
                    'block_name': 'Site Total'
                })
            else:
                totals['combiner_details'].append({
                    'part_number': 'NO MATCH',
                    'description': f'No CB found: {group_info["max_inputs"]} inputs, {breaker_size}A, {fuse_holder_rating}',
                    'max_inputs': group_info['max_inputs'],
                    'breaker_size': breaker_size,
                    'fuse_holder_rating': fuse_holder_rating,
                    'strings_per_cb': group_info['max_inputs'],
                    'quantity': group_info['count'],
                    'block_name': 'Site Total'
                })
        
        self.last_totals = totals
        return True

    def _rebuild_combiner_totals_from_assignments(self):
        """Fallback: rebuild combiner BOM totals from last_combiner_assignments.
        
        Used when the Device Configurator isn't available or isn't in QE mode.
        """
        if not self.last_combiner_assignments or not hasattr(self, 'last_totals') or not self.last_totals:
            return
        
        totals = self.last_totals
        totals['combiners_by_breaker'] = {}
        totals['combiner_details'] = []
        
        module_isc = self.selected_module.isc if self.selected_module else 0
        nec_factor = 1.56
        if self.current_project:
            nec_factor = getattr(self.current_project, 'nec_safety_factor', 1.56)
        
        for cb in self.last_combiner_assignments:
            bs = cb['breaker_size']
            if bs not in totals['combiners_by_breaker']:
                totals['combiners_by_breaker'][bs] = 0
            totals['combiners_by_breaker'][bs] += 1
        
        for bs, qty in totals['combiners_by_breaker'].items():
            max_h_strings = 1
            max_inputs = 0
            for cb in self.last_combiner_assignments:
                if cb['breaker_size'] == bs:
                    num_inputs = len(cb['connections'])
                    max_inputs = max(max_inputs, num_inputs)
                    for conn in cb['connections']:
                        max_h_strings = max(max_h_strings, conn['num_strings'])
            
            fuse_current = module_isc * nec_factor * max_h_strings
            fuse_holder_rating = self.get_fuse_holder_category(fuse_current)
            
            matched_cb = self.find_combiner_box(max_inputs, bs, fuse_holder_rating)
            if matched_cb:
                totals['combiner_details'].append({
                    'part_number': matched_cb.get('part_number', ''),
                    'description': matched_cb.get('description', ''),
                    'max_inputs': matched_cb.get('max_inputs', 0),
                    'breaker_size': bs,
                    'fuse_holder_rating': fuse_holder_rating,
                    'strings_per_cb': max_inputs,
                    'quantity': qty,
                    'block_name': 'Site Total'
                })
            else:
                totals['combiner_details'].append({
                    'part_number': 'NO MATCH',
                    'description': f'No CB found: {max_inputs} inputs, {bs}A, {fuse_holder_rating}',
                    'max_inputs': max_inputs,
                    'breaker_size': bs,
                    'fuse_holder_rating': fuse_holder_rating,
                    'strings_per_cb': max_inputs,
                    'quantity': qty,
                    'block_name': 'Site Total'
                })
        
        self.last_totals = totals

    def _refresh_combiner_results_from_assignments(self):
        """Rebuild combiner BOM totals and refresh display.
        
        Called after manual CB edits in the site preview to keep the estimate
        results in sync without re-running the full calculation.
        """
        if not self.last_combiner_assignments or not hasattr(self, 'last_totals') or not self.last_totals:
            return
        
        # Read from Device Configurator if available, else fallback
        if not self._read_combiner_bom_from_device_config():
            self._rebuild_combiner_totals_from_assignments()
        
        self._redraw_results_tree()

    def _redraw_results_tree(self):
        """Redraw the results treeview from self.last_totals without recalculating.
        
        This is a display-only refresh — it does NOT re-run calculate_estimate().
        """
        if not hasattr(self, 'last_totals') or not self.last_totals:
            return
        
        totals = self.last_totals
        
        # Clear existing results
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        self.checked_items.clear()
        
        def insert_section(label):
            self.results_tree.insert('', 'end', values=('', f'--- {label} ---', '', '', '', '', ''), tags=('section',))

        def insert_row(item, part_number, qty, unit, unit_cost='', ext_cost=''):
            iid = self.results_tree.insert('', 'end', values=('☑', item, part_number, qty, unit, unit_cost, ext_cost), tags=('checked',))
            self.checked_items.add(iid)
        
        # Combiner Boxes
        if totals.get('combiners_by_breaker'):
            insert_section('COMBINER BOXES')
            total_cbs = 0
            for breaker_size in sorted(totals['combiners_by_breaker'].keys()):
                qty = totals['combiners_by_breaker'][breaker_size]
                total_cbs += qty
                insert_row(f"Combiner Box ({breaker_size}A breaker)", '', qty, 'ea')
            if len(totals['combiners_by_breaker']) > 1:
                insert_row('Total Combiner Boxes', '', total_cbs, 'ea')
            
            for detail in totals.get('combiner_details', []):
                if detail['part_number'] != 'NO MATCH':
                    insert_row(
                        f"  └ Site Total: ({detail['max_inputs']}-input, {detail['fuse_holder_rating']})",
                        detail['part_number'], detail['quantity'], 'ea'
                    )
                else:
                    insert_row(
                        f"  └ Site Total: ⚠ {detail['description']}",
                        '', detail['quantity'], 'ea'
                    )
        
        # Inverters
        if totals.get('string_inverters', 0) > 0:
            insert_section('INVERTERS')
            inv_summary = totals.get('inverter_summary', {})
            actual_ratio = inv_summary.get('actual_dc_ac', 0)
            inv_name = "Inverter"
            if self.selected_inverter:
                inv_name = f"{self.selected_inverter.manufacturer} {self.selected_inverter.model}"
            insert_row(f"{inv_name} (DC:AC {actual_ratio:.2f})", '', totals['string_inverters'], 'ea')
        
        # Harnesses
        if totals.get('harnesses_by_size'):
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
        
        # Extenders — split by wire gauge, then by polarity
        ext_gauges = sorted(set(g for (_, g) in totals['extenders_pos_by_length'].keys()) |
                           set(g for (_, g) in totals['extenders_neg_by_length'].keys()))
        for gauge in ext_gauges:
            # Positive
            pos_items = {length: qty for (length, g), qty in totals['extenders_pos_by_length'].items() if g == gauge}
            if pos_items:
                insert_section(f'EXTENDERS — POSITIVE ({gauge})')
                for length in sorted(pos_items.keys()):
                    qty = pos_items[length]
                    e_pn, e_unit, e_ext = self.lookup_part_and_price('extender', polarity='positive', length_ft=length, qty=qty, num_strings=self._gauge_to_string_count('extender', gauge))
                    insert_row(f"Extender {length}ft (Pos)", e_pn, qty, 'ea', e_unit, e_ext)

            # Negative
            neg_items = {length: qty for (length, g), qty in totals['extenders_neg_by_length'].items() if g == gauge}
            if neg_items:
                insert_section(f'EXTENDERS — NEGATIVE ({gauge})')
                for length in sorted(neg_items.keys()):
                    qty = neg_items[length]
                    e_pn, e_unit, e_ext = self.lookup_part_and_price('extender', polarity='negative', length_ft=length, qty=qty, num_strings=self._gauge_to_string_count('extender', gauge))
                    insert_row(f"Extender {length}ft (Neg)", e_pn, qty, 'ea', e_unit, e_ext)

        # Whips — split by wire gauge, then by polarity
        whip_gauges = sorted(set(g for (_, g) in totals.get('whips_by_length', {}).keys()))
        if whip_gauges:
            display = totals.get('_display', {})
            topology = display.get('topology', self.topology_var.get() if hasattr(self, 'topology_var') else '')
            device_label = 'to inverter' if topology == 'Distributed String' else 'to combiner'

            for gauge in whip_gauges:
                gauge_items = {length: qty for (length, g), qty in totals['whips_by_length'].items() if g == gauge}
                if not gauge_items:
                    continue

                # Positive whips (half of each length bucket)
                insert_section(f'WHIPS — POSITIVE ({device_label}) ({gauge})')
                for length in sorted(gauge_items.keys()):
                    qty = gauge_items[length] // 2
                    w_pn, w_unit, w_ext = self.lookup_part_and_price('whip', polarity='positive', length_ft=length, qty=qty, num_strings=self._gauge_to_string_count('whip', gauge))
                    insert_row(f"Whip {length}ft (Pos)", w_pn, qty, 'ea', w_unit, w_ext)

                # Negative whips
                insert_section(f'WHIPS — NEGATIVE ({device_label}) ({gauge})')
                for length in sorted(gauge_items.keys()):
                    qty = gauge_items[length] // 2
                    w_pn, w_unit, w_ext = self.lookup_part_and_price('whip', polarity='negative', length_ft=length, qty=qty, num_strings=self._gauge_to_string_count('whip', gauge))
                    insert_row(f"Whip {length}ft (Neg)", w_pn, qty, 'ea', w_unit, w_ext)
        
        # DC Feeders
        if totals.get('dc_feeder_count', 0) > 0:
            display = totals.get('_display', {})
            use_routed = display.get('use_routed', False)
            dc_feeder_avg_ft = display.get('dc_feeder_avg_ft', 0)

            if totals['dc_feeder_count'] > 0:
                routed_dc_avg = totals['dc_feeder_total_ft'] / totals['dc_feeder_count']
            else:
                routed_dc_avg = dc_feeder_avg_ft
            dc_avg_display = routed_dc_avg if use_routed else dc_feeder_avg_ft
            dc_label_suffix = " (routed)" if use_routed else ""
            dc_wire_size = self.get_wire_size_for('dc_feeder')
            insert_section(f'DC FEEDERS ({dc_wire_size})')
            insert_row(f"DC Feeder {dc_wire_size} — avg {dc_avg_display:.0f}ft{dc_label_suffix} × {totals['dc_feeder_count']} runs (pos)", '', f"{totals['dc_feeder_total_ft']:.0f}", 'ft')
            insert_row(f"DC Feeder {dc_wire_size} — avg {dc_avg_display:.0f}ft{dc_label_suffix} × {totals['dc_feeder_count']} runs (neg)", '', f"{totals['dc_feeder_total_ft']:.0f}", 'ft')

        # AC Homeruns
        if totals.get('ac_homerun_count', 0) > 0:
            display = totals.get('_display', {})
            use_routed = display.get('use_routed', False)
            ac_homerun_avg_ft = display.get('ac_homerun_avg_ft', 0)

            if totals['ac_homerun_count'] > 0:
                routed_ac_avg = totals['ac_homerun_total_ft'] / totals['ac_homerun_count']
            else:
                routed_ac_avg = ac_homerun_avg_ft
            ac_avg_display = routed_ac_avg if use_routed else ac_homerun_avg_ft
            ac_label_suffix = " (routed)" if use_routed else ""
            ac_wire_size = self.get_wire_size_for('ac_homerun')
            insert_section(f'AC HOMERUNS ({ac_wire_size})')
            insert_row(f"AC Homerun {ac_wire_size} — avg {ac_avg_display:.0f}ft{ac_label_suffix} × {totals['ac_homerun_count']} runs", '', f"{totals['ac_homerun_total_ft']:.0f}", 'ft')
    
    def _build_connections_from_harness_map(self, harness_map, tracker_segment_map, module_isc, nec_factor):
        """Convert an allocation harness_map into Device Configurator connection dicts."""
        connections = []
        
        # Group harness_map entries by tracker to assign harness labels
        tracker_harness_counter = {}  # tracker_idx -> next harness number
        
        for entry in harness_map:
            tidx = entry['tracker_idx']
            strings_taken = entry['strings_taken']
            spt = entry['strings_per_tracker']
            is_split = entry.get('is_split', False)
            
            # Get segment info for this tracker
            t_info = tracker_segment_map[tidx] if tidx < len(tracker_segment_map) else None
            wire_gauge = t_info['wire_gauge'] if t_info else '10 AWG'
            original_harnesses = t_info['harness_sizes'] if t_info else [spt]
            
            if not is_split:
                # Full tracker — one connection per harness in the config
                for h_idx, h_size in enumerate(original_harnesses):
                    connections.append({
                        'tracker_idx': tidx,
                        'tracker_label': f"T{tidx + 1:02d}",
                        'harness_label': f"H{h_idx + 1:02d}",
                        'num_strings': h_size,
                        'module_isc': module_isc,
                        'nec_factor': nec_factor,
                        'wire_gauge': wire_gauge,
                    })
            else:
                # Split tracker — distribute strings_taken across harnesses
                remaining = strings_taken
                harness_cursor = tracker_harness_counter.get(tidx, 0)
                
                while remaining > 0 and harness_cursor < len(original_harnesses):
                    h_size = original_harnesses[harness_cursor]
                    take = min(remaining, h_size)
                    connections.append({
                        'tracker_idx': tidx,
                        'tracker_label': f"T{tidx + 1:02d}",
                        'harness_label': f"H{harness_cursor + 1:02d}",
                        'num_strings': take,
                        'module_isc': module_isc,
                        'nec_factor': nec_factor,
                        'wire_gauge': wire_gauge,
                    })
                    remaining -= take
                    if take >= h_size:
                        harness_cursor += 1
                
                tracker_harness_counter[tidx] = harness_cursor
        
        return connections
    
    def _get_wire_gauge_for_segment(self, seg, cable_type):
        """Get the wire gauge for a segment from the wire sizing table."""
        spt = seg.get('strings_per_tracker', 1)
        harness_sizes = self.parse_harness_config(self._get_effective_harness_config(seg))
        max_harness = max(harness_sizes) if harness_sizes else spt
        
        # Look up from self.wire_sizing
        if hasattr(self, 'wire_sizing') and self.wire_sizing:
            # Wire sizing is keyed by string count
            sizing = self.wire_sizing.get(str(max_harness)) or self.wire_sizing.get(str(spt))
            if sizing and cable_type in sizing:
                return sizing[cable_type]
        
        return '10 AWG'

    def _adjust_harnesses_for_splits(self, totals):
        """Adjust harness counts based on inverter allocation split trackers.
        
        Also builds self._split_tracker_details for use by whip and extender
        calculations — maps each split tracker to per-device harness assignments.
        """
        self._split_tracker_details = {}  # tracker_idx -> split info
        
        inv_summary = totals.get('inverter_summary', {})
        allocation_result = inv_summary.get('allocation_result')
        
        if not allocation_result:
            return
        
        # Collect split tracker info from harness_map
        split_trackers = {}
        
        for inv_idx, inv in enumerate(allocation_result['inverters']):
            for entry in inv['harness_map']:
                if entry['is_split']:
                    tidx = entry['tracker_idx']
                    if tidx not in split_trackers:
                        split_trackers[tidx] = []
                    # Tag with which inverter/device this portion belongs to
                    entry_with_inv = dict(entry)
                    entry_with_inv['inv_idx'] = inv_idx
                    split_trackers[tidx].append(entry_with_inv)
        
        if not split_trackers:
            return
                
        for tidx, entries in split_trackers.items():
            spt = entries[0]['strings_per_tracker']
            original_harness_sizes = self._get_harness_config_for_tracker_type(spt)
            
            # Sort portions largest-first for greedy harness distribution
            portions = sorted(entries, key=lambda e: e['strings_taken'], reverse=True)
            remaining_harnesses = sorted(original_harness_sizes, reverse=True)
                        
            # Build per-portion harness assignments
            portion_details = []
            for portion in portions:
                amount = portion['strings_taken']
                inv_idx = portion['inv_idx']
                assigned_harnesses = []
                assigned = 0
                
                while assigned < amount and remaining_harnesses:
                    h = remaining_harnesses[0]
                    if assigned + h <= amount:
                        # Whole harness fits in this portion
                        assigned_harnesses.append(h)
                        remaining_harnesses.pop(0)
                        assigned += h
                    else:
                        # Must split this harness at the boundary
                        needed = amount - assigned
                        assigned_harnesses.append(needed)
                        leftover = h - needed
                        remaining_harnesses.pop(0)
                        if leftover > 0:
                            remaining_harnesses.insert(0, leftover)
                        assigned = amount
                
                portion_details.append({
                    'inv_idx': inv_idx,
                    'strings_taken': amount,
                    'harnesses': assigned_harnesses,
                })
            
            # Any remaining harnesses (shouldn't happen but be safe)
            if remaining_harnesses:
                portion_details[-1]['harnesses'].extend(remaining_harnesses)
            
            # Store for whip and extender calculations
            self._split_tracker_details[tidx] = {
                'spt': spt,
                'original_config': original_harness_sizes,
                'portions': portion_details,
            }
            
            # Adjust harness totals: remove originals, add new
            for size in original_harness_sizes:
                if size in totals['harnesses_by_size']:
                    totals['harnesses_by_size'][size] -= 1
                    if totals['harnesses_by_size'][size] <= 0:
                        del totals['harnesses_by_size'][size]
            
            all_new = []
            for p in portion_details:
                all_new.extend(p['harnesses'])
            for size in all_new:
                if size not in totals['harnesses_by_size']:
                    totals['harnesses_by_size'][size] = 0
                totals['harnesses_by_size'][size] += 1
                    
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
        
        if topology == 'Distributed String':
            # Physical input count matters — strings connect directly to inverter
            input_limited = self.selected_inverter.get_total_string_capacity()
            strings_per_inv = min(power_based_strings, input_limited)
        else:
            # Centralized String or Central Inverter — combiners aggregate strings
            # Only power/ratio target applies, not physical input count
            strings_per_inv = power_based_strings
        
        strings_per_inv = max(strings_per_inv, 1)  # At least 1
        self._updating_spi = True

        self.strings_per_inverter_var.set(str(strings_per_inv))
        self._updating_spi = False
        
        # Calculate actual DC:AC ratio achieved
        actual_ratio = self.selected_inverter.dc_ac_ratio(
            strings_per_inv, module_wattage, modules_per_string
        )
        

        # Show warning if input limit is capping the ratio (Distributed String only)
        if topology == 'Distributed String' and strings_per_inv < power_based_strings:
            self.inverter_info_label.config(
                text=f"{self.selected_inverter.rated_power_kw}kW AC  |  Capped by string inputs ({self.selected_inverter.get_total_string_capacity()})  |  Max DC:AC ≈ {actual_ratio:.2f}",
                foreground='orange'
            )
        else:
            # No capping — reset to normal inverter info
            inv = self.selected_inverter
            type_str = inv.inverter_type.value if hasattr(inv, 'inverter_type') else 'String'
            self.inverter_info_label.config(
                text=f"{inv.rated_power_kw}kW AC  |  {inv.max_dc_power_kw}kW DC  |  {inv.get_total_string_capacity()} inputs  |  {type_str}",
                foreground='black'
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
        ttk.Button(top_bar, text="Copy", command=self.copy_estimate, width=6).pack(side='left', padx=2)
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
        
        # Global settings + Wire sizing side-by-side container
        settings_container = ttk.Frame(bottom_frame)
        settings_container.pack(fill='x', pady=(0, 10))
        
        settings_frame = ttk.LabelFrame(settings_container, text="Global Settings", padding="5")
        settings_frame.pack(side='left', fill='y')
        
        # Row 1: Module display (derived from templates)
        module_row = ttk.Frame(settings_frame)
        module_row.pack(fill='x', pady=(0, 5))
        
        ttk.Label(module_row, text="Module:", font=('Helvetica', 9, 'bold')).pack(side='left', padx=(0, 5))
        self.module_info_label = ttk.Label(module_row, text="No templates linked — link a tracker template to a segment", foreground='gray')
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
        self.topology_var.trace_add('write', lambda *args: (self._auto_unlock_allocation(), self._update_distance_hints(), self._on_topology_changed_wire_sizing(), self._mark_stale(), self._schedule_autosave()))
        
        ttk.Label(topology_row, text="DC:AC Ratio:").pack(side='left', padx=(0, 5))
        self.dc_ac_ratio_var = tk.StringVar(value='1.25')
        ttk.Spinbox(
            topology_row, from_=1.0, to=2.0, increment=0.05,
            textvariable=self.dc_ac_ratio_var, width=6, format='%.2f'
        ).pack(side='left', padx=(0, 15))
        self.dc_ac_ratio_var.trace_add('write', lambda *args: (self._auto_unlock_allocation(), self._update_strings_per_inverter(), self._mark_stale(), self._schedule_autosave()))
        
        ttk.Label(topology_row, text="LV Collection:").pack(side='left', padx=(0, 5))
        self.lv_collection_var = tk.StringVar(value='Wire Harness')
        lv_collection_combo = ttk.Combobox(
            topology_row,
            textvariable=self.lv_collection_var,
            values=['String HR', 'Wire Harness', 'Trunk Bus'],
            state='readonly',
            width=14
        )
        lv_collection_combo.pack(side='left', padx=(0, 15))
        self.disable_combobox_scroll(lv_collection_combo)
        self.lv_collection_var.trace_add('write', lambda *args: (self._on_lv_collection_changed(), self._mark_stale(), self._schedule_autosave()))

        ttk.Label(topology_row, text="Breaker Size:").pack(side='left', padx=(0, 5))
        self.breaker_size_var = tk.StringVar(value='400')
        breaker_combo = ttk.Combobox(
            topology_row,
            textvariable=self.breaker_size_var,
            values=['200', '300', '400', '600', '800'],
            state='readonly',
            width=6
        )
        breaker_combo.pack(side='left', padx=(0, 15))
        self.disable_combobox_scroll(breaker_combo)
        self.breaker_size_var.trace_add('write', lambda *args: (self._on_breaker_changed_wire_sizing(), self._mark_stale(), self._schedule_autosave()))

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
        
        # Modules per string — hidden var, derived from template
        self.modules_per_string_var = tk.StringVar(value=str(getattr(self, 'modules_per_string_default', 28)))
        self.modules_per_string_var.trace_add('write', lambda *args: (self._update_strings_per_inverter(), self._mark_stale(), self._schedule_autosave()))
        
        ttk.Label(settings_inner, text="Row Spacing (ft):").pack(side='left', padx=(0, 5))
        self.row_spacing_var = tk.StringVar(value=str(getattr(self, 'row_spacing_default', 20.0)))
        ttk.Spinbox(settings_inner, from_=1, to=100, textvariable=self.row_spacing_var, width=6).pack(side='left', padx=(0, 15))
        self.row_spacing_var.trace_add('write', lambda *args: (self._mark_stale(), self._schedule_autosave()))

        # Wire sizing dict — populated dynamically based on harness configs
        # Keys: 'temp_rating', 'feeder_material', 'by_string_count', 'dc_feeder', 'ac_homerun', 'user_overrides'
        self.wire_sizing = {
            'temp_rating': '90C',
            'feeder_material': 'aluminum',
            'by_string_count': {},
            'dc_feeder': '',
            'ac_homerun': '',
            'user_overrides': {}
        }
        
        ttk.Label(settings_inner, text="Polarity:").pack(side='left', padx=(10, 5))
        self.polarity_convention_var = tk.StringVar(value='Negative Always South')
        polarity_combo = ttk.Combobox(
            settings_inner,
            textvariable=self.polarity_convention_var,
            values=[
                'Negative Always South',
                'Negative Always North',
                'Negative Toward Device',
                'Positive Toward Device'
            ],
            state='readonly',
            width=22
        )
        polarity_combo.pack(side='left')
        self.disable_combobox_scroll(polarity_combo)
        self.polarity_convention_var.trace_add('write', lambda *args: (self._mark_stale(), self._schedule_autosave()))
        
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
        
        self.use_routed_var = tk.BooleanVar(value=False)
        self.use_routed_cb = ttk.Checkbutton(
            distance_row, text="Use Routed Distances",
            variable=self.use_routed_var,
            command=lambda: (self._mark_stale(), self._update_distance_hints(), self._schedule_autosave())
        )
        self.use_routed_cb.pack(side='left', padx=(10, 0))
        
        # Topology hint label
        self.distance_hint_label = ttk.Label(distance_row, text="", foreground='gray')
        self.distance_hint_label.pack(side='left', padx=(5, 0))
        self._update_distance_hints()
        
        # Wire Sizing frame (right side of Global Settings)
        self._build_wire_sizing_frame(settings_container)
        
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
        """Setup the left row list panel"""
        # Label
        tree_label = ttk.Label(parent, text="Tracker Groups", font=('Helvetica', 10, 'bold'))
        tree_label.pack(anchor='w', pady=(0, 5))
        
        # Listbox frame
        list_frame = ttk.Frame(parent)
        list_frame.pack(fill='both', expand=True)
        
        self.group_listbox = tk.Listbox(list_frame, selectmode='browse', font=('Helvetica', 10))
        self.group_listbox.pack(side='left', fill='both', expand=True)
        
        # Scrollbar
        list_scroll = ttk.Scrollbar(list_frame, orient='vertical', command=self.group_listbox.yview)
        list_scroll.pack(side='right', fill='y')
        self.group_listbox.configure(yscrollcommand=list_scroll.set)
        
        # Bind selection
        self.group_listbox.bind('<<ListboxSelect>>', self.on_group_select)
        
        # Buttons frame
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill='x', pady=(10, 0))
        
        ttk.Button(btn_frame, text="+ Group", command=self.add_group).pack(side='left', padx=(0, 5))
        ttk.Button(btn_frame, text="Copy", command=self.copy_selected_group).pack(side='left', padx=(0, 5))
        ttk.Button(btn_frame, text="Delete", command=self.delete_selected_group).pack(side='left', padx=(0, 5))
        
        # Move buttons
        move_frame = ttk.Frame(parent)
        move_frame.pack(fill='x', pady=(5, 0))
        ttk.Button(move_frame, text="▲ Up", command=self.move_group_up).pack(side='left', padx=(0, 5))
        ttk.Button(move_frame, text="▼ Down", command=self.move_group_down).pack(side='left')

    def setup_details_panel(self, parent):
        """Setup the right details panel"""
        # Details label
        self.details_label = ttk.Label(parent, text="Details", font=('Helvetica', 10, 'bold'))
        self.details_label.pack(anchor='w', pady=(0, 5))
        
        # Container for dynamic content
        self.details_container = ttk.Frame(parent)
        self.details_container.pack(fill='both', expand=True)
        
        # Placeholder
        self.placeholder_label = ttk.Label(self.details_container, text="Select a group to view details", foreground='gray')
        self.placeholder_label.pack(pady=20)

    def clear_details_panel(self):
        """Clear the details panel"""
        self._harness_combos = []
        for widget in self.details_container.winfo_children():
            widget.destroy()
        
        self.placeholder_label = ttk.Label(self.details_container, text="Select a group to view details", foreground='gray')
        self.placeholder_label.pack(pady=20)
        
        self.details_label.config(text="Details")

    def on_group_select(self, event):
        """Handle row selection changes"""
        if getattr(self, '_updating_listbox', False):
            return
        
        sel = self.group_listbox.curselection()
        if not sel:
            return
        
        new_idx = sel[0]
        
        # Sync the previous row's listbox text before switching (inline, no selection_set)
        if self.selected_group_idx is not None and self.selected_group_idx != new_idx and self.selected_group_idx < len(self.groups):
            prev = self.selected_group_idx
            group = self.groups[prev]
            total_trackers = sum(seg['quantity'] for seg in group['segments'])
            total_strings = sum(seg['quantity'] * seg['strings_per_tracker'] for seg in group['segments'])
            display = f"{group['name']}  ({total_trackers}T / {total_strings}S)"
            self._updating_listbox = True
            self.group_listbox.delete(prev)
            self.group_listbox.insert(prev, display)
            self.group_listbox.selection_clear(0, tk.END)
            self.group_listbox.selection_set(new_idx)
            self._updating_listbox = False
        
        if new_idx < 0 or new_idx >= len(self.groups):
            self.clear_details_panel()
            return
        
        # Clear existing details
        self._harness_combos = []
        for widget in self.details_container.winfo_children():
            widget.destroy()
        
        self.selected_group_idx = new_idx
        self.show_group_details(new_idx)

    def show_group_details(self, group_idx: int):
        """Show segment editor for the selected group"""
        if group_idx < 0 or group_idx >= len(self.groups):
            return
        
        group = self.groups[group_idx]
        self.details_label.config(text=f"Group: {group['name']}")
        
        # Create scrollable frame
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
        
        # Mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
        
        # Row name
        form_frame = ttk.Frame(scrollable_frame, padding="10")
        form_frame.pack(fill='x')
        
        ttk.Label(form_frame, text="Group Name:").grid(row=0, column=0, sticky='w', pady=5)
        name_var = tk.StringVar(value=group['name'])
        name_entry = ttk.Entry(form_frame, textvariable=name_var, width=25)
        name_entry.grid(row=0, column=1, sticky='w', pady=5, padx=(10, 0))
        
        def update_name(*args):
            group['name'] = name_var.get()
            self.details_label.config(text=f"Group: {name_var.get()}")
            self._schedule_autosave()
        name_var.trace_add('write', update_name)
        
        # Device Position dropdown
        ttk.Label(form_frame, text="Device Position:").grid(row=1, column=0, sticky='w', pady=5)
        device_pos_var = tk.StringVar(value=group.get('device_position', 'middle'))
        device_pos_combo = ttk.Combobox(form_frame, textvariable=device_pos_var,
                                         values=['north', 'middle', 'south'],
                                         state='readonly', width=10)
        device_pos_combo.grid(row=1, column=1, sticky='w', pady=5, padx=(10, 0))
        
        def on_device_position_change(*args):
            group['device_position'] = device_pos_var.get()
            self._mark_stale()
            self._schedule_autosave()
        device_pos_combo.bind('<<ComboboxSelected>>', on_device_position_change)
        
        # Driveline Angle
        ttk.Label(form_frame, text="Driveline Angle (°):").grid(row=2, column=0, sticky='w', pady=5)
        angle_var = tk.StringVar(value=str(group.get('driveline_angle', 0.0)))
        angle_spinbox = ttk.Spinbox(form_frame, from_=0, to=45, increment=0.5,
                                     textvariable=angle_var, width=8)
        angle_spinbox.grid(row=2, column=1, sticky='w', pady=5, padx=(10, 0))
        
        def on_angle_change(*args):
            try:
                val = float(angle_var.get())
                val = max(0.0, min(45.0, val))
                group['driveline_angle'] = val
            except ValueError:
                pass
            self._mark_stale()
            self._schedule_autosave()
        angle_var.trace_add('write', on_angle_change)
        
        # Horizontal container for segments (left) and template info (right)
        content_row = ttk.Frame(scrollable_frame)
        content_row.pack(fill='both', expand=True, pady=(10, 0), padx=10)
        
        # Segments section (left side)
        seg_frame = ttk.LabelFrame(content_row, text="Segments (left to right)", padding="10")
        seg_frame.pack(side='left', fill='both', expand=True)
        
        # Summary card — template info (right side)
        summary_info = self.get_group_summary_info(group)
        if summary_info:
            summary_frame = ttk.LabelFrame(content_row, text="Template Info", padding="10")
            summary_frame.pack(side='left', fill='y', padx=(10, 0))
            
            for label, value in summary_info:
                if label == "" and value == "":
                    ttk.Separator(summary_frame, orient='horizontal').pack(fill='x', pady=5)
                    continue
                info_row = ttk.Frame(summary_frame)
                info_row.pack(fill='x', pady=1)
                ttk.Label(info_row, text=label, font=('Helvetica', 9, 'bold'), width=18, anchor='e').pack(side='left')
                ttk.Label(info_row, text=value, font=('Helvetica', 9), foreground='#333333').pack(side='left', padx=(8, 0))
        
        # String count display
        self.group_string_count_label = ttk.Label(seg_frame, text="", font=('Helvetica', 10, 'bold'))
        self.group_string_count_label.pack(anchor='w', pady=(0, 10))
        
        # Headers
        header_frame = ttk.Frame(seg_frame)
        header_frame.pack(fill='x')
        ttk.Label(header_frame, text="Qty", width=6).pack(side='left', padx=2)
        ttk.Label(header_frame, text="Tracker Template", width=30).pack(side='left', padx=2)
        ttk.Label(header_frame, text="Harness Config", width=14).pack(side='left', padx=2)
        ttk.Label(header_frame, text="", width=9).pack(side='left', padx=2)
        
        # Container for segment rows
        self.segment_rows_container = ttk.Frame(seg_frame)
        self.segment_rows_container.pack(fill='x', pady=(5, 0))
        
        # Add existing segment rows
        for i, seg in enumerate(group['segments']):
            self._add_segment_ui(group, group_idx, i, seg)
        
        # Update count
        self._update_group_string_count(group)
        
        # Add segment button
        add_btn = ttk.Button(seg_frame, text="+ Add Segment",
                            command=lambda: self._add_segment_to_group(group, group_idx))
        add_btn.pack(anchor='w', pady=(10, 0))
    
    def _add_segment_ui(self, group: dict, group_idx: int, seg_idx: int, segment: dict):
        """Add a segment configuration row to the UI"""
        row_frame = ttk.Frame(self.segment_rows_container)
        row_frame.pack(fill='x', pady=2)
        
        # Quantity
        qty_var = tk.StringVar(value=str(segment['quantity']))
        qty_spinbox = ttk.Spinbox(row_frame, from_=1, to=10000, textvariable=qty_var, width=6)
        qty_spinbox.pack(side='left', padx=2)
        
        # Template dropdown
        compatible_keys = self.get_compatible_templates(group)
        template_display_map = {}  # display_name -> template_key
        display_names = []
        
        # Add "Unlinked" option for backward compat
        unlinked_label = f"Unlinked ({segment.get('strings_per_tracker', 3)}S)"
        display_names.append(unlinked_label)
        template_display_map[unlinked_label] = None
        
        for key in compatible_keys:
            display = self.get_template_display_name(key)
            display_names.append(display)
            template_display_map[display] = key
        
        # Determine current selection
        current_ref = segment.get('template_ref')
        if current_ref and current_ref in self.enabled_templates:
            current_display = self.get_template_display_name(current_ref)
        else:
            current_display = unlinked_label
        
        template_var = tk.StringVar(value=current_display)
        template_combo = ttk.Combobox(row_frame, textvariable=template_var,
                                      values=display_names, width=28, state='readonly')
        template_combo.pack(side='left', padx=2)
        self.disable_combobox_scroll(template_combo)
        
        # Harness config (still user-specified)
        harness_var = tk.StringVar(value=segment['harness_config'])
        harness_options = self.get_harness_options(segment.get('strings_per_tracker', 3))
        harness_combo = ttk.Combobox(row_frame, textvariable=harness_var,
                                     values=harness_options, width=12)
        harness_combo.pack(side='left', padx=2)
        self.disable_combobox_scroll(harness_combo)
        
        # Track combo for LV collection method disabling
        self._harness_combos.append((harness_combo, harness_var, segment))
        if self.lv_collection_var.get() == 'String HR':
            spt = segment.get('strings_per_tracker', 1)
            override_display = '+'.join(['1'] * int(spt))
            original_config = segment['harness_config']
            harness_var.set(override_display)
            segment['harness_config'] = original_config  # restore — don't overwrite saved config
            harness_combo.config(state='disabled')
        elif self.lv_collection_var.get() == 'Trunk Bus':
            harness_combo.config(state='disabled')
        
        # Move up button
        up_btn = ttk.Button(row_frame, text="▲", width=2,
                           command=lambda si=seg_idx: self._move_segment(group, group_idx, si, -1))
        up_btn.pack(side='left', padx=(2, 0))
        if seg_idx == 0:
            up_btn.config(state='disabled')
        
        # Move down button
        down_btn = ttk.Button(row_frame, text="▼", width=2,
                             command=lambda si=seg_idx: self._move_segment(group, group_idx, si, 1))
        down_btn.pack(side='left', padx=(0, 2))
        if seg_idx == len(group['segments']) - 1:
            down_btn.config(state='disabled')
        
        # Delete button
        del_btn = ttk.Button(row_frame, text="×", width=3,
                            command=lambda: self._delete_segment(group, group_idx, seg_idx))
        del_btn.pack(side='left', padx=2)
        
        # Callbacks
        def on_template_change(*args):
            selected_display = template_var.get()
            selected_key = template_display_map.get(selected_display)
            segment['template_ref'] = selected_key
            
            if selected_key and selected_key in self.enabled_templates:
                # Derive strings_per_tracker from template
                tdata = self.enabled_templates[selected_key]
                new_spt = tdata.get('strings_per_tracker', 3)
                segment['strings_per_tracker'] = new_spt
                
                # Update harness options for new string count
                new_options = self.get_harness_options(new_spt)
                harness_combo['values'] = new_options
                if harness_var.get() not in new_options:
                    harness_var.set(new_options[0])
                    segment['harness_config'] = new_options[0]
            
            self._update_group_string_count(group)
            self._auto_unlock_allocation()
            self._mark_stale()
            self._schedule_autosave()
            
            # Update derived module from templates
            self._derive_module_from_templates()
            
            # Refresh wire sizing for new template/harness configs
            self._refresh_wire_sizing_for_segments()
            
            # Rebuild details panel to refresh summary card
            for widget in self.details_container.winfo_children():
                widget.destroy()
            self.show_group_details(group_idx)
        template_combo.bind('<<ComboboxSelected>>', on_template_change)
        
        def on_qty_change(*args):
            try:
                segment['quantity'] = max(1, int(qty_var.get()))
            except ValueError:
                pass
            self._update_group_string_count(group)
            self._auto_unlock_allocation()
            self._mark_stale()
            self._schedule_autosave()
        qty_var.trace_add('write', on_qty_change)
        
        def on_harness_change(*args):
            segment['harness_config'] = harness_var.get()
            self._refresh_wire_sizing_for_segments()
            self._mark_stale()
            self._schedule_autosave()
        harness_var.trace_add('write', on_harness_change)
    
    def _add_segment_to_group(self, group: dict, group_idx: int):
        """Add a new segment to the group and refresh UI"""
        # Default to first compatible template if available
        compatible = self.get_compatible_templates(group)
        if compatible:
            default_ref = compatible[0]
            default_spt = self.enabled_templates[default_ref].get('strings_per_tracker', 3)
        else:
            default_ref = None
            default_spt = 3
        
        group['segments'].append({
            'quantity': 1,
            'strings_per_tracker': default_spt,
            'harness_config': str(default_spt),
            'template_ref': default_ref
        })
        # Rebuild the details panel to show the new segment
        for widget in self.details_container.winfo_children():
            widget.destroy()
        self.show_group_details(group_idx)
        self._refresh_wire_sizing_for_segments()
        self._auto_unlock_allocation()
        self._mark_stale()
        self._schedule_autosave()
    
    def _move_segment(self, group: dict, group_idx: int, seg_idx: int, direction: int):
        """Move a segment up (-1) or down (+1) within the group"""
        new_idx = seg_idx + direction
        if new_idx < 0 or new_idx >= len(group['segments']):
            return
        group['segments'][seg_idx], group['segments'][new_idx] = group['segments'][new_idx], group['segments'][seg_idx]
        # Rebuild segment editor to fix indices
        self._harness_combos = []
        for widget in self.details_container.winfo_children():
            widget.destroy()
        self.show_group_details(group_idx)
        self._refresh_wire_sizing_for_segments()
        self._auto_unlock_allocation()
        self._mark_stale()
        self._schedule_autosave()

    def _delete_segment(self, group: dict, group_idx: int, seg_idx: int):
        """Delete a segment from the group"""
        if len(group['segments']) <= 1:
            return  # Keep at least one segment
        del group['segments'][seg_idx]
        # Rebuild segment editor to fix indices
        for widget in self.details_container.winfo_children():
            widget.destroy()
        self.show_group_details(group_idx)
        self._auto_unlock_allocation()
        self._mark_stale()
        self._schedule_autosave()
    
    def _update_group_string_count(self, group: dict):
        """Update the string/tracker count label for a group"""
        if not hasattr(self, 'group_string_count_label'):
            return
        total_trackers = sum(seg['quantity'] for seg in group['segments'])
        total_strings = sum(seg['quantity'] * seg['strings_per_tracker'] for seg in group['segments'])
        self.group_string_count_label.config(
            text=f"Total Strings: {total_strings:,}  |  Total Trackers: {total_trackers:,}")

    # ==================== Calculation Methods ====================

    def calculate_estimate(self):
        """Calculate and display the rolled-up BOM estimate"""
        # Clear previous results
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        self.checked_items.clear()
        
        # Sync all group listbox text before calculating
        self._refresh_group_listbox(preserve_selection=True)
        
        # Aggregated totals
        totals = {
            'combiners_by_breaker': {},  # {breaker_size: quantity}
            'combiner_details': [],  # list of matched CB info dicts
            'string_inverters': 0,
            'trackers_by_string': {},  # {num_strings: quantity}
            'harnesses_by_size': {},   # {num_strings: quantity}
            'whips_by_length': {},     # {length_ft: quantity}
            'extenders_pos_by_length': {},  # {length_ft: quantity}
            'extenders_neg_by_length': {},  # {length_ft: quantity}
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
        
        # Validate — need at least one linked template or a legacy fallback module
        self._derive_module_from_templates()
        # Check for Trunk Bus — not yet implemented
        lv_method = self.lv_collection_var.get() if hasattr(self, 'lv_collection_var') else 'Wire Harness'
        if lv_method == 'Trunk Bus':
            from tkinter import messagebox
            messagebox.showinfo("Not Yet Implemented", "Trunk Bus collection method is not yet implemented.")
            return
        if not self.selected_module:
            from tkinter import messagebox
            messagebox.showwarning(
                "No Module Available",
                "No tracker templates are linked and no legacy module data was found.\n\n"
                "Please link a tracker template to at least one segment before calculating."
            )
            return
        
        # ==================== Build tracker sequence from groups ====================
        try:
            modules_per_string = int(self.modules_per_string_var.get())
        except ValueError:
            modules_per_string = 28

        tracker_sequence = []
        total_all_trackers = 0
        total_all_strings = 0
        total_all_harnesses = 0
        max_harness_strings = 0
        self._tracker_to_segment = []
        
        # Track unique modules across all groups
        unique_modules = {}  # "Manufacturer Model (WattageW)" -> module_spec_dict
        
        # Per-segment module data for geometry calculations
        segment_module_data = []  # list of {module_spec_dict, modules_per_string, qty, spt}

        for group in self.groups:
            for seg in group['segments']:
                qty = seg['quantity']
                spt = seg['strings_per_tracker']
                harness_config = self._get_effective_harness_config(seg)

                if qty <= 0:
                    continue
                
                # Resolve module from template
                seg_module = None
                seg_mps = modules_per_string
                ref = seg.get('template_ref')
                if ref and ref in self.enabled_templates:
                    tdata = self.enabled_templates[ref]
                    seg_module = tdata.get('module_spec', {})
                    seg_mps = tdata.get('modules_per_string', modules_per_string)
                    
                    # Track unique modules
                    mod_key = f"{seg_module.get('manufacturer', '?')} {seg_module.get('model', '?')} ({seg_module.get('wattage', '?')}W)"
                    if mod_key not in unique_modules:
                        unique_modules[mod_key] = seg_module
                
                segment_module_data.append({
                    'module_spec': seg_module,
                    'modules_per_string': seg_mps,
                    'qty': qty,
                    'spt': spt
                })

                # Determine if this segment has partial strings
                full_spt = int(spt)
                has_partial = (spt != full_spt)
                
                # Add to flat tracker sequence (one entry per physical tracker)
                # For partial-string trackers, pair adjacent trackers:
                #   Left of pair: full_spt + 1 (owns the shared string)
                #   Right of pair: full_spt
                #   Unpaired odd tracker: full_spt (half-string not counted)
                num_pairs = qty // 2 if has_partial else 0
                unpaired = qty % 2 if has_partial else 0
                
                for i in range(qty):
                    if has_partial:
                        if i % 2 == 0 and i + 1 < qty:
                            effective_spt = full_spt + 1  # Left of pair — owns shared string
                        elif i % 2 == 1:
                            effective_spt = full_spt  # Right of pair
                        else:
                            effective_spt = full_spt  # Unpaired odd tracker
                    else:
                        effective_spt = spt
                    
                    tracker_sequence.append(effective_spt)
                    self._tracker_to_segment.append({
                        'group_idx': self.groups.index(group),
                        'seg': seg,
                        'device_position': group.get('device_position', 'middle'),
                    })

                total_all_trackers += qty
                if has_partial:
                    total_all_strings += full_spt * qty + num_pairs  # Only paired halves count
                else:
                    total_all_strings += qty * spt

                # Count trackers by effective string count
                if has_partial:
                    num_pairs_count = qty // 2
                    # Left of pair
                    left_spt = full_spt + 1
                    if left_spt not in totals['trackers_by_string']:
                        totals['trackers_by_string'][left_spt] = 0
                    totals['trackers_by_string'][left_spt] += num_pairs_count
                    # Right of pair + any unpaired
                    right_count = num_pairs_count + (qty % 2)
                    if full_spt not in totals['trackers_by_string']:
                        totals['trackers_by_string'][full_spt] = 0
                    totals['trackers_by_string'][full_spt] += right_count
                else:
                    if spt not in totals['trackers_by_string']:
                        totals['trackers_by_string'][spt] = 0
                    totals['trackers_by_string'][spt] += qty

                # Count harnesses by size
                harness_sizes = self.parse_harness_config(harness_config)
                for size in harness_sizes:
                    if size > max_harness_strings:
                        max_harness_strings = size
                    if size not in totals['harnesses_by_size']:
                        totals['harnesses_by_size'][size] = 0
                    totals['harnesses_by_size'][size] += qty
                    total_all_harnesses += qty
        
        # Build harness-count-per-spt lookup for whip calculation
        harness_count_by_spt = {}
        harness_sizes_by_spt = {}
        for group in self.groups:
            for seg in group['segments']:
                spt = seg['strings_per_tracker']
                if spt not in harness_count_by_spt:
                    harness_sizes = self.parse_harness_config(self._get_effective_harness_config(seg))
                    harness_count_by_spt[spt] = len(harness_sizes)
                    harness_sizes_by_spt[spt] = harness_sizes

        # Warn about unpaired partial strings
        unpaired_warnings = []
        for group in self.groups:
            for seg in group['segments']:
                seg_spt = seg['strings_per_tracker']
                if seg_spt != int(seg_spt) and seg['quantity'] % 2 != 0:
                    ref = seg.get('template_ref', 'Unlinked')
                    unpaired_warnings.append(f"{group['name']}: {seg['quantity']}x {seg_spt}S has 1 unpaired half-string")
        
        if unpaired_warnings:
            from tkinter import messagebox
            messagebox.showwarning(
                "Unpaired Partial Strings",
                "The following segments have an odd number of partial-string trackers, "
                "leaving half-strings unpaired:\n\n" + "\n".join(unpaired_warnings)
            )

        # ==================== Module geometry (primary module for global calcs) ====================
        module_isc = self.selected_module.isc
        module_width_mm = self.selected_module.width_mm
        module_width_ft = module_width_mm / 304.8
        string_length_ft = module_width_ft * modules_per_string

        # ==================== Build spatial tracker entries ====================
        try:
            row_spacing_ft = float(self.row_spacing_var.get())
        except (ValueError, AttributeError):
            row_spacing_ft = 20.0

        fallback_width_ft = 6.0
        fallback_length_ft = 180.0

        # Pre-compute tracker count per group for auto-layout fallback
        group_tracker_counts = []
        for group in self.groups:
            count = sum(seg['quantity'] for seg in group['segments'] if seg['quantity'] > 0)
            group_tracker_counts.append(count)

        tracker_entries = []
        flat_idx = 0
        auto_x_cursor = 0.0  # Running X for auto-layout

        for grp_idx, group in enumerate(self.groups):
            saved_x = group.get('position_x')
            saved_y = group.get('position_y')

            if saved_x is not None and saved_y is not None:
                group_x = saved_x
                group_y = saved_y
            else:
                # Auto-layout: stack groups left-to-right at y=0
                group_x = auto_x_cursor
                group_y = 0.0


            # Compute group reference motor Y (same as SitePreviewWindow uses)
            # This is the motor_y of the FIRST segment's template in the group
            group_ref_motor_y_ft = 0.0
            for seg in group['segments']:
                ref_check = seg.get('template_ref')
                if ref_check and ref_check in self.enabled_templates:
                    tdata_check = self.enabled_templates[ref_check]
                    if tdata_check.get('has_motor', True):
                        # Compute this template's motor_y using same logic
                        ms_c = tdata_check.get('module_spec', {})
                        orient_c = tdata_check.get('module_orientation', 'Portrait')
                        mps_c = tdata_check.get('modules_per_string', 28)
                        spacing_c = tdata_check.get('module_spacing_m', 0.02)
                        placement_c = tdata_check.get('motor_placement_type', 'between_strings')
                        pos_after_c = tdata_check.get('motor_position_after_string', None)
                        str_idx_c = tdata_check.get('motor_string_index', None)
                        split_n_c = tdata_check.get('motor_split_north', mps_c // 2)
                        mod_along_c = (ms_c.get('width_mm', 1000) if orient_c == 'Portrait' else ms_c.get('length_mm', 2000)) / 1000
                        
                        # Partial string on north adds offset
                        spt_c = tdata_check.get('strings_per_tracker', 1)
                        partial_north_m_c = 0
                        if spt_c != int(spt_c) and tdata_check.get('partial_string_side', 'north') == 'north':
                            partial_north_mods_c = round((spt_c - int(spt_c)) * mps_c)
                            partial_north_m_c = partial_north_mods_c * (mod_along_c + spacing_c)
                        
                        if placement_c == 'between_strings':
                            p = pos_after_c if pos_after_c is not None else (str_idx_c if str_idx_c is not None else 1)
                            mn = p * mps_c
                            motor_y_m = partial_north_m_c + (mn * mod_along_c + (mn - 1) * spacing_c + spacing_c) if mn > 0 else 0.0
                            group_ref_motor_y_ft = motor_y_m * 3.28084
                        elif placement_c == 'middle_of_string':
                            s = str_idx_c if str_idx_c is not None else 1
                            mb = (s - 1) * mps_c + split_n_c
                            motor_y_m = partial_north_m_c + (mb * mod_along_c + (mb - 1) * spacing_c + spacing_c)
                            group_ref_motor_y_ft = motor_y_m * 3.28084
                        break  # Use first template's motor as reference

            # Driveline angle for this group
            driveline_angle_deg = group.get('driveline_angle', 0.0)
            driveline_tan = math.tan(math.radians(driveline_angle_deg)) if driveline_angle_deg > 0 else 0.0

            tracker_within_group = 0
            for seg in group['segments']:
                qty = seg['quantity']
                spt = seg['strings_per_tracker']
                if qty <= 0:
                    continue

                ref = seg.get('template_ref')
                dims = self._get_estimate_tracker_dims_ft(ref)
                t_length = dims[1] if dims else fallback_length_ft
                
                # Partial string pairing (same logic as tracker_sequence)
                full_spt = int(spt)
                has_partial = (spt != full_spt)

                for i in range(qty):
                    if has_partial:
                        if i % 2 == 0 and i + 1 < qty:
                            effective_spt = full_spt + 1
                        else:
                            effective_spt = full_spt
                    else:
                        effective_spt = int(spt)
                    
                    local_x_offset = tracker_within_group * row_spacing_ft
                    tracker_entries.append({
                        'original_idx': flat_idx,
                        'spt': effective_spt,
                        'x': group_x + local_x_offset,
                        'y': group_y + local_x_offset * driveline_tan,
                        'length_ft': t_length,
                        'motor_y_ft': group_ref_motor_y_ft,
                    })
                    flat_idx += 1
                    tracker_within_group += 1

            # Advance auto-layout cursor for next group
            group_width = group_tracker_counts[grp_idx] * row_spacing_ft
            auto_x_cursor += group_width + row_spacing_ft * 2  # Extra gap between groups

        # ==================== Allocation ====================
        allocation_result = None

        if self.selected_inverter and strings_per_inv > 0 and total_all_strings > 0:
            if self.allocation_locked and self.locked_allocation_result is not None:
                # Use the locked (frozen) allocation — skip spatial recalculation
                allocation_result = self.locked_allocation_result
                spatial_runs = allocation_result.get('spatial_runs', 1)
            elif tracker_entries:
                allocation_result = allocate_strings_spatial(
                    tracker_entries, strings_per_inv, row_spacing_ft
                )
                spatial_runs = allocation_result.get('spatial_runs', 1)
            else:
                allocation_result = allocate_strings_sequential(tracker_sequence, strings_per_inv)

            module_wattage = self.selected_module.wattage
            # Use site-level DC:AC (total DC power / total AC capacity)
            # Not per-inverter nominal, since allocation doesn't give every inverter the same count
            total_alloc_strings = allocation_result['summary']['total_strings']
            total_alloc_invs = allocation_result['summary']['total_inverters']
            if total_alloc_invs > 0:
                total_dc_kw = (total_alloc_strings * modules_per_string * module_wattage) / 1000
                total_ac_kw = total_alloc_invs * self.selected_inverter.rated_power_kw
                actual_dc_ac = round(total_dc_kw / total_ac_kw, 3)
            else:
                actual_dc_ac = 0.0            

            totals['string_inverters'] = allocation_result['summary']['total_inverters']
            totals['inverter_summary'] = {
                'strings_per_inverter': strings_per_inv,
                'total_inverters': allocation_result['summary']['total_inverters'],
                'total_split_trackers': allocation_result['summary']['total_split_trackers'],
                'total_strings': allocation_result['summary']['total_strings'],
                'actual_dc_ac': actual_dc_ac,
                'allocation_result': allocation_result,
                'allocations': [],  # Backward compat — will be replaced in next update
            }

        # ==================== Topology-driven device & combiner counting ===================

        total_inverters_count = totals.get('inverter_summary', {}).get('total_inverters', 0)
        num_devices = 0
        num_combiners = 0
        strings_per_cb = 0

        # Global breaker size (will add UI field later, default 400A for now)
        try:
            breaker_size = int(self.breaker_size_var.get())
        except (ValueError, AttributeError):
            breaker_size = 400

        if topology == 'Distributed String':
            num_devices = total_inverters_count

        elif topology == 'Centralized String':
            num_devices = total_inverters_count
            num_combiners = total_inverters_count
            strings_per_cb = strings_per_inv

        elif topology == 'Central Inverter':
            # Find largest matching CB to minimize combiner count
            fuse_current = module_isc * nec_factor * max(max_harness_strings, 1)
            fuse_holder_rating = self.get_fuse_holder_category(fuse_current)
            combiner_library = self.load_combiner_library()

            matching_cbs = [
                cb_data for cb_data in combiner_library.values()
                if (cb_data.get('breaker_size', 0) == breaker_size and
                    cb_data.get('fuse_holder_rating', '') == fuse_holder_rating)
            ]

            if matching_cbs:
                matching_cbs.sort(key=lambda c: c.get('max_inputs', 0), reverse=True)
                max_inputs = matching_cbs[0].get('max_inputs', 24)
            else:
                max_inputs = 24  # fallback

            if total_all_strings > 0:
                num_combiners = math.ceil(total_all_strings / max_inputs)
                strings_per_cb = math.ceil(total_all_strings / num_combiners)
            num_devices = num_combiners

        # Preliminary combiner count (used for DC feeder/AC homerun distance calc)
        # Actual CB part matching is done later via Device Configurator or fallback
        if num_combiners > 0:
            if breaker_size not in totals['combiners_by_breaker']:
                totals['combiners_by_breaker'][breaker_size] = 0
            totals['combiners_by_breaker'][breaker_size] += num_combiners

        # ==================== Harness split adjustment ====================
        # NOTE: _adjust_harnesses_for_splits will be updated to use harness_map
        # in the next batch. For now it's a no-op since allocations=[] above.
        self._adjust_harnesses_for_splits(totals)

        # ==================== Whip calculation ====================
        try:
            row_spacing = float(self.row_spacing_var.get())
        except ValueError:
            row_spacing = 20.0

        if total_all_trackers > 0 and num_devices > 0:
            whip_distances = self.calculate_whip_distances_from_positions(
                allocation_result, topology, num_devices, row_spacing
            )

            split_details = getattr(self, '_split_tracker_details', {})
            seen_whip_trackers = set()
            
            for entry in whip_distances:
                if len(entry) == 4:
                    distance_ft, spt, tidx, inv_idx = entry
                else:
                    distance_ft, spt = entry[0], entry[1]
                    tidx, inv_idx = -1, -1
                whip_length = self.round_whip_length(distance_ft)
                
                if tidx in split_details:
                    # Split tracker — find this portion's harness count
                    portion_harnesses = 0
                    for portion in split_details[tidx]['portions']:
                        if portion['inv_idx'] == inv_idx:
                            portion_harnesses = len(portion['harnesses'])
                            break
                    
                    if portion_harnesses == 0:
                        continue
                    
                    num_harnesses = portion_harnesses
                else:
                    # Non-split tracker — skip duplicates, use original harness count
                    if tidx in seen_whip_trackers:
                        continue
                    seen_whip_trackers.add(tidx)
                    num_harnesses = harness_count_by_spt.get(spt, 1)
                
                # Determine individual harness sizes for wire gauge lookup
                if tidx in split_details:
                    ind_harness_sizes = split_details[tidx]['portions'][0]['harnesses']
                    for portion in split_details[tidx]['portions']:
                        if portion['inv_idx'] == inv_idx:
                            ind_harness_sizes = portion['harnesses']
                            break
                else:
                    ind_harness_sizes = harness_sizes_by_spt.get(spt, [spt])
                
                for h_str_count in ind_harness_sizes:
                    gauge = self.get_wire_size_for('whip', h_str_count)
                    key = (whip_length, gauge)
                    if key not in totals['whips_by_length']:
                        totals['whips_by_length'][key] = 0
                    totals['whips_by_length'][key] += 2  # pos + neg
                    totals['total_whip_length'] += whip_length * 2

        # ==================== Extenders ====================
        split_details = getattr(self, '_split_tracker_details', {})
        tracker_seg_map = getattr(self, '_tracker_to_segment', [])
        
        # Count how many split trackers exist per (group_idx, seg identity) so we can reduce bulk qty
        split_tracker_seg_counts = {}  # (group_idx, id(seg)) -> count of split trackers
        for tidx in split_details:
            if tidx < len(tracker_seg_map):
                info = tracker_seg_map[tidx]
                key = (info['group_idx'], id(info['seg']))
                split_tracker_seg_counts[key] = split_tracker_seg_counts.get(key, 0) + 1
                
        # Process non-split trackers in bulk (original logic minus split count)
        for group_idx, group in enumerate(self.groups):
            device_position = group.get('device_position', 'middle')
            for seg in group['segments']:
                if seg['quantity'] <= 0:
                    continue
                
                key = (group_idx, id(seg))
                num_splits_in_seg = split_tracker_seg_counts.get(key, 0)
                non_split_qty = seg['quantity'] - num_splits_in_seg
                
                if non_split_qty > 0:
                    extender_pairs = self.calculate_extender_lengths_per_segment(seg, device_position)
                    harness_sizes = self.parse_harness_config(self._get_effective_harness_config(seg))
                    for pair_idx, (pos_len, neg_len) in enumerate(extender_pairs):
                        h_str_count = harness_sizes[pair_idx] if pair_idx < len(harness_sizes) else 1
                        gauge = self.get_wire_size_for('extender', h_str_count)
                        pos_rounded = self.round_whip_length(pos_len)
                        neg_rounded = self.round_whip_length(neg_len)
                        pos_key = (pos_rounded, gauge)
                        neg_key = (neg_rounded, gauge)
                        if pos_key not in totals['extenders_pos_by_length']:
                            totals['extenders_pos_by_length'][pos_key] = 0
                        totals['extenders_pos_by_length'][pos_key] += non_split_qty
                        if neg_key not in totals['extenders_neg_by_length']:
                            totals['extenders_neg_by_length'][neg_key] = 0
                        totals['extenders_neg_by_length'][neg_key] += non_split_qty
                        
        # Process split trackers individually — each portion gets its own extenders
        for tidx, details in split_details.items():
            if tidx >= len(tracker_seg_map):
                continue
            
            seg_info = tracker_seg_map[tidx]
            seg = seg_info['seg']
            device_position = seg_info['device_position']
            
            for portion in details['portions']:
                # Build a temporary segment matching this portion's harness config
                portion_config = '+'.join(str(h) for h in portion['harnesses'])
                temp_seg = dict(seg)
                temp_seg['harness_config'] = portion_config
                temp_seg['quantity'] = 1
                
                extender_pairs = self.calculate_extender_lengths_per_segment(temp_seg, device_position)
                portion_harness_sizes = portion['harnesses']
                
                for pair_idx, (pos_len, neg_len) in enumerate(extender_pairs):
                    h_str_count = portion_harness_sizes[pair_idx] if pair_idx < len(portion_harness_sizes) else 1
                    gauge = self.get_wire_size_for('extender', h_str_count)
                    pos_rounded = self.round_whip_length(pos_len)
                    neg_rounded = self.round_whip_length(neg_len)
                    pos_key = (pos_rounded, gauge)
                    neg_key = (neg_rounded, gauge)
                    if pos_key not in totals['extenders_pos_by_length']:
                        totals['extenders_pos_by_length'][pos_key] = 0
                    totals['extenders_pos_by_length'][pos_key] += 1
                    if neg_key not in totals['extenders_neg_by_length']:
                        totals['extenders_neg_by_length'][neg_key] = 0
                    totals['extenders_neg_by_length'][neg_key] += 1

        # ==================== DC Feeder and AC Homerun ====================
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
        
        use_routed = self.use_routed_var.get() if hasattr(self, 'use_routed_var') else False

        if use_routed and self.pads and allocation_result:
            # Use routed Manhattan distances from devices to pads
            try:
                routed = self.calculate_routed_feeder_distances(
                    allocation_result, topology, row_spacing
                )
            except Exception as e:
                print(f"[Routed distance error] {e}")
                import traceback
                traceback.print_exc()
                routed = {'feeder_distances': [], 'feeder_total_ft': 0, 'feeder_count': 0}
            
            if topology == 'Distributed String':
                # Devices are SIs, cables are AC homeruns
                # Devices are SIs, cables are AC homeruns
                totals['dc_feeder_count'] = 0
                totals['dc_feeder_total_ft'] = 0
                totals['ac_homerun_count'] = routed['feeder_count']
                totals['ac_homerun_total_ft'] = routed['feeder_total_ft']
            elif topology == 'Centralized String':
                # CBs route DC feeders to pads, inverters at pads have short AC
                totals['dc_feeder_count'] = routed['feeder_count']
                totals['dc_feeder_total_ft'] = routed['feeder_total_ft']
                totals['ac_homerun_count'] = total_inverters
                totals['ac_homerun_total_ft'] = total_inverters * ac_homerun_avg_ft
            elif topology == 'Central Inverter':
                totals['dc_feeder_count'] = routed['feeder_count']
                totals['dc_feeder_total_ft'] = routed['feeder_total_ft']
                totals['ac_homerun_count'] = 1
                totals['ac_homerun_total_ft'] = ac_homerun_avg_ft
            
            # Store routed details for potential export
            totals['routed_feeder_details'] = routed['feeder_distances']
        else:
            # Use average distance inputs
            if topology == 'Distributed String':
                totals['dc_feeder_count'] = 0
                totals['dc_feeder_total_ft'] = 0
                totals['ac_homerun_count'] = total_inverters
                totals['ac_homerun_total_ft'] = total_inverters * ac_homerun_avg_ft
            elif topology == 'Centralized String':
                totals['dc_feeder_count'] = total_combiners
                totals['dc_feeder_total_ft'] = total_combiners * dc_feeder_avg_ft
                totals['ac_homerun_count'] = total_inverters
                totals['ac_homerun_total_ft'] = total_inverters * ac_homerun_avg_ft
            elif topology == 'Central Inverter':
                totals['dc_feeder_count'] = total_combiners
                totals['dc_feeder_total_ft'] = total_combiners * dc_feeder_avg_ft
                totals['ac_homerun_count'] = 1
                totals['ac_homerun_total_ft'] = ac_homerun_avg_ft

        # ==================== Display Results ====================
        
        total_inverters = totals.get('inverter_summary', {}).get('total_inverters', 0)
        total_combiners = sum(totals['combiners_by_breaker'].values())

        # Stash display context into totals so _redraw_results_tree has everything
        totals['_display'] = {
            'topology': topology,
            'use_routed': use_routed if 'use_routed' in dir() else False,
            'dc_feeder_avg_ft': dc_feeder_avg_ft if 'dc_feeder_avg_ft' in dir() else 0,
            'ac_homerun_avg_ft': ac_homerun_avg_ft if 'ac_homerun_avg_ft' in dir() else 0,
        }

        # Store totals for Excel export
        self.last_totals = totals
        self._results_stale = False
        
        # Build structured combiner assignments if they don't exist yet
        if not (self.last_combiner_assignments and any(cb.get('connections') for cb in self.last_combiner_assignments)):
            self._build_combiner_assignments(totals, topology)
        
        # Read combiner BOM from Device Configurator (single source of truth)
        # Falls back to simple assignment-based totals if DC isn't available
        if not self._read_combiner_bom_from_device_config():
            self._rebuild_combiner_totals_from_assignments()
            
            self.last_totals = totals
            totals['combiners_by_breaker'] = {}
            totals['combiner_details'] = []
            
            # Group by breaker size
            for cb in self.last_combiner_assignments:
                bs = cb['breaker_size']
                if bs not in totals['combiners_by_breaker']:
                    totals['combiners_by_breaker'][bs] = 0
                totals['combiners_by_breaker'][bs] += 1
            
            # Rebuild combiner details with per-breaker matching
            module_isc = self.selected_module.isc if self.selected_module else 0
            for bs, qty in totals['combiners_by_breaker'].items():
                # Find max harness strings for fuse holder rating
                max_h_strings = 1
                max_inputs = 0
                for cb in self.last_combiner_assignments:
                    if cb['breaker_size'] == bs:
                        num_inputs = len(cb['connections'])
                        max_inputs = max(max_inputs, num_inputs)
                        for conn in cb['connections']:
                            max_h_strings = max(max_h_strings, conn['num_strings'])
                
                fuse_current = module_isc * 1.56 * max_h_strings
                fuse_holder_rating = self.get_fuse_holder_category(fuse_current)
                strings_per_cb = max_inputs
                
                matched_cb = self.find_combiner_box(strings_per_cb, bs, fuse_holder_rating)
                if matched_cb:
                    totals['combiner_details'].append({
                        'part_number': matched_cb.get('part_number', ''),
                        'description': matched_cb.get('description', ''),
                        'max_inputs': matched_cb.get('max_inputs', 0),
                        'breaker_size': bs,
                        'fuse_holder_rating': fuse_holder_rating,
                        'strings_per_cb': strings_per_cb,
                        'quantity': qty,
                        'block_name': 'Site Total'
                    })
                else:
                    totals['combiner_details'].append({
                        'part_number': 'NO MATCH',
                        'description': f'No CB found: {strings_per_cb} inputs, {bs}A, {fuse_holder_rating}',
                        'max_inputs': strings_per_cb,
                        'breaker_size': bs,
                        'fuse_holder_rating': fuse_holder_rating,
                        'strings_per_cb': strings_per_cb,
                        'quantity': qty,
                        'block_name': 'Site Total'
                    })
            
            # Update last_totals since we modified totals
            self.last_totals = totals

        self._redraw_results_tree()

    def _write_combiner_sheet(self, wb):
        """Write a Combiner Boxes sheet to the workbook from Device Configurator data."""
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        
        # Get Device Configurator
        main_app = getattr(self, 'main_app', None)
        if not main_app or not hasattr(main_app, 'device_configurator'):
            return
        dc = main_app.device_configurator
        if not hasattr(dc, 'combiner_configs') or not dc.combiner_configs:
            return
        
        ws = wb.create_sheet("Combiner Boxes")
        
        # Styles
        title_font = Font(bold=True, size=14)
        header_font = Font(bold=True, size=11, color="FFFFFF")
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        cb_header_font = Font(bold=True, size=11)
        cb_header_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        center_align = Alignment(horizontal='center', vertical='center')
        mismatch_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        
        row = 1
        ws.merge_cells(f'A{row}:H{row}')
        ws.cell(row=row, column=1, value="Combiner Box Configuration Details").font = title_font
        row += 2
        
        # NEC factor info
        nec_factor = float(dc.nec_factor_var.get()) if hasattr(dc, 'nec_factor_var') else 1.56
        ws.cell(row=row, column=1, value="NEC Safety Factor:").font = Font(bold=True)
        ws.cell(row=row, column=2, value=nec_factor)
        row += 2
        
        # Column headers
        headers = ['Combiner Box', 'Tracker', 'Harness', 'Strings', 'Isc (A)',
                    'Harness Current (A)', 'Fuse Size (A)', 'Cable Size',
                    'Total Current (A)', 'Breaker Size (A)']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align
            cell.border = thin_border
        header_row = row
        row += 1
        
        # Data rows — one section per combiner box
        for combiner_id in sorted(dc.combiner_configs.keys()):
            config = dc.combiner_configs[combiner_id]
            
            for i, conn in enumerate(config.connections):
                cb_cell = ws.cell(row=row, column=1, value=combiner_id)
                cb_cell.border = thin_border
                cb_cell.alignment = center_align
                ws.cell(row=row, column=2, value=conn.tracker_id).border = thin_border
                ws.cell(row=row, column=2).alignment = center_align
                ws.cell(row=row, column=3, value=conn.harness_id).border = thin_border
                ws.cell(row=row, column=3).alignment = center_align
                ws.cell(row=row, column=4, value=conn.num_strings).border = thin_border
                ws.cell(row=row, column=4).alignment = center_align
                ws.cell(row=row, column=5, value=round(conn.module_isc, 2)).border = thin_border
                ws.cell(row=row, column=5).alignment = center_align
                ws.cell(row=row, column=6, value=round(conn.harness_current, 2)).border = thin_border
                ws.cell(row=row, column=6).alignment = center_align
                
                fuse_cell = ws.cell(row=row, column=7, value=conn.get_display_fuse_size())
                fuse_cell.border = thin_border
                fuse_cell.alignment = center_align
                
                cable_cell = ws.cell(row=row, column=8, value=conn.get_display_cable_size())
                cable_cell.border = thin_border
                cable_cell.alignment = center_align
                if conn.is_cable_size_mismatch():
                    cable_cell.fill = mismatch_fill
                
                # Total current and breaker on first connection row only
                if i == 0:
                    ws.cell(row=row, column=9, value=round(config.total_input_current, 2)).border = thin_border
                    ws.cell(row=row, column=9).alignment = center_align
                    ws.cell(row=row, column=10, value=config.get_display_breaker_size()).border = thin_border
                    ws.cell(row=row, column=10).alignment = center_align
                else:
                    ws.cell(row=row, column=9, value='').border = thin_border
                    ws.cell(row=row, column=10, value='').border = thin_border
                
                row += 1
            
            # Blank row between combiner boxes
            row += 1
        
        # Auto-fit columns
        for col_idx in range(1, len(headers) + 1):
            max_length = len(headers[col_idx - 1])  # Start with header width
            col_letter = get_column_letter(col_idx)
            for cell in ws[col_letter]:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max_length + 3, 30)

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
        
        # Build suggested filename (matching BOM Generator convention)
        def clean_filename(s):
            return "".join(c for c in s if c.isalnum() or c in (' ', '-', '_')).strip()
        
        client = "Unknown_Client"
        project_name = "Unknown_Project"
        estimate_name = "Estimate"
        
        if self.current_project and self.current_project.metadata:
            client = clean_filename(self.current_project.metadata.client or "Unknown_Client")
            project_name = clean_filename(self.current_project.metadata.name or "Unknown_Project")
        if self.estimate_id and self.current_project:
            est_data = self.current_project.quick_estimates.get(self.estimate_id, {})
            estimate_name = clean_filename(est_data.get('name', 'Estimate'))
        
        lv_method_tag = clean_filename(self.lv_collection_var.get() if hasattr(self, 'lv_collection_var') else 'Wire Harness')
        suggested_filename = f"{client}_{project_name}_Ampacity Quick eBOM_{estimate_name}_{lv_method_tag}.xlsx"
        
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
            info_ws = wb.active
            info_ws.title = "Project Info"
            ws = wb.create_sheet("Quick Estimate BOM")
            
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
            
            # ========== PROJECT INFO SHEET ==========
            info_row = 1
            
            info_ws.merge_cells(f'A{info_row}:E{info_row}')
            info_ws.cell(row=info_row, column=1, value="Quick Estimate — Project Info").font = title_font
            info_row += 2
            
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
            
            # List all unique modules used
            if hasattr(self, 'last_totals') and self.last_totals.get('unique_modules'):
                seg_mod_data = self.last_totals.get('segment_module_data', [])
                for i, (mod_key, mod_data) in enumerate(self.last_totals['unique_modules'].items()):
                    prefix = "Module:" if i == 0 else f"Module {i+1}:"
                    info_items.append((prefix, mod_key))
                    info_items.append(("  Isc:", f"{mod_data.get('isc', '?')} A"))
                    info_items.append(("  Width:", f"{mod_data.get('width_mm', '?')} mm"))
                    for smd in seg_mod_data:
                        if smd.get('module_spec') == mod_data:
                            info_items.append(("  Modules/String:", str(smd['modules_per_string'])))
                            break
            else:
                info_items.append(("Module:", f"{self.selected_module.manufacturer} {self.selected_module.model} ({self.selected_module.wattage}W)"))
                info_items.append(("Module Isc:", f"{module_isc} A"))
                info_items.append(("Module Width:", f"{module_width_mm} mm"))
            info_items.append(("Row Spacing:", f"{row_spacing} ft"))
            info_items.append(("LV Collection Method:", self.lv_collection_var.get() if hasattr(self, 'lv_collection_var') else 'Wire Harness'))
            
            if self.selected_inverter:
                inv = self.selected_inverter
                info_items.append(("Inverter:", f"{inv.manufacturer} {inv.model} ({inv.rated_power_kw}kW AC)"))
                info_items.append(("Topology:", self.topology_var.get()))
                info_items.append(("DC:AC Ratio (target):", self.dc_ac_ratio_var.get()))
                if hasattr(self, 'last_totals') and self.last_totals.get('inverter_summary'):
                    inv_sum = self.last_totals['inverter_summary']
                    info_items.append(("DC:AC Ratio (actual):", f"{inv_sum.get('actual_dc_ac', 0):.2f}"))
                    info_items.append(("Strings per Inverter (target):", str(inv_sum.get('strings_per_inverter', ''))))
                    info_items.append(("Total Inverters:", str(inv_sum.get('total_inverters', ''))))
                    info_items.append(("Split Trackers:", str(inv_sum.get('total_split_trackers', ''))))
            
            if self.estimate_id and self.current_project:
                est_data = self.current_project.quick_estimates.get(self.estimate_id, {})
                info_items.append(("Estimate:", est_data.get('name', '')))
            
            for label, value in info_items:
                info_ws.cell(row=info_row, column=1, value=label).font = label_font
                info_ws.cell(row=info_row, column=2, value=value)
                info_row += 1
            
            # Add copper price from pricing manager
            try:
                from src.utils.pricing_lookup import PricingLookup
                _pricing = PricingLookup()
                copper_price = _pricing.get_current_copper_price()
                active_tier = _pricing.get_active_tier()
                info_ws.cell(row=info_row, column=1, value="Copper Price:").font = label_font
                info_ws.cell(row=info_row, column=2, value=f"${copper_price:.2f}/lb")
                info_row += 1
            except Exception:
                pass
            
            info_row += 1
            
            # ========== GROUP CONFIGURATION SUMMARY (info sheet) ==========
            info_ws.merge_cells(f'A{info_row}:E{info_row}')
            info_ws.cell(row=info_row, column=1, value="Group Configuration Summary").font = title_font
            info_row += 1
            
            row_headers = ['Group', 'Segment Configs', 'Total Strings', 'Total Trackers']
            for col, header in enumerate(row_headers, 1):
                cell = info_ws.cell(row=info_row, column=col, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = center_align
                cell.border = thin_border
            info_row += 1
            
            for r in self.groups:
                group_strings = sum(int(s['quantity'] * s['strings_per_tracker']) for s in r['segments'])
                group_trackers = sum(s['quantity'] for s in r['segments'])
                seg_summary = ", ".join(
                    f"{s['quantity']}x{s['strings_per_tracker']}S({s['harness_config']})"
                    for s in r['segments'] if s['quantity'] > 0
                )
                group_data = [r['name'], seg_summary, group_strings, group_trackers]
                for col, value in enumerate(group_data, 1):
                    cell = info_ws.cell(row=info_row, column=col, value=value)
                    cell.border = thin_border
                    cell.alignment = center_align
                info_row += 1
            
            info_row += 1
            
            # ========== TRACKER SUMMARY (info sheet) ==========
            if hasattr(self, 'last_totals') and self.last_totals.get('trackers_by_string'):
                info_ws.merge_cells(f'A{info_row}:E{info_row}')
                info_ws.cell(row=info_row, column=1, value="Tracker Summary").font = title_font
                info_row += 1
                
                tracker_headers = ['Tracker Type', 'Quantity', 'Unit']
                for col, header in enumerate(tracker_headers, 1):
                    cell = info_ws.cell(row=info_row, column=col, value=header)
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = center_align
                    cell.border = thin_border
                info_row += 1
                
                total_trackers = 0
                for strings in sorted(self.last_totals['trackers_by_string'].keys()):
                    qty = self.last_totals['trackers_by_string'][strings]
                    total_trackers += qty
                    info_ws.cell(row=info_row, column=1, value=f"{strings}-String Trackers").border = thin_border
                    cell_qty = info_ws.cell(row=info_row, column=2, value=qty)
                    cell_qty.border = thin_border
                    cell_qty.alignment = center_align
                    cell_unit = info_ws.cell(row=info_row, column=3, value='ea')
                    cell_unit.border = thin_border
                    cell_unit.alignment = center_align
                    info_row += 1
                
                total_label = info_ws.cell(row=info_row, column=1, value="Total Trackers")
                total_label.font = label_font
                total_label.border = thin_border
                total_qty = info_ws.cell(row=info_row, column=2, value=total_trackers)
                total_qty.font = label_font
                total_qty.border = thin_border
                total_qty.alignment = center_align
                total_unit = info_ws.cell(row=info_row, column=3, value='ea')
                total_unit.border = thin_border
                total_unit.alignment = center_align
                info_row += 2
            
            # ========== INVERTER ALLOCATION (info sheet) ==========
            if hasattr(self, 'last_totals') and self.last_totals.get('inverter_summary'):
                inv_sum = self.last_totals['inverter_summary']
                
                info_ws.merge_cells(f'A{info_row}:E{info_row}')
                info_ws.cell(row=info_row, column=1, value="Inverter Allocation Summary").font = title_font
                info_row += 1
                
                alloc_headers = ['Inverter', 'Strings', 'Trackers', 'Pattern']
                for col, header in enumerate(alloc_headers, 1):
                    cell = info_ws.cell(row=info_row, column=col, value=header)
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = center_align
                    cell.border = thin_border
                info_row += 1
                
                allocation_result = inv_sum.get('allocation_result')
                if allocation_result:
                    for inv_idx, inv in enumerate(allocation_result['inverters']):
                        pattern_str = '-'.join(str(s) for s in inv['pattern'])
                        
                        inv_row = [
                            f"Inverter {inv_idx + 1}",
                            inv['total_strings'],
                            len(inv['tracker_indices']),
                            f"[{pattern_str}]"
                        ]
                        for col, value in enumerate(inv_row, 1):
                            cell = info_ws.cell(row=info_row, column=col, value=value)
                            cell.border = thin_border
                            cell.alignment = center_align
                        info_row += 1
            
            # Auto-fit columns on info sheet
            for col_idx in range(1, 6):
                max_length = 0
                col_letter = get_column_letter(col_idx)
                for cell in info_ws[col_letter]:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                info_ws.column_dimensions[col_letter].width = min(max_length + 4, 55)
            
            # ========== PROJECT INFO HEADER (BOM sheet) ==========
            row = 1
            ws.merge_cells(f'A{row}:F{row}')
            ws.cell(row=row, column=1, value="Quick Estimate — Project Info").font = title_font
            row += 2
            
            # Build the same info_items list used on the Project Info sheet
            bom_info_items = []
            if self.current_project and self.current_project.metadata:
                meta = self.current_project.metadata
                if meta.name:
                    bom_info_items.append(("Project Name:", meta.name))
                if meta.client:
                    bom_info_items.append(("Customer:", meta.client))
                if meta.location:
                    bom_info_items.append(("Location:", meta.location))
            
            if hasattr(self, 'last_totals') and self.last_totals.get('unique_modules'):
                seg_mod_data = self.last_totals.get('segment_module_data', [])
                for i, (mod_key, mod_data) in enumerate(self.last_totals['unique_modules'].items()):
                    prefix = "Module:" if i == 0 else f"Module {i+1}:"
                    bom_info_items.append((prefix, mod_key))
                    bom_info_items.append(("  Isc:", f"{mod_data.get('isc', '?')} A"))
                    bom_info_items.append(("  Width:", f"{mod_data.get('width_mm', '?')} mm"))
                    for smd in seg_mod_data:
                        if smd.get('module_spec') == mod_data:
                            bom_info_items.append(("  Modules/String:", str(smd['modules_per_string'])))
                            break
            elif self.selected_module:
                bom_info_items.append(("Module:", f"{self.selected_module.manufacturer} {self.selected_module.model} ({self.selected_module.wattage}W)"))
                bom_info_items.append(("Module Isc:", f"{module_isc} A"))
                bom_info_items.append(("Module Width:", f"{module_width_mm} mm"))
            bom_info_items.append(("Row Spacing:", f"{row_spacing} ft"))
            
            if self.selected_inverter:
                inv = self.selected_inverter
                bom_info_items.append(("Inverter:", f"{inv.manufacturer} {inv.model} ({inv.rated_power_kw}kW AC)"))
                bom_info_items.append(("Topology:", self.topology_var.get()))
                bom_info_items.append(("DC:AC Ratio (target):", self.dc_ac_ratio_var.get()))
                if hasattr(self, 'last_totals') and self.last_totals.get('inverter_summary'):
                    inv_sum = self.last_totals['inverter_summary']
                    bom_info_items.append(("DC:AC Ratio (actual):", f"{inv_sum.get('actual_dc_ac', 0):.2f}"))
                    bom_info_items.append(("Strings per Inverter:", str(inv_sum.get('strings_per_inverter', ''))))
                    bom_info_items.append(("Total Inverters:", str(inv_sum.get('total_inverters', ''))))
                    bom_info_items.append(("Split Trackers:", str(inv_sum.get('total_split_trackers', ''))))
            
            if self.estimate_id and self.current_project:
                est_data = self.current_project.quick_estimates.get(self.estimate_id, {})
                bom_info_items.append(("Estimate:", est_data.get('name', '')))
            
            bom_info_items.append(("Date:", datetime.now().strftime('%Y-%m-%d')))
            
            # Add copper price from pricing manager
            try:
                from src.utils.pricing_lookup import PricingLookup
                _pricing = PricingLookup()
                copper_price = _pricing.get_current_copper_price()
                active_tier = _pricing.get_active_tier()
                bom_info_items.append(("Copper Price:", f"${copper_price:.2f}/lb"))
            except Exception:
                pass
            
            for label, value in bom_info_items:
                ws.cell(row=row, column=1, value=label).font = label_font
                ws.cell(row=row, column=2, value=value)
                row += 1
            
            row += 1
            
            # ========== BOM RESULTS SECTION (BOM sheet) ==========
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
                # Convert qty to number for Excel
                qty_val = ''
                if qty:
                    try:
                        qty_val = int(qty) if '.' not in str(qty) else float(qty)
                    except (ValueError, TypeError):
                        qty_val = qty
                cell_qty = ws.cell(row=row, column=3, value=qty_val)
                cell_unit = ws.cell(row=row, column=4, value=unit if unit else '')
                # Convert unit_cost to number for Excel
                unit_cost_val = ''
                if unit_cost:
                    try:
                        cleaned = str(unit_cost).replace('$', '').replace(',', '').strip()
                        unit_cost_val = float(cleaned) if cleaned else ''
                    except (ValueError, TypeError):
                        unit_cost_val = unit_cost
                cell_unit_cost = ws.cell(row=row, column=5, value=unit_cost_val)
                if isinstance(unit_cost_val, float):
                    cell_unit_cost.number_format = '"$"#,##0.00'

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

            # ========== COMBINER BOXES SHEET ==========
            self._write_combiner_sheet(wb)
            
            # ========== AUTO-FIT COLUMNS (BOM sheet) ==========
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
    
    def __init__(self, parent, inv_summary, topology, colors, groups, enabled_templates, row_spacing_ft,
                 num_devices=0, device_label='CB', initial_inspect=False, pads=None, device_names=None):
        super().__init__(parent)
        self.title("Site Preview — Inverter Allocation")
        self.geometry("1100x750")
        self.minsize(600, 400)
        
        self.inv_summary = inv_summary
        self.topology = topology
        self.colors = colors
        self.groups = groups or []
        self.enabled_templates = enabled_templates or {}
        self.row_spacing_ft = row_spacing_ft
        self.num_devices = num_devices
        self.device_label = device_label
        self.inspect_mode = initial_inspect
        self.selected_device_idx = None
        self.selected_pad_inspect_idx = None  # Pad selected in inspect mode
        self.pads = list(pads) if pads else []  # Deep copy so we don't mutate caller's list
        self.device_names = dict(device_names) if device_names else {}  # {device_idx: "custom_name"}
        self.selected_pad_idx = None
        self.placing_pad = False  # True when in "click to place" mode
        self.assigning_devices = False  # True when Assign Devices dialog is open
        
        # Zoom and pan state
        self.scale = 1.0
        self.pan_x = 0
        self.pan_y = 0
        self.dragging_canvas = False
        self.dragging_group = False
        self.drag_start_x = 0
        self.drag_start_y = 0
        self._drag_moved = False
        self._drag_group_start_x = 0
        self._drag_group_start_y = 0
        self.selected_group_idx = None
        self.align_on_motor = True
        
        # Read lock state from parent (QuickEstimate)
        self.allocation_locked = getattr(self.master, 'allocation_locked', False)
        
        self.setup_ui()
        self.build_layout_data()
        self._recolor_from_cb_assignments()
        self.after(50, self.fit_and_redraw)
    
    def _get_preview_tracker_dims_ft(self, template_ref):
        """Compute physical (width_ft, length_ft) for a tracker from its template.
        
        Width = E-W dimension (across tracker, short side)
        Length = N-S dimension (along tracker, long side)
        
        Returns (width_ft, length_ft) or None if template not found.
        """
        if not template_ref or template_ref not in self.enabled_templates:
            return None
        
        tdata = self.enabled_templates[template_ref]
        module_spec = tdata.get('module_spec', {})
        
        module_length_mm = module_spec.get('length_mm', 2000)
        module_width_mm = module_spec.get('width_mm', 1000)
        orientation = tdata.get('module_orientation', 'Portrait')
        modules_per_string = tdata.get('modules_per_string', 28)
        strings_per_tracker = tdata.get('strings_per_tracker', 2)
        modules_high = tdata.get('modules_high', 1)
        module_spacing_m = tdata.get('module_spacing_m', 0.02)
        has_motor = tdata.get('has_motor', True)
        motor_gap_m = tdata.get('motor_gap_m', 1.0) if has_motor else 0
        
        # Module dimensions along vs across the tracker
        if orientation == 'Portrait':
            # Portrait: module width runs N-S (along tracker), length runs E-W (across)
            mod_along_m = module_width_mm / 1000
            mod_across_m = module_length_mm / 1000
        else:
            # Landscape: module length runs N-S (along tracker), width runs E-W (across)
            mod_along_m = module_length_mm / 1000
            mod_across_m = module_width_mm / 1000
        
        # N-S length: all modules in one string laid end-to-end, times strings, plus gaps and motor
        full_spt = int(strings_per_tracker)
        partial_mods = round((strings_per_tracker - full_spt) * modules_per_string) if strings_per_tracker != full_spt else 0
        modules_in_row = full_spt * modules_per_string + partial_mods
        tracker_length_m = (modules_in_row * mod_along_m + 
                           (modules_in_row - 1) * module_spacing_m +
                           motor_gap_m)
        
        # E-W width: module across dimension times modules_high
        tracker_width_m = mod_across_m * modules_high
        
        # Convert to feet
        m_to_ft = 3.28084

        return (tracker_width_m * m_to_ft, tracker_length_m * m_to_ft)

    def get_motor_position_in_tracker(self, template_ref):
        """Compute the motor's Y offset from the tracker top (north end), in feet.
        
        Returns (motor_y_offset_ft, motor_gap_ft, has_motor) or (0, 0, False).
        """
        if not template_ref or template_ref not in self.enabled_templates:
            return 0, 0, False
        
        tdata = self.enabled_templates[template_ref]
        has_motor = tdata.get('has_motor', True)
        if not has_motor:
            return 0, 0, False
        
        module_spec = tdata.get('module_spec', {})
        module_length_mm = module_spec.get('length_mm', 2000)
        module_width_mm = module_spec.get('width_mm', 1000)
        orientation = tdata.get('module_orientation', 'Portrait')
        modules_per_string = tdata.get('modules_per_string', 28)
        strings_per_tracker = tdata.get('strings_per_tracker', 2)
        module_spacing_m = tdata.get('module_spacing_m', 0.02)
        motor_gap_m = tdata.get('motor_gap_m', 1.0)
        motor_placement = tdata.get('motor_placement_type', 'between_strings')
        motor_position_after_string = tdata.get('motor_position_after_string', None)
        motor_string_index_raw = tdata.get('motor_string_index', None)
        motor_split_north = tdata.get('motor_split_north', modules_per_string // 2)
        
        if orientation == 'Portrait':
            mod_along_m = module_width_mm / 1000
        else:
            mod_along_m = module_length_mm / 1000
        
        m_to_ft = 3.28084
        
        # Partial string on north pushes motor further south
        partial_north_m = 0
        spt_val = tdata.get('strings_per_tracker', 1)
        if spt_val != int(spt_val) and tdata.get('partial_string_side', 'north') == 'north':
            partial_north_mods = round((spt_val - int(spt_val)) * modules_per_string)
            partial_north_m = partial_north_mods * (mod_along_m + module_spacing_m)
        
        if motor_placement == 'between_strings':
            pos_after = motor_position_after_string if motor_position_after_string is not None else (motor_string_index_raw if motor_string_index_raw is not None else 1)
            modules_north = pos_after * modules_per_string
            if modules_north > 0:
                motor_y_m = partial_north_m + (modules_north * mod_along_m + 
                            (modules_north - 1) * module_spacing_m +
                            module_spacing_m)
            else:
                motor_y_m = partial_north_m
        elif motor_placement == 'middle_of_string':
            string_idx = motor_string_index_raw if motor_string_index_raw is not None else 1
            modules_before_split = (string_idx - 1) * modules_per_string + motor_split_north
            motor_y_m = partial_north_m + (modules_before_split * mod_along_m + 
                        (modules_before_split - 1) * module_spacing_m +
                        module_spacing_m)
        else:
            # Fallback: center
            dims = self.get_tracker_dimensions_ft(template_ref)
            if dims:
                return dims[1] / 2, motor_gap_m * m_to_ft, True
            return 0, 0, False
        
        return motor_y_m * m_to_ft, motor_gap_m * m_to_ft, True
    
    def setup_ui(self):
        """Create the preview window UI"""
        # Top bar with controls
        top_bar = ttk.Frame(self, padding="5")
        top_bar.pack(fill='x')
        
        ttk.Button(top_bar, text="Fit to Window", command=self.fit_and_redraw).pack(side='left', padx=2)
        ttk.Button(top_bar, text="Zoom In", command=lambda: self.zoom(1.3)).pack(side='left', padx=2)
        ttk.Button(top_bar, text="Zoom Out", command=lambda: self.zoom(0.7)).pack(side='left', padx=2)
        ttk.Button(top_bar, text="Reset Positions", command=self._reset_positions).pack(side='left', padx=2)
        ttk.Button(top_bar, text="Refresh Allocation", command=self._refresh_allocation).pack(side='left', padx=2)
        
        self.lock_btn = ttk.Button(top_bar, text="Lock Allocation", command=self._toggle_allocation_lock)
        self.lock_btn.pack(side='left', padx=2)
        self._update_lock_button()
        
        ttk.Separator(top_bar, orient='vertical').pack(side='left', fill='y', padx=8, pady=2)
        
        self.align_motor_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            top_bar, text="Align on Motor",
            variable=self.align_motor_var,
            command=self._on_alignment_toggle
        ).pack(side='left', padx=4)
        
        ttk.Separator(top_bar, orient='vertical').pack(side='left', fill='y', padx=8, pady=2)
        
        ttk.Label(top_bar, text="Mode:").pack(side='left', padx=(0, 4))
        
        self.inspect_mode_var = tk.BooleanVar(value=self.inspect_mode)
        
        toggle_frame = ttk.Frame(top_bar)
        toggle_frame.pack(side='left', padx=4)
        
        self.toggle_canvas = tk.Canvas(toggle_frame, width=52, height=24,
                                        highlightthickness=0, bg=top_bar.winfo_toplevel().cget('bg'))
        self.toggle_canvas.pack(side='left')
        self.toggle_canvas.bind('<Button-1>', self._on_toggle_click)
        
        self.toggle_label = ttk.Label(toggle_frame, text="Layout", foreground='#333333')
        self.toggle_label.pack(side='left', padx=(4, 0))
        
        self._draw_toggle()

        # Sync label to initial state
        if self.inspect_mode:
            self.toggle_label.config(text="Inspect", foreground='#4CAF50')
        
        ttk.Separator(top_bar, orient='vertical').pack(side='left', fill='y', padx=8, pady=2)
        
        self.add_pad_btn = ttk.Button(top_bar, text="+ Add Pad", command=self._add_pad)
        self.add_pad_btn.pack(side='left', padx=4)
        
        self.assign_btn = ttk.Button(top_bar, text="Assign Devices", command=self._show_assignment_dialog)
        self.assign_btn.pack(side='left', padx=4)
        
        if self.device_label == 'CB':
            self.edit_cbs_btn = ttk.Button(top_bar, text="Edit CBs", command=self._show_cb_assignment_dialog)
            self.edit_cbs_btn.pack(side='left', padx=4)

        self.show_routes_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            top_bar, text="Show Routes",
            variable=self.show_routes_var,
            command=self.draw
        ).pack(side='left', padx=4)
        
        self.zoom_label = ttk.Label(top_bar, text="100%")
        self.zoom_label.pack(side='left', padx=10)
        
        # Summary info
        num_inv = self.inv_summary.get('total_inverters', 0)
        total_str = self.inv_summary.get('total_strings', 0)
        actual_ratio = self.inv_summary.get('actual_dc_ac', 0)
        split = self.inv_summary.get('total_split_trackers', 0)
        
        summary_text = f"{num_inv} Inverters  |  {total_str} Strings  |  DC:AC: {actual_ratio:.2f}  |  {split} Split Trackers  |  {self.topology}"
        self.summary_label = ttk.Label(top_bar, text=summary_text, foreground='#333333')
        self.summary_label.pack(side='right', padx=10)
        
        # Canvas
        canvas_frame = ttk.Frame(self)
        canvas_frame.pack(fill='both', expand=True)
        
        self.canvas = tk.Canvas(canvas_frame, bg='white', highlightthickness=0)
        self.canvas.pack(fill='both', expand=True)
        
        # Bind events
        self.canvas.bind('<MouseWheel>', self.on_mousewheel)
        self.canvas.bind('<Button-4>', lambda e: self.zoom(1.1))
        self.canvas.bind('<Button-5>', lambda e: self.zoom(0.9))
        # Left-click: group select/drag only
        self.canvas.bind('<ButtonPress-1>', self.on_press)
        self.canvas.bind('<Double-Button-1>', self._on_device_double_click)
        self.canvas.bind('<B1-Motion>', self.on_motion)
        self.canvas.bind('<ButtonRelease-1>', self.on_release)
        # Middle-click: pan canvas
        self.canvas.bind('<ButtonPress-2>', self.on_pan_press)
        self.canvas.bind('<B2-Motion>', self.on_pan_motion)
        self.canvas.bind('<ButtonRelease-2>', self.on_pan_release)
        self.canvas.bind('<Button-3>', self._on_pad_right_click)
        self.canvas.bind('<Configure>', lambda e: self.draw())
        
        # Bottom legend (rebuildable)
        self.legend_frame = ttk.Frame(self, padding="5")
        self.legend_frame.pack(fill='x')
        self._build_legend()
    
    def _build_legend(self):
        """Build or rebuild the bottom legend with color swatches and allocation summary."""
        # Clear existing legend contents
        for child in self.legend_frame.winfo_children():
            child.destroy()
        
        # Color swatches
        swatch_frame = ttk.Frame(self.legend_frame)
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
        
        # Allocation summary
        allocation_result = self.inv_summary.get('allocation_result')
        if allocation_result:
            summary = allocation_result['summary']
            max_spi = summary.get('max_strings_per_inverter', 0)
            min_spi = summary.get('min_strings_per_inverter', 0)
            n_larger = summary.get('num_larger_inverters', 0)
            n_smaller = summary.get('num_smaller_inverters', 0)
            split_count = summary.get('total_split_trackers', 0)
            if max_spi == min_spi:
                size_str = f"All inverters: {max_spi} strings"
            else:
                size_str = f"{n_larger} inverters × {max_spi} strings + {n_smaller} inverters × {min_spi} strings"
            spatial_runs = allocation_result.get('spatial_runs', 1)
            runs_str = f"  |  {spatial_runs} spatial run(s)" if spatial_runs > 1 else ""
            ttk.Label(self.legend_frame, text=f"{size_str}  |  {split_count} split tracker(s){runs_str}",
                     font=('Helvetica', 9), foreground='#555555').pack(anchor='w')

    def build_layout_data(self):
        """Build a group-based layout of trackers with physical dimensions from templates.
        
        World units are in feet. Trackers run N-S (Y axis), spaced E-W (X axis).
        Each group has an (x, y) position in world-space representing its top-left corner.
        """
        self.group_layout = []
        self.selected_group_idx = None
        
        allocation_result = self.inv_summary.get('allocation_result')
        if not allocation_result:
            self.world_width = 0
            self.world_height = 0
            return
        
        # Build tracker_idx -> assignments from harness_map
        tracker_map = {}
        for inv_idx, inv in enumerate(allocation_result['inverters']):
            color = self.colors[inv_idx % len(self.colors)]
            for entry in inv['harness_map']:
                tidx = entry['tracker_idx']
                if tidx not in tracker_map:
                    tracker_map[tidx] = {
                        'strings_per_tracker': entry['strings_per_tracker'],
                        'assignments': []
                    }
                tracker_map[tidx]['assignments'].append({
                    'color': color,
                    'strings': entry['strings_taken'],
                    'inv_idx': inv_idx
                })
        
        # Split tracker_map into groups with template dimensions
        global_idx = 0
        max_tracker_length_ft = 0
        max_tracker_width_ft = 0
        
        # Fallback dimensions for unlinked trackers
        fallback_width_ft = 6.0
        fallback_length_ft = 180.0
        
        for grp_idx, group_data in enumerate(self.groups):
            group_trackers = []
            group_motor_y = None  # motor Y offset for first tracker in group (used for alignment)
            
            for seg in group_data['segments']:
                ref = seg.get('template_ref')
                dims = self._get_preview_tracker_dims_ft(ref)
                
                for _ in range(seg['quantity']):
                    if global_idx in tracker_map:
                        tracker = tracker_map[global_idx].copy()
                        if dims:
                            tracker['width_ft'] = dims[0]
                            tracker['length_ft'] = dims[1]
                        else:
                            tracker['width_ft'] = fallback_width_ft
                            tracker['length_ft'] = fallback_length_ft
                        tracker['template_ref'] = ref
                        
                        # Motor position
                        motor_y, motor_gap, has_motor = self.get_motor_position_in_tracker(ref)
                        tracker['motor_y_ft'] = motor_y
                        tracker['motor_gap_ft'] = motor_gap
                        tracker['has_motor'] = has_motor
                        # Partial string info from template
                        if ref and ref in self.enabled_templates:
                            tdata_ps = self.enabled_templates[ref]
                            raw_spt = tdata_ps.get('strings_per_tracker', 1)
                            if raw_spt != int(raw_spt):
                                mps_ps = tdata_ps.get('modules_per_string', 28)
                                tracker['partial_module_count'] = round((raw_spt - int(raw_spt)) * mps_ps)
                                tracker['partial_string_side'] = tdata_ps.get('partial_string_side', 'north')
                                tracker['full_string_count'] = int(raw_spt)
                            else:
                                tracker['partial_module_count'] = 0
                                tracker['partial_string_side'] = 'north'
                                tracker['full_string_count'] = int(raw_spt)
                        
                        if group_motor_y is None and has_motor:
                            group_motor_y = motor_y

                        group_trackers.append(tracker)
                        
                        max_tracker_width_ft = max(max_tracker_width_ft, tracker['width_ft'])
                        max_tracker_length_ft = max(max_tracker_length_ft, tracker['length_ft'])
                    global_idx += 1
            
            # Group dimensions (bounding box of all its trackers laid out E-W)
            num_trackers = len(group_trackers)
            if num_trackers > 0:
                group_max_width = max(t['width_ft'] for t in group_trackers)
                group_width = group_max_width + (num_trackers - 1) * self.row_spacing_ft
                group_length = max(t['length_ft'] for t in group_trackers)
            else:
                group_width = 0
                group_length = 0
            
            # Compute string length for NS snap offset (from first linked template)
            string_length_ft = 0
            for seg in group_data['segments']:
                ref = seg.get('template_ref')
                if ref and ref in self.enabled_templates:
                    tdata = self.enabled_templates[ref]
                    ms = tdata.get('module_spec', {})
                    mps = tdata.get('modules_per_string', 28)
                    orientation = tdata.get('module_orientation', 'Portrait')
                    if orientation == 'Portrait':
                        mod_along_m = ms.get('width_mm', 1000) / 1000
                    else:
                        mod_along_m = ms.get('length_mm', 2000) / 1000
                    spacing_m = tdata.get('module_spacing_m', 0.02)
                    string_length_ft = (mps * mod_along_m + (mps - 1) * spacing_m) * 3.28084
                    break
            
            # Compute visual bounding box offsets considering motor alignment
            # When align_on_motor is active, trackers shift vertically so their
            # motors match the group's reference motor_y. This affects the actual
            # visual extent of the group.
            ref_motor = group_motor_y or 0
            visual_min_y_offset = 0.0
            visual_max_y_offset = 0.0
            
            # Driveline angle: each tracker offset in Y by t_idx * pitch * tan(angle)
            driveline_angle_deg = group_data.get('driveline_angle', 0.0)
            driveline_angle_rad = math.radians(driveline_angle_deg)
            driveline_tan = math.tan(driveline_angle_rad) if driveline_angle_deg > 0 else 0.0
            
            visual_min_y_base = 0.0
            visual_max_y_base = 0.0
            
            if group_trackers and ref_motor is not None:
                for t_i, t in enumerate(group_trackers):
                    t_motor = t.get('motor_y_ft', 0)
                    t_length_val = t.get('length_ft', group_length)
                    y_offset = ref_motor - t_motor  # Motor alignment shift (no angle)
                    angle_y = t_i * self.row_spacing_ft * driveline_tan
                    # Base bounds (no angle) — for parallelogram overlap checking
                    visual_min_y_base = min(visual_min_y_base, y_offset)
                    visual_max_y_base = max(visual_max_y_base, y_offset + t_length_val)
                    # Full bounds (with angle) — for bounding box and selection highlight
                    visual_min_y_offset = min(visual_min_y_offset, y_offset + angle_y)
                    visual_max_y_offset = max(visual_max_y_offset, y_offset + angle_y + t_length_val)
            
            self.group_layout.append({
                'name': group_data['name'],
                'trackers': group_trackers,
                'width_ft': group_width,
                'length_ft': group_length,
                'motor_y_ft': group_motor_y or 0,
                'string_length_ft': string_length_ft,
                'group_idx': grp_idx,
                'visual_min_y': visual_min_y_offset,
                'visual_max_y': visual_max_y_offset,
                'visual_min_y_base': visual_min_y_base,
                'visual_max_y_base': visual_max_y_base,
                'driveline_angle': driveline_angle_deg,
                'driveline_tan': driveline_tan,
            })

        # Flat list for backward compat
        self.tracker_list = [tracker_map[i] for i in sorted(tracker_map.keys())]
        
        # Store global metrics
        self.tracker_pitch_ft = self.row_spacing_ft
        self.tracker_gap_ft = max(self.row_spacing_ft - max_tracker_width_ft, 1.0)
        self.max_tracker_width_ft = max_tracker_width_ft if max_tracker_width_ft > 0 else fallback_width_ft
        self.max_tracker_length_ft = max_tracker_length_ft if max_tracker_length_ft > 0 else fallback_length_ft
        self.group_ns_gap_ft = self.max_tracker_length_ft * 0.15
        
        # Assign initial positions from saved data or auto-layout
        self._assign_group_positions()
        
        # Compute device (CB/SI) positions per group
        self._compute_device_positions()
        
        # Compute world bounds from actual positions (including devices)
        self._update_world_bounds()
    
    def _assign_group_positions(self):
        """Assign (x, y) positions to each group. Use saved positions if available,
        otherwise auto-layout stacking groups left-to-right."""
        for grp_idx, layout in enumerate(self.group_layout):
            group_data = self.groups[grp_idx] if grp_idx < len(self.groups) else {}
            
            saved_x = group_data.get('position_x')
            saved_y = group_data.get('position_y')
            
            if saved_x is not None and saved_y is not None:
                layout['x'] = saved_x
                layout['y'] = saved_y
            else:
                # Auto-layout: stack groups left to right with spacing
                # Each group is placed one group-width + gap to the right
                group_spacing = self.max_tracker_length_ft * 0.1
                layout['x'] = grp_idx * (layout['width_ft'] + group_spacing)
                layout['y'] = 0
    
    def _compute_device_positions(self):
        """Compute world-space positions for combiner boxes / string inverters.
        
        For Distributed String and Centralized String: derives placement from
        the allocation result so each device maps 1:1 to an inverter.
        
        For Central Inverter: distributes CBs proportionally across groups
        by tracker count.
        
        Device Y is determined by the group's device_position setting:
          - 'north': offset above northernmost tracker edge
          - 'south': offset below southernmost tracker edge  
          - 'middle': at the motor/driveline Y
        """
        self.device_positions = []
        
        if not self.group_layout:
            return
        
        device_width_ft = 4.0
        device_height_ft = 3.0
        offset_ft = 5.0
        
        alloc = self.inv_summary.get('allocation_result', {}) if hasattr(self, 'inv_summary') else {}
        inverters = alloc.get('inverters', [])
        
        if self.topology in ('Distributed String', 'Centralized String') and inverters:
            self._compute_devices_from_allocation(
                inverters, device_width_ft, device_height_ft, offset_ft
            )
        elif self.topology == 'Central Inverter':
            self._compute_devices_proportional(
                device_width_ft, device_height_ft, offset_ft
            )
    
    def _compute_devices_from_allocation(self, inverters, device_width_ft, device_height_ft, offset_ft):
        """Place one device per inverter, positioned at the center of that inverter's trackers."""
        pitch = self.tracker_pitch_ft
        max_width = self.max_tracker_width_ft
        
        # Build a lookup: global_tracker_idx -> (group_idx, local_tracker_idx)
        tracker_to_group = {}
        running = 0
        for grp_idx, grp in enumerate(self.group_layout):
            for local_idx in range(len(grp['trackers'])):
                tracker_to_group[running] = (grp_idx, local_idx)
                running += 1
        
        for inv_idx, inv in enumerate(inverters):
            harness_map = inv.get('harness_map', [])
            if not harness_map:
                continue
            
            # Find which trackers this inverter uses
            inv_tracker_indices = [entry['tracker_idx'] for entry in harness_map]
            
            # Determine majority group
            group_counts = {}
            for tidx in inv_tracker_indices:
                if tidx in tracker_to_group:
                    grp_idx = tracker_to_group[tidx][0]
                    group_counts[grp_idx] = group_counts.get(grp_idx, 0) + 1
            
            if not group_counts:
                continue
            
            primary_grp_idx = max(group_counts, key=group_counts.get)
            group_data = self.group_layout[primary_grp_idx]
            group_source = self.groups[primary_grp_idx] if primary_grp_idx < len(self.groups) else {}
            device_position = group_source.get('device_position', 'middle')
            
            gx = group_data['x']
            gy = group_data['y']
            
            # Compute X from the center of this inverter's trackers within the primary group
            local_indices = []
            for tidx in inv_tracker_indices:
                if tidx in tracker_to_group and tracker_to_group[tidx][0] == primary_grp_idx:
                    local_indices.append(tracker_to_group[tidx][1])
            
            if local_indices:
                center_local = (min(local_indices) + max(local_indices)) / 2.0
                device_x = gx + center_local * pitch + (max_width - device_width_ft) / 2
            else:
                center_local = 0
                device_x = gx
            
            # Driveline angle Y offset based on device's X position in group
            angle_y_offset = center_local * pitch * group_data.get('driveline_tan', 0.0)
            
            # Compute Y based on position setting
            if device_position == 'north':
                vis_min = group_data.get('visual_min_y', 0)
                device_y = gy + vis_min - offset_ft - device_height_ft + angle_y_offset
            elif device_position == 'south':
                vis_max = group_data.get('visual_max_y', group_data['length_ft'])
                device_y = gy + vis_max + offset_ft + angle_y_offset
            else:  # 'middle'
                motor_y = group_data.get('motor_y_ft', group_data['length_ft'] / 2)
                device_y = gy + motor_y - device_height_ft / 2 + angle_y_offset
            
            # Build assigned_strings from this inverter's harness_map
            assigned_strings = {}
            for entry in harness_map:
                tidx = entry['tracker_idx']
                strings_taken = entry['strings_taken']
                spt = entry['strings_per_tracker']
                is_split = entry.get('is_split', False)
                split_pos = entry.get('split_position', 'full')
                
                if tidx not in assigned_strings:
                    assigned_strings[tidx] = set()
                
                if is_split and split_pos == 'tail':
                    # Tail of a split: strings are at the END of the tracker
                    start_idx = spt - strings_taken
                    for s in range(start_idx, spt):
                        assigned_strings[tidx].add(s)
                else:
                    # Head of split or full tracker: strings from the front
                    existing = len(assigned_strings[tidx])
                    for s in range(existing, existing + strings_taken):
                        assigned_strings[tidx].add(s)
            
            dev_idx = len(self.device_positions)
            label = self.device_names.get(dev_idx, f"{self.device_label}-{inv_idx + 1:02d}")
            
            self.device_positions.append({
                'x': device_x,
                'y': device_y,
                'width_ft': device_width_ft,
                'height_ft': device_height_ft,
                'label': label,
                'group_idx': primary_grp_idx,
                'device_position': device_position,
                'assigned_strings': assigned_strings,
            })
    
    def _compute_devices_proportional(self, device_width_ft, device_height_ft, offset_ft):
        """Distribute devices proportionally across groups for Central Inverter topology."""
        pitch = self.tracker_pitch_ft
        max_width = self.max_tracker_width_ft
        
        total_trackers = sum(len(g['trackers']) for g in self.group_layout)
        if total_trackers <= 0 or self.num_devices <= 0:
            return
        
        global_device_idx = 0
        
        for grp_idx, group_data in enumerate(self.group_layout):
            group_trackers = group_data['trackers']
            num_trackers_in_group = len(group_trackers)
            if num_trackers_in_group == 0:
                continue
            
            group_share = num_trackers_in_group / total_trackers
            group_device_count = max(1, round(group_share * self.num_devices))
            remaining = self.num_devices - global_device_idx
            group_device_count = min(group_device_count, remaining)
            
            if group_device_count <= 0:
                continue
            
            group_source = self.groups[grp_idx] if grp_idx < len(self.groups) else {}
            device_position = group_source.get('device_position', 'middle')
            
            gx = group_data['x']
            gy = group_data['y']
            
            # Base Y (before driveline angle offset)
            if device_position == 'north':
                vis_min = group_data.get('visual_min_y', 0)
                base_device_y = gy + vis_min - offset_ft - device_height_ft
            elif device_position == 'south':
                vis_max = group_data.get('visual_max_y', group_data['length_ft'])
                base_device_y = gy + vis_max + offset_ft
            else:
                motor_y = group_data.get('motor_y_ft', group_data['length_ft'] / 2)
                base_device_y = gy + motor_y - device_height_ft / 2
            
            driveline_tan = group_data.get('driveline_tan', 0.0)
            
            # Even spacing within group
            group_global_start = sum(len(self.group_layout[g]['trackers']) for g in range(grp_idx))
            base_group_size = num_trackers_in_group // group_device_count
            extra = num_trackers_in_group % group_device_count
            
            tracker_start = 0
            for dev_i in range(group_device_count):
                sub_size = base_group_size + (1 if dev_i < extra else 0)
                if sub_size <= 0:
                    continue
                
                center_tracker = tracker_start + sub_size / 2.0 - 0.5
                device_x = gx + center_tracker * pitch + (max_width - device_width_ft) / 2
                device_y = base_device_y + center_tracker * pitch * driveline_tan
                
                # All strings in tracker range belong to this CB
                assigned_strings = {}
                for local_idx in range(tracker_start, tracker_start + sub_size):
                    global_idx = group_global_start + local_idx
                    if local_idx < len(group_trackers):
                        spt = group_trackers[local_idx].get('strings_per_tracker', 0)
                        assigned_strings[global_idx] = set(range(int(spt)))
                
                self.device_positions.append({
                    'x': device_x,
                    'y': device_y,
                    'width_ft': device_width_ft,
                    'height_ft': device_height_ft,
                    'label': self.device_names.get(global_device_idx, f"CB-{global_device_idx + 1:02d}"),
                    'group_idx': grp_idx,
                    'device_position': device_position,
                    'assigned_strings': assigned_strings,
                })
                
                global_device_idx += 1
                tracker_start += sub_size
        
        # Fill any remaining devices
        while global_device_idx < self.num_devices and self.device_positions:
            last = self.device_positions[-1].copy()
            last['label'] = self.device_names.get(global_device_idx, f"CB-{global_device_idx + 1:02d}")
            last['x'] += device_width_ft + 2
            last['assigned_strings'] = {}
            self.device_positions.append(last)
            global_device_idx += 1

    def _update_world_bounds(self):
        """Recompute world_width and world_height from actual group positions."""
        if not self.group_layout:
            self.world_width = 0
            self.world_height = 0
            return
        
        min_x = min(g['x'] for g in self.group_layout)
        max_x = max(g['x'] + g['width_ft'] for g in self.group_layout)
        min_y = min(g['y'] for g in self.group_layout)
        max_y = max(g['y'] + g['length_ft'] for g in self.group_layout)
        
        # Include device positions in bounds
        if hasattr(self, 'device_positions') and self.device_positions:
            for dev in self.device_positions:
                min_x = min(min_x, dev['x'])
                max_x = max(max_x, dev['x'] + dev['width_ft'])
                min_y = min(min_y, dev['y'])
                max_y = max(max_y, dev['y'] + dev['height_ft'])

        # Include pads in bounds
        if hasattr(self, 'pads') and self.pads:
            for pad in self.pads:
                pw = pad.get('width_ft', 10.0)
                ph = pad.get('height_ft', 8.0)
                min_x = min(min_x, pad['x'])
                max_x = max(max_x, pad['x'] + pw)
                min_y = min(min_y, pad['y'])
                max_y = max(max_y, pad['y'] + ph)
        
        # Add margin
        margin = self.max_tracker_width_ft * 2
        self.world_min_x = min_x - margin
        self.world_min_y = min_y - margin
        self.world_width = (max_x - min_x) + margin * 2
        self.world_height = (max_y - min_y) + margin * 2
    
    def _save_group_positions(self):
        """Save current group positions back to the group data dicts."""
        for grp_idx, layout in enumerate(self.group_layout):
            if grp_idx < len(self.groups):
                self.groups[grp_idx]['position_x'] = layout['x']
                self.groups[grp_idx]['position_y'] = layout['y']
    
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
        
        # Center on actual content bounds
        min_x = getattr(self, 'world_min_x', 0)
        min_y = getattr(self, 'world_min_y', 0)
        
        scaled_w = self.world_width * self.scale
        scaled_h = self.world_height * self.scale
        self.pan_x = (cw - scaled_w) / 2 - min_x * self.scale
        self.pan_y = (ch - scaled_h) / 2 - min_y * self.scale
    
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
    
    def canvas_to_world(self, cx, cy):
        """Convert canvas pixel coordinates to world-space feet."""
        if self.scale == 0:
            return 0, 0
        wx = (cx - self.pan_x) / self.scale
        wy = (cy - self.pan_y) / self.scale
        return wx, wy
    
    def hit_test_group(self, cx, cy):
        """Return the index of the group under canvas coords (cx, cy), or None."""
        wx, wy = self.canvas_to_world(cx, cy)
        # Check in reverse order so topmost (last drawn) is hit first
        for i in range(len(self.group_layout) - 1, -1, -1):
            g = self.group_layout[i]
            vis_min = g.get('visual_min_y', 0)
            vis_max = g.get('visual_max_y', g['length_ft'])
            if (g['x'] <= wx <= g['x'] + g['width_ft'] and
                g['y'] + vis_min <= wy <= g['y'] + vis_max):
                return i
        return None
    
    def on_press(self, event):
        """Handle left mouse press — place pad, select/drag group or pad, or select device."""
        self.drag_start_x = event.x
        self.drag_start_y = event.y
        self._drag_moved = False
        self._dragging_pad = False
        
        # Pad placement mode — click to place
        if self.placing_pad:
            wx, wy = self.canvas_to_world(event.x, event.y)
            self._place_pad_at(wx, wy)
            return
        
        if self.inspect_mode:
            # Check pads first (drawn on top)
            hit_pad = self.hit_test_pad(event.x, event.y)
            if hit_pad is not None:
                if self.selected_pad_inspect_idx == hit_pad:
                    self.selected_pad_inspect_idx = None
                else:
                    self.selected_pad_inspect_idx = hit_pad
                self.selected_device_idx = None
                self.dragging_canvas = False
                self.draw()
                self.dragging_group = False
                return
            
            # Then check devices
            hit_dev = self.hit_test_device(event.x, event.y)
            if hit_dev is not None:
                if self.selected_device_idx == hit_dev:
                    self.selected_device_idx = None
                else:
                    self.selected_device_idx = hit_dev
                self.selected_pad_inspect_idx = None
                self.dragging_canvas = False
                self.draw()
            else:
                # Empty space — clear selections and start panning
                self.selected_device_idx = None
                self.selected_pad_inspect_idx = None
                self.dragging_canvas = True
                self.draw()
            self.dragging_group = False
            return
        
        # Layout mode — check pads first (they're on top visually)
        hit_pad = self.hit_test_pad(event.x, event.y)
        if hit_pad is not None:
            self.selected_pad_idx = hit_pad
            self._dragging_pad = True
            pad = self.pads[hit_pad]
            self._drag_pad_start_x = pad['x']
            self._drag_pad_start_y = pad['y']
            self.selected_group_idx = None
            self.dragging_group = False
            self.draw()
            return
        
        self.selected_pad_idx = None
        
        hit = self.hit_test_group(event.x, event.y)
        if hit is not None:
            self.selected_group_idx = hit
            self.dragging_group = True
            g = self.group_layout[hit]
            self._drag_group_start_x = g['x']
            self._drag_group_start_y = g['y']
            self.draw()
        else:
            self.selected_group_idx = None
            self.dragging_group = False
            self.draw()

    def _on_device_double_click(self, event):
        """Handle double-click on canvas — rename device if clicked on one."""
        dev_idx = self.hit_test_device(event.x, event.y)
        if dev_idx is None:
            return
        
        dev = self.device_positions[dev_idx]
        current_name = dev.get('label', f"{self.device_label}-{dev_idx + 1:02d}")
        
        from tkinter import simpledialog
        new_name = simpledialog.askstring(
            "Rename Device",
            f"Enter new name for {current_name}:",
            initialvalue=current_name,
            parent=self
        )
        
        if new_name and new_name.strip():
            new_name = new_name.strip()
            self.device_names[dev_idx] = new_name
            dev['label'] = new_name
            self.draw()
    
    def on_motion(self, event):
        """Handle mouse drag — move group, move pad, or pan canvas."""
        dx_px = event.x - self.drag_start_x
        dy_px = event.y - self.drag_start_y
        
        if abs(dx_px) > 3 or abs(dy_px) > 3:
            self._drag_moved = True
        
        if getattr(self, '_dragging_pad', False) and self.selected_pad_idx is not None:
            dx_world = dx_px / self.scale if self.scale != 0 else 0
            dy_world = dy_px / self.scale if self.scale != 0 else 0
            self.pads[self.selected_pad_idx]['x'] = self._drag_pad_start_x + dx_world
            self.pads[self.selected_pad_idx]['y'] = self._drag_pad_start_y + dy_world
            self.draw()
        elif getattr(self, 'dragging_group', False) and self.selected_group_idx is not None:
            dx_world = dx_px / self.scale if self.scale != 0 else 0
            dy_world = dy_px / self.scale if self.scale != 0 else 0
            
            new_x = self._drag_group_start_x + dx_world
            new_y = self._drag_group_start_y + dy_world
            
            shift_held = event.state & 0x1
            if shift_held:
                # Shift held — constrain to N/S movement only (lock X)
                new_x = self._drag_group_start_x
            else:
                # Normal drag — apply snapping
                new_x, new_y = self._snap_group_position(
                    self.selected_group_idx, new_x, new_y
                )
            
            self.group_layout[self.selected_group_idx]['x'] = new_x
            self.group_layout[self.selected_group_idx]['y'] = new_y
            
            self.draw()
    
    def on_release(self, event):
        """Handle mouse release — finalize group or pad position."""
        if getattr(self, '_dragging_pad', False) and self._drag_moved:
            self._update_world_bounds()
        
        if getattr(self, 'dragging_group', False) and self._drag_moved:
            self._update_world_bounds()
            self._save_group_positions()
            
            overlaps = self._check_overlaps()
            if overlaps:
                names = set()
                for i, j in overlaps:
                    names.add(self.group_layout[i].get('name', f'Group {i+1}'))
                    names.add(self.group_layout[j].get('name', f'Group {j+1}'))
        
        self.dragging_group = False
        self.dragging_canvas = False
        self._dragging_pad = False

    def on_pan_press(self, event):
        """Handle middle mouse press — start panning."""
        self.dragging_canvas = True
        self.drag_start_x = event.x
        self.drag_start_y = event.y

    def on_pan_motion(self, event):
        """Handle middle mouse drag — pan canvas."""
        if self.dragging_canvas:
            dx_px = event.x - self.drag_start_x
            dy_px = event.y - self.drag_start_y
            self.pan_x += dx_px
            self.pan_y += dy_px
            self.drag_start_x = event.x
            self.drag_start_y = event.y
            self.draw()

    def on_pan_release(self, event):
        """Handle middle mouse release — stop panning."""
        self.dragging_canvas = False
    
    def _snap_group_position(self, group_idx, raw_x, raw_y):
        """Apply snapping to a group's proposed position.
        
        EW (X): Snap to row-spacing pitch grid.
        NS (Y): Snap motor to align with ANY nearby group's motor.
        Checks all groups and picks the closest snap candidate.
        """
        group = self.group_layout[group_idx]
        my_motor_offset = group.get('motor_y_ft', 0)
        string_len = group.get('string_length_ft', 0)
        
        snapped_x = raw_x
        snapped_y = raw_y
        
        # EW snap: align to pitch grid
        pitch = self.tracker_pitch_ft
        if pitch > 0:
            snapped_x = round(raw_x / pitch) * pitch
        
        # NS snap: check motor alignment against ALL other groups
        snap_threshold = self.max_tracker_length_ft * 0.15
        best_snap_y = None
        best_snap_dist = float('inf')
        
        for i, g in enumerate(self.group_layout):
            if i == group_idx:
                continue
            
            neighbor_motor_world_y = g['y'] + g.get('motor_y_ft', 0)
            
            # Target Y so motors align
            motor_aligned_y = neighbor_motor_world_y - my_motor_offset
            
            # Also compute string-offset positions
            candidates = [
                motor_aligned_y,                     # motor alignment
                motor_aligned_y + string_len,        # offset +1 string south
                motor_aligned_y - string_len,        # offset +1 string north
            ]
            
            for candidate_y in candidates:
                dist = abs(raw_y - candidate_y)
                if dist < best_snap_dist and dist < snap_threshold:
                    best_snap_dist = dist
                    best_snap_y = candidate_y
        
        if best_snap_y is not None:
            snapped_y = best_snap_y
        
        return snapped_x, snapped_y
    
    def draw(self):
        """Draw the site layout with to-scale trackers at their group positions.
        
        X = E-W (tracker width + row spacing gaps)
        Y = N-S (tracker length, north at top)
        World units = feet.
        """
        self.canvas.delete('all')
        
        if not self.group_layout:
            return
        
        pitch = getattr(self, 'tracker_pitch_ft', 20)
        max_width = getattr(self, 'max_tracker_width_ft', 6)
        
        for group_idx, group_data in enumerate(self.group_layout):
            gx = group_data['x']
            gy = group_data['y']
            is_selected = (group_idx == self.selected_group_idx)
            
            # Draw selection highlight behind group (using visual bounds)
            if is_selected:
                pad = max_width * 0.3
                vis_min = group_data.get('visual_min_y', 0)
                vis_max = group_data.get('visual_max_y', group_data['length_ft'])
                hx1, hy1 = self.world_to_canvas(gx - pad, gy + vis_min - pad)
                hx2, hy2 = self.world_to_canvas(
                    gx + group_data['width_ft'] + pad,
                    gy + vis_max + pad
                )
                self.canvas.create_rectangle(
                    hx1, hy1, hx2, hy2,
                    fill='', outline='#4A90D9', width=2, dash=(6, 3)
                )
            
            # Draw group label
            label_x, label_y = self.world_to_canvas(
                gx - max_width * 0.5,
                gy + group_data['length_ft'] / 2
            )
            font_size = max(6, min(11, int(9 * self.scale)))
            self.canvas.create_text(
                label_x, label_y,
                text=group_data['name'], font=('Helvetica', font_size),
                fill='#4A90D9' if is_selected else '#333333', anchor='e'
            )
            
            for t_idx, tracker in enumerate(group_data['trackers']):
                spt = tracker['strings_per_tracker']
                assignments = tracker['assignments']
                t_width = tracker.get('width_ft', max_width)
                t_length = tracker.get('length_ft', 100)
                
                # X position within group: center-to-center = pitch
                tx = gx + t_idx * pitch
                # Center tracker within pitch slot
                tx_offset = (max_width - t_width) / 2 if max_width > t_width else 0
                
                # Driveline angle: offset each tracker in Y
                angle_y_offset = t_idx * pitch * group_data.get('driveline_tan', 0.0)
                
                # Align tracker vertically within group
                if getattr(self, 'align_on_motor', False) and tracker.get('has_motor', False) and group_data.get('motor_y_ft', None) is not None:
                    # Motor alignment: offset so this tracker's motor Y matches group's reference motor Y
                    reference_motor_y = group_data['motor_y_ft']
                    ty = gy + (reference_motor_y - tracker['motor_y_ft']) + angle_y_offset
                else:
                    # Center alignment fallback
                    max_group_length = group_data['length_ft']
                    ty_offset = (max_group_length - t_length) / 2
                    ty = gy + ty_offset + angle_y_offset              
                # Per-string height — adjust for partial strings
                partial_mods = tracker.get('partial_module_count', 0)
                partial_side = tracker.get('partial_string_side', 'north')
                full_str_count = tracker.get('full_string_count', spt)
                mps_for_height = 26  # fallback
                
                ref = tracker.get('template_ref')
                if partial_mods > 0 and ref and ref in self.enabled_templates:
                    mps_for_height = self.enabled_templates[ref].get('modules_per_string', 26)
                
                if partial_mods > 0 and full_str_count > 0:
                    total_mods = full_str_count * mps_for_height + partial_mods
                    module_extent = t_length
                    full_height = (module_extent * mps_for_height / total_mods) if total_mods > 0 else module_extent
                    partial_height = (module_extent * partial_mods / total_mods) if total_mods > 0 else 0
                    
                    # Build height list per effective string slot
                    # Always include the partial band (even for right-of-pair trackers)
                    string_heights = []
                    has_owned_partial = (spt > full_str_count)
                    draw_spt = spt + (1 if not has_owned_partial else 0)  # Add unowned partial band
                    
                    if partial_side == 'north':
                        string_heights.append(partial_height)  # Always draw partial band
                        for _ in range(full_str_count):
                            string_heights.append(full_height)
                    else:  # south
                        for _ in range(full_str_count):
                            string_heights.append(full_height)
                        string_heights.append(partial_height)  # Always draw partial band
                else:
                    string_height = t_length / spt if spt > 0 else t_length
                    string_heights = [string_height] * int(spt)
                
                # Build string colors
                string_colors = []
                for assignment in assignments:
                    for _ in range(assignment['strings']):
                        string_colors.append(assignment['color'])
                
                # Determine global tracker index for device highlighting
                global_tracker_idx = sum(
                    len(self.group_layout[g]['trackers']) for g in range(group_idx)
                ) + t_idx
                
                # Check if we're highlighting a selected device or pad
                highlighting = False
                selected_strings = set()
                
                if self.inspect_mode and hasattr(self, 'device_positions') and self.device_positions:
                    if self.selected_device_idx is not None:
                        highlighting = True
                        dev = self.device_positions[self.selected_device_idx]
                        assigned = dev.get('assigned_strings', {})
                        selected_strings = assigned.get(global_tracker_idx, set())
                    elif self.selected_pad_inspect_idx is not None:
                        highlighting = True
                        # Collect strings from ALL devices assigned to this pad
                        pad = self.pads[self.selected_pad_inspect_idx] if self.selected_pad_inspect_idx < len(self.pads) else None
                        if pad:
                            for dev_idx in pad.get('assigned_devices', []):
                                if dev_idx < len(self.device_positions):
                                    dev = self.device_positions[dev_idx]
                                    assigned = dev.get('assigned_strings', {})
                                    selected_strings.update(assigned.get(global_tracker_idx, set()))
                
                # Draw each string (including unowned partial bands)
                draw_count = len(string_heights)
                for s_idx in range(draw_count):
                    # Determine if this is an unowned partial band
                    is_unowned_partial = (partial_mods > 0 and spt <= full_str_count and
                                         ((partial_side == 'north' and s_idx == 0) or
                                          (partial_side == 'south' and s_idx == draw_count - 1)))
                    
                    if is_unowned_partial:
                        color = '#D4C878'  # Muted gold for unowned partial
                    else:
                        # Map drawing index back to allocation string index
                        # Skip the partial band position to get the right color
                        if partial_mods > 0 and partial_side == 'north':
                            color_idx = s_idx - 1  # Partial is at 0, so shift down
                            if spt > full_str_count and s_idx == 0:
                                color_idx = 0  # Owned partial gets first color
                        elif partial_mods > 0 and partial_side == 'south':
                            if spt > full_str_count and s_idx == draw_count - 1:
                                color_idx = spt - 1  # Owned partial gets last color
                            else:
                                color_idx = s_idx
                        else:
                            color_idx = s_idx
                        color = string_colors[color_idx] if 0 <= color_idx < len(string_colors) else '#D0D0D0'
                    
                    if highlighting:
                        if s_idx in selected_strings:
                            outline_color = '#FF6600'
                            outline_width = 2
                        else:
                            # Dim unselected strings
                            color = '#E0E0E0'
                            outline_color = '#CCCCCC'
                            outline_width = 1
                    elif self.assigning_devices:
                        # Grey out all trackers while Assign Devices dialog is open
                        color = '#E0E0E0'
                        outline_color = '#CCCCCC'
                        outline_width = 1
                    else:
                        outline_color = '#555555'
                        outline_width = 1
                    
                    sy = ty + sum(string_heights[:s_idx])
                    sh = string_heights[s_idx] if s_idx < len(string_heights) else string_heights[-1]
                    
                    sx1, sy1 = self.world_to_canvas(tx + tx_offset, sy)
                    sx2, sy2 = self.world_to_canvas(tx + tx_offset + t_width, sy + sh)
                    
                    self.canvas.create_rectangle(
                        sx1, sy1, sx2, sy2,
                        fill=color, outline=outline_color, width=outline_width
                    )
                
                # Tracker outline
                ox1, oy1 = self.world_to_canvas(tx + tx_offset - 0.5, ty - 0.5)
                ox2, oy2 = self.world_to_canvas(tx + tx_offset + t_width + 0.5, ty + t_length + 0.5)
                
                self.canvas.create_rectangle(
                    ox1, oy1, ox2, oy2,
                    fill='', outline='#222222', width=1
                )
                
                # Motor indicator
                if tracker.get('has_motor', False):
                    motor_y = tracker['motor_y_ft']
                    motor_gap = tracker['motor_gap_ft']
                    
                    motor_world_y = ty + motor_y
                    motor_x1 = tx + tx_offset - 0.3
                    motor_x2 = tx + tx_offset + t_width + 0.3
                    
                    mx1, my1 = self.world_to_canvas(motor_x1, motor_world_y)
                    mx2, my2 = self.world_to_canvas(motor_x2, motor_world_y + motor_gap)
                    
                    self.canvas.create_rectangle(
                        mx1, my1, mx2, my2,
                        fill='#666666', outline='#444444', width=1
                    )
                    
                    motor_cx = (mx1 + mx2) / 2
                    motor_cy = (my1 + my2) / 2
                    dot_r = max(2, min(4, 3 * self.scale))
                    self.canvas.create_oval(
                        motor_cx - dot_r, motor_cy - dot_r,
                        motor_cx + dot_r, motor_cy + dot_r,
                        fill='#FF8800', outline='#CC6600', width=1
                    )
                
                # Tracker label
                label_cx, label_cy = self.world_to_canvas(
                    tx + tx_offset + t_width / 2,
                    ty + t_length + 2
                )
                pixel_width = abs(ox2 - ox1)
                if pixel_width > 14:
                    lbl_size = max(6, min(9, int(8 * self.scale)))
                    self.canvas.create_text(
                        label_cx, label_cy,
                        text=f"T{t_idx+1}", font=('Helvetica', lbl_size), fill='#555555'
                    )
        
        # Draw devices (CB/SI)
        self._draw_devices()

        # Draw routes (behind pads)
        self._draw_routes()
        
        # Draw pads
        self._draw_pads()
        
        # Motor alignment line (if groups share a motor Y and alignment is on)
        if getattr(self, 'align_on_motor', False):
            self._draw_motor_alignment_lines()
        
        # Overlap warnings
        self._draw_overlap_warnings()

        # Scale bar
        self._draw_scale_bar()
        
        # Compass
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
    
    def _draw_motor_alignment_lines(self):
        """Draw a driveline across each group at its motor Y position,
        following the driveline angle if set."""
        for group_data in self.group_layout:
            motor_y = group_data.get('motor_y_ft', 0)
            if motor_y <= 0:
                continue
            
            overhang = self.max_tracker_width_ft * 0.5
            driveline_tan = group_data.get('driveline_tan', 0.0)
            
            left_x = group_data['x'] - overhang
            right_x = group_data['x'] + group_data['width_ft'] + overhang
            
            left_y = group_data['y'] + motor_y
            # Angle offset based on horizontal span from group origin
            right_y = left_y + (right_x - group_data['x']) * driveline_tan
            left_y = left_y + (-overhang) * driveline_tan
            
            x1, y1 = self.world_to_canvas(left_x, left_y)
            x2, y2 = self.world_to_canvas(right_x, right_y)
            
            self.canvas.create_line(
                x1, y1, x2, y2,
                fill='#FF8800', width=2, dash=(6, 3)
            )

    def _on_alignment_toggle(self):
        """Handle motor alignment checkbox toggle."""
        self.align_on_motor = self.align_motor_var.get()
        self.draw()

    def _on_inspect_toggle(self):
        """Handle inspect mode toggle."""
        self.selected_device_idx = None
        self.selected_pad_inspect_idx = None
        if self.inspect_mode:
            self.selected_group_idx = None
        self.draw()

    def _draw_toggle(self):
        """Draw the slider toggle switch on the canvas."""
        self.toggle_canvas.delete('all')
        w, h = 52, 24
        r = h // 2  # radius for rounded ends
        
        if self.inspect_mode:
            # ON state — green track
            track_color = '#4CAF50'
            knob_x = w - r
        else:
            # OFF state — gray track
            track_color = '#BDBDBD'
            knob_x = r
        
        # Draw rounded track
        self.toggle_canvas.create_oval(0, 0, h, h, fill=track_color, outline=track_color)
        self.toggle_canvas.create_oval(w - h, 0, w, h, fill=track_color, outline=track_color)
        self.toggle_canvas.create_rectangle(r, 0, w - r, h, fill=track_color, outline=track_color)
        
        # Draw knob
        knob_r = r - 2
        self.toggle_canvas.create_oval(
            knob_x - knob_r, 2, knob_x + knob_r, h - 2,
            fill='white', outline='#999999', width=1
        )
    
    def _on_toggle_click(self, event=None):
        """Handle click on the toggle switch."""
        if self.inspect_mode:
            # Switching back to Layout — confirm
            from tkinter import messagebox
            if not messagebox.askyesno("Switch Mode",
                                        "Switch back to Layout mode? Groups will be draggable again.",
                                        parent=self):
                return
        
        self.inspect_mode = not self.inspect_mode
        self.inspect_mode_var.set(self.inspect_mode)
        self._draw_toggle()
        self.toggle_label.config(
            text="Inspect" if self.inspect_mode else "Layout",
            foreground='#4CAF50' if self.inspect_mode else '#333333'
        )
        self._on_inspect_toggle()
    
    def hit_test_device(self, cx, cy):
        """Return the index of the device under canvas coords (cx, cy), or None."""
        if not hasattr(self, 'device_positions') or not self.device_positions:
            return None
        wx, wy = self.canvas_to_world(cx, cy)
        for i, dev in enumerate(self.device_positions):
            if (dev['x'] <= wx <= dev['x'] + dev['width_ft'] and
                dev['y'] <= wy <= dev['y'] + dev['height_ft']):
                return i
        return None
    
    def _reset_positions(self):
        """Reset all group positions to auto-layout and clear saved positions."""
        from tkinter import messagebox
        if not messagebox.askyesno("Reset Positions",
                                    "This will reset all group positions to the default layout. Continue?"):
            return
        
        for grp_idx, layout in enumerate(self.group_layout):
            group_spacing = self.max_tracker_length_ft * 0.1
            layout['x'] = grp_idx * (layout['width_ft'] + group_spacing)
            layout['y'] = 0
        
        self._update_world_bounds()
        self._save_group_positions()
        self.fit_and_redraw()

    def _refresh_allocation(self):
        """Re-run string allocation using current group positions, then refresh preview."""
        self._save_group_positions()
        
        parent = self.master
        if hasattr(parent, 'calculate_estimate'):
            parent.calculate_estimate()
            
            inv_summary = getattr(parent, 'last_totals', {}).get('inverter_summary', {})
            if inv_summary and inv_summary.get('allocation_result'):
                self.inv_summary = inv_summary
                self.build_layout_data()
                self._build_legend()
                
                # Update top bar summary
                num_inv = inv_summary.get('total_inverters', 0)
                total_str = inv_summary.get('total_strings', 0)
                actual_ratio = inv_summary.get('actual_dc_ac', 0)
                split = inv_summary.get('total_split_trackers', 0)
                spatial_runs = inv_summary.get('allocation_result', {}).get('spatial_runs', 1)
                lock_str = "  |  🔒 LOCKED" if self.allocation_locked else ""
                self.summary_label.config(
                    text=f"{num_inv} Inverters  |  {total_str} Strings  |  DC:AC: {actual_ratio:.2f}  |  "
                         f"{split} Split Trackers  |  {spatial_runs} Run(s)  |  {self.topology}{lock_str}"
                )
                
                self.draw()

    def _toggle_allocation_lock(self):
        """Toggle the allocation lock on/off."""
        parent = self.master
        
        if self.allocation_locked:
            # Unlock
            self.allocation_locked = False
            parent.allocation_locked = False
            parent.locked_allocation_result = None
            self._update_lock_button()
        else:
            # Lock — snapshot the current allocation
            inv_summary = getattr(parent, 'last_totals', {}).get('inverter_summary', {})
            alloc = inv_summary.get('allocation_result')
            if not alloc:
                from tkinter import messagebox
                messagebox.showwarning(
                    "No Allocation",
                    "Run Refresh Allocation first to generate an allocation to lock.",
                    parent=self
                )
                return
            
            import copy
            self.allocation_locked = True
            parent.allocation_locked = True
            parent.locked_allocation_result = copy.deepcopy(alloc)
            self._update_lock_button()

    def _update_lock_button(self):
        """Update the lock button text and style to reflect current state."""
        if self.allocation_locked:
            self.lock_btn.config(text="🔒 Unlock Allocation")
        else:
            self.lock_btn.config(text="🔓 Lock Allocation")

    def _check_overlaps(self):
        """Check for overlapping groups and return list of overlapping pair indices."""
        overlaps = []
        
        for i in range(len(self.group_layout)):
            for j in range(i + 1, len(self.group_layout)):
                gi = self.group_layout[i]
                gj = self.group_layout[j]
                
                i_x1 = gi['x']
                i_x2 = gi['x'] + gi['width_ft']
                j_x1 = gj['x']
                j_x2 = gj['x'] + gj['width_ft']
                
                # Check X overlap first
                if i_x1 >= j_x2 or i_x2 <= j_x1:
                    continue
                
                # Driveline angle: Y bounds shift linearly with X
                i_tan = gi.get('driveline_tan', 0.0)
                j_tan = gj.get('driveline_tan', 0.0)
                i_vis_min = gi.get('visual_min_y_base', gi.get('visual_min_y', 0))
                i_vis_max = gi.get('visual_max_y_base', gi.get('visual_max_y', gi['length_ft']))
                j_vis_min = gj.get('visual_min_y_base', gj.get('visual_min_y', 0))
                j_vis_max = gj.get('visual_max_y_base', gj.get('visual_max_y', gj['length_ft']))
                
                # Check Y overlap at both ends of the X overlap region
                x_overlap_left = max(i_x1, j_x1)
                x_overlap_right = min(i_x2, j_x2)
                
                has_overlap = False
                for x_check in [x_overlap_left, x_overlap_right]:
                    i_y1 = gi['y'] + i_vis_min + (x_check - i_x1) * i_tan
                    i_y2 = gi['y'] + i_vis_max + (x_check - i_x1) * i_tan
                    j_y1 = gj['y'] + j_vis_min + (x_check - j_x1) * j_tan
                    j_y2 = gj['y'] + j_vis_max + (x_check - j_x1) * j_tan
                    
                    if i_y1 < j_y2 and i_y2 > j_y1:
                        has_overlap = True
                        break
                
                if has_overlap:
                    overlaps.append((i, j))
        
        return overlaps
    
    def _draw_overlap_warnings(self):
        """Draw red warning highlights around overlapping groups."""
        overlaps = self._check_overlaps()
        if not overlaps:
            return
        
        # Collect unique group indices that are involved in overlaps
        overlap_indices = set()
        for i, j in overlaps:
            overlap_indices.add(i)
            overlap_indices.add(j)
        
        max_width = getattr(self, 'max_tracker_width_ft', 6)
        pad = max_width * 0.3
        
        for idx in overlap_indices:
            g = self.group_layout[idx]
            vis_min = g.get('visual_min_y', 0)
            vis_max = g.get('visual_max_y', g['length_ft'])
            
            hx1, hy1 = self.world_to_canvas(g['x'] - pad, g['y'] + vis_min - pad)
            hx2, hy2 = self.world_to_canvas(
                g['x'] + g['width_ft'] + pad,
                g['y'] + vis_max + pad
            )
            
            self.canvas.create_rectangle(
                hx1, hy1, hx2, hy2,
                fill='', outline='#FF0000', width=3, dash=(8, 4),
                tags='overlap_warning'
            )
            
            # Warning label
            label_x = (hx1 + hx2) / 2
            label_y = hy1 - 8
            font_size = max(7, min(10, int(9 * self.scale)))
            self.canvas.create_text(
                label_x, label_y,
                text=f"⚠ Overlap", font=('Helvetica', font_size, 'bold'),
                fill='#FF0000', tags='overlap_warning'
            )

    def _draw_devices(self):
        """Draw combiner box / string inverter rectangles on the canvas."""
        if not hasattr(self, 'device_positions') or not self.device_positions:
            return
        
        # Build device_idx -> pad_idx lookup
        device_to_pad = {}
        if self.pads:
            for pad_idx, pad in enumerate(self.pads):
                for dev_idx in pad.get('assigned_devices', []):
                    device_to_pad[dev_idx] = pad_idx
        
        # Pad colors for device outlines (when pads exist)
        PAD_COLORS = ['#C62828', '#1565C0', '#2E7D32', '#E65100', '#6A1B9A',
                      '#00838F', '#AD1457', '#4E342E']
        
        for dev_idx, dev in enumerate(self.device_positions):
            dx = dev['x']
            dy = dev['y']
            dw = dev['width_ft']
            dh = dev['height_ft']
            label = dev['label']
            is_selected = (self.selected_device_idx == dev_idx)
            
            x1, y1 = self.world_to_canvas(dx, dy)
            x2, y2 = self.world_to_canvas(dx + dw, dy + dh)
            
            # Device fill color
            if self.device_label == 'CB':
                fill_color = '#FFB74D' if is_selected else '#FF9800'
            else:
                fill_color = '#64B5F6' if is_selected else '#2196F3'
            
            # Outline color: pad color if assigned, else default
            if dev_idx in device_to_pad and self.pads:
                pad_idx = device_to_pad[dev_idx]
                outline_color = PAD_COLORS[pad_idx % len(PAD_COLORS)]
                outline_width = 3
            else:
                outline_color = '#E65100' if self.device_label == 'CB' else '#0D47A1'
                outline_width = 3 if is_selected else 2
            
            if is_selected:
                outline_width = 4
            
            self.canvas.create_rectangle(
                x1, y1, x2, y2,
                fill=fill_color, outline=outline_color, width=outline_width
            )
            
            # Label — offset above device for readability
            cx = (x1 + x2) / 2
            font_size = max(7, min(14, int(10 * self.scale)))
            label_y = y1 - font_size - 2
            self.canvas.create_text(
                cx, label_y,
                text=label, font=('Helvetica', font_size, 'bold'),
                fill='#333333', anchor='s'
            )

    def _draw_scale_bar(self):
        """Draw a scale bar in the bottom-left corner showing real-world distance."""
        if self.scale <= 0:
            return
        
        self.canvas.update_idletasks()
        ch = self.canvas.winfo_height()
        
        # Pick a nice round scale bar length in feet
        target_px = 120  # target pixel width for the bar
        target_ft = target_px / self.scale
        
        # Round to a nice number
        nice_values = [5, 10, 20, 25, 50, 100, 200, 250, 500, 1000]
        bar_ft = nice_values[0]
        for v in nice_values:
            if v <= target_ft:
                bar_ft = v
            else:
                break
        
        bar_px = bar_ft * self.scale
        
        x1 = 20
        y1 = ch - 25
        x2 = x1 + bar_px
        
        self.canvas.create_line(x1, y1, x2, y1, fill='#333333', width=2)
        self.canvas.create_line(x1, y1 - 5, x1, y1 + 5, fill='#333333', width=2)
        self.canvas.create_line(x2, y1 - 5, x2, y1 + 5, fill='#333333', width=2)
        
        self.canvas.create_text(
            (x1 + x2) / 2, y1 - 10,
            text=f"{bar_ft} ft", font=('Helvetica', 9), fill='#333333'
        )

    def _add_pad(self):
        """Enter pad placement mode — next click on canvas places a new pad."""
        if self.inspect_mode:
            from tkinter import messagebox
            messagebox.showinfo("Locked", "Switch to Layout mode to add pads.", parent=self)
            return
        
        self.placing_pad = True
        self.canvas.config(cursor='crosshair')
        self.add_pad_btn.config(state='disabled')
    
    def _place_pad_at(self, wx, wy):
        """Create a new pad at the given world coordinates."""
        pad_num = len(self.pads) + 1
        # Auto-assign all devices to the first pad if it's the only one
        if len(self.pads) == 0:
            all_device_indices = list(range(len(self.device_positions)))
        else:
            all_device_indices = []
        
        label_char = chr(ord('A') + (pad_num - 1) % 26)
        
        self.pads.append({
            'label': f"Pad {label_char}",
            'x': wx - 5,  # Center the 10ft-wide pad on click
            'y': wy - 4,  # Center the 8ft-tall pad on click
            'width_ft': 10.0,
            'height_ft': 8.0,
            'assigned_devices': all_device_indices,
        })
        
        self.placing_pad = False
        self.canvas.config(cursor='')
        self.add_pad_btn.config(state='normal')
        self.draw()
    
    def _draw_pads(self):
        """Draw inverter pad rectangles on the canvas."""
        if not self.pads:
            return
        
        PAD_COLORS = ['#C62828', '#1565C0', '#2E7D32', '#E65100', '#6A1B9A',
                      '#00838F', '#AD1457', '#4E342E']
        
        for pad_idx, pad in enumerate(self.pads):
            px = pad['x']
            py = pad['y']
            pw = pad.get('width_ft', 10.0)
            ph = pad.get('height_ft', 8.0)
            label = pad.get('label', f'Pad {pad_idx+1}')
            is_selected = (self.selected_pad_idx == pad_idx)
            
            x1, y1 = self.world_to_canvas(px, py)
            x2, y2 = self.world_to_canvas(px + pw, py + ph)
            
            base_color = PAD_COLORS[pad_idx % len(PAD_COLORS)]
            outline_width = 3 if is_selected else 2
            
            self.canvas.create_rectangle(
                x1, y1, x2, y2,
                fill=base_color, outline='#222222', width=outline_width
            )
            
            # Label
            cx = (x1 + x2) / 2
            cy = (y1 + y2) / 2
            font_size = max(6, min(10, int(8 * self.scale)))
            self.canvas.create_text(
                cx, cy,
                text=label, font=('Helvetica', font_size, 'bold'),
                fill='white'
            )
            
            # Device count subtitle
            num_assigned = len(pad.get('assigned_devices', []))
            if num_assigned > 0:
                sub_size = max(5, min(8, int(6 * self.scale)))
                self.canvas.create_text(
                    cx, cy + font_size + 2,
                    text=f"({num_assigned} devices)", font=('Helvetica', sub_size),
                    fill='#CCCCCC'
                )
    
    def hit_test_pad(self, cx, cy):
        """Return the index of the pad under canvas coords, or None."""
        if not self.pads:
            return None
        wx, wy = self.canvas_to_world(cx, cy)
        for i, pad in enumerate(self.pads):
            pw = pad.get('width_ft', 10.0)
            ph = pad.get('height_ft', 8.0)
            if (pad['x'] <= wx <= pad['x'] + pw and
                pad['y'] <= wy <= pad['y'] + ph):
                return i
        return None

    def hit_test_device(self, cx, cy):
        """Return the index of the device under canvas coords, or None."""
        if not hasattr(self, 'device_positions') or not self.device_positions:
            return None
        wx, wy = self.canvas_to_world(cx, cy)
        for i, dev in enumerate(self.device_positions):
            if (dev['x'] <= wx <= dev['x'] + dev['width_ft'] and
                dev['y'] <= wy <= dev['y'] + dev['height_ft']):
                return i
        return None
    
    def _show_assignment_dialog(self):
        """Show a dialog to assign devices to pads."""
        if not self.pads:
            from tkinter import messagebox
            messagebox.showinfo("No Pads", "Add at least one pad first.", parent=self)
            return
        
        if not hasattr(self, 'device_positions') or not self.device_positions:
            from tkinter import messagebox
            messagebox.showinfo("No Devices", "No devices to assign. Run Calculate Estimate first.", parent=self)
            return
        
        dialog = tk.Toplevel(self)
        dialog.title("Assign Devices to Pads")
        dialog.transient(self)
        # No grab_set() — allow pan/zoom on canvas behind dialog
        
        self.assigning_devices = True
        self.draw()
        
        # Size based on device count
        num_devices = len(self.device_positions)
        dialog_height = min(600, 120 + num_devices * 28)
        dialog.geometry(f"500x{dialog_height}")
        dialog.minsize(400, 200)
        

        # Instructions
        ttk.Label(dialog, text="Assign each device to a collection pad:",
                  font=('Helvetica', 10)).pack(anchor='w', padx=10, pady=(10, 0))
        ttk.Label(dialog, text="Tip: Drag the blue handle to fill multiple rows with the same pad",
                  font=('Helvetica', 8), foreground='gray').pack(anchor='w', padx=10, pady=(0, 5))
        
        # Build pad label list for dropdowns
        pad_labels = [pad.get('label', f'Pad {i+1}') for i, pad in enumerate(self.pads)]
        
        # Build reverse lookup: device_idx -> pad_idx
        device_to_pad = {}
        for pad_idx, pad in enumerate(self.pads):
            for dev_idx in pad.get('assigned_devices', []):
                device_to_pad[dev_idx] = pad_idx
        
        # Scrollable frame
        container = ttk.Frame(dialog)
        container.pack(fill='both', expand=True, padx=10, pady=5)
        
        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient='vertical', command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        
        scroll_frame.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=scroll_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')
        canvas.bind('<Enter>', lambda e: canvas.bind_all('<MouseWheel>', _on_mousewheel))
        canvas.bind('<Leave>', lambda e: canvas.unbind_all('<MouseWheel>'))
        
        # Headers
        header = ttk.Frame(scroll_frame)
        header.pack(fill='x', pady=(0, 5))
        ttk.Label(header, text="Device", font=('Helvetica', 9, 'bold'), width=12).pack(side='left', padx=5)
        ttk.Label(header, text="Strings", font=('Helvetica', 9, 'bold'), width=8).pack(side='left', padx=5)
        ttk.Label(header, text="Group", font=('Helvetica', 9, 'bold'), width=12).pack(side='left', padx=5)
        ttk.Label(header, text="Pad", font=('Helvetica', 9, 'bold'), width=12).pack(side='left', padx=5)
        ttk.Label(header, text="", width=1).pack(side='left', padx=2)
        
        ttk.Separator(scroll_frame, orient='horizontal').pack(fill='x', pady=2)
        
        # One row per device
        pad_vars = []
        pad_combos = []
        
        # Drag-fill handle state
        _fill = {'active': False, 'source_idx': None, 'value': None}
        _fill_handles = []
        
        def _handle_press(event, idx):
            """Start fill-drag from this row's handle"""
            _fill['active'] = True
            _fill['source_idx'] = idx
            _fill['value'] = pad_vars[idx].get()
            event.widget.config(cursor='sb_v_double_arrow')
        
        def _handle_motion(event):
            """Fill combos as drag passes over handles"""
            if not _fill['active']:
                return
            my = event.y_root
            src = _fill['source_idx']
            for i, handle in enumerate(_fill_handles):
                try:
                    hy = handle.winfo_rooty()
                    hh = handle.winfo_height()
                    row_center = hy + hh / 2
                    src_center = _fill_handles[src].winfo_rooty() + _fill_handles[src].winfo_height() / 2
                    in_range = min(my, src_center) <= row_center <= max(my, src_center)
                    if in_range:
                        pad_vars[i].set(_fill['value'])
                        pad_combos[i].focus_set()
                except tk.TclError:
                    pass
        
        def _handle_release(event):
            """End fill-drag"""
            if _fill['active']:
                event.widget.config(cursor='')
            _fill['active'] = False
            _fill['source_idx'] = None
            _fill['value'] = None

        for dev_idx, dev in enumerate(self.device_positions):
            row = ttk.Frame(scroll_frame)
            row.pack(fill='x', pady=1)
            
            # Device label
            ttk.Label(row, text=dev['label'], width=12).pack(side='left', padx=5)
            
            # String count
            num_strings = sum(len(v) for v in dev.get('assigned_strings', {}).values())
            ttk.Label(row, text=str(num_strings), width=8).pack(side='left', padx=5)
            
            # Group name
            grp_idx = dev.get('group_idx', 0)
            grp_name = self.groups[grp_idx]['name'] if grp_idx < len(self.groups) else '?'
            ttk.Label(row, text=grp_name, width=12).pack(side='left', padx=5)
            
            # Pad dropdown
            current_pad = device_to_pad.get(dev_idx, 0)
            if current_pad >= len(pad_labels):
                current_pad = 0
            var = tk.StringVar(value=pad_labels[current_pad])
            combo = ttk.Combobox(row, textvariable=var, values=pad_labels,
                                 state='readonly', width=12)
            combo.pack(side='left', padx=5)
            pad_vars.append(var)
            pad_combos.append(combo)
            
            # Fill handle — small draggable square
            handle = tk.Frame(row, width=10, height=18, bg='#4A90D9', cursor='sb_v_double_arrow')
            handle.pack(side='left', padx=(2, 0))
            handle.pack_propagate(False)
            handle.bind('<ButtonPress-1>', lambda e, i=dev_idx: _handle_press(e, i))
            handle.bind('<B1-Motion>', _handle_motion)
            handle.bind('<ButtonRelease-1>', _handle_release)
            _fill_handles.append(handle)
        
        # Buttons
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill='x', padx=10, pady=10)
        
        def _assign_all_to(pad_label):
            for var in pad_vars:
                var.set(pad_label)
        
        if len(self.pads) > 1:
            ttk.Label(btn_frame, text="Quick assign all:").pack(side='left', padx=(0, 5))
            for label in pad_labels:
                ttk.Button(btn_frame, text=label,
                           command=lambda l=label: _assign_all_to(l)).pack(side='left', padx=2)
        
        def _apply():
            # Rebuild pad assignments from dropdown values
            for pad in self.pads:
                pad['assigned_devices'] = []
            
            for dev_idx, var in enumerate(pad_vars):
                selected_label = var.get()
                for pad_idx, pad in enumerate(self.pads):
                    if pad.get('label', f'Pad {pad_idx+1}') == selected_label:
                        pad['assigned_devices'].append(dev_idx)
                        break
            
            self.assigning_devices = False
            self.draw()
            dialog.destroy()
        
        def _cancel():
            self.assigning_devices = False
            self.draw()
            dialog.destroy()
        
        dialog.protocol("WM_DELETE_WINDOW", _cancel)
        
        ttk.Button(btn_frame, text="Apply", command=_apply).pack(side='right', padx=(5, 0))
        ttk.Button(btn_frame, text="Cancel", command=_cancel).pack(side='right')
        
        # Center on parent
        dialog.update_idletasks()
        px = self.winfo_rootx()
        py = self.winfo_rooty()
        pw = self.winfo_width()
        ph = self.winfo_height()
        dw = dialog.winfo_width()
        dh = dialog.winfo_height()
        dialog.geometry(f"+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")

    def _show_cb_assignment_dialog(self):
        """Show dialog to reassign harness connections between combiner boxes."""
        if not hasattr(self, 'device_positions') or not self.device_positions:
            from tkinter import messagebox
            messagebox.showinfo("No Devices", "No combiner boxes to edit.", parent=self)
            return
        
        # Build working data from last_combiner_assignments (accurate per-harness)
        parent_qe = self.master
        assignments = getattr(parent_qe, 'last_combiner_assignments', [])
        
        if not assignments:
            from tkinter import messagebox
            messagebox.showinfo("No Data", "No combiner assignments found. Run Calculate Estimate first.", parent=self)
            return
        
        dialog = tk.Toplevel(self)
        dialog.title("Edit Combiner Box Assignments")
        dialog.geometry("750x500")
        dialog.transient(self)
        dialog.grab_set()
        
        # Deep copy so we can cancel without side effects
        cb_data = []
        for cb in assignments:
            cb_data.append({
                'name': cb['combiner_name'],
                'dev_idx': cb.get('device_idx'),
                'module_isc': cb.get('module_isc', 0),
                'nec_factor': cb.get('nec_factor', 1.56),
                'connections': [dict(c) for c in cb['connections']],
            })
        
        # Save original state for cancel revert
        import copy
        original_cb_data = copy.deepcopy(cb_data)
        
        # --- Layout ---
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill='both', expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=0)
        main_frame.columnconfigure(2, weight=2)
        main_frame.rowconfigure(0, weight=1)
        
        # --- Left panel: CB list ---
        left_frame = ttk.LabelFrame(main_frame, text="Combiner Boxes", padding="5")
        left_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 5))
        left_frame.rowconfigure(0, weight=1)
        left_frame.columnconfigure(0, weight=1)
        
        cb_listbox = tk.Listbox(left_frame, exportselection=False)
        cb_listbox.grid(row=0, column=0, sticky='nsew')
        
        cb_scroll = ttk.Scrollbar(left_frame, orient='vertical', command=cb_listbox.yview)
        cb_scroll.grid(row=0, column=1, sticky='ns')
        cb_listbox.config(yscrollcommand=cb_scroll.set)
        
        def refresh_cb_list(select_idx=None):
            cb_listbox.delete(0, tk.END)
            for i, cb in enumerate(cb_data):
                n_conn = len(cb['connections'])
                total_str = sum(c['num_strings'] for c in cb['connections'])
                cb_listbox.insert(tk.END, f"{cb['name']}  ({n_conn} inputs, {total_str} strings)")
            if select_idx is not None and select_idx < len(cb_data):
                cb_listbox.selection_set(select_idx)
                cb_listbox.see(select_idx)
                on_cb_select(None)
        
        def _update_live_preview():
            """Push current cb_data to device_positions and tracker colors, then redraw."""
            # Rebuild device_names and assigned_strings on device_positions
            for dev in self.device_positions:
                dev['assigned_strings'] = {}
            
            for cb_idx, cb in enumerate(cb_data):
                if cb_idx < len(self.device_positions):
                    self.device_positions[cb_idx]['label'] = cb['name']
                    self.device_names[cb_idx] = cb['name']
                    
                    assigned = {}
                    for conn in cb['connections']:
                        tidx = conn['tracker_idx']
                        n_str = conn['num_strings']
                        if tidx not in assigned:
                            assigned[tidx] = set()
                        existing = len(assigned[tidx])
                        for s in range(existing, existing + n_str):
                            assigned[tidx].add(s)
                    self.device_positions[cb_idx]['assigned_strings'] = assigned
            
            # Rebuild tracker color assignments from cb_data
            # Build tracker_idx -> list of (cb_idx, strings_taken)
            tracker_cb_map = {}  # tidx -> [(cb_idx, strings_taken), ...]
            for cb_idx, cb in enumerate(cb_data):
                for conn in cb['connections']:
                    tidx = conn['tracker_idx']
                    if tidx not in tracker_cb_map:
                        tracker_cb_map[tidx] = []
                    tracker_cb_map[tidx].append((cb_idx, conn['num_strings']))
            
            # Walk through group_layout trackers and update their assignments
            global_idx = 0
            for group_data in self.group_layout:
                for tracker in group_data['trackers']:
                    if global_idx in tracker_cb_map:
                        new_assignments = []
                        for cb_idx, strings_taken in tracker_cb_map[global_idx]:
                            color = self.colors[cb_idx % len(self.colors)]
                            new_assignments.append({
                                'color': color,
                                'strings': strings_taken,
                                'inv_idx': cb_idx,
                            })
                        tracker['assignments'] = new_assignments
                    else:
                        # Unassigned tracker — grey
                        tracker['assignments'] = [{
                            'color': '#CCCCCC',
                            'strings': tracker['strings_per_tracker'],
                            'inv_idx': -1,
                        }]
                    global_idx += 1
            
            self.draw()
        
        # CB buttons
        cb_btn_frame = ttk.Frame(left_frame)
        cb_btn_frame.grid(row=1, column=0, columnspan=2, pady=(5, 0), sticky='ew')
        
        def add_cb():
            next_num = len(cb_data) + 1
            module_isc = cb_data[0]['module_isc'] if cb_data else 0
            nec_factor = cb_data[0]['nec_factor'] if cb_data else 1.56
            cb_data.append({
                'name': f"CB-{next_num:02d}",
                'dev_idx': None,
                'module_isc': module_isc,
                'nec_factor': nec_factor,
                'connections': [],
            })
            refresh_cb_list(select_idx=len(cb_data) - 1)
            _update_live_preview()
        
        def remove_cb():
            sel = cb_listbox.curselection()
            if not sel:
                return
            idx = sel[0]
            cb = cb_data[idx]
            if cb['connections']:
                from tkinter import messagebox
                if not messagebox.askyesno(
                    "Remove CB",
                    f"'{cb['name']}' has {len(cb['connections'])} connections.\n"
                    "They will become unassigned. Continue?",
                    parent=dialog
                ):
                    return
            cb_data.pop(idx)
            new_sel = min(idx, len(cb_data) - 1) if cb_data else None
            refresh_cb_list(select_idx=new_sel)
            _update_live_preview()
        
        def rename_cb():
            sel = cb_listbox.curselection()
            if not sel:
                return
            idx = sel[0]
            from tkinter import simpledialog
            new_name = simpledialog.askstring(
                "Rename CB", f"New name for {cb_data[idx]['name']}:",
                initialvalue=cb_data[idx]['name'], parent=dialog
            )
            if new_name and new_name.strip():
                cb_data[idx]['name'] = new_name.strip()
                refresh_cb_list(select_idx=idx)
                _update_live_preview()
        
        ttk.Button(cb_btn_frame, text="Add CB", command=add_cb).pack(side='left', padx=2)
        ttk.Button(cb_btn_frame, text="Remove", command=remove_cb).pack(side='left', padx=2)
        ttk.Button(cb_btn_frame, text="Rename", command=rename_cb).pack(side='left', padx=2)
        
        # --- Middle: Move button ---
        mid_frame = ttk.Frame(main_frame)
        mid_frame.grid(row=0, column=1, padx=5, sticky='ns')
        mid_frame.rowconfigure(0, weight=1)
        mid_frame.rowconfigure(2, weight=1)
        
        def move_selected():
            cb_sel = cb_listbox.curselection()
            if not cb_sel:
                from tkinter import messagebox
                messagebox.showinfo("Select CB", "Select a source CB on the left first.", parent=dialog)
                return
            src_idx = cb_sel[0]
            
            conn_sel = conn_tree.selection()
            if not conn_sel:
                from tkinter import messagebox
                messagebox.showinfo("Select Connections", "Select connections to move on the right.", parent=dialog)
                return
            
            target_names = [(i, cb['name']) for i, cb in enumerate(cb_data) if i != src_idx]
            if not target_names:
                return
            
            pick = tk.Toplevel(dialog)
            pick.title("Move to...")
            pick.transient(dialog)
            pick.grab_set()
            pick.geometry("250x300")
            
            ttk.Label(pick, text="Select target CB:").pack(padx=10, pady=(10, 5))
            target_lb = tk.Listbox(pick, exportselection=False)
            target_lb.pack(fill='both', expand=True, padx=10)
            for i, name in target_names:
                target_lb.insert(tk.END, name)
            if target_names:
                target_lb.selection_set(0)
            
            def do_move():
                t_sel = target_lb.curselection()
                if not t_sel:
                    pick.destroy()
                    return
                target_idx = target_names[t_sel[0]][0]
                
                selected_iids = conn_tree.selection()
                all_iids = list(conn_tree.get_children())
                indices_to_move = [all_iids.index(iid) for iid in selected_iids if iid in all_iids]
                
                moved = []
                for ci in sorted(indices_to_move, reverse=True):
                    if ci < len(cb_data[src_idx]['connections']):
                        moved.append(cb_data[src_idx]['connections'].pop(ci))
                
                cb_data[target_idx]['connections'].extend(reversed(moved))
                
                pick.destroy()
                refresh_cb_list(select_idx=src_idx)
                _update_live_preview()
            
            ttk.Button(pick, text="Move", command=do_move).pack(pady=10)
        
        ttk.Button(mid_frame, text="Move →", command=move_selected).grid(row=1, column=0)
        
        # --- Right panel: Connections for selected CB ---
        right_frame = ttk.LabelFrame(main_frame, text="Connections", padding="5")
        right_frame.grid(row=0, column=2, sticky='nsew', padx=(5, 0))
        right_frame.rowconfigure(0, weight=1)
        right_frame.columnconfigure(0, weight=1)
        
        conn_columns = ('tracker', 'harness', 'strings')
        conn_tree = ttk.Treeview(right_frame, columns=conn_columns, show='headings', selectmode='extended')
        conn_tree.heading('tracker', text='Tracker')
        conn_tree.heading('harness', text='Harness')
        conn_tree.heading('strings', text='# Strings')
        conn_tree.column('tracker', width=80, anchor='center')
        conn_tree.column('harness', width=80, anchor='center')
        conn_tree.column('strings', width=80, anchor='center')
        conn_tree.grid(row=0, column=0, sticky='nsew')
        
        conn_scroll = ttk.Scrollbar(right_frame, orient='vertical', command=conn_tree.yview)
        conn_scroll.grid(row=0, column=1, sticky='ns')
        conn_tree.config(yscrollcommand=conn_scroll.set)
        
        def on_cb_select(event):
            sel = cb_listbox.curselection()
            if not sel:
                conn_tree.delete(*conn_tree.get_children())
                return
            idx = sel[0]
            conn_tree.delete(*conn_tree.get_children())
            for conn in cb_data[idx]['connections']:
                conn_tree.insert('', 'end', values=(
                    conn['tracker_label'],
                    conn['harness_label'],
                    conn['num_strings']
                ))
        
        cb_listbox.bind('<<ListboxSelect>>', on_cb_select)
        
        # --- Warning label ---
        warning_var = tk.StringVar(value="")
        warning_label = ttk.Label(dialog, textvariable=warning_var, foreground='orange')
        warning_label.pack(fill='x', padx=10)
        
        # --- Bottom buttons ---
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill='x', padx=10, pady=10)
        
        def apply_changes():
            # Write back to parent's last_combiner_assignments
            BREAKER_SIZES = [100, 125, 150, 175, 200, 225, 250, 300, 350, 400, 450, 500, 600, 700, 800]
            
            new_assignments = []
            for cb_idx, cb in enumerate(cb_data):
                connections = cb['connections']
                total_current = sum(
                    c['num_strings'] * cb['module_isc'] * cb['nec_factor']
                    for c in connections
                )
                calc_breaker = 400
                for bs in BREAKER_SIZES:
                    if bs >= total_current:
                        calc_breaker = bs
                        break
                
                new_assignments.append({
                    'combiner_name': cb['name'],
                    'device_idx': cb_idx,
                    'breaker_size': calc_breaker,
                    'module_isc': cb['module_isc'],
                    'nec_factor': cb['nec_factor'],
                    'connections': connections,
                })
            
            parent_qe.last_combiner_assignments = new_assignments
            
            # Final preview update
            _update_live_preview()
            
            dialog.destroy()
        
        def cancel():
            # Restore from saved snapshot
            for cb_idx, snap in enumerate(original_cb_data):
                if cb_idx < len(cb_data):
                    cb_data[cb_idx] = snap
                else:
                    cb_data.append(snap)
            # Trim if we added extra CBs
            while len(cb_data) > len(original_cb_data):
                cb_data.pop()
            _update_live_preview()
            dialog.destroy()
        
        dialog.protocol("WM_DELETE_WINDOW", cancel)
        
        ttk.Button(btn_frame, text="Apply", command=apply_changes).pack(side='right', padx=(5, 0))
        ttk.Button(btn_frame, text="Cancel", command=cancel).pack(side='right')
        
        # Initial population
        refresh_cb_list(select_idx=0)
        
        # Center on parent
        dialog.update_idletasks()
        px = self.winfo_rootx()
        py = self.winfo_rooty()
        pw = self.winfo_width()
        ph = self.winfo_height()
        dw = dialog.winfo_width()
        dh = dialog.winfo_height()
        dialog.geometry(f"+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")

    def _on_pad_right_click(self, event):
        """Show context menu for pads."""
        hit = self.hit_test_pad(event.x, event.y)
        if hit is None:
            return
        
        self.selected_pad_idx = hit
        self.draw()
        
        menu = tk.Menu(self, tearoff=0)
        
        def _rename():
            from tkinter import simpledialog
            current = self.pads[hit].get('label', f'Pad {hit+1}')
            new_name = simpledialog.askstring("Rename Pad", "New label:", 
                                              initialvalue=current, parent=self)
            if new_name and new_name.strip():
                self.pads[hit]['label'] = new_name.strip()
                self.draw()
        
        def _delete():
            from tkinter import messagebox
            label = self.pads[hit].get('label', f'Pad {hit+1}')
            if not messagebox.askyesno("Delete Pad", f"Delete '{label}'?", parent=self):
                return
            
            # Reassign devices to first remaining pad if any
            orphaned = self.pads[hit].get('assigned_devices', [])
            del self.pads[hit]
            
            if self.pads and orphaned:
                self.pads[0]['assigned_devices'] = list(
                    set(self.pads[0].get('assigned_devices', []) + orphaned)
                )
            
            # Fix device indices in remaining pads (indices > hit shift down)
            # Not needed — pad indices don't change, only the list position
            
            self.selected_pad_idx = None
            self.draw()
        
        menu.add_command(label="Rename", command=_rename)
        menu.add_command(label="Delete", command=_delete)
        menu.tk_popup(event.x_root, event.y_root)

    def _draw_routes(self):
        """Draw L-shaped Manhattan routes from each device to its assigned pad."""
        if not self.show_routes_var.get():
            return
        if not self.pads or not hasattr(self, 'device_positions') or not self.device_positions:
            return
        
        PAD_COLORS = ['#C62828', '#1565C0', '#2E7D32', '#E65100', '#6A1B9A',
                      '#00838F', '#AD1457', '#4E342E']
        
        # Build device -> pad lookup
        device_to_pad = {}
        for pad_idx, pad in enumerate(self.pads):
            for dev_idx in pad.get('assigned_devices', []):
                device_to_pad[dev_idx] = pad_idx
        
        for dev_idx, dev in enumerate(self.device_positions):
            pad_idx = device_to_pad.get(dev_idx)
            if pad_idx is None or pad_idx >= len(self.pads):
                continue
            
            pad = self.pads[pad_idx]
            
            # Device center
            dev_cx = dev['x'] + dev['width_ft'] / 2
            dev_cy = dev['y'] + dev['height_ft'] / 2
            
            # Pad center
            pad_cx = pad['x'] + pad.get('width_ft', 10.0) / 2
            pad_cy = pad['y'] + pad.get('height_ft', 8.0) / 2
            
            # L-shaped route: go E-W first, then N-S
            corner_x = pad_cx
            corner_y = dev_cy
            
            # Convert to canvas coords
            cx1, cy1 = self.world_to_canvas(dev_cx, dev_cy)
            cx_corner, cy_corner = self.world_to_canvas(corner_x, corner_y)
            cx2, cy2 = self.world_to_canvas(pad_cx, pad_cy)
            
            color = PAD_COLORS[pad_idx % len(PAD_COLORS)]
            
            # Determine line style based on topology
            if self.topology == 'Distributed String':
                dash_pattern = (4, 4)  # Dashed for AC
            else:
                dash_pattern = ()  # Solid for DC
            
            line_width = 1
            
            # Draw E-W leg
            self.canvas.create_line(
                cx1, cy1, cx_corner, cy_corner,
                fill=color, width=line_width, dash=dash_pattern
            )
            
            # Draw N-S leg
            self.canvas.create_line(
                cx_corner, cy_corner, cx2, cy2,
                fill=color, width=line_width, dash=dash_pattern
            )

            # Show distance label if this device is selected in inspect mode
            if self.inspect_mode and self.selected_device_idx == dev_idx:
                ew_dist = abs(dev_cx - pad_cx)
                ns_dist = abs(dev_cy - pad_cy)
                total_dist = ew_dist + ns_dist  # Manhattan
                
                # Place label at the corner of the L
                font_size = max(6, min(10, int(8 * self.scale)))
                self.canvas.create_text(
                    cx_corner, cy_corner - 8,
                    text=f"{total_dist:.0f} ft",
                    font=('Helvetica', font_size, 'bold'),
                    fill=color
                )

    def _recolor_from_cb_assignments(self):
        """Recolor tracker assignments from parent's last_combiner_assignments.
        
        Called after build_layout_data() to override the default inverter-based
        coloring with CB-based coloring when manual CB edits exist.
        """
        parent_qe = self.master
        assignments = getattr(parent_qe, 'last_combiner_assignments', [])
        if not assignments:
            return
        
        # Build tracker_idx -> [(cb_idx, strings_taken), ...]
        tracker_cb_map = {}
        for cb_idx, cb in enumerate(assignments):
            for conn in cb.get('connections', []):
                tidx = conn['tracker_idx']
                if tidx not in tracker_cb_map:
                    tracker_cb_map[tidx] = []
                tracker_cb_map[tidx].append((cb_idx, conn['num_strings']))
        
        if not tracker_cb_map:
            return
        
        # Walk through group_layout and update tracker assignments
        global_idx = 0
        for group_data in self.group_layout:
            for tracker in group_data['trackers']:
                if global_idx in tracker_cb_map:
                    new_assignments = []
                    for cb_idx, strings_taken in tracker_cb_map[global_idx]:
                        color = self.colors[cb_idx % len(self.colors)]
                        new_assignments.append({
                            'color': color,
                            'strings': strings_taken,
                            'inv_idx': cb_idx,
                        })
                    tracker['assignments'] = new_assignments
                global_idx += 1