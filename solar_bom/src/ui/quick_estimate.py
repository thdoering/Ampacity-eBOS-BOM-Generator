import tkinter as tk
from tkinter import ttk
from typing import Optional, Dict, List, Any
from pathlib import Path
import json
import uuid
import math
from datetime import datetime
from src.utils.string_allocation import allocate_strings, allocate_strings_sequential


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
        self.selected_group_idx = None
        self._updating_listbox = False
        self.enabled_templates = self.load_enabled_templates()

        # Global settings defaults
        self.module_width_default = 1134
        self.modules_per_string_default = 28
        self.row_spacing_default = 20.0
        self.wire_gauge_default = '10 AWG'
        
        # Track currently selected item
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
            
            total_modules = int(spt) * int(mps) if isinstance(spt, int) and isinstance(mps, int) else '?'
            tracker_length_m = (int(mps) * int(spt) * module_dim_along_tracker + 
                               (int(mps) * int(spt) - 1) * spacing_m +
                               (motor_gap_m if has_motor else 0)) if isinstance(spt, int) and isinstance(mps, int) else None
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
        
        # Default to first enabled template if available
        enabled_keys = list(self.enabled_templates.keys())
        if enabled_keys:
            default_ref = enabled_keys[0]
            default_spt = self.enabled_templates[default_ref].get('strings_per_tracker', 3)
        else:
            default_ref = None
            default_spt = 3
        
        group = {
            'name': f"Group {group_num}",
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
        
        self._mark_stale()
        self._schedule_autosave()
        return idx
    
    def copy_selected_group(self):
        """Copy the currently selected group"""
        sel = self.group_listbox.curselection()
        if not sel:
            return
        
        import copy
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
            total_strings = sum(seg['quantity'] * seg['strings_per_tracker'] for seg in group['segments'])
            display = f"{group['name']}  ({total_trackers}T / {total_strings}S)"
            self.group_listbox.insert(tk.END, display)
        
        if preserve_selection and old_idx is not None and old_idx < len(self.groups):
            self.group_listbox.selection_set(old_idx)
        self._updating_listbox = False

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
        for group_num in range(1, total_rows + 1):
            # Find which group this row belongs to
            for group in groups:
                if group['row_start'] <= group_num <= group['row_end']:
                    distance_ft = abs(group_num - group['cb_position']) * row_spacing_ft
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
        if hasattr(self, 'breaker_size_var'):
            self.breaker_size_var.set('400')
        if hasattr(self, 'dc_feeder_distance_var'):
            self.dc_feeder_distance_var.set(str(estimate_data.get('dc_feeder_distance', 500)))
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
            self.groups = saved_groups
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
        
        self._loading = False

        # Derive module from templates
        self._derive_module_from_templates()

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

        estimate_data['wire_gauge'] = self.wire_gauge_var.get()
        
        # Save inverter selection
        if self.selected_inverter:
            estimate_data['inverter_name'] = self.inverter_select_var.get()
        
        # Save topology and DC:AC ratio
        estimate_data['topology'] = self.topology_var.get()
        estimate_data['breaker_size'] = self.breaker_size_var.get()
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
        
        # Save groups (new format)
        estimate_data['groups'] = self.groups
        estimate_data['subarrays'] = {}
        
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
            'groups': self.groups
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
        """Clear the groups and details when switching/deleting estimates"""
        # Clear groups
        self.groups.clear()
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
        
        if not inv_summary or not inv_summary.get('allocation_result'):
            from tkinter import messagebox
            messagebox.showinfo("No Data", "Run Calculate Estimate first to generate preview data.")
            return
        
        topology = self.topology_var.get()
        SitePreviewWindow(self, inv_summary, topology, self.INVERTER_COLORS, self.groups)

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
        for group in self.groups:
            for seg in group['segments']:
                if seg['strings_per_tracker'] == strings_per_tracker and seg['quantity'] > 0:
                    return self.parse_harness_config(seg['harness_config'])
        # Fallback: single harness equal to string count
        return [strings_per_tracker]

    def _adjust_harnesses_for_splits(self, totals):
        """Adjust harness counts based on inverter allocation split trackers.
        
        Uses the harness_map from allocate_strings_sequential() to identify
        split trackers and replace their original harness configs with
        harnesses matching the split amounts.
        
        E.g., a 3-string tracker with harness config '3' split 1/2 →
              remove one 3-string harness, add one 1-string + one 2-string.
        """
        inv_summary = totals.get('inverter_summary', {})
        allocation_result = inv_summary.get('allocation_result')
        
        if not allocation_result:
            return
        
        # Collect split tracker info from harness_map
        # Each split tracker appears in multiple inverters with position head/tail/middle
        split_trackers = {}  # tracker_idx -> list of harness_map entries
        
        for inv in allocation_result['inverters']:
            for entry in inv['harness_map']:
                if entry['is_split']:
                    tidx = entry['tracker_idx']
                    if tidx not in split_trackers:
                        split_trackers[tidx] = []
                    split_trackers[tidx].append(entry)
        
        if not split_trackers:
            return
        
        split_count = 0
        
        for tidx, entries in split_trackers.items():
            spt = entries[0]['strings_per_tracker']
            original_harness_sizes = self._get_harness_config_for_tracker_type(spt)
            
            # Determine new harness sizes after split by distributing
            # strings_taken across existing harnesses, keeping whole
            # harnesses intact when possible.
            split_amounts = sorted([e['strings_taken'] for e in entries], reverse=True)
            remaining_harnesses = sorted(original_harness_sizes, reverse=True)
            
            new_harnesses = []
            for amount in split_amounts:
                assigned = 0
                while assigned < amount and remaining_harnesses:
                    h = remaining_harnesses[0]
                    if assigned + h <= amount:
                        # Whole harness fits in this split portion
                        new_harnesses.append(h)
                        remaining_harnesses.pop(0)
                        assigned += h
                    else:
                        # Must split this harness at the boundary
                        needed = amount - assigned
                        new_harnesses.append(needed)
                        leftover = h - needed
                        remaining_harnesses.pop(0)
                        if leftover > 0:
                            remaining_harnesses.insert(0, leftover)
                        assigned = amount
            
            # Add any remaining harnesses that weren't consumed
            new_harnesses.extend(remaining_harnesses)
            
            # Remove original harness(es) for this one split tracker
            for size in original_harness_sizes:
                if size in totals['harnesses_by_size']:
                    totals['harnesses_by_size'][size] -= 1
                    if totals['harnesses_by_size'][size] <= 0:
                        del totals['harnesses_by_size'][size]
            
            # Add the new (possibly unchanged) harnesses
            for size in new_harnesses:
                if size not in totals['harnesses_by_size']:
                    totals['harnesses_by_size'][size] = 0
                totals['harnesses_by_size'][size] += 1
            
            split_count += 1
        
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
        self.topology_var.trace_add('write', lambda *args: (self._update_distance_hints(), self._mark_stale(), self._schedule_autosave()))
        
        ttk.Label(topology_row, text="DC:AC Ratio:").pack(side='left', padx=(0, 5))
        self.dc_ac_ratio_var = tk.StringVar(value='1.25')
        ttk.Spinbox(
            topology_row, from_=1.0, to=2.0, increment=0.05,
            textvariable=self.dc_ac_ratio_var, width=6, format='%.2f'
        ).pack(side='left', padx=(0, 15))
        self.dc_ac_ratio_var.trace_add('write', lambda *args: (self._update_strings_per_inverter(), self._mark_stale(), self._schedule_autosave()))
        
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
        self.breaker_size_var.trace_add('write', lambda *args: (self._mark_stale(), self._schedule_autosave()))

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
        ttk.Label(header_frame, text="", width=4).pack(side='left', padx=2)
        
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
            self._mark_stale()
            self._schedule_autosave()
            
            # Update derived module from templates
            self._derive_module_from_templates()
            
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
            self._mark_stale()
            self._schedule_autosave()
        qty_var.trace_add('write', on_qty_change)
        
        def on_harness_change(*args):
            segment['harness_config'] = harness_var.get()
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
        
        # Validate — need at least one linked template or a legacy fallback module
        self._derive_module_from_templates()
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
        
        # Track unique modules across all groups
        unique_modules = {}  # "Manufacturer Model (WattageW)" -> module_spec_dict
        
        # Per-segment module data for geometry calculations
        segment_module_data = []  # list of {module_spec_dict, modules_per_string, qty, spt}

        for group in self.groups:
            for seg in group['segments']:
                qty = seg['quantity']
                spt = seg['strings_per_tracker']
                harness_config = seg['harness_config']

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

                # Add to flat tracker sequence (one entry per physical tracker)
                for _ in range(qty):
                    tracker_sequence.append(spt)

                total_all_trackers += qty
                total_all_strings += qty * spt

                # Count trackers by string count
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
        
        # Store unique modules for Excel export
        totals['segment_module_data'] = segment_module_data

        # ==================== Module geometry (primary module for global calcs) ====================
        module_isc = self.selected_module.isc
        module_width_mm = self.selected_module.width_mm
        module_width_ft = module_width_mm / 304.8
        string_length_ft = module_width_ft * modules_per_string

        # ==================== Allocation ====================
        allocation_result = None

        if self.selected_inverter and strings_per_inv > 0 and total_all_strings > 0:
            allocation_result = allocate_strings_sequential(tracker_sequence, strings_per_inv)

            module_wattage = self.selected_module.wattage
            actual_dc_ac = self.selected_inverter.dc_ac_ratio(
                strings_per_inv, module_wattage, modules_per_string
            )

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
            fuse_current = module_isc * 1.56 * max(max_harness_strings, 1)
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

        # Combiner box matching
        if num_combiners > 0:
            if breaker_size not in totals['combiners_by_breaker']:
                totals['combiners_by_breaker'][breaker_size] = 0
            totals['combiners_by_breaker'][breaker_size] += num_combiners

            fuse_current = module_isc * 1.56 * max(max_harness_strings, 1)
            fuse_holder_rating = self.get_fuse_holder_category(fuse_current)

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
                    'block_name': 'Site Total'
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
                    'block_name': 'Site Total'
                })

        # ==================== Whip calculation ====================
        try:
            row_spacing = float(self.row_spacing_var.get())
        except ValueError:
            row_spacing = 20.0

        if total_all_trackers > 0 and num_devices > 0:
            whip_distances = self.calculate_cb_whip_distances(
                total_all_trackers, num_devices, row_spacing
            )
            for distance_ft, device_idx in whip_distances:
                whip_length = self.round_whip_length(distance_ft)
                whips_at_length = 2  # pos + neg per tracker group
                if whip_length not in totals['whips_by_length']:
                    totals['whips_by_length'][whip_length] = 0
                totals['whips_by_length'][whip_length] += whips_at_length
                totals['total_whip_length'] += whip_length * whips_at_length

        # ==================== Extenders ====================
        totals['extenders_short'] = total_all_harnesses
        totals['extenders_long'] = total_all_harnesses

        short_extender_length = self.round_whip_length(10)
        long_extender_length = self.round_whip_length(string_length_ft)

        # ==================== Harness split adjustment ====================
        # NOTE: _adjust_harnesses_for_splits will be updated to use harness_map
        # in the next batch. For now it's a no-op since allocations=[] above.
        self._adjust_harnesses_for_splits(totals)

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
            inv_name = f"{self.selected_inverter.manufacturer} {self.selected_inverter.model}" if self.selected_inverter else "Inverter"
            inv_summary = totals.get('inverter_summary', {})
            actual_ratio = inv_summary.get('actual_dc_ac', 0)
            insert_row(f"{inv_name} (DC:AC {actual_ratio:.2f})", '', totals['string_inverters'], 'ea')
            # Show allocation summary
            if allocation_result:
                summary = allocation_result['summary']
                split_count = summary.get('total_split_trackers', 0)
                max_spi = summary.get('max_strings_per_inverter', 0)
                min_spi = summary.get('min_strings_per_inverter', 0)
                n_larger = summary.get('num_larger_inverters', 0)
                n_smaller = summary.get('num_smaller_inverters', 0)
                if max_spi == min_spi:
                    size_str = f"all {max_spi} strings"
                else:
                    size_str = f"{n_larger}x{max_spi}str + {n_smaller}x{min_spi}str"
                insert_row(
                    f"  Allocation: {size_str}, {split_count} split tracker(s)",
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
            
            # ========== ROW SUMMARY SECTION ==========
            ws.merge_cells(f'A{row}:E{row}')
            cell = ws.cell(row=row, column=1, value="Group Configuration Summary")
            cell.font = title_font
            row += 1
            
            # Row summary headers
            row_headers = ['Group', 'Segment Configs', 'Total Strings', 'Total Trackers']
            for col, header in enumerate(row_headers, 1):
                cell = ws.cell(row=row, column=col, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = center_align
                cell.border = thin_border
            row += 1
            
            # Group data
            for r in self.groups:
                group_strings = sum(s['quantity'] * s['strings_per_tracker'] for s in r['segments'])
                group_trackers = sum(s['quantity'] for s in r['segments'])
                seg_summary = ", ".join(
                    f"{s['quantity']}x{s['strings_per_tracker']}S({s['harness_config']})"
                    for s in r['segments'] if s['quantity'] > 0
                )
                group_data = [r['name'], seg_summary, group_strings, group_trackers]
                for col, value in enumerate(group_data, 1):
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
                            cell = ws.cell(row=row, column=col, value=value)
                            cell.border = thin_border
                            cell.alignment = center_align
                        row += 1
                
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
    
    def __init__(self, parent, inv_summary, topology, colors, groups=None):
        super().__init__(parent)
        self.title("Site Preview — Inverter Allocation")
        self.geometry("1100x750")
        self.minsize(600, 400)
        
        self.inv_summary = inv_summary
        self.topology = topology
        self.colors = colors
        self.groups = groups or []
        
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
        
        # Color swatches group
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
            ttk.Label(legend_frame, text=f"{size_str}  |  {split_count} split tracker(s)",
                     font=('Helvetica', 9), foreground='#555555').pack(anchor='w')
    
    def build_layout_data(self):
        """Build a group-based layout of trackers with their inverter color assignments"""
        self.group_layout = []  # List of groups, each group is a list of tracker dicts
        
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
        
        # Split tracker_map into groups based on self.groups segment data
        global_idx = 0
        for group_data in self.groups:
            group_trackers = []
            for seg in group_data['segments']:
                for _ in range(seg['quantity']):
                    if global_idx in tracker_map:
                        group_trackers.append(tracker_map[global_idx])
                    global_idx += 1
            self.group_layout.append({
                'name': group_data['name'],
                'trackers': group_trackers
            })
        
        # Also keep flat list for backward compat
        self.tracker_list = [tracker_map[i] for i in sorted(tracker_map.keys())]
        
        # World-space dimensions
        self.tracker_w = 8
        self.group_gap = 6
        self.string_h = 30
        self.string_gap = 2
        self.group_v_gap = 20  # Vertical gap between groups
        
        if not self.tracker_list:
            self.world_width = 0
            self.world_height = 0
            return
        
        max_strings = max(t['strings_per_tracker'] for t in self.tracker_list)
        max_trackers_in_group = max(len(r['trackers']) for r in self.group_layout) if self.group_layout else 0
        
        self.world_width = max_trackers_in_group * (self.tracker_w + self.group_gap) - self.group_gap
        
        tracker_height = max_strings * (self.string_h + self.string_gap) - self.string_gap
        num_groups = len(self.group_layout)
        self.world_height = num_groups * (tracker_height + self.group_v_gap) - self.group_v_gap if num_groups > 0 else 0
    
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
        """Draw the site layout on the canvas with groups stacked N-S"""
        self.canvas.delete('all')
        
        if not self.group_layout:
            return
        
        max_strings = max(t['strings_per_tracker'] for t in self.tracker_list) if self.tracker_list else 3
        tracker_height = max_strings * (self.string_h + self.string_gap) - self.string_gap
        
        for group_idx, group_data in enumerate(self.group_layout):
            # Y offset for this group
            group_y_offset = group_idx * (tracker_height + self.group_v_gap)
            
            # Draw group label
            label_x, label_y = self.world_to_canvas(-15, group_y_offset + tracker_height / 2)
            font_size = max(6, min(10, int(9 * self.scale)))
            self.canvas.create_text(
                label_x, label_y,
                text=group_data['name'], font=('Helvetica', font_size),
                fill='#333333', anchor='e'
            )
            
            for t_idx, tracker in enumerate(group_data['trackers']):
                spt = tracker['strings_per_tracker']
                assignments = tracker['assignments']
                
                # X position for this tracker (E-W)
                wx = t_idx * (self.tracker_w + self.group_gap)
                
                # Build string colors
                string_colors = []
                for assignment in assignments:
                    for _ in range(assignment['strings']):
                        string_colors.append(assignment['color'])
                
                # Draw each string
                for s_idx in range(spt):
                    if s_idx < len(string_colors):
                        color = string_colors[s_idx]
                    else:
                        color = '#D0D0D0'
                    
                    wy = group_y_offset + s_idx * (self.string_h + self.string_gap)
                    
                    sx1, sy1 = self.world_to_canvas(wx, wy)
                    sx2, sy2 = self.world_to_canvas(wx + self.tracker_w, wy + self.string_h)
                    
                    self.canvas.create_rectangle(
                        sx1, sy1, sx2, sy2,
                        fill=color, outline='#444444', width=1
                    )
                
                # Tracker outline
                ty1_world = group_y_offset
                ty2_world = group_y_offset + spt * (self.string_h + self.string_gap) - self.string_gap
                
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
                        text=f"T{t_idx+1}", font=('Helvetica', font_size), fill='#555555'
                    )
        
        # Compass indicator
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