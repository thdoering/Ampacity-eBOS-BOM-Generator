import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
from pathlib import Path
from typing import Optional, Callable, Dict, Set
from ..models.inverter import InverterSpec, MPPTChannel, MPPTConfig, InverterType
from ..utils.inverter_library import load_merged_inverter_specs, save_user_inverters

class InverterManager(ttk.Frame):
    def __init__(self, parent,
                 on_inverter_selected: Optional[Callable[[InverterSpec], None]] = None,
                 current_project_getter=None,
                 on_inverter_assignment_changed=None):
        super().__init__(parent)
        self.parent = parent
        self.on_inverter_selected = on_inverter_selected
        self.current_project_getter = current_project_getter  # callable → Project or None
        self.on_inverter_assignment_changed = on_inverter_assignment_changed
        self.inverters: Dict[str, InverterSpec] = {}
        self.factory_keys: Set[str] = set()

        self.setup_ui()
        self.load_inverters()
        
    def setup_ui(self):
        """Create and arrange UI components"""
        # Main container with padding
        main_container = ttk.Frame(self, padding="10")
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Left side - Inverter List
        list_frame = ttk.LabelFrame(main_container, text="Inverter Library", padding="5")
        list_frame.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Search bar
        search_frame = ttk.Frame(list_frame)
        search_frame.grid(row=0, column=0, columnspan=2, padx=5, pady=(5, 2), sticky=(tk.W, tk.E))
        ttk.Label(search_frame, text="Search:").pack(side='left', padx=(0, 4))
        self.search_var = tk.StringVar()
        self.search_var.trace_add('write', lambda *_: self.update_inverter_list())
        ttk.Entry(search_frame, textvariable=self.search_var).pack(side='left', fill='x', expand=True)

        self.inverter_tree = ttk.Treeview(list_frame, height=15, show='tree')
        self.inverter_tree.column('#0', width=420)
        self.inverter_tree.grid(row=1, column=0, padx=5, pady=5, sticky=(tk.N, tk.S, tk.E, tk.W))
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.inverter_tree.yview)
        scrollbar.grid(row=1, column=1, pady=5, sticky=(tk.N, tk.S))
        self.inverter_tree.configure(yscrollcommand=scrollbar.set)
        self.inverter_tree.bind('<<TreeviewSelect>>', self.on_inverter_select)
        self.inverter_tree.bind('<Button-3>', self._show_context_menu)

        button_frame = ttk.Frame(list_frame)
        button_frame.grid(row=2, column=0, padx=5, pady=5)
        
        ttk.Button(button_frame, text="Import OND", command=self.import_ond).grid(row=0, column=0, padx=2)
        ttk.Button(button_frame, text="Delete", command=self.delete_inverter).grid(row=0, column=1, padx=2)
        
        # Right side - Inverter Editor
        editor_frame = ttk.LabelFrame(main_container, text="Inverter Details", padding="5")
        editor_frame.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Basic Info
        ttk.Label(editor_frame, text="Manufacturer:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.manufacturer_var = tk.StringVar()
        ttk.Entry(editor_frame, textvariable=self.manufacturer_var).grid(row=0, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(editor_frame, text="Model:").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.model_var = tk.StringVar()
        ttk.Entry(editor_frame, textvariable=self.model_var).grid(row=1, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(editor_frame, text="Inverter Type:").grid(row=2, column=0, padx=5, pady=2, sticky=tk.W)
        self.inverter_type_var = tk.StringVar(value=InverterType.STRING.value)
        type_combo = ttk.Combobox(editor_frame, textvariable=self.inverter_type_var, state='readonly')
        type_combo['values'] = [t.value for t in InverterType]
        type_combo.grid(row=2, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(editor_frame, text="Rated AC Power (kW):").grid(row=3, column=0, padx=5, pady=2, sticky=tk.W)
        self.power_var = tk.StringVar()
        ttk.Entry(editor_frame, textvariable=self.power_var).grid(row=3, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(editor_frame, text="Max DC Power (kW):").grid(row=4, column=0, padx=5, pady=2, sticky=tk.W)
        self.max_dc_power_var = tk.StringVar()
        ttk.Entry(editor_frame, textvariable=self.max_dc_power_var).grid(row=4, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        # MPPT Configuration
        mppt_frame = ttk.LabelFrame(editor_frame, text="MPPT Configuration", padding="5")
        mppt_frame.grid(row=5, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        ttk.Label(mppt_frame, text="Number of MPPTs:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.num_mppts_var = tk.StringVar(value="1")
        ttk.Entry(mppt_frame, textvariable=self.num_mppts_var, width=8).grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)
        
        ttk.Label(mppt_frame, text="Inputs per MPPT:").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.inputs_per_mppt_var = tk.StringVar(value="1")
        ttk.Entry(mppt_frame, textvariable=self.inputs_per_mppt_var, width=8).grid(row=1, column=1, padx=5, pady=2, sticky=tk.W)
        
        ttk.Label(mppt_frame, text="Max Current per MPPT (A):").grid(row=2, column=0, padx=5, pady=2, sticky=tk.W)
        self.mppt_current_var = tk.StringVar(value="180")
        ttk.Entry(mppt_frame, textvariable=self.mppt_current_var, width=8).grid(row=2, column=1, padx=5, pady=2, sticky=tk.W)
        
        # Total inputs display (auto-calculated)
        self.total_inputs_label = ttk.Label(mppt_frame, text="Total String Inputs: 1", foreground="gray")
        self.total_inputs_label.grid(row=0, column=2, padx=(15, 5), pady=2, sticky=tk.W)
        
        # Bind auto-update of total inputs
        self.num_mppts_var.trace_add('write', lambda *_: self._update_total_inputs())
        self.inputs_per_mppt_var.trace_add('write', lambda *_: self._update_total_inputs())
        
        # Voltage Limits
        voltage_frame = ttk.LabelFrame(editor_frame, text="Voltage & Current Limits", padding="5")
        voltage_frame.grid(row=6, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        ttk.Label(voltage_frame, text="Max System Voltage (Vdc):").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.max_dc_voltage_var = tk.StringVar(value="1500")
        ttk.Entry(voltage_frame, textvariable=self.max_dc_voltage_var).grid(row=0, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(voltage_frame, text="MPPT Voltage Min (V):").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.mppt_voltage_min_var = tk.StringVar(value="855")
        ttk.Entry(voltage_frame, textvariable=self.mppt_voltage_min_var).grid(row=1, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(voltage_frame, text="MPPT Voltage Max (V):").grid(row=2, column=0, padx=5, pady=2, sticky=tk.W)
        self.mppt_voltage_max_var = tk.StringVar(value="1425")
        ttk.Entry(voltage_frame, textvariable=self.mppt_voltage_max_var).grid(row=2, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(voltage_frame, text="Max Short Circuit Current (A):").grid(row=3, column=0, padx=5, pady=2, sticky=tk.W)
        self.max_isc_var = tk.StringVar(value="")
        ttk.Entry(voltage_frame, textvariable=self.max_isc_var).grid(row=3, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(voltage_frame, text="Nominal AC Voltage (V):").grid(row=4, column=0, padx=5, pady=2, sticky=tk.W)
        self.nominal_ac_voltage_var = tk.StringVar(value="480")
        ttk.Entry(voltage_frame, textvariable=self.nominal_ac_voltage_var).grid(row=4, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(voltage_frame, text="Max AC Output Current (A):").grid(row=5, column=0, padx=5, pady=2, sticky=tk.W)
        self.max_ac_current_var = tk.StringVar(value="")
        ttk.Entry(voltage_frame, textvariable=self.max_ac_current_var).grid(row=5, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        # Buttons
        btn_frame = ttk.Frame(editor_frame)
        btn_frame.grid(row=7, column=0, columnspan=2, pady=10)
        
        ttk.Button(btn_frame, text="Save Inverter", command=self.save_inverter).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="New / Clear", command=self.clear_form).pack(side='left', padx=5)

    def _update_total_inputs(self):
        """Update the total string inputs display label"""
        try:
            num_mppts = int(self.num_mppts_var.get())
            inputs_per = int(self.inputs_per_mppt_var.get())
            total = num_mppts * inputs_per
            self.total_inputs_label.config(text=f"Total String Inputs: {total}")
        except (ValueError, TypeError):
            self.total_inputs_label.config(text="Total String Inputs: --")

    def _calculate_default_ac_current(self, rated_power_kw, nominal_ac_voltage):
        """Calculate default max AC output current from rated power and voltage.
        Assumes 3-phase: I = P / (V × √3)"""
        import math
        if nominal_ac_voltage <= 0:
            return 40.0
        return round(rated_power_kw * 1000 / (nominal_ac_voltage * math.sqrt(3)), 1)
            
    def create_inverter_spec(self) -> Optional[InverterSpec]:
        """Create InverterSpec from current UI values"""
        try:
            # Build MPPT channel list from simplified inputs
            num_mppts = int(self.num_mppts_var.get())
            inputs_per_mppt = int(self.inputs_per_mppt_var.get())
            mppt_current = float(self.mppt_current_var.get())
            mppt_v_min = float(self.mppt_voltage_min_var.get())
            mppt_v_max = float(self.mppt_voltage_max_var.get())
            rated_power_kw = float(self.power_var.get())
            
            # Auto-calculate max DC power if left blank
            max_dc_power_str = self.max_dc_power_var.get().strip()
            if max_dc_power_str:
                max_dc_power_kw = float(max_dc_power_str)
            else:
                # Default: rated AC / 0.985 (CEC efficiency estimate)
                max_dc_power_kw = round(rated_power_kw / 0.985, 1)
            
            # Per-MPPT max power = total max DC power / num MPPTs
            per_mppt_power_w = (max_dc_power_kw * 1000) / num_mppts
            
            channels = []
            for _ in range(num_mppts):
                channels.append(MPPTChannel(
                    max_input_current=mppt_current,
                    min_voltage=mppt_v_min,
                    max_voltage=mppt_v_max,
                    max_power=per_mppt_power_w,
                    num_string_inputs=inputs_per_mppt
                ))
            
            # Parse optional max short circuit current
            max_isc_str = self.max_isc_var.get().strip()
            max_isc = float(max_isc_str) if max_isc_str else None
            
            inverter = InverterSpec(
                manufacturer=self.manufacturer_var.get(),
                model=self.model_var.get(),
                inverter_type=InverterType(self.inverter_type_var.get()),
                rated_power_kw=rated_power_kw,
                max_dc_power_kw=max_dc_power_kw,
                max_efficiency=98.0,
                mppt_channels=channels,
                mppt_configuration=MPPTConfig.INDEPENDENT,
                max_dc_voltage=float(self.max_dc_voltage_var.get()),
                startup_voltage=mppt_v_min,
                nominal_ac_voltage=float(self.nominal_ac_voltage_var.get()),
                max_ac_current=float(self.max_ac_current_var.get()) if self.max_ac_current_var.get().strip() else self._calculate_default_ac_current(rated_power_kw, float(self.nominal_ac_voltage_var.get())),
                power_factor=0.99,
                dimensions_mm=(1000, 600, 300),
                weight_kg=75.0,
                ip_rating="IP65",
                max_short_circuit_current=max_isc
            )
            
            inverter.validate()
            return inverter
            
        except (ValueError, TypeError) as e:
            messagebox.showerror("Error", str(e))
            return None
            
    def save_inverter(self):
        """Save current inverter"""
        inverter = self.create_inverter_spec()
        if inverter:
            name = f"{inverter.manufacturer} {inverter.model}"
            if name in self.inverters:
                if not messagebox.askyesno("Confirm", f"Inverter '{name}' already exists. Overwrite?"):
                    return
                    
            self.inverters[name] = inverter
            self.save_inverters()
            self.update_inverter_list()
            
            if self.on_inverter_selected:
                self.on_inverter_selected(inverter)
                
            messagebox.showinfo("Success", f"Inverter '{name}' saved successfully")
            
    def load_inverters(self):
        """Load inverters from merged factory + user library."""
        try:
            self.inverters, self.factory_keys = load_merged_inverter_specs()
            self.update_inverter_list()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load inverters: {str(e)}")
            self.inverters = {}
            self.factory_keys = set()
            
    def save_inverters(self):
        """Save user-owned inverters (non-factory) to the user library."""
        save_user_inverters(self.inverters, self.factory_keys)
            
    def update_inverter_list(self):
        """Update the inverter treeview grouped by manufacturer."""
        for item in self.inverter_tree.get_children():
            self.inverter_tree.delete(item)

        filter_text = self.search_var.get().strip().lower() if hasattr(self, 'search_var') else ''

        # Build inverter_key -> list of estimate names that use it
        inv_to_estimates: Dict[str, list] = {}
        project = self.current_project_getter() if self.current_project_getter else None
        if project:
            for est_id, est_data in project.quick_estimates.items():
                inv_id = est_data.get('inverter_id')
                if inv_id:
                    inv_to_estimates.setdefault(inv_id, []).append(est_data.get('name', est_id))

        manufacturers: Dict[str, list] = {}
        for inv_key, inverter in self.inverters.items():
            manufacturers.setdefault(inverter.manufacturer, []).append((inv_key, inverter))

        for manufacturer in sorted(manufacturers):
            inv_list = manufacturers[manufacturer]

            if filter_text:
                mfr_matches = filter_text in manufacturer.lower()
                if mfr_matches:
                    visible = inv_list
                else:
                    visible = [(ik, inv) for ik, inv in inv_list if filter_text in inv.model.lower()]
                if not visible:
                    continue
                is_open = True
            else:
                visible = inv_list
                is_open = False

            # Mark the manufacturer node if any of its inverters are assigned
            mfr_has_assignment = any(inv_to_estimates.get(ik) for ik, _ in visible)
            mfr_label = f"● {manufacturer}" if mfr_has_assignment else manufacturer
            parent = self.inverter_tree.insert('', 'end', text=mfr_label, open=is_open)

            for inv_key, inverter in sorted(visible, key=lambda x: x[1].model):
                label = f"{inverter.model} ({inverter.rated_power_kw} kW)"
                assigned = inv_to_estimates.get(inv_key)
                if assigned:
                    label += f"  — selected ({', '.join(assigned)})"
                self.inverter_tree.insert(parent, 'end', text=label, values=(inv_key,))
            
    def import_ond(self):
        """Import inverter from OND file"""
        # TODO: Implement OND file parsing
        messagebox.showinfo("Not Implemented", "OND file import not yet implemented")
        
    def delete_inverter(self):
        """Delete selected inverter or all inverters under a manufacturer node"""
        selection = self.inverter_tree.selection()
        if not selection:
            return

        item = selection[0]
        values = self.inverter_tree.item(item, 'values')

        if values:
            # Leaf node — single inverter
            inv_key = values[0]
            if inv_key in self.factory_keys:
                messagebox.showinfo("Factory Entry", f"'{inv_key}' is part of the factory library and cannot be deleted.")
                return
            if messagebox.askyesno("Confirm", f"Delete inverter '{inv_key}'?"):
                del self.inverters[inv_key]
                self.save_inverters()
                self.update_inverter_list()
        else:
            # Manufacturer node — delete all children
            manufacturer = self.inverter_tree.item(item, 'text')
            children = self.inverter_tree.get_children(item)
            if not children:
                return
            deletable = [
                self.inverter_tree.item(c, 'values')[0]
                for c in children
                if self.inverter_tree.item(c, 'values')[0] not in self.factory_keys
            ]
            factory_count = len(children) - len(deletable)
            if not deletable:
                messagebox.showinfo("Factory Entries", f"All inverters under '{manufacturer}' are factory entries and cannot be deleted.")
                return
            msg = f"Delete {len(deletable)} user inverter(s) for '{manufacturer}'?"
            if factory_count:
                msg += f"\n({factory_count} factory entr{'y' if factory_count == 1 else 'ies'} will be kept.)"
            if messagebox.askyesno("Confirm", msg):
                for key in deletable:
                    del self.inverters[key]
                self.save_inverters()
                self.update_inverter_list()

    def clear_form(self):
        """Clear the editor form for entering a new inverter"""
        self.manufacturer_var.set('')
        self.model_var.set('')
        self.inverter_type_var.set(InverterType.STRING.value)
        self.power_var.set('')
        self.max_dc_power_var.set('')
        self.num_mppts_var.set('1')
        self.inputs_per_mppt_var.set('1')
        self.mppt_current_var.set('180')
        self.max_dc_voltage_var.set('1500')
        self.mppt_voltage_min_var.set('855')
        self.mppt_voltage_max_var.set('1425')
        self.max_isc_var.set('')
        self.nominal_ac_voltage_var.set('480')
        self.max_ac_current_var.set('')
        self._update_total_inputs()
        
        for item in self.inverter_tree.selection():
            self.inverter_tree.selection_remove(item)
            
    def on_inverter_select(self, event=None):
        """Handle inverter selection"""
        selection = self.inverter_tree.selection()
        if not selection:
            return

        values = self.inverter_tree.item(selection[0], 'values')
        if not values:
            return  # manufacturer node, not a leaf

        inv_key = values[0]
        inverter = self.inverters.get(inv_key)
        if not inverter:
            return
        
        # Update UI with selected inverter
        self.manufacturer_var.set(inverter.manufacturer)
        self.model_var.set(inverter.model)
        self.inverter_type_var.set(inverter.inverter_type.value if hasattr(inverter, 'inverter_type') and inverter.inverter_type else InverterType.STRING.value)
        self.power_var.set(str(inverter.rated_power_kw))
        self.max_dc_power_var.set(str(inverter.max_dc_power_kw))
        self.max_dc_voltage_var.set(str(inverter.max_dc_voltage))
        max_isc = getattr(inverter, 'max_short_circuit_current', None)
        self.max_isc_var.set(str(max_isc) if max_isc else "")
        self.nominal_ac_voltage_var.set(str(getattr(inverter, 'nominal_ac_voltage', 480.0)))
        self.max_ac_current_var.set(str(getattr(inverter, 'max_ac_current', '')))
        
        # Derive simplified MPPT fields from channel list
        channels = inverter.mppt_channels
        num_mppts = len(channels)
        self.num_mppts_var.set(str(num_mppts))
        
        if channels:
            self.inputs_per_mppt_var.set(str(channels[0].num_string_inputs))
            self.mppt_current_var.set(str(channels[0].max_input_current))
            self.mppt_voltage_min_var.set(str(channels[0].min_voltage))
            self.mppt_voltage_max_var.set(str(channels[0].max_voltage))

        self._update_total_inputs()

    def _show_context_menu(self, event):
        """Right-click handler on the inverter tree — shows Assign to estimate submenu."""
        row_id = self.inverter_tree.identify_row(event.y)
        if not row_id:
            return

        values = self.inverter_tree.item(row_id, 'values')
        if not values:
            return  # manufacturer node — no-op

        inv_key = values[0]

        menu = tk.Menu(self, tearoff=0)
        assign_menu = tk.Menu(menu, tearoff=0)
        menu.add_cascade(label="Assign to estimate ▶", menu=assign_menu)

        project = self.current_project_getter() if self.current_project_getter else None

        if not project:
            assign_menu.add_command(label="(no project loaded)", state='disabled')
        elif not project.quick_estimates:
            assign_menu.add_command(label="(no estimates)", state='disabled')
        else:
            for est_id, est_data in project.quick_estimates.items():
                est_name = est_data.get('name', est_id)
                assign_menu.add_command(
                    label=est_name,
                    command=lambda eid=est_id, ik=inv_key: self._assign_inverter_to_estimate(ik, eid)
                )

        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _assign_inverter_to_estimate(self, inverter_key, estimate_id):
        """Assign an inverter to a specific estimate and notify listeners."""
        project = self.current_project_getter() if self.current_project_getter else None
        if not project or estimate_id not in project.quick_estimates:
            return

        project.quick_estimates[estimate_id]['inverter_id'] = inverter_key
        self.update_inverter_list()

        if self.on_inverter_assignment_changed:
            self.on_inverter_assignment_changed(inverter_key, estimate_id)