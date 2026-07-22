"""Review/confirm dialog for a drawing-extraction import plan.

Consumes an :class:`ImportPlan` (from ``utils.extraction_import``) and lets the
user resolve it — map/create modules, review-and-edit motor fields, match or
skip the inverter, and choose which empty project-meta fields to fill. On accept
it returns an :class:`ImportDecisions`; ``main.py`` applies those decisions.

This dialog performs no library writes of its own except through the standard
Module Manager creation flow (which persists new modules exactly as the manager
normally does). Template/inverter/project changes are returned as decisions and
applied by the caller.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from dataclasses import dataclass, field
from typing import Optional, Dict, List, Any

from ..models.module import ModuleSpec
from ..utils.extraction_import import (
    ImportPlan, TemplateImportEntry, ModuleImportEntry, derive_motor_fields,
)
from ..utils.module_library import load_merged_module_specs
from ..utils.inverter_library import load_merged_inverter_specs


# Sentinel shown in the "map to existing" dropdown before a choice is made.
_UNMAPPED = "-- Select a module --"
_SKIP_INVERTER = "-- Skip (don't set an inverter) --"


@dataclass
class ImportDecisions:
    """Resolved decisions returned by the dialog on accept."""
    # Placeholder label -> resolved ModuleSpec (from library or freshly created).
    resolved_modules: Dict[str, ModuleSpec] = field(default_factory=dict)
    # Each entry: {'name', 'module_ref', 'template_data'} with motor fields
    # corrected by the user and module_spec still un-embedded (caller embeds it).
    templates: List[Dict[str, Any]] = field(default_factory=list)
    # Selected library inverter key, or None to skip.
    inverter_name: Optional[str] = None
    # ProjectMetadata attribute -> value, only for fields the user chose to fill.
    project_meta_fills: Dict[str, str] = field(default_factory=dict)


def _join_location(*parts) -> str:
    """Join non-empty location parts into a single comma-separated string.

    The extraction's ``address`` is usually the full one-line address, which
    already contains ``city_state_zip``; appending that again would repeat it
    ("Pine Hill Rd, Cross Plains, WI 53528, Cross Plains, WI 53528"). Skip any
    part already present in what has been kept so far. A street-only address
    (the case this join exists for) still picks up the city/state/zip.
    """
    def _comparable(s: str) -> str:
        # Collapse whitespace and casefold so extraction spacing/case drift
        # ("Dexter, ME 04930" vs "dexter,  me 04930") still counts as a repeat.
        return ' '.join(s.split()).casefold()

    kept: List[str] = []
    for part in parts:
        if not part:
            continue
        text = str(part).strip()
        if not text:
            continue
        if _comparable(text) in _comparable(", ".join(kept)):
            continue
        kept.append(text)
    return ", ".join(kept)


class ExtractionImportDialog(tk.Toplevel):
    def __init__(self, parent, plan: ImportPlan, current_project=None):
        super().__init__(parent)
        self.title("Import Drawing Extraction")
        self.resizable(True, True)
        self.plan = plan
        self.current_project = current_project
        self.result: Optional[ImportDecisions] = None

        # Library data used for mapping/matching.
        self._lib_modules, _ = load_merged_module_specs()      # {key: ModuleSpec}
        self._lib_inverters, _ = load_merged_inverter_specs()  # {key: InverterSpec}

        # Per-module-label UI state.
        self._module_choice_vars: Dict[str, tk.StringVar] = {}
        self._module_showall_vars: Dict[str, tk.BooleanVar] = {}
        self._module_combos: Dict[str, ttk.Combobox] = {}
        self._resolved_modules: Dict[str, ModuleSpec] = {}
        # Newly created specs, keyed by their display string, so they show up in
        # every row's dropdown after creation.
        self._created_specs: Dict[str, ModuleSpec] = {}

        # Per-template motor-field vars, keyed by template index.
        self._tpl_vars: Dict[int, Dict[str, tk.StringVar]] = {}
        self._tpl_ack_vars: Dict[int, tk.BooleanVar] = {}

        # Inverter / project-meta state.
        self._inverter_var: Optional[tk.StringVar] = None
        # Per meta field: a checkbox (write this field at all?) and an editable
        # value seeded from the extraction — the user can correct a wrong value
        # or supply one the extraction missed.
        self._meta_fill_vars: Dict[str, tk.BooleanVar] = {}
        self._meta_entry_vars: Dict[str, tk.StringVar] = {}

        self._build_ui()
        self._center_on_parent(parent)
        self.grab_set()  # Modal (after build so geometry is settled)

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _center_on_parent(self, parent):
        self.update_idletasks()
        try:
            px, py = parent.winfo_rootx(), parent.winfo_rooty()
            pw, ph = parent.winfo_width(), parent.winfo_height()
            w, h = self.winfo_width(), self.winfo_height()
            x = px + max(0, (pw - w) // 2)
            y = py + max(0, (ph - h) // 2)
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

    def _build_ui(self):
        # Scrollable body + fixed button bar.
        outer = ttk.Frame(self)
        outer.pack(fill='both', expand=True)

        canvas = tk.Canvas(outer, borderwidth=0, highlightthickness=0, width=760, height=620)
        vsb = ttk.Scrollbar(outer, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        body = ttk.Frame(canvas, padding=10)
        body_id = canvas.create_window((0, 0), window=body, anchor='nw')
        body.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>', lambda e: canvas.itemconfigure(body_id, width=e.width))
        # Mouse-wheel scrolling while the pointer is over the canvas.
        canvas.bind('<Enter>', lambda e: canvas.bind_all(
            '<MouseWheel>', lambda ev: canvas.yview_scroll(int(-ev.delta / 120), 'units')))
        canvas.bind('<Leave>', lambda e: canvas.unbind_all('<MouseWheel>'))

        self._build_modules_section(body)
        self._build_templates_section(body)
        self._build_inverter_section(body)
        self._build_project_section(body)
        self._build_hints_section(body)
        if self.plan.warnings:
            self._build_plan_warnings_section(body)

        # Button bar.
        btn_bar = ttk.Frame(self, padding=(10, 6))
        btn_bar.pack(fill='x', side='bottom')
        ttk.Button(btn_bar, text="Cancel", command=self._on_cancel).pack(side='right', padx=(5, 0))
        ttk.Button(btn_bar, text="Accept & Import", command=self._on_accept).pack(side='right')

    def _section(self, parent, title):
        frame = ttk.LabelFrame(parent, text=title, padding=8)
        frame.pack(fill='x', expand=True, pady=(0, 10))
        return frame

    def _build_modules_section(self, parent):
        frame = self._section(parent, "Modules — map each placeholder to a real module")
        if not self.plan.modules:
            ttk.Label(frame, text="No modules in extraction.").pack(anchor='w')
            return

        for entry in self.plan.modules:
            self._build_module_row(frame, entry)

    def _build_module_row(self, parent, entry: ModuleImportEntry):
        row = ttk.Frame(parent, padding=(0, 4))
        row.pack(fill='x', expand=True)

        watt = f"{entry.wattage:g} W" if entry.wattage is not None else "wattage unknown"
        header = f"{entry.label}  ({watt})"
        if entry.model:
            header += f"   model: {entry.model}"
        ttk.Label(row, text=header, font=('TkDefaultFont', 9, 'bold')).grid(
            row=0, column=0, columnspan=3, sticky='w')

        if entry.layout_hints:
            hints = "  ·  ".join(f"{k}={v}" for k, v in entry.layout_hints.items())
            ttk.Label(row, text=f"Layout hints (read-only): {hints}",
                      foreground='gray40').grid(row=1, column=0, columnspan=3, sticky='w')

        ttk.Label(row, text="Map to existing:").grid(row=2, column=0, sticky='w', pady=(2, 0))

        choice_var = tk.StringVar(value=_UNMAPPED)
        self._module_choice_vars[entry.label] = choice_var
        combo = ttk.Combobox(row, textvariable=choice_var, state='readonly', width=48)
        combo.grid(row=2, column=1, sticky='we', padx=5, pady=(2, 0))
        combo.bind('<<ComboboxSelected>>', lambda e, lbl=entry.label: self._on_module_mapped(lbl))
        self._module_combos[entry.label] = combo

        ttk.Button(row, text="Create new…",
                   command=lambda ent=entry: self._create_new_module(ent)).grid(
            row=2, column=2, sticky='w', padx=5, pady=(2, 0))

        showall_var = tk.BooleanVar(value=False)
        self._module_showall_vars[entry.label] = showall_var
        ttk.Checkbutton(row, text="Show all wattages", variable=showall_var,
                        command=lambda lbl=entry.label: self._refresh_module_combo(lbl)).grid(
            row=3, column=1, sticky='w', padx=5)

        ttk.Separator(parent, orient='horizontal').pack(fill='x', pady=2)
        row.columnconfigure(1, weight=1)

        self._refresh_module_combo(entry.label)

    def _module_display(self, key: str, spec: ModuleSpec) -> str:
        return f"{key} ({spec.wattage:g}W)"

    def _candidate_specs(self) -> Dict[str, ModuleSpec]:
        """All selectable specs (library + freshly created), keyed by display."""
        out = {}
        for key, spec in self._lib_modules.items():
            out[self._module_display(key, spec)] = spec
        out.update(self._created_specs)
        return out

    def _refresh_module_combo(self, label: str):
        """Populate one row's dropdown, filtered by wattage unless 'show all'."""
        entry = next(e for e in self.plan.modules if e.label == label)
        show_all = self._module_showall_vars[label].get()
        candidates = self._candidate_specs()

        values = []
        for disp, spec in sorted(candidates.items()):
            if show_all or entry.wattage is None:
                values.append(disp)
            elif round(spec.wattage) == round(entry.wattage):
                values.append(disp)

        combo = self._module_combos[label]
        combo['values'] = values
        # Preserve an existing valid selection; otherwise reset to unmapped.
        if self._module_choice_vars[label].get() not in values:
            if label not in self._resolved_modules:
                self._module_choice_vars[label].set(_UNMAPPED)

    def _on_module_mapped(self, label: str):
        disp = self._module_choice_vars[label].get()
        spec = self._candidate_specs().get(disp)
        if spec is not None:
            self._resolved_modules[label] = spec

    def _create_new_module(self, entry: ModuleImportEntry):
        """Open the standard Module Manager creation flow pre-filled with the
        placeholder wattage and drawing label."""
        from .module_manager import ModuleManager

        win = tk.Toplevel(self)
        win.title(f"Create Module for '{entry.label}'")
        win.transient(self)

        last = {'spec': None}

        def on_saved(spec):
            last['spec'] = spec

        mgr = ModuleManager(win, on_module_selected=on_saved)
        mgr.pack(fill='both', expand=True)
        # Pre-fill wattage and drawing label as a starting point.
        if entry.wattage is not None:
            mgr.wattage_var.set(f"{entry.wattage:g}")
        mgr.model_var.set(entry.label)

        bar = ttk.Frame(win, padding=(8, 6))
        bar.pack(fill='x', side='bottom')
        ttk.Label(bar, text="Save the module, then click 'Use This Module'.",
                  foreground='gray40').pack(side='left')

        def use_it():
            spec = last['spec']
            if spec is None:
                messagebox.showwarning(
                    "No module saved",
                    "Save the module first (Save Module), then click 'Use This Module'.",
                    parent=win)
                return
            self._commit_created_module(entry.label, spec)
            win.destroy()

        ttk.Button(bar, text="Cancel", command=win.destroy).pack(side='right', padx=(5, 0))
        ttk.Button(bar, text="Use This Module", command=use_it).pack(side='right')

        win.grab_set()

    def _commit_created_module(self, label: str, spec: ModuleSpec):
        key = f"{spec.manufacturer} {spec.model}"
        disp = self._module_display(key, spec)
        self._created_specs[disp] = spec
        self._resolved_modules[label] = spec
        # Refresh every row's dropdown so the new module is selectable elsewhere,
        # then select it for the row that created it.
        for lbl in self._module_choice_vars:
            self._refresh_module_combo(lbl)
        self._module_choice_vars[label].set(disp)
        self._refresh_module_combo(label)
        self._module_choice_vars[label].set(disp)

    def _build_templates_section(self, parent):
        frame = self._section(parent, "Tracker templates — motor fields need review")
        if not self.plan.templates:
            ttk.Label(frame, text="No tracker templates in extraction.").pack(anchor='w')
            return

        for idx, tpl in enumerate(self.plan.templates):
            self._build_template_row(frame, idx, tpl)

    def _build_template_row(self, parent, idx: int, tpl: TemplateImportEntry):
        box = ttk.LabelFrame(parent, text=tpl.name, padding=6)
        box.pack(fill='x', expand=True, pady=4)

        td = tpl.template_data
        summary = (f"module_ref: {tpl.module_ref}   ·   "
                   f"{td['module_orientation']}, {td['modules_high']}-high   ·   "
                   f"{td['modules_per_string']} mod/string × {td['strings_per_tracker']} strings"
                   f"   ·   {tpl.modules_per_tracker} mod/tracker")
        ttk.Label(box, text=summary).grid(row=0, column=0, columnspan=4, sticky='w')

        r = 1
        # Warnings.
        for w in tpl.warnings:
            ttk.Label(box, text=f"⚠ {w}", foreground='#b00020', wraplength=700,
                      justify='left').grid(row=r, column=0, columnspan=4, sticky='w')
            r += 1
        # Acknowledge checkbox for invariant failures.
        if any('check failed' in w for w in tpl.warnings):
            ack = tk.BooleanVar(value=False)
            self._tpl_ack_vars[idx] = ack
            ttk.Checkbutton(box, text="I acknowledge this template's invariant warning(s)",
                            variable=ack).grid(row=r, column=0, columnspan=4, sticky='w')
            r += 1

        # Informational notes (e.g. we overrode the drawing's stated motor position).
        for n in tpl.notes:
            ttk.Label(box, text=f"ℹ {n}", foreground='#0a5', wraplength=700,
                      justify='left').grid(row=r, column=0, columnspan=4, sticky='w')
            r += 1

        # Read-only raw motor context from the drawing.
        rm = tpl.raw_motor
        ctx = (f"Drawing motor values (context only): placement={rm.get('motor_placement')}, "
               f"after_string={rm.get('motor_after_string')}, in_string={rm.get('motor_in_string')}, "
               f"split={rm.get('split_north')}/{rm.get('split_south')} (tracker-wide)")
        ttk.Label(box, text=ctx, foreground='gray40', wraplength=700,
                  justify='left').grid(row=r, column=0, columnspan=4, sticky='w', pady=(4, 2))
        r += 1

        ttk.Label(box, text="Review motor position:", font=('TkDefaultFont', 9, 'bold')).grid(
            row=r, column=0, columnspan=4, sticky='w')
        r += 1

        # Single editable input: modules north of the motor (matches the drawing).
        # South and the app's placement/split are derived and shown live.
        mps = td['modules_per_string']
        full_strings = int(td['strings_per_tracker'])
        total = mps * full_strings

        vars_ = {'mps': mps, 'full_strings': full_strings, 'total': total}
        north_seed = (tpl.modules_north_of_motor
                      if tpl.modules_north_of_motor is not None
                      else td.get('motor_split_north', 0))
        north_var = tk.StringVar(value=str(north_seed))
        vars_['north'] = north_var

        ttk.Label(box, text=f"Modules north of motor (of {total}):").grid(
            row=r, column=0, columnspan=2, sticky='w')
        ttk.Entry(box, textvariable=north_var, width=8).grid(row=r, column=2, sticky='w', padx=5)
        r += 1

        interp_label = ttk.Label(box, text="", foreground='gray20', wraplength=700, justify='left')
        interp_label.grid(row=r, column=0, columnspan=4, sticky='w')

        def _update_interp(*_a, _lbl=interp_label, _nv=north_var, _mps=mps,
                           _fs=full_strings, _total=total):
            txt = _nv.get().strip()
            try:
                n = int(txt)
            except ValueError:
                _lbl.config(text="→ enter a whole number", foreground='#b00020')
                return
            try:
                d = derive_motor_fields(n, _mps, _fs)
            except ValueError:
                _lbl.config(text=f"→ must be between 0 and {_total}", foreground='#b00020')
                return
            south = _total - n
            if d['motor_placement_type'] == 'between_strings':
                desc = f"between strings, after string {d['motor_position_after_string']}"
            else:
                desc = (f"in string {d['motor_string_index']}, "
                        f"split {d['motor_split_north']}/{d['motor_split_south']} within it")
            _lbl.config(text=f"→ South: {south}.  Motor {desc}.", foreground='gray20')

        north_var.trace_add('write', _update_interp)
        _update_interp()

        self._tpl_vars[idx] = vars_
        box.columnconfigure(1, weight=1)

    def _build_inverter_section(self, parent):
        frame = self._section(parent, "Inverter — match only (never created)")
        name = self.plan.inverter.name
        qty = self.plan.inverter.qty

        matched_key = self._match_inverter(name)

        if name:
            ttk.Label(frame, text=f"Extraction inverter: {name}"
                                  + (f"  (qty {qty})" if qty is not None else "")).pack(anchor='w')
        else:
            ttk.Label(frame, text="Extraction did not name an inverter.").pack(anchor='w')

        if matched_key:
            ttk.Label(frame, text=f"✓ Matched library inverter: {matched_key}",
                      foreground='#0a7d18').pack(anchor='w', pady=(2, 0))
            self._inverter_var = tk.StringVar(value=matched_key)
        else:
            if name:
                ttk.Label(frame, text="No exact library match. Pick one below or skip.",
                          foreground='#b00020').pack(anchor='w', pady=(2, 0))
            self._inverter_var = tk.StringVar(value=_SKIP_INVERTER)
            picker = ttk.Combobox(frame, textvariable=self._inverter_var, state='readonly', width=50)
            picker['values'] = [_SKIP_INVERTER] + sorted(self._lib_inverters.keys())
            picker.pack(anchor='w', pady=(2, 0))

    def _match_inverter(self, name: Optional[str]) -> Optional[str]:
        if not name:
            return None
        if name in self._lib_inverters:
            return name
        lowered = {k.lower(): k for k in self._lib_inverters}
        return lowered.get(name.lower())

    def _build_project_section(self, parent):
        frame = self._section(parent, "Project info — fills only currently-empty fields")
        pm = self.plan.project_meta
        md = getattr(self.current_project, 'metadata', None)

        # (attr, extracted value, label)
        location_value = _join_location(pm.address, pm.city_state_zip, pm.coordinates)
        fillables = [
            ('client', pm.customer, 'Client (from customer)'),
            ('name', pm.name, 'Project name'),
            ('location', location_value, 'Location (address / city-state-zip / coordinates)'),
        ]

        any_row = False
        for attr, value, caption in fillables:
            current = getattr(md, attr, None) if md else None
            if current:
                # Never overwrite a field the project already has.
                if value:
                    any_row = True
                    ttk.Label(frame, text=f"{caption}: '{value}'  —  already set to '{current}', skipping",
                              foreground='gray40').pack(anchor='w')
                continue

            # Editable row: the checkbox decides whether the field is written,
            # the entry carries the value. Seeded from the extraction, but the
            # user can correct a wrong value or type one it missed.
            any_row = True
            fill_var = tk.BooleanVar(value=True)
            entry_var = tk.StringVar(value=value or '')
            self._meta_fill_vars[attr] = fill_var
            self._meta_entry_vars[attr] = entry_var

            row = ttk.Frame(frame)
            row.pack(anchor='w', fill='x', pady=1)
            ttk.Checkbutton(row, text=f"Fill {caption}:", variable=fill_var,
                            width=22).pack(side='left')
            ttk.Entry(row, textvariable=entry_var, width=40).pack(side='left', padx=5)
            if not value:
                ttk.Label(row, text="not found in the drawing",
                          foreground='gray40').pack(side='left')
        if not any_row:
            ttk.Label(frame, text="No project-meta fields to fill.").pack(anchor='w')

        # Cross-check display only (never written).
        cc = []
        if pm.dc_capacity_kw is not None:
            cc.append(f"DC {pm.dc_capacity_kw} kW")
        if pm.ac_capacity_kw is not None:
            cc.append(f"AC {pm.ac_capacity_kw} kW")
        if pm.dc_ac_ratio is not None:
            cc.append(f"DC/AC {pm.dc_ac_ratio}")
        if pm.total_modules is not None:
            cc.append(f"{pm.total_modules} modules")
        if pm.total_strings is not None:
            cc.append(f"{pm.total_strings} strings")
        if pm.inverter_qty is not None:
            cc.append(f"{pm.inverter_qty} inverters")
        if self.plan.tracker_manufacturer:
            cc.append(f"tracker mfr {self.plan.tracker_manufacturer}")
        if cc:
            ttk.Label(frame, text="Cross-check (display only, never written): " + "  ·  ".join(cc),
                      foreground='gray40', wraplength=700, justify='left').pack(anchor='w', pady=(6, 0))

    def _build_hints_section(self, parent):
        hints = []
        for tpl in self.plan.templates:
            q = tpl.layout_hints.get('quantity')
            if q is not None:
                hints.append(f"{tpl.name}: qty {q}")
        if not hints:
            return
        frame = self._section(parent, "Layout hints (read-only)")
        for h in hints:
            ttk.Label(frame, text=h, foreground='gray40').pack(anchor='w')

    def _build_plan_warnings_section(self, parent):
        frame = self._section(parent, "Import warnings")
        for w in self.plan.warnings:
            ttk.Label(frame, text=f"⚠ {w}", foreground='#b00020',
                      wraplength=700, justify='left').pack(anchor='w')

    # ------------------------------------------------------------------
    # Accept / validate
    # ------------------------------------------------------------------

    def _on_cancel(self):
        self.result = None
        self.destroy()

    def _on_accept(self):
        # 1) Every module must be mapped or created.
        unresolved = [e.label for e in self.plan.modules if e.label not in self._resolved_modules]
        if unresolved:
            messagebox.showwarning(
                "Unresolved modules",
                "Map or create a module for:\n\n" + "\n".join(f"  • {u}" for u in unresolved),
                parent=self)
            return

        # 1b) Every template must join to a resolved module. Without this, main.py
        # silently skips the template and the import is a no-op the user can't see.
        orphaned = [t.name for t in self.plan.templates
                    if t.module_ref not in self._resolved_modules]
        if orphaned:
            messagebox.showwarning(
                "Templates without a module",
                "These templates reference a module that isn't in the extraction, "
                "so they cannot be imported:\n\n"
                + "\n".join(f"  • {n}" for n in orphaned)
                + "\n\nFix the extraction's module_ref values and re-import.",
                parent=self)
            return

        # 2) Invariant-failing templates must be acknowledged.
        for idx, tpl in enumerate(self.plan.templates):
            if idx in self._tpl_ack_vars and not self._tpl_ack_vars[idx].get():
                messagebox.showwarning(
                    "Acknowledge warning",
                    f"Template '{tpl.name}' has an invariant warning. "
                    f"Check its acknowledgement box to proceed.",
                    parent=self)
                return

        # 3) Validate + collect corrected motor fields per template.
        resolved_templates = []
        for idx, tpl in enumerate(self.plan.templates):
            corrected = self._validate_template_motor(idx, tpl)
            if corrected is None:
                return  # message already shown
            resolved_templates.append({
                'name': tpl.name,
                'module_ref': tpl.module_ref,
                'template_data': corrected,
            })

        # 4) Project-meta: write a field only if it's checked AND has a value —
        # an unchecked box skips it, and an empty box has nothing to write.
        fills = {}
        for attr, fill_var in self._meta_fill_vars.items():
            if not fill_var.get():
                continue
            value = self._meta_entry_vars[attr].get().strip()
            if value:
                fills[attr] = value

        # 5) Inverter selection.
        inv = None
        if self._inverter_var is not None:
            chosen = self._inverter_var.get()
            if chosen and chosen != _SKIP_INVERTER:
                inv = chosen

        self.result = ImportDecisions(
            resolved_modules=dict(self._resolved_modules),
            templates=resolved_templates,
            inverter_name=inv,
            project_meta_fills=fills,
        )
        self.destroy()

    def _validate_template_motor(self, idx: int, tpl: TemplateImportEntry) -> Optional[Dict[str, Any]]:
        """Return a corrected copy of template_data with motor fields derived
        from the user's 'modules north of motor' input, or None (after showing a
        message) if the input is invalid."""
        vars_ = self._tpl_vars[idx]
        td = dict(tpl.template_data)  # shallow copy is fine (all scalars)

        def fail(msg):
            messagebox.showwarning("Motor position", f"Template '{tpl.name}': {msg}", parent=self)
            return None

        try:
            north = int(vars_['north'].get())
        except ValueError:
            return fail("enter a whole number for 'modules north of motor'.")

        try:
            motor_fields = derive_motor_fields(north, vars_['mps'], vars_['full_strings'])
        except ValueError:
            return fail(f"'modules north of motor' must be between 0 and {vars_['total']}.")

        td.update(motor_fields)
        return td
