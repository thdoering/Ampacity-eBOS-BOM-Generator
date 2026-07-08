"""Throwaway manual test for Change 2 (review dialog). Safe to delete.

Run from the project root:  python verify_change2.py

Opens the extraction-import review dialog against the Caledon sample with no
current project. Map the two placeholder modules (any library module, or use
'Create new…'), review a template's motor fields, then click Accept & Import.
The resolved decisions are printed to the console.
"""
import json
import tkinter as tk

from src.utils.extraction_import import build_import_plan
from src.ui.extraction_import_dialog import ExtractionImportDialog

with open('docs/samples/caledon_extraction.json') as f:
    data = json.load(f)

plan = build_import_plan(data)

root = tk.Tk()
root.title("Change 2 test host")
root.geometry("400x200")
tk.Label(root, text="Change 2 dialog test.\nClose this window when done.").pack(pady=40)


def open_dialog():
    dlg = ExtractionImportDialog(root, plan, current_project=None)
    root.wait_window(dlg)
    if dlg.result is None:
        print("\n--- Dialog cancelled (result is None) ---")
        return
    d = dlg.result
    print("\n--- ImportDecisions ---")
    print("Resolved modules:")
    for label, spec in d.resolved_modules.items():
        print(f"    {label!r} -> {spec.manufacturer} {spec.model} ({spec.wattage}W, "
              f"{spec.length_mm}x{spec.width_mm}mm)")
    print(f"\nTemplates ({len(d.templates)}):")
    for t in d.templates:
        td = t['template_data']
        motor = (f"placement={td['motor_placement_type']}, "
                 f"after={td['motor_position_after_string']}, "
                 f"in_string={td['motor_string_index']}, "
                 f"split={td['motor_split_north']}/{td['motor_split_south']}")
        print(f"    {t['name']}  (ref={t['module_ref']})  {motor}")
    print(f"\nInverter: {d.inverter_name}")
    print(f"Project meta fills: {d.project_meta_fills}")


tk.Button(root, text="Open Import Dialog", command=open_dialog).pack()
root.mainloop()
