"""Adapter for drawing-extraction JSON (``ExtractionResult`` from the external
``ampacity-rfp`` plugin).

This module is a *pure* adapter: it parses/validates an ``ExtractionResult``
dict and produces an in-memory "import plan" describing what an import would do.
It performs no UI work and no library writes. The review dialog
(``extraction_import_dialog``) consumes the plan; ``main.py`` applies the
resolved decisions.

Key facts about the input format that drive this design (see
``docs/extraction-output.schema.json`` and ``docs/samples/caledon_extraction.json``):

* Modules are placeholders. ``ModuleInfo.brand`` is a drawing-local label like
  ``"Module Type 1"`` — not a manufacturer — and carries no electrical or
  dimensional spec. Modules are resolved by the user, never auto-created here.
* ``TrackerTemplate.module_ref`` is a join key equal to a module's ``brand``.
* ``module_spacing_m`` is always a placeholder (0.01); we ignore the skill value
  and use the app default.
* Motor fields can be internally inconsistent and are always review-only.
"""

from dataclasses import dataclass, field
from typing import Optional, List, Dict, Any

from ..models.tracker import TrackerTemplate


# App defaults for fields the skill does not reliably provide. Sourced from the
# TrackerTemplate dataclass so they can't drift from the model.
_DEFAULT_MODULE_SPACING_M: float = TrackerTemplate.__dataclass_fields__['module_spacing_m'].default
_DEFAULT_MOTOR_GAP_M: float = TrackerTemplate.__dataclass_fields__['motor_gap_m'].default


@dataclass
class ModuleImportEntry:
    """A placeholder module from the extraction, to be resolved by the user."""
    label: str                       # brand — placeholder label and join key
    wattage: Optional[float]         # size_watts (used to pre-filter the mapping dropdown)
    model: Optional[str]             # often null
    layout_hints: Dict[str, Any] = field(default_factory=dict)  # read-only display only


@dataclass
class TemplateImportEntry:
    """A tracker template mapped into this app's ``template_data`` shape, minus
    the embedded module spec (which is resolved from ``module_ref`` later)."""
    name: str                        # template name (used to build the library key)
    module_ref: str                  # join key -> ModuleImportEntry.label
    template_data: Dict[str, Any]    # app-shape dict, WITHOUT module_spec
    modules_per_tracker: Optional[int]  # used only for invariant checks / display
    # Raw extraction motor values, for read-only context in the review dialog.
    # The skill's split_north/split_south are tracker-wide (sum to
    # modules_per_tracker); the app's motor fields are derived from them (see
    # derive_motor_fields). The drawing's motor_placement/after/in_string are
    # unreliable and intentionally NOT used for the derivation.
    raw_motor: Dict[str, Any] = field(default_factory=dict)
    # "Modules north of the motor" the derivation started from (the trustworthy
    # tracker-wide north split). The dialog seeds its one editable motor input
    # from this. None when it could not be determined.
    modules_north_of_motor: Optional[int] = None
    layout_hints: Dict[str, Any] = field(default_factory=dict)  # read-only display only
    warnings: List[str] = field(default_factory=list)
    # Informational notes (e.g. "we overrode the drawing's stated motor position").
    notes: List[str] = field(default_factory=list)
    # Motor position is always surfaced for review (see module docstring).
    motor_needs_review: bool = True


@dataclass
class InverterImportInfo:
    """Inverter is match-only; never created from extraction data."""
    name: Optional[str]
    qty: Optional[int]


@dataclass
class ProjectMetaImport:
    """Project metadata preview. Fillable fields are applied only where the
    current project's field is empty. Cross-check fields are display-only."""
    # Fillable (only when target is empty)
    customer: Optional[str] = None
    name: Optional[str] = None
    address: Optional[str] = None
    coordinates: Optional[str] = None
    city_state_zip: Optional[str] = None
    # Cross-check display only — never written (the app computes these)
    dc_capacity_kw: Optional[float] = None
    ac_capacity_kw: Optional[float] = None
    dc_ac_ratio: Optional[float] = None
    total_modules: Optional[int] = None
    total_strings: Optional[int] = None
    inverter_qty: Optional[int] = None


@dataclass
class ImportPlan:
    """Full in-memory plan describing a proposed import."""
    modules: List[ModuleImportEntry]
    templates: List[TemplateImportEntry]
    inverter: InverterImportInfo
    project_meta: ProjectMetaImport
    tracker_manufacturer: Optional[str]
    warnings: List[str] = field(default_factory=list)  # plan-level warnings


class ExtractionImportError(ValueError):
    """Raised when the input is not a structurally valid ExtractionResult."""
    pass


# ---------------------------------------------------------------------------
# Field mapping: extraction TrackerTemplate -> app template_data shape
# ---------------------------------------------------------------------------

# Extraction orientation values ("Portrait"/"Landscape") already match the app's
# ModuleOrientation vocabulary, so no translation table is needed.

_LAYOUT_HINT_MODULE_KEYS = (
    'quantity', 'string_size', 'string_quantity', 'row_spacing_ft', 'driveline_angle',
)


def derive_motor_fields(modules_north_of_motor: int, modules_per_string: int,
                        full_strings: int) -> Dict[str, Any]:
    """Translate a tracker-wide "modules north of the motor" count into the
    app's motor placement fields.

    The count is trustworthy because it is encoded in the template name and
    matches the tracker-wide split; the drawing's stated motor_placement/
    after_string/in_string fields are not used. If the count lands on a string
    boundary the motor sits *between* strings; otherwise it sits in the *middle*
    of one string and the remainder is that string's north/south split.

    Raises ValueError if the count is outside ``0..modules_per_string*full_strings``.
    """
    total = modules_per_string * full_strings
    if not (0 <= modules_north_of_motor <= total):
        raise ValueError(
            f"modules north of motor ({modules_north_of_motor}) out of range 0..{total}"
        )
    if modules_north_of_motor % modules_per_string == 0:
        return {
            'motor_placement_type': 'between_strings',
            'motor_position_after_string': modules_north_of_motor // modules_per_string,
            'motor_string_index': 1,
            'motor_split_north': 0,
            'motor_split_south': 0,
        }
    in_string = modules_north_of_motor // modules_per_string + 1
    split_n = modules_north_of_motor % modules_per_string
    return {
        'motor_placement_type': 'middle_of_string',
        'motor_position_after_string': 0,
        'motor_string_index': in_string,
        'motor_split_north': split_n,
        'motor_split_south': modules_per_string - split_n,
    }


def _template_to_app_shape(tpl: Dict[str, Any]) -> Dict[str, Any]:
    """Map the non-motor part of one extraction template into the app's
    ``template_data`` shape. Motor fields are added separately (derived).

    The returned dict deliberately omits ``module_spec`` (resolved later from
    ``module_ref``) and forces ``module_spacing_m`` to the app default.
    """
    orientation = tpl.get('orientation') or 'Portrait'
    return {
        'module_orientation': orientation,
        'modules_per_string': tpl.get('modules_per_string'),
        'strings_per_tracker': tpl.get('strings_per_tracker'),
        # Ignore skill module_spacing_m (placeholder) — use app default.
        'module_spacing_m': _DEFAULT_MODULE_SPACING_M,
        'has_motor': True,
        'motor_gap_m': _DEFAULT_MOTOR_GAP_M,
        'modules_high': tpl.get('modules_high') or 1,
        'source_point_config': None,
        'partial_string_side': 'north',
    }


def _compute_motor_fields(tpl: Dict[str, Any]):
    """Derive the app motor fields for one template from its tracker-wide north
    split. Returns ``(motor_fields, modules_north_of_motor, note)``.

    ``note`` is an informational string when the drawing's stated motor fields
    disagreed with the derivation (so the dialog can say we overrode them), or
    when derivation was not possible (fell back to the drawing's raw values).
    """
    mps = tpl.get('modules_per_string')
    spt = tpl.get('strings_per_tracker')
    split_n = tpl.get('split_north')

    fallback = {
        'motor_placement_type': tpl.get('motor_placement') or 'between_strings',
        'motor_position_after_string': tpl.get('motor_after_string') or 0,
        'motor_string_index': tpl.get('motor_in_string') or 1,
        'motor_split_north': 0,
        'motor_split_south': 0,
    }

    if mps is None or spt is None or split_n is None:
        return fallback, None, (
            "Could not derive the motor position (missing split or module counts) — "
            "review the motor fields carefully."
        )

    try:
        full_strings = int(spt)
        derived = derive_motor_fields(split_n, mps, full_strings)
    except (ValueError, TypeError):
        return fallback, None, (
            "Could not derive the motor position from the tracker split — "
            "review the motor fields carefully."
        )

    # Transparency note when the drawing's stated motor fields disagree.
    disagree = False
    stated_placement = tpl.get('motor_placement')
    if stated_placement and stated_placement != derived['motor_placement_type']:
        disagree = True
    elif (derived['motor_placement_type'] == 'between_strings'
          and tpl.get('motor_after_string') not in (None, derived['motor_position_after_string'])):
        disagree = True
    elif (derived['motor_placement_type'] == 'middle_of_string'
          and tpl.get('motor_in_string') not in (None, derived['motor_string_index'])):
        disagree = True

    note = None
    if disagree:
        south = mps * full_strings - split_n
        note = (
            f"Motor position derived from the tracker split ({split_n} north / {south} south, "
            "matching the template name); the drawing's stated motor fields differed and were "
            "not used."
        )
    return derived, split_n, note


def _check_template_invariants(tpl: Dict[str, Any]) -> List[str]:
    """Return warning strings for any failed invariant on a good template:

    * split_north + split_south == modules_per_tracker
    * modules_per_string * strings_per_tracker == modules_per_tracker

    Checks are skipped (no warning) when a needed value is absent, since a
    missing field is a separate concern from an inconsistent one.
    """
    warnings: List[str] = []

    mpt = tpl.get('modules_per_tracker')
    split_n = tpl.get('split_north')
    split_s = tpl.get('split_south')
    mps = tpl.get('modules_per_string')
    spt = tpl.get('strings_per_tracker')

    if mpt is not None and split_n is not None and split_s is not None:
        if split_n + split_s != mpt:
            warnings.append(
                f"Split check failed: split_north ({split_n}) + split_south ({split_s}) "
                f"= {split_n + split_s}, expected modules_per_tracker ({mpt})."
            )

    if mpt is not None and mps is not None and spt is not None:
        if mps * spt != mpt:
            warnings.append(
                f"Count check failed: modules_per_string ({mps}) x strings_per_tracker "
                f"({spt}) = {mps * spt}, expected modules_per_tracker ({mpt})."
            )

    return warnings


def build_import_plan(data: Dict[str, Any]) -> ImportPlan:
    """Parse a validated ``ExtractionResult`` dict into an :class:`ImportPlan`.

    Raises :class:`ExtractionImportError` when the input is not a structurally
    valid ExtractionResult (not a dict, missing ``project``, or with
    ``modules``/``tracker_templates`` of the wrong type).
    """
    if not isinstance(data, dict):
        raise ExtractionImportError("Extraction file must contain a JSON object.")

    project = data.get('project')
    if not isinstance(project, dict):
        raise ExtractionImportError(
            "Extraction file is missing the required 'project' object."
        )

    raw_modules = data.get('modules', [])
    if not isinstance(raw_modules, list):
        raise ExtractionImportError("'modules' must be a list.")

    raw_templates = data.get('tracker_templates', [])
    if not isinstance(raw_templates, list):
        raise ExtractionImportError("'tracker_templates' must be a list.")

    plan_warnings: List[str] = []

    # --- Modules -----------------------------------------------------------
    modules: List[ModuleImportEntry] = []
    module_labels = set()
    for m in raw_modules:
        if not isinstance(m, dict):
            plan_warnings.append("Skipped a module entry that was not an object.")
            continue
        label = m.get('brand')
        if not label:
            plan_warnings.append("Skipped a module entry with no 'brand' label.")
            continue
        module_labels.add(label)
        modules.append(ModuleImportEntry(
            label=label,
            wattage=m.get('size_watts'),
            model=m.get('model'),
            layout_hints={k: m.get(k) for k in _LAYOUT_HINT_MODULE_KEYS if m.get(k) is not None},
        ))

    # --- Tracker templates -------------------------------------------------
    templates: List[TemplateImportEntry] = []
    for t in raw_templates:
        if not isinstance(t, dict):
            plan_warnings.append("Skipped a tracker template entry that was not an object.")
            continue
        name = t.get('name')
        module_ref = t.get('module_ref')
        if not name or not module_ref:
            plan_warnings.append(
                "Skipped a tracker template missing required 'name' or 'module_ref'."
            )
            continue

        warnings = _check_template_invariants(t)

        # Resolve the join key to a known module label.
        if module_ref not in module_labels:
            warnings.append(
                f"module_ref '{module_ref}' has no matching module in the extraction."
            )

        template_data = _template_to_app_shape(t)
        motor_fields, north_of_motor, motor_note = _compute_motor_fields(t)
        template_data.update(motor_fields)

        templates.append(TemplateImportEntry(
            name=name,
            module_ref=module_ref,
            template_data=template_data,
            modules_per_tracker=t.get('modules_per_tracker'),
            raw_motor={
                'motor_placement': t.get('motor_placement'),
                'motor_after_string': t.get('motor_after_string'),
                'motor_in_string': t.get('motor_in_string'),
                'split_north': t.get('split_north'),
                'split_south': t.get('split_south'),
            },
            modules_north_of_motor=north_of_motor,
            layout_hints={'quantity': t.get('quantity')} if t.get('quantity') is not None else {},
            warnings=warnings,
            notes=[motor_note] if motor_note else [],
            motor_needs_review=True,
        ))

    # --- Inverter (match-only) --------------------------------------------
    inverter = InverterImportInfo(
        name=project.get('inverter'),
        qty=project.get('inverter_qty'),
    )

    # --- Project meta ------------------------------------------------------
    project_meta = ProjectMetaImport(
        customer=project.get('customer') or None,
        name=project.get('name') or None,
        address=project.get('address'),
        coordinates=project.get('coordinates'),
        city_state_zip=project.get('city_state_zip'),
        dc_capacity_kw=project.get('dc_capacity_kw'),
        ac_capacity_kw=project.get('ac_capacity_kw'),
        dc_ac_ratio=project.get('dc_ac_ratio'),
        total_modules=project.get('total_modules'),
        total_strings=project.get('total_strings'),
        inverter_qty=project.get('inverter_qty'),
    )

    return ImportPlan(
        modules=modules,
        templates=templates,
        inverter=inverter,
        project_meta=project_meta,
        tracker_manufacturer=project.get('tracker_manufacturer'),
        warnings=plan_warnings,
    )
