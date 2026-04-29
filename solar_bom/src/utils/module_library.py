import sys
import json
from pathlib import Path
from typing import Dict, Set, Tuple


def get_factory_path() -> Path:
    """Return path to the factory module library, resolving correctly in dev and bundled modes."""
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS) / 'data' / 'module_library_factory.json'
    return Path(__file__).parent.parent.parent / 'data' / 'module_library_factory.json'


def get_user_path() -> Path:
    """Return path to the user module library."""
    return Path('data/module_templates.json')


def _parse_hierarchical(data: dict) -> Dict[str, dict]:
    """Convert {Manufacturer: {Model: module_data}} to flat {key: module_data}."""
    flat = {}
    for manufacturer, models in data.items():
        if not isinstance(models, dict):
            continue
        for model, module_data in models.items():
            flat[f"{manufacturer} {model}"] = module_data
    return flat


def load_factory_modules() -> Dict[str, dict]:
    """Load factory modules. Returns flat {module_key: module_data}. Tolerates missing/empty file."""
    path = get_factory_path()
    try:
        if path.exists():
            with open(path, 'r') as f:
                data = json.load(f)
            if data:
                return _parse_hierarchical(data)
    except Exception as e:
        print(f"Warning: Could not load factory module library: {e}")
    return {}


def load_user_modules() -> Dict[str, dict]:
    """Load user modules. Returns flat {module_key: module_data}. Handles hierarchical and old flat formats."""
    path = get_user_path()
    try:
        if path.exists():
            with open(path, 'r') as f:
                data = json.load(f)
            if data:
                first_value = next(iter(data.values()))
                if isinstance(first_value, dict) and not any(
                    key in first_value for key in ('manufacturer', 'model', 'type')
                ):
                    # Hierarchical format
                    return _parse_hierarchical(data)
                else:
                    # Old flat format
                    return dict(data)
    except Exception as e:
        print(f"Warning: Could not load user module library: {e}")
    return {}


def load_merged_modules() -> Tuple[Dict[str, dict], Set[str]]:
    """
    Merge factory and user module libraries.
    Returns (merged_dict, factory_keys) where merged_dict is flat {module_key: module_data}.
    Factory wins on conflict. factory_keys is the set of keys sourced from the factory library.
    """
    user_modules = load_user_modules()
    factory_modules = load_factory_modules()

    merged = dict(user_modules)
    merged.update(factory_modules)  # factory wins on conflict

    return merged, set(factory_modules.keys())


def save_user_modules(modules: dict, factory_keys: Set[str]) -> None:
    """Save only non-factory entries to the user file in hierarchical format. Never writes the factory file."""
    path = get_user_path()
    path.parent.mkdir(exist_ok=True)

    hierarchical_data = {}
    for key, module in modules.items():
        if key in factory_keys:
            continue
        manufacturer = module.manufacturer
        model = module.model
        if manufacturer not in hierarchical_data:
            hierarchical_data[manufacturer] = {}
        hierarchical_data[manufacturer][model] = {
            **module.__dict__,
            'type': module.type.value,
            'default_orientation': module.default_orientation.value,
        }

    with open(path, 'w') as f:
        json.dump(hierarchical_data, f, indent=2)


def is_module_in_factory(manufacturer: str, model: str) -> bool:
    """Return True if a module with the given manufacturer and model exists in the factory library."""
    factory = load_factory_modules()
    return f"{manufacturer} {model}" in factory
