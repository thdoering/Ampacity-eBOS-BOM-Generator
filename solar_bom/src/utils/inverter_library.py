import sys
import json
from pathlib import Path
from typing import Dict, Optional, Set, Tuple


def get_factory_path() -> Path:
    """Return path to the factory inverter library, resolving correctly in dev and bundled modes."""
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS) / 'data' / 'inverter_library_factory.json'
    return Path(__file__).parent.parent.parent / 'data' / 'inverter_library_factory.json'


def get_user_path() -> Path:
    """Return path to the user inverter library."""
    return Path('data/inverters.json')


def _parse_hierarchical(data: dict) -> Dict[str, dict]:
    """Convert {Manufacturer: {Model: inverter_data}} to flat {key: inverter_data}."""
    flat = {}
    for manufacturer, models in data.items():
        if not isinstance(models, dict):
            continue
        for model, inverter_data in models.items():
            flat[f"{manufacturer} {model}"] = inverter_data
    return flat


def load_factory_inverters() -> Dict[str, dict]:
    """Load factory inverters. Returns flat {inverter_key: raw_dict}. Tolerates missing/empty file."""
    path = get_factory_path()
    try:
        if path.exists():
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if data:
                return _parse_hierarchical(data)
    except Exception as e:
        print(f"Warning: Could not load factory inverter library: {e}")
    return {}


def load_user_inverters() -> Dict[str, dict]:
    """Load user inverters. Returns flat {inverter_key: raw_dict}. Handles hierarchical and flat formats."""
    path = get_user_path()
    try:
        if path.exists():
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if data:
                first_value = next(iter(data.values()))
                if isinstance(first_value, dict) and not any(
                    key in first_value for key in ('manufacturer', 'model', 'inverter_type')
                ):
                    return _parse_hierarchical(data)
                else:
                    return dict(data)
    except Exception as e:
        print(f"Warning: Could not load user inverter library: {e}")
    return {}


def load_merged_inverters() -> Tuple[Dict[str, dict], Set[str]]:
    """
    Merge factory and user inverter libraries.
    Returns (merged_dict, factory_keys) where merged_dict is flat {inverter_key: raw_dict}.
    Factory wins on conflict. factory_keys is the set of keys sourced from the factory library.
    """
    user_inverters = load_user_inverters()
    factory_inverters = load_factory_inverters()

    merged = dict(user_inverters)
    merged.update(factory_inverters)  # factory wins on conflict

    return merged, set(factory_inverters.keys())


def save_user_inverters(inverters: dict, factory_keys: Set[str]) -> None:
    """Save only non-factory entries to the user file in flat format. Never writes the factory file."""
    path = get_user_path()
    path.parent.mkdir(exist_ok=True)

    data = {}
    for name, inverter in inverters.items():
        if name in factory_keys:
            continue
        inv_dict = {
            'manufacturer': inverter.manufacturer,
            'model': inverter.model,
            'inverter_type': inverter.inverter_type.value,
            'rated_power_kw': inverter.rated_power_kw,
            'max_dc_power_kw': inverter.max_dc_power_kw,
            'max_efficiency': inverter.max_efficiency,
            'mppt_channels': [ch.__dict__ for ch in inverter.mppt_channels],
            'mppt_configuration': inverter.mppt_configuration.value,
            'max_dc_voltage': inverter.max_dc_voltage,
            'startup_voltage': inverter.startup_voltage,
            'nominal_ac_voltage': inverter.nominal_ac_voltage,
            'max_ac_current': inverter.max_ac_current,
            'power_factor': inverter.power_factor,
            'dimensions_mm': list(inverter.dimensions_mm),
            'weight_kg': inverter.weight_kg,
            'ip_rating': inverter.ip_rating,
            'max_short_circuit_current': getattr(inverter, 'max_short_circuit_current', None),
            'temperature_range': list(inverter.temperature_range) if getattr(inverter, 'temperature_range', None) else None,
            'altitude_limit_m': getattr(inverter, 'altitude_limit_m', None),
            'communication_protocol': getattr(inverter, 'communication_protocol', None),
        }
        data[name] = inv_dict

    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


def deserialize_inverter_spec(name: str, specs: dict) -> 'Optional[InverterSpec]':
    """Convert a raw inverter dict to an InverterSpec. Returns None on failure."""
    from ..models.inverter import InverterSpec, MPPTChannel, MPPTConfig, InverterType
    try:
        rated_power = specs.get('rated_power_kw', specs.get('rated_power', 10.0))
        max_dc_power = specs.get('max_dc_power_kw', float(rated_power) * 1.5)
        return InverterSpec(
            manufacturer=specs.get('manufacturer', 'Unknown'),
            model=specs.get('model', 'Unknown'),
            inverter_type=InverterType(specs.get('inverter_type', 'String')),
            rated_power_kw=float(rated_power),
            max_dc_power_kw=float(max_dc_power),
            max_efficiency=float(specs.get('max_efficiency', 98.0)),
            mppt_channels=[MPPTChannel(**ch) for ch in specs.get('mppt_channels', [])],
            mppt_configuration=MPPTConfig(specs.get('mppt_configuration', 'Independent')),
            max_dc_voltage=float(specs.get('max_dc_voltage', 1500)),
            startup_voltage=float(specs.get('startup_voltage', 150)),
            nominal_ac_voltage=float(specs.get('nominal_ac_voltage', 400.0)),
            max_ac_current=float(specs.get('max_ac_current', 40.0)),
            power_factor=float(specs.get('power_factor', 0.99)),
            dimensions_mm=tuple(specs.get('dimensions_mm', (1000, 600, 300))),
            weight_kg=float(specs.get('weight_kg', 75.0)),
            ip_rating=specs.get('ip_rating', 'IP65'),
            max_short_circuit_current=specs.get('max_short_circuit_current'),
            temperature_range=tuple(specs['temperature_range']) if specs.get('temperature_range') else None,
            altitude_limit_m=specs.get('altitude_limit_m'),
            communication_protocol=specs.get('communication_protocol'),
        )
    except Exception as e:
        print(f"Warning: Failed to deserialize inverter '{name}': {e}")
        return None


def load_merged_inverter_specs() -> Tuple[Dict[str, 'InverterSpec'], Set[str]]:
    """
    Merge factory and user libraries and deserialize to InverterSpec objects.
    Returns (inverters_dict, factory_keys) where inverters_dict is {display_name: InverterSpec}.
    Factory wins on conflict.
    """
    merged_raw, factory_keys = load_merged_inverters()
    inverters = {}
    for name, specs in merged_raw.items():
        inv = deserialize_inverter_spec(name, specs)
        if inv:
            inverters[name] = inv
    return inverters, factory_keys


def is_inverter_in_factory(manufacturer: str, model: str) -> bool:
    """Return True if an inverter with the given manufacturer and model exists in the factory library."""
    factory = load_factory_inverters()
    return f"{manufacturer} {model}" in factory
