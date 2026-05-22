"""
Cable Sizing Service for Solar eBOS BOM Generator

This module provides functions to calculate recommended cable sizes
for all four cable types in a harness assembly based on electrical load.
"""

from typing import Dict, Optional, List, Tuple
import json
import os

# NEC Table 310.15(B)(16) - Ampacity for 90°C rated cables (THWN-2, XHHW-2)
# Standard sizes used in solar installations
CABLE_AMPACITY_90C = {
    "10 AWG": 40,
    "8 AWG": 55,
    "6 AWG": 75,
    "4 AWG": 95,
    "2 AWG": 130,
    "1/0 AWG": 170,
    "2/0 AWG": 195,
    "4/0 AWG": 260
}

# Order of cable sizes for iteration (smallest to largest)
CABLE_SIZE_ORDER = [
    "10 AWG", "8 AWG", "6 AWG", 
    "4 AWG", "2 AWG", "1/0 AWG", "2/0 AWG", "4/0 AWG"
]

# Extended order including kcmil sizes (for feeders/homeruns)
CABLE_SIZE_ORDER_EXTENDED = [
    "10 AWG", "8 AWG", "6 AWG", "4 AWG", "3 AWG", "2 AWG", "1 AWG",
    "1/0 AWG", "2/0 AWG", "3/0 AWG", "4/0 AWG",
    "250 kcmil", "300 kcmil", "350 kcmil", "400 kcmil", "500 kcmil",
    "600 kcmil", "700 kcmil", "750 kcmil", "800 kcmil", "900 kcmil", "1000 kcmil"
]

# Cached NEC table data
_nec_table_cache = None

def _load_nec_table() -> dict:
    """Load NEC Table 310.16 data from JSON file, with caching."""
    global _nec_table_cache
    if _nec_table_cache is not None:
        return _nec_table_cache
    
    data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'data')
    table_path = os.path.join(data_dir, 'nec_table_310_16.json')
    
    try:
        with open(table_path, 'r') as f:
            _nec_table_cache = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"Warning: Could not load NEC Table 310.16: {e}")
        _nec_table_cache = {"copper": {}, "aluminum": {}}
    
    return _nec_table_cache


def get_nec_ampacity(cable_size: str, material: str = "copper", temp_rating: str = "75C") -> float:
    """
    Look up ampacity from NEC Table 310.16.
    
    Args:
        cable_size: Cable size string (e.g., "10 AWG", "500 kcmil")
        material: "copper" or "aluminum"
        temp_rating: "60C", "75C", or "90C"
        
    Returns:
        float: Ampacity in amperes, or 0 if not found
    """
    table = _load_nec_table()
    material_data = table.get(material, {})
    size_data = material_data.get(cable_size, {})
    return size_data.get(temp_rating, 0)


def get_available_sizes(material: str = "copper") -> List[str]:
    """
    Get the ordered list of available cable sizes for a given material.
    
    Args:
        material: "copper" or "aluminum"
        
    Returns:
        List of cable size strings in order from smallest to largest
    """
    table = _load_nec_table()
    material_data = table.get(material, {})
    # Return sizes in the extended order, filtered to only those in the table
    return [s for s in CABLE_SIZE_ORDER_EXTENDED if s in material_data]


def recommend_cable_size(current: float, material: str = "copper", temp_rating: str = "75C") -> str:
    """
    Recommend the smallest cable size adequate for the given current.
    
    Args:
        current: Required ampacity (already includes any NEC factors)
        material: "copper" or "aluminum"
        temp_rating: "60C", "75C", or "90C"
        
    Returns:
        str: Recommended cable size, or largest available if none adequate
    """
    available = get_available_sizes(material)
    for size in available:
        ampacity = get_nec_ampacity(size, material, temp_rating)
        if ampacity >= current:
            return size
    
    # Return largest available
    return available[-1] if available else "1000 kcmil"


def recommend_lv_cable_sizes(num_strings: int, module_isc: float, 
                              nec_factor: float = 1.56, temp_rating: str = "75C") -> Dict[str, str]:
    """
    Recommend cable sizes for LV collection (harness, extender, whip).
    Always uses copper. Extender and whip are floored at the harness size.
    
    Args:
        num_strings: Number of strings in the harness
        module_isc: Module short circuit current in amperes
        nec_factor: NEC safety factor (default 1.56)
        temp_rating: Temperature rating column to use
        
    Returns:
        Dict with 'harness', 'extender', 'whip' cable sizes
    """
    required_current = num_strings * module_isc * nec_factor
    harness_size = recommend_cable_size(required_current, "copper", temp_rating)
    
    # Extender and whip are at least as large as harness
    return {
        'harness': harness_size,
        'extender': harness_size,
        'whip': harness_size
    }


def recommend_dc_feeder_size(breaker_rating: float, material: str = "aluminum", 
                              temp_rating: str = "75C") -> str:
    """
    Recommend DC feeder cable size based on combiner breaker rating.
    Cable ampacity must be >= breaker rating per NEC 240.4.
    
    Args:
        breaker_rating: Combiner output breaker rating in amps
        material: "copper" or "aluminum"
        temp_rating: Temperature rating column to use
        
    Returns:
        str: Recommended cable size
    """
    return recommend_cable_size(breaker_rating, material, temp_rating)


def get_block_dc_ocpd_rating(block, device_configurator=None) -> Optional[float]:
    """Return the OCPD rating (A) driving the DC feeder for a block, or None if unknown.

    Prefers the largest combiner output breaker from device_configurator; falls back to
    inverter MPPT channel max_input_current × 1.25 (NEC continuous load factor).
    """
    breaker = None

    if device_configurator is not None:
        combiner_configs = getattr(device_configurator, 'combiner_configs', {})
        block_id = getattr(block, 'block_id', None)
        for cfg in combiner_configs.values():
            if getattr(cfg, 'block_id', None) == block_id:
                size = cfg.get_display_breaker_size()
                breaker = max(breaker, size) if breaker is not None else size

    if breaker is None:
        inv = getattr(block, 'inverter', None)
        if inv is not None:
            channels = getattr(inv, 'mppt_channels', [])
            total_dc = sum(getattr(ch, 'max_input_current', 0.0) for ch in channels)
            breaker = total_dc * 1.25

    return float(breaker) if breaker else None


def recommend_block_dc_feeder_size(block, device_configurator=None) -> str:
    """Recommend DC feeder cable size for a block (legacy — aluminum 75 °C).

    Use autosize_dc_feeder_for_block() for NEC-2023 project-aware sizing.
    """
    ocpd = get_block_dc_ocpd_rating(block, device_configurator)
    if not ocpd:
        return '4/0 AWG'
    return recommend_dc_feeder_size(ocpd, material='aluminum', temp_rating='75C')


def recommend_ac_homerun_size(max_ac_current: float, material: str = "aluminum",
                               temp_rating: str = "75C") -> str:
    """
    Recommend AC homerun cable size based on inverter max output current.
    Uses 1.25 NEC continuous load factor on the AC side.
    
    Args:
        max_ac_current: Inverter maximum AC output current in amps
        material: "copper" or "aluminum"
        temp_rating: Temperature rating column to use
        
    Returns:
        str: Recommended cable size
    """
    required_current = max_ac_current * 1.25
    return recommend_cable_size(required_current, material, temp_rating)

def calculate_string_cable_size(module_isc: float, nec_factor: float = 1.56) -> str:
    """
    Calculate required cable size for a single string connection.
    
    String cables carry current from one string of modules to the harness
    connection point. Per NEC, we use 125% of Isc for continuous current.
    
    Args:
        module_isc: Module short circuit current in amperes
        nec_factor: NEC safety factor (default 1.56 for continuous current)
        
    Returns:
        str: Recommended AWG cable size
    """
    # Calculate current with NEC factor
    current = module_isc * nec_factor
    
    # Find appropriate cable size
    for cable_size in CABLE_SIZE_ORDER:
        if CABLE_AMPACITY_90C[cable_size] >= current:
            return cable_size
    
    # If current exceeds all standard sizes, return largest
    return "4/0 AWG"

def calculate_harness_cable_size(num_strings: int, module_isc: float, nec_factor: float = 1.56) -> str:
    """
    Calculate required cable size for harness cables.
    
    Harness cables combine current from multiple strings and carry it
    to the extender connection point.
    
    Args:
        num_strings: Number of strings combined in the harness
        module_isc: Module short circuit current in amperes
        nec_factor: NEC safety factor (default 1.56 for continuous current)
        
    Returns:
        str: Recommended AWG cable size
    """
    # Calculate combined current with NEC factor
    current = num_strings * module_isc * nec_factor
    
    # Find appropriate cable size
    for cable_size in CABLE_SIZE_ORDER:
        if CABLE_AMPACITY_90C[cable_size] >= current:
            return cable_size
    
    return "4/0 AWG"

def calculate_extender_cable_size(num_strings: int, module_isc: float, nec_factor: float = 1.56) -> str:
    """
    Calculate required cable size for extender cables.
    
    Extender cables carry the combined current from the harness to
    the whip connection point. They carry the same current as harness cables.
    
    Args:
        num_strings: Number of strings in the harness assembly
        module_isc: Module short circuit current in amperes
        nec_factor: NEC safety factor (default 1.56 for continuous current)
        
    Returns:
        str: Recommended AWG cable size
    """
    # Extender carries same current as harness
    return calculate_harness_cable_size(num_strings, module_isc, nec_factor)

def calculate_whip_cable_size(num_strings: int, module_isc: float, nec_factor: float = 1.56) -> str:
    """
    Calculate required cable size for whip cables.
    
    Whip cables make the final connection from the extender to the
    device (combiner box or inverter). They carry the same current
    as harness and extender cables.
    
    Args:
        num_strings: Number of strings in the harness assembly
        module_isc: Module short circuit current in amperes
        nec_factor: NEC safety factor (default 1.56 for continuous current)
        
    Returns:
        str: Recommended AWG cable size
    """
    # Whip carries same current as harness and extender
    return calculate_harness_cable_size(num_strings, module_isc, nec_factor)

def calculate_all_cable_sizes(num_strings: int, module_isc: float, nec_factor: float = 1.56) -> Dict[str, str]:
    """
    Calculate recommended cable sizes for all components of a harness assembly.
    
    This function calculates appropriate cable sizes for string, harness,
    extender, and whip cables based on the electrical load.
    
    Args:
        num_strings: Number of strings in the harness assembly
        module_isc: Module short circuit current in amperes
        nec_factor: NEC safety factor (default 1.56 for continuous current)
        
    Returns:
        Dict containing recommended sizes for each cable type:
            - 'string': AWG size for string cables
            - 'harness': AWG size for harness cables
            - 'extender': AWG size for extender cables
            - 'whip': AWG size for whip cables
    """
    return {
        'string': calculate_string_cable_size(module_isc, nec_factor),
        'harness': calculate_harness_cable_size(num_strings, module_isc, nec_factor),
        'extender': calculate_extender_cable_size(num_strings, module_isc, nec_factor),
        'whip': calculate_whip_cable_size(num_strings, module_isc, nec_factor)
    }

def get_cable_ampacity(cable_size: str) -> float:
    """
    Get the ampacity rating for a given cable size.
    
    Args:
        cable_size: AWG cable size string
        
    Returns:
        float: Ampacity in amperes at 90°C
    """
    return CABLE_AMPACITY_90C.get(cable_size, 0)

def validate_cable_size_for_current(cable_size: str, current: float, nec_factor: float = 1.56) -> bool:
    """
    Validate if a cable size is adequate for the given current.
    
    Args:
        cable_size: AWG cable size string
        current: Base current in amperes (before NEC factor)
        nec_factor: NEC safety factor (default 1.56)
        
    Returns:
        bool: True if cable size is adequate, False otherwise
    """
    required_ampacity = current * nec_factor
    cable_ampacity = get_cable_ampacity(cable_size)
    return cable_ampacity >= required_ampacity

def get_next_larger_cable_size(cable_size: str) -> Optional[str]:
    """
    Get the next larger standard cable size.
    
    Args:
        cable_size: Current AWG cable size
        
    Returns:
        str or None: Next larger size, or None if already at maximum
    """
    try:
        current_index = CABLE_SIZE_ORDER.index(cable_size)
        if current_index < len(CABLE_SIZE_ORDER) - 1:
            return CABLE_SIZE_ORDER[current_index + 1]
    except ValueError:
        pass
    return None

def get_cable_size_index(cable_size: str) -> int:
    """
    Get the index of a cable size for comparison purposes.
    Lower index means smaller cable.
    
    Args:
        cable_size: AWG cable size string
        
    Returns:
        int: Index in size order, or -1 if not found
    """
    try:
        return CABLE_SIZE_ORDER.index(cable_size)
    except ValueError:
        return -1

def is_cable_size_larger(size1: str, size2: str) -> bool:
    """
    Check if size1 is larger than size2.
    
    Args:
        size1: First AWG cable size
        size2: Second AWG cable size
        
    Returns:
        bool: True if size1 is larger than size2
    """
    idx1 = get_cable_size_index(size1)
    idx2 = get_cable_size_index(size2)
    
    if idx1 == -1 or idx2 == -1:
        return False
    
    return idx1 > idx2

def calculate_fuse_size(total_current: float) -> int:
    """
    Calculate required fuse size based on total current.
    
    Args:
        total_current: Total current including NEC factor
        
    Returns:
        int: Recommended fuse size in amperes
    """
    # Standard fuse sizes per NEC
    FUSE_SIZES = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 60, 70, 80, 90, 100]
    
    # Find the smallest fuse that can handle the current
    for size in FUSE_SIZES:
        if size >= total_current:
            return size
    
    # If current exceeds standard sizes, return maximum
    return 100


def recommend_trunk_cable_size(num_strings: int, module_isc: float, 
                                nec_factor: float = 1.56,
                                material: str = "copper",
                                temp_rating: str = "75C") -> str:
    """
    Recommend cable size for a trunk bus segment.
    
    The trunk bus carries the combined current of all strings in the LBD block.
    
    Args:
        num_strings: Total number of strings on the trunk bus segment
        module_isc: Module short circuit current in amperes
        nec_factor: NEC safety factor (default 1.56)
        material: "copper" or "aluminum"
        temp_rating: Temperature rating column to use
        
    Returns:
        str: Recommended cable size string
    """
    required_current = num_strings * module_isc * nec_factor
    return recommend_cable_size(required_current, material, temp_rating)


# Standard LBD sizes in amperes (250A to 500A in 50A increments)
LBD_SIZES = [250, 300, 350, 400, 450, 500]


def select_lbd_size(num_strings: int, module_isc: float,
                     nec_factor: float = 1.56) -> int:
    """
    Auto-select the smallest LBD (Load Break Disconnect) rating
    that can handle the block's current.

    Args:
        num_strings: Number of strings in the LBD block
        module_isc: Module short circuit current in amperes
        nec_factor: NEC safety factor (default 1.56)

    Returns:
        int: LBD ampere rating (250, 300, 350, 400, 450, or 500)
    """
    required_amps = num_strings * module_isc * nec_factor

    for size in LBD_SIZES:
        if size >= required_amps:
            return size

    # If current exceeds all standard sizes, return largest
    return LBD_SIZES[-1]


# ---------------------------------------------------------------------------
# NEC 2023 enhanced sizing — new data file caches
# ---------------------------------------------------------------------------

_nec_table_310_17_cache = None
_ambient_correction_cache = None
_ccc_adjustment_cache = None
_insulation_types_cache = None
_chapter9_table8_cache = None


def _data_path(filename: str) -> str:
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'data', filename)


def _load_table_310_17() -> dict:
    global _nec_table_310_17_cache
    if _nec_table_310_17_cache is None:
        try:
            with open(_data_path('nec_table_310_17.json'), 'r') as f:
                _nec_table_310_17_cache = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError) as e:
            print(f"Warning: Could not load NEC Table 310.17: {e}")
            _nec_table_310_17_cache = {"copper": {}, "aluminum": {}}
    return _nec_table_310_17_cache


def _load_ambient_correction() -> dict:
    global _ambient_correction_cache
    if _ambient_correction_cache is None:
        try:
            with open(_data_path('nec_ambient_correction.json'), 'r') as f:
                _ambient_correction_cache = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError) as e:
            print(f"Warning: Could not load NEC ambient correction table: {e}")
            _ambient_correction_cache = {}
    return _ambient_correction_cache


def _load_ccc_adjustment() -> dict:
    global _ccc_adjustment_cache
    if _ccc_adjustment_cache is None:
        try:
            with open(_data_path('nec_ccc_adjustment.json'), 'r') as f:
                _ccc_adjustment_cache = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError) as e:
            print(f"Warning: Could not load NEC CCC adjustment table: {e}")
            _ccc_adjustment_cache = {"bins": []}
    return _ccc_adjustment_cache


def _load_insulation_types() -> dict:
    global _insulation_types_cache
    if _insulation_types_cache is None:
        try:
            with open(_data_path('insulation_types.json'), 'r') as f:
                _insulation_types_cache = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError) as e:
            print(f"Warning: Could not load insulation types: {e}")
            _insulation_types_cache = {}
    return _insulation_types_cache


def _load_chapter9_table8() -> dict:
    global _chapter9_table8_cache
    if _chapter9_table8_cache is None:
        try:
            with open(_data_path('nec_chapter_9_table_8.json'), 'r') as f:
                _chapter9_table8_cache = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError) as e:
            print(f"Warning: Could not load NEC Chapter 9 Table 8: {e}")
            _chapter9_table8_cache = {"copper": {}, "aluminum": {}}
    return _chapter9_table8_cache


# ---------------------------------------------------------------------------
# Public lookup functions
# ---------------------------------------------------------------------------

def get_base_ampacity(gauge: str, material: str, temp_rating_c: int,
                      installation_method: str) -> float:
    """Base ampacity from NEC 310.16 (conduit/buried) or 310.17 (free_air)."""
    temp_key = f"{temp_rating_c}C"
    if installation_method == 'free_air':
        table = _load_table_310_17()
    else:
        table = _load_nec_table()
    return float(table.get(material, {}).get(gauge, {}).get(temp_key, 0))


def get_ambient_correction(ambient_c: float, conductor_temp_rating_c: int) -> float:
    """Correction factor from NEC 2023 Table 310.15(B)(1)(a)."""
    temp_key = f"{conductor_temp_rating_c}C"
    bins = _load_ambient_correction().get(temp_key, [])
    if not bins:
        return 1.0
    # Clamp below-table ambients to the first bin
    if ambient_c <= bins[0]['min_c']:
        return bins[0]['factor']
    for b in bins:
        if b['min_c'] <= ambient_c <= b['max_c']:
            return b['factor']
    return 0.0  # above table max — conductor cannot operate at this ambient


def get_ccc_adjustment(num_ccc: int) -> float:
    """Adjustment factor from NEC 2023 Table 310.15(C)(1). Returns 1.0 for ≤3 CCCs."""
    for b in _load_ccc_adjustment().get('bins', []):
        if b['min_ccc'] <= num_ccc <= b['max_ccc']:
            return b['factor']
    return 0.35  # fallback for very large counts


def get_termination_cap_ampacity(gauge: str, material: str, termination_temp_c: int) -> float:
    """Termination ampacity cap per NEC 110.14(C). Always uses Table 310.16."""
    temp_key = f"{termination_temp_c}C"
    table = _load_nec_table()
    return float(table.get(material, {}).get(gauge, {}).get(temp_key, 0))


def get_required_ampacity(isc_total_a: float, ocpd_rating_a: float) -> Tuple[float, str]:
    """Required ampacity: max(Isc × 1.5625, OCPD). Returns (value, source_label)."""
    nec_690_8 = isc_total_a * 1.5625
    if ocpd_rating_a >= nec_690_8:
        return ocpd_rating_a, "OCPD rating (NEC 240.4)"
    return nec_690_8, "NEC 690.8 (Isc × 1.5625)"


def get_voltage_drop_pct(current_a: float, one_way_length_ft: float, gauge: str,
                          material: str, source_voltage: float) -> float:
    """DC voltage drop percentage: 2 × I × L × R_per_kft / 1000 / V × 100."""
    r_per_kft = _load_chapter9_table8().get(material, {}).get(gauge, 0.0)
    if r_per_kft == 0 or source_voltage == 0:
        return 0.0
    return (2.0 * current_a * one_way_length_ft * r_per_kft) / 1000.0 / source_voltage * 100.0


def autosize_conductor(
    isc_total_a: float,
    ocpd_rating_a: float,
    material: str,
    insulation_type: str,
    installation_method: str,
    ambient_c: float,
    ccc_count: int,
    termination_temp_c: int,
    one_way_length_ft: float,
    source_voltage: float,
    vd_target_pct: float,
) -> dict:
    """
    Select the smallest standard gauge satisfying both ampacity (NEC 690.8 /
    110.14(C) / ambient+CCC derating) and voltage drop.

    Returns a structured breakdown dict suitable for display and audit.
    If one_way_length_ft <= 0, the VD check is skipped (vd_passes=True).
    """
    insulation_data = _load_insulation_types().get(insulation_type, {})
    conductor_temp_c = insulation_data.get('temp_rating_c', 90)
    temp_key = f"{conductor_temp_c}C"

    if installation_method == 'free_air':
        table_label = f"NEC 2023 Table 310.17, {conductor_temp_c}°C"
    else:
        table_label = f"NEC 2023 Table 310.16, {conductor_temp_c}°C"

    # NEC 110.14(C) termination cap: PV Wire in free air uses 90C-rated MC4 connectors,
    # so the cap does not apply — termination and conductor are both rated 90C.
    pv_wire_free_air = (insulation_type == 'PV Wire' and installation_method == 'free_air')
    if pv_wire_free_air:
        term_source = "N/A - PV Wire free air (MC4 connectors rated 90C, NEC 110.14(C) cap not applied)"
    else:
        term_source = f"NEC 110.14(C), {termination_temp_c}°C terminals"
    required_a, req_source = get_required_ampacity(isc_total_a, ocpd_rating_a)

    available = get_available_sizes(material)
    if not available:
        available = CABLE_SIZE_ORDER_EXTENDED

    def _calc_for(gauge):
        base = get_base_ampacity(gauge, material, conductor_temp_c, installation_method)
        af = get_ambient_correction(ambient_c, conductor_temp_c)
        cf = 1.0 if installation_method == 'free_air' else get_ccc_adjustment(ccc_count)
        adj = base * af * cf
        if pv_wire_free_air:
            tc = adj  # no termination cap: MC4 connectors are 90C-rated
        else:
            tc = get_termination_cap_ampacity(gauge, material, termination_temp_c)
        final = min(adj, tc)
        return base, af, cf, adj, tc, final

    def _vd(gauge):
        if one_way_length_ft <= 0 or source_voltage <= 0:
            return 0.0
        return get_voltage_drop_pct(required_a, one_way_length_ft, gauge, material, source_voltage)

    # Pass 1: find minimum gauge that satisfies ampacity
    amp_idx = None
    for i, gauge in enumerate(available):
        _, _, _, _, _, final = _calc_for(gauge)
        if final >= required_a:
            amp_idx = i
            break

    def _build_result(gauge, amp_passes, vd_pct, vd_passes, binding):
        base, af, cf, adj, tc, final = _calc_for(gauge)
        return {
            "gauge": gauge,
            "material": material,
            "installation_method": installation_method,
            "insulation_type": insulation_type,
            "conductor_temp_rating_c": conductor_temp_c,
            "termination_temp_rating_c": termination_temp_c,
            "base_ampacity": base,
            "base_ampacity_source": table_label,
            "ambient_temp_c": ambient_c,
            "ambient_correction": round(af, 4),
            "ccc_count": ccc_count,
            "ccc_adjustment": round(cf, 4),
            "adjusted_ampacity": round(adj, 2),
            "termination_capped_ampacity": round(tc, 2),
            "termination_cap_source": term_source,
            "final_ampacity": round(final, 2),
            "required_ampacity": round(required_a, 2),
            "required_ampacity_source": req_source,
            "ampacity_passes": amp_passes,
            "vd_pct": round(vd_pct, 3),
            "vd_target_pct": vd_target_pct,
            "vd_passes": vd_passes,
            "binding_constraint": binding,
        }

    if amp_idx is None:
        # No gauge satisfies ampacity — return largest with failure noted
        gauge = available[-1]
        vd = _vd(gauge)
        return _build_result(gauge, False, vd, vd <= vd_target_pct or one_way_length_ft <= 0, "ampacity")

    # Pass 2: starting at amp_idx, find first gauge that also satisfies VD
    for idx in range(amp_idx, len(available)):
        gauge = available[idx]
        vd = _vd(gauge)
        vd_ok = (one_way_length_ft <= 0) or (vd <= vd_target_pct)
        if vd_ok:
            binding = "voltage_drop" if idx > amp_idx else "ampacity"
            return _build_result(gauge, True, vd, True, binding)

    # All gauges satisfy ampacity but VD still fails on largest
    gauge = available[-1]
    vd = _vd(gauge)
    return _build_result(gauge, True, vd, False, "voltage_drop")


# ---------------------------------------------------------------------------
# Project-aware wrappers — pull per-cable-type settings from wire_sizing_settings
# ---------------------------------------------------------------------------

def autosize_harness_for_block(
    num_strings: int,
    module_isc: float,
    wire_sizing_settings: dict,
    cable_type: str = 'harness',
    one_way_length_ft: float = 0.0,
    source_voltage: float = 0.0,
) -> dict:
    """
    Autosize a LV cable (harness/extender/whip) using project wire sizing settings.
    Passes num_strings * module_isc as isc_total_a; NEC 690.8 factor applied internally.
    """
    ambient_c = wire_sizing_settings.get('ambient_temp_c', 30)
    per = wire_sizing_settings.get('per_cable_type', {}).get(cable_type, {})
    result = autosize_conductor(
        isc_total_a=num_strings * module_isc,
        ocpd_rating_a=0.0,
        material=per.get('material', 'copper'),
        insulation_type=per.get('insulation_type', 'PV Wire'),
        installation_method=per.get('installation_method', 'free_air'),
        ambient_c=ambient_c,
        ccc_count=per.get('circuits_sharing_raceway', 1),
        termination_temp_c=per.get('termination_temp_c', 90),
        one_way_length_ft=one_way_length_ft,
        source_voltage=source_voltage,
        vd_target_pct=per.get('vd_target_pct', 2.0),
    )
    print(
        f"[DBG autosize_harness] {cable_type} {num_strings}-str | "
        f"isc={module_isc:.2f}A req={result['required_ampacity']:.1f}A | "
        f"install={result['installation_method']} insulation={result['insulation_type']} "
        f"mat={result['material']} term={result['termination_temp_rating_c']}C | "
        f"base={result['base_ampacity']:.0f}A amb*{result['ambient_correction']:.3f} "
        f"ccc*{result['ccc_adjustment']:.3f} adj={result['adjusted_ampacity']:.1f}A "
        f"term_cap={result['termination_capped_ampacity']:.0f}A "
        f"final={result['final_ampacity']:.0f}A -> {result['gauge']} "
        f"({'PASS' if result['ampacity_passes'] else 'FAIL'})"
    )
    return result


def autosize_dc_feeder_for_block(
    ocpd_rating_a: float,
    wire_sizing_settings: dict,
    one_way_length_ft: float = 0.0,
    source_voltage: float = 0.0,
) -> dict:
    """
    Autosize a DC feeder cable using project wire sizing settings.
    ocpd_rating_a is the combiner output breaker rating.
    """
    ambient_c = wire_sizing_settings.get('ambient_temp_c', 30)
    per = wire_sizing_settings.get('per_cable_type', {}).get('dc_feeder', {})
    return autosize_conductor(
        isc_total_a=0.0,
        ocpd_rating_a=ocpd_rating_a,
        material=per.get('material', 'copper'),
        insulation_type=per.get('insulation_type', 'PV Wire'),
        installation_method=per.get('installation_method', 'conduit'),
        ambient_c=ambient_c,
        ccc_count=per.get('circuits_sharing_raceway', 1),
        termination_temp_c=per.get('termination_temp_c', 90),
        one_way_length_ft=one_way_length_ft,
        source_voltage=source_voltage,
        vd_target_pct=per.get('vd_target_pct', 2.0),
    )


def autosize_ac_homerun_for_block(
    max_ac_current_a: float,
    wire_sizing_settings: dict,
    one_way_length_ft: float = 0.0,
    source_voltage: float = 0.0,
) -> dict:
    """
    Autosize an AC homerun cable using project wire sizing settings.
    NEC 210.20 continuous load factor (×1.25) is applied internally via ocpd_rating_a.
    """
    ambient_c = wire_sizing_settings.get('ambient_temp_c', 30)
    per = wire_sizing_settings.get('per_cable_type', {}).get('ac_homerun', {})
    return autosize_conductor(
        isc_total_a=0.0,
        ocpd_rating_a=max_ac_current_a * 1.25,
        material=per.get('material', 'aluminum'),
        insulation_type=per.get('insulation_type', 'XHHW-2'),
        installation_method=per.get('installation_method', 'conduit'),
        ambient_c=ambient_c,
        ccc_count=per.get('circuits_sharing_raceway', 1),
        termination_temp_c=per.get('termination_temp_c', 90),
        one_way_length_ft=one_way_length_ft,
        source_voltage=source_voltage,
        vd_target_pct=per.get('vd_target_pct', 2.0),
    )