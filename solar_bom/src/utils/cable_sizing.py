"""
Cable Sizing Service for Solar eBOS BOM Generator

This module provides functions to calculate recommended cable sizes
for all four cable types in a harness assembly based on electrical load.
"""

from typing import Dict, Optional

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

def calculate_string_cable_size(module_isc: float, nec_factor: float = 1.25) -> str:
    """
    Calculate required cable size for a single string connection.
    
    String cables carry current from one string of modules to the harness
    connection point. Per NEC, we use 125% of Isc for continuous current.
    
    Args:
        module_isc: Module short circuit current in amperes
        nec_factor: NEC safety factor (default 1.25 for continuous current)
        
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

def calculate_harness_cable_size(num_strings: int, module_isc: float, nec_factor: float = 1.25) -> str:
    """
    Calculate required cable size for harness cables.
    
    Harness cables combine current from multiple strings and carry it
    to the extender connection point.
    
    Args:
        num_strings: Number of strings combined in the harness
        module_isc: Module short circuit current in amperes
        nec_factor: NEC safety factor (default 1.25 for continuous current)
        
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

def calculate_extender_cable_size(num_strings: int, module_isc: float, nec_factor: float = 1.25) -> str:
    """
    Calculate required cable size for extender cables.
    
    Extender cables carry the combined current from the harness to
    the whip connection point. They carry the same current as harness cables.
    
    Args:
        num_strings: Number of strings in the harness assembly
        module_isc: Module short circuit current in amperes
        nec_factor: NEC safety factor (default 1.25 for continuous current)
        
    Returns:
        str: Recommended AWG cable size
    """
    # Extender carries same current as harness
    return calculate_harness_cable_size(num_strings, module_isc, nec_factor)

def calculate_whip_cable_size(num_strings: int, module_isc: float, nec_factor: float = 1.25) -> str:
    """
    Calculate required cable size for whip cables.
    
    Whip cables make the final connection from the extender to the
    device (combiner box or inverter). They carry the same current
    as harness and extender cables.
    
    Args:
        num_strings: Number of strings in the harness assembly
        module_isc: Module short circuit current in amperes
        nec_factor: NEC safety factor (default 1.25 for continuous current)
        
    Returns:
        str: Recommended AWG cable size
    """
    # Whip carries same current as harness and extender
    return calculate_harness_cable_size(num_strings, module_isc, nec_factor)

def calculate_all_cable_sizes(num_strings: int, module_isc: float, nec_factor: float = 1.25) -> Dict[str, str]:
    """
    Calculate recommended cable sizes for all components of a harness assembly.
    
    This function calculates appropriate cable sizes for string, harness,
    extender, and whip cables based on the electrical load.
    
    Args:
        num_strings: Number of strings in the harness assembly
        module_isc: Module short circuit current in amperes
        nec_factor: NEC safety factor (default 1.25 for continuous current)
        
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

def validate_cable_size_for_current(cable_size: str, current: float, nec_factor: float = 1.25) -> bool:
    """
    Validate if a cable size is adequate for the given current.
    
    Args:
        cable_size: AWG cable size string
        current: Base current in amperes (before NEC factor)
        nec_factor: NEC safety factor (default 1.25)
        
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