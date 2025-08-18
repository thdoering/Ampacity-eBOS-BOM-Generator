from typing import List, Dict, Tuple
from math import pi, sqrt

def voltage_drop(
    current: float,
    length: float,
    conductor_size: float,
    temperature: float = 25.0,
    material: str = "copper"
) -> float:
    """
    Calculate voltage drop in a DC circuit.
    
    Args:
        current: Current in amperes
        length: One-way length of conductor in meters
        conductor_size: Conductor cross-sectional area in mm²
        temperature: Conductor temperature in Celsius (default 25°C)
        material: Conductor material ("copper" or "aluminum")
        
    Returns:
        float: Voltage drop in volts
    """
    # Resistivity at 20°C (ohm·mm²/m)
    resistivity_20C = {
        "copper": 0.0172,
        "aluminum": 0.0282
    }
    
    # Temperature coefficient (/°C)
    temp_coeff = {
        "copper": 0.00393,
        "aluminum": 0.00403
    }
    
    if material.lower() not in resistivity_20C:
        raise ValueError("Material must be 'copper' or 'aluminum'")
    
    # Calculate resistivity at operating temperature
    base_resistivity = resistivity_20C[material.lower()]
    alpha = temp_coeff[material.lower()]
    resistivity = base_resistivity * (1 + alpha * (temperature - 20))
    
    # Calculate resistance (length * 2 for round trip)
    resistance = (resistivity * length * 2) / conductor_size
    
    # Calculate voltage drop
    return current * resistance

def power_loss(voltage_drop: float, current: float) -> float:
    """
    Calculate power loss in watts.
    
    Args:
        voltage_drop: Voltage drop in volts
        current: Current in amperes
        
    Returns:
        float: Power loss in watts
    """
    return voltage_drop * current

def conductor_ampacity(
    size_mm2: float,
    insulation_temp: float = 90.0,
    ambient_temp: float = 30.0,
    material: str = "copper"
) -> float:
    """
    Calculate conductor ampacity based on NEC guidelines.
    This is a simplified calculation - actual values should be looked up in NEC tables.
    
    Args:
        size_mm2: Conductor size in mm²
        insulation_temp: Maximum insulation temperature in Celsius
        ambient_temp: Ambient temperature in Celsius
        material: Conductor material ("copper" or "aluminum")
        
    Returns:
        float: Ampacity in amperes
    """
    # Base ampacity coefficients (A/mm²) at 30°C ambient for 90°C insulation
    base_ampacity_coeff = {
        "copper": 5.0,
        "aluminum": 3.8
    }
    
    if material.lower() not in base_ampacity_coeff:
        raise ValueError("Material must be 'copper' or 'aluminum'")
    
    # Temperature correction factor
    temp_factor = sqrt((insulation_temp - ambient_temp) / (insulation_temp - 30))
    
    # Calculate ampacity
    base_ampacity = base_ampacity_coeff[material.lower()] * size_mm2
    return base_ampacity * temp_factor

def required_conductor_size(
    current: float,
    voltage_drop_limit: float,
    length: float,
    system_voltage: float,
    material: str = "copper"
) -> float:
    """
    Calculate required conductor size based on voltage drop limit.
    
    Args:
        current: Current in amperes
        voltage_drop_limit: Maximum allowed voltage drop percentage
        length: One-way length of conductor in meters
        system_voltage: System voltage in volts
        material: Conductor material ("copper" or "aluminum")
        
    Returns:
        float: Required conductor size in mm²
    """
    # Resistivity at 20°C (ohm·mm²/m)
    resistivity = {
        "copper": 0.0172,
        "aluminum": 0.0282
    }
    
    if material.lower() not in resistivity:
        raise ValueError("Material must be 'copper' or 'aluminum'")
    
    # Calculate maximum allowed voltage drop
    max_voltage_drop = system_voltage * (voltage_drop_limit / 100)
    
    # Calculate required size
    size = (2 * length * current * resistivity[material.lower()]) / max_voltage_drop
    return size

def string_electrical_characteristics(
    modules_in_series: int,
    voc: float,
    vmp: float,
    isc: float,
    imp: float,
    temp_coeff_voc: float,
    min_temp: float = -10.0,
    max_temp: float = 70.0
) -> Dict[str, float]:
    """
    Calculate string electrical characteristics including temperature effects.
    
    Args:
        modules_in_series: Number of modules in series
        voc: Module open circuit voltage at STC
        vmp: Module maximum power voltage at STC
        isc: Module short circuit current at STC
        imp: Module maximum power current at STC
        temp_coeff_voc: Temperature coefficient for Voc (%/°C)
        min_temp: Minimum cell temperature in Celsius
        max_temp: Maximum cell temperature in Celsius
        
    Returns:
        Dict containing:
            - voc_max: Maximum string Voc (at min temp)
            - voc_min: Minimum string Voc (at max temp)
            - vmp_max: Maximum string Vmp (at min temp)
            - vmp_min: Minimum string Vmp (at max temp)
            - isc: String short circuit current
            - imp: String operating current
    """
    # Convert temperature coefficient to decimal form
    temp_coeff_decimal = temp_coeff_voc / 100
    
    # Calculate temperature adjustments
    delta_t_min = min_temp - 25  # Difference from STC
    delta_t_max = max_temp - 25  # Difference from STC
    
    # Calculate voltage variations
    voc_max = voc * (1 + temp_coeff_decimal * delta_t_min) * modules_in_series
    voc_min = voc * (1 + temp_coeff_decimal * delta_t_max) * modules_in_series
    
    # Assume same temperature coefficient for Vmp (simplified)
    vmp_max = vmp * (1 + temp_coeff_decimal * delta_t_min) * modules_in_series
    vmp_min = vmp * (1 + temp_coeff_decimal * delta_t_max) * modules_in_series
    
    return {
        "voc_max": voc_max,
        "voc_min": voc_min,
        "vmp_max": vmp_max,
        "vmp_min": vmp_min,
        "isc": isc,  # Current doesn't change in series
        "imp": imp   # Current doesn't change in series
    }

def conductor_fill_ratio(
    cables: List[Dict[str, float]],
    conduit_size: float
) -> float:
    """
    Calculate conduit fill ratio according to NEC requirements.
    
    Args:
        cables: List of dictionaries containing cable outer diameters in mm
               [{"diameter": float}, ...]
        conduit_size: Inner diameter of conduit in mm
        
    Returns:
        float: Fill ratio as a decimal
    """
    # Calculate total cable area
    total_cable_area = sum(
        pi * (cable["diameter"] / 2) ** 2
        for cable in cables
    )
    
    # Calculate conduit area
    conduit_area = pi * (conduit_size / 2) ** 2
    
    # Calculate fill ratio
    return total_cable_area / conduit_area

def wire_harness_compatibility(
    num_strings: int,
    string_current: float,
    harness_rating: float,
    temperature_rating: float = 90.0,
    ambient_temperature: float = 30.0
) -> Tuple[bool, float]:
    """
    Check wire harness compatibility with number of strings.
    
    Args:
        num_strings: Number of strings to be combined
        string_current: Current per string in amperes
        harness_rating: Harness ampacity rating at standard conditions
        temperature_rating: Harness temperature rating in Celsius
        ambient_temperature: Ambient temperature in Celsius
        
    Returns:
        Tuple containing:
            - bool: True if combination is acceptable
            - float: Actual current as percentage of adjusted rating
    """
    # Temperature correction factor
    temp_factor = sqrt((temperature_rating - ambient_temperature) / 
                      (temperature_rating - 30))
    
    # Adjusted harness rating
    adjusted_rating = harness_rating * temp_factor
    
    # Total current
    total_current = num_strings * string_current
    
    # Calculate utilization percentage
    utilization = (total_current / adjusted_rating) * 100
    
    # Check if combination is acceptable (typically want to stay under 80%)
    is_acceptable = utilization <= 80
    
    return is_acceptable, utilization

def calculate_nec_current(isc: float) -> float:
    """Calculate NEC-compliant current (125% of Isc)"""
    return isc * 1.25
    
def calculate_conductor_required_ampacity(isc: float) -> float:
    """Calculate required conductor ampacity per NEC (156.25% of Isc)"""
    # 156.25% comes from 125% * 125% (double 125% factor)
    return isc * 1.5625

def get_ampacity_for_wire_gauge(wire_gauge: str, temperature_rating: int = 90) -> float:
    """
    Get ampacity for given wire gauge and temperature rating
    Based on NEC Table 310.15(B)(16)
    """
    # NEC Table 310.15(B)(16) values for common wire sizes (90°C column)
    ampacity_table = {
        "14 AWG": 25,
        "12 AWG": 30,
        "10 AWG": 40,
        "8 AWG": 55,
        "6 AWG": 75,
        "4 AWG": 95,
        "2 AWG": 130,
        "1/0 AWG": 170,
        "2/0 AWG": 195,
        "4/0 AWG": 260
    }
    
    return ampacity_table.get(wire_gauge, 0)

def validate_device_inputs(
    num_inputs_required: int,
    num_inputs_available: int
) -> Tuple[bool, str]:
    """
    Validate if there are enough device inputs for the wiring configuration.
    
    Args:
        num_inputs_required: Number of inputs needed for current configuration
        num_inputs_available: Number of inputs available on device
        
    Returns:
        Tuple containing:
            - bool: True if valid, False if error
            - str: Error message if invalid, empty string if valid
    """
    if num_inputs_required > num_inputs_available:
        return False, f"Configuration requires {num_inputs_required} inputs, but device only has {num_inputs_available} inputs"
    return True, ""

def validate_input_current(
    input_current: float,
    max_input_current: float
) -> Tuple[bool, str, float]:
    """
    Validate if the current flowing into an input is within limits.
    
    Args:
        input_current: Current flowing into the input (A)
        max_input_current: Maximum rated current for the input (A)
        
    Returns:
        Tuple containing:
            - bool: True if valid, False if error
            - str: Error message if invalid, empty string if valid
            - float: Utilization percentage
    """
    # Apply NEC 125% safety factor
    design_current = input_current * 1.25
    utilization = (design_current / max_input_current) * 100
    
    if design_current > max_input_current:
        return False, f"Input current ({design_current:.1f}A) exceeds maximum rated current ({max_input_current:.1f}A)", utilization
    return True, "", utilization

def validate_mppt_channel(
    channel_current: float,
    max_channel_current: float
) -> Tuple[bool, str, float]:
    """
    Validate if the current flowing into an MPPT channel is within limits.
    
    Args:
        channel_current: Current flowing into the MPPT channel (A)
        max_channel_current: Maximum rated current for the MPPT channel (A)
        
    Returns:
        Tuple containing:
            - bool: True if valid, False if error
            - str: Error message if invalid, empty string if valid
            - float: Utilization percentage
    """
    # Apply NEC 125% safety factor
    design_current = channel_current * 1.25
    utilization = (design_current / max_channel_current) * 100
    
    if design_current > max_channel_current:
        return False, f"MPPT channel current ({design_current:.1f}A) exceeds maximum rated current ({max_channel_current:.1f}A)", utilization
    return True, "", utilization

def calculate_harness_inputs_required(
    trackers: List[Dict],
    harness_groupings: Dict[str, List[Dict]],
    wiring_type: str
) -> int:
    """
    Calculate the number of inputs required based on wiring configuration.
    
    Args:
        trackers: List of tracker information dictionaries
        harness_groupings: Dictionary mapping tracker IDs to lists of harness groups
        wiring_type: Type of wiring ('String Homerun' or 'Wire Harness')
        
    Returns:
        int: Number of inputs required
    """
    if wiring_type == 'String Homerun':
        # Each string requires its own input
        return sum(len(tracker.get('strings', [])) for tracker in trackers)
    else:  # Wire Harness
        input_count = 0
        for tracker_idx, tracker in enumerate(trackers):
            tracker_id = str(tracker_idx)
            if tracker_id in harness_groupings and harness_groupings[tracker_id]:
                # Each harness group requires one input
                input_count += len(harness_groupings[tracker_id])
            else:
                # Default: one harness per tracker if not specified
                input_count += 1
        return input_count
    
# Standard fuse sizes in amps
STANDARD_FUSE_SIZES = [5, 10, 15, 20, 25, 30, 32, 35, 40, 45, 50, 60, 70, 80, 90]

# Standard breaker sizes in amps
STANDARD_BREAKER_SIZES = [100, 125, 150, 175, 200, 225, 250, 275, 300, 320, 325, 350, 400, 450, 500, 600, 700, 800]

def calculate_fuse_size(current_amps: float) -> int:
    """Calculate required fuse size based on current"""
    for size in STANDARD_FUSE_SIZES:
        if size >= current_amps:
            return size
    return STANDARD_FUSE_SIZES[-1]

def calculate_breaker_size(current_amps: float) -> int:
    """Calculate required breaker size based on current"""
    for size in STANDARD_BREAKER_SIZES:
        if size >= current_amps:
            return size
    return STANDARD_BREAKER_SIZES[-1]

def validate_cable_for_current(cable_size: str, current: float, nec_factor: float = 1.25) -> bool:
    """
    Validate if a cable size is adequate for the given current.
    
    Args:
        cable_size: AWG cable size string
        current: Current in amperes
        nec_factor: NEC safety factor (default 1.25)
        
    Returns:
        bool: True if cable is adequately sized
    """
    ampacity = get_ampacity_for_wire_gauge(cable_size)
    if ampacity == 0:
        return True  # Unknown size, assume OK
    
    nec_current = current * nec_factor
    return nec_current <= ampacity

def get_cable_load_percentage(cable_size: str, current: float, nec_factor: float = 1.25) -> float:
    """
    Calculate the load percentage for a cable.
    
    Args:
        cable_size: AWG cable size string
        current: Current in amperes
        nec_factor: NEC safety factor (default 1.25)
        
    Returns:
        float: Load percentage (0-100+)
    """
    ampacity = get_ampacity_for_wire_gauge(cable_size)
    if ampacity == 0:
        return 0  # Unknown size
    
    nec_current = current * nec_factor
    return (nec_current / ampacity) * 100

def natural_sort_key(text):
    """
    Create a key for natural sorting that handles various formats including:
    - Simple numeric: "Block_01", "Block_2" 
    - Dotted notation: "2.5.01", "3.1.20"
    - Mixed alphanumeric: "Area2Block1"
    - Non-numeric: "MainBlock"
    
    Args:
        text: String to create sort key for
        
    Returns:
        Tuple suitable for sorting
    """
    import re
    
    # Split the text into parts, separating numbers from text
    parts = []
    
    # Handle dotted notation specially
    if '.' in text:
        # Split by dots first
        segments = text.split('.')
        for segment in segments:
            # For each segment, separate numbers and text
            tokens = re.split(r'(\d+)', segment)
            for token in tokens:
                if token:  # Skip empty strings
                    if token.isdigit():
                        parts.append(int(token))
                    else:
                        parts.append(token.lower())
    else:
        # Regular splitting for non-dotted strings
        tokens = re.split(r'(\d+)', text)
        for token in tokens:
            if token:  # Skip empty strings
                if token.isdigit():
                    parts.append(int(token))
                else:
                    parts.append(token.lower())
    
    return tuple(parts)