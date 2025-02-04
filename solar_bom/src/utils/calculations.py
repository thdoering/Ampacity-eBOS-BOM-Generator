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