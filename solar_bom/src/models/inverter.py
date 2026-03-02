from dataclasses import dataclass
from typing import List, Dict, Optional
from enum import Enum

class InverterType(Enum):
    """Enumeration of inverter types"""
    STRING = "String"
    CENTRAL = "Central"

class MPPTConfig(Enum):
    """Enumeration of MPPT configuration types"""
    INDEPENDENT = "Independent"
    PARALLEL = "Parallel"
    SYMMETRIC = "Symmetric"

@dataclass
class MPPTChannel:
    """Data class representing a single MPPT channel"""
    max_input_current: float
    min_voltage: float
    max_voltage: float
    max_power: float
    num_string_inputs: int

@dataclass
class InverterSpec:
    """Data class representing inverter specifications"""
    
    # Basic inverter info
    manufacturer: str
    model: str
    inverter_type: InverterType
    rated_power_kw: float       # AC rated power in kW
    max_dc_power_kw: float      # Maximum DC input power in kW
    max_efficiency: float
    
    # Input specifications
    mppt_channels: List[MPPTChannel]
    mppt_configuration: MPPTConfig
    max_dc_voltage: float
    startup_voltage: float
    
    # Output specifications
    nominal_ac_voltage: float
    max_ac_current: float
    power_factor: float
    
    # Physical specifications
    dimensions_mm: tuple[float, float, float]  # length, width, depth
    weight_kg: float
    ip_rating: str
    
    # Optional specifications
    temperature_range: Optional[tuple[float, float]] = None  # min, max °C
    altitude_limit_m: Optional[float] = None
    communication_protocol: Optional[str] = None
    max_short_circuit_current: Optional[float] = None  # Max Isc rating (A) - hard ceiling for string count
    
    # Backward compatibility alias
    @property
    def rated_power(self) -> float:
        """Alias for rated_power_kw for backward compatibility"""
        return self.rated_power_kw
    
    @property
    def max_ac_power_w(self) -> float:
        """AC power in watts (for SLD editor compatibility)"""
        return self.rated_power_kw * 1000
    
    def validate(self) -> bool:
        """
        Validate that inverter specifications are within reasonable bounds
        Returns True if valid, raises ValueError if invalid
        """
        if self.rated_power_kw <= 0:
            raise ValueError("Rated AC power must be positive")
            
        if self.max_dc_power_kw <= 0:
            raise ValueError("Max DC power must be positive")
            
        if not (0 < self.max_efficiency <= 100):
            raise ValueError("Efficiency must be between 0 and 100%")
            
        if not self.mppt_channels:
            raise ValueError("Inverter must have at least one MPPT channel")
            
        for channel in self.mppt_channels:
            if channel.max_input_current <= 0:
                raise ValueError("MPPT max input current must be positive")
            if channel.min_voltage <= 0 or channel.max_voltage <= channel.min_voltage:
                raise ValueError("Invalid MPPT voltage range")
            if channel.max_power <= 0:
                raise ValueError("MPPT max power must be positive")
            if channel.num_string_inputs <= 0:
                raise ValueError("Number of string inputs must be positive")
                
        if self.max_dc_voltage <= 0:
            raise ValueError("Maximum DC voltage must be positive")
            
        if self.startup_voltage <= 0:
            raise ValueError("Startup voltage must be positive")
            
        if self.nominal_ac_voltage <= 0:
            raise ValueError("Nominal AC voltage must be positive")
            
        if self.max_ac_current <= 0:
            raise ValueError("Maximum AC current must be positive")
            
        if not (0 < self.power_factor <= 1):
            raise ValueError("Power factor must be between 0 and 1")
            
        return True
    
    def get_total_string_capacity(self) -> int:
        """Calculate total number of string inputs available"""
        return sum(channel.num_string_inputs for channel in self.mppt_channels)
    
    def get_max_power_per_mppt(self) -> Dict[int, float]:
        """Return dictionary of max power capacity per MPPT channel"""
        return {i: channel.max_power for i, channel in enumerate(self.mppt_channels)}
    
    def max_strings_for_module(self, module_wattage: float, modules_per_string: int) -> int:
        """
        Calculate max strings this inverter can accept based on DC power limit.
        
        Args:
            module_wattage: Module power in watts (e.g., 600)
            modules_per_string: Number of modules per string
            
        Returns:
            Maximum number of strings (limited by both power and physical inputs)
        """
        string_power_kw = (module_wattage * modules_per_string) / 1000
        if string_power_kw <= 0:
            return 0
        power_limited = int(self.max_dc_power_kw / string_power_kw)
        input_limited = self.get_total_string_capacity()
        return min(power_limited, input_limited)
    
    def dc_ac_ratio(self, num_strings: int, module_wattage: float, modules_per_string: int) -> float:
        """
        Calculate the DC:AC ratio for a given string count.
        
        Args:
            num_strings: Number of strings connected to this inverter
            module_wattage: Module power in watts
            modules_per_string: Number of modules per string
            
        Returns:
            DC:AC ratio (e.g., 1.25)
        """
        if self.rated_power_kw <= 0:
            return 0.0
        dc_power_kw = (num_strings * modules_per_string * module_wattage) / 1000
        return round(dc_power_kw / self.rated_power_kw, 3)
    
    def strings_for_target_ratio(self, target_ratio: float, module_wattage: float, 
                                  modules_per_string: int) -> int:
        """
        Calculate the number of strings needed to achieve a target DC:AC ratio.
        
        Args:
            target_ratio: Desired DC:AC ratio (e.g., 1.25)
            module_wattage: Module power in watts
            modules_per_string: Number of modules per string
            
        Returns:
            Number of strings (capped by inverter physical limits)
        """
        string_power_kw = (module_wattage * modules_per_string) / 1000
        if string_power_kw <= 0:
            return 0
        target_dc_kw = target_ratio * self.rated_power_kw
        calculated_strings = round(target_dc_kw / string_power_kw)
        max_strings = self.max_strings_for_module(module_wattage, modules_per_string)
        return min(calculated_strings, max_strings)
    
    def __str__(self) -> str:
        type_str = self.inverter_type.value if self.inverter_type else ""
        return f"{self.manufacturer} {self.model} ({self.rated_power_kw}kW {type_str})"