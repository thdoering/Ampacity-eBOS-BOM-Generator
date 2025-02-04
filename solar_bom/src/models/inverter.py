from dataclasses import dataclass
from typing import List, Dict, Optional
from enum import Enum

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
    rated_power: float
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
    
    def validate(self) -> bool:
        """
        Validate that inverter specifications are within reasonable bounds
        Returns True if valid, raises ValueError if invalid
        """
        if self.rated_power <= 0:
            raise ValueError("Rated power must be positive")
            
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
    
    def __str__(self) -> str:
        return f"{self.manufacturer} {self.model} ({self.rated_power}kW)"