from dataclasses import dataclass
from typing import Optional
from enum import Enum

class ModuleType(Enum):
    """Enumeration of supported module types"""
    MONO_PERC = "Mono PERC"
    BIFACIAL = "Bifacial"
    THIN_FILM = "Thin Film"
    
class ModuleOrientation(Enum):
    """Enumeration of module orientation options"""
    PORTRAIT = "Portrait"
    LANDSCAPE = "Landscape"

@dataclass
class ModuleSpec:
    """Data class representing solar module specifications"""
    
    # Basic module info
    manufacturer: str
    model: str
    type: ModuleType
    
    # Physical specifications
    length_mm: float
    width_mm: float
    depth_mm: float
    weight_kg: float
    
    # Electrical specifications
    wattage: float
    vmp: float  # Maximum power voltage
    imp: float  # Maximum power current
    voc: float  # Open circuit voltage
    isc: float  # Short circuit current
    max_system_voltage: float
    
    # Optional specifications
    efficiency: Optional[float] = None
    temperature_coefficient_pmax: Optional[float] = None  # %/°C
    temperature_coefficient_voc: Optional[float] = None   # %/°C  
    temperature_coefficient_isc: Optional[float] = None   # %/°C
    bifaciality_factor: Optional[float] = None
    
    # Default mounting configuration
    default_orientation: ModuleOrientation = ModuleOrientation.PORTRAIT
    cells_per_module: int = 72
    
    def validate(self) -> bool:
        """
        Validate that module specifications are within reasonable bounds
        Returns True if valid, raises ValueError if invalid
        """
        if self.length_mm <= 0 or self.width_mm <= 0 or self.depth_mm <= 0:
            raise ValueError("Module dimensions must be positive")
            
        if self.weight_kg <= 0:
            raise ValueError("Module weight must be positive")
            
        if self.wattage <= 0:
            raise ValueError("Module wattage must be positive")
            
        if self.vmp <= 0 or self.imp <= 0 or self.voc <= 0 or self.isc <= 0:
            raise ValueError("Electrical specifications must be positive")
            
        if self.max_system_voltage <= 0:
            raise ValueError("Maximum system voltage must be positive")
            
        if self.efficiency is not None and (self.efficiency <= 0 or self.efficiency > 100):
            raise ValueError("Efficiency must be between 0 and 100%")
            
        if self.bifaciality_factor is not None and (self.bifaciality_factor <= 0 or self.bifaciality_factor > 1):
            raise ValueError("Bifaciality factor must be between 0 and 1")
            
        return True
    
    def get_area_m2(self) -> float:
        """Calculate module area in square meters"""
        return (self.length_mm * self.width_mm) / 1_000_000
    
    def get_power_density(self) -> float:
        """Calculate power density in W/m²"""
        return self.wattage / self.get_area_m2()
    
    @property
    def dimensions_mm(self) -> tuple[float, float, float]:
        """Return module dimensions as (length, width, depth) tuple"""
        return (self.length_mm, self.width_mm, self.depth_mm)
    
    def __str__(self) -> str:
        return f"{self.manufacturer} {self.model} ({self.wattage}W)"