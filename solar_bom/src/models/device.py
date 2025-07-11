from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple
from enum import Enum

@dataclass
class HarnessConnection:
    """Represents a single harness connection to a combiner box"""
    block_id: str
    tracker_id: str  # e.g., "T01", "T02"
    harness_id: str  # e.g., "H01", "H02"
    num_strings: int
    module_isc: float
    nec_factor: float = 1.25
    
    # Wiring config reference
    actual_cable_size: str = "8 AWG"  # From wiring config
    
    # Calculated values
    harness_current: float = field(init=False)
    calculated_fuse_size: int = field(init=False)
    calculated_cable_size: str = field(init=False)
    
    # User overrides
    user_fuse_size: Optional[int] = None
    user_cable_size: Optional[str] = None
    fuse_manually_set: bool = False
    cable_manually_set: bool = False
    
    def __post_init__(self):
        self.harness_current = self.num_strings * self.module_isc * self.nec_factor
        self.calculated_fuse_size = self._calculate_fuse_size()
        self.calculated_cable_size = self._calculate_cable_size()
    
    def _calculate_fuse_size(self) -> int:
        """Calculate required fuse size based on harness current"""
        # Standard fuse sizes
        FUSE_SIZES = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 60, 70, 80, 90]
        
        for size in FUSE_SIZES:
            if size >= self.harness_current:
                return size
        return 90  # Max standard size
    
    def _calculate_cable_size(self) -> str:
        """Calculate required cable size based on fuse rating"""
        fuse_size = self.user_fuse_size if self.user_fuse_size else self.calculated_fuse_size
        
        # Cable ampacity at 90°C
        if fuse_size <= 40:
            return "10 AWG"
        elif fuse_size <= 55:
            return "8 AWG"
        elif fuse_size <= 75:
            return "6 AWG"
        else:
            return "4 AWG"
    
    def get_display_fuse_size(self) -> int:
        """Get fuse size to display (user override or calculated)"""
        return self.user_fuse_size if self.user_fuse_size else self.calculated_fuse_size
    
    def get_display_cable_size(self) -> str:
        """Get cable size to display (user override or calculated)"""
        return self.user_cable_size if self.user_cable_size else self.calculated_cable_size
    
    def is_cable_size_mismatch(self) -> bool:
        """Check if calculated/selected cable size differs from wiring config"""
        display_size = self.get_display_cable_size()
        return display_size != self.actual_cable_size

@dataclass
class CombinerBoxConfig:
    """Configuration for a single combiner box"""
    combiner_id: str  # e.g., "CB-01"
    block_id: str
    connections: List[HarnessConnection] = field(default_factory=list)
    
    # Calculated values
    total_input_current: float = field(init=False)
    calculated_breaker_size: int = field(init=False)
    
    # User override
    user_breaker_size: Optional[int] = None
    breaker_manually_set: bool = False
    
    def __post_init__(self):
        self.calculate_totals()
    
    def calculate_totals(self):
        """Calculate total current and breaker size"""
        if self.connections:
            sum_current = sum(conn.harness_current for conn in self.connections)
            self.total_input_current = sum_current * 1.25  # Additional NEC factor
            self.calculated_breaker_size = self._calculate_breaker_size()
        else:
            self.total_input_current = 0
            self.calculated_breaker_size = 100
    
    def _calculate_breaker_size(self) -> int:
        """Calculate required breaker size"""
        BREAKER_SIZES = [100, 125, 150, 175, 200, 225, 250, 300, 350, 400, 450, 500, 600, 700, 800]
        
        for size in BREAKER_SIZES:
            if size >= self.total_input_current:
                return size
        return BREAKER_SIZES[-1]
    
    def get_display_breaker_size(self) -> int:
        """Get breaker size to display (user override or calculated)"""
        return self.user_breaker_size if self.user_breaker_size else self.calculated_breaker_size