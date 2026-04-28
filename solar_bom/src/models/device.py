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
    nec_factor: float = 1.56
    
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
        # Use combiner box fuse sizes (not inline harness fuses)
        from ..utils.calculations import COMBINER_BOX_FUSE_SIZES
        
        for size in COMBINER_BOX_FUSE_SIZES:
            if size >= self.harness_current:
                return size
        return 65  # Max standard combiner box fuse size
    
    def _calculate_cable_size(self) -> str:
        """Calculate required cable size based on harness current"""
        from ..utils.cable_sizing import calculate_harness_cable_size
        
        # Use the new cable sizing service to calculate based on actual current
        # Note: harness_current already includes NEC factor, so we pass factor=1.0
        return calculate_harness_cable_size(
            self.num_strings, 
            self.module_isc, 
            self.nec_factor
        )
    
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
    connections: List[HarnessConnection]
    
    # Add these new fields below the existing ones
    use_whips: bool = True  # Default to using whips
    whip_length_ft: int = 3  # Default whip length
    
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
            # harness_current already includes NEC factor, so just sum them
            self.total_input_current = sum(conn.harness_current for conn in self.connections)
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


@dataclass
class StringInverterConfig:
    """Configuration for a string inverter in the device tab."""
    inverter_id: str        # e.g., "SI-01"
    block_id: str
    inverter_spec: object   # InverterSpec (avoid circular import)
    connections: List[HarnessConnection]
    mppt_assignments: Dict[str, int] = field(default_factory=dict)  # harness_id -> channel_idx

    def __post_init__(self):
        pass  # connections carry their own harness_current

    def calculate_total_dc_current(self) -> float:
        """Sum of harness currents across all connections."""
        return sum(conn.harness_current for conn in self.connections)

    def get_mppt_max_current(self, channel_idx: int) -> float:
        """
        Return the MPPT max input current for a given channel index.
        Independent config → per-channel max_input_current.
        Shared config → sum of all channels' max_input_current.
        """
        try:
            from .inverter import MPPTConfig
            channels = getattr(self.inverter_spec, 'mppt_channels', [])
            if not channels:
                return 0.0
            config = getattr(self.inverter_spec, 'mppt_config', MPPTConfig.INDEPENDENT)
            if config == MPPTConfig.SHARED:
                return sum(getattr(ch, 'max_input_current', 0.0) for ch in channels)
            # Independent: return the specific channel
            if 0 <= channel_idx < len(channels):
                return getattr(channels[channel_idx], 'max_input_current', 0.0)
            return 0.0
        except Exception:
            return 0.0

    def get_max_ac_output_current(self) -> float:
        """Return inverter max AC output current."""
        return float(getattr(self.inverter_spec, 'max_ac_current', 0.0))

