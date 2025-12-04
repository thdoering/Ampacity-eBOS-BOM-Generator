"""
Pricing Lookup Utility
Provides functions to look up component prices from pricing_data.json
"""
import json
import os
from typing import Dict, Any, Optional, Tuple


class PricingLookup:
    """Utility class for looking up component prices"""
    
    def __init__(self):
        self.pricing_data: Dict[str, Any] = {}
        self.load_pricing_data()
    
    def get_pricing_file_path(self) -> str:
        """Get the path to the pricing data file"""
        current_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(os.path.dirname(current_dir))
        return os.path.join(project_root, 'data', 'pricing_data.json')
    
    def load_pricing_data(self):
        """Load pricing data from JSON file"""
        filepath = self.get_pricing_file_path()
        
        if os.path.exists(filepath):
            try:
                with open(filepath, 'r') as f:
                    self.pricing_data = json.load(f)
            except Exception as e:
                print(f"Warning: Failed to load pricing data: {e}")
                self.pricing_data = {}
        else:
            self.pricing_data = {}
    
    def get_current_copper_price(self) -> float:
        """Get the current copper price setting"""
        return self.pricing_data.get('settings', {}).get('current_copper_price', 4.5)
    
    def get_copper_tiers(self) -> list:
        """Get the list of copper price tiers"""
        return self.pricing_data.get('settings', {}).get('copper_price_tiers', [4.0, 4.5, 5.0, 5.5, 6.0])
    
    def get_active_tier(self, copper_price: Optional[float] = None) -> str:
        """
        Get the active tier key for the given copper price.
        Returns the tier as a string (e.g., "4.5") for use as a dictionary key.
        """
        if copper_price is None:
            copper_price = self.get_current_copper_price()
        
        tiers = self.get_copper_tiers()
        
        # Find the appropriate tier (highest tier that copper price is >= to)
        active_tier = tiers[0]
        for tier in tiers:
            if copper_price >= tier:
                active_tier = tier
            else:
                break
        
        return str(active_tier)
    
    def get_price(self, part_number: str, copper_price: Optional[float] = None) -> Optional[float]:
        """
        Look up the price for a given part number.
        
        Args:
            part_number: The part number to look up
            copper_price: Optional copper price override (uses current setting if not provided)
        
        Returns:
            The unit price, or None if not found
        """
        if not part_number or part_number == 'N/A':
            return None
        
        # Clean up part number
        part_number = str(part_number).strip()
        
        # Check fuses first (flat pricing)
        fuses = self.pricing_data.get('fuses', {})
        if part_number in fuses:
            price = fuses[part_number]
            return round(float(price), 2) if price else None
        
        # Get active tier for copper-indexed items
        tier_key = self.get_active_tier(copper_price)
        
        # Check extenders
        price = self._lookup_copper_indexed('extenders', part_number, tier_key)
        if price is not None:
            return price
        
        # Check whips
        price = self._lookup_copper_indexed('whips', part_number, tier_key)
        if price is not None:
            return price
        
        # Check harnesses
        price = self._lookup_copper_indexed('harnesses', part_number, tier_key)
        if price is not None:
            return price
        
        # Check combiner boxes (flat pricing)
        combiner_boxes = self.pricing_data.get('combiner_boxes', {})
        if part_number in combiner_boxes:
            price = combiner_boxes[part_number]
            return round(float(price), 2) if price else None
        
        return None
    
    def _lookup_copper_indexed(self, category: str, part_number: str, tier_key: str) -> Optional[float]:
        """Look up a price in a copper-indexed category (extenders, whips, harnesses)"""
        category_data = self.pricing_data.get(category, {})
        
        for subcategory, items in category_data.items():
            if part_number in items:
                prices = items[part_number]
                if isinstance(prices, dict) and tier_key in prices:
                    return round(float(prices[tier_key]), 2)
                # Try with float key in case of type mismatch
                elif isinstance(prices, dict):
                    for key, value in prices.items():
                        if float(key) == float(tier_key):
                            return round(float(value), 2)
        
        return None
    
    def get_price_with_details(self, part_number: str, copper_price: Optional[float] = None) -> Tuple[Optional[float], str]:
        """
        Look up the price and return details about where it came from.
        
        Returns:
            Tuple of (price, source_description)
        """
        price = self.get_price(part_number, copper_price)
        
        if price is None:
            return None, "Not found in pricing data"
        
        tier_key = self.get_active_tier(copper_price)
        return price, f"Copper tier ${tier_key}/lb"


# Singleton instance for convenience
_pricing_lookup_instance = None

def get_pricing_lookup() -> PricingLookup:
    """Get the singleton PricingLookup instance"""
    global _pricing_lookup_instance
    if _pricing_lookup_instance is None:
        _pricing_lookup_instance = PricingLookup()
    return _pricing_lookup_instance

def lookup_price(part_number: str, copper_price: Optional[float] = None) -> Optional[float]:
    """Convenience function to look up a price"""
    return get_pricing_lookup().get_price(part_number, copper_price)

def reload_pricing_data():
    """Reload pricing data from disk"""
    global _pricing_lookup_instance
    _pricing_lookup_instance = None
    return get_pricing_lookup()