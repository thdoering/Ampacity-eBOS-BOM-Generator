"""
String-to-Inverter Rolling Allocation Engine

Calculates how strings from trackers get distributed across inverters
using a rolling allocation pattern.

Example: 3-string trackers, 10 strings per inverter, 30 trackers
  Inverter 1: [3, 3, 3, 1]  - 3 full trackers + 1 string from tracker 4
  Inverter 2: [2, 3, 3, 2]  - 2 leftover from tracker 4, 2 full, 2 from tracker 7
  Inverter 3: [1, 3, 3, 3]  - 1 leftover from tracker 7, 3 full trackers
  Pattern repeats...
"""

from typing import List, Dict, Any, Tuple
from math import gcd


def compute_allocation_cycle(strings_per_tracker: int, strings_per_inverter: int) -> List[List[int]]:
    """
    Compute the repeating inverter allocation cycle.
    
    Each entry in the returned list represents one inverter's allocation pattern,
    showing how many strings come from each consecutive tracker.
    
    Args:
        strings_per_tracker: Number of strings on each tracker (e.g., 3)
        strings_per_inverter: Number of strings assigned to each inverter (e.g., 10)
        
    Returns:
        List of inverter patterns. Each pattern is a list of ints showing
        strings taken from consecutive trackers.
        
    Example:
        compute_allocation_cycle(3, 10) returns:
        [[3, 3, 3, 1], [2, 3, 3, 2], [1, 3, 3, 3]]
    """
    if strings_per_tracker <= 0 or strings_per_inverter <= 0:
        return []
    
    # LCM determines when the pattern repeats
    lcm = (strings_per_tracker * strings_per_inverter) // gcd(strings_per_tracker, strings_per_inverter)
    total_strings_in_cycle = lcm
    inverters_in_cycle = total_strings_in_cycle // strings_per_inverter
    trackers_in_cycle = total_strings_in_cycle // strings_per_tracker
    
    # Walk through trackers, filling inverters
    cycle = []
    remaining_in_tracker = strings_per_tracker  # Strings left in current tracker
    
    for inv_idx in range(inverters_in_cycle):
        pattern = []
        strings_needed = strings_per_inverter
        
        while strings_needed > 0:
            take = min(strings_needed, remaining_in_tracker)
            pattern.append(take)
            strings_needed -= take
            remaining_in_tracker -= take
            
            if remaining_in_tracker == 0:
                remaining_in_tracker = strings_per_tracker
        
        cycle.append(pattern)
    
    return cycle


def allocate_strings(strings_per_tracker: int, strings_per_inverter: int, 
                     num_trackers: int) -> Dict[str, Any]:
    """
    Perform full string-to-inverter allocation for a block.
    
    Args:
        strings_per_tracker: Number of strings on each tracker
        strings_per_inverter: Number of strings assigned to each inverter
        num_trackers: Total number of trackers in the block
        
    Returns:
        Dictionary containing:
            - cycle: The repeating allocation pattern
            - inverters: List of inverter assignments, each with:
                - pattern: List of ints (strings from each tracker)
                - tracker_indices: List of (tracker_index, strings_taken) tuples
                - total_strings: Total strings on this inverter
                - full_trackers: Count of trackers fully assigned to this inverter
                - split_trackers: Count of trackers partially assigned
            - summary:
                - total_inverters: Number of inverters needed
                - full_inverters: Inverters with exactly strings_per_inverter
                - partial_inverter_strings: Strings on the last inverter (if partial)
                - total_strings: Total strings allocated
                - total_split_trackers: Number of trackers shared between inverters
                - cycle_length: Number of inverters per repeating cycle
    """
    if strings_per_tracker <= 0 or strings_per_inverter <= 0 or num_trackers <= 0:
        return {
            'cycle': [],
            'inverters': [],
            'summary': {
                'total_inverters': 0,
                'full_inverters': 0,
                'partial_inverter_strings': 0,
                'total_strings': 0,
                'total_split_trackers': 0,
                'cycle_length': 0
            }
        }
    
    cycle = compute_allocation_cycle(strings_per_tracker, strings_per_inverter)
    total_strings = num_trackers * strings_per_tracker
    
    # Walk through all trackers, assigning to inverters using the cycle pattern
    inverters = []
    remaining_in_tracker = strings_per_tracker
    current_tracker_idx = 0
    strings_allocated = 0
    cycle_idx = 0
    
    while strings_allocated < total_strings:
        strings_for_this_inverter = min(strings_per_inverter, total_strings - strings_allocated)
        
        inv_data = {
            'pattern': [],
            'tracker_indices': [],
            'total_strings': 0,
            'full_trackers': 0,
            'split_trackers': 0
        }
        
        strings_needed = strings_for_this_inverter
        
        while strings_needed > 0 and current_tracker_idx < num_trackers:
            take = min(strings_needed, remaining_in_tracker)
            inv_data['pattern'].append(take)
            inv_data['tracker_indices'].append((current_tracker_idx, take))
            inv_data['total_strings'] += take
            strings_needed -= take
            remaining_in_tracker -= take
            
            if remaining_in_tracker == 0:
                current_tracker_idx += 1
                remaining_in_tracker = strings_per_tracker
        
        # Count full vs split trackers for this inverter
        for tracker_idx, strings_taken in inv_data['tracker_indices']:
            if strings_taken == strings_per_tracker:
                inv_data['full_trackers'] += 1
            else:
                inv_data['split_trackers'] += 1
        
        strings_allocated += inv_data['total_strings']
        inverters.append(inv_data)
        cycle_idx = (cycle_idx + 1) % max(len(cycle), 1)
    
    # Count total split trackers (trackers that appear in more than one inverter)
    tracker_appearances = {}
    for inv in inverters:
        for tracker_idx, strings_taken in inv['tracker_indices']:
            if tracker_idx not in tracker_appearances:
                tracker_appearances[tracker_idx] = 0
            tracker_appearances[tracker_idx] += 1
    
    total_split_trackers = sum(1 for count in tracker_appearances.values() if count > 1)
    
    # Build summary
    full_inverters = sum(1 for inv in inverters if inv['total_strings'] == strings_per_inverter)
    last_inv = inverters[-1] if inverters else None
    partial_strings = last_inv['total_strings'] if last_inv and last_inv['total_strings'] < strings_per_inverter else 0
    
    return {
        'cycle': cycle,
        'inverters': inverters,
        'summary': {
            'total_inverters': len(inverters),
            'full_inverters': full_inverters,
            'partial_inverter_strings': partial_strings,
            'total_strings': strings_allocated,
            'total_split_trackers': total_split_trackers,
            'cycle_length': len(cycle)
        }
    }


def format_allocation_summary(allocation: Dict[str, Any], strings_per_tracker: int) -> str:
    """
    Format allocation results as a human-readable string for display.
    
    Args:
        allocation: Result from allocate_strings()
        strings_per_tracker: Strings per tracker (for display context)
        
    Returns:
        Formatted multi-line string summary
    """
    if not allocation['inverters']:
        return "No allocation data"
    
    summary = allocation['summary']
    lines = []
    lines.append(f"Total Inverters: {summary['total_inverters']}")
    lines.append(f"Full Inverters: {summary['full_inverters']}")
    if summary['partial_inverter_strings'] > 0:
        lines.append(f"Partial Inverter: {summary['partial_inverter_strings']} strings")
    lines.append(f"Total Strings: {summary['total_strings']}")
    lines.append(f"Split Trackers: {summary['total_split_trackers']}")
    
    # Show the cycle pattern
    cycle = allocation['cycle']
    if cycle:
        lines.append(f"\nRepeating cycle ({len(cycle)} inverters):")
        for i, pattern in enumerate(cycle):
            pattern_str = '-'.join(str(s) for s in pattern)
            lines.append(f"  Inverter {i+1}: [{pattern_str}]")
    
    return '\n'.join(lines)