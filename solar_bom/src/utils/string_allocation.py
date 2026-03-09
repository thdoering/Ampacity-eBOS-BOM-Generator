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
from math import gcd, ceil


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
    Perform balanced string-to-inverter allocation for a block.
    
    Uses balanced distribution: instead of packing every inverter to max capacity
    (leaving the last one underfilled), distributes strings so all inverters are
    as close to equal as possible.
    
    Example: 3-string trackers, max 10 strings/inv, 31 trackers (93 strings)
      Greedy:   9 × 10 strings + 1 × 3 strings  (wasteful last inverter)
      Balanced: 3 × 10 strings + 7 × 9 strings   (all inverters near capacity)
    
    Args:
        strings_per_tracker: Number of strings on each tracker
        strings_per_inverter: Maximum number of strings assigned to each inverter
        num_trackers: Total number of trackers in the block
        
    Returns:
        Dictionary containing:
            - cycle: List of unique patterns seen across inverters
            - inverters: List of inverter assignments, each with:
                - pattern: List of ints (strings from each tracker)
                - tracker_indices: List of (tracker_index, strings_taken) tuples
                - total_strings: Total strings on this inverter
                - target_strings: The target this inverter was assigned
                - full_trackers: Count of trackers fully assigned to this inverter
                - split_trackers: Count of trackers partially assigned
            - summary:
                - total_inverters: Number of inverters needed
                - full_inverters: Inverters with exactly strings_per_inverter strings
                - partial_inverter_strings: 0 for balanced (kept for backward compat)
                - total_strings: Total strings allocated
                - total_split_trackers: Number of trackers shared between inverters
                - cycle_length: Number of unique patterns
                - max_strings_per_inverter: Largest inverter size used
                - min_strings_per_inverter: Smallest inverter size used
                - num_larger_inverters: How many inverters get the larger size
                - num_smaller_inverters: How many inverters get the smaller size
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
                'cycle_length': 0,
                'max_strings_per_inverter': 0,
                'min_strings_per_inverter': 0,
                'num_larger_inverters': 0,
                'num_smaller_inverters': 0,
            }
        }
    
    total_strings = num_trackers * strings_per_tracker
    n_inv = ceil(total_strings / strings_per_inverter)
    base = total_strings // n_inv
    remainder = total_strings % n_inv
    
    # Build target sizes: 'remainder' inverters get base+1, rest get base
    # Larger inverters go first in sequence
    targets = [base + 1] * remainder + [base] * (n_inv - remainder)
    
    # Walk through all trackers, assigning to inverters by target
    inverters = []
    remaining_in_tracker = strings_per_tracker
    current_tracker_idx = 0
    
    for inv_idx, target in enumerate(targets):
        inv_data = {
            'pattern': [],
            'tracker_indices': [],
            'total_strings': 0,
            'target_strings': target,
            'full_trackers': 0,
            'split_trackers': 0
        }
        
        strings_needed = target
        
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
        
        inverters.append(inv_data)
    
    # Count total split trackers (trackers that appear in more than one inverter)
    tracker_appearances = {}
    for inv in inverters:
        for tracker_idx, strings_taken in inv['tracker_indices']:
            if tracker_idx not in tracker_appearances:
                tracker_appearances[tracker_idx] = 0
            tracker_appearances[tracker_idx] += 1
    
    total_split_trackers = sum(1 for count in tracker_appearances.values() if count > 1)
    
    # Collect unique patterns for display
    unique_patterns = []
    seen = set()
    for inv in inverters:
        key = tuple(inv['pattern'])
        if key not in seen:
            seen.add(key)
            unique_patterns.append(inv['pattern'])
    
    # Build summary
    full_inverters = sum(1 for inv in inverters if inv['total_strings'] == strings_per_inverter)
    max_spi = base + 1 if remainder > 0 else base
    min_spi = base
    
    return {
        'cycle': unique_patterns,
        'inverters': inverters,
        'summary': {
            'total_inverters': len(inverters),
            'full_inverters': full_inverters,
            'partial_inverter_strings': 0,
            'total_strings': sum(inv['total_strings'] for inv in inverters),
            'total_split_trackers': total_split_trackers,
            'cycle_length': len(unique_patterns),
            'max_strings_per_inverter': max_spi,
            'min_strings_per_inverter': min_spi,
            'num_larger_inverters': remainder,
            'num_smaller_inverters': n_inv - remainder,
        }
    }

def allocate_strings_sequential(tracker_sequence: List[int], 
                                 max_strings_per_inverter: int) -> Dict[str, Any]:
    """
    Balanced string-to-inverter allocation for a physically-ordered tracker sequence.
    
    Unlike allocate_strings() which assumes uniform tracker sizes, this accepts
    a mixed sequence representing the physical left-to-right, row-by-row order
    of trackers on site. Each entry is the string count for that tracker.
    
    Uses balanced distribution: all inverters are within 1 string of each other.
    
    Args:
        tracker_sequence: Ordered list of string counts per tracker.
            Example: [3,3,3,2,2,3,3,1] means 8 trackers in physical order,
            with varying string counts.
        max_strings_per_inverter: Maximum strings any inverter can accept.
        
    Returns:
        Dictionary containing:
            - inverters: List of inverter assignments, each with:
                - pattern: List of ints (strings taken from each consecutive tracker)
                - tracker_indices: List of (tracker_index, strings_taken) tuples
                - total_strings: Total strings on this inverter
                - target_strings: The balanced target for this inverter
                - full_trackers: Count of trackers fully assigned to this inverter
                - split_trackers: Count of trackers partially assigned
                - harness_map: List of dicts with split info per tracker:
                    - tracker_idx: Index in the original sequence
                    - strings_per_tracker: Original string count of this tracker
                    - strings_taken: How many strings assigned to this inverter
                    - is_split: Whether this tracker is shared with another inverter
                    - split_position: 'head' (first part), 'tail' (last part), 
                                     'middle' (rare, 3+ way split), or 'full'
            - summary:
                - total_inverters: Number of inverters needed
                - total_strings: Total strings allocated
                - total_trackers: Total trackers in sequence
                - total_split_trackers: Trackers shared between inverters
                - max_strings_per_inverter: Largest inverter size used
                - min_strings_per_inverter: Smallest inverter size used
                - num_larger_inverters: How many get the larger size
                - num_smaller_inverters: How many get the smaller size
                - tracker_type_counts: Dict of {string_count: quantity} across all trackers
    """
    if not tracker_sequence or max_strings_per_inverter <= 0:
        return {
            'inverters': [],
            'summary': {
                'total_inverters': 0,
                'total_strings': 0,
                'total_trackers': 0,
                'total_split_trackers': 0,
                'max_strings_per_inverter': 0,
                'min_strings_per_inverter': 0,
                'num_larger_inverters': 0,
                'num_smaller_inverters': 0,
                'tracker_type_counts': {},
            }
        }
    
    total_strings = sum(tracker_sequence)
    num_trackers = len(tracker_sequence)
    n_inv = ceil(total_strings / max_strings_per_inverter)
    base = total_strings // n_inv
    remainder = total_strings % n_inv
    
    # Build target sizes: larger inverters first
    targets = [base + 1] * remainder + [base] * (n_inv - remainder)
    
    # Walk through trackers in physical order, filling inverters by target
    inverters = []
    current_tracker_idx = 0
    remaining_in_tracker = tracker_sequence[0] if tracker_sequence else 0
    
    for inv_idx, target in enumerate(targets):
        inv_data = {
            'pattern': [],
            'tracker_indices': [],
            'total_strings': 0,
            'target_strings': target,
            'full_trackers': 0,
            'split_trackers': 0,
            'harness_map': []
        }
        
        strings_needed = target
        
        while strings_needed > 0 and current_tracker_idx < num_trackers:
            take = min(strings_needed, remaining_in_tracker)
            inv_data['pattern'].append(take)
            inv_data['tracker_indices'].append((current_tracker_idx, take))
            inv_data['total_strings'] += take
            strings_needed -= take
            remaining_in_tracker -= take
            
            # Build harness info for this tracker contribution
            harness_entry = {
                'tracker_idx': current_tracker_idx,
                'strings_per_tracker': tracker_sequence[current_tracker_idx],
                'strings_taken': take,
                'is_split': take < tracker_sequence[current_tracker_idx],
                'split_position': 'full'  # Will be updated below
            }
            inv_data['harness_map'].append(harness_entry)
            
            if remaining_in_tracker == 0:
                current_tracker_idx += 1
                if current_tracker_idx < num_trackers:
                    remaining_in_tracker = tracker_sequence[current_tracker_idx]
        
        # Count full vs split trackers
        for entry in inv_data['harness_map']:
            if entry['is_split']:
                inv_data['split_trackers'] += 1
            else:
                inv_data['full_trackers'] += 1
        
        inverters.append(inv_data)
    
    # Second pass: determine split positions (head/tail/middle)
    # Track how many times each tracker appears and in what order
    tracker_appearances = {}  # tracker_idx -> list of inverter indices
    for inv_idx, inv in enumerate(inverters):
        for entry in inv['harness_map']:
            tidx = entry['tracker_idx']
            if tidx not in tracker_appearances:
                tracker_appearances[tidx] = []
            tracker_appearances[tidx].append(inv_idx)
    
    for tidx, inv_indices in tracker_appearances.items():
        if len(inv_indices) == 1:
            # Not split — already marked 'full'
            continue
        for pos, inv_idx in enumerate(inv_indices):
            # Find this tracker's entry in this inverter's harness_map
            for entry in inverters[inv_idx]['harness_map']:
                if entry['tracker_idx'] == tidx:
                    if pos == 0:
                        entry['split_position'] = 'head'
                    elif pos == len(inv_indices) - 1:
                        entry['split_position'] = 'tail'
                    else:
                        entry['split_position'] = 'middle'
                    entry['is_split'] = True
                    break
    
    total_split_trackers = sum(1 for appearances in tracker_appearances.values() 
                                if len(appearances) > 1)
    
    # Tracker type counts
    tracker_type_counts = {}
    for spt in tracker_sequence:
        tracker_type_counts[spt] = tracker_type_counts.get(spt, 0) + 1
    
    max_spi = base + 1 if remainder > 0 else base
    min_spi = base
    
    return {
        'inverters': inverters,
        'summary': {
            'total_inverters': len(inverters),
            'total_strings': sum(inv['total_strings'] for inv in inverters),
            'total_trackers': num_trackers,
            'total_split_trackers': total_split_trackers,
            'max_strings_per_inverter': max_spi,
            'min_strings_per_inverter': min_spi,
            'num_larger_inverters': remainder,
            'num_smaller_inverters': n_inv - remainder,
            'tracker_type_counts': tracker_type_counts,
        }
    }

def allocate_strings_spatial(tracker_entries: List[Dict], 
                              max_strings_per_inverter: int,
                              pitch_ft: float,
                              row_threshold_ft: float = None) -> Dict[str, Any]:
    """
    Spatially-aware string-to-inverter allocation.
    
    Clusters trackers into contiguous spatial runs based on physical position,
    then allocates independently within each run. Inverters never span across
    runs. This prevents unrealistic wiring like connecting trackers on opposite
    sides of the site.
    
    Row breaks: trackers with Y-center difference > row_threshold_ft.
    X breaks: adjacent trackers on the same row with X gap > 2× pitch.
    
    Args:
        tracker_entries: List of dicts, each with:
            - original_idx: int, index in the flat tracker_sequence
            - spt: int, strings per tracker
            - x: float, world X position (E-W) in feet
            - y: float, world Y position (N-S) in feet
            - length_ft: float, tracker N-S length in feet
        max_strings_per_inverter: Maximum strings any inverter can accept.
        pitch_ft: Row spacing / pitch in feet (E-W distance between tracker centers).
        row_threshold_ft: Max Y-center difference to be considered same row.
            Defaults to half the max tracker length if None.
    
    Returns:
        Same structure as allocate_strings_sequential(), with tracker_idx values
        referring to the original flat sequence indices.
    """
    if not tracker_entries or max_strings_per_inverter <= 0:
        return {
            'inverters': [],
            'summary': {
                'total_inverters': 0,
                'total_strings': 0,
                'total_trackers': 0,
                'total_split_trackers': 0,
                'max_strings_per_inverter': 0,
                'min_strings_per_inverter': 0,
                'num_larger_inverters': 0,
                'num_smaller_inverters': 0,
                'tracker_type_counts': {},
            }
        }
    
    # Compute row threshold from tracker lengths if not specified
    if row_threshold_ft is None:
        max_length = max(e.get('length_ft', 180.0) for e in tracker_entries)
        row_threshold_ft = max_length * 0.5
    
    x_gap_threshold = pitch_ft * 2.0
    
    # --- Step 1: Cluster into rows by driveline (motor) Y similarity ---
    # Compute driveline world-Y for each tracker: group_y + motor_y_offset
    # This is more accurate than Y-center because trackers of different heights
    # align on the driveline, not on their bounding-box center.
    for entry in tracker_entries:
        motor_y_offset = entry.get('motor_y_ft', entry.get('length_ft', 180.0) / 2.0)
        entry['driveline_y'] = entry['y'] + motor_y_offset
    
    # Sort by driveline Y to group into rows
    sorted_by_y = sorted(tracker_entries, key=lambda e: e['driveline_y'])
    
    y_debug = [f'T{e["original_idx"]}:y={e["y"]:.1f},motor={e.get("motor_y_ft", 0):.1f},dl={e["driveline_y"]:.1f}' for e in sorted_by_y]    
    rows = []  # list of lists of tracker entries
    current_row = [sorted_by_y[0]]
    current_row_y = sorted_by_y[0]['driveline_y']
    
    for entry in sorted_by_y[1:]:
        # Compare to the LAST entry added to the row (nearest neighbor chaining)
        # instead of the first entry, so rows can gradually span a wider Y range
        last_in_row_y = current_row[-1]['driveline_y']
        y_diff = abs(entry['driveline_y'] - last_in_row_y)
        if y_diff <= row_threshold_ft:
            current_row.append(entry)
        else:
            rows.append(current_row)
            current_row = [entry]
    rows.append(current_row)
    
    for r_idx, row in enumerate(rows):
        y_range = f"dl_y=[{min(e['driveline_y'] for e in row):.1f} .. {max(e['driveline_y'] for e in row):.1f}]"
    
    # --- Step 2: Within each row, sort by X and split into runs by X-gap ---
    runs = []  # list of lists of tracker entries, each run is independently allocated
    
    for r_idx, row in enumerate(rows):
        row_sorted = sorted(row, key=lambda e: e['x'])
        
        current_run = [row_sorted[0]]
        for i in range(1, len(row_sorted)):
            prev_x = row_sorted[i - 1]['x']
            curr_x = row_sorted[i]['x']
            gap = curr_x - prev_x
            
            if gap > x_gap_threshold:
                runs.append(current_run)
                current_run = [row_sorted[i]]
            else:
                current_run.append(row_sorted[i])
        runs.append(current_run)
    
    # --- Step 3: Allocate per run, remap indices ---
    all_inverters = []
    global_inv_offset = 0
    
    for run in runs:
        # Build the local tracker sequence for this run
        local_sequence = [e['spt'] for e in run]
        # Map from local index -> original flat index
        local_to_original = [e['original_idx'] for e in run]
        
        run_result = allocate_strings_sequential(local_sequence, max_strings_per_inverter)
        
        # Remap tracker_idx in each inverter's data back to original indices
        for inv in run_result['inverters']:
            remapped_inv = {
                'pattern': inv['pattern'],
                'tracker_indices': [
                    (local_to_original[local_idx], strings_taken)
                    for local_idx, strings_taken in inv['tracker_indices']
                ],
                'total_strings': inv['total_strings'],
                'target_strings': inv['target_strings'],
                'full_trackers': inv['full_trackers'],
                'split_trackers': inv['split_trackers'],
                'harness_map': [],
            }
            
            for entry in inv['harness_map']:
                remapped_entry = dict(entry)
                remapped_entry['tracker_idx'] = local_to_original[entry['tracker_idx']]
                remapped_inv['harness_map'].append(remapped_entry)
            
            all_inverters.append(remapped_inv)
    
    # --- Step 4: Build merged summary ---
    total_strings = sum(inv['total_strings'] for inv in all_inverters)
    total_trackers = len(tracker_entries)
    
    # Count split trackers globally
    tracker_appearances = {}
    for inv in all_inverters:
        for entry in inv['harness_map']:
            tidx = entry['tracker_idx']
            if tidx not in tracker_appearances:
                tracker_appearances[tidx] = 0
            tracker_appearances[tidx] += 1
    total_split_trackers = sum(1 for count in tracker_appearances.values() if count > 1)
    
    # Tracker type counts
    tracker_type_counts = {}
    for e in tracker_entries:
        tracker_type_counts[e['spt']] = tracker_type_counts.get(e['spt'], 0) + 1
    
    # Min/max strings per inverter
    if all_inverters:
        inv_sizes = [inv['total_strings'] for inv in all_inverters]
        max_spi = max(inv_sizes)
        min_spi = min(inv_sizes)
        num_larger = sum(1 for s in inv_sizes if s == max_spi) if max_spi != min_spi else 0
        num_smaller = sum(1 for s in inv_sizes if s == min_spi) if max_spi != min_spi else 0
    else:
        max_spi = min_spi = num_larger = num_smaller = 0
    
    return {
        'inverters': all_inverters,
        'spatial_runs': len(runs),
        'summary': {
            'total_inverters': len(all_inverters),
            'total_strings': total_strings,
            'total_trackers': total_trackers,
            'total_split_trackers': total_split_trackers,
            'max_strings_per_inverter': max_spi,
            'min_strings_per_inverter': min_spi,
            'num_larger_inverters': num_larger,
            'num_smaller_inverters': num_smaller,
            'tracker_type_counts': tracker_type_counts,
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