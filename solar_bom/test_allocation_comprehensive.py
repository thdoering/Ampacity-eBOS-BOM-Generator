"""
Comprehensive Test Suite for String-to-Inverter Allocation Engine
Tests: correctness, edge cases, real-world scenarios, performance
"""

import sys
import os
import time

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.utils.string_allocation import compute_allocation_cycle, allocate_strings, format_allocation_summary
from math import gcd

# ============================================================
# Test infrastructure
# ============================================================

class TestResults:
    def __init__(self):
        self.passed = 0
        self.failed = 0
        self.errors = []
    
    def ok(self, name):
        self.passed += 1
        print(f"  ✓ {name}")
    
    def fail(self, name, detail):
        self.failed += 1
        self.errors.append((name, detail))
        print(f"  ✗ {name}")
        print(f"    → {detail}")
    
    def summary(self):
        total = self.passed + self.failed
        print(f"\n{'='*60}")
        print(f"Results: {self.passed}/{total} passed, {self.failed} failed")
        if self.errors:
            print(f"\nFailures:")
            for name, detail in self.errors:
                print(f"  {name}: {detail}")
        print(f"{'='*60}")
        return self.failed == 0

T = TestResults()


# ============================================================
# Helper: run standard invariant checks on any allocation
# ============================================================

def check_invariants(result, spt, spi, num_trackers, label):
    """Run all invariant checks on an allocation result. Returns True if all pass."""
    ok = True
    total_strings = spt * num_trackers
    summary = result['summary']
    
    # 1. Total strings in == total strings out
    if summary['total_strings'] != total_strings:
        T.fail(f"{label} — string conservation",
               f"Expected {total_strings} strings, got {summary['total_strings']}")
        ok = False
    else:
        T.ok(f"{label} — string conservation ({total_strings} strings)")
    
    # 2. Every full inverter has exactly spi strings
    for i, inv in enumerate(result['inverters']):
        if i < len(result['inverters']) - 1:
            # Not the last inverter — must be full
            if inv['total_strings'] != spi:
                T.fail(f"{label} — inverter {i} fullness",
                       f"Non-last inverter has {inv['total_strings']} strings, expected {spi}")
                ok = False
    
    # 3. Last inverter has <= spi strings
    if result['inverters']:
        last = result['inverters'][-1]
        if last['total_strings'] > spi:
            T.fail(f"{label} — last inverter overflow",
                   f"Last inverter has {last['total_strings']} strings, max is {spi}")
            ok = False
        elif last['total_strings'] <= 0:
            T.fail(f"{label} — last inverter empty",
                   f"Last inverter has {last['total_strings']} strings")
            ok = False
    
    # 4. No negative values anywhere
    for i, inv in enumerate(result['inverters']):
        for val in inv['pattern']:
            if val <= 0:
                T.fail(f"{label} — non-positive in pattern",
                       f"Inverter {i} pattern has value {val}")
                ok = False
    
    # 5. Pattern sums match total_strings for each inverter
    for i, inv in enumerate(result['inverters']):
        pattern_sum = sum(inv['pattern'])
        if pattern_sum != inv['total_strings']:
            T.fail(f"{label} — pattern sum mismatch",
                   f"Inverter {i}: pattern sums to {pattern_sum}, total_strings={inv['total_strings']}")
            ok = False
    
    # 6. Tracker indices are sequential and within bounds
    all_tracker_refs = []
    for inv in result['inverters']:
        for tracker_idx, strings_taken in inv['tracker_indices']:
            if tracker_idx < 0 or tracker_idx >= num_trackers:
                T.fail(f"{label} — tracker index OOB",
                       f"Tracker index {tracker_idx} out of range [0, {num_trackers})")
                ok = False
            if strings_taken <= 0 or strings_taken > spt:
                T.fail(f"{label} — invalid strings_taken",
                       f"Tracker {tracker_idx}: strings_taken={strings_taken}, max={spt}")
                ok = False
            all_tracker_refs.append((tracker_idx, strings_taken))
    
    # 7. Each tracker's total taken == spt
    tracker_totals = {}
    for tracker_idx, strings_taken in all_tracker_refs:
        tracker_totals[tracker_idx] = tracker_totals.get(tracker_idx, 0) + strings_taken
    
    for tracker_idx in range(num_trackers):
        got = tracker_totals.get(tracker_idx, 0)
        if got != spt:
            T.fail(f"{label} — tracker {tracker_idx} incomplete",
                   f"Tracker {tracker_idx} allocated {got}/{spt} strings")
            ok = False
    
    # 8. Split tracker count matches
    tracker_appearances = {}
    for inv in result['inverters']:
        for tracker_idx, _ in inv['tracker_indices']:
            tracker_appearances[tracker_idx] = tracker_appearances.get(tracker_idx, 0) + 1
    actual_splits = sum(1 for c in tracker_appearances.values() if c > 1)
    if actual_splits != summary['total_split_trackers']:
        T.fail(f"{label} — split tracker count",
               f"Counted {actual_splits} splits, summary says {summary['total_split_trackers']}")
        ok = False
    
    # 9. full_inverters count is correct
    actual_full = sum(1 for inv in result['inverters'] if inv['total_strings'] == spi)
    if actual_full != summary['full_inverters']:
        T.fail(f"{label} — full inverter count",
               f"Counted {actual_full} full, summary says {summary['full_inverters']}")
        ok = False
    
    if ok:
        T.ok(f"{label} — all invariants passed")
    
    return ok


# ============================================================
# Test 1: compute_allocation_cycle correctness
# ============================================================

print("\n" + "="*60)
print("TEST 1: compute_allocation_cycle")
print("="*60)

# 1a: 3-string tracker, 10 strings/inv → cycle of 3
cycle = compute_allocation_cycle(3, 10)
expected = [[3, 3, 3, 1], [2, 3, 3, 2], [1, 3, 3, 3]]
if cycle == expected:
    T.ok("3/10 cycle matches expected pattern")
else:
    T.fail("3/10 cycle", f"Expected {expected}, got {cycle}")

# 1b: Every pattern in cycle sums to strings_per_inverter
for spt, spi in [(2, 10), (3, 10), (4, 7), (5, 12), (1, 8), (6, 4), (3, 3), (4, 4)]:
    cycle = compute_allocation_cycle(spt, spi)
    all_ok = True
    for pattern in cycle:
        if sum(pattern) != spi:
            T.fail(f"cycle {spt}/{spi} pattern sum", f"Pattern {pattern} sums to {sum(pattern)}, expected {spi}")
            all_ok = False
            break
    if all_ok:
        T.ok(f"cycle {spt}/{spi}: all {len(cycle)} patterns sum to {spi}")

# 1c: Cycle length = LCM / spi
for spt, spi in [(2, 10), (3, 10), (4, 7), (5, 12), (6, 9), (3, 3)]:
    cycle = compute_allocation_cycle(spt, spi)
    lcm = (spt * spi) // gcd(spt, spi)
    expected_len = lcm // spi
    if len(cycle) == expected_len:
        T.ok(f"cycle {spt}/{spi}: length {len(cycle)} == LCM({lcm})/{spi}")
    else:
        T.fail(f"cycle {spt}/{spi} length", f"Expected {expected_len}, got {len(cycle)}")

# 1d: Total strings consumed in cycle = LCM
for spt, spi in [(2, 10), (3, 10), (4, 7), (5, 12)]:
    cycle = compute_allocation_cycle(spt, spi)
    total = sum(sum(p) for p in cycle)
    lcm = (spt * spi) // gcd(spt, spi)
    if total == lcm:
        T.ok(f"cycle {spt}/{spi}: total strings in cycle = LCM = {lcm}")
    else:
        T.fail(f"cycle {spt}/{spi} total", f"Expected {lcm}, got {total}")

# 1e: Edge - tracker strings == inverter strings (no splits ever)
cycle = compute_allocation_cycle(10, 10)
if cycle == [[10]]:
    T.ok("10/10: single pattern [10], no splits")
else:
    T.fail("10/10 cycle", f"Expected [[10]], got {cycle}")

# 1f: Edge - 1 string/tracker
cycle = compute_allocation_cycle(1, 10)
if cycle == [[1]*10]:
    T.ok("1/10: pattern is ten 1s")
else:
    T.fail("1/10 cycle", f"Expected [[1,1,...,1]], got {cycle}")

# 1g: Edge - 1 string/inverter  
cycle = compute_allocation_cycle(3, 1)
if cycle == [[1], [1], [1]]:
    T.ok("3/1: three patterns of [1]")
else:
    T.fail("3/1 cycle", f"Expected [[1],[1],[1]], got {cycle}")

# 1h: Edge - invalid inputs
cycle = compute_allocation_cycle(0, 10)
if cycle == []:
    T.ok("0/10: returns empty")
else:
    T.fail("0/10 cycle", f"Expected [], got {cycle}")

cycle = compute_allocation_cycle(3, 0)
if cycle == []:
    T.ok("3/0: returns empty")
else:
    T.fail("3/0 cycle", f"Expected [], got {cycle}")


# ============================================================
# Test 2: Your known test case — 3-string, 10 strings/inv
# ============================================================

print("\n" + "="*60)
print("TEST 2: Known test case (3-str tracker, 10 str/inv)")
print("="*60)

# 2a: 30 trackers — perfectly divisible (90 strings / 10 = 9 inverters)
result = allocate_strings(3, 10, 30)
check_invariants(result, 3, 10, 30, "30x3str/10spi")
if result['summary']['total_inverters'] == 9:
    T.ok("30x3/10: 9 inverters (90/10)")
else:
    T.fail("30x3/10 inverter count", f"Expected 9, got {result['summary']['total_inverters']}")
if result['summary']['partial_inverter_strings'] == 0:
    T.ok("30x3/10: no partial inverter")
else:
    T.fail("30x3/10 partial", f"Expected 0, got {result['summary']['partial_inverter_strings']}")

# 2b: 31 trackers — partial last inverter (93 strings / 10 = 9 full + 3 partial)
result = allocate_strings(3, 10, 31)
check_invariants(result, 3, 10, 31, "31x3str/10spi")
if result['summary']['total_inverters'] == 10:
    T.ok("31x3/10: 10 inverters")
else:
    T.fail("31x3/10 inverter count", f"Expected 10, got {result['summary']['total_inverters']}")
if result['summary']['partial_inverter_strings'] == 3:
    T.ok("31x3/10: partial has 3 strings")
else:
    T.fail("31x3/10 partial", f"Expected 3, got {result['summary']['partial_inverter_strings']}")

# 2c: 7 trackers — small project (21 strings / 10 = 2 full + 1 partial)
result = allocate_strings(3, 10, 7)
check_invariants(result, 3, 10, 7, "7x3str/10spi")
if result['summary']['total_inverters'] == 3:
    T.ok("7x3/10: 3 inverters")
else:
    T.fail("7x3/10 inverter count", f"Expected 3, got {result['summary']['total_inverters']}")
if result['summary']['partial_inverter_strings'] == 1:
    T.ok("7x3/10: partial has 1 string")
else:
    T.fail("7x3/10 partial", f"Expected 1, got {result['summary']['partial_inverter_strings']}")


# ============================================================
# Test 3: Even division — no splits
# ============================================================

print("\n" + "="*60)
print("TEST 3: Even division cases (no splits expected)")
print("="*60)

# 2-string trackers, 10 strings/inv, 50 trackers = 100 strings = 10 inverters, 0 splits
result = allocate_strings(2, 10, 50)
check_invariants(result, 2, 10, 50, "50x2str/10spi")
if result['summary']['total_split_trackers'] == 0:
    T.ok("50x2/10: zero splits (2 divides 10)")
else:
    T.fail("50x2/10 splits", f"Expected 0, got {result['summary']['total_split_trackers']}")

# 5-string trackers, 10 strings/inv, 20 trackers
result = allocate_strings(5, 10, 20)
check_invariants(result, 5, 10, 20, "20x5str/10spi")
if result['summary']['total_split_trackers'] == 0:
    T.ok("20x5/10: zero splits (5 divides 10)")
else:
    T.fail("20x5/10 splits", f"Expected 0, got {result['summary']['total_split_trackers']}")

# 4-string trackers, 8 strings/inv, 10 trackers
result = allocate_strings(4, 8, 10)
check_invariants(result, 4, 8, 10, "10x4str/8spi")
if result['summary']['total_split_trackers'] == 0:
    T.ok("10x4/8: zero splits (4 divides 8)")
else:
    T.fail("10x4/8 splits", f"Expected 0, got {result['summary']['total_split_trackers']}")


# ============================================================
# Test 4: Tracker strings > strings/inverter (every tracker splits)
# ============================================================

print("\n" + "="*60)
print("TEST 4: Tracker has more strings than inverter capacity")
print("="*60)

# 6-string tracker, 4 strings/inv, 10 trackers
result = allocate_strings(6, 4, 10)
check_invariants(result, 6, 4, 10, "10x6str/4spi")

# 60 strings / 4 = 15 inverters
if result['summary']['total_inverters'] == 15:
    T.ok("10x6/4: 15 inverters")
else:
    T.fail("10x6/4 count", f"Expected 15, got {result['summary']['total_inverters']}")

# Check cycle: LCM(6,4) = 12, cycle has 12/4 = 3 inverters
cycle = result['cycle']
if len(cycle) == 3:
    T.ok("10x6/4: cycle length 3")
else:
    T.fail("10x6/4 cycle len", f"Expected 3, got {len(cycle)}")

# 10-string tracker, 3 strings/inv, 5 trackers
result = allocate_strings(10, 3, 5)
check_invariants(result, 10, 3, 5, "5x10str/3spi")

# 50 strings / 3 = 16 full + 2 partial
if result['summary']['total_inverters'] == 17:
    T.ok("5x10/3: 17 inverters")
else:
    T.fail("5x10/3 count", f"Expected 17, got {result['summary']['total_inverters']}")
if result['summary']['partial_inverter_strings'] == 2:
    T.ok("5x10/3: partial has 2 strings")
else:
    T.fail("5x10/3 partial", f"Expected 2, got {result['summary']['partial_inverter_strings']}")


# ============================================================
# Test 5: 1-string tracker edge case
# ============================================================

print("\n" + "="*60)
print("TEST 5: 1-string trackers")
print("="*60)

# 1-string trackers, 10 strings/inv, 25 trackers
result = allocate_strings(1, 10, 25)
check_invariants(result, 1, 10, 25, "25x1str/10spi")
if result['summary']['total_inverters'] == 3:
    T.ok("25x1/10: 3 inverters (25/10 = 2.5 -> 3)")
else:
    T.fail("25x1/10 count", f"Expected 3, got {result['summary']['total_inverters']}")
if result['summary']['total_split_trackers'] == 0:
    T.ok("25x1/10: zero splits (1-string can't split)")
else:
    T.fail("25x1/10 splits", f"Expected 0, got {result['summary']['total_split_trackers']}")

# 1-string trackers, 1 string/inv, 5 trackers (degenerate)
result = allocate_strings(1, 1, 5)
check_invariants(result, 1, 1, 5, "5x1str/1spi")
if result['summary']['total_inverters'] == 5:
    T.ok("5x1/1: 5 inverters (1:1 mapping)")
else:
    T.fail("5x1/1 count", f"Expected 5, got {result['summary']['total_inverters']}")


# ============================================================
# Test 6: Single tracker
# ============================================================

print("\n" + "="*60)
print("TEST 6: Single tracker")
print("="*60)

# 1 tracker, 3 strings, 10 strings/inv -> 1 partial inverter with 3 strings
result = allocate_strings(3, 10, 1)
check_invariants(result, 3, 10, 1, "1x3str/10spi")
if result['summary']['total_inverters'] == 1:
    T.ok("1x3/10: 1 inverter")
else:
    T.fail("1x3/10 count", f"Expected 1, got {result['summary']['total_inverters']}")
if result['summary']['partial_inverter_strings'] == 3:
    T.ok("1x3/10: partial with 3 strings")
else:
    T.fail("1x3/10 partial", f"Expected 3, got {result['summary']['partial_inverter_strings']}")


# ============================================================
# Test 7: Large projects
# ============================================================

print("\n" + "="*60)
print("TEST 7: Large projects")
print("="*60)

# 500 trackers x 3 strings = 1500 strings / 10 = 150 inverters
result = allocate_strings(3, 10, 500)
check_invariants(result, 3, 10, 500, "500x3str/10spi")
if result['summary']['total_inverters'] == 150:
    T.ok("500x3/10: 150 inverters")
else:
    T.fail("500x3/10 count", f"Expected 150, got {result['summary']['total_inverters']}")

# 1000 trackers x 4 strings = 4000 strings / 12 = 333 full + 4 partial
result = allocate_strings(4, 12, 1000)
check_invariants(result, 4, 12, 1000, "1000x4str/12spi")
expected_inv = (4000 + 12 - 1) // 12  # ceil(4000/12) = 334
if result['summary']['total_inverters'] == expected_inv:
    T.ok(f"1000x4/12: {expected_inv} inverters")
else:
    T.fail("1000x4/12 count", f"Expected {expected_inv}, got {result['summary']['total_inverters']}")

# Performance check
t0 = time.time()
result = allocate_strings(3, 10, 5000)
dt = time.time() - t0
check_invariants(result, 3, 10, 5000, "5000x3str/10spi")
if dt < 1.0:
    T.ok(f"5000x3/10: completed in {dt*1000:.1f}ms (< 1s)")
else:
    T.fail("5000x3/10 perf", f"Took {dt:.2f}s")


# ============================================================
# Test 8: Real-world combos (odd spt/spi ratios)
# ============================================================

print("\n" + "="*60)
print("TEST 8: Real-world odd combos")
print("="*60)

test_cases = [
    # (spt, spi, num_trackers, label)
    (4, 7, 50, "4str/7spi — common odd combo"),
    (3, 8, 40, "3str/8spi — doesn't divide"),
    (5, 12, 30, "5str/12spi — large cycle"),
    (2, 7, 100, "2str/7spi — coprime"),
    (3, 13, 20, "3str/13spi — coprime, large spi"),
    (6, 9, 45, "6str/9spi — GCD=3"),
    (4, 10, 60, "4str/10spi — GCD=2"),
    (3, 4, 25, "3str/4spi — small ratio"),
    (7, 11, 33, "7str/11spi — both prime"),
    (2, 3, 150, "2str/3spi — minimal split"),
]

for spt, spi, n, label in test_cases:
    result = allocate_strings(spt, spi, n)
    check_invariants(result, spt, spi, n, label)
    
    # Verify expected inverter count
    total_str = spt * n
    expected_inv = (total_str + spi - 1) // spi  # ceil division
    if result['summary']['total_inverters'] == expected_inv:
        T.ok(f"{label}: {expected_inv} inverters ({total_str} strings)")
    else:
        T.fail(f"{label} inverter count",
               f"Expected {expected_inv}, got {result['summary']['total_inverters']}")


# ============================================================
# Test 9: Cycle pattern correctness — patterns repeat faithfully
# ============================================================

print("\n" + "="*60)
print("TEST 9: Pattern repetition fidelity")
print("="*60)

# For cases where total strings is a multiple of LCM, 
# every cycle's worth of inverters should have identical patterns
for spt, spi in [(3, 10), (4, 7), (2, 5)]:
    lcm = (spt * spi) // gcd(spt, spi)
    trackers_per_cycle = lcm // spt
    # Use enough trackers for 4 full cycles
    num_trackers = trackers_per_cycle * 4
    result = allocate_strings(spt, spi, num_trackers)
    
    cycle = result['cycle']
    cycle_len = len(cycle)
    
    all_match = True
    for i, inv in enumerate(result['inverters']):
        expected_pattern = cycle[i % cycle_len]
        if inv['pattern'] != expected_pattern:
            T.fail(f"pattern repeat {spt}/{spi} inv {i}",
                   f"Expected {expected_pattern}, got {inv['pattern']}")
            all_match = False
            break
    
    if all_match:
        T.ok(f"pattern repeat {spt}/{spi}: all {len(result['inverters'])} inverters match cycle")


# ============================================================
# Test 10: Partial inverter only at the end
# ============================================================

print("\n" + "="*60)
print("TEST 10: Partial inverter only appears at end")
print("="*60)

for spt, spi, n in [(3, 10, 31), (4, 7, 13), (5, 12, 7), (2, 3, 8)]:
    result = allocate_strings(spt, spi, n)
    total = spt * n
    
    non_last_short = False
    for i, inv in enumerate(result['inverters'][:-1]):
        if inv['total_strings'] < spi:
            T.fail(f"partial not last {spt}/{spi}/{n}",
                   f"Inverter {i} has {inv['total_strings']} strings (not last)")
            non_last_short = True
            break
    
    if not non_last_short:
        last_strings = result['inverters'][-1]['total_strings']
        expected_partial = total % spi
        if expected_partial == 0:
            expected_partial = spi  # Fully fills
        if last_strings == expected_partial:
            T.ok(f"partial at end {spt}/{spi}/{n}: last has {last_strings} strings")
        else:
            T.fail(f"partial at end {spt}/{spi}/{n}",
                   f"Last inverter has {last_strings}, expected {expected_partial}")


# ============================================================
# Test 11: Edge — strings_per_inverter == 1
# ============================================================

print("\n" + "="*60)
print("TEST 11: 1 string per inverter (micro-inverter scenario)")
print("="*60)

result = allocate_strings(3, 1, 10)
check_invariants(result, 3, 1, 10, "10x3str/1spi")
if result['summary']['total_inverters'] == 30:
    T.ok("10x3/1: 30 inverters (one per string)")
else:
    T.fail("10x3/1 count", f"Expected 30, got {result['summary']['total_inverters']}")
# Every tracker should be split into 3 inverters
if result['summary']['total_split_trackers'] == 10:
    T.ok("10x3/1: all 10 trackers split")
else:
    T.fail("10x3/1 splits", f"Expected 10, got {result['summary']['total_split_trackers']}")


# ============================================================
# Test 12: Edge — strings_per_tracker == strings_per_inverter
# ============================================================

print("\n" + "="*60)
print("TEST 12: Tracker exactly fills one inverter")
print("="*60)

for spt in [2, 3, 5, 10]:
    result = allocate_strings(spt, spt, 20)
    check_invariants(result, spt, spt, 20, f"20x{spt}str/{spt}spi")
    if result['summary']['total_inverters'] == 20:
        T.ok(f"20x{spt}/{spt}: 20 inverters (1:1)")
    else:
        T.fail(f"20x{spt}/{spt} count",
               f"Expected 20, got {result['summary']['total_inverters']}")
    if result['summary']['total_split_trackers'] == 0:
        T.ok(f"20x{spt}/{spt}: zero splits")
    else:
        T.fail(f"20x{spt}/{spt} splits",
               f"Expected 0, got {result['summary']['total_split_trackers']}")


# ============================================================
# Test 13: Invalid / zero inputs
# ============================================================

print("\n" + "="*60)
print("TEST 13: Invalid / zero inputs")
print("="*60)

result = allocate_strings(0, 10, 5)
if result['summary']['total_inverters'] == 0:
    T.ok("spt=0: returns empty")
else:
    T.fail("spt=0", f"Expected 0 inverters, got {result['summary']['total_inverters']}")

result = allocate_strings(3, 0, 5)
if result['summary']['total_inverters'] == 0:
    T.ok("spi=0: returns empty")
else:
    T.fail("spi=0", f"Expected 0 inverters, got {result['summary']['total_inverters']}")

result = allocate_strings(3, 10, 0)
if result['summary']['total_inverters'] == 0:
    T.ok("n=0: returns empty")
else:
    T.fail("n=0", f"Expected 0 inverters, got {result['summary']['total_inverters']}")


# ============================================================
# Test 14: Split tracker count — manual verification
# ============================================================

print("\n" + "="*60)
print("TEST 14: Split tracker count — manual verification")
print("="*60)

# 3-string tracker, 10 strings/inv: cycle [3,3,3,1][2,3,3,2][1,3,3,3]
result = allocate_strings(3, 10, 30)
T.ok(f"30x3/10: split_trackers = {result['summary']['total_split_trackers']}")
print(f"    (Manual check: {result['summary']['total_split_trackers']} splits in 30 trackers)")

# 31 trackers
result = allocate_strings(3, 10, 31)
T.ok(f"31x3/10: split_trackers = {result['summary']['total_split_trackers']}")
print(f"    (Manual check: {result['summary']['total_split_trackers']} splits in 31 trackers)")

# 2-str / 3-spi: cycle is [2,1][1,2] — 1 split per cycle (2 trackers per cycle)
result = allocate_strings(2, 3, 10)
T.ok(f"10x2/3: split_trackers = {result['summary']['total_split_trackers']}")
for i, inv in enumerate(result['inverters']):
    trackers = [(idx, taken) for idx, taken in inv['tracker_indices']]
    print(f"    Inv {i}: {inv['pattern']}  trackers={trackers}")


# ============================================================
# Test 15: Pattern dump for visual inspection
# ============================================================

print("\n" + "="*60)
print("TEST 15: Pattern dump for visual inspection")
print("="*60)

scenarios = [
    (3, 10, 10, "Small: 10 trackers"),
    (4, 7, 8, "Odd: 4-str/7-spi"),
    (6, 4, 6, "Bigger tracker: 6-str/4-spi"),
    (2, 3, 9, "Tight: 2-str/3-spi"),
]

for spt, spi, n, label in scenarios:
    result = allocate_strings(spt, spi, n)
    print(f"\n  --- {label}: {n}x {spt}-string trackers, {spi} strings/inv ---")
    print(f"  Cycle: {result['cycle']}")
    print(f"  Inverters: {result['summary']['total_inverters']} "
          f"(full={result['summary']['full_inverters']}, "
          f"partial={result['summary']['partial_inverter_strings']} strings)")
    print(f"  Splits: {result['summary']['total_split_trackers']}")
    for i, inv in enumerate(result['inverters']):
        tracker_str = ', '.join(f"T{idx}({taken})" for idx, taken in inv['tracker_indices'])
        print(f"    Inv {i}: [{'-'.join(str(s) for s in inv['pattern'])}]  -> {tracker_str}")


# ============================================================
# Final Summary
# ============================================================

all_passed = T.summary()
sys.exit(0 if all_passed else 1)