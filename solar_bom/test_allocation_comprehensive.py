"""
Comprehensive Test Suite for String-to-Inverter Allocation Engine (Balanced)
"""

import sys
import os
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.utils.string_allocation import compute_allocation_cycle, allocate_strings, format_allocation_summary
from math import gcd, ceil

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


def check_invariants(result, spt, spi, num_trackers, label):
    """Run invariant checks for balanced allocation."""
    ok = True
    total_strings = spt * num_trackers
    summary = result['summary']
    
    # 1. String conservation
    if summary['total_strings'] != total_strings:
        T.fail(f"{label} — string conservation",
               f"Expected {total_strings} strings, got {summary['total_strings']}")
        ok = False
    else:
        T.ok(f"{label} — string conservation ({total_strings} strings)")
    
    # 2. No inverter exceeds max spi
    for i, inv in enumerate(result['inverters']):
        if inv['total_strings'] > spi:
            T.fail(f"{label} — inverter {i} overflow",
                   f"Inverter has {inv['total_strings']} strings, max is {spi}")
            ok = False
    
    # 3. Balanced: all inverters within 1 string of each other
    sizes = [inv['total_strings'] for inv in result['inverters']]
    if sizes:
        if max(sizes) - min(sizes) > 1:
            T.fail(f"{label} — balance check",
                   f"Inverter sizes range from {min(sizes)} to {max(sizes)} (gap > 1)")
            ok = False
    
    # 4. No zero or negative values in patterns
    for i, inv in enumerate(result['inverters']):
        for val in inv['pattern']:
            if val <= 0:
                T.fail(f"{label} — non-positive in pattern",
                       f"Inverter {i} pattern has value {val}")
                ok = False
    
    # 5. Pattern sums match total_strings per inverter
    for i, inv in enumerate(result['inverters']):
        pattern_sum = sum(inv['pattern'])
        if pattern_sum != inv['total_strings']:
            T.fail(f"{label} — pattern sum mismatch",
                   f"Inverter {i}: pattern sums to {pattern_sum}, total_strings={inv['total_strings']}")
            ok = False
    
    # 6. Tracker indices in bounds
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
    
    # 7. Each tracker fully consumed
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
    
    # 9. Balanced summary fields correct
    n_inv = len(result['inverters'])
    if n_inv > 0:
        expected_n = ceil(total_strings / spi)
        if n_inv != expected_n:
            T.fail(f"{label} — inverter count",
                   f"Expected {expected_n}, got {n_inv}")
            ok = False
        
        base = total_strings // n_inv
        remainder = total_strings % n_inv
        expected_max = base + 1 if remainder > 0 else base
        expected_min = base
        
        if summary.get('max_strings_per_inverter') != expected_max:
            T.fail(f"{label} — max_spi",
                   f"Expected {expected_max}, got {summary.get('max_strings_per_inverter')}")
            ok = False
        if summary.get('min_strings_per_inverter') != expected_min:
            T.fail(f"{label} — min_spi",
                   f"Expected {expected_min}, got {summary.get('min_strings_per_inverter')}")
            ok = False
        if summary.get('num_larger_inverters') != remainder:
            T.fail(f"{label} — num_larger",
                   f"Expected {remainder}, got {summary.get('num_larger_inverters')}")
            ok = False
    
    if ok:
        T.ok(f"{label} — all invariants passed")
    
    return ok


# ============================================================
# Test 1: compute_allocation_cycle (unchanged)
# ============================================================

print("\n" + "="*60)
print("TEST 1: compute_allocation_cycle")
print("="*60)

cycle = compute_allocation_cycle(3, 10)
expected = [[3, 3, 3, 1], [2, 3, 3, 2], [1, 3, 3, 3]]
if cycle == expected:
    T.ok("3/10 cycle matches expected pattern")
else:
    T.fail("3/10 cycle", f"Expected {expected}, got {cycle}")

for spt, spi in [(2, 10), (3, 10), (4, 7), (5, 12), (1, 8), (6, 4), (3, 3), (4, 4)]:
    cycle = compute_allocation_cycle(spt, spi)
    all_ok = all(sum(p) == spi for p in cycle)
    if all_ok:
        T.ok(f"cycle {spt}/{spi}: all {len(cycle)} patterns sum to {spi}")
    else:
        T.fail(f"cycle {spt}/{spi} pattern sum", "Pattern sum mismatch")

for spt, spi in [(2, 10), (3, 10), (4, 7), (5, 12), (6, 9), (3, 3)]:
    cycle = compute_allocation_cycle(spt, spi)
    lcm = (spt * spi) // gcd(spt, spi)
    expected_len = lcm // spi
    if len(cycle) == expected_len:
        T.ok(f"cycle {spt}/{spi}: length {len(cycle)} == LCM({lcm})/{spi}")
    else:
        T.fail(f"cycle {spt}/{spi} length", f"Expected {expected_len}, got {len(cycle)}")

cycle = compute_allocation_cycle(10, 10)
if cycle == [[10]]:
    T.ok("10/10: single pattern [10]")
else:
    T.fail("10/10 cycle", f"Expected [[10]], got {cycle}")

cycle = compute_allocation_cycle(1, 10)
if cycle == [[1]*10]:
    T.ok("1/10: pattern is ten 1s")
else:
    T.fail("1/10 cycle", f"Got {cycle}")

cycle = compute_allocation_cycle(3, 1)
if cycle == [[1], [1], [1]]:
    T.ok("3/1: three patterns of [1]")
else:
    T.fail("3/1 cycle", f"Got {cycle}")

if compute_allocation_cycle(0, 10) == []:
    T.ok("0/10: returns empty")
else:
    T.fail("0/10", "Not empty")

if compute_allocation_cycle(3, 0) == []:
    T.ok("3/0: returns empty")
else:
    T.fail("3/0", "Not empty")


# ============================================================
# Test 2: Known test case — balanced
# ============================================================

print("\n" + "="*60)
print("TEST 2: Known test case — balanced (3-str, max 10 str/inv)")
print("="*60)

# 30 trackers: 90/10 = 9 inverters, all get 10 (evenly divisible)
result = allocate_strings(3, 10, 30)
check_invariants(result, 3, 10, 30, "30x3str/10spi")

# 31 trackers: 93 strings, ceil(93/10)=10 inv, 93/10 = 9r3 -> 3x10 + 7x9
result = allocate_strings(3, 10, 31)
check_invariants(result, 3, 10, 31, "31x3str/10spi")
s = result['summary']
if s['num_larger_inverters'] == 3 and s['max_strings_per_inverter'] == 10:
    T.ok("31x3/10: 3 inverters get 10 strings")
else:
    T.fail("31x3/10 larger", f"Got {s['num_larger_inverters']}x{s['max_strings_per_inverter']}")
if s['num_smaller_inverters'] == 7 and s['min_strings_per_inverter'] == 9:
    T.ok("31x3/10: 7 inverters get 9 strings")
else:
    T.fail("31x3/10 smaller", f"Got {s['num_smaller_inverters']}x{s['min_strings_per_inverter']}")

# 7 trackers: 21 strings, ceil(21/10)=3 inv, 21/3 = 7r0 -> all get 7
result = allocate_strings(3, 10, 7)
check_invariants(result, 3, 10, 7, "7x3str/10spi")
if result['summary']['max_strings_per_inverter'] == 7 and result['summary']['min_strings_per_inverter'] == 7:
    T.ok("7x3/10: all 3 inverters get 7 strings")
else:
    T.fail("7x3/10 sizes", f"max={result['summary']['max_strings_per_inverter']}, min={result['summary']['min_strings_per_inverter']}")


# ============================================================
# Test 3: Even division — no splits
# ============================================================

print("\n" + "="*60)
print("TEST 3: Even division (no splits expected)")
print("="*60)

result = allocate_strings(2, 10, 50)
check_invariants(result, 2, 10, 50, "50x2str/10spi")
if result['summary']['total_split_trackers'] == 0:
    T.ok("50x2/10: zero splits")
else:
    T.fail("50x2/10 splits", f"Got {result['summary']['total_split_trackers']}")

result = allocate_strings(5, 10, 20)
check_invariants(result, 5, 10, 20, "20x5str/10spi")

result = allocate_strings(4, 8, 10)
check_invariants(result, 4, 8, 10, "10x4str/8spi")


# ============================================================
# Test 4: Tracker strings > strings/inverter
# ============================================================

print("\n" + "="*60)
print("TEST 4: Tracker > inverter capacity")
print("="*60)

result = allocate_strings(6, 4, 10)
check_invariants(result, 6, 4, 10, "10x6str/4spi")

result = allocate_strings(10, 3, 5)
check_invariants(result, 10, 3, 5, "5x10str/3spi")


# ============================================================
# Test 5: 1-string trackers
# ============================================================

print("\n" + "="*60)
print("TEST 5: 1-string trackers")
print("="*60)

result = allocate_strings(1, 10, 25)
check_invariants(result, 1, 10, 25, "25x1str/10spi")
if result['summary']['total_split_trackers'] == 0:
    T.ok("25x1/10: zero splits")
else:
    T.fail("25x1/10 splits", f"Got {result['summary']['total_split_trackers']}")

result = allocate_strings(1, 1, 5)
check_invariants(result, 1, 1, 5, "5x1str/1spi")


# ============================================================
# Test 6: Single tracker
# ============================================================

print("\n" + "="*60)
print("TEST 6: Single tracker")
print("="*60)

result = allocate_strings(3, 10, 1)
check_invariants(result, 3, 10, 1, "1x3str/10spi")
if result['summary']['total_inverters'] == 1:
    T.ok("1x3/10: 1 inverter")
else:
    T.fail("1x3/10 count", f"Got {result['summary']['total_inverters']}")


# ============================================================
# Test 7: Large projects
# ============================================================

print("\n" + "="*60)
print("TEST 7: Large projects")
print("="*60)

result = allocate_strings(3, 10, 500)
check_invariants(result, 3, 10, 500, "500x3str/10spi")

result = allocate_strings(4, 12, 1000)
check_invariants(result, 4, 12, 1000, "1000x4str/12spi")

t0 = time.time()
result = allocate_strings(3, 10, 5000)
dt = time.time() - t0
check_invariants(result, 3, 10, 5000, "5000x3str/10spi")
if dt < 1.0:
    T.ok(f"5000x3/10: completed in {dt*1000:.1f}ms")
else:
    T.fail("5000x3/10 perf", f"Took {dt:.2f}s")


# ============================================================
# Test 8: Real-world odd combos
# ============================================================

print("\n" + "="*60)
print("TEST 8: Real-world odd combos")
print("="*60)

test_cases = [
    (4, 7, 50, "4str/7spi"),
    (3, 8, 40, "3str/8spi"),
    (5, 12, 30, "5str/12spi"),
    (2, 7, 100, "2str/7spi"),
    (3, 13, 20, "3str/13spi"),
    (6, 9, 45, "6str/9spi"),
    (4, 10, 60, "4str/10spi"),
    (3, 4, 25, "3str/4spi"),
    (7, 11, 33, "7str/11spi"),
    (2, 3, 150, "2str/3spi"),
]

for spt, spi, n, label in test_cases:
    result = allocate_strings(spt, spi, n)
    check_invariants(result, spt, spi, n, label)


# ============================================================
# Test 9: spt == spi (1:1 mapping)
# ============================================================

print("\n" + "="*60)
print("TEST 9: Tracker exactly fills one inverter")
print("="*60)

for spt in [2, 3, 5, 10]:
    result = allocate_strings(spt, spt, 20)
    check_invariants(result, spt, spt, 20, f"20x{spt}str/{spt}spi")
    if result['summary']['total_split_trackers'] == 0:
        T.ok(f"20x{spt}/{spt}: zero splits")
    else:
        T.fail(f"20x{spt}/{spt} splits", f"Got {result['summary']['total_split_trackers']}")


# ============================================================
# Test 10: Micro-inverter (1 str/inv)
# ============================================================

print("\n" + "="*60)
print("TEST 10: 1 string per inverter")
print("="*60)

result = allocate_strings(3, 1, 10)
check_invariants(result, 3, 1, 10, "10x3str/1spi")
if result['summary']['total_split_trackers'] == 10:
    T.ok("10x3/1: all 10 trackers split")
else:
    T.fail("10x3/1 splits", f"Got {result['summary']['total_split_trackers']}")


# ============================================================
# Test 11: Invalid inputs
# ============================================================

print("\n" + "="*60)
print("TEST 11: Invalid inputs")
print("="*60)

for spt, spi, n, label in [(0, 10, 5, "spt=0"), (3, 0, 5, "spi=0"), (3, 10, 0, "n=0")]:
    result = allocate_strings(spt, spi, n)
    if result['summary']['total_inverters'] == 0:
        T.ok(f"{label}: returns empty")
    else:
        T.fail(label, f"Got {result['summary']['total_inverters']} inverters")


# ============================================================
# Test 12: Detail dump for visual review
# ============================================================

print("\n" + "="*60)
print("TEST 12: Detail dump for visual review")
print("="*60)

scenarios = [
    (3, 10, 31, "31x 3-str, max 10/inv"),
    (4, 7, 8, "8x 4-str, max 7/inv"),
    (6, 4, 6, "6x 6-str, max 4/inv"),
    (2, 3, 9, "9x 2-str, max 3/inv"),
    (3, 10, 10, "10x 3-str, max 10/inv"),
    (1, 10, 25, "25x 1-str, max 10/inv"),
]

for spt, spi, n, label in scenarios:
    result = allocate_strings(spt, spi, n)
    s = result['summary']
    print(f"\n  --- {label} ---")
    print(f"  {s['total_inverters']} inverters: "
          f"{s['num_larger_inverters']}x{s['max_strings_per_inverter']} + "
          f"{s['num_smaller_inverters']}x{s['min_strings_per_inverter']}")
    print(f"  Splits: {s['total_split_trackers']}")
    for i, inv in enumerate(result['inverters']):
        tracker_str = ', '.join(f"T{idx}({taken})" for idx, taken in inv['tracker_indices'])
        print(f"    Inv {i} ({inv['total_strings']}str): [{'-'.join(str(x) for x in inv['pattern'])}]  -> {tracker_str}")
    T.ok(f"{label}: printed for review")


# ============================================================
# Final Summary
# ============================================================

all_passed = T.summary()
sys.exit(0 if all_passed else 1)