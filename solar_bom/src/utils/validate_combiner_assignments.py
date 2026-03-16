"""
Combiner Assignment Validator
=============================

Validates the integrity of last_combiner_assignments data.
Can be used:
  1. Inline after calculate_estimate() or _rebuild_from_device_strings()
  2. As a pytest test module

Usage (inline):
    from src.utils.validate_combiner_assignments import validate_assignments
    issues = validate_assignments(self.last_combiner_assignments)
    if issues:
        print("[VALIDATION FAILED]")
        for issue in issues:
            print(f"  {issue}")

Usage (pytest):
    python -m pytest src/utils/validate_combiner_assignments.py -v
"""

from collections import Counter


def validate_assignments(combiner_assignments, expected_spt=None, verbose=False):
    """Validate combiner assignments for duplicates, missing strings, and bad positions.

    Args:
        combiner_assignments: list of CB dicts, each with 'connections' list.
        expected_spt: optional dict {tracker_idx: strings_per_tracker} for completeness check.
                      If None, inferred from total strings per tracker across all CBs.
        verbose: if True, print the full inventory even when valid.

    Returns:
        list of issue strings. Empty list = all good.
    """
    issues = []
    all_strings = []  # (tracker_idx, physical_pos)

    # ── 1. Collect every string from every connection ──
    for cb_idx, cb in enumerate(combiner_assignments):
        cb_name = cb.get('combiner_name', f'CB-{cb_idx + 1}')
        for conn in cb.get('connections', []):
            tidx = conn['tracker_idx']
            start = conn.get('start_string_pos', None)
            n = conn['num_strings']
            hlabel = conn.get('harness_label', '?')

            # Check: start_string_pos must exist
            if start is None:
                issues.append(
                    f"MISSING_POS: {cb_name} T{tidx+1:02d} {hlabel} has no start_string_pos"
                )
                continue

            # Check: start_string_pos must be non-negative int
            if not isinstance(start, int) or start < 0:
                issues.append(
                    f"BAD_POS: {cb_name} T{tidx+1:02d} {hlabel} start_string_pos={start}"
                )
                continue

            for s in range(start, start + n):
                all_strings.append((tidx, s))

    # ── 2. Duplicate check ──
    counts = Counter(all_strings)
    duplicates = {k: v for k, v in counts.items() if v > 1}
    if duplicates:
        for (tidx, pos), count in sorted(duplicates.items()):
            issues.append(f"DUPLICATE: T{tidx+1:02d} position {pos} appears {count} times")

    # ── 3. Missing string check ──
    # Build SPT map: either from expected_spt or inferred from data
    tracker_spt = {}
    if expected_spt:
        tracker_spt = dict(expected_spt)
    else:
        # Infer: total strings for each tracker across all CBs
        tracker_totals = {}
        for cb in combiner_assignments:
            for conn in cb.get('connections', []):
                tidx = conn['tracker_idx']
                tracker_totals[tidx] = tracker_totals.get(tidx, 0) + conn['num_strings']
        tracker_spt = tracker_totals

    for tidx, spt in sorted(tracker_spt.items()):
        expected = set(range(spt))
        actual = set(s for t, s in all_strings if t == tidx)
        missing = expected - actual
        extra = actual - expected
        if missing:
            issues.append(
                f"MISSING: T{tidx+1:02d} expected {spt} strings, "
                f"missing positions {sorted(missing)}"
            )
        if extra:
            issues.append(
                f"EXTRA: T{tidx+1:02d} expected {spt} strings, "
                f"extra positions {sorted(extra)}"
            )

    # ── 4. Harness label ordering check ──
    # Within each CB, harness labels for the same tracker should be in H01, H02, ... order
    # and their start_string_pos should be ascending
    for cb_idx, cb in enumerate(combiner_assignments):
        cb_name = cb.get('combiner_name', f'CB-{cb_idx + 1}')
        tracker_conns = {}  # tidx -> [(harness_label, start_pos, num_strings)]
        for conn in cb.get('connections', []):
            tidx = conn['tracker_idx']
            if tidx not in tracker_conns:
                tracker_conns[tidx] = []
            tracker_conns[tidx].append((
                conn.get('harness_label', '?'),
                conn.get('start_string_pos', -1),
                conn['num_strings'],
            ))

        for tidx, conns in tracker_conns.items():
            if len(conns) < 2:
                continue
            # Check position ordering
            positions = [c[1] for c in conns if c[1] >= 0]
            if positions and positions != sorted(positions):
                issues.append(
                    f"ORDER: {cb_name} T{tidx+1:02d} harness positions not ascending: "
                    f"{[(c[0], c[1]) for c in conns]}"
                )

    # ── 5. Contiguity check ──
    # Each device's strings from a given tracker should be contiguous
    for cb_idx, cb in enumerate(combiner_assignments):
        cb_name = cb.get('combiner_name', f'CB-{cb_idx + 1}')
        tracker_positions = {}  # tidx -> set of positions
        for conn in cb.get('connections', []):
            tidx = conn['tracker_idx']
            start = conn.get('start_string_pos', None)
            if start is None:
                continue
            if tidx not in tracker_positions:
                tracker_positions[tidx] = set()
            for s in range(start, start + conn['num_strings']):
                tracker_positions[tidx].add(s)

        for tidx, positions in tracker_positions.items():
            if not positions:
                continue
            min_p = min(positions)
            max_p = max(positions)
            expected_contiguous = set(range(min_p, max_p + 1))
            if positions != expected_contiguous:
                gaps = expected_contiguous - positions
                issues.append(
                    f"CONTIGUITY: {cb_name} T{tidx+1:02d} has gaps at positions {sorted(gaps)}"
                )

    # ── Verbose output ──
    if verbose or issues:
        print("\n[VALIDATE] === Combiner Assignment Inventory ===")
        for cb_idx, cb in enumerate(combiner_assignments):
            cb_name = cb.get('combiner_name', f'CB-{cb_idx + 1}')
            print(f"  {cb_name}:")
            for conn in cb.get('connections', []):
                tidx = conn['tracker_idx']
                start = conn.get('start_string_pos', '?')
                n = conn['num_strings']
                hlabel = conn.get('harness_label', '?')
                print(f"    T{tidx+1:02d} {hlabel}: {n} strings, start_pos={start}")

        if issues:
            print(f"  [{len(issues)} ISSUES FOUND]:")
            for issue in issues:
                print(f"    {issue}")
        else:
            print("  [ALL CHECKS PASSED]")
        print("[VALIDATE] === End ===\n")

    return issues


def print_inventory(combiner_assignments):
    """Convenience: just print the inventory with full validation."""
    validate_assignments(combiner_assignments, verbose=True)


# ═══════════════════════════════════════════════════════════════
# Pytest tests using known-good data from development testing
# ═══════════════════════════════════════════════════════════════

def _make_conn(tidx, hlabel, num_strings, start_pos):
    """Helper to build a connection dict for testing."""
    return {
        'tracker_idx': tidx,
        'tracker_label': f'T{tidx+1:02d}',
        'harness_label': hlabel,
        'num_strings': num_strings,
        'start_string_pos': start_pos,
        'module_isc': 18.42,
        'nec_factor': 1.56,
        'wire_gauge': '10 AWG',
    }


def _make_cb(name, connections):
    """Helper to build a CB dict for testing."""
    return {
        'combiner_name': name,
        'device_idx': 0,
        'breaker_size': 400,
        'module_isc': 18.42,
        'nec_factor': 1.56,
        'connections': connections,
    }


class TestValidateAssignments:
    """Test the validator itself with known scenarios."""

    def test_clean_non_split(self):
        """Non-split trackers with correct positions pass validation."""
        assignments = [
            _make_cb('CB-01', [
                _make_conn(0, 'H01', 6, 0),
                _make_conn(0, 'H02', 2, 6),  # 8-string, config [6,2]
                _make_conn(1, 'H01', 6, 0),
                _make_conn(1, 'H02', 7, 6),  # 13-string, config [6,7]
            ]),
        ]
        issues = validate_assignments(assignments)
        assert issues == [], f"Unexpected issues: {issues}"

    def test_clean_split_tracker(self):
        """Split tracker across two CBs with correct positions passes."""
        assignments = [
            _make_cb('CB-01', [
                _make_conn(0, 'H01', 6, 0),   # T01 head: 6 strings from pos 0
            ]),
            _make_cb('CB-02', [
                _make_conn(0, 'H01', 4, 6),   # T01 tail: 4 strings from pos 6
            ]),
        ]
        issues = validate_assignments(assignments)
        assert issues == [], f"Unexpected issues: {issues}"

    def test_duplicate_detected(self):
        """Overlapping start_string_pos produces DUPLICATE issue."""
        assignments = [
            _make_cb('CB-01', [
                _make_conn(0, 'H01', 6, 0),
                _make_conn(0, 'H02', 3, 0),  # BUG: overlaps H01
            ]),
        ]
        issues = validate_assignments(assignments)
        assert any('DUPLICATE' in i for i in issues), f"Expected DUPLICATE, got: {issues}"

    def test_missing_detected(self):
        """Gap in positions produces MISSING issue."""
        assignments = [
            _make_cb('CB-01', [
                _make_conn(0, 'H01', 6, 0),
                # H02 missing — positions 6-7 not covered for an 8-string tracker
            ]),
        ]
        # Provide expected SPT so validator knows what to expect
        issues = validate_assignments(assignments, expected_spt={0: 8})
        assert any('MISSING' in i for i in issues), f"Expected MISSING, got: {issues}"

    def test_missing_start_pos(self):
        """Connection without start_string_pos produces MISSING_POS issue."""
        assignments = [
            _make_cb('CB-01', [
                {
                    'tracker_idx': 0,
                    'tracker_label': 'T01',
                    'harness_label': 'H01',
                    'num_strings': 6,
                    # No start_string_pos!
                    'module_isc': 18.42,
                    'nec_factor': 1.56,
                    'wire_gauge': '10 AWG',
                },
            ]),
        ]
        issues = validate_assignments(assignments)
        assert any('MISSING_POS' in i for i in issues), f"Expected MISSING_POS, got: {issues}"

    def test_contiguity_gap(self):
        """Non-contiguous positions within a CB produce CONTIGUITY issue."""
        assignments = [
            _make_cb('CB-01', [
                _make_conn(0, 'H01', 3, 0),
                _make_conn(0, 'H02', 3, 5),  # Gap at positions 3,4
            ]),
        ]
        issues = validate_assignments(assignments)
        assert any('CONTIGUITY' in i for i in issues), f"Expected CONTIGUITY, got: {issues}"

    def test_harness_order(self):
        """Out-of-order harness positions within a CB produce ORDER issue."""
        assignments = [
            _make_cb('CB-01', [
                _make_conn(0, 'H02', 4, 6),   # H02 before H01
                _make_conn(0, 'H01', 6, 0),
            ]),
        ]
        issues = validate_assignments(assignments)
        assert any('ORDER' in i for i in issues), f"Expected ORDER, got: {issues}"

    def test_real_project_snapshot(self):
        """Snapshot from actual project debug output — should pass all checks."""
        assignments = [
            _make_cb('CB-01', [
                _make_conn(0, 'H01', 6, 0),
                _make_conn(0, 'H02', 2, 6),
                _make_conn(1, 'H01', 6, 0),
                _make_conn(1, 'H02', 2, 6),
                _make_conn(2, 'H01', 6, 0),
                _make_conn(2, 'H02', 3, 6),
                _make_conn(3, 'H01', 6, 0),
                _make_conn(3, 'H02', 3, 6),
                _make_conn(4, 'H01', 6, 0),
                _make_conn(4, 'H02', 4, 6),
                _make_conn(5, 'H01', 6, 0),
                _make_conn(5, 'H02', 4, 6),
                _make_conn(6, 'H01', 6, 0),   # T07 head
            ]),
            _make_cb('CB-02', [
                _make_conn(6, 'H01', 4, 6),   # T07 tail
                _make_conn(7, 'H01', 6, 0),
                _make_conn(7, 'H02', 5, 6),
                _make_conn(8, 'H01', 6, 0),
                _make_conn(8, 'H02', 5, 6),
                _make_conn(9, 'H01', 6, 0),
                _make_conn(9, 'H02', 5, 6),
                _make_conn(10, 'H01', 6, 0),
                _make_conn(10, 'H02', 6, 6),
                _make_conn(11, 'H01', 6, 0),  # T12 head (11 strings)
                _make_conn(11, 'H02', 5, 6),
            ]),
            _make_cb('CB-03', [
                _make_conn(11, 'H01', 1, 11),  # T12 tail (1 string)
                _make_conn(12, 'H01', 6, 0),
                _make_conn(12, 'H02', 6, 6),
                _make_conn(13, 'H01', 6, 0),
                _make_conn(13, 'H02', 7, 6),
                _make_conn(14, 'H01', 6, 0),
                _make_conn(14, 'H02', 7, 6),
                _make_conn(15, 'H01', 6, 0),
                _make_conn(15, 'H02', 7, 6),
                _make_conn(16, 'H01', 6, 0),  # T17 head (8 strings)
                _make_conn(16, 'H02', 2, 6),
            ]),
            _make_cb('CB-04', [
                _make_conn(16, 'H01', 5, 8),  # T17 tail (5 strings)
                _make_conn(17, 'H01', 6, 0),
                _make_conn(17, 'H02', 7, 6),
                _make_conn(18, 'H01', 6, 0),
                _make_conn(18, 'H02', 7, 6),
                _make_conn(19, 'H01', 6, 0),
                _make_conn(19, 'H02', 7, 6),
                _make_conn(20, 'H01', 2, 0),
                _make_conn(20, 'H02', 7, 2),
                _make_conn(21, 'H01', 7, 0),
            ]),
        ]

        expected_spt = {
            0: 8, 1: 8, 2: 9, 3: 9, 4: 10, 5: 10,
            6: 10,  # split
            7: 11, 8: 11, 9: 11, 10: 12,
            11: 12,  # split
            12: 12, 13: 13, 14: 13, 15: 13,
            16: 13,  # split
            17: 13, 18: 13, 19: 13, 20: 9, 21: 7,
        }

        issues = validate_assignments(assignments, expected_spt=expected_spt)
        assert issues == [], f"Snapshot validation failed:\n" + "\n".join(issues)

    def test_edit_devices_moved_string(self):
        """After moving 1 string from T12 to CB-03, split harness should be 5+6 not 6+5."""
        assignments = [
            _make_cb('CB-02', [
                _make_conn(11, 'H01', 5, 1),  # T12: lost string 0, now starts at 1
                _make_conn(11, 'H02', 6, 6),  # Full south harness
            ]),
            _make_cb('CB-03', [
                _make_conn(11, 'H01', 1, 0),  # T12's moved string
            ]),
        ]
        expected_spt = {11: 12}
        issues = validate_assignments(assignments, expected_spt=expected_spt)
        assert issues == [], f"Edit Devices scenario failed:\n" + "\n".join(issues)


# Allow running directly: python validate_combiner_assignments.py
if __name__ == '__main__':
    import sys
    # Run the real project snapshot test
    test = TestValidateAssignments()
    
    tests = [
        ('clean_non_split', test.test_clean_non_split),
        ('clean_split_tracker', test.test_clean_split_tracker),
        ('duplicate_detected', test.test_duplicate_detected),
        ('missing_detected', test.test_missing_detected),
        ('missing_start_pos', test.test_missing_start_pos),
        ('contiguity_gap', test.test_contiguity_gap),
        ('harness_order', test.test_harness_order),
        ('real_project_snapshot', test.test_real_project_snapshot),
        ('edit_devices_moved_string', test.test_edit_devices_moved_string),
    ]
    
    passed = 0
    failed = 0
    for name, fn in tests:
        try:
            fn()
            print(f"  PASS: {name}")
            passed += 1
        except AssertionError as e:
            print(f"  FAIL: {name}")
            print(f"        {e}")
            failed += 1
    
    print(f"\n{passed} passed, {failed} failed")
    sys.exit(1 if failed else 0)