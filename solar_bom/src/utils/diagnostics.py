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
# Extender Diagnostics
# ═══════════════════════════════════════════════════════════════

def validate_extenders(totals, groups, tracker_seg_map, split_tracker_details,
                       enabled_templates=None, lv_collection_method='Wire Harness',
                       max_ns_offset=0, verbose=False):
    """Validate extender BOM data using physics-based invariants.

    These checks don't rely on hardcoded expected values — they verify
    relationships that must always hold regardless of project inputs.

    Args:
        totals: the totals dict from calculate_estimate containing
                'extenders_pos_by_length' and 'extenders_neg_by_length'.
        groups: the QE groups list (each with 'segments', 'device_position').
        tracker_seg_map: list of dicts from _tracker_to_segment.
        split_tracker_details: dict from _split_tracker_details.
        enabled_templates: dict of template data (for tracker length calc).
        verbose: if True, print details even when passing.

    Returns:
        list of issue strings. Empty list = all good.
    """
    issues = []
    pos_by_length = totals.get('extenders_pos_by_length', {})
    neg_by_length = totals.get('extenders_neg_by_length', {})

    # ── 1. Total positive count must equal total negative count ──
    total_pos = sum(pos_by_length.values())
    total_neg = sum(neg_by_length.values())
    if total_pos != total_neg:
        issues.append(
            f"EXT_COUNT_MISMATCH: Total positive extenders ({total_pos}) "
            f"!= total negative extenders ({total_neg})"
        )

    # ── 2. No extender length should be zero or negative ──
    for (length, gauge), qty in pos_by_length.items():
        if length <= 0:
            issues.append(
                f"EXT_BAD_LENGTH: Positive extender has length={length}ft "
                f"(gauge={gauge}, qty={qty})"
            )
    for (length, gauge), qty in neg_by_length.items():
        if length <= 0:
            issues.append(
                f"EXT_BAD_LENGTH: Negative extender has length={length}ft "
                f"(gauge={gauge}, qty={qty})"
            )

    # ── 3. Extender count should equal total harness count across all trackers ──
    # Every harness on every tracker produces exactly one pos and one neg extender.
    expected_harness_count = 0
    split_tidxs = set(split_tracker_details.keys()) if split_tracker_details else set()

    # Count from non-split trackers (bulk segments)
    split_seg_counts = {}  # (group_idx, id(seg)) -> number of splits
    for tidx in split_tidxs:
        if tidx < len(tracker_seg_map):
            info = tracker_seg_map[tidx]
            key = (info['group_idx'], id(info['seg']))
            split_seg_counts[key] = split_seg_counts.get(key, 0) + 1

    for group_idx, group in enumerate(groups):
        for seg in group.get('segments', []):
            qty = seg.get('quantity', 0)
            if qty <= 0:
                continue
            key = (group_idx, id(seg))
            num_splits = split_seg_counts.get(key, 0)
            non_split_qty = qty - num_splits

            # Parse harness config to count harnesses per tracker
            # For String HR, every string is its own harness regardless of stored config
            harness_config = seg.get('harness_config', str(seg.get('strings_per_tracker', 1)))
            if lv_collection_method == 'String HR':
                num_harnesses = int(seg.get('strings_per_tracker', 1))
            else:
                num_harnesses = len(harness_config.split('+'))
            expected_harness_count += non_split_qty * num_harnesses

    # Count from split tracker portions
    for tidx, details in (split_tracker_details or {}).items():
        for portion in details.get('portions', []):
            expected_harness_count += len(portion.get('harnesses', []))

    if total_pos != expected_harness_count:
        issues.append(
            f"EXT_HARNESS_COUNT: Expected {expected_harness_count} extenders "
            f"(one per harness), got {total_pos} positive"
        )
    if total_neg != expected_harness_count:
        issues.append(
            f"EXT_HARNESS_COUNT: Expected {expected_harness_count} extenders "
            f"(one per harness), got {total_neg} negative"
        )

    # ── 4. For device_position='middle', non-split extender pairs should ──
    #    show a pos/neg length difference ≈ one string length (~20-40ft).
    #    This is a soft check (warning, not failure) since rounding affects it.
    for group_idx, group in enumerate(groups):
        device_pos = group.get('device_position', 'middle')
        if device_pos != 'middle':
            continue
        for seg in group.get('segments', []):
            harness_config = seg.get('harness_config', str(seg.get('strings_per_tracker', 1)))
            harness_sizes = [int(float(x)) for x in harness_config.split('+')]
            if len(harness_sizes) < 2:
                continue
            # For 2-harness configs with device=middle, H0 and H1 should have
            # complementary pos/neg patterns (one ~25ft, other ~5ft)
            # We don't enforce this strictly since split trackers differ

    # ── 5. No extender should exceed the longest tracker in the project ──
    max_tracker_length_ft = 0
    if enabled_templates:
        m_to_ft = 3.28084
        for tname, tdata in enabled_templates.items():
            mod_spec = tdata.get('module_spec', {})
            orientation = tdata.get('module_orientation', 'Portrait')
            mps = tdata.get('modules_per_string', 28)
            spacing_m = tdata.get('module_spacing_m', 0.02)
            has_motor = tdata.get('has_motor', True)
            motor_gap_m = tdata.get('motor_gap_m', 1.0) if has_motor else 0
            spt = tdata.get('strings_per_tracker', 1)

            if orientation == 'Portrait':
                mod_along_m = mod_spec.get('width_mm', 1000) / 1000
            else:
                mod_along_m = mod_spec.get('length_mm', 2000) / 1000

            full_spt = int(spt)
            total_mods = full_spt * mps
            tracker_len_m = (total_mods * mod_along_m +
                             (total_mods - 1) * spacing_m +
                             motor_gap_m)
            tracker_len_ft = tracker_len_m * m_to_ft
            max_tracker_length_ft = max(max_tracker_length_ft, tracker_len_ft)

    max_allowed = (max_tracker_length_ft + max_ns_offset) * 1.1 if max_tracker_length_ft > 0 else 0
    if max_allowed > 0:
        for (length, gauge), qty in pos_by_length.items():
            if length > max_allowed:
                issues.append(
                    f"EXT_TOO_LONG: Positive extender {length}ft exceeds "
                    f"max allowed ({max_allowed:.0f}ft = tracker {max_tracker_length_ft:.0f}ft + "
                    f"inter-row {max_ns_offset:.0f}ft + 10%), gauge={gauge}, qty={qty}"
                )
        for (length, gauge), qty in neg_by_length.items():
            if length > max_allowed:
                issues.append(
                    f"EXT_TOO_LONG: Negative extender {length}ft exceeds "
                    f"max allowed ({max_allowed:.0f}ft = tracker {max_tracker_length_ft:.0f}ft + "
                    f"inter-row {max_ns_offset:.0f}ft + 10%), gauge={gauge}, qty={qty}"
                )

    # ── Verbose output ──
    if verbose or issues:
        print("\n[EXTENDER DIAG] === Extender Validation ===")
        print(f"  Total positive extenders: {total_pos}")
        print(f"  Total negative extenders: {total_neg}")
        print(f"  Expected (from harness count): {expected_harness_count}")
        if pos_by_length:
            print(f"  Positive breakdown: {dict(pos_by_length)}")
        if neg_by_length:
            print(f"  Negative breakdown: {dict(neg_by_length)}")
        if max_tracker_length_ft > 0:
            print(f"  Max tracker length: {max_tracker_length_ft:.0f}ft")
        if issues:
            print(f"  [{len(issues)} ISSUES]:")
            for issue in issues:
                print(f"    {issue}")
        else:
            print("  [ALL CHECKS PASSED]")
        print("[EXTENDER DIAG] === End ===\n")

    return issues


# ═══════════════════════════════════════════════════════════════
# Split Tracker Diagnostics
# ═══════════════════════════════════════════════════════════════

def validate_split_details(split_tracker_details, tracker_seg_map, verbose=False):
    """Validate split tracker portion data for internal consistency.

    Args:
        split_tracker_details: dict {tidx: {'spt': int, 'portions': [...]}}
        tracker_seg_map: list of dicts from _tracker_to_segment.
        verbose: if True, print details even when passing.

    Returns:
        list of issue strings. Empty list = all good.
    """
    issues = []

    if not split_tracker_details:
        if verbose:
            print("\n[SPLIT DIAG] No split trackers to validate.\n")
        return issues

    for tidx, details in split_tracker_details.items():
        spt = details.get('spt', 0)
        portions = details.get('portions', [])
        label = f"T{tidx + 1:02d}"

        # ── 1. Portion string counts must sum to total spt ──
        total_strings = sum(p.get('strings_taken', 0) for p in portions)
        if total_strings != int(spt):
            issues.append(
                f"SPLIT_SUM: {label} portions sum to {total_strings} strings, "
                f"expected {int(spt)}"
            )

        # ── 2. Harness counts within each portion must sum to strings_taken ──
        for p_idx, portion in enumerate(portions):
            harnesses = portion.get('harnesses', [])
            h_sum = sum(harnesses)
            expected = portion.get('strings_taken', 0)
            if h_sum != expected:
                issues.append(
                    f"SPLIT_HARNESS_SUM: {label} portion {p_idx} "
                    f"(inv_idx={portion.get('inv_idx', '?')}) "
                    f"harnesses sum to {h_sum}, expected {expected}"
                )

        # ── 3. start_pos should not overlap between portions ──
        all_positions = set()
        for p_idx, portion in enumerate(portions):
            start = portion.get('start_pos', None)
            if start is None:
                issues.append(
                    f"SPLIT_NO_START: {label} portion {p_idx} missing start_pos"
                )
                continue
            count = portion.get('strings_taken', 0)
            for s in range(start, start + count):
                if s in all_positions:
                    issues.append(
                        f"SPLIT_OVERLAP: {label} position {s} claimed by "
                        f"multiple portions"
                    )
                all_positions.add(s)

        # ── 4. All positions 0..spt-1 should be covered ──
        expected_positions = set(range(int(spt)))
        missing = expected_positions - all_positions
        extra = all_positions - expected_positions
        if missing:
            issues.append(
                f"SPLIT_MISSING: {label} missing positions {sorted(missing)}"
            )
        if extra:
            issues.append(
                f"SPLIT_EXTRA: {label} extra positions {sorted(extra)}"
            )

        # ── 5. No zero-size harness ──
        for p_idx, portion in enumerate(portions):
            for h_idx, h_size in enumerate(portion.get('harnesses', [])):
                if h_size <= 0:
                    issues.append(
                        f"SPLIT_ZERO_HARNESS: {label} portion {p_idx} "
                        f"harness {h_idx} has size {h_size}"
                    )

    # ── Verbose output ──
    if verbose or issues:
        print("\n[SPLIT DIAG] === Split Tracker Validation ===")
        for tidx, details in sorted(split_tracker_details.items()):
            label = f"T{tidx + 1:02d}"
            spt = details.get('spt', 0)
            print(f"  {label} (spt={spt}):")
            for p_idx, portion in enumerate(details.get('portions', [])):
                print(f"    Portion {p_idx}: inv_idx={portion.get('inv_idx', '?')}, "
                      f"start_pos={portion.get('start_pos', '?')}, "
                      f"strings={portion.get('strings_taken', '?')}, "
                      f"harnesses={portion.get('harnesses', [])}")
        if issues:
            print(f"  [{len(issues)} ISSUES]:")
            for issue in issues:
                print(f"    {issue}")
        else:
            print("  [ALL CHECKS PASSED]")
        print("[SPLIT DIAG] === End ===\n")

    return issues


# ═══════════════════════════════════════════════════════════════
# Whip Diagnostics
# ═══════════════════════════════════════════════════════════════

def validate_whips(totals, groups=None, split_tracker_details=None, tracker_seg_map=None,
                   lv_collection_method='Wire Harness', verbose=False):
    """Validate whip BOM data using invariants.

    Args:
        totals: the totals dict from calculate_estimate containing 'whips_by_length'.
        groups: the QE groups list (for expected count calculation).
        split_tracker_details: dict from _split_tracker_details.
        tracker_seg_map: list of dicts from _tracker_to_segment.
        lv_collection_method: 'Wire Harness' or 'String HR'.
        verbose: if True, print details even when passing.

    Returns:
        list of issue strings. Empty list = all good.
    """
    issues = []
    whips_by_length = totals.get('whips_by_length', {})

    # ── 1. Every whip entry should have qty divisible by 2 (pos + neg) ──
    for (length, gauge), qty in whips_by_length.items():
        if qty % 2 != 0:
            issues.append(
                f"WHIP_ODD_QTY: Whip {length}ft ({gauge}) has qty={qty}, "
                f"expected even (pos+neg pairs)"
            )

    # ── 2. No zero or negative lengths ──
    for (length, gauge), qty in whips_by_length.items():
        if length <= 0:
            issues.append(
                f"WHIP_BAD_LENGTH: Whip has length={length}ft (gauge={gauge}, qty={qty})"
            )

    # ── 3. Total whip length should match sum of (length * qty) ──
    calculated_total = sum(length * qty for (length, gauge), qty in whips_by_length.items())
    stored_total = totals.get('total_whip_length', 0)
    if abs(calculated_total - stored_total) > 1:  # Allow 1ft rounding tolerance
        issues.append(
            f"WHIP_TOTAL_MISMATCH: Calculated total={calculated_total}ft, "
            f"stored total={stored_total}ft"
        )

    # ── 4. Whip count should equal harness count × 2 (pos + neg) ──
    total_whips = sum(whips_by_length.values())
    if groups is not None:
        # Calculate expected harness count same way as validate_harnesses
        split_tidxs = set(split_tracker_details.keys()) if split_tracker_details else set()
        split_seg_counts = {}
        if tracker_seg_map:
            for tidx in split_tidxs:
                if tidx < len(tracker_seg_map):
                    info = tracker_seg_map[tidx]
                    key = (info['group_idx'], id(info['seg']))
                    split_seg_counts[key] = split_seg_counts.get(key, 0) + 1

        expected_harnesses = 0
        for group_idx, group in enumerate(groups):
            for seg in group.get('segments', []):
                qty = seg.get('quantity', 0)
                if qty <= 0:
                    continue
                harness_config = seg.get('harness_config', str(seg.get('strings_per_tracker', 1)))
                if lv_collection_method == 'String HR':
                    num_harnesses = int(seg.get('strings_per_tracker', 1))
                else:
                    num_harnesses = len(harness_config.split('+'))
                key = (group_idx, id(seg))
                num_splits = split_seg_counts.get(key, 0)
                expected_harnesses += (qty - num_splits) * num_harnesses

        for tidx, details in (split_tracker_details or {}).items():
            for portion in details.get('portions', []):
                expected_harnesses += len(portion.get('harnesses', []))

        expected_whips = expected_harnesses * 2
        if total_whips != expected_whips:
            issues.append(
                f"WHIP_COUNT_MISMATCH: BOM has {total_whips} whip cables, "
                f"expected {expected_whips} ({expected_harnesses} harnesses × 2 pos/neg)"
            )

    # ── 5. Whip count should equal extender count ──
    total_pos_ext = sum(totals.get('extenders_pos_by_length', {}).values())
    total_neg_ext = sum(totals.get('extenders_neg_by_length', {}).values())
    total_ext = total_pos_ext + total_neg_ext
    if total_ext > 0 and total_whips != total_ext:
        issues.append(
            f"WHIP_EXT_COUNT_MISMATCH: {total_whips} whips != "
            f"{total_ext} extenders ({total_pos_ext} pos + {total_neg_ext} neg) — "
            f"every harness should have one whip and one extender per polarity"
        )

    if verbose or issues:
        print("\n[WHIP DIAG] === Whip Validation ===")
        print(f"  Total whip cables: {total_whips}")
        print(f"  Total whip length: {stored_total}ft")
        if whips_by_length:
            print(f"  Breakdown: {dict(whips_by_length)}")
        if groups is not None:
            print(f"  Expected harnesses: {expected_harnesses}")
            print(f"  Expected whips: {expected_whips}")
        if total_ext > 0:
            print(f"  Total extenders: {total_ext}")
        if issues:
            print(f"  [{len(issues)} ISSUES]:")
            for issue in issues:
                print(f"    {issue}")
        else:
            print("  [ALL CHECKS PASSED]")
        print("[WHIP DIAG] === End ===\n")

    return issues

def validate_whip_extender_relationship(totals, qe_widget, verbose=False):
    """Validate that whips are E-W only and far-away tracker extenders are correct.

    Invariants:
      1. No whip should exceed the max E-W span of the project.
      2. For trackers with significant inter-row N-S offset to their device,
         the positive and negative extenders must be asymmetric (differ by ~ns_offset).
      3. The longer extender must be >= abs(ns_offset).

    Args:
        totals: the totals dict from calculate_estimate.
        qe_widget: the QuickEstimate widget instance.
        verbose: if True, print details to console.

    Returns:
        list of issue strings. Empty list = all good.
    """
    issues = []
    ns_offsets = getattr(qe_widget, '_tracker_ns_to_device', {})
    whips_by_length = totals.get('whips_by_length', {})
    groups = getattr(qe_widget, 'groups', [])
    row_spacing = 20.0
    try:
        row_spacing = float(qe_widget.row_spacing_var.get())
    except (ValueError, AttributeError):
        pass

    # ── 1. Max whip should not exceed max E-W span ──
    # Max E-W span = largest group tracker count × row_spacing
    max_ew_span = 0
    for group in groups:
        group_trackers = sum(seg.get('quantity', 0) for seg in group.get('segments', []))
        max_ew_span = max(max_ew_span, group_trackers * row_spacing)

    if max_ew_span > 0 and whips_by_length:
        max_whip = max(length for (length, gauge) in whips_by_length.keys())
        if max_whip > max_ew_span * 1.1:  # 10% tolerance for rounding
            issues.append(
                f"WHIP_TOO_LONG: Max whip {max_whip}ft exceeds max E-W span "
                f"{max_ew_span:.0f}ft — whips should be E-W only"
            )

    # ── 2. Far-away tracker extender checks ──
    split_details = getattr(qe_widget, '_split_tracker_details', {})
    tracker_seg_map = getattr(qe_widget, '_tracker_to_segment', [])
    assignments = getattr(qe_widget, 'last_combiner_assignments', [])

    for (tidx, inv_idx), signed_ns in ns_offsets.items():
        if abs(signed_ns) < 10:
            continue  # Only check trackers with meaningful inter-row offset

        if tidx >= len(tracker_seg_map):
            continue

        seg_info = tracker_seg_map[tidx]
        seg = seg_info['seg']
        device_position = seg_info['device_position']

        # Get the extender pairs for this tracker with offset
        if tidx in split_details:
            for portion in split_details[tidx]['portions']:
                if portion['inv_idx'] != inv_idx:
                    continue
                pairs = qe_widget.calculate_extender_lengths_per_segment(
                    seg, device_position, portion.get('start_pos', 0),
                    target_y_offset=signed_ns,
                    harness_sizes_override=portion['harnesses'])

                for pair_idx, (pos_len, neg_len) in enumerate(pairs):
                    pos_r = qe_widget.round_whip_length(pos_len)
                    neg_r = qe_widget.round_whip_length(neg_len)
                    longer = max(pos_r, neg_r)
                    ns_abs = abs(signed_ns)

                    # Check 2a: asymmetry — pos and neg should differ
                    if abs(pos_r - neg_r) < 10:
                        issues.append(
                            f"EXT_SYMMETRIC: T{tidx+1:02d} portion(inv={inv_idx}) H{pair_idx+1} "
                            f"has pos={pos_r}ft neg={neg_r}ft but N-S offset={ns_abs:.0f}ft — "
                            f"expected asymmetric extenders"
                        )

                    # Check 2b: longer extender must cover the inter-row gap
                    if longer < ns_abs * 0.8:  # 20% tolerance
                        issues.append(
                            f"EXT_TOO_SHORT: T{tidx+1:02d} portion(inv={inv_idx}) H{pair_idx+1} "
                            f"longer ext={longer}ft but N-S offset={ns_abs:.0f}ft — "
                            f"extender should cover inter-row distance"
                        )
        else:
            # Non-split tracker
            pairs = qe_widget.calculate_extender_lengths_per_segment(
                seg, device_position,
                target_y_offset=signed_ns,
                harness_sizes_override=None)

            for pair_idx, (pos_len, neg_len) in enumerate(pairs):
                pos_r = qe_widget.round_whip_length(pos_len)
                neg_r = qe_widget.round_whip_length(neg_len)
                longer = max(pos_r, neg_r)
                ns_abs = abs(signed_ns)

                if abs(pos_r - neg_r) < 10:
                    issues.append(
                        f"EXT_SYMMETRIC: T{tidx+1:02d} (inv={inv_idx}) H{pair_idx+1} "
                        f"has pos={pos_r}ft neg={neg_r}ft but N-S offset={ns_abs:.0f}ft — "
                        f"expected asymmetric extenders"
                    )

                if longer < ns_abs * 0.8:
                    issues.append(
                        f"EXT_TOO_SHORT: T{tidx+1:02d} (inv={inv_idx}) H{pair_idx+1} "
                        f"longer ext={longer}ft but N-S offset={ns_abs:.0f}ft — "
                        f"extender should cover inter-row distance"
                    )

    if verbose or issues:
        print("\n[WHIP/EXT RELATIONSHIP] === Validation ===")
        print(f"  Max E-W span: {max_ew_span:.0f}ft")
        print(f"  Trackers with inter-row offset: {sum(1 for v in ns_offsets.values() if abs(v) >= 10)}")
        if issues:
            print(f"  [{len(issues)} ISSUES]:")
            for issue in issues:
                print(f"    {issue}")
        else:
            print("  [ALL CHECKS PASSED]")
        print("[WHIP/EXT RELATIONSHIP] === End ===\n")

    return issues


# ═══════════════════════════════════════════════════════════════
# Harness Diagnostics
# ═══════════════════════════════════════════════════════════════

def validate_harnesses(totals, groups, split_tracker_details, tracker_seg_map=None,
                       lv_collection_method='Wire Harness', verbose=False):
    """Validate harness BOM counts against tracker/segment data.

    Independently recalculates expected harness counts from segment configs
    and split tracker adjustments, then compares to the BOM totals.

    Args:
        totals: the totals dict containing 'harnesses_by_size'.
        groups: the QE groups list.
        split_tracker_details: dict from _split_tracker_details.
        tracker_seg_map: list of dicts from _tracker_to_segment.
        lv_collection_method: 'Wire Harness' or 'String HR'.
        verbose: if True, print details even when passing.

    Returns:
        list of issue strings. Empty list = all good.
    """
    issues = []
    harnesses_by_size = totals.get('harnesses_by_size', {})
    actual_total = sum(harnesses_by_size.values())

    # ── 1. No zero-size harness entries ──
    for size, qty in harnesses_by_size.items():
        if size <= 0:
            issues.append(f"HARNESS_BAD_SIZE: Harness size={size}, qty={qty}")
        if qty <= 0:
            issues.append(f"HARNESS_BAD_QTY: Harness size={size}, qty={qty}")

    # ── 2. Calculate expected harness count from segments ──
    # Step 2a: bulk count from segments (before split adjustment)
    split_tidxs = set(split_tracker_details.keys()) if split_tracker_details else set()

    # Count how many split trackers belong to each (group_idx, seg) pair
    split_seg_counts = {}
    if tracker_seg_map:
        for tidx in split_tidxs:
            if tidx < len(tracker_seg_map):
                info = tracker_seg_map[tidx]
                key = (info['group_idx'], id(info['seg']))
                split_seg_counts[key] = split_seg_counts.get(key, 0) + 1

    expected_by_size = {}
    for group_idx, group in enumerate(groups):
        for seg in group.get('segments', []):
            qty = seg.get('quantity', 0)
            if qty <= 0:
                continue

            harness_config = seg.get('harness_config', str(seg.get('strings_per_tracker', 1)))
            if lv_collection_method == 'String HR':
                harness_sizes = [1] * int(seg.get('strings_per_tracker', 1))
            else:
                harness_sizes = [int(float(x)) for x in harness_config.split('+')]

            # Subtract split trackers — they get counted separately
            key = (group_idx, id(seg))
            num_splits = split_seg_counts.get(key, 0)
            non_split_qty = qty - num_splits

            for size in harness_sizes:
                expected_by_size[size] = expected_by_size.get(size, 0) + non_split_qty

    # Step 2b: add split tracker portions
    for tidx, details in (split_tracker_details or {}).items():
        for portion in details.get('portions', []):
            for h_size in portion.get('harnesses', []):
                expected_by_size[h_size] = expected_by_size.get(h_size, 0) + 1

    expected_total = sum(expected_by_size.values())

    # ── 3. Compare totals ──
    if actual_total != expected_total:
        issues.append(
            f"HARNESS_COUNT_MISMATCH: BOM has {actual_total} harnesses, "
            f"expected {expected_total} from segments + splits"
        )

    # ── 4. Compare per-size breakdown ──
    all_sizes = set(list(harnesses_by_size.keys()) + list(expected_by_size.keys()))
    for size in sorted(all_sizes):
        actual_qty = harnesses_by_size.get(size, 0)
        expected_qty = expected_by_size.get(size, 0)
        if actual_qty != expected_qty:
            issues.append(
                f"HARNESS_SIZE_MISMATCH: {size}-string harnesses: "
                f"BOM has {actual_qty}, expected {expected_qty}"
            )

    if verbose or issues:
        print("\n[HARNESS DIAG] === Harness Validation ===")
        print(f"  Actual total: {actual_total}")
        print(f"  Expected total: {expected_total}")
        if harnesses_by_size:
            print(f"  Actual by size: {dict(harnesses_by_size)}")
        if expected_by_size:
            print(f"  Expected by size: {dict(expected_by_size)}")
        if issues:
            print(f"  [{len(issues)} ISSUES]:")
            for issue in issues:
                print(f"    {issue}")
        else:
            print("  [ALL CHECKS PASSED]")
        print("[HARNESS DIAG] === End ===\n")

    return issues


# ═══════════════════════════════════════════════════════════════
# Master Diagnostic Runner
# ═══════════════════════════════════════════════════════════════

def run_all_diagnostics(qe_widget, verbose=True):
    """Run all diagnostic checks against a Quick Estimate widget's current state.

    Call this from a "Run Diagnostics" button in the QE UI.

    Args:
        qe_widget: the QuickEstimate widget instance (must have calculated already).
        verbose: if True, print full details to console.

    Returns:
        dict with keys:
            'all_passed': bool
            'sections': list of dicts, each with:
                'name': str (section label)
                'passed': bool
                'issues': list of str
            'project_info': list of (label, value) tuples
    """
    sections = []

    # ── Collect project info ──
    project_info = []

    # Project name
    project = getattr(qe_widget, 'current_project', None)
    if project and hasattr(project, 'metadata') and project.metadata:
        meta = project.metadata
        if meta.name:
            project_info.append(("Project", meta.name))
        if meta.client:
            project_info.append(("Customer", meta.client))
        if meta.location:
            project_info.append(("Location", meta.location))

    # Estimate name
    estimate_id = getattr(qe_widget, 'estimate_id', None)
    if estimate_id and project:
        est_data = project.quick_estimates.get(estimate_id, {})
        if est_data.get('name'):
            project_info.append(("Estimate", est_data['name']))

    # Module
    selected_module = getattr(qe_widget, 'selected_module', None)
    if selected_module:
        project_info.append(("Module", f"{selected_module.manufacturer} {selected_module.model} ({selected_module.wattage}W)"))
        project_info.append(("Module Isc", f"{selected_module.isc} A"))

    # Inverter
    selected_inverter = getattr(qe_widget, 'selected_inverter', None)
    if selected_inverter:
        project_info.append(("Inverter", f"{selected_inverter.manufacturer} {selected_inverter.model} ({selected_inverter.rated_power_kw}kW AC)"))

    # Topology and settings
    topology_var = getattr(qe_widget, 'topology_var', None)
    topology_str = topology_var.get() if topology_var else ''
    if topology_str:
        project_info.append(("Topology", topology_str))

    lv_var = getattr(qe_widget, 'lv_collection_var', None)
    if lv_var:
        project_info.append(("LV Collection", lv_var.get()))

    breaker_var = getattr(qe_widget, 'breaker_size_var', None)
    if breaker_var:
        project_info.append(("Breaker Size", f"{breaker_var.get()}A"))

    spi_var = getattr(qe_widget, 'strings_per_inverter_var', None)
    if spi_var and spi_var.get() != '--':
        project_info.append(("Strings/Inverter", spi_var.get()))

    dc_ac_var = getattr(qe_widget, 'dc_ac_ratio_var', None)
    if dc_ac_var:
        project_info.append(("DC:AC Ratio (target)", dc_ac_var.get()))

    row_spacing_var = getattr(qe_widget, 'row_spacing_var', None)
    if row_spacing_var:
        project_info.append(("Row Spacing", f"{row_spacing_var.get()} ft"))

    # NEC factor
    nec_factor = 1.56
    if project:
        nec_factor = getattr(project, 'nec_safety_factor', 1.56)
    project_info.append(("NEC Safety Factor", f"{nec_factor}"))

    # Segment summary
    groups = getattr(qe_widget, 'groups', [])
    total_trackers = 0
    total_strings = 0
    for group in groups:
        for seg in group.get('segments', []):
            qty = seg.get('quantity', 0)
            spt = seg.get('strings_per_tracker', 0)
            total_trackers += qty
            total_strings += int(qty * spt)
    project_info.append(("Total Trackers", str(total_trackers)))
    project_info.append(("Total Strings", str(total_strings)))
    project_info.append(("Groups", str(len(groups))))

    # Segment details
    for grp_idx, group in enumerate(groups):
        grp_name = group.get('name', f'Group {grp_idx + 1}')
        for seg_idx, seg in enumerate(group.get('segments', [])):
            qty = seg.get('quantity', 0)
            spt = seg.get('strings_per_tracker', 0)
            config = seg.get('harness_config', str(int(spt)))
            project_info.append((f"  {grp_name} Seg{seg_idx+1}", f"{qty}x {spt}S, harness={config}"))

    if verbose:
        print("\n" + "=" * 60)
        print("PROJECT INFO")
        print("=" * 60)
        max_label = max((len(label) for label, _ in project_info), default=0)
        for label, value in project_info:
            print(f"  {label:<{max_label + 2}} {value}")
        print("=" * 60)

    # ── 1. Combiner Assignments ──
    cb_issues = []
    topology = getattr(qe_widget, 'topology_var', None)
    topology_str = topology.get() if topology else ''
    assignments = getattr(qe_widget, 'last_combiner_assignments', None)
    if topology_str == 'Distributed String':
        cb_issues = []  # No combiner assignments expected for this topology
    elif assignments:
        cb_issues = validate_assignments(assignments, verbose=verbose)
    else:
        cb_issues = ['NO_DATA: No combiner assignments found. Run Calculate first.']
    sections.append({
        'name': 'Combiner Assignments',
        'passed': len(cb_issues) == 0,
        'issues': cb_issues,
    })

    # ── 2. Split Tracker Details ──
    split_details = getattr(qe_widget, '_split_tracker_details', {})
    tracker_seg_map = getattr(qe_widget, '_tracker_to_segment', [])
    split_issues = validate_split_details(split_details, tracker_seg_map, verbose=verbose)
    sections.append({
        'name': 'Split Trackers',
        'passed': len(split_issues) == 0,
        'issues': split_issues,
    })

    # ── 3. Extenders ──
    totals = getattr(qe_widget, 'last_totals', None)
    groups = getattr(qe_widget, 'groups', [])
    enabled_templates = getattr(qe_widget, 'enabled_templates', None)
    ext_issues = []
    lv_method = getattr(qe_widget, 'lv_collection_var', None)
    lv_collection_method = lv_method.get() if lv_method else 'Wire Harness'
    if totals:
        ns_offsets = getattr(qe_widget, '_tracker_ns_to_device', {})
        max_ns_offset = max((abs(v) for v in ns_offsets.values()), default=0)
        ext_issues = validate_extenders(
            totals, groups, tracker_seg_map, split_details,
            enabled_templates=enabled_templates,
            lv_collection_method=lv_collection_method,
            max_ns_offset=max_ns_offset,
            verbose=verbose,
        )
    else:
        ext_issues = ['NO_DATA: No totals found. Run Calculate first.']
    sections.append({
        'name': 'Extenders',
        'passed': len(ext_issues) == 0,
        'issues': ext_issues,
    })

    # ── 4. Whips ──
    whip_issues = []
    if totals:
        whip_issues = validate_whips(
            totals, groups=groups, split_tracker_details=split_details,
            tracker_seg_map=tracker_seg_map,
            lv_collection_method=lv_collection_method,
            verbose=verbose,
        )
    else:
        whip_issues = ['NO_DATA: No totals found. Run Calculate first.']
    sections.append({
        'name': 'Whips',
        'passed': len(whip_issues) == 0,
        'issues': whip_issues,
    })

    # ── 4b. Whip/Extender Relationship ──
    rel_issues = []
    if totals:
        rel_issues = validate_whip_extender_relationship(totals, qe_widget, verbose=verbose)
    else:
        rel_issues = ['NO_DATA: No totals found. Run Calculate first.']
    sections.append({
        'name': 'Whip/Extender Relationship',
        'passed': len(rel_issues) == 0,
        'issues': rel_issues,
    })

    # ── 5. Harnesses ──
    harness_issues = []
    if totals:
        harness_issues = validate_harnesses(
            totals, groups, split_details,
            tracker_seg_map=tracker_seg_map,
            lv_collection_method=lv_collection_method,
            verbose=verbose,
        )
    else:
        harness_issues = ['NO_DATA: No totals found. Run Calculate first.']
    sections.append({
        'name': 'Harnesses',
        'passed': len(harness_issues) == 0,
        'issues': harness_issues,
    })

    all_passed = all(s['passed'] for s in sections)

    # ── Summary ──
    if verbose:
        print("\n" + "=" * 60)
        print("DIAGNOSTIC SUMMARY")
        print("=" * 60)
        for section in sections:
            status = "PASS" if section['passed'] else "FAIL"
            issue_count = len(section['issues'])
            print(f"  [{status}] {section['name']}"
                  + (f" ({issue_count} issues)" if issue_count > 0 else ""))
        print("=" * 60)
        print(f"  Overall: {'ALL PASSED' if all_passed else 'ISSUES FOUND'}")
        print("=" * 60 + "\n")

    return {
        'all_passed': all_passed,
        'sections': sections,
        'project_info': project_info,
    }


def format_diagnostic_report(result):
    """Format diagnostic results as a human-readable string for display in a dialog.

    Args:
        result: dict returned by run_all_diagnostics.

    Returns:
        str suitable for display in a messagebox or text widget.
    """
    lines = []
    lines.append("=" * 50)
    lines.append("  QUICK ESTIMATE DIAGNOSTIC REPORT")
    lines.append("=" * 50)
    lines.append("")

    # Project info section
    project_info = result.get('project_info', [])
    if project_info:
        lines.append("PROJECT INFO")
        lines.append("-" * 50)
        max_label = max((len(label) for label, _ in project_info), default=0)
        for label, value in project_info:
            lines.append(f"  {label:<{max_label + 2}} {value}")
        lines.append("")

    for section in result['sections']:
        status = "PASS" if section['passed'] else "FAIL"
        lines.append(f"[{status}] {section['name']}")
        if section['issues']:
            for issue in section['issues']:
                lines.append(f"      {issue}")
        lines.append("")

    lines.append("-" * 50)
    if result['all_passed']:
        lines.append("  Result: ALL CHECKS PASSED")
    else:
        total_issues = sum(len(s['issues']) for s in result['sections'])
        lines.append(f"  Result: {total_issues} ISSUE(S) FOUND")
    lines.append("-" * 50)

    return "\n".join(lines)


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