"""Shared device-position and tracker-world-coordinate geometry.

Used by both src/ui/site_preview.py (the reference implementation this module
was extracted from) and src/ui/quick_estimate.py, so a combiner-box nudge and
the whip-length math it drives agree on where the device actually is.

Pure functions only — no Tk, no imports of quick_estimate or site_preview.
Callers pass in whatever group/template/step state they hold.
"""

import math


def tracker_top_y(group, tracker, local_x_idx=None):
    """World-Y of a tracker's north (top) edge.

    `group` must provide 'y', 'length_ft', 'row_spacing_ft', 'driveline_tan',
    and 'tracker_alignment' ('top' | 'bottom' | 'motor'), plus 'motor_y_ft'
    (the group's reference motor offset) when alignment is 'motor'.
    `tracker` must provide 'length_ft' and — for 'motor' alignment —
    'has_motor' / 'motor_y_ft'. `local_x_idx` is used only if the tracker
    dict itself has no 'local_x_idx' entry.
    """
    gy = group['y']
    group_length = group.get('length_ft', 0)
    pitch = group.get('row_spacing_ft', 0)
    driveline_tan = group.get('driveline_tan', 0.0)
    alignment = group.get('tracker_alignment', 'motor')
    motor_ref = group.get('motor_y_ft', None)

    t_length = tracker.get('length_ft', group_length)
    t_lx = tracker.get('local_x_idx', local_x_idx if local_x_idx is not None else 0)
    t_ang = t_lx * pitch * driveline_tan

    if alignment == 'top':
        return gy + t_ang
    elif alignment == 'bottom':
        return gy + (group_length - t_length) + t_ang
    elif tracker.get('has_motor', False) and motor_ref is not None:
        return gy + (motor_ref - tracker.get('motor_y_ft', 0)) + t_ang
    else:
        return gy + (group_length - t_length) / 2 + t_ang


def tracker_motor_row_y(group, tracker, local_x_idx=None):
    """World-Y of a tracker's motor row (the device 'middle' anchor, before
    centering the device box on it)."""
    ty = tracker_top_y(group, tracker, local_x_idx)
    t_length = tracker.get('length_ft', group.get('length_ft', 0))
    t_motor = tracker.get('motor_y_ft', t_length / 2)
    motor_gap = tracker.get('motor_gap_ft', 0.0)
    return ty + t_motor + motor_gap / 2


def tracker_ns_anchors(group, tracker, offset_ft, device_height_ft, local_x_idx=None):
    """The three device-Y snap options for one tracker, as (y_ft, zone_type) tuples.

    Rounded to 4 decimals — these feed a snap-to-nearest-anchor step resolver,
    not a rendered position, so float noise must not create phantom anchors.
    """
    ty = tracker_top_y(group, tracker, local_x_idx)
    t_length = tracker.get('length_ft', group.get('length_ft', 0))
    motor_row_y = tracker_motor_row_y(group, tracker, local_x_idx)
    return [
        (round(ty - offset_ft - device_height_ft, 4), 'north'),
        (round(motor_row_y - device_height_ft / 2, 4), 'middle'),
        (round(ty + t_length + offset_ft, 4), 'south'),
    ]


def device_position_for_inverter(harness_map, tracker_to_group, group_layout, groups,
                                  device_width_ft, device_height_ft, offset_ft, max_width_ft):
    """Resolve one inverter's device position from its harness_map.

    `tracker_to_group[tracker_idx] = (group_idx, local_idx)`.
    `group_layout[group_idx]` is shaped like site_preview.py's build_layout_data
    output: 'x', 'y', 'length_ft', 'row_spacing_ft', 'driveline_tan',
    'tracker_alignment', 'motor_y_ft', 'visual_min_y', 'visual_max_y', and
    'trackers' (a list of tracker dicts with 'local_x_idx' / 'length_ft' /
    'motor_y_ft' / 'has_motor' / 'motor_gap_ft' / 'strings_per_tracker').
    `groups[group_idx].get('device_position')` selects 'north' / 'south' / 'middle'.

    Returns None if the inverter has no usable harness_map / group assignment.
    Otherwise returns a dict with:
        'x', 'y'               -- device position (pre middle-x-bias; see below)
        'device_position'      -- resolved zone ('north' | 'south' | 'middle')
        'primary_group_idx'    -- the group this device's X/Y were anchored to
        'ns_anchors'           -- sorted (y_ft, zone_type) snap candidates,
                                   built from every tracker this inverter touches
        'middle_bias_context'  -- None unless device_position == 'middle'; else
                                   the ingredients a caller needs to run its own
                                   row-gap bias correction on 'x' (this module
                                   deliberately does not apply one, since the two
                                   callers use different bias logic today)

    The primary group is chosen by total strings_taken per group (not harness_map
    entry count), ties breaking toward the lower group index — a split tracker
    contributing several entries to the same group can't outvote a group with
    fewer, larger entries, and the result is stable regardless of dict iteration
    order.
    """
    inv_tracker_indices = [entry['tracker_idx'] for entry in harness_map]

    group_weights = {}
    for entry in harness_map:
        tidx = entry['tracker_idx']
        if tidx in tracker_to_group:
            grp_idx = tracker_to_group[tidx][0]
            group_weights[grp_idx] = group_weights.get(grp_idx, 0) + entry.get('strings_taken', 1)

    if not group_weights:
        return None

    max_weight = max(group_weights.values())
    primary_grp_idx = min(g for g, w in group_weights.items() if w == max_weight)

    group_data = group_layout[primary_grp_idx]
    group_source = groups[primary_grp_idx] if primary_grp_idx < len(groups) else {}
    device_position = group_source.get('device_position', 'middle')

    gx = group_data['x']
    gy = group_data['y']

    local_indices = [
        tracker_to_group[tidx][1] for tidx in inv_tracker_indices
        if tidx in tracker_to_group and tracker_to_group[tidx][0] == primary_grp_idx
    ]

    group_trackers_list = group_data.get('trackers', [])
    pitch = group_data.get('row_spacing_ft', 0)

    local_x_indices = [
        group_trackers_list[li].get('local_x_idx', li)
        for li in local_indices if li < len(group_trackers_list)
    ]

    if local_x_indices:
        center_local_x = (min(local_x_indices) + max(local_x_indices)) / 2.0
        device_x = gx + center_local_x * pitch + (max_width_ft - device_width_ft) / 2
    else:
        center_local_x = 0
        device_x = gx

    angle_y_offset = center_local_x * pitch * group_data.get('driveline_tan', 0.0)

    middle_bias_context = None

    if device_position in ('north', 'south'):
        # Step 1: find the most extreme group by visual extent.
        anchor_grp_idx = primary_grp_idx
        anchor_val = None
        for grp_idx_s, grp_data_s in enumerate(group_layout):
            has_any = any(
                tidx in tracker_to_group and tracker_to_group[tidx][0] == grp_idx_s
                for tidx in inv_tracker_indices
            )
            if not has_any:
                continue
            grp_gy_s = grp_data_s['y']
            if device_position == 'south':
                vis_max = grp_data_s.get('visual_max_y', grp_data_s.get('length_ft', 0))
                val = grp_gy_s + vis_max
                if anchor_val is None or val > anchor_val:
                    anchor_val = val
                    anchor_grp_idx = grp_idx_s
            else:  # 'north'
                vis_min = grp_data_s.get('visual_min_y', 0)
                val = grp_gy_s + vis_min
                if anchor_val is None or val < anchor_val:
                    anchor_val = val
                    anchor_grp_idx = grp_idx_s

        # Step 2: within the anchor group, use the closest-to-device-X tracker.
        anchor_grp_data = group_layout[anchor_grp_idx]
        anchor_gy = anchor_grp_data['y']
        anchor_trackers = anchor_grp_data.get('trackers', [])

        anchor_local = [
            tracker_to_group[tidx][1]
            for tidx in inv_tracker_indices
            if tidx in tracker_to_group and tracker_to_group[tidx][0] == anchor_grp_idx
        ]

        closest_local_idx = None
        closest_dist = float('inf')
        for li in anchor_local:
            if li >= len(anchor_trackers):
                continue
            t_lx = anchor_trackers[li].get('local_x_idx', li)
            dist = abs(t_lx - center_local_x)
            if dist < closest_dist:
                closest_dist = dist
                closest_local_idx = li

        if closest_local_idx is not None:
            t = anchor_trackers[closest_local_idx]
            t_length = t.get('length_ft', anchor_grp_data.get('length_ft', 0))
            ty = tracker_top_y(anchor_grp_data, t, t.get('local_x_idx', closest_local_idx))

            if device_position == 'north':
                device_y = ty - offset_ft - device_height_ft
            else:
                device_y = ty + t_length + offset_ft
        else:
            # Fallback to anchor group bounds
            if device_position == 'north':
                vis_min = anchor_grp_data.get('visual_min_y', 0)
                device_y = anchor_gy + vis_min - offset_ft - device_height_ft
            else:
                vis_max = anchor_grp_data.get('visual_max_y', anchor_grp_data.get('length_ft', 0))
                device_y = anchor_gy + vis_max + offset_ft
    else:  # 'middle' or fallback
        group_length = group_data.get('length_ft', 0)
        group_motor_y_ref = group_data.get('motor_y_ft', None)

        m_closest_idx = None
        m_closest_dist = float('inf')
        for li in local_indices:
            if li < len(group_trackers_list):
                t_lx = group_trackers_list[li].get('local_x_idx', li)
                dist = abs(t_lx - center_local_x)
                if dist < m_closest_dist:
                    m_closest_dist = dist
                    m_closest_idx = li

        if m_closest_idx is not None:
            t = group_trackers_list[m_closest_idx]
            device_y = tracker_motor_row_y(
                group_data, t, t.get('local_x_idx', m_closest_idx)
            ) - device_height_ft / 2
        else:
            fallback_motor = group_motor_y_ref if group_motor_y_ref is not None else group_length / 2
            fallback_gap = group_trackers_list[0].get('motor_gap_ft', 0.0) if group_trackers_list else 0.0
            device_y = gy + fallback_motor + fallback_gap / 2 - device_height_ft / 2 + angle_y_offset

        if local_x_indices:
            spt_map = {
                group_trackers_list[li].get('local_x_idx', li):
                group_trackers_list[li].get('strings_per_tracker', 1)
                for li in local_indices if li < len(group_trackers_list)
            }
            middle_bias_context = {
                'center_local_x': center_local_x,
                'local_x_indices': local_x_indices,
                'spt_map': spt_map,
                'pitch': pitch,
                'group_x': gx,
                'group_num_trackers': len(group_trackers_list),
            }

    # N-S snap anchors from every tracker this inverter touches (not just the
    # primary group) — a device spanning groups can snap to any of them.
    ns_typed_set = set()
    for tidx in inv_tracker_indices:
        if tidx not in tracker_to_group:
            continue
        g_idx, l_idx = tracker_to_group[tidx]
        gd = group_layout[g_idx]
        g_trackers = gd.get('trackers', [])
        if l_idx >= len(g_trackers):
            continue
        t = g_trackers[l_idx]
        for anchor in tracker_ns_anchors(gd, t, offset_ft, device_height_ft, t.get('local_x_idx', l_idx)):
            ns_typed_set.add(anchor)
    ns_anchors = sorted(ns_typed_set, key=lambda a: a[0])

    return {
        'x': device_x,
        'y': device_y,
        'device_position': device_position,
        'primary_group_idx': primary_grp_idx,
        'ns_anchors': ns_anchors,
        'middle_bias_context': middle_bias_context,
    }


def tracker_dims_ft(template):
    """(width_ft, length_ft) for a tracker from its template dict.

    Ported from site_preview.py's _get_preview_tracker_dims_ft (the reference
    implementation) so quick_estimate.py — which has its own, slightly
    different dims formula for its unrelated rendering needs — can build the
    same tracker geometry site_preview.py does for position/whip purposes.
    Returns None if `template` is falsy.
    """
    if not template:
        return None

    module_spec = template.get('module_spec', {})
    module_length_mm = module_spec.get('length_mm', 2000)
    module_width_mm = module_spec.get('width_mm', 1000)
    orientation = template.get('module_orientation', 'Portrait')
    modules_per_string = template.get('modules_per_string', 28)
    strings_per_tracker = template.get('strings_per_tracker', 2)
    modules_high = template.get('modules_high', 1)
    module_spacing_m = template.get('module_spacing_m', 0.02)
    has_motor = template.get('has_motor', True)
    motor_gap_m = template.get('motor_gap_m', 1.0) if has_motor else 0

    if orientation == 'Portrait':
        mod_along_m = module_width_mm / 1000
        mod_across_m = module_length_mm / 1000
    else:
        mod_along_m = module_length_mm / 1000
        mod_across_m = module_width_mm / 1000

    full_spt = int(strings_per_tracker)
    partial_mods = round((strings_per_tracker - full_spt) * modules_per_string) if strings_per_tracker != full_spt else 0
    modules_in_row = full_spt * modules_per_string + partial_mods
    tracker_length_m = (modules_in_row * mod_along_m +
                         (modules_in_row - 1) * module_spacing_m +
                         motor_gap_m)
    tracker_width_m = mod_across_m * modules_high

    m_to_ft = 3.28084
    return (tracker_width_m * m_to_ft, tracker_length_m * m_to_ft)


def motor_position_in_tracker(template):
    """(motor_y_offset_ft, motor_gap_ft, has_motor) for a tracker template.

    Ported from site_preview.py's get_motor_position_in_tracker (the reference
    implementation) — quick_estimate.py never had an equivalent of its own,
    which is exactly why its whip math used to be motor-blind.
    """
    if not template:
        return 0, 0, False

    has_motor = template.get('has_motor', True)
    if not has_motor:
        return 0, 0, False

    module_spec = template.get('module_spec', {})
    module_length_mm = module_spec.get('length_mm', 2000)
    module_width_mm = module_spec.get('width_mm', 1000)
    orientation = template.get('module_orientation', 'Portrait')
    modules_per_string = template.get('modules_per_string', 28)
    module_spacing_m = template.get('module_spacing_m', 0.02)
    motor_gap_m = template.get('motor_gap_m', 1.0)
    motor_placement = template.get('motor_placement_type', 'between_strings')
    motor_position_after_string = template.get('motor_position_after_string', None)
    motor_string_index_raw = template.get('motor_string_index', None)
    motor_split_north = template.get('motor_split_north', modules_per_string // 2)

    if orientation == 'Portrait':
        mod_along_m = module_width_mm / 1000
    else:
        mod_along_m = module_length_mm / 1000

    m_to_ft = 3.28084

    # Partial string on north pushes motor further south
    partial_north_m = 0
    spt_val = template.get('strings_per_tracker', 1)
    if spt_val != int(spt_val) and template.get('partial_string_side', 'north') == 'north':
        partial_north_mods = round((spt_val - int(spt_val)) * modules_per_string)
        partial_north_m = partial_north_mods * (mod_along_m + module_spacing_m)

    if motor_placement == 'between_strings':
        pos_after = motor_position_after_string if motor_position_after_string is not None else (motor_string_index_raw if motor_string_index_raw is not None else 1)
        modules_north = pos_after * modules_per_string
        if modules_north > 0:
            motor_y_m = partial_north_m + (modules_north * mod_along_m +
                        (modules_north - 1) * module_spacing_m +
                        module_spacing_m)
        else:
            motor_y_m = partial_north_m
    elif motor_placement == 'middle_of_string':
        string_idx = motor_string_index_raw if motor_string_index_raw is not None else 1
        modules_before_split = (string_idx - 1) * modules_per_string + motor_split_north
        motor_y_m = partial_north_m + (modules_before_split * mod_along_m +
                    (modules_before_split - 1) * module_spacing_m +
                    module_spacing_m)
    else:
        # Fallback: center on the tracker
        dims = tracker_dims_ft(template)
        if dims:
            return dims[1] / 2, motor_gap_m * m_to_ft, True
        return 0, 0, False

    return motor_y_m * m_to_ft, motor_gap_m * m_to_ft, True


def build_group_layout(groups, templates, default_row_spacing_ft=20.0):
    """Build a group_layout + tracker_to_group index from raw project groups
    and the enabled_templates registry — everything device_position_for_inverter
    and the tracker_* helpers need, without requiring a pre-built layout.

    Mirrors site_preview.py's build_layout_data (template dims, motor position,
    per-tracker local_x_idx, driveline shear, visual N-S bounds), minus fields
    that exist only for rendering (colors, string assignments, group width).

    Group position resolves the way quick_estimate.py's whip/feeder math always
    has: saved position_x/position_y when present, else an auto X cursor and
    Y=0 — quick_estimate.py has no richer auto-layout of its own to preserve.

    Returns (group_layout, tracker_to_group, max_tracker_width_ft).
    tracker_to_group maps a global tracker index to (group_idx, local_idx).
    max_tracker_width_ft is the largest tracker width across every group —
    apply_middle_x_bias's boundary-clamp equivalent of site_preview.py's
    self.max_tracker_width_ft.
    """
    group_layout = []
    tracker_to_group = {}
    global_idx = 0
    auto_x_cursor = 0.0
    max_tracker_width_ft = 0.0

    for grp_idx, group_data in enumerate(groups):
        grp_pitch = group_data.get('row_spacing_ft', default_row_spacing_ft)

        saved_x = group_data.get('position_x')
        gx = saved_x if saved_x is not None else auto_x_cursor
        saved_y = group_data.get('position_y')
        gy = saved_y if saved_y is not None else 0.0

        group_trackers = []
        group_motor_y = None
        local_x_counter = 0

        for seg in group_data.get('segments', []):
            ref = seg.get('template_ref')
            template = templates.get(ref) if ref else None
            dims = tracker_dims_ft(template)

            for _ in range(seg.get('quantity', 0)):
                tracker = {}
                if dims:
                    tracker['width_ft'] = dims[0]
                    tracker['length_ft'] = dims[1]
                else:
                    tracker['width_ft'] = 6.0
                    tracker['length_ft'] = 180.0
                max_tracker_width_ft = max(max_tracker_width_ft, tracker['width_ft'])

                motor_y, motor_gap, has_motor = motor_position_in_tracker(template)
                tracker['motor_y_ft'] = motor_y
                tracker['motor_gap_ft'] = motor_gap
                tracker['has_motor'] = has_motor
                tracker['strings_per_tracker'] = seg.get('strings_per_tracker', 1)
                tracker['local_x_idx'] = local_x_counter

                if group_motor_y is None and has_motor:
                    group_motor_y = motor_y

                group_trackers.append(tracker)
                tracker_to_group[global_idx] = (grp_idx, local_x_counter)

                local_x_counter += 1
                global_idx += 1

        group_length = max((t['length_ft'] for t in group_trackers), default=0)

        driveline_angle_deg = group_data.get('driveline_angle', 0.0)
        driveline_tan = math.tan(math.radians(driveline_angle_deg)) if driveline_angle_deg != 0 else 0.0
        tracker_alignment = group_data.get('tracker_alignment', 'motor')
        ref_motor = group_motor_y or 0

        visual_min_y_offset = 0.0
        visual_max_y_offset = 0.0
        for t in group_trackers:
            t_length_val = t.get('length_ft', group_length)
            t_local_x = t.get('local_x_idx', 0)
            if tracker_alignment == 'top':
                y_offset = 0.0
            elif tracker_alignment == 'bottom':
                y_offset = group_length - t_length_val
            else:  # 'motor'
                y_offset = (ref_motor or 0) - t.get('motor_y_ft', 0)
            angle_y = t_local_x * grp_pitch * driveline_tan
            visual_min_y_offset = min(visual_min_y_offset, y_offset + angle_y)
            visual_max_y_offset = max(visual_max_y_offset, y_offset + angle_y + t_length_val)

        group_layout.append({
            'x': gx,
            'y': gy,
            'length_ft': group_length,
            'motor_y_ft': group_motor_y or 0,
            'row_spacing_ft': grp_pitch,
            'driveline_tan': driveline_tan,
            'tracker_alignment': tracker_alignment,
            'visual_min_y': visual_min_y_offset,
            'visual_max_y': visual_max_y_offset,
            'trackers': group_trackers,
        })

        group_tracker_count = sum(seg.get('quantity', 0) for seg in group_data.get('segments', []))
        auto_x_cursor += group_tracker_count * grp_pitch + grp_pitch * 2

    return group_layout, tracker_to_group, max_tracker_width_ft


def apply_middle_x_bias(device_x, device_y, center_local, local_indices,
                         strings_per_tracker_map, pitch, group_x, group_num_trackers,
                         max_tracker_width_ft, pads):
    """Shift device_x into the row-spacing gap for 'middle' placement.

    Ported from site_preview.py's _apply_middle_x_bias (the reference
    implementation) so quick_estimate.py's whip/feeder math picks the same
    bias direction — its own simpler "always bias east" rule could disagree
    with this on a strings tie broken toward the nearest pad, throwing off
    E-W distance by a full row pitch for exactly the devices this tie-break
    matters for.

    Biases east if more strings are east of center_local, west if more are
    west. Tie-breaks toward the nearest pad (defaults east when no pads
    exist). Falls back to the opposite direction if the bias would leave the
    group's bounding rect.
    """
    # If center_local is already at a gap position (half-integer like 0.5, 1.5 ...),
    # the device_x is already centered in the row gap — no bias needed.
    if abs(center_local % 1 - 0.5) < 0.01:
        return device_x

    half_pitch = pitch / 2.0
    east_strings = sum(
        strings_per_tracker_map.get(i, 1) for i in local_indices if i > center_local
    )
    west_strings = sum(
        strings_per_tracker_map.get(i, 1) for i in local_indices if i < center_local
    )

    if east_strings > west_strings:
        bias = half_pitch
    elif west_strings > east_strings:
        bias = -half_pitch
    else:
        bias = half_pitch  # default east
        if pads:
            nearest_pad = min(
                pads,
                key=lambda p: (device_x - (p['x'] + p.get('width_ft', 10.0) / 2)) ** 2
                              + (device_y - (p['y'] + p.get('height_ft', 8.0) / 2)) ** 2
            )
            pad_cx = nearest_pad['x'] + nearest_pad.get('width_ft', 10.0) / 2
            bias = half_pitch if pad_cx >= device_x else -half_pitch

    x_min = group_x
    x_max = group_x + max(group_num_trackers - 1, 0) * pitch + max_tracker_width_ft

    new_x = device_x + bias
    if new_x < x_min or new_x > x_max:
        # Bias pushed outside group bounds — try opposite side
        new_x = max(x_min, min(x_max, device_x - bias))

    return new_x


def resolve_ns_step(steps, dev_idx, base_y, anchors):
    """Resolve a stored N-S step count against freshly built anchors.

    Same semantics as site_preview.py's _resolve_ns_step: `steps` is the
    caller's own {device_idx: int} nudge-step dict, mutated in place (clamped)
    if the stored step overruns the anchor list.

    Returns (resolved_y, zone_type), or (base_y, None) when there's no step
    to apply or no anchors to snap to — the caller should leave position/zone
    alone in that case.
    """
    step = steps.get(dev_idx, 0)
    if step == 0 or not anchors:
        return base_y, None

    nearest_idx = min(range(len(anchors)), key=lambda i: abs(anchors[i][0] - base_y))
    target_idx = nearest_idx + step
    clamped_idx = max(0, min(len(anchors) - 1, target_idx))
    if clamped_idx != target_idx:
        steps[dev_idx] = clamped_idx - nearest_idx

    target_y, target_type = anchors[clamped_idx]
    return target_y, target_type


def physical_anchor_y(resolved_y, resolved_zone_type, device_height_ft):
    """Convert a device_position_for_inverter render-corner Y into the physical
    connection-point Y for the resolved zone.

    device_position_for_inverter's 'y' (and each ns_anchors entry) is the
    top-left corner of a device_height_ft-tall rendered box: the motor row
    minus half that height for 'middle', or the offset_ft standoff edge minus
    the full height for 'north' ('south' has no height term — its anchor
    already sits on the near edge). Callers with no rendered box of their own
    call device_position_for_inverter with the SAME device_height_ft used
    elsewhere (so nudge-step indices resolve to the same anchor everywhere)
    and then use this to recover the physical point for their own math.
    """
    if resolved_zone_type == 'north':
        return resolved_y + device_height_ft
    elif resolved_zone_type == 'middle':
        return resolved_y + device_height_ft / 2.0
    else:  # 'south', or no zone resolved
        return resolved_y
