"""
Corridor routing helpers.

Corridors are horizontal polylines (in world feet) that DC feeder / AC homerun
cables travel along before turning row-direction to reach their assigned pad.
The three-leg path is:
    Leg 1 — row-direction from device to corridor entry point
    Leg 2 — along the corridor from entry to exit
    Leg 3 — row-direction from corridor exit point to pad
"""

import math


def intersect_horizontal(y, polyline):
    """
    Return all points where polyline crosses the horizontal line at world-y == y.

    Returns a list of (x, segment_idx, t) where t ∈ [0, 1] is the parameter
    along segment segment_idx.  Exactly-on-a-vertex intersections are returned
    once for the segment that starts at that vertex (or the last segment if it
    is the last vertex).
    """
    results = []
    for i in range(len(polyline) - 1):
        x1, y1 = polyline[i]
        x2, y2 = polyline[i + 1]
        dy = y2 - y1
        if dy == 0:
            continue  # horizontal segment, skip
        t = (y - y1) / dy
        if 0.0 <= t <= 1.0:
            x = x1 + t * (x2 - x1)
            results.append((x, i, t))
    return results


def nearest_endpoint(point, polyline):
    """
    Return (segment_idx, t) for the polyline endpoint closest to point in
    y-coordinate.  t is 0.0 for the start of segment_idx, 1.0 for the end.
    Falls back gracefully when no horizontal intersection exists.
    """
    px, py = point
    best_dist = float('inf')
    best = (0, 0.0)
    for i, (vx, vy) in enumerate(polyline):
        dist = abs(vy - py)
        if dist < best_dist:
            best_dist = dist
            seg_idx = max(0, i - 1) if i == len(polyline) - 1 else i
            t = 1.0 if i == len(polyline) - 1 and i > 0 else 0.0
            best = (seg_idx, t)
    return best


def _arc_param_to_length(polyline, seg_idx, t):
    """Return total arc length from polyline start to (seg_idx, t)."""
    length = 0.0
    for i in range(seg_idx):
        x1, y1 = polyline[i]
        x2, y2 = polyline[i + 1]
        length += math.hypot(x2 - x1, y2 - y1)
    x1, y1 = polyline[seg_idx]
    x2, y2 = polyline[seg_idx + 1]
    length += t * math.hypot(x2 - x1, y2 - y1)
    return length


def polyline_arc_length(polyline, p1, p2):
    """
    Return the distance along the polyline between two (segment_idx, t) positions.
    The order of p1 and p2 does not matter — the absolute value is returned.
    """
    l1 = _arc_param_to_length(polyline, p1[0], p1[1])
    l2 = _arc_param_to_length(polyline, p2[0], p2[1])
    return abs(l2 - l1)


def pick_entry(device_xy, polyline):
    """
    Find the best entry point on the polyline for a device at device_xy.

    Strategy:
    1. Find all horizontal intersections at device_y.
    2. Pick the one with the smallest |x - device_x| (fewest horizontal feet).
    3. If no intersection exists, fall back to nearest_endpoint.

    Returns (entry_x, (segment_idx, t)).
    """
    device_x, device_y = device_xy
    hits = intersect_horizontal(device_y, polyline)
    if hits:
        best_x, best_seg, best_t = min(hits, key=lambda h: abs(h[0] - device_x))
        return best_x, (best_seg, best_t)
    # Fallback
    seg_idx, t = nearest_endpoint(device_xy, polyline)
    seg_x1, seg_y1 = polyline[seg_idx]
    seg_x2, seg_y2 = polyline[seg_idx + 1] if seg_idx + 1 < len(polyline) else (seg_x1, seg_y1)
    entry_x = seg_x1 + t * (seg_x2 - seg_x1)
    entry_y = seg_y1 + t * (seg_y2 - seg_y1)
    return entry_x, (seg_idx, t)


def three_leg_distance(device_xy, pad_xy, polyline):
    """
    Compute the total routed distance for a device assigned to a corridor.

    Returns (total_ft, path_geom) where path_geom is a list of (x, y) world
    coords representing the three-leg polyline:
        [device_xy, corridor_entry, ... corridor segments ..., corridor_exit, pad_xy]
    """
    device_x, device_y = device_xy
    pad_x, pad_y = pad_xy

    entry_x, entry_pos = pick_entry(device_xy, polyline)
    exit_x, exit_pos = pick_entry(pad_xy, polyline)

    leg1 = abs(device_x - entry_x)
    leg2 = polyline_arc_length(polyline, entry_pos, exit_pos)
    leg3 = abs(pad_x - exit_x)
    total = leg1 + leg2 + leg3

    # Build path geometry
    entry_y = device_y  # entry is at the device's y (horizontal intersection)
    exit_y = pad_y      # exit is at the pad's y (horizontal intersection)

    # Collect corridor vertices between entry and exit
    entry_len = _arc_param_to_length(polyline, entry_pos[0], entry_pos[1])
    exit_len = _arc_param_to_length(polyline, exit_pos[0], exit_pos[1])
    if entry_len <= exit_len:
        from_len, to_len = entry_len, exit_len
        forward = True
    else:
        from_len, to_len = exit_len, entry_len
        forward = False

    # Vertex indices between from_len and to_len (exclusive of end-segment fractional pts)
    middle_pts = []
    acc = 0.0
    for i in range(len(polyline) - 1):
        x1, y1 = polyline[i]
        x2, y2 = polyline[i + 1]
        seg_len = math.hypot(x2 - x1, y2 - y1)
        if acc + seg_len > from_len and acc < to_len:
            if acc >= from_len:
                middle_pts.append((x1, y1))
        acc += seg_len
    if not forward:
        middle_pts.reverse()

    entry_pt = (entry_x, entry_y)
    exit_pt = (exit_x, exit_y)

    if forward:
        corridor_pts = [entry_pt] + middle_pts + [exit_pt]
    else:
        corridor_pts = [exit_pt] + middle_pts + [entry_pt]

    path_geom = [device_xy] + corridor_pts + [pad_xy]
    return total, path_geom
