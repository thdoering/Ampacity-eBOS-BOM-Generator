"""
Site PDF Generator — Renders the Quick Estimate string allocation site preview
to an 11x17 landscape PDF with a titleblock sidebar.

Uses matplotlib for vector PDF output (crisp at any zoom level).
"""

import math
from datetime import datetime
from typing import Dict, List, Any, Optional, Tuple

import matplotlib
matplotlib.use('Agg')  # Non-interactive backend for PDF generation
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, Rectangle
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.patheffects as pe


# Page dimensions in inches (11x17 landscape)
PAGE_WIDTH = 17.0
PAGE_HEIGHT = 11.0

# Page border margin
PAGE_MARGIN = 0.25

# Layout regions (inches)
SIDEBAR_WIDTH = 2.75
DRAWING_LEFT = PAGE_MARGIN + 0.15
DRAWING_BOTTOM = PAGE_MARGIN + 0.15
DRAWING_TOP = PAGE_MARGIN + 0.15
DRAWING_RIGHT = 0.2  # gap between drawing and sidebar
DRAWING_WIDTH = PAGE_WIDTH - PAGE_MARGIN - SIDEBAR_WIDTH - DRAWING_LEFT - DRAWING_RIGHT
DRAWING_HEIGHT = PAGE_HEIGHT - DRAWING_BOTTOM - DRAWING_TOP - PAGE_MARGIN

# Sidebar position (hugs right side, inside page border)
SIDEBAR_LEFT = PAGE_WIDTH - PAGE_MARGIN - SIDEBAR_WIDTH
SIDEBAR_BOTTOM = PAGE_MARGIN

# Pad colors (match site_preview.py)
PAD_COLORS = ['#C62828', '#1565C0', '#2E7D32', '#E65100', '#6A1B9A',
              '#00838F', '#AD1457', '#4E342E']


def generate_site_pdf(filepath: str,
                      group_layout: List[Dict],
                      device_positions: List[Dict],
                      pads: List[Dict],
                      colors: List[str],
                      topology: str,
                      device_label: str,
                      project_info: Dict[str, Any],
                      show_routes: bool = True,
                      align_on_motor: bool = True) -> bool:
    """
    Generate an 11x17 landscape PDF of the string allocation site preview.

    Args:
        filepath: Output PDF path.
        group_layout: List of group dicts from SitePreviewWindow.build_layout_data().
        device_positions: List of device position dicts.
        pads: List of pad dicts.
        colors: List of hex color strings for inverter assignments.
        topology: 'Distributed String', 'Centralized String', or 'Central Inverter'.
        device_label: 'CB' or 'SI'.
        project_info: Dict with keys like 'project_name', 'customer', 'location',
                      'estimate_name', 'topology', 'module_info', 'total_strings',
                      'total_devices', 'dc_ac_ratio', 'split_trackers', 'revision'.
        show_routes: Whether to draw L-shaped routes from devices to pads.
        align_on_motor: Whether to draw motor alignment lines.

    Returns:
        True on success, False on error.
    """
    try:
        fig = plt.figure(figsize=(PAGE_WIDTH, PAGE_HEIGHT))

        # Main drawing axes (world-coordinate space)
        ax = fig.add_axes([
            DRAWING_LEFT / PAGE_WIDTH,
            DRAWING_BOTTOM / PAGE_HEIGHT,
            DRAWING_WIDTH / PAGE_WIDTH,
            DRAWING_HEIGHT / PAGE_HEIGHT
        ])

        # Compute world bounds from layout data
        world_bounds = _compute_world_bounds(group_layout, device_positions, pads)
        if world_bounds is None:
            # Empty layout — still produce a PDF with titleblock
            ax.set_xlim(0, 100)
            ax.set_ylim(100, 0)
        else:
            xmin, xmax, ymin, ymax = world_bounds
            # Add padding (5% on each side)
            dx = (xmax - xmin) or 100
            dy = (ymax - ymin) or 100
            pad_x = dx * 0.05
            pad_y = dy * 0.05
            xmin -= pad_x
            xmax += pad_x
            ymin -= pad_y
            ymax += pad_y

            # Fit to drawing area aspect ratio
            aspect_drawing = DRAWING_WIDTH / DRAWING_HEIGHT
            aspect_world = dx / dy if dy > 0 else 1.0

            if aspect_world > aspect_drawing:
                # World is wider — expand Y to match
                needed_dy = dx / aspect_drawing
                extra = (needed_dy - dy) / 2
                ymin -= extra
                ymax += extra
            else:
                # World is taller — expand X to match
                needed_dx = dy * aspect_drawing
                extra = (needed_dx - dx) / 2
                xmin -= extra
                xmax += extra

            ax.set_xlim(xmin, xmax)
            ax.set_ylim(ymax, ymin)  # Invert Y so north is up

        ax.set_aspect('equal', adjustable='box')
        ax.axis('off')

        # --- Draw site elements ---
        max_width = _get_max_tracker_width(group_layout)

        _draw_groups(ax, group_layout, max_width)
        _draw_devices(ax, device_positions, device_label, pads)
        if show_routes:
            _draw_routes(ax, device_positions, pads, topology)
        _draw_pads(ax, pads)
        if align_on_motor:
            _draw_motor_alignment_lines(ax, group_layout)
        _draw_compass(ax)
        _draw_scale_bar(ax)
        _draw_watermark(ax)

        # --- Draw sidebar titleblock ---
        _draw_sidebar(fig, project_info)

        # --- Draw page border ---
        _draw_page_border(fig)

        # --- Save ---
        fig.savefig(filepath, format='pdf', dpi=150,
                    bbox_inches=None, pad_inches=0)
        plt.close(fig)
        return True

    except Exception as e:
        print(f"Error generating site PDF: {e}")
        import traceback
        traceback.print_exc()
        plt.close('all')
        return False


# ---------------------------------------------------------------------------
# World bounds
# ---------------------------------------------------------------------------

def _compute_world_bounds(group_layout, device_positions, pads):
    """Compute the bounding box of all drawn elements in world feet."""
    if not group_layout:
        return None

    xs, ys = [], []

    for group in group_layout:
        gx = group['x']
        gy = group['y']
        pitch = group.get('row_spacing_ft', 20)
        max_w = _get_max_tracker_width([group])

        for t_idx, tracker in enumerate(group['trackers']):
            t_width = tracker.get('width_ft', max_w)
            t_length = tracker.get('length_ft', 100)
            tx = gx + t_idx * pitch
            tx_offset = (max_w - t_width) / 2 if max_w > t_width else 0

            angle_y_offset = t_idx * pitch * math.tan(math.radians(group.get('driveline_angle', 0)))
            ty = gy + tracker.get('y_offset', 0) + angle_y_offset

            xs.extend([tx + tx_offset, tx + tx_offset + t_width])
            ys.extend([ty, ty + t_length])

    for dev in (device_positions or []):
        xs.extend([dev['x'], dev['x'] + dev['width_ft']])
        ys.extend([dev['y'], dev['y'] + dev['height_ft']])

    for pad in (pads or []):
        pw = pad.get('width_ft', 10)
        ph = pad.get('height_ft', 8)
        xs.extend([pad['x'], pad['x'] + pw])
        ys.extend([pad['y'], pad['y'] + ph])

    if not xs:
        return None

    return min(xs), max(xs), min(ys), max(ys)


# ---------------------------------------------------------------------------
# Drawing helpers
# ---------------------------------------------------------------------------

def _get_max_tracker_width(group_layout):
    """Return the maximum tracker width across all groups."""
    max_w = 6.0
    for group in group_layout:
        for tracker in group.get('trackers', []):
            w = tracker.get('width_ft', 6.0)
            if w > max_w:
                max_w = w
    return max_w


def _draw_groups(ax, group_layout, max_width):
    """Draw all groups with color-coded string rectangles, tracker outlines, and motors."""
    for group_idx, group_data in enumerate(group_layout):
        pitch = group_data.get('row_spacing_ft', 20)
        gx = group_data['x']
        gy = group_data['y']

        for t_idx, tracker in enumerate(group_data['trackers']):
            spt = tracker['strings_per_tracker']
            assignments = tracker.get('assignments', [])
            t_width = tracker.get('width_ft', max_width)
            t_length = tracker.get('length_ft', 100)

            tx = gx + t_idx * pitch
            tx_offset = (max_width - t_width) / 2 if max_width > t_width else 0

            # Driveline angle Y offset
            angle = group_data.get('driveline_angle', 0)
            angle_y_offset = t_idx * pitch * math.tan(math.radians(angle))
            ty = gy + tracker.get('y_offset', 0) + angle_y_offset

            # --- String heights (handle half-string trackers) ---
            half_string_side = tracker.get('half_string_side', None)
            if half_string_side and spt > 0:
                full_str_count = int(spt)
                frac = spt - full_str_count
                if frac > 0:
                    full_str_count_int = full_str_count
                    full_height = t_length / (full_str_count_int + frac)
                    partial_height = full_height * frac
                    string_heights = []
                    if half_string_side == 'north':
                        string_heights.append(partial_height)
                        for _ in range(full_str_count_int):
                            string_heights.append(full_height)
                    else:
                        for _ in range(full_str_count_int):
                            string_heights.append(full_height)
                        string_heights.append(partial_height)
                else:
                    string_height = t_length / spt if spt > 0 else t_length
                    string_heights = [string_height] * int(spt)
            else:
                int_spt = int(spt) if spt == int(spt) else int(spt) + 1
                if int_spt > 0:
                    string_height = t_length / int_spt
                    string_heights = [string_height] * int_spt
                else:
                    string_heights = [t_length]

            # Build string colors from assignments
            string_colors = []
            for assignment in assignments:
                for _ in range(assignment.get('strings', 0)):
                    string_colors.append(assignment.get('color', '#DDDDDD'))

            # Pad to match string count
            while len(string_colors) < len(string_heights):
                string_colors.append('#DDDDDD')

            # Draw each string rectangle
            for s_idx in range(len(string_heights)):
                sy = ty + sum(string_heights[:s_idx])
                sh = string_heights[s_idx]
                color = string_colors[s_idx] if s_idx < len(string_colors) else '#DDDDDD'

                rect = Rectangle(
                    (tx + tx_offset, sy), t_width, sh,
                    facecolor=color, edgecolor='#555555', linewidth=0.3
                )
                ax.add_patch(rect)

            # Tracker outline
            outline = Rectangle(
                (tx + tx_offset - 0.5, ty - 0.5),
                t_width + 1.0, t_length + 1.0,
                facecolor='none', edgecolor='#222222', linewidth=0.4
            )
            ax.add_patch(outline)

            # Motor indicator
            if tracker.get('has_motor', False):
                motor_y = tracker['motor_y_ft']
                motor_gap = tracker.get('motor_gap_ft', 1.0)
                motor_world_y = ty + motor_y
                motor_x1 = tx + tx_offset - 0.3
                motor_w = t_width + 0.6
                motor_rect = Rectangle(
                    (motor_x1, motor_world_y), motor_w, motor_gap,
                    facecolor='#666666', edgecolor='#444444', linewidth=0.3
                )
                ax.add_patch(motor_rect)
                # Motor dot
                motor_cx = motor_x1 + motor_w / 2
                motor_cy = motor_world_y + motor_gap / 2
                ax.plot(motor_cx, motor_cy, 'o', color='#FF8800',
                        markersize=2, markeredgecolor='#CC6600', markeredgewidth=0.3)


def _draw_devices(ax, device_positions, device_label, pads):
    """Draw CB/SI device markers."""
    if not device_positions:
        return

    # Build device -> pad lookup for outline coloring
    device_to_pad = {}
    if pads:
        for pad_idx, pad in enumerate(pads):
            for dev_idx in pad.get('assigned_devices', []):
                device_to_pad[dev_idx] = pad_idx

    for dev_idx, dev in enumerate(device_positions):
        dx = dev['x']
        dy = dev['y']
        dw = dev['width_ft']
        dh = dev['height_ft']
        label = dev.get('label', f'{device_label}{dev_idx + 1}')

        # Fill color
        if device_label == 'CB':
            fill_color = '#FF9800'
        else:
            fill_color = '#2196F3'

        # Outline color from pad assignment
        if dev_idx in device_to_pad and pads:
            pad_idx = device_to_pad[dev_idx]
            outline_color = PAD_COLORS[pad_idx % len(PAD_COLORS)]
        else:
            outline_color = '#E65100' if device_label == 'CB' else '#0D47A1'

        rect = Rectangle(
            (dx, dy), dw, dh,
            facecolor=fill_color, edgecolor=outline_color, linewidth=0.8
        )
        ax.add_patch(rect)

        # Label above device
        cx = dx + dw / 2
        ax.text(cx, dy - 1.0, label,
                fontsize=5, fontweight='bold', color='#333333',
                ha='center', va='bottom', fontfamily='sans-serif')


def _draw_routes(ax, device_positions, pads, topology):
    """Draw L-shaped Manhattan routes from devices to their assigned pads."""
    if not pads or not device_positions:
        return

    device_to_pad = {}
    for pad_idx, pad in enumerate(pads):
        for dev_idx in pad.get('assigned_devices', []):
            device_to_pad[dev_idx] = pad_idx

    for dev_idx, dev in enumerate(device_positions):
        pad_idx = device_to_pad.get(dev_idx)
        if pad_idx is None or pad_idx >= len(pads):
            continue

        pad = pads[pad_idx]
        dev_cx = dev['x'] + dev['width_ft'] / 2
        dev_cy = dev['y'] + dev['height_ft'] / 2
        pad_cx = pad['x'] + pad.get('width_ft', 10) / 2
        pad_cy = pad['y'] + pad.get('height_ft', 8) / 2

        color = PAD_COLORS[pad_idx % len(PAD_COLORS)]
        linestyle = '--' if topology == 'Distributed String' else '-'

        # L-shaped: E-W first, then N-S
        ax.plot([dev_cx, pad_cx, pad_cx], [dev_cy, dev_cy, pad_cy],
                color=color, linewidth=0.4, linestyle=linestyle, alpha=0.6)


def _draw_pads(ax, pads):
    """Draw inverter pad rectangles."""
    if not pads:
        return

    for pad_idx, pad in enumerate(pads):
        px = pad['x']
        py = pad['y']
        pw = pad.get('width_ft', 10)
        ph = pad.get('height_ft', 8)
        label = pad.get('label', f'Pad {pad_idx + 1}')
        base_color = PAD_COLORS[pad_idx % len(PAD_COLORS)]

        rect = Rectangle(
            (px, py), pw, ph,
            facecolor=base_color, edgecolor='#222222', linewidth=0.6
        )
        ax.add_patch(rect)

        # Label
        cx = px + pw / 2
        cy = py + ph / 2
        ax.text(cx, cy, label, fontsize=5, fontweight='bold',
                color='white', ha='center', va='center', fontfamily='sans-serif')

        # Device count
        num_assigned = len(pad.get('assigned_devices', []))
        if num_assigned > 0:
            ax.text(cx, cy + ph * 0.25, f"({num_assigned} devices)",
                    fontsize=3.5, color='#CCCCCC', ha='center', va='center',
                    fontfamily='sans-serif')


def _draw_motor_alignment_lines(ax, group_layout):
    """Draw driveline across each group at its motor Y position."""
    for group in group_layout:
        motor_y_ft = group.get('motor_y_ft')
        if motor_y_ft is None:
            continue

        gx = group['x']
        gy = group['y']
        trackers = group.get('trackers', [])
        if not trackers:
            continue

        pitch = group.get('row_spacing_ft', 20)
        angle = group.get('driveline_angle', 0)
        n_trackers = len(trackers)

        # Start and end X
        x_start = gx - 2
        x_end = gx + (n_trackers - 1) * pitch + trackers[-1].get('width_ft', 6) + 2

        # Y positions at start and end (following angle)
        y_start = gy + motor_y_ft
        y_end = gy + motor_y_ft + (x_end - x_start) * math.tan(math.radians(angle))

        ax.plot([x_start, x_end], [y_start, y_end],
                color='#FF8800', linewidth=0.4, linestyle='--', alpha=0.5)


def _draw_compass(ax):
    """Draw a north arrow in the top-right corner of the drawing area."""
    # Use axes-relative coordinates
    ax.annotate('', xy=(0.97, 0.97), xytext=(0.97, 0.92),
                xycoords='axes fraction', textcoords='axes fraction',
                arrowprops=dict(arrowstyle='->', color='#333333', lw=1.2))
    ax.text(0.97, 0.99, 'N', transform=ax.transAxes,
            fontsize=7, fontweight='bold', color='#333333',
            ha='center', va='top', fontfamily='sans-serif')


def _draw_scale_bar(ax):
    """Draw a scale bar in the bottom-left corner."""
    xlim = ax.get_xlim()
    ylim = ax.get_ylim()  # ylim[0] > ylim[1] because Y is inverted

    # Target bar = ~15% of view width
    view_width = abs(xlim[1] - xlim[0])
    target_ft = view_width * 0.15

    nice_values = [5, 10, 20, 25, 50, 100, 200, 250, 500, 1000, 2000]
    bar_ft = nice_values[0]
    for v in nice_values:
        if v <= target_ft:
            bar_ft = v
        else:
            break

    # Position at bottom-left in world coords
    # ylim[0] is the max Y (bottom of drawing since inverted)
    x_start = xlim[0] + view_width * 0.03
    y_pos = ylim[0] - abs(ylim[0] - ylim[1]) * 0.03  # near bottom

    ax.plot([x_start, x_start + bar_ft], [y_pos, y_pos],
            color='#333333', linewidth=1.0)
    # End ticks
    tick_h = abs(ylim[0] - ylim[1]) * 0.01
    ax.plot([x_start, x_start], [y_pos - tick_h, y_pos + tick_h],
            color='#333333', linewidth=1.0)
    ax.plot([x_start + bar_ft, x_start + bar_ft], [y_pos - tick_h, y_pos + tick_h],
            color='#333333', linewidth=1.0)

    ax.text(x_start + bar_ft / 2, y_pos - tick_h * 2, f"{bar_ft} ft",
            fontsize=6, color='#333333', ha='center', va='top',
            fontfamily='sans-serif')


def _draw_watermark(ax):
    """Draw diagonal 'PRELIMINARY — FOR REVIEW ONLY' watermark across drawing area."""
    ax.text(0.5, 0.5, 'PRELIMINARY — FOR REVIEW ONLY',
            transform=ax.transAxes,
            fontsize=28, color='#FF0000', alpha=0.12,
            ha='center', va='center', rotation=30,
            fontweight='bold', fontfamily='sans-serif',
            path_effects=[pe.withStroke(linewidth=0.5, foreground='#FF000010')])


# ---------------------------------------------------------------------------
# Page Border
# ---------------------------------------------------------------------------

def _draw_page_border(fig):
    """Draw a border around the entire page that lines up with the titleblock."""
    border_ax = fig.add_axes([0, 0, 1, 1], zorder=-1)
    border_ax.set_xlim(0, PAGE_WIDTH)
    border_ax.set_ylim(0, PAGE_HEIGHT)
    border_ax.axis('off')

    border_ax.add_patch(Rectangle(
        (PAGE_MARGIN, PAGE_MARGIN),
        PAGE_WIDTH - 2 * PAGE_MARGIN,
        PAGE_HEIGHT - 2 * PAGE_MARGIN,
        facecolor='none', edgecolor='black', linewidth=1.2,
        clip_on=False
    ))


# ---------------------------------------------------------------------------
# Sidebar / Titleblock
# ---------------------------------------------------------------------------

def _draw_sidebar(fig, project_info):
    """Draw an engineering-drawing-style titleblock on the right side.

    Layout matches Ampacity's standard site plan titleblock:
      - Customer Approval box at top
      - Ampacity Renewables branding + company address
      - Large vertical project info box (rotated text)
      - Description / Rev / Issued table
      - Sheet No, Original Plan Size
      - Copyright disclaimer
    """
    SIDEBAR_H = PAGE_HEIGHT - 2 * PAGE_MARGIN  # height inside border

    ax = fig.add_axes([
        SIDEBAR_LEFT / PAGE_WIDTH,
        SIDEBAR_BOTTOM / PAGE_HEIGHT,
        SIDEBAR_WIDTH / PAGE_WIDTH,
        SIDEBAR_H / PAGE_HEIGHT
    ])
    ax.set_xlim(0, SIDEBAR_WIDTH)
    ax.set_ylim(0, SIDEBAR_H)
    ax.axis('off')

    W = SIDEBAR_WIDTH   # 2.75 inches
    H = SIDEBAR_H       # ~10.5 inches
    LW = 0.8            # standard border linewidth

    # Outer border
    ax.add_patch(Rectangle((0, 0), W, H,
                           facecolor='white', edgecolor='black', linewidth=LW * 1.5))

    # =====================================================================
    # Vertical Y positions (from bottom = 0 upward = H)
    # =====================================================================
    COPYRIGHT_H = 0.55
    PLAN_SIZE_H = 0.25
    SHEET_NO_H = 0.30
    REV_H = 0.25
    DESC_H = 0.45
    # Project box fills the remaining middle space
    # Branding box at top
    BRANDING_H = 1.30
    CUSTOMER_H = 0.35

    # Stack from bottom
    y_copyright_bot = 0.0
    y_copyright_top = COPYRIGHT_H

    y_plansize_bot = y_copyright_top
    y_plansize_top = y_plansize_bot + PLAN_SIZE_H

    y_sheetno_bot = y_plansize_top
    y_sheetno_top = y_sheetno_bot + SHEET_NO_H

    y_rev_bot = y_sheetno_top
    y_rev_top = y_rev_bot + REV_H

    y_desc_bot = y_rev_top
    y_desc_top = y_desc_bot + DESC_H

    y_branding_top = H
    y_branding_bot = H - BRANDING_H

    y_customer_top = y_branding_bot
    y_customer_bot = y_customer_top - CUSTOMER_H

    y_project_bot = y_desc_top
    y_project_top = y_customer_bot

    def _hline(y):
        ax.plot([0, W], [y, y], color='black', linewidth=LW, clip_on=False)

    def _vline(x, yb, yt):
        ax.plot([x, x], [yb, yt], color='black', linewidth=LW, clip_on=False)

    def _box(x, yb, w, h, **kw):
        ax.add_patch(Rectangle((x, yb), w, h,
                               facecolor='none', edgecolor='black',
                               linewidth=LW, **kw))

    # Draw all horizontal dividers
    for yy in [y_copyright_top, y_plansize_top, y_sheetno_top,
               y_rev_top, y_desc_top, y_project_top, y_customer_top,
               y_branding_bot]:
        _hline(yy)

    # =====================================================================
    # 1. COPYRIGHT (bottom strip)
    # =====================================================================
    ax.text(W / 2, y_copyright_bot + COPYRIGHT_H / 2,
            '\u00A92026 AMPACITY: THE DRAWINGS,\n'
            'SPECIFICATIONS, AND OTHER DOCUMENTS\n'
            'RELATED TO THIS PROJECT ARE\n'
            'PROPRIETARY AND CONFIDENTIAL.',
            fontsize=3.2, color='#333333', ha='center', va='center',
            fontfamily='sans-serif', fontstyle='italic', linespacing=1.3)

    # =====================================================================
    # 2. ORIGINAL PLAN SIZE
    # =====================================================================
    ax.text(W / 2, y_plansize_bot + PLAN_SIZE_H / 2,
            'ORIGINAL PLAN SIZE: 11\u00d717',
            fontsize=5, fontweight='bold', color='black',
            ha='center', va='center', fontfamily='sans-serif')

    # =====================================================================
    # 3. SHEET NO
    # =====================================================================
    _vline(W * 0.38, y_sheetno_bot, y_sheetno_top)
    ax.text(W * 0.19, y_sheetno_bot + SHEET_NO_H / 2,
            'SHEET NO:',
            fontsize=4.5, fontweight='bold', color='black',
            ha='center', va='center', fontfamily='sans-serif')
    ax.text(W * 0.69, y_sheetno_bot + SHEET_NO_H / 2,
            'STRING ALLOCATION',
            fontsize=5, fontweight='bold', color='black',
            ha='center', va='center', fontfamily='sans-serif')

    # =====================================================================
    # 4. REV / ISSUED / DESCRIPTION header row
    # =====================================================================
    col_rev_w = 0.30
    col_issued_w = 0.50
    # rest is description

    _vline(col_rev_w, y_rev_bot, y_desc_top)
    _vline(col_rev_w + col_issued_w, y_rev_bot, y_desc_top)

    # Header labels
    ax.text(col_rev_w / 2, y_rev_bot + REV_H / 2,
            'REV', fontsize=4.5, fontweight='bold', color='black',
            ha='center', va='center', fontfamily='sans-serif')
    ax.text(col_rev_w + col_issued_w / 2, y_rev_bot + REV_H / 2,
            'ISSUED', fontsize=4.5, fontweight='bold', color='black',
            ha='center', va='center', fontfamily='sans-serif')
    ax.text((col_rev_w + col_issued_w + W) / 2, y_rev_bot + REV_H / 2,
            'DESCRIPTION', fontsize=4.5, fontweight='bold', color='black',
            ha='center', va='center', fontfamily='sans-serif')

    # =====================================================================
    # 5. DESCRIPTION content row
    # =====================================================================
    rev = project_info.get('revision', '0')
    date_str = datetime.now().strftime('%m/%d/%Y')

    ax.text(col_rev_w / 2, y_desc_bot + DESC_H / 2,
            str(rev), fontsize=5, color='black',
            ha='center', va='center', fontfamily='sans-serif')
    ax.text(col_rev_w + col_issued_w / 2, y_desc_bot + DESC_H / 2,
            date_str, fontsize=4.5, color='black',
            ha='center', va='center', fontfamily='sans-serif')
    ax.text(col_rev_w + col_issued_w + 0.06, y_desc_bot + DESC_H / 2,
            'PRELIMINARY',
            fontsize=4.5, color='black',
            ha='left', va='center', fontfamily='sans-serif')

    # =====================================================================
    # 6. PROJECT INFO (large box with rotated text)
    # =====================================================================
    _box(0, y_project_bot, W, y_project_top - y_project_bot)

    project_lines = []
    customer = project_info.get('customer', '')
    project_name = project_info.get('project_name', '')
    location = project_info.get('location', '')
    estimate = project_info.get('estimate_name', '')
    topology = project_info.get('topology', '')
    module_info = project_info.get('module_info', '')
    total_strings = project_info.get('total_strings', '')
    total_devices = project_info.get('total_devices', '')
    dc_ac = project_info.get('dc_ac_ratio', '')
    split = project_info.get('split_trackers', '')

    if customer:
        project_lines.append(customer.upper())
    if project_name:
        project_lines.append(project_name.upper())
    if location:
        project_lines.append(location.upper())
    if estimate:
        project_lines.append(f'ESTIMATE: {estimate.upper()}')

    # Add technical summary lines
    tech_parts = []
    if topology:
        tech_parts.append(topology)
    if total_devices:
        tech_parts.append(f'{total_devices}')
    if total_strings:
        tech_parts.append(f'{total_strings} Strings')
    if dc_ac:
        tech_parts.append(f'DC:AC {dc_ac}')
    if split:
        tech_parts.append(f'{split} Split Trackers')
    if tech_parts:
        project_lines.append(' | '.join(tech_parts))

    if module_info:
        project_lines.append(module_info.upper())

    project_text = '\n'.join(project_lines)
    project_box_cy = (y_project_bot + y_project_top) / 2
    project_box_height = y_project_top - y_project_bot

    # Rotated 90° text (read from bottom to top)
    ax.text(W / 2, project_box_cy, project_text,
            fontsize=6.5, fontweight='bold', color='black',
            ha='center', va='center', fontfamily='sans-serif',
            rotation=90, linespacing=1.5)

    # =====================================================================
    # 7. CUSTOMER APPROVAL
    # =====================================================================
    _box(0, y_customer_bot, W, CUSTOMER_H)
    ax.text(0.08, y_customer_bot + CUSTOMER_H / 2,
            'CUSTOMER APPROVAL:',
            fontsize=4.5, fontweight='bold', color='black',
            ha='left', va='center', fontfamily='sans-serif')
    # Line for signature
    ax.plot([W * 0.58, W * 0.95],
            [y_customer_bot + CUSTOMER_H * 0.35, y_customer_bot + CUSTOMER_H * 0.35],
            color='black', linewidth=0.4)

    # =====================================================================
    # 8. BRANDING (top box)
    # =====================================================================
    _box(0, y_branding_bot, W, BRANDING_H)

    # Ampacity Renewables title
    brand_cy = y_branding_bot + BRANDING_H * 0.68
    ax.text(W / 2, brand_cy, 'AMPACITY',
            fontsize=13, fontweight='bold', color='#1a5276',
            ha='center', va='center', fontfamily='sans-serif')
    ax.text(W / 2, brand_cy - 0.19, 'R E N E W A B L E S',
            fontsize=5.5, fontweight='bold', color='#1a5276',
            ha='center', va='center', fontfamily='sans-serif')

    # Company info
    company_lines = (
        'AMPACITY, LLC\n'
        '305 DELA VINA AVE\n'
        'MONTEREY, CA 93940\n'
        'CSLB #1000391  ROC #355463'
    )
    ax.text(W / 2, y_branding_bot + BRANDING_H * 0.20, company_lines,
            fontsize=4, color='#333333',
            ha='center', va='center', fontfamily='sans-serif',
            linespacing=1.3)