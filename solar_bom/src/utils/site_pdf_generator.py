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
                      align_on_motor: bool = True,
                      wiring_specs: Optional[List[Dict]] = None) -> bool:
    """
    Generate an 11x17 landscape PDF of the string allocation site preview,
    optionally followed by DC cabling wiring diagram pages.

    Args:
        filepath: Output PDF path.
        group_layout: List of group dicts from SitePreviewWindow.build_layout_data().
        device_positions: List of device position dicts.
        pads: List of pad dicts.
        colors: List of hex color strings for inverter assignments.
        topology: 'Distributed String', 'Centralized String', or 'Central Inverter'.
        device_label: 'CB' or 'SI'.
        project_info: Dict with project metadata for the titleblock.
        show_routes: Whether to draw L-shaped routes from devices to pads.
        align_on_motor: Whether to draw motor alignment lines.
        wiring_specs: Optional list of dicts describing unique tracker wiring configs.

    Returns:
        True on success, False on error.
    """
    from matplotlib.backends.backend_pdf import PdfPages

    try:
        with PdfPages(filepath) as pdf:
            # ===== Page 1: Site Preview =====
            fig = _create_site_page(group_layout, device_positions, pads,
                                     colors, topology, device_label, project_info,
                                     show_routes, align_on_motor)
            pdf.savefig(fig)
            plt.close(fig)

            # ===== Pages 2+: Wiring Diagrams =====
            if wiring_specs:
                wiring_figs = _create_wiring_pages(wiring_specs, project_info)
                for wfig in wiring_figs:
                    pdf.savefig(wfig)
                    plt.close(wfig)

        return True

    except Exception as e:
        print(f"Error generating site PDF: {e}")
        import traceback
        traceback.print_exc()
        plt.close('all')
        return False


def _create_site_page(group_layout, device_positions, pads, colors,
                       topology, device_label, project_info,
                       show_routes, align_on_motor):
    """Create the Page 1 figure (site preview)."""
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

    return fig


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

            # Motor alignment (must match _draw_groups logic)
            group_motor_y = group.get('motor_y_ft', None)
            if tracker.get('has_motor', False) and group_motor_y is not None:
                ty = gy + (group_motor_y - tracker['motor_y_ft']) + angle_y_offset
            else:
                max_group_length = group.get('length_ft', tracker.get('length_ft', 100))
                ty = gy + (max_group_length - tracker.get('length_ft', 100)) / 2 + angle_y_offset

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

            # Motor alignment: match site_preview.py draw() logic
            group_motor_y = group_data.get('motor_y_ft', None)
            if tracker.get('has_motor', False) and group_motor_y is not None:
                ty = gy + (group_motor_y - tracker['motor_y_ft']) + angle_y_offset
            else:
                # Center fallback
                max_group_length = group_data.get('length_ft', t_length)
                ty = gy + (max_group_length - t_length) / 2 + angle_y_offset

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


# ===========================================================================
# Wiring Diagram Pages
# ===========================================================================

MAX_DIAGRAMS_PER_PAGE = 2


def _create_wiring_pages(wiring_specs, project_info):
    """Create figures for wiring diagram pages, up to 2 diagrams per page."""
    figures = []
    for page_start in range(0, len(wiring_specs), MAX_DIAGRAMS_PER_PAGE):
        page_specs = wiring_specs[page_start:page_start + MAX_DIAGRAMS_PER_PAGE]
        fig = _create_single_wiring_page(page_specs, project_info, page_start)
        figures.append(fig)
    return figures


def _create_single_wiring_page(specs, project_info, start_idx):
    """Create one 11x17 page with 1-2 wiring diagrams plus titleblock."""
    fig = plt.figure(figsize=(PAGE_WIDTH, PAGE_HEIGHT))

    n = len(specs)
    # Drawing area (inside border, excluding titleblock)
    draw_left = PAGE_MARGIN + 0.3
    draw_right_edge = PAGE_WIDTH - PAGE_MARGIN - SIDEBAR_WIDTH - 0.2
    draw_bottom = PAGE_MARGIN + 0.3
    draw_top = PAGE_HEIGHT - PAGE_MARGIN - 0.3
    draw_w = draw_right_edge - draw_left
    draw_h = draw_top - draw_bottom

    for i, spec in enumerate(specs):
        letter = chr(ord('A') + start_idx + i)

        # Divide height evenly, top diagram first
        slot_h = draw_h / n
        slot_bottom = draw_top - (i + 1) * slot_h + 0.15
        slot_top = draw_top - i * slot_h - 0.15

        ax = fig.add_axes([
            draw_left / PAGE_WIDTH,
            slot_bottom / PAGE_HEIGHT,
            draw_w / PAGE_WIDTH,
            (slot_top - slot_bottom) / PAGE_HEIGHT
        ])

        _draw_single_wiring_diagram(ax, spec, letter)

    _draw_sidebar(fig, project_info)
    _draw_page_border(fig)

    # --- Compass rose (north = left in horizontal wiring layout) ---
    cx = (draw_right_edge - 0.6) / PAGE_WIDTH  # figure-fraction X
    cy = (draw_top - 0.5) / PAGE_HEIGHT         # figure-fraction Y
    r = 0.018  # radius in figure-fraction units
    tick = r * 0.35

    compass_ax = fig.add_axes([cx - r * 1.8, cy - r * 1.8, r * 3.6, r * 3.6])
    compass_ax.set_xlim(-1.2, 1.2)
    compass_ax.set_ylim(-1.2, 1.2)
    compass_ax.set_aspect('equal')
    compass_ax.axis('off')

    # Outer circle
    circle = plt.Circle((0, 0), 1.0, fill=False, edgecolor='#333333', linewidth=0.8)
    compass_ax.add_patch(circle)

    # North arrow (points LEFT, filled triangle)
    compass_ax.fill([-1.0, -0.3, -0.3], [0, 0.18, -0.18], color='#333333')
    # South tick (right)
    compass_ax.plot([0.6, 1.0], [0, 0], color='#333333', linewidth=0.6)
    # East tick (up)
    compass_ax.plot([0, 0], [0.6, 1.0], color='#333333', linewidth=0.6)
    # West tick (down)
    compass_ax.plot([0, 0], [-0.6, -1.0], color='#333333', linewidth=0.6)

    # Labels
    compass_ax.text(-1.05, 0, 'N', fontsize=5, fontweight='bold', color='#333333',
                    ha='right', va='center', fontfamily='sans-serif')
    compass_ax.text(1.05, 0, 'S', fontsize=4, color='#777777',
                    ha='left', va='center', fontfamily='sans-serif')
    compass_ax.text(0, 1.05, 'E', fontsize=4, color='#777777',
                    ha='center', va='bottom', fontfamily='sans-serif')
    compass_ax.text(0, -1.05, 'W', fontsize=4, color='#777777',
                    ha='center', va='top', fontfamily='sans-serif')

    return fig


# ---------------------------------------------------------------------------
# Single wiring diagram
# ---------------------------------------------------------------------------

# Layout constants (data-coordinate units)
# Module dimensions are computed per-diagram from real specs.
# These are defaults / spacing constants.
_MOD_GAP = 0.08     # gap between modules within a string
_STRING_GAP = 0.6    # gap between adjacent strings
_MOTOR_GAP = 2.5     # gap for driveline / motor
_MM_TO_DU = 0.00065  # mm to data-unit conversion (1mm = 0.00065 DU)


def _draw_single_wiring_diagram(ax, spec, letter):
    """Draw one horizontal tracker wiring diagram.

    spec keys:
        strings_per_tracker, modules_per_string, harness_sizes,
        has_motor, motor_position_after_string, polarity_convention,
        wire_gauges: {string, harness: {size: gauge}, whip: {size: gauge}},
    """
    spt = int(spec['strings_per_tracker'])
    mps = int(spec['modules_per_string'])
    harness_sizes = spec['harness_sizes']  # e.g. [7, 6]
    has_motor = spec.get('has_motor', True)
    motor_after = spec.get('motor_position_after_string', spt // 2)
    if has_motor:
        motor_after = max(0, min(motor_after, spt))
    polarity = spec.get('polarity_convention', 'positive_north')
    device_position = spec.get('device_position', 'south')  # 'north', 'south', or 'middle'
    wire_gauges = spec.get('wire_gauges', {})

    string_gauge = wire_gauges.get('string', '10 AWG')
    harness_gauge_map = wire_gauges.get('harness', {})
    whip_gauge_map = wire_gauges.get('whip', {})

    # --- Module dimensions from real specs ---
    mod_width_mm = spec.get('module_width_mm', 1134)
    mod_length_mm = spec.get('module_length_mm', 2384)
    orientation = spec.get('module_orientation', 'Portrait')

    # In the horizontal wiring diagram:
    #   X direction = along the tracker (N-S) = module "along" dimension
    #   Y direction = across the tracker (E-W) = module "across" dimension
    if orientation == 'Portrait':
        mod_along_mm = mod_width_mm    # N-S = width in portrait
        mod_across_mm = mod_length_mm  # E-W = length in portrait
    else:
        mod_along_mm = mod_length_mm
        mod_across_mm = mod_width_mm

    _MOD_W = mod_along_mm * _MM_TO_DU   # width in data units (X direction)
    _MOD_H = mod_across_mm * _MM_TO_DU  # height in data units (Y direction)

    # --- Compute X positions for each string ---
    string_w = mps * _MOD_W + (mps - 1) * _MOD_GAP
    string_x_starts = []
    x_cursor = 0.0
    if has_motor and motor_after == 0:
        x_cursor += _MOTOR_GAP
    for s_idx in range(spt):
        string_x_starts.append(x_cursor)
        x_cursor += string_w
        if has_motor and motor_after > 0 and s_idx == motor_after - 1:
            x_cursor += _MOTOR_GAP
        elif s_idx < spt - 1:
            x_cursor += _STRING_GAP

    total_w = x_cursor

    # --- Fixed scale coordinate system ---
    # Use a constant scale so modules are the same physical size on every diagram.
    # _SCALE data units = 1 inch on paper.
    _SCALE = 7.0
    ax_pos = ax.get_position()
    ax_w_in = ax_pos.width * PAGE_WIDTH
    ax_h_in = ax_pos.height * PAGE_HEIGHT
    x_range = ax_w_in * _SCALE
    y_range = ax_h_in * _SCALE

    # Center the tracker horizontally in the available space
    x_margin = (x_range - total_w) / 2 if x_range > total_w else x_range * 0.02
    x_min = -x_margin
    x_max = x_min + x_range

    y_base = 0.0
    y_top = y_range

    def Y(frac):
        """Convert fraction (0=bottom, 1=top) to data Y coordinate."""
        return y_base + frac * y_range

    ax.set_xlim(x_min, x_max)
    ax.set_ylim(y_base, y_top)
    ax.set_aspect('equal', adjustable='box')
    ax.axis('off')

    # Key vertical positions (as fractions)
    F_TITLE = 0.95
    F_HARNESS_BRACKET = 0.85
    F_STRING_LABEL = 0.80
    F_MOD_TOP = 0.73
    F_MOD_BOT = 0.62
    F_POLARITY = 0.58
    F_ROUTE_TOP = 0.54
    F_MERGE = 0.40
    F_LABEL_BOT = 0.05

    mod_h = _MOD_H  # real module height from dimensions

    # --- Determine polarity per string ---
    polarity_lower = polarity.lower().replace(' ', '_')
    if 'negative' in polarity_lower and 'south' in polarity_lower:
        pos_on_north = True
    elif 'negative' in polarity_lower and 'north' in polarity_lower:
        pos_on_north = False
    elif 'positive' in polarity_lower and 'north' in polarity_lower:
        pos_on_north = True
    elif 'positive' in polarity_lower and 'south' in polarity_lower:
        pos_on_north = False
    else:
        pos_on_north = True
    polarity_at = [pos_on_north] * spt

    # --- Module Y position: center modules at the F_MOD vertical zone ---
    mod_center_y = Y((F_MOD_TOP + F_MOD_BOT) / 2)
    mod_bot_y = mod_center_y - mod_h / 2
    mod_top_y = mod_center_y + mod_h / 2

    # --- Draw modules ---
    for s_idx in range(spt):
        sx = string_x_starts[s_idx]
        for m in range(mps):
            mx = sx + m * (_MOD_W + _MOD_GAP)
            rect = Rectangle((mx, mod_bot_y), _MOD_W, mod_h,
                              facecolor='#D6EAF8', edgecolor='#2C5282', linewidth=0.3)
            ax.add_patch(rect)

    # --- Motor gap indicator ---
    if has_motor:
        if motor_after == 0:
            mg_left = 0
            mg_right = _MOTOR_GAP
        elif motor_after >= spt:
            mg_left = string_x_starts[spt - 1] + string_w
            mg_right = total_w
        else:
            mg_left = string_x_starts[motor_after - 1] + string_w + _STRING_GAP * 0.15
            mg_right = string_x_starts[motor_after] - _STRING_GAP * 0.15
        mw = mg_right - mg_left
        if mw > 0.2:
            motor_rect = Rectangle((mg_left, mod_bot_y - mod_h * 0.05),
                                    mw, mod_h * 1.1,
                                    facecolor='#AAAAAA', edgecolor='#555555', linewidth=0.4)
            ax.add_patch(motor_rect)

    # --- Polarity markers ---
    for s_idx in range(spt):
        sx = string_x_starts[s_idx]
        left_x = sx - 0.3
        right_x = sx + string_w + 0.3
        y_pol = mod_center_y

        if polarity_at[s_idx]:
            left_label, right_label = '(+)', '(-)'
            left_color, right_color = '#CC0000', '#0000CC'
        else:
            left_label, right_label = '(-)', '(+)'
            left_color, right_color = '#0000CC', '#CC0000'

        ax.text(left_x, y_pol, left_label, fontsize=3.5, color=left_color,
                ha='right', va='center', fontfamily='sans-serif', fontweight='bold')
        ax.text(right_x, y_pol, right_label, fontsize=3.5, color=right_color,
                ha='left', va='center', fontfamily='sans-serif', fontweight='bold')

    # --- String labels ---
    for s_idx in range(spt):
        sx = string_x_starts[s_idx]
        cx = sx + string_w / 2
        ax.text(cx, Y(F_STRING_LABEL), f'{mps}-MODULE STRING',
                fontsize=2.8, color='black', ha='center', va='bottom',
                fontfamily='sans-serif')

    # --- Harness bracket labels ---
    harness_string_offset = 0
    for h_idx, h_size in enumerate(harness_sizes):
        h_start_str = harness_string_offset
        h_end_str = harness_string_offset + h_size - 1
        x_left = string_x_starts[h_start_str]
        x_right = string_x_starts[h_end_str] + string_w
        cx = (x_left + x_right) / 2

        bracket_y = Y(F_HARNESS_BRACKET)
        tick = y_range * 0.015
        ax.plot([x_left, x_left, x_right, x_right],
                [bracket_y - tick, bracket_y, bracket_y, bracket_y - tick],
                color='black', linewidth=0.5)
        ax.text(cx, bracket_y + tick * 0.5, f'{h_size}-STRING HARNESS',
                fontsize=3.5, color='black', ha='center', va='bottom',
                fontfamily='sans-serif', fontweight='bold')

        harness_string_offset += h_size

    # --- Compute device X target for home run routing ---
    if device_position == 'south':
        _device_x = total_w + 2.0
    elif device_position == 'north':
        _device_x = -2.0
    else:
        # Middle — motor gap center
        if has_motor and 0 < motor_after < spt:
            _device_x = (string_x_starts[motor_after - 1] + string_w + string_x_starts[motor_after]) / 2
        else:
            _device_x = total_w / 2

    # --- Routing: positive (red) and negative (blue) ---
    harness_string_offset = 0
    for h_idx, h_size in enumerate(harness_sizes):
        h_strings = list(range(harness_string_offset, harness_string_offset + h_size))
        h_gauge = harness_gauge_map.get(str(h_size), harness_gauge_map.get(h_size, '10 AWG'))
        w_gauge = whip_gauge_map.get(str(h_size), whip_gauge_map.get(h_size, '10 AWG'))

        pos_xs = []
        neg_xs = []
        for s_idx in h_strings:
            sx = string_x_starts[s_idx]
            left_x = sx + _MOD_W * 0.5
            right_x = sx + string_w - _MOD_W * 0.5
            if polarity_at[s_idx]:
                pos_xs.append(left_x)
                neg_xs.append(right_x)
            else:
                pos_xs.append(right_x)
                neg_xs.append(left_x)

        # Bias merge point toward device location
        if device_position == 'south':
            pos_merge_x = max(pos_xs)
        elif device_position == 'north':
            pos_merge_x = min(pos_xs)
        else:
            # Middle — bias toward motor gap
            if harness_string_offset < motor_after:
                pos_merge_x = max(pos_xs)  # north harness → merge at south end (toward motor)
            else:
                pos_merge_x = min(pos_xs)  # south harness → merge at north end (toward motor)
        _draw_harness_routing(ax, pos_xs, pos_merge_x,
                               mod_bot_y, mod_bot_y - y_range * 0.06,
                               '#CC0000', h_size, h_gauge, string_gauge,
                               'positive', label_y=Y(F_MERGE) - y_range * 0.08,
                               y_range=y_range, device_x=_device_x,
                               hr_offset=0, device_position=device_position)

        if device_position == 'south':
            neg_merge_x = max(neg_xs)
        elif device_position == 'north':
            neg_merge_x = min(neg_xs)
        else:
            if harness_string_offset < motor_after:
                neg_merge_x = max(neg_xs)
            else:
                neg_merge_x = min(neg_xs)
        _draw_harness_routing(ax, neg_xs, neg_merge_x,
                               Y(F_ROUTE_TOP), Y(F_MERGE),
                               '#0000CC', h_size, h_gauge, string_gauge,
                               'negative', label_y=Y(F_MERGE) - y_range * 0.08,
                               y_range=y_range, device_x=_device_x,
                               hr_offset=0, device_position=device_position)

        harness_string_offset += h_size

    # --- Title ---
    circle_x = total_w * 0.02
    title_x = circle_x + 1.5  # close to circle
    title_y = Y(F_TITLE)
    template_name = spec.get('template_name', '')
    title_suffix = f' ({template_name})' if template_name else ''
    title_text = f'TYPICAL WIRING DIAGRAM: {spt}-STRING TRACKER{title_suffix}'

    ax.text(circle_x, title_y, letter,
            fontsize=10, fontweight='bold', color='black',
            ha='center', va='center', fontfamily='sans-serif',
            bbox=dict(boxstyle='circle,pad=0.3', facecolor='white',
                      edgecolor='black', linewidth=0.8))
    ax.text(title_x, title_y, title_text,
            fontsize=6, fontweight='bold', color='black',
            ha='left', va='center', fontfamily='sans-serif')
    # Underline
    from matplotlib.font_manager import FontProperties
    fp = FontProperties(size=6, weight='bold', family='sans-serif')
    r = ax.get_figure().canvas.get_renderer() if hasattr(ax.get_figure().canvas, 'get_renderer') else None
    # Approximate underline width from character count
    approx_w = len(title_text) * 0.42
    ax.plot([title_x, title_x + approx_w],
            [title_y - y_range * 0.012, title_y - y_range * 0.012],
            color='black', linewidth=0.6)
    ax.text(title_x, title_y - y_range * 0.04, 'Scale: NTS',
            fontsize=4, color='#555555', ha='left', va='center',
            fontfamily='sans-serif')


def _draw_harness_routing(ax, string_xs, merge_x, y_top, y_merge,
                           color, harness_size, harness_gauge, string_gauge,
                           polarity_label, label_y, y_range=100, device_x=None,
                           hr_offset=0, hr_base_y=None, device_position='south'):
    """Draw the collection routing for one polarity of one harness group."""
    n = len(string_xs)
    stub_len = (y_top - y_merge) * 0.3

    # Per-string inline fuses: positive multi-string harnesses only
    draw_fuses = (polarity_label == 'positive' and harness_size > 1)
    fuse_h = y_range * 0.015
    fuse_w = 0.8
    fuse_label_drawn = False

    for i, sx in enumerate(string_xs):
        ax.plot([sx, sx], [y_top, y_top - stub_len],
                color=color, linewidth=0.4, solid_capstyle='butt')

        if draw_fuses:
            fuse_top_y = y_top - stub_len
            fuse_bot_y = fuse_top_y - fuse_h
            ax.add_patch(Rectangle((sx - fuse_w / 2, fuse_bot_y), fuse_w, fuse_h,
                                    facecolor='white', edgecolor=color, linewidth=0.4))
            ax.plot([sx, sx], [fuse_bot_y, y_merge],
                    color=color, linewidth=0.3, alpha=0.6)
            if not fuse_label_drawn:
                ax.text(sx + fuse_w, fuse_bot_y + fuse_h / 2, 'INLINE FUSE (TYP)',
                        fontsize=2.5, color='#555555', ha='left', va='center',
                        fontfamily='sans-serif')
                fuse_label_drawn = True
        else:
            ax.plot([sx, sx], [y_top - stub_len, y_merge],
                    color=color, linewidth=0.3, alpha=0.6)

    # Horizontal collection bus at merge level
    if n > 1:
        ax.plot([min(string_xs), max(string_xs)],
                [y_merge, y_merge],
                color=color, linewidth=0.5)

    # Y-connector dot at merge point
    if n > 1:
        ax.plot(merge_x, y_merge, 'o', color=color, markersize=2.5,
                markeredgecolor='black', markeredgewidth=0.3)
        ax.text(merge_x + 1.5, y_merge, 'Y-CONNECTOR (TYP)',
                fontsize=2.5, color='#555555', ha='left', va='center',
                fontfamily='sans-serif')

    # Extender route from Y-connector — horizontal toward device
    hr_drop = y_range * 0.06
    dist_to_device = abs(device_x - merge_x) if device_x is not None else 0
    print(f"  [EXTENDER] {polarity_label} merge_x={merge_x:.2f} device_x={device_x} "
          f"dist={dist_to_device:.2f} device_position={device_position}")
    if device_position == 'middle' and device_x is not None and abs(device_x - merge_x) > 0.5:
        # Horizontal run toward motor gap center
        inset = 0.8 if merge_x < device_x else -0.8
        end_x = device_x - inset
        ax.plot([merge_x, end_x], [y_merge, y_merge],
                color=color, linewidth=0.5)
    elif device_x is not None and abs(device_x - merge_x) > 0.5:
        # Horizontal run toward device (north = left, south = right)
        inset = 0.8 if merge_x < device_x else -0.8
        end_x = device_x - inset
        ax.plot([merge_x, end_x], [y_merge, y_merge],
                color=color, linewidth=0.5)
    else:
        # Fallback: vertical drop
        ax.plot([merge_x, merge_x], [y_merge, y_merge - hr_drop],
                color=color, linewidth=0.5)

    # Bottom label — source circuit for 1-string, harness home run for multi-string
    if n == 1:
        bot_label = f'1-STRING SOURCE CIRCUIT\n#{string_gauge} CU PV WIRE (TYP)'
    else:
        bot_label = (f'{harness_size}-STRING HARNESS EXTENDER\n'
                     f'#{harness_gauge} CU PV WIRE (TYP)')
    ax.text(merge_x, label_y, bot_label,
            fontsize=3, color=color, ha='center', va='top',
            fontfamily='sans-serif', fontweight='bold', linespacing=1.4)