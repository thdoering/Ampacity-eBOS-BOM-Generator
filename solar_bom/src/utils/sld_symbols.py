"""
ANSI Electrical Symbol Library for Single Line Diagrams
"""

import tkinter as tk
from typing import Dict, List, Tuple, Optional, Any
from dataclasses import dataclass
import math


@dataclass
class SymbolDefinition:
    """Definition of a symbol's geometry and properties"""
    symbol_type: str
    default_width: float
    default_height: float
    default_color: str
    default_outline: str
    connection_points: Dict[str, Tuple[str, float]]  # port_name -> (side, offset)


class ANSISymbols:
    """ANSI standard electrical symbols for SLD"""
    
    # Color scheme
    COLORS = {
        'pv_array': {
            'fill': '#4A90E2',  # Blue
            'outline': '#2C5282',  # Dark blue
            'text': '#FFFFFF'
        },
        'inverter': {
            'fill': '#E24A4A',  # Red
            'outline': '#822C2C',  # Dark red
            'text': '#FFFFFF'
        },
        'combiner': {
            'fill': '#808080',  # Gray
            'outline': '#404040',  # Dark gray
            'text': '#FFFFFF'
        },
        'transformer': {
            'fill': '#90E24A',  # Green
            'outline': '#528C2C',  # Dark green
            'text': '#000000'
        },
        'meter': {
            'fill': '#E2E24A',  # Yellow
            'outline': '#828C2C',  # Dark yellow
            'text': '#000000'
        },
        'disconnect': {
            'fill': '#FFFFFF',  # White
            'outline': '#000000',  # Black
            'text': '#000000'
        },
        'breaker': {
            'fill': '#FFFFFF',  # White
            'outline': '#000000',  # Black
            'text': '#000000'
        }
    }
    
    # Symbol definitions with connection points
    SYMBOLS = {
        'pv_array': SymbolDefinition(
            symbol_type='pv_array',
            default_width=150,
            default_height=100,
            default_color='#4A90E2',
            default_outline='#2C5282',
            connection_points={
                'dc_positive': ('right', 0.3),
                'dc_negative': ('right', 0.7)
            }
        ),
        'inverter': SymbolDefinition(
            symbol_type='inverter',
            default_width=150,
            default_height=100,
            default_color='#E24A4A',
            default_outline='#822C2C',
            connection_points={
                'dc_positive_in': ('left', 0.3),
                'dc_negative_in': ('left', 0.7),
                'ac_l1': ('right', 0.25),
                'ac_l2': ('right', 0.5),
                'ac_l3': ('right', 0.75)
            }
        ),
        'combiner': SymbolDefinition(
            symbol_type='combiner',
            default_width=80,
            default_height=80,
            default_color='#808080',
            default_outline='#404040',
            connection_points={
                'input_1': ('left', 0.2),
                'input_2': ('left', 0.4),
                'input_3': ('left', 0.6),
                'input_4': ('left', 0.8),
                'output_positive': ('right', 0.35),
                'output_negative': ('right', 0.65)
            }
        )
    }
    
    @classmethod
    def draw_symbol(cls, canvas: tk.Canvas, symbol_type: str, x: float, y: float, 
                   width: Optional[float] = None, height: Optional[float] = None,
                   label: str = "", element_id: str = "", **kwargs) -> Dict[str, Any]:
        """
        Draw an ANSI symbol on the canvas
        
        Args:
            canvas: tkinter Canvas to draw on
            symbol_type: Type of symbol ('pv_array', 'inverter', 'combiner', etc.)
            x: X coordinate of top-left corner
            y: Y coordinate of top-left corner
            width: Optional width override
            height: Optional height override
            label: Text label for the symbol
            element_id: Unique identifier for the element
            **kwargs: Additional drawing options
            
        Returns:
            Dictionary with canvas item IDs and connection point coordinates
        """
        if symbol_type not in cls.SYMBOLS:
            raise ValueError(f"Unknown symbol type: {symbol_type}")
        
        symbol_def = cls.SYMBOLS[symbol_type]
        colors = cls.COLORS.get(symbol_type, cls.COLORS['pv_array'])
        
        # Use default dimensions if not specified
        width = width or symbol_def.default_width
        height = height or symbol_def.default_height
        
        # Get custom colors if provided
        fill_color = kwargs.get('fill', colors['fill'])
        outline_color = kwargs.get('outline', colors['outline'])
        text_color = kwargs.get('text_color', colors['text'])
        outline_width = kwargs.get('outline_width', 2)
        
        # Tags for the element
        tags = ['element', symbol_type, element_id] if element_id else ['element', symbol_type]
        
        # Dictionary to store canvas item IDs
        items = {
            'shapes': [],
            'text': [],
            'ports': [],
            'all': []
        }
        
        # Draw based on symbol type
        if symbol_type == 'pv_array':
            items['shapes'].extend(cls._draw_pv_array(
                canvas, x, y, width, height, fill_color, outline_color, outline_width, tags
            ))
        elif symbol_type == 'inverter':
            items['shapes'].extend(cls._draw_inverter(
                canvas, x, y, width, height, fill_color, outline_color, outline_width, tags
            ))
        elif symbol_type == 'combiner':
            items['shapes'].extend(cls._draw_combiner(
                canvas, x, y, width, height, fill_color, outline_color, outline_width, tags
            ))
        elif symbol_type == 'transformer':
            items['shapes'].extend(cls._draw_transformer(
                canvas, x, y, width, height, fill_color, outline_color, outline_width, tags
            ))
        else:
            # Default rectangle for unknown types
            rect_id = canvas.create_rectangle(
                x, y, x + width, y + height,
                fill=fill_color,
                outline=outline_color,
                width=outline_width,
                tags=tags
            )
            items['shapes'].append(rect_id)
        
        # Add label if provided
        if label:
            label_tags = tags.copy()
            label_tags[0] = 'label'  # Replace 'element' with 'label'
            if element_id:
                label_tags.append(f"{element_id}_label")
            
            text_id = canvas.create_text(
                x + width / 2,
                y + height / 2,
                text=label,
                fill=text_color,
                font=('Arial', 10, 'bold'),
                tags=label_tags
            )
            items['text'].append(text_id)
        
        # Calculate connection point coordinates
        connection_points = {}
        for port_name, (side, offset) in symbol_def.connection_points.items():
            port_x, port_y = cls._calculate_port_position(x, y, width, height, side, offset)
            connection_points[port_name] = (port_x, port_y)
            
            # Optionally draw connection points (for debugging)
            if kwargs.get('show_ports', False):
                port_id = cls._draw_connection_port(
                    canvas, port_x, port_y, port_name, tags
                )
                items['ports'].append(port_id)
        
        # Collect all items
        items['all'] = items['shapes'] + items['text'] + items['ports']
        
        return {
            'items': items,
            'connection_points': connection_points,
            'bounds': (x, y, x + width, y + height)
        }
    
    @classmethod
    def _draw_pv_array(cls, canvas: tk.Canvas, x: float, y: float, 
                       width: float, height: float, fill: str, 
                       outline: str, outline_width: int, tags: List[str]) -> List[int]:
        """Draw PV array symbol (rectangle with diagonal lines)"""
        items = []
        
        # Main rectangle
        rect_id = canvas.create_rectangle(
            x, y, x + width, y + height,
            fill=fill,
            outline=outline,
            width=outline_width,
            tags=tags
        )
        items.append(rect_id)
        
        # Diagonal lines to represent solar cells
        line_spacing = 15
        line_tags = tags.copy()
        line_tags.append('pv_lines')
        
        # Calculate diagonal lines
        for i in range(int((width + height) / line_spacing)):
            line_x = x + i * line_spacing
            if line_x < x + width:
                # Start from top edge
                start_x = line_x
                start_y = y
                end_x = max(x, line_x - height)
                end_y = min(y + height, y + (line_x - x))
            else:
                # Start from right edge
                overflow = line_x - (x + width)
                start_x = x + width
                start_y = y + overflow
                end_x = max(x, x + width - height + overflow)
                end_y = y + height
            
            if start_y <= y + height and end_y >= y:
                line_id = canvas.create_line(
                    start_x, start_y, end_x, end_y,
                    fill=outline,
                    width=1,
                    tags=line_tags
                )
                items.append(line_id)
        
        # Add DC label
        dc_text_id = canvas.create_text(
            x + width - 20, y + 15,
            text="DC",
            fill='white',
            font=('Arial', 8),
            tags=tags
        )
        items.append(dc_text_id)
        
        return items
    
    @classmethod
    def _draw_inverter(cls, canvas: tk.Canvas, x: float, y: float, 
                      width: float, height: float, fill: str, 
                      outline: str, outline_width: int, tags: List[str]) -> List[int]:
        """Draw inverter symbol (rectangle with AC/DC notation and wave symbol)"""
        items = []
        
        # Main rectangle
        rect_id = canvas.create_rectangle(
            x, y, x + width, y + height,
            fill=fill,
            outline=outline,
            width=outline_width,
            tags=tags
        )
        items.append(rect_id)
        
        # Vertical divider line
        divider_x = x + width * 0.5
        divider_id = canvas.create_line(
            divider_x, y + 10, divider_x, y + height - 10,
            fill='white',
            width=2,
            tags=tags
        )
        items.append(divider_id)
        
        # DC label on left side
        dc_text_id = canvas.create_text(
            x + width * 0.25, y + 15,
            text="DC",
            fill='white',
            font=('Arial', 10, 'bold'),
            tags=tags
        )
        items.append(dc_text_id)
        
        # AC label on right side
        ac_text_id = canvas.create_text(
            x + width * 0.75, y + 15,
            text="AC",
            fill='white',
            font=('Arial', 10, 'bold'),
            tags=tags
        )
        items.append(ac_text_id)
        
        # Sine wave on AC side
        wave_points = []
        wave_start_x = x + width * 0.6
        wave_end_x = x + width * 0.9
        wave_y = y + height * 0.5
        wave_amplitude = height * 0.15
        
        for i in range(21):
            t = i / 20
            wave_x = wave_start_x + (wave_end_x - wave_start_x) * t
            wave_offset = wave_amplitude * math.sin(t * 2 * math.pi)
            wave_points.extend([wave_x, wave_y + wave_offset])
        
        if len(wave_points) >= 4:
            wave_id = canvas.create_line(
                *wave_points,
                fill='white',
                width=2,
                smooth=True,
                tags=tags
            )
            items.append(wave_id)
        
        return items
    
    @classmethod
    def _draw_combiner(cls, canvas: tk.Canvas, x: float, y: float, 
                      width: float, height: float, fill: str, 
                      outline: str, outline_width: int, tags: List[str]) -> List[int]:
        """Draw combiner box symbol (square with junction symbol)"""
        items = []
        
        # Main square/rectangle
        rect_id = canvas.create_rectangle(
            x, y, x + width, y + height,
            fill=fill,
            outline=outline,
            width=outline_width,
            tags=tags
        )
        items.append(rect_id)
        
        # Draw junction symbol (multiple inputs to single output)
        center_x = x + width / 2
        center_y = y + height / 2
        
        # Input lines on left
        for i in range(3):
            line_y = y + height * (0.25 + i * 0.25)
            line_id = canvas.create_line(
                x + 10, line_y,
                center_x - 10, center_y,
                fill='white',
                width=2,
                tags=tags
            )
            items.append(line_id)
        
        # Central junction point
        junction_id = canvas.create_oval(
            center_x - 5, center_y - 5,
            center_x + 5, center_y + 5,
            fill='white',
            outline='white',
            tags=tags
        )
        items.append(junction_id)
        
        # Output line on right
        output_id = canvas.create_line(
            center_x + 10, center_y,
            x + width - 10, center_y,
            fill='white',
            width=3,
            tags=tags
        )
        items.append(output_id)
        
        # Add "CB" label
        cb_text_id = canvas.create_text(
            center_x, y + height - 15,
            text="CB",
            fill='white',
            font=('Arial', 8, 'bold'),
            tags=tags
        )
        items.append(cb_text_id)
        
        return items
    
    @classmethod
    def _draw_transformer(cls, canvas: tk.Canvas, x: float, y: float, 
                         width: float, height: float, fill: str, 
                         outline: str, outline_width: int, tags: List[str]) -> List[int]:
        """Draw transformer symbol (two circles/coils)"""
        items = []
        
        # Draw two circles representing transformer coils
        coil_radius = min(width, height) * 0.3
        center_y = y + height / 2
        
        # Primary coil (left)
        left_center = x + width * 0.35
        left_coil_id = canvas.create_oval(
            left_center - coil_radius, center_y - coil_radius,
            left_center + coil_radius, center_y + coil_radius,
            fill='',
            outline=outline,
            width=outline_width,
            tags=tags
        )
        items.append(left_coil_id)
        
        # Secondary coil (right)
        right_center = x + width * 0.65
        right_coil_id = canvas.create_oval(
            right_center - coil_radius, center_y - coil_radius,
            right_center + coil_radius, center_y + coil_radius,
            fill='',
            outline=outline,
            width=outline_width,
            tags=tags
        )
        items.append(right_coil_id)
        
        # Core lines (vertical lines between coils)
        for i in range(2):
            line_x = x + width * (0.45 + i * 0.1)
            line_id = canvas.create_line(
                line_x, center_y - coil_radius,
                line_x, center_y + coil_radius,
                fill=outline,
                width=outline_width,
                tags=tags
            )
            items.append(line_id)
        
        return items
    
    @classmethod
    def _draw_connection_port(cls, canvas: tk.Canvas, x: float, y: float, 
                             port_name: str, tags: List[str]) -> int:
        """Draw a connection port indicator (small circle)"""
        port_radius = 4
        port_tags = tags.copy()
        port_tags.append('port')
        port_tags.append(f"port_{port_name}")
        
        # Determine color based on port type
        if 'positive' in port_name or 'l1' in port_name:
            color = '#FF0000'  # Red for positive/L1
        elif 'negative' in port_name or 'neutral' in port_name:
            color = '#0000FF'  # Blue for negative/neutral
        elif 'l2' in port_name:
            color = '#00FF00'  # Green for L2
        elif 'l3' in port_name:
            color = '#FFA500'  # Orange for L3
        else:
            color = '#808080'  # Gray for others
        
        port_id = canvas.create_oval(
            x - port_radius, y - port_radius,
            x + port_radius, y + port_radius,
            fill=color,
            outline='black',
            width=1,
            tags=port_tags
        )
        
        return port_id
    
    @classmethod
    def _calculate_port_position(cls, x: float, y: float, width: float, height: float,
                                side: str, offset: float) -> Tuple[float, float]:
        """
        Calculate the absolute position of a connection port
        
        Args:
            x, y: Top-left corner of the symbol
            width, height: Dimensions of the symbol
            side: Which side ('top', 'bottom', 'left', 'right')
            offset: Position along the side (0.0 to 1.0)
            
        Returns:
            Tuple of (port_x, port_y)
        """
        if side == 'top':
            port_x = x + width * offset
            port_y = y
        elif side == 'bottom':
            port_x = x + width * offset
            port_y = y + height
        elif side == 'left':
            port_x = x
            port_y = y + height * offset
        elif side == 'right':
            port_x = x + width
            port_y = y + height * offset
        else:
            # Default to center
            port_x = x + width / 2
            port_y = y + height / 2
        
        return (port_x, port_y)
    
    @classmethod
    def get_symbol_info(cls, symbol_type: str) -> Optional[SymbolDefinition]:
        """Get information about a symbol type"""
        return cls.SYMBOLS.get(symbol_type)
    
    @classmethod
    def get_available_symbols(cls) -> List[str]:
        """Get list of available symbol types"""
        return list(cls.SYMBOLS.keys())
    
    @classmethod
    def get_symbol_color(cls, symbol_type: str, color_type: str = 'fill') -> str:
        """
        Get color for a symbol type
        
        Args:
            symbol_type: Type of symbol
            color_type: 'fill', 'outline', or 'text'
            
        Returns:
            Color hex string
        """
        colors = cls.COLORS.get(symbol_type, cls.COLORS['pv_array'])
        return colors.get(color_type, '#000000')
    
    @classmethod
    def draw_string_symbol(cls, canvas: tk.Canvas, x: float, y: float, 
                          num_modules: int, string_label: str = "",
                          element_id: str = "", **kwargs) -> Dict[str, Any]:
        """
        Draw a string of PV modules connected in series
        """
        # Module dimensions (small rectangles)
        module_width = 8
        module_height = 15
        module_spacing = 2
        
        # Calculate how many modules to show (max 6 for space)
        modules_to_show = min(num_modules, 6)
        
        # Tags for the element - element_id should be in all tags
        tags = [element_id, 'element', 'string'] if element_id else ['element', 'string']
        
        items = {
            'shapes': [],
            'text': [],
            'all': []
        }
        
        # Draw modules
        current_x = x
        for i in range(modules_to_show):
            # Draw module rectangle
            rect_id = canvas.create_rectangle(
                current_x, y,
                current_x + module_width, y + module_height,
                fill='#4A90E2',
                outline='#2C5282',
                width=1,
                tags=tags
            )
            items['shapes'].append(rect_id)
            
            # Draw connection line to next module (except for last)
            if i < modules_to_show - 1:
                line_id = canvas.create_line(
                    current_x + module_width, y + module_height/2,
                    current_x + module_width + module_spacing, y + module_height/2,
                    fill='#2C5282',
                    width=2,
                    tags=tags
                )
                items['shapes'].append(line_id)
            
            current_x += module_width + module_spacing
        
        # If there are more modules than shown, add "..." indicator
        if num_modules > modules_to_show:
            dots_id = canvas.create_text(
                current_x + 5, y + module_height/2,
                text="...",
                fill='#2C5282',
                font=('Arial', 10, 'bold'),
                tags=tags
            )
            items['text'].append(dots_id)
            current_x += 20
            
            # Draw last module
            rect_id = canvas.create_rectangle(
                current_x, y,
                current_x + module_width, y + module_height,
                fill='#4A90E2',
                outline='#2C5282',
                width=1,
                tags=tags
            )
            items['shapes'].append(rect_id)
            current_x += module_width
        
        # Calculate total width
        total_width = current_x - x
        
        # NO VERTICAL TEXT - Only draw a simple label below if provided
        total_height = module_height + 5
        
        # Draw positive and negative terminals
        # Positive (left side)
        pos_x = x - 10
        pos_y = y + module_height/2
        pos_circle = canvas.create_oval(
            pos_x - 3, pos_y - 3,
            pos_x + 3, pos_y + 3,
            fill='red',
            outline='darkred',
            tags=tags
        )
        items['shapes'].append(pos_circle)
        
        pos_line = canvas.create_line(
            pos_x + 3, pos_y,
            x, pos_y,
            fill='red',
            width=2,
            tags=tags
        )
        items['shapes'].append(pos_line)
        
        # Add + symbol
        plus_id = canvas.create_text(
            pos_x, pos_y - 8,
            text="+",
            fill='red',
            font=('Arial', 8, 'bold'),
            tags=tags
        )
        items['text'].append(plus_id)
        
        # Negative (right side)
        neg_x = current_x + 10
        neg_y = y + module_height/2
        neg_circle = canvas.create_oval(
            neg_x - 3, neg_y - 3,
            neg_x + 3, neg_y + 3,
            fill='blue',
            outline='darkblue',
            tags=tags
        )
        items['shapes'].append(neg_circle)
        
        neg_line = canvas.create_line(
            current_x, neg_y,
            neg_x - 3, neg_y,
            fill='blue',
            width=2,
            tags=tags
        )
        items['shapes'].append(neg_line)
        
        # Add - symbol
        minus_id = canvas.create_text(
            neg_x, neg_y - 8,
            text="−",
            fill='blue',
            font=('Arial', 10, 'bold'),
            tags=tags
        )
        items['text'].append(minus_id)
        
        # Collect all items
        items['all'] = items['shapes'] + items['text']
        
        # Connection points
        connection_points = {
            'dc_positive': (pos_x - 5, pos_y),
            'dc_negative': (neg_x + 5, neg_y)
        }
        
        return {
            'items': items,
            'connection_points': connection_points,
            'bounds': (pos_x - 10, y - 10, neg_x + 10, y + total_height)
        }