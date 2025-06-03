import json
import os
from PIL import Image, ImageDraw, ImageFont
from typing import Dict, Any, Optional
from pathlib import Path

class HarnessDrawingGenerator:
    """Generator for harness technical drawings"""
    
    def __init__(self, harness_library_path: str = 'data/harness_library.json'):
        """Initialize the generator with harness library"""
        self.harness_library_path = harness_library_path
        self.harness_library = self.load_harness_library()
        
        # Drawing constants
        self.canvas_width = 1200
        self.canvas_height = 900
        self.margin = 80
        self.drawing_area_width = self.canvas_width - (2 * self.margin)
        self.drawing_area_height = 600
        
        # Colors
        self.bg_color = (255, 255, 255)  # White
        self.line_color = (0, 0, 0)      # Black
        self.pos_color = (255, 0, 0)     # Red for positive
        self.neg_color = (0, 0, 255)     # Blue for negative
        self.dim_color = (100, 100, 100) # Gray for dimensions
        
        # Try to load fonts, fall back to default if not available
        try:
            self.title_font = ImageFont.truetype("arial.ttf", 16)
            self.label_font = ImageFont.truetype("arial.ttf", 12)
            self.dim_font = ImageFont.truetype("arial.ttf", 10)
            self.table_font = ImageFont.truetype("arial.ttf", 9)
        except OSError:
            # Fall back to default font if Arial not available
            self.title_font = ImageFont.load_default()
            self.label_font = ImageFont.load_default()
            self.dim_font = ImageFont.load_default()
            self.table_font = ImageFont.load_default()
    
    def load_harness_library(self) -> Dict[str, Any]:
        """Load the harness library from JSON file"""
        try:
            with open(self.harness_library_path, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"Warning: Harness library file not found at {self.harness_library_path}")
            return {}
        except json.JSONDecodeError:
            print(f"Warning: Invalid JSON in harness library file {self.harness_library_path}")
            return {}
    
    def generate_harness_drawing(self, part_number: str, output_dir: str = 'harness_drawings') -> bool:
        """Generate a technical drawing for a specific harness part number"""
        if part_number not in self.harness_library:
            print(f"Error: Part number {part_number} not found in harness library")
            return False
        
        harness_spec = self.harness_library[part_number]
        
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        
        # Create canvas
        image = Image.new('RGB', (self.canvas_width, self.canvas_height), self.bg_color)
        draw = ImageDraw.Draw(image)
        
        # Draw the harness diagram
        self.draw_harness_diagram(draw, harness_spec)
        
        # Draw the specifications table
        self.draw_specifications_table(draw, harness_spec)
        
        # Draw title and part number
        self.draw_title_block(draw, harness_spec)
        
        # Save the image
        filename = f"{part_number}_harness_drawing.png"
        filepath = os.path.join(output_dir, filename)
        
        try:
            image.save(filepath, 'PNG', dpi=(300, 300))
            print(f"Generated harness drawing: {filepath}")
            return True
        except Exception as e:
            print(f"Error saving harness drawing: {str(e)}")
            return False
    
    def draw_harness_diagram(self, draw: ImageDraw, harness_spec: Dict[str, Any]):
        """Draw the main harness diagram with trunk and string drops"""
        num_strings = harness_spec['num_strings']
        spacing_ft = harness_spec['string_spacing_ft']
        polarity = harness_spec['polarity']
        
        # Calculate positions
        diagram_start_y = 150
        trunk_y = diagram_start_y + 100
        
        # Calculate total width needed for the harness
        total_harness_width = (num_strings - 1) * 60  # 60 pixels per string spacing
        if total_harness_width > self.drawing_area_width - 200:
            # Scale down if too wide
            string_spacing_px = (self.drawing_area_width - 200) // (num_strings - 1)
            total_harness_width = (num_strings - 1) * string_spacing_px
        else:
            string_spacing_px = 60
        
        # Center the harness horizontally
        harness_start_x = self.margin + (self.drawing_area_width - total_harness_width) // 2
        harness_end_x = harness_start_x + total_harness_width
        
        # Draw main trunk line
        trunk_color = self.pos_color if polarity == 'positive' else self.neg_color
        trunk_thickness = 4
        
        draw.line([harness_start_x, trunk_y, harness_end_x, trunk_y], 
                 fill=trunk_color, width=trunk_thickness)
        
        # Draw string drops
        drop_length = 80
        drop_thickness = 2
        
        for i in range(num_strings):
            drop_x = harness_start_x + (i * string_spacing_px)
            drop_start_y = trunk_y
            drop_end_y = trunk_y + drop_length
            
            # Draw vertical drop line
            draw.line([drop_x, drop_start_y, drop_x, drop_end_y], 
                     fill=trunk_color, width=drop_thickness)
            
            # Draw connector symbol at end of drop
            connector_size = 8
            draw.rectangle([drop_x - connector_size//2, drop_end_y - connector_size//2,
                          drop_x + connector_size//2, drop_end_y + connector_size//2],
                         outline=trunk_color, width=2)
            
            # Label each string
            string_label = f"S{i+1}"
            bbox = draw.textbbox((0, 0), string_label, font=self.label_font)
            text_width = bbox[2] - bbox[0]
            draw.text((drop_x - text_width//2, drop_end_y + 15), string_label, 
                     fill=self.line_color, font=self.label_font)
        
        # Draw trunk connectors at ends
        connector_length = 30
        connector_thickness = 4
        
        # Left trunk connector
        draw.line([harness_start_x - connector_length, trunk_y, harness_start_x, trunk_y], 
                 fill=trunk_color, width=connector_thickness)
        
        # Right trunk connector  
        draw.line([harness_end_x, trunk_y, harness_end_x + connector_length, trunk_y], 
                 fill=trunk_color, width=connector_thickness)
        
        # Draw spacing dimensions
        self.draw_spacing_dimensions(draw, harness_start_x, trunk_y, string_spacing_px, 
                                   num_strings, spacing_ft)
        
        # Draw wire gauge labels
        self.draw_wire_gauge_labels(draw, harness_spec, harness_start_x, harness_end_x, trunk_y)
        
        # Draw fuse indicators if fused
        if harness_spec.get('fused', False):
            self.draw_fuse_indicators(draw, harness_start_x, trunk_y, string_spacing_px, 
                                    num_strings, harness_spec.get('fuse_rating', ''))
    
    def draw_spacing_dimensions(self, draw: ImageDraw, start_x: int, trunk_y: int, 
                              spacing_px: int, num_strings: int, spacing_ft: float):
        """Draw dimension lines showing string spacing"""
        if num_strings < 2:
            return
            
        dim_y = trunk_y - 40
        
        # Draw dimension line for first spacing
        dim_start_x = start_x
        dim_end_x = start_x + spacing_px
        
        # Dimension line
        draw.line([dim_start_x, dim_y, dim_end_x, dim_y], fill=self.dim_color, width=1)
        
        # Tick marks
        tick_height = 8
        draw.line([dim_start_x, dim_y - tick_height//2, dim_start_x, dim_y + tick_height//2], 
                 fill=self.dim_color, width=1)
        draw.line([dim_end_x, dim_y - tick_height//2, dim_end_x, dim_y + tick_height//2], 
                 fill=self.dim_color, width=1)
        
        # Dimension text
        dim_text = f"{spacing_ft}'"
        bbox = draw.textbbox((0, 0), dim_text, font=self.dim_font)
        text_width = bbox[2] - bbox[0]
        text_x = dim_start_x + (spacing_px - text_width) // 2
        draw.text((text_x, dim_y - 25), dim_text, fill=self.dim_color, font=self.dim_font)
        
        # Add "TYP" notation if more than 2 strings
        if num_strings > 2:
            draw.text((text_x + text_width + 5, dim_y - 25), "TYP", 
                     fill=self.dim_color, font=self.dim_font)
    
    def draw_wire_gauge_labels(self, draw: ImageDraw, harness_spec: Dict[str, Any], 
                             start_x: int, end_x: int, trunk_y: int):
        """Draw wire gauge labels on trunk and drops"""
        drop_gauge = harness_spec['drop_wire_gauge']
        trunk_gauge = harness_spec['trunk_wire_gauge']
        
        # Label trunk wire gauge
        trunk_center_x = (start_x + end_x) // 2
        bbox = draw.textbbox((0, 0), trunk_gauge, font=self.label_font)
        text_width = bbox[2] - bbox[0]
        draw.text((trunk_center_x - text_width//2, trunk_y + 15), trunk_gauge, 
                 fill=self.line_color, font=self.label_font)
        
        # Label drop wire gauge (on first drop)
        drop_x = start_x
        bbox = draw.textbbox((0, 0), drop_gauge, font=self.label_font)
        text_width = bbox[2] - bbox[0]
        draw.text((drop_x - text_width - 10, trunk_y + 40), drop_gauge, 
                 fill=self.line_color, font=self.label_font)
    
    def draw_fuse_indicators(self, draw: ImageDraw, start_x: int, trunk_y: int, 
                           spacing_px: int, num_strings: int, fuse_rating: str):
        """Draw fuse symbols on fused harnesses"""
        fuse_size = 12
        
        for i in range(num_strings):
            fuse_x = start_x + (i * spacing_px)
            fuse_y = trunk_y - 30
            
            # Draw fuse symbol (rectangle with line through it)
            draw.rectangle([fuse_x - fuse_size//2, fuse_y - fuse_size//2,
                          fuse_x + fuse_size//2, fuse_y + fuse_size//2],
                         outline=self.line_color, width=2)
            
            # Draw line connecting fuse to drop
            draw.line([fuse_x, fuse_y + fuse_size//2, fuse_x, trunk_y], 
                     fill=self.line_color, width=1)
            
            # Label first fuse with rating
            if i == 0 and fuse_rating:
                bbox = draw.textbbox((0, 0), fuse_rating, font=self.dim_font)
                text_width = bbox[2] - bbox[0]
                draw.text((fuse_x - text_width//2, fuse_y - fuse_size//2 - 20), 
                         fuse_rating, fill=self.line_color, font=self.dim_font)
    
    def draw_specifications_table(self, draw: ImageDraw, harness_spec: Dict[str, Any]):
        """Draw specifications table"""
        table_start_x = self.margin
        table_start_y = self.canvas_height - 200
        table_width = self.canvas_width - (2 * self.margin)
        
        # Table title
        draw.text((table_start_x, table_start_y - 30), "SPECIFICATIONS", 
                 fill=self.line_color, font=self.title_font)
        
        # Table data
        specs = [
            ("Part Number:", harness_spec['part_number']),
            ("ATPI Part Number:", harness_spec['atpi_part_number']),
            ("Description:", harness_spec['description']),
            ("Number of Strings:", str(harness_spec['num_strings'])),
            ("Polarity:", harness_spec['polarity'].title()),
            ("String Spacing:", f"{harness_spec['string_spacing_ft']}'"),
            ("Drop Wire Gauge:", harness_spec['drop_wire_gauge']),
            ("Trunk Wire Gauge:", harness_spec['trunk_wire_gauge']),
            ("Connector Type:", harness_spec['connector_type']),
        ]
        
        if harness_spec.get('fused', False):
            specs.append(("Fuse Rating:", harness_spec.get('fuse_rating', 'N/A')))
        
        # Draw table rows
        row_height = 20
        col1_width = 150
        
        for i, (label, value) in enumerate(specs):
            y_pos = table_start_y + (i * row_height)
            
            # Draw row background (alternating)
            if i % 2 == 0:
                draw.rectangle([table_start_x, y_pos - 2, 
                              table_start_x + table_width, y_pos + row_height - 2],
                             fill=(245, 245, 245))
            
            # Draw text
            draw.text((table_start_x + 10, y_pos), label, 
                     fill=self.line_color, font=self.table_font)
            draw.text((table_start_x + col1_width, y_pos), value, 
                     fill=self.line_color, font=self.table_font)
        
        # Draw table border
        table_height = len(specs) * row_height
        draw.rectangle([table_start_x, table_start_y - 2, 
                       table_start_x + table_width, table_start_y + table_height - 2],
                     outline=self.line_color, width=1)
    
    def draw_title_block(self, draw: ImageDraw, harness_spec: Dict[str, Any]):
        """Draw title block with part number and polarity"""
        title = f"HARNESS ASSEMBLY - {harness_spec['part_number']}"
        polarity_text = f"({harness_spec['polarity'].upper()})"
        
        # Main title
        bbox = draw.textbbox((0, 0), title, font=self.title_font)
        title_width = bbox[2] - bbox[0]
        title_x = (self.canvas_width - title_width) // 2
        draw.text((title_x, 30), title, fill=self.line_color, font=self.title_font)
        
        # Polarity indicator
        polarity_color = self.pos_color if harness_spec['polarity'] == 'positive' else self.neg_color
        bbox = draw.textbbox((0, 0), polarity_text, font=self.label_font)
        polarity_width = bbox[2] - bbox[0]
        polarity_x = (self.canvas_width - polarity_width) // 2
        draw.text((polarity_x, 55), polarity_text, fill=polarity_color, font=self.label_font)
    
    def generate_all_harness_drawings(self, output_dir: str = 'harness_drawings') -> int:
        """Generate drawings for all harnesses in the library"""
        success_count = 0
        
        for part_number in self.harness_library.keys():
            if self.generate_harness_drawing(part_number, output_dir):
                success_count += 1
        
        print(f"Generated {success_count} harness drawings in {output_dir}/")
        return success_count
    
    def get_available_harnesses(self) -> Dict[str, str]:
        """Get list of available harnesses with descriptions"""
        return {part_num: spec['description'] 
                for part_num, spec in self.harness_library.items()}