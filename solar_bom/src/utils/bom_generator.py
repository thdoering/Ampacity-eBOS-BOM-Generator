import pandas as pd
import json
import os
from typing import Dict, List, Any, Optional
from ..models.block import BlockConfig, WiringType
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from ..models.device import HarnessConnection, CombinerBoxConfig


class BOMGenerator:
    """Utility class for generating Bill of Materials from block configurations"""
    
    # Constants for BOM calculations
    CABLE_WASTE_FACTOR = 1.05  # 5% extra for waste/installation
    
    def __init__(self, blocks: Dict[str, BlockConfig], project=None):
        """
        Initialize BOM Generator
        
        Args:
            blocks: Dictionary of block configurations
            project: Project object (optional)
        """
        self.blocks = blocks
        self.project = project
        # Load library json files
        self.harness_library = self.load_harness_library()
        self.fuse_library = self.load_fuse_library()
        self.extender_library = self.load_extender_library()
        self.whip_library = self.load_whip_library()
    
    def calculate_cable_quantities(self) -> Dict[str, Dict[str, Any]]:
        """
        Calculate cable quantities by block and type, separating positive and negative
        
        Returns:
            Dictionary with quantities by block and component type
        """
        quantities = {}
        
        for block_id, block in self.blocks.items():
            
            block_quantities = {}
            
            # Get tracker counts categorized by strings per tracker
            tracker_counts = {}
            for pos in block.tracker_positions:
                if pos.template:
                    strings_per_tracker = pos.template.strings_per_tracker
                    key = f"{strings_per_tracker}-String Tracker"
                    if key not in tracker_counts:
                        tracker_counts[key] = 0
                    tracker_counts[key] += 1
            
            # Add trackers to quantities (as structural components)
            for tracker_type, count in tracker_counts.items():
                block_quantities[tracker_type] = {
                    'description': f"{tracker_type}",
                    'quantity': count,
                    'unit': 'units',
                    'category': 'Structural'
                }
            
            # If block has no wiring config, add a note and store the block with just trackers
            if not block.wiring_config:
                block_quantities["No Wiring Configuration"] = {
                    'description': f"Block requires wiring configuration",
                    'quantity': 1,
                    'unit': 'warning',
                    'category': 'Warnings'
                }
                quantities[block_id] = block_quantities
                continue
                
            cable_lengths = block.calculate_cable_lengths()
            
            if block.wiring_config.wiring_type == WiringType.HOMERUN:
                # For homerun, we track string cable length - split by polarity
                if 'string_cable_positive' in cable_lengths:
                    length = cable_lengths['string_cable_positive'] * self.CABLE_WASTE_FACTOR
                    # Convert from meters to feet (1m = 3.28084ft)
                    length_feet = round(length * 3.28084, 1)
                    
                    # Get the cable size for more specific description
                    cable_size = block.wiring_config.string_cable_size
                    
                    block_quantities[f'Positive String Wire ({cable_size})'] = {
                        'description': f'DC Positive String Wire {cable_size}',
                        'quantity': length_feet,
                        'unit': 'feet',
                        'category': 'eBOS'
                    }
                
                if 'string_cable_negative' in cable_lengths:
                    length = cable_lengths['string_cable_negative'] * self.CABLE_WASTE_FACTOR
                    # Convert from meters to feet (1m = 3.28084ft)
                    length_feet = round(length * 3.28084, 1)
                    
                    cable_size = block.wiring_config.string_cable_size
                    
                    block_quantities[f'Negative String Wire ({cable_size})'] = {
                        'description': f'DC Negative String Wire {cable_size}',
                        'quantity': length_feet,
                        'unit': 'feet',
                        'category': 'eBOS'
                    }
                    
                # Get the whip cable size
                whip_cable_size = getattr(block.wiring_config, 'whip_cable_size', '6 AWG')  # Default to 6 AWG if not specified

                # Split whip cable by polarity
                if 'whip_cable_positive' in cable_lengths:
                    whip_length_feet = round(cable_lengths['whip_cable_positive'] * 3.28084 * self.CABLE_WASTE_FACTOR, 1)

                    block_quantities[f'Positive Whip Cable ({whip_cable_size})'] = {
                        'description': self.get_whip_description_format(whip_cable_size).format(length="TOTAL"),
                        'quantity': whip_length_feet,
                        'unit': 'feet',
                        'category': 'eBOS'
                    }
                
                if 'whip_cable_negative' in cable_lengths:
                    whip_length_feet = round(cable_lengths['whip_cable_negative'] * 3.28084 * self.CABLE_WASTE_FACTOR, 1)

                    block_quantities[f'Negative Whip Cable ({whip_cable_size})'] = {
                        'description': self.get_whip_description_format(whip_cable_size).format(length="TOTAL"),
                        'quantity': whip_length_feet,
                        'unit': 'feet',
                        'category': 'eBOS'
                    }
                    
            else:  # HARNESS configuration
                # Add harnesses by number of strings they connect - split by polarity
                harness_counts = self._count_harnesses_by_size(block)
                
                # Add each harness type, separately for positive and negative
                for string_count, count in harness_counts.items():
                    harness_cable_size = block.wiring_config.harness_cable_size
                    string_cable_size = block.wiring_config.string_cable_size
                    
                    # Calculate string spacing for harness matching
                    if block.tracker_template and block.tracker_template.module_spec:
                        module_spec = block.tracker_template.module_spec
                        modules_per_string = block.tracker_template.modules_per_string
                        module_spacing_m = block.tracker_template.module_spacing_m
                        
                        string_spacing_ft = self.calculate_string_spacing_ft(
                            modules_per_string, module_spec.width_mm, module_spacing_m
                        )
                    else:
                        string_spacing_ft = 102.0  # Default spacing
                        
                    # Get descriptions from harness library
                    pos_description = self.get_harness_description(
                        string_count, 'positive', string_spacing_ft, harness_cable_size, string_cable_size
                    )
                    neg_description = self.get_harness_description(
                        string_count, 'negative', string_spacing_ft, harness_cable_size, string_cable_size
                    )

                    block_quantities[f'Positive {string_count}-String Harness'] = {
                        'description': pos_description,
                        'quantity': count,
                        'unit': 'units',
                        'category': 'eBOS'
                    }

                    block_quantities[f'Negative {string_count}-String Harness'] = {
                        'description': neg_description,
                        'quantity': count,
                        'unit': 'units',
                        'category': 'eBOS'
                    }
                
                # Count fuses by rating
                fuse_counts = self._count_fuses_by_rating(block)
                for rating, count in fuse_counts.items():
                    block_quantities[f'DC String Fuse {rating}A'] = {
                        'description': self.get_fuse_description(rating),
                        'quantity': count,
                        'unit': 'units',
                        'category': 'eBOS'
                    }
                
                # Split whip cables by polarity
                whip_cable_size = getattr(block.wiring_config, 'whip_cable_size', '6 AWG')  # Default to 6 AWG if not specified
                
                if 'whip_cable_positive' in cable_lengths:
                    whip_length_feet = round(cable_lengths['whip_cable_positive'] * 3.28084 * self.CABLE_WASTE_FACTOR, 1)

                    block_quantities[f'Positive Whip Cable ({whip_cable_size})'] = {
                        'description': self.get_whip_description_format(whip_cable_size).format(length="TOTAL"),
                        'quantity': whip_length_feet,
                        'unit': 'feet',
                        'category': 'eBOS'
                    }
                
                if 'whip_cable_negative' in cable_lengths:
                    whip_length_feet = round(cable_lengths['whip_cable_negative'] * 3.28084 * self.CABLE_WASTE_FACTOR, 1)

                    block_quantities[f'Negative Whip Cable ({whip_cable_size})'] = {
                        'description': self.get_whip_description_format(whip_cable_size).format(length="TOTAL"),
                        'quantity': whip_length_feet,
                        'unit': 'feet',
                        'category': 'eBOS'
                    }

                # Split extender cable by polarity
                extender_cable_size = getattr(block.wiring_config, 'extender_cable_size', '8 AWG')

                if 'extender_cable_positive' in cable_lengths:
                    extender_length_feet = round(cable_lengths['extender_cable_positive'] * 3.28084 * self.CABLE_WASTE_FACTOR, 1)

                    block_quantities[f'Positive Extender Cable ({extender_cable_size})'] = {
                        'description': self.get_extender_description_format(extender_cable_size).format(length="TOTAL"),
                        'quantity': extender_length_feet,
                        'unit': 'feet',
                        'category': 'Extender Cables'
                    }

                if 'extender_cable_negative' in cable_lengths:
                    extender_length_feet = round(cable_lengths['extender_cable_negative'] * 3.28084 * self.CABLE_WASTE_FACTOR, 1)

                    block_quantities[f'Negative Extender Cable ({extender_cable_size})'] = {
                        'description': self.get_extender_description_format(extender_cable_size).format(length="TOTAL"),
                        'quantity': extender_length_feet,
                        'unit': 'feet',
                        'category': 'Extender Cables'
                    }
            
            quantities[block_id] = block_quantities
        
        # Store the original quantities before segment analysis
        original_quantities = {}
        for block_id, block_quantities in quantities.items():
            original_quantities[block_id] = {k: v.copy() for k, v in block_quantities.items() if 'Whip Cable' in k and v['unit'] == 'feet'}
        
        quantities = self.analyze_wire_segments(quantities)

        return quantities
        
    def _count_harnesses_by_size(self, block: BlockConfig) -> Dict[int, int]:
        """
        Count harnesses by number of strings they connect
        
        Args:
            block: Block configuration
            
        Returns:
            Dictionary mapping string count to harness count
        """
        harness_counts = {}
        
        if not block.wiring_config or block.wiring_config.wiring_type != WiringType.HARNESS:
            return harness_counts
        
        # Check if we have custom harness groupings
        has_custom_groupings = (hasattr(block.wiring_config, 'harness_groupings') and 
                            block.wiring_config.harness_groupings)
        
        if has_custom_groupings:
            # Count custom harness groups
            for string_count, harness_groups in block.wiring_config.harness_groupings.items():
                # Count trackers with this string count
                tracker_count = sum(1 for pos in block.tracker_positions if len(pos.strings) == string_count)
                
                # For each harness group configuration, multiply by number of trackers
                for harness in harness_groups:
                    # The key is the number of strings in this harness group
                    actual_string_count = len(harness.string_indices)
                    if actual_string_count not in harness_counts:
                        harness_counts[actual_string_count] = 0
                    harness_counts[actual_string_count] += tracker_count
                
                # Check for unconfigured strings
                all_configured_indices = set()
                for harness in harness_groups:
                    all_configured_indices.update(harness.string_indices)
                
                unconfigured_strings = string_count - len(all_configured_indices)
                if unconfigured_strings > 0:
                    # Add default harness for unconfigured strings
                    if unconfigured_strings not in harness_counts:
                        harness_counts[unconfigured_strings] = 0
                    harness_counts[unconfigured_strings] += tracker_count
        else:
            # Use the default behavior - one harness per tracker
            for pos in block.tracker_positions:
                if pos.template:
                    string_count = len(pos.strings)
                    if string_count not in harness_counts:
                        harness_counts[string_count] = 0
                    harness_counts[string_count] += 1
        
        return harness_counts
    
    def generate_summary_data(self, quantities: Dict[str, Dict[str, Any]]) -> pd.DataFrame:
        """
        Generate summary data for BOM with categories
        
        Args:
            quantities: Component quantities by block
            
        Returns:
            DataFrame with summary data
        """
        # Initialize dictionaries to track totals by component, description, and category
        component_totals = {}
        
        # Sum up quantities for each component type across all blocks
        for block_id, block_quantities in quantities.items():
            for component_type, details in block_quantities.items():
                description = details['description']
                quantity = details['quantity']
                unit = details['unit']
                category = details['category']
                
                # Skip tracker entries (structural category with "Tracker" in component type)
                if category == 'Structural' and 'Tracker' in component_type:
                    continue
                    
                key = (component_type, description, unit, category)
                if key not in component_totals:
                    component_totals[key] = 0
                component_totals[key] += quantity
        
        # Convert to DataFrame
        summary_data = []
        for (component_type, description, unit, category), quantity in component_totals.items():
            summary_data.append({
                'Category': category,
                'Component Type': component_type,
                'Part Number': '',
                'Description': description,
                'Quantity': round(quantity, 1) if unit == 'feet' else int(quantity),
                'Unit': unit
            })
        
        # Sort by category then component type, with numerical sorting for segments
        def sort_key(item):
            category = item['Category']
            component_type = item['Component Type']
            
            # For segment entries, extract the numeric part for proper numerical sorting
            if 'Segment' in component_type and 'ft' in component_type:
                import re
                # Extract the numeric part from patterns like "Positive Whip Cable Segment 115ft (8 AWG)"
                match = re.search(r'Segment (\d+)ft', component_type)
                if match:
                    length = int(match.group(1))
                    # Create a sort key that groups by cable type and sorts numerically by length
                    cable_type = component_type.split(' Segment')[0]  # e.g., "Positive Whip Cable"
                    return (category, cable_type, length)
            
            # For non-segment entries, sort normally
            return (category, component_type, 0)

        summary_data = sorted(summary_data, key=sort_key)

        # Add part numbers to summary data
        for item in summary_data:
            item['Part Number'] = self.get_component_part_number(item)

        # Create DataFrame with specific column order
        df = pd.DataFrame(summary_data)
        if not df.empty:
            # Reorder columns to put Part Number after Component Type
            columns = list(df.columns)
            if 'Part Number' in columns and 'Component Type' in columns:
                # Remove Part Number from its current position
                columns.remove('Part Number')
                # Insert it after Component Type
                ct_index = columns.index('Component Type')
                columns.insert(ct_index + 1, 'Part Number')
                df = df[columns]

        return df
    
    def get_component_part_number(self, item):
        """Get part number for a summary data item"""
        component_type = item['Component Type']
        description = item['Description']
        
        if 'Harness' in component_type:                
            # Extract info from description to find harness part number
            polarity = 'positive' if 'Positive' in description else 'negative'
            
            # Extract string count
            import re
            string_match = re.search(r'(\d+)-String', description)
            if string_match:
                num_strings = int(string_match.group(1))
                
                # Get module specs from first block
                if self.blocks:
                    first_block = next(iter(self.blocks.values()))
                    if first_block.tracker_template and first_block.tracker_template.module_spec:
                        module_spec = first_block.tracker_template.module_spec
                        modules_per_string = first_block.tracker_template.modules_per_string
                        module_spacing_m = first_block.tracker_template.module_spacing_m
                        
                        string_spacing_ft = self.calculate_string_spacing_ft(
                            modules_per_string, module_spec.width_mm, module_spacing_m
                        )
                        
                        # Get trunk cable size from wiring config
                        trunk_cable_size = getattr(first_block.wiring_config, 'harness_cable_size', '8 AWG')
                        
                        return self.find_matching_harness_part_number(
                            num_strings, polarity, string_spacing_ft, trunk_cable_size
                        )
        
        elif 'Fuse' in component_type:
            # Extract fuse rating from description
            import re
            rating_match = re.search(r'(\d+)A', description)
            if rating_match:
                rating = int(rating_match.group(1))
                return self.get_fuse_part_number_by_rating(rating)
        
        # Handle whip cable segments
        elif 'Whip Cable Segment' in component_type:
            return self.get_whip_segment_part_number_from_item(item)
        
        # Handle extender cable segments  
        elif 'Extender Cable Segment' in component_type:
            return self.get_extender_segment_part_number_from_item(item)
        
        # Handle string cable segments (use extender part numbers)
        elif 'String Cable Segment' in component_type:
            return self.get_extender_segment_part_number_from_item(item)
        
        return "N/A"
    
    def get_whip_segment_part_number_from_item(self, item):
        """Get whip cable segment part number from item data"""
        try:
            component_type = item['Component Type']
            
            # Extract polarity from component type
            polarity = 'positive' if 'Positive' in component_type else 'negative'
            
            # Extract length from component type (e.g., "Positive Whip Cable Segment 25ft (8 AWG)")
            import re
            length_match = re.search(r'(\d+)ft', component_type)
            if not length_match:
                return "N/A"
            length_ft = int(length_match.group(1))
            
            # Extract wire gauge from component type
            gauge_match = re.search(r'\(([^)]+)\)', component_type)
            if not gauge_match:
                return "N/A"
            wire_gauge = gauge_match.group(1)
            
            # Find matching whip part number
            return self.find_matching_whip_part_number(wire_gauge, polarity, length_ft)
            
        except Exception as e:
            print(f"Error getting whip segment part number: {e}")
            return "N/A"

    def get_extender_segment_part_number_from_item(self, item):
        """Get extender cable segment part number from item data"""
        try:
            component_type = item['Component Type']
            
            # Extract polarity from component type
            polarity = 'positive' if 'Positive' in component_type else 'negative'
            
            # Extract length from component type (e.g., "Positive Extender Cable Segment 30ft (8 AWG)")
            import re
            length_match = re.search(r'(\d+)ft', component_type)
            if not length_match:
                return "N/A"
            length_ft = int(length_match.group(1))
            
            # Extract wire gauge from component type
            gauge_match = re.search(r'\(([^)]+)\)', component_type)
            if not gauge_match:
                return "N/A"
            wire_gauge = gauge_match.group(1)
            
            # Find matching extender part number
            return self.find_matching_extender_part_number(wire_gauge, polarity, length_ft)
            
        except Exception as e:
            print(f"Error getting extender segment part number: {e}")
            return "N/A"
    
    def generate_detailed_data(self, quantities: Dict[str, Dict[str, Any]]) -> pd.DataFrame:
        """
        Generate detailed data for BOM
        
        Args:
            quantities: Component quantities by block
            
        Returns:
            DataFrame with detailed data
        """
        detailed_data = []
        
        for block_id, block_quantities in quantities.items():
            for component_type, details in block_quantities.items():
                description = details['description']
                quantity = details['quantity']
                unit = details['unit']
                category = details['category']
                
                # Skip tracker entries (structural category with "Tracker" in component type)
                if category == 'Structural' and 'Tracker' in component_type:
                    continue
                    
                item_data = {
                    'Block': block_id,
                    'Category': category,
                    'Component Type': component_type,
                    'Description': description,
                    'Quantity': round(quantity, 1) if unit == 'feet' else int(quantity),
                    'Unit': unit
                }

                # Add part number
                item_data['Part Number'] = self.get_component_part_number(item_data)

                detailed_data.append(item_data)
        
        # Sort by block ID, category, and component type
        detailed_data = sorted(detailed_data, key=lambda x: (x['Block'], x['Category'], x['Component Type']))
        
        return pd.DataFrame(detailed_data)


    def export_bom_to_excel_with_preview_data(self, filepath: str, project_info: Optional[Dict[str, Any]] = None, 
                                          preview_data: List[Dict] = None) -> bool:
        """
        Export BOM to Excel using preview data from UI
        
        Args:
            filepath: Path to save the Excel file
            project_info: Optional dictionary with project information
            preview_data: Data from the UI preview with correct part numbers
            
        Returns:
            True if export successful, False otherwise
        """
        writer = None
        try:
            # Use preview data if provided
            if preview_data:
                # Convert preview data to DataFrame
                summary_data = pd.DataFrame(preview_data)
                
                # Add Category column (infer from component type)
                def get_category(component_type):
                    if 'Whip Cable Segment' in component_type:
                        return 'eBOS Segments'
                    elif 'Extender Cable Segment' in component_type:
                        return 'Extender Cable Segments'
                    elif 'String Cable' in component_type:
                        return 'eBOS'
                    elif 'Harness' in component_type:
                        return 'eBOS'
                    elif 'Fuse' in component_type:
                        return 'Electrical'
                    else:
                        return 'Other'
                
                summary_data['Category'] = summary_data['Component Type'].apply(get_category)
                
                # Reorder columns to match expected format
                column_order = ['Category', 'Component Type', 'Part Number', 'Description', 'Quantity', 'Unit']
                summary_data = summary_data[column_order]
            else:
                # Fall back to original method
                quantities = self.calculate_cable_quantities()
                summary_data = self.generate_summary_data(quantities)
            
            # Generate detailed data (this can still use the original method)
            quantities = self.calculate_cable_quantities() 
            detailed_data = self.generate_detailed_data(quantities)
            
            # Generate project info if not provided
            if project_info is None:
                project_info = self.generate_project_info()
            
            # Create Excel writer
            writer = pd.ExcelWriter(filepath, engine='openpyxl')
            
            # Write summary data
            summary_data.to_excel(writer, sheet_name='BOM Summary', index=False, startrow=13)  # Start after project info
            
            # Write detailed data  
            detailed_data.to_excel(writer, sheet_name='Block Details', index=False)

            # Add combiner box sheet if there are any combiner boxes
            combiner_box_count = self.count_combiner_boxes()
            if combiner_box_count > 0:
                # Check if device configurator has combiner configs
                from ..ui.device_configurator import DeviceConfigurator
                device_configs = {}
                
                # Try to get device configurations from the UI if available
                if hasattr(self, 'parent') and hasattr(self.parent, 'device_configurator'):
                    device_configurator = self.parent.device_configurator
                    if hasattr(device_configurator, 'combiner_configs'):
                        device_configs = device_configurator.combiner_configs
                
                # Generate combiner box data
                combiner_data = self.generate_combiner_box_data(device_configs)
                combiner_data.to_excel(writer, sheet_name='Combiner Boxes', index=False)
            
            # Format the sheets (same as original method)
            workbook = writer.book
            summary_sheet = writer.sheets['BOM Summary']
            
            # Add project info to summary sheet
            row = 1
            summary_sheet.merge_cells(f'A{row}:F{row}')
            project_info_cell = summary_sheet.cell(row=row, column=1, value="Project Information")
            project_info_cell.font = Font(bold=True, size=14)
            project_info_cell.alignment = Alignment(horizontal='center')
            
            if project_info:
                # First set of info in columns A and B
                main_info = ['Project Name', 'Customer', 'Location', 'System Size (kW DC)', 
                            'Number of Modules', 'Module Manufacturer', 'Module Model', 
                            'Inverter Manufacturer', 'Inverter Model', 'DC Collection']
                
                # Additional info for columns C and D
                additional_info = {
                    'String Size': project_info.get('String Size', 'Unknown'),
                    'Number of Strings': project_info.get('Number of Strings', 0),
                    'Module Wiring': project_info.get('Module Wiring', 'Unknown'),
                    'Module Dimensions': project_info.get('Module Dimensions', 'Unknown'),
                    'Number of Combiner Boxes': project_info.get('Number of Combiner Boxes', 0)
                }
                
                # Write main info with units
                row = 2
                for key in main_info:
                    if key in project_info:
                        value = project_info[key]
                        
                        # Add units to specific fields
                        if key == 'System Size (kW DC)' and isinstance(value, (int, float)):
                            value = f"{round(value, 2)} kW"
                        elif key == 'Number of Modules' and isinstance(value, int):
                            value = f"{value} modules"
                        
                        summary_sheet.cell(row=row, column=1, value=key).font = Font(bold=True)
                        summary_sheet.cell(row=row, column=2, value=value)
                        row += 1
                
                # Write additional info in columns C and D
                row = 2
                for key, value in additional_info.items():
                    # Add units to specific fields
                    if key == 'String Size' and value != 'Unknown':
                        value = f"{value} modules per string"
                    elif key == 'Number of Strings' and isinstance(value, int):
                        value = f"{value} strings"
                    elif key == 'Number of Combiner Boxes' and isinstance(value, int):
                        if value == 1:
                            value = f"{value} combiner box"
                        else:
                            value = f"{value} combiner boxes"
                    
                    summary_sheet.cell(row=row, column=3, value=key).font = Font(bold=True)
                    summary_sheet.cell(row=row, column=4, value=value)
                    row += 1

            # Add section header for BOM
            row = 12
            summary_sheet.merge_cells(f'A{row}:E{row}')
            summary_sheet.cell(row=row, column=1, value="Bill of Materials").font = Font(bold=True, size=14)
            
            # Format sheets
            self._format_excel_sheet(workbook['BOM Summary'], summary_data, start_row=13)
            self._format_excel_sheet(workbook['Block Details'], detailed_data)
            
            # Format combiner box sheet if it exists
            if 'Combiner Boxes' in workbook.sheetnames and combiner_box_count > 0:
                self._format_combiner_sheet(workbook['Combiner Boxes'], combiner_data)
            
            # Add filter
            summary_sheet.auto_filter.ref = f"A13:F{13 + len(summary_data)}"
            
            # Save and open
            writer.close()
            writer = None
            
            try:
                os.startfile(filepath)
            except Exception as e:
                print(f"File was saved but could not be opened automatically: {str(e)}")

            return True
            
        except Exception as e:
            print(f"Error exporting BOM: {str(e)}")
            if isinstance(e, PermissionError):
                raise
            return False
        finally:
            if writer is not None:
                try:
                    writer.close()
                except:
                    pass

    def filter_data_by_checked_components(self, data_df, checked_components, is_detailed=False):
        """Filter DataFrame based on checked components"""
        if not checked_components:
            return data_df
        
        # Create set of checked component descriptions for fast lookup
        checked_descriptions = {comp['description'] for comp in checked_components}
        
        # Filter DataFrame
        if is_detailed:
            # For detailed data, filter by Description column
            filtered_df = data_df[data_df['Description'].isin(checked_descriptions)]
        else:
            # For summary data, filter by Description column
            filtered_df = data_df[data_df['Description'].isin(checked_descriptions)]
        
        return filtered_df.reset_index(drop=True)                       

    def _format_excel_sheet(self, worksheet, data: pd.DataFrame, start_row: int = 1):
        """
        Format Excel worksheet
        
        Args:
            worksheet: openpyxl worksheet
            data: DataFrame with data
            start_row: Row to start formatting from (default=1)
        """
        # Define styles
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        centered_alignment = Alignment(horizontal='center')
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Warning style
        warning_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")  # Light red
        
        # Format headers
        for col_num, column_title in enumerate(data.columns, 1):
            cell = worksheet.cell(row=start_row, column=col_num)
            cell.value = column_title
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = centered_alignment
            cell.border = border

        # Apply center alignment to Part Number column data
        if 'Part Number' in data.columns:
            part_number_col = list(data.columns).index('Part Number') + 1
            for row_num in range(start_row + 1, start_row + len(data) + 2):  # +2 to include last row
                cell = worksheet.cell(row=row_num, column=part_number_col)
                cell.alignment = centered_alignment
        
        for i, row in enumerate(worksheet.iter_rows(min_row=start_row+1, 
                                                    max_row=start_row+len(data)+1,  # +1 to ensure last row is included
                                                    min_col=1, 
                                                    max_col=len(data.columns)), 1):
            # Get the row index
            row_index = start_row + i
            
            # Check if this is a warning row
            is_warning = False
            if 'Unit' in data.columns:
                unit_col_idx = list(data.columns).index('Unit')
                unit_value = worksheet.cell(row=row_index, column=unit_col_idx + 1).value
                if unit_value and 'warning' in str(unit_value).lower():
                    is_warning = True
            
            # Apply warning formatting if this is a warning row
            if is_warning:
                for cell in row:
                    cell.fill = warning_fill
            
            # Add borders to all cells
            for cell in row:
                cell.border = border
        
        # Auto-adjust column width
        for column in worksheet.columns:
            max_length = 0
            column_name = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column_name].width = adjusted_width

    def generate_project_info(self) -> Dict[str, Any]:
        """
        Generate general project information for BOM header
        
        Returns:
            Dictionary with project info
        """
        # Initialize with default values
        info = {
            'System Size (kW DC)': 0,
            'Number of Modules': 0,
            'Module Manufacturer': 'Unknown',
            'Module Model': 'Unknown',
            'DC Collection': 'Unknown',
            'Inverter Manufacturer': 'Unknown',
            'Inverter Model': 'Unknown',
            'String Size': 0,
            'Number of Strings': 0,
            'Module Wiring': 'Unknown',
            'Module Dimensions': 'Unknown',
            'Number of Combiner Boxes': 0
        }
        
        # Calculate total system size and module count
        total_modules = 0
        module_manufacturer = set()
        module_model = set()
        inverter_manufacturer = set()
        inverter_model = set()
        dc_collection_types = set()
        
        # New counters for additional info
        total_strings = 0
        string_sizes = set()
        module_dimensions = set()
        
        for block_id, block in self.blocks.items():
            # Count modules
            if block.tracker_template and block.tracker_template.module_spec:
                module_spec = block.tracker_template.module_spec
                block_modules = 0
                
                # Count modules directly from strings
                for pos in block.tracker_positions:
                    for string in pos.strings:
                        block_modules += string.num_modules
                
                total_modules += block_modules
                
                # Add module info
                module_manufacturer.add(module_spec.manufacturer)
                module_model.add(module_spec.model)
                
                # Calculate system size
                info['System Size (kW DC)'] += (block_modules * module_spec.wattage) / 1000
                
                # Get module dimensions
                dim_str = f"{int(module_spec.length_mm)} mm x {int(module_spec.width_mm)} mm"
                module_dimensions.add(dim_str)

            # Count combiner boxes
            combiner_box_count = self.count_combiner_boxes()
            info['Number of Combiner Boxes'] = combiner_box_count
            
            # Count strings and get string size
            if block.tracker_positions:
                for pos in block.tracker_positions:
                    total_strings += len(pos.strings)
            
            # Get string size from tracker template
            if block.tracker_template:
                string_sizes.add(block.tracker_template.modules_per_string)
            
            # Add inverter info
            if block.inverter:
                inverter_manufacturer.add(block.inverter.manufacturer)
                inverter_model.add(block.inverter.model)
            
            # Add DC collection type
            if block.wiring_config:
                dc_collection_types.add(block.wiring_config.wiring_type.value)
        
        # Set the collected values
        info['Number of Modules'] = total_modules
        info['Module Manufacturer'] = ', '.join(module_manufacturer) if module_manufacturer else 'Unknown'
        info['Module Model'] = ', '.join(module_model) if module_model else 'Unknown'
        info['Inverter Manufacturer'] = ', '.join(inverter_manufacturer) if inverter_manufacturer else 'Unknown'
        info['Inverter Model'] = ', '.join(inverter_model) if inverter_model else 'Unknown'
        info['DC Collection'] = ', '.join(dc_collection_types) if dc_collection_types else 'Unknown'
        
        # Set new fields
        info['String Size'] = ', '.join(str(s) for s in sorted(string_sizes)) if string_sizes else 'Unknown'
        info['Number of Strings'] = total_strings
        info['Module Dimensions'] = ', '.join(module_dimensions) if module_dimensions else 'Unknown'
        info['Number of Combiner Boxes'] = len(self.blocks)
        
        # Get module wiring from project
        if hasattr(self.project, 'wiring_mode'):
            wiring_mode = "Leapfrog" if self.project.wiring_mode == 'leapfrog' else "Daisy Chain"
            info['Module Wiring'] = wiring_mode
        
        return info

    def analyze_wire_segments(self, quantities: Dict[str, Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
        """Analyze wire segments and add segment counts to quantities"""
        for block_id, block in self.blocks.items():
            block_quantities = quantities.get(block_id, {})
            
            # Skip blocks without wiring config
            if not block.wiring_config:
                continue
            
            # Get realistic cable routes if available, otherwise use regular routes
            cable_routes = getattr(block.wiring_config, 'realistic_cable_routes', {})
            if not cable_routes:
                cable_routes = block.wiring_config.cable_routes
            
            # Group routes by type and analyze segments
            string_segments_pos = []
            string_segments_neg = []
            whip_segments_pos = []
            whip_segments_neg = []
            
            for route_id, route in cable_routes.items():
                # Skip routes with less than 2 points
                if len(route) < 2:
                    continue
                    
                # Calculate segment length
                segment_length = 0
                for i in range(len(route) - 1):
                    dx = route[i+1][0] - route[i][0]
                    dy = route[i+1][1] - route[i][1]
                    segment_length += (dx**2 + dy**2)**0.5
                
                # Convert to feet
                segment_length_feet = segment_length * 3.28084
                
                # Categorize by route type and polarity
                if "pos_src" in route_id or "pos_node" in route_id or "pos_string" in route_id:
                    string_segments_pos.append(segment_length_feet)
                elif "neg_src" in route_id or "neg_node" in route_id or "neg_string" in route_id:
                    string_segments_neg.append(segment_length_feet)
                elif "pos_dev" in route_id or "pos_main" in route_id or "whip_pos" in route_id or "pos_whip" in route_id:
                    # Add underground routing component if enabled
                    if hasattr(block, 'underground_routing') and block.underground_routing:
                        underground_addition_m = 2 * (block.pile_reveal_m + block.trench_depth_m)
                        underground_addition_ft = underground_addition_m * 3.28084
                        segment_length_feet += underground_addition_ft
                    whip_segments_pos.append(segment_length_feet)
                elif "neg_dev" in route_id or "neg_main" in route_id or "whip_neg" in route_id or "neg_whip" in route_id:
                    # Add underground routing component if enabled
                    if hasattr(block, 'underground_routing') and block.underground_routing:
                        underground_addition_m = 2 * (block.pile_reveal_m + block.trench_depth_m)
                        underground_addition_ft = underground_addition_m * 3.28084
                        segment_length_feet += underground_addition_ft
                    whip_segments_neg.append(segment_length_feet)
            
            # Process string segments only for HOMERUN wiring
            if block.wiring_config.wiring_type == WiringType.HOMERUN:
                string_size = block.wiring_config.string_cable_size
                self._add_segment_analysis(block_quantities, string_segments_pos, 
                                        string_size, "Positive String Cable", 5)
                self._add_segment_analysis(block_quantities, string_segments_neg, 
                                        string_size, "Negative String Cable", 5)
                
                # Calculate and add total string entries from segments
                self.calculate_totals_from_segments(block_quantities, string_size, "Positive String Cable")
                self.calculate_totals_from_segments(block_quantities, string_size, "Negative String Cable")
            
            # Process whip segments for all wiring types
            whip_size = getattr(block.wiring_config, 'whip_cable_size', "8 AWG")
            self._add_segment_analysis(block_quantities, whip_segments_pos, 
                                    whip_size, "Positive Whip Cable", 1)
            self._add_segment_analysis(block_quantities, whip_segments_neg, 
                                    whip_size, "Negative Whip Cable", 1)
            
            # Calculate and add total whip entries from segments
            self.calculate_totals_from_segments(block_quantities, whip_size, "Positive Whip Cable")
            self.calculate_totals_from_segments(block_quantities, whip_size, "Negative Whip Cable")

            # Process extender segments
            extender_segments_pos = []
            extender_segments_neg = []

            for route_id, route in cable_routes.items():
                # Skip routes with less than 2 points
                if len(route) < 2:
                    continue
                    
                # Calculate segment length
                segment_length = 0
                for i in range(len(route) - 1):
                    dx = route[i+1][0] - route[i][0]
                    dy = route[i+1][1] - route[i][1]
                    segment_length += (dx**2 + dy**2)**0.5
                
                # Convert to feet
                segment_length_feet = segment_length * 3.28084
                
                # Categorize extender routes
                if "pos_extender" in route_id:
                    extender_segments_pos.append(segment_length_feet)
                elif "neg_extender" in route_id:
                    extender_segments_neg.append(segment_length_feet)

            # Process extender segments
            extender_size = getattr(block.wiring_config, 'extender_cable_size', "8 AWG")
            self._add_segment_analysis(block_quantities, extender_segments_pos, 
                                    extender_size, "Positive Extender Cable", 5)
            self._add_segment_analysis(block_quantities, extender_segments_neg, 
                                    extender_size, "Negative Extender Cable", 5)

            # Calculate and add total extender entries from segments
            self.calculate_totals_from_segments(block_quantities, extender_size, "Positive Extender Cable")
            self.calculate_totals_from_segments(block_quantities, extender_size, "Negative Extender Cable")
                    
            # Update quantities
            quantities[block_id] = block_quantities
        
        return quantities
    
    def load_harness_library(self):
        """Load harness library from JSON file"""
        try:
            # Get the path relative to this file (src/utils/bom_generator.py)
            current_dir = os.path.dirname(os.path.abspath(__file__))  # src/utils/
            project_root = os.path.dirname(os.path.dirname(current_dir))  # Go up two levels to get to project root
            library_path = os.path.join(project_root, 'data', 'harness_library.json')
            
            with open(library_path, 'r') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading harness library: {e}")
            return {}
        
    def load_extender_library(self):
        """Load the extender library from JSON file"""
        try:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            project_root = os.path.dirname(os.path.dirname(current_dir))
            library_path = os.path.join(project_root, 'data', 'extender_library.json')
            
            with open(library_path, 'r') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading extender library: {e}")
            return {}

    def load_whip_library(self):
        """Load the whip library from JSON file"""
        try:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            project_root = os.path.dirname(os.path.dirname(current_dir))
            library_path = os.path.join(project_root, 'data', 'whip_library.json')
            
            with open(library_path, 'r') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading whip library: {e}")
            return {}
        
    def get_whip_description_format(self, wire_gauge):
        """Get whip description format from library"""
        # Find any whip with matching wire gauge to get description format
        for spec in self.whip_library.values():
            if spec.get('wire_gauge') == wire_gauge:
                # Extract format without the length
                desc = spec.get('description', '')
                # Replace the length part with placeholder
                import re
                return re.sub(r'LENGTH \(FT\): \d+', 'LENGTH (FT): {length}', desc)
        # Default format if not found
        return f"1500 VDC {wire_gauge} WHIP WITH MC4 CONNECTOR AND BLUNT CUT END, LENGTH (FT): {{length}}"

    def get_extender_description_format(self, wire_gauge):
        """Get extender description format from library"""
        # Find any extender with matching wire gauge to get description format
        for spec in self.extender_library.values():
            if spec.get('wire_gauge') == wire_gauge:
                # Extract format without the length
                desc = spec.get('description', '')
                # Replace the length part with placeholder
                import re
                return re.sub(r'LENGTH \(FT\): \d+', 'LENGTH (FT): {length}', desc)
        # Default format if not found
        return f"1500 VDC {wire_gauge} EXTENDER WITH MC4 CONNECTORS, LENGTH (FT): {{length}}"
    
    def get_harness_description(self, num_strings, polarity, string_spacing_ft, trunk_cable_size, string_cable_size):
        """Get harness description from library based on matching part number"""      
        # First find the matching part number
        part_number = self.find_matching_harness_part_number(
            num_strings, polarity, string_spacing_ft, trunk_cable_size
        )
        
        if part_number == "N/A" or " or " in part_number:
            # If no match or multiple matches, return a generic description
            pol_text = "Positive" if polarity == 'positive' else "Negative"
            string_text = "String" if num_strings == 1 else "String"
            fuse_text = ""
            
            # Try to determine if this should be fused based on polarity and string count
            if polarity == 'positive' and num_strings > 1:
                # Make an educated guess about fuse rating based on string count
                if num_strings <= 2:
                    fuse_text = "20A Fuses, " if string_spacing_ft > 110 else "Unfused, "
                else:
                    fuse_text = "30A Fuses, " if string_spacing_ft > 120 else "20A Fuses, "
            
            return f"{num_strings} {string_text}, {pol_text}, {fuse_text}{string_cable_size} Drops w/{trunk_cable_size} Trunk, {int(string_spacing_ft)}' string length, MC4 connectors"
        
        # If single part number found, get its description
        if part_number in self.harness_library:
            return self.harness_library[part_number].get('description', f"Harness {part_number}")
        
        # Handle multiple matches by using first one's description format
        if " or " in part_number:
            first_part = part_number.split(" or ")[0]
            if first_part in self.harness_library:
                return self.harness_library[first_part].get('description', f"Harness {first_part}")
        
        return f"Harness {part_number}"
    
    def get_fuse_description(self, fuse_rating_amps):
        """Get fuse description from library based on rating"""
        # Find the fuse part number by rating
        part_number = self.get_fuse_part_number_by_rating(fuse_rating_amps)
        
        # If we have a valid part number, get its description
        if part_number and part_number != "N/A" and not part_number.startswith("FUSE-"):
            if part_number in self.fuse_library:
                return self.fuse_library[part_number].get('description', f'DC String Fuse {fuse_rating_amps}A')
        
        # Default description if not found
        return f'DC String Fuse {fuse_rating_amps}A'

    def calculate_string_spacing_ft(self, modules_per_string, module_width_mm, module_spacing_m):
        """Calculate string spacing in feet based on module specs"""
        try:
            # Calculate total string length: (modules * width) + ((modules-1) * spacing between modules)
            total_length_mm = (modules_per_string * module_width_mm + 
                            (modules_per_string - 1) * module_spacing_m * 1000)
            
            # Convert to feet
            total_length_ft = total_length_mm / 1000 * 3.28084
            
            return round(total_length_ft, 1)
        except Exception as e:
            print(f"Error calculating string spacing: {e}")
            return 0

    def find_matching_harness_part_number(self, num_strings, polarity, calculated_spacing_ft, trunk_cable_size=None):
        """Find matching harness part number from library"""
        try:
            matches = []
            
            # Available harness spacing options based on your library
            # First Solar uses 26', Standard uses 102, 113, 122, 133
            available_spacings = [26.0, 102.0, 113.0, 122.0, 133.0]
            
            # Find the next largest available spacing
            target_spacing = None
            for spacing in sorted(available_spacings):
                if spacing >= calculated_spacing_ft:
                    target_spacing = spacing
                    break
            
            if target_spacing is None:
                target_spacing = max(available_spacings)  # Use largest if calculated is bigger than all
            
            # print(f"Calculated spacing: {calculated_spacing_ft}ft, Target spacing: {target_spacing}ft, Strings: {num_strings}, Polarity: {polarity}")
            
            # Search for matching harnesses
            for part_number, spec in self.harness_library.items():
                # Skip comment entries
                if part_number.startswith('_comment_'):
                    continue
                    
                # Check basic criteria
                if (spec.get('num_strings') == num_strings and 
                    spec.get('polarity') == polarity and
                    abs(spec.get('string_spacing_ft', 0) - target_spacing) < 0.1):  # Allow small tolerance
                    
                    # If trunk cable size is specified, filter by it
                    if trunk_cable_size:
                        spec_trunk_size = spec.get('trunk_cable_size', spec.get('trunk_wire_gauge'))
                        if spec_trunk_size != trunk_cable_size:
                            continue
                            
                    matches.append(part_number)
            
            if matches:
                if len(matches) == 1:
                    return matches[0]
                else:
                    return " or ".join(sorted(matches))  # Sort for consistent output
            else:
                # print(f"No matches found for {num_strings} strings, {polarity}, {target_spacing}ft")
                return "N/A"
                
        except Exception as e:
            print(f"Error finding harness match: {e}")
            return "N/A"
        
    def find_matching_extender_part_number(self, wire_gauge, polarity, required_length_ft):
        """Find matching extender part number from library"""
        try:
            # Round up to next 5ft increment for extender lengths
            target_length = ((required_length_ft - 1) // 5 + 1) * 5
            target_length = max(10, target_length)  # Remove the max 300 clamp
            
            # If over 300ft, return custom part number with wire gauge and polarity
            if target_length > 300:
                polarity_code = 'P' if polarity == 'positive' else 'N'
                awg_size = wire_gauge.split()[0]  # Extract just the number from "8 AWG"
                return f"EXT-{awg_size}-{polarity_code}-{target_length}-CUSTOM"
            
            # Search for matching extender
            for part_number, spec in self.extender_library.items():
                # Skip comment entries
                if part_number.startswith('_comment_'):
                    continue
                if (spec.get('wire_gauge') == wire_gauge and 
                    spec.get('polarity') == polarity and
                    spec.get('length_ft') == target_length):
                    return part_number
            
            print(f"No extender found for {wire_gauge}, {polarity}, {target_length}ft")
            return "N/A"
            
        except Exception as e:
            print(f"Error finding extender part number: {e}")
            return "N/A"

    def find_matching_whip_part_number(self, wire_gauge, polarity, required_length_ft):
        """Find matching whip part number from library"""
        try:
            # Round up to next 5ft increment for whip lengths
            target_length = ((required_length_ft - 1) // 5 + 1) * 5
            target_length = max(10, target_length)  # Remove the max 300 clamp

            # If over 300ft, return custom part number with wire gauge and polarity
            if target_length > 300:
                polarity_code = 'P' if polarity == 'positive' else 'N'
                awg_size = wire_gauge.split()[0]  # Extract just the number from "8 AWG"
                return f"WHI-{awg_size}-{polarity_code}-{target_length}-CUSTOM"
            
            # Search for matching whip
            for part_number, spec in self.whip_library.items():
                # Skip comment entries
                if part_number.startswith('_comment_'):
                    continue
                if (spec.get('wire_gauge') == wire_gauge and 
                    spec.get('polarity') == polarity and
                    spec.get('length_ft') == target_length):
                    return part_number
            
            print(f"No whip found for {wire_gauge}, {polarity}, {target_length}ft")
            return "N/A"
            
        except Exception as e:
            print(f"Error finding whip part number: {e}")
            return "N/A"
        
    def get_fuse_part_number_by_rating(self, fuse_rating_amps):
        """Get fuse part number by rating from fuse library"""
        try:
            # Find exact match first
            for part_number, spec in self.fuse_library.items():
                if spec.get('fuse_rating_amps') == fuse_rating_amps:
                    return part_number
            
            # If no exact match, find the next higher rating
            available_ratings = []
            for spec in self.fuse_library.values():
                rating = spec.get('fuse_rating_amps')
                if rating and rating >= fuse_rating_amps:
                    available_ratings.append((rating, spec.get('part_number')))
            
            if available_ratings:
                # Sort by rating and return the lowest that meets the requirement
                available_ratings.sort(key=lambda x: x[0])
                return available_ratings[0][1]
            
            # If no suitable fuse found, return placeholder
            return f"FUSE-{fuse_rating_amps}A-NOT-FOUND"
            
        except Exception as e:
            print(f"Error getting fuse part number: {e}")
            return f"FUSE-{fuse_rating_amps}A-ERROR"

    def _add_segment_analysis(self, block_quantities, segments, cable_size, 
                   prefix, length_increment=5):
        """Add segment analysis for a specific cable type"""
        
        if not segments:
            return
            
        # Add waste factor
        segments = [s * self.CABLE_WASTE_FACTOR for s in segments]
        
        # Always use 5ft increment for whip cables, otherwise use the passed-in increment
        actual_increment = 5 if "Whip Cable" in prefix else length_increment
        
        # Round up to nearest increment, with minimum of 10ft
        rounded_segments = []
        for s in segments:
            rounded_length = actual_increment * ((s + actual_increment - 0.1) // actual_increment + 1)
            # Ensure minimum length of 10ft
            rounded_length = max(10, rounded_length)
            rounded_segments.append(rounded_length)
        
        # Count segments by length
        segment_counts = {}
        for length in rounded_segments:
            if length not in segment_counts:
                segment_counts[length] = 0
            segment_counts[length] += 1
            
        # Add segment counts to quantities
        for length, count in segment_counts.items():
            segment_key = f"{prefix} Segment {int(length)}ft ({cable_size})"
            
            # Determine category based on cable type
            if "Extender Cable" in prefix:
                category = 'Extender Cable Segments'
            else:
                category = 'eBOS Segments'
            
            # Generate proper description based on cable type
            if "Whip Cable" in prefix:
                base_desc = self.get_whip_description_format(cable_size)
                description = base_desc.format(length=int(length))
            elif "Extender Cable" in prefix:
                base_desc = self.get_extender_description_format(cable_size)
                description = base_desc.format(length=int(length))
            else:
                description = f"{int(length)}ft {prefix} Segment ({cable_size})"

            block_quantities[segment_key] = {
                'description': description,
                'quantity': count,
                'unit': 'segments',
                'category': category
            }

    def calculate_totals_from_segments(self, block_quantities, cable_size, prefix):
        """Calculate and add total cable entry from segment entries"""
        total_length = 0
        segment_entries = {}
        
        # Find all segment entries for this cable type
        for key, details in block_quantities.items():
            if prefix in key and "Segment" in key and details['unit'] == 'segments':
                # Extract length from key (format: "{prefix} Segment {length}ft ({cable_size})")
                try:
                    segment_str = key.split("Segment ")[1].split("ft")[0]
                    length = float(segment_str)
                    count = details['quantity']
                    total_length += length * count
                    segment_entries[key] = details
                except (ValueError, IndexError):
                    continue
        
        # Only add total if we found segments
        if total_length > 0:
            # For whip cables, ensure the total is an integer (no decimal places)
            # This maintains the 5ft increment pattern
            if "Whip Cable" in prefix:
                total_length = int(total_length)
            
            total_key = f"{prefix} ({cable_size})"

            # Determine category based on cable type
            if "Extender Cable" in prefix:
                category = 'Extender Cables'
            else:
                category = 'eBOS'

            # Generate proper description for totals
            if "Whip Cable" in prefix:
                description = self.get_whip_description_format(cable_size).format(length="TOTAL")
            elif "Extender Cable" in prefix:
                description = self.get_extender_description_format(cable_size).format(length="TOTAL")
            else:
                description = f'DC {prefix} {cable_size}'

            block_quantities[total_key] = {
                'description': description,
                'quantity': total_length if "Whip Cable" in prefix or "Extender Cable" in prefix else round(total_length, 1),
                'unit': 'feet',
                'category': category
            }
            
        return total_length > 0
    
    def _count_fuses_by_rating(self, block: BlockConfig) -> Dict[int, int]:
        """
        Count fuses by rating
        
        Args:
            block: Block configuration
            
        Returns:
            Dictionary mapping fuse rating to count
        """
        fuse_counts = {}
        
        if not block.wiring_config or block.wiring_config.wiring_type != WiringType.HARNESS:
            return fuse_counts
            
        # Check all string counts
        for string_count, harness_groups in getattr(block.wiring_config, 'harness_groupings', {}).items():
            # Count trackers with this string count - this is the key fix
            tracker_count = sum(1 for pos in block.tracker_positions if len(pos.strings) == string_count)
            
            for harness in harness_groups:
                # Only count fuses if use_fuse is True and there's more than one string
                use_fuse = getattr(harness, 'use_fuse', len(harness.string_indices) > 1)
                if use_fuse and len(harness.string_indices) > 1:
                    rating = getattr(harness, 'fuse_rating_amps', 15)
                    
                    # Count one fuse per string in the harness, multiplied by tracker count
                    num_strings = len(harness.string_indices)
                    fuse_counts[rating] = fuse_counts.get(rating, 0) + (num_strings * tracker_count)
        
        # Check for trackers without custom harness configuration
        if not hasattr(block.wiring_config, 'harness_groupings') or not block.wiring_config.harness_groupings:
            # Calculate default fuse rating
            default_rating = 15
            if block.tracker_template and block.tracker_template.module_spec:
                isc = block.tracker_template.module_spec.isc
                nec_min = isc * 1.25
                for rating in [5, 10, 15, 20, 25, 30, 35, 40, 45]:
                    if rating >= nec_min:
                        default_rating = rating
                        break
            
            # Count trackers with multiple strings
            for pos in block.tracker_positions:
                string_count = len(pos.strings)
                if string_count > 1:  # Only count for 2+ strings
                    # Add one fuse per string (positive side only)
                    fuse_counts[default_rating] = fuse_counts.get(default_rating, 0) + string_count
        
        return fuse_counts
    
    def load_fuse_library(self):
        """Load fuse library from JSON file"""
        try:
            # Get the path relative to this file
            current_dir = os.path.dirname(os.path.abspath(__file__))
            project_root = os.path.dirname(os.path.dirname(current_dir))
            library_path = os.path.join(project_root, 'data', 'fuse_library.json')
            
            with open(library_path, 'r') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading fuse library: {e}")
            return {}
        
    def generate_combiner_box_data(self, device_configs: Dict[str, 'CombinerBoxConfig']) -> pd.DataFrame:
        """
        Generate combiner box configuration data for Excel export
        
        Args:
            device_configs: Dictionary of combiner box configurations
            
        Returns:
            DataFrame with combiner box data
        """
        combiner_data = []
        
        for combiner_id in sorted(device_configs.keys()):
            config = device_configs[combiner_id]
            
            # Add a header row for each combiner box
            for i, conn in enumerate(config.connections):
                row_data = {
                    'Combiner': combiner_id,
                    'Tracker': conn.tracker_id,
                    'Harness': conn.harness_id,
                    '# Strings': conn.num_strings,
                    'Module Isc': f"{conn.module_isc:.2f}",
                    'NEC Safety Factor': conn.nec_factor,
                    'Harness Current': f"{conn.harness_current:.2f}",
                    'Fuse Size': conn.get_display_fuse_size(),
                    'Cable Size': conn.get_display_cable_size(),
                    'Total Current': '',  # Only on first row
                    'Breaker Size': ''    # Only on first row
                }
                
                # Add total current and breaker size only to first row
                if i == 0:
                    row_data['Total Current'] = f"{config.total_input_current:.2f}"
                    row_data['Breaker Size'] = config.get_display_breaker_size()
                
                combiner_data.append(row_data)
            
            # Add empty row between combiner boxes (except after last one)
            if combiner_id != sorted(device_configs.keys())[-1]:
                combiner_data.append({})
        
        return pd.DataFrame(combiner_data)
    
    def count_combiner_boxes(self) -> int:
        """Count the number of combiner boxes in the project"""
        count = 0
        for block in self.blocks.values():
            if hasattr(block, 'device_type'):
                from ..models.block import DeviceType
                if block.device_type == DeviceType.COMBINER_BOX:
                    count += 1
        return count
    
    def _format_combiner_sheet(self, worksheet, data: pd.DataFrame):
        """
        Format the combiner box sheet with special handling for cable mismatches
        
        Args:
            worksheet: openpyxl worksheet
            data: DataFrame with combiner data
        """
        # Use standard formatting first
        self._format_excel_sheet(worksheet, data)
        
        # Additional formatting for cable size mismatches
        red_font = Font(color="FF0000")
        
        # Find Cable Size column
        cable_col = None
        for col in range(1, worksheet.max_column + 1):
            if worksheet.cell(row=1, column=col).value == 'Cable Size':
                cable_col = col
                break
        
        if cable_col:
            # Check each cable size cell for mismatches
            # This would need integration with actual mismatch detection
            # For now, we'll leave it as a placeholder for future enhancement
            pass