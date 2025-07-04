import pandas as pd
import json
import os
from typing import Dict, List, Any, Optional
from ..models.block import BlockConfig, WiringType
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


class BOMGenerator:
    """Utility class for generating Bill of Materials from block configurations"""
    
    # Constants for BOM calculations
    CABLE_WASTE_FACTOR = 1.05  # 5% extra for waste/installation
    
    def __init__(self, blocks: Dict[str, BlockConfig]):
        """
        Initialize BOM Generator with block configurations
        
        Args:
            blocks: Dictionary of block configurations (id -> BlockConfig)
        """
        self.blocks = blocks
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
                        'description': f'DC Positive Whip Cable {whip_cable_size}',
                        'quantity': whip_length_feet,
                        'unit': 'feet',
                        'category': 'eBOS'
                    }
                
                if 'whip_cable_negative' in cable_lengths:
                    whip_length_feet = round(cable_lengths['whip_cable_negative'] * 3.28084 * self.CABLE_WASTE_FACTOR, 1)

                    block_quantities[f'Negative Whip Cable ({whip_cable_size})'] = {
                        'description': f'DC Negative Whip Cable {whip_cable_size}',
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
                    
                    block_quantities[f'Positive {string_count}-String Harness'] = {
                        'description': f'Positive {string_count}-String Harness ({harness_cable_size} trunk, {string_cable_size} drops)',
                        'quantity': count,
                        'unit': 'units',
                        'category': 'eBOS'
                    }
                    
                    block_quantities[f'Negative {string_count}-String Harness'] = {
                        'description': f'Negative {string_count}-String Harness ({harness_cable_size} trunk, {string_cable_size} drops)',
                        'quantity': count,
                        'unit': 'units',
                        'category': 'eBOS'
                    }
                
                # Count fuses by rating
                fuse_counts = self._count_fuses_by_rating(block)
                for rating, count in fuse_counts.items():
                    block_quantities[f'DC String Fuse {rating}A'] = {
                        'description': f'DC String Fuse {rating}A',
                        'quantity': count,
                        'unit': 'units',
                        'category': 'eBOS'
                    }
                
                # Split whip cables by polarity
                whip_cable_size = getattr(block.wiring_config, 'whip_cable_size', '6 AWG')  # Default to 6 AWG if not specified
                
                if 'whip_cable_positive' in cable_lengths:
                    whip_length_feet = round(cable_lengths['whip_cable_positive'] * 3.28084 * self.CABLE_WASTE_FACTOR, 1)

                    block_quantities[f'Positive Whip Cable ({whip_cable_size})'] = {
                        'description': f'DC Positive Whip Cable {whip_cable_size}',
                        'quantity': whip_length_feet,
                        'unit': 'feet',
                        'category': 'eBOS'
                    }
                
                if 'whip_cable_negative' in cable_lengths:
                    whip_length_feet = round(cable_lengths['whip_cable_negative'] * 3.28084 * self.CABLE_WASTE_FACTOR, 1)

                    block_quantities[f'Negative Whip Cable ({whip_cable_size})'] = {
                        'description': f'DC Negative Whip Cable {whip_cable_size}',
                        'quantity': whip_length_feet,
                        'unit': 'feet',
                        'category': 'eBOS'
                    }

                # Split extender cable by polarity
                extender_cable_size = getattr(block.wiring_config, 'extender_cable_size', '8 AWG')

                if 'extender_cable_positive' in cable_lengths:
                    extender_length_feet = round(cable_lengths['extender_cable_positive'] * 3.28084 * self.CABLE_WASTE_FACTOR, 1)

                    block_quantities[f'Positive Extender Cable ({extender_cable_size})'] = {
                        'description': f'DC Positive Extender Cable {extender_cable_size}',
                        'quantity': extender_length_feet,
                        'unit': 'feet',
                        'category': 'Extender Cables'
                    }

                if 'extender_cable_negative' in cable_lengths:
                    extender_length_feet = round(cable_lengths['extender_cable_negative'] * 3.28084 * self.CABLE_WASTE_FACTOR, 1)

                    block_quantities[f'Negative Extender Cable ({extender_cable_size})'] = {
                        'description': f'DC Negative Extender Cable {extender_cable_size}',
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

    
    def export_bom_to_excel(self, filepath: str, project_info: Optional[Dict[str, Any]] = None, 
                       checked_components: Optional[List[Dict]] = None) -> bool:
        """
        Export BOM to Excel file
        
        Args:
            filepath: Path to save the Excel file
            project_info: Optional dictionary with project information
            
        Returns:
            True if export successful, False otherwise
        """
        writer = None
        try:
            # Calculate quantities
            quantities = self.calculate_cable_quantities()
            
            # Generate summary and detailed data
            summary_data = self.generate_summary_data(quantities)
            detailed_data = self.generate_detailed_data(quantities)

            # Filter data based on checked components if provided
            if checked_components:
                summary_data = self.filter_data_by_checked_components(summary_data, checked_components)
                detailed_data = self.filter_data_by_checked_components(detailed_data, checked_components, is_detailed=True)

            # Check for missing blocks
            all_block_ids = set(self.blocks.keys())
            blocks_with_quantities = set(quantities.keys())
            missing_blocks = all_block_ids - blocks_with_quantities
            
            if missing_blocks:
                for block_id in sorted(missing_blocks):
                    # Add empty entries for missing blocks to make them visible
                    if block_id not in quantities:
                        quantities[block_id] = {
                            "Missing Data": {
                                'description': "Block data missing in BOM generation",
                                'quantity': 0,
                                'unit': 'note',
                                'category': 'Warnings'
                            }
                        }
            
            # Re-generate data with missing blocks included
            if missing_blocks:
                detailed_data = self.generate_detailed_data(quantities)
            
            # Generate project info if not provided
            if project_info is None:
                project_info = self.generate_project_info()
            
            # Create Excel writer
            writer = pd.ExcelWriter(filepath, engine='openpyxl')
            
            # Write summary data
            summary_data.to_excel(writer, sheet_name='BOM Summary', index=False, startrow=10)  # Start after project info

            # Write detailed data  
            detailed_data.to_excel(writer, sheet_name='Block Details', index=False)
            
            # Get workbook
            workbook = writer.book
            
            # Get summary worksheet
            summary_sheet = writer.sheets['BOM Summary']
            
            # Add project info to summary sheet
            row = 1
            summary_sheet.merge_cells(f'A{row}:E{row}')
            summary_sheet.cell(row=row, column=1, value="Project Information").font = Font(bold=True, size=14)
            
            if project_info:
                for i, (key, value) in enumerate(project_info.items()):
                    # Format system size to 2 decimal places
                    if key == 'System Size (kW DC)' and isinstance(value, (int, float)):
                        value = round(value, 2)
                        
                    row = i + 2
                    summary_sheet.cell(row=row, column=1, value=key).font = Font(bold=True)
                    summary_sheet.cell(row=row, column=2, value=value)
            else:
                # Generate project info from blocks if not provided
                generated_info = self.generate_project_info()
                for i, (key, value) in enumerate(generated_info.items()):
                    # Format system size to 2 decimal places
                    if key == 'System Size (kW DC)' and isinstance(value, (int, float)):
                        value = round(value, 2)
                        
                    row = i + 2
                    summary_sheet.cell(row=row, column=1, value=key).font = Font(bold=True)
                    summary_sheet.cell(row=row, column=2, value=str(value))  
            
            # Add section header for BOM
            row = 9
            summary_sheet.merge_cells(f'A{row}:E{row}')
            summary_sheet.cell(row=row, column=1, value="Bill of Materials").font = Font(bold=True, size=14)
            
            # Format summary sheet
            self._format_excel_sheet(workbook['BOM Summary'], summary_data, start_row=10)
            
            # Format detailed sheet
            self._format_excel_sheet(workbook['Block Details'], detailed_data)
            
            # Add filter to summary sheet for easier sorting
            summary_sheet.auto_filter.ref = f"A10:E{10 + len(summary_data)}"
            
            # Save the Excel file
            writer.close()
            writer = None  # Set to None after closing to avoid double close in finally block
            
            # Open the Excel file (Windows specific)
            try:
                os.startfile(filepath)
            except Exception as e:
                print(f"File was saved but could not be opened automatically: {str(e)}")

            return True
        except Exception as e:
            print(f"Error exporting BOM: {str(e)}")
            # Re-raise permission errors to be caught by the calling function
            if isinstance(e, PermissionError):
                raise
            return False
        finally:
            # Make sure the writer is closed
            if writer is not None:
                try:
                    writer.close()
                except:
                    pass

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
            summary_data.to_excel(writer, sheet_name='BOM Summary', index=False, startrow=10)
            
            # Write detailed data  
            detailed_data.to_excel(writer, sheet_name='Block Details', index=False)
            
            # Format the sheets (same as original method)
            workbook = writer.book
            summary_sheet = writer.sheets['BOM Summary']
            
            # Add project info to summary sheet
            row = 1
            summary_sheet.merge_cells(f'A{row}:F{row}')
            summary_sheet.cell(row=row, column=1, value="Project Information").font = Font(bold=True, size=14)
            
            if project_info:
                for i, (key, value) in enumerate(project_info.items()):
                    if key == 'System Size (kW DC)' and isinstance(value, (int, float)):
                        value = round(value, 2)
                    row = i + 2
                    summary_sheet.cell(row=row, column=1, value=key).font = Font(bold=True)
                    summary_sheet.cell(row=row, column=2, value=value)
            
            # Add section header for BOM
            row = 9
            summary_sheet.merge_cells(f'A{row}:F{row}')
            summary_sheet.cell(row=row, column=1, value="Bill of Materials").font = Font(bold=True, size=14)
            
            # Format sheets
            self._format_excel_sheet(workbook['BOM Summary'], summary_data, start_row=10)
            self._format_excel_sheet(workbook['Block Details'], detailed_data)
            
            # Add filter
            summary_sheet.auto_filter.ref = f"A10:F{10 + len(summary_data)}"
            
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
        
        # Category styles
        category_fill = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")
        
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
            for row_num in range(start_row + 1, start_row + len(data) + 1):
                cell = worksheet.cell(row=row_num, column=part_number_col)
                cell.alignment = centered_alignment
        
        # Format data cells and apply auto-width
        current_category = None
        
        for i, row in enumerate(worksheet.iter_rows(min_row=start_row+1, 
                                                    max_row=start_row+len(data), 
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
            
            # Check if category has changed
            if 'Category' in data.columns:
                category_col_idx = list(data.columns).index('Category')
                category = worksheet.cell(row=row_index, column=category_col_idx + 1).value
                
                if category != current_category:
                    current_category = category
                    # Add subtle highlight for category rows
                    for cell in row:
                        cell.fill = category_fill
            
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
            'Inverter Model': 'Unknown'
        }
        
        # Calculate total system size and module count
        total_modules = 0
        module_manufacturer = set()
        module_model = set()
        inverter_manufacturer = set()
        inverter_model = set()
        dc_collection_types = set()
        
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
            
            # Add inverter info
            if block.inverter:
                inverter_manufacturer.add(block.inverter.manufacturer)
                inverter_model.add(block.inverter.model)
            
            # Add DC collection type
            if block.wiring_config:
                dc_collection_types.add(block.wiring_config.wiring_type.value)
        
        # Update info dict
        info['Number of Modules'] = total_modules
        info['Module Manufacturer'] = ', '.join(module_manufacturer) if module_manufacturer else 'Unknown'
        info['Module Model'] = ', '.join(module_model) if module_model else 'Unknown'
        info['Inverter Manufacturer'] = ', '.join(inverter_manufacturer) if inverter_manufacturer else 'Unknown'
        info['Inverter Model'] = ', '.join(inverter_model) if inverter_model else 'Unknown'
        info['DC Collection'] = ', '.join(dc_collection_types) if dc_collection_types else 'Unknown'
        
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
                    whip_segments_pos.append(segment_length_feet)
                elif "neg_dev" in route_id or "neg_main" in route_id or "whip_neg" in route_id or "neg_whip" in route_id:
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

            # If over 300ft, return custom part number
            if target_length > 300:
                return f"EXT-CUSTOM-{target_length}"
            
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

            # If over 300ft, return custom part number
            if target_length > 300:
                return f"WHI-CUSTOM-{target_length}"
            
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
            
            block_quantities[segment_key] = {
                'description': f"{int(length)}ft {prefix} Segment ({cable_size})",
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

            block_quantities[total_key] = {
                'description': f'DC {prefix} {cable_size}',
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