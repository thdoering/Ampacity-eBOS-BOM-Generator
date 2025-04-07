import pandas as pd
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
            
            quantities[block_id] = block_quantities
        
        # Store the original quantities before segment analysis
        original_quantities = {}
        for block_id, block_quantities in quantities.items():
            original_quantities[block_id] = {k: v.copy() for k, v in block_quantities.items() if 'Whip Cable' in k and v['unit'] == 'feet'}
        
        quantities = self.analyze_wire_segments(quantities)
        
        # Restore total whip cable entries that might have been overwritten
        for block_id, original_block_quantities in original_quantities.items():
            if block_id in quantities:
                for item_key, item_details in original_block_quantities.items():
                    quantities[block_id][item_key] = item_details

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
            
        # Count strings per tracker
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
                'Description': description,
                'Quantity': round(quantity, 1) if unit == 'feet' else int(quantity),
                'Unit': unit
            })
        
        # Sort by category then component type
        summary_data = sorted(summary_data, key=lambda x: (x['Category'], x['Component Type']))
        
        return pd.DataFrame(summary_data)
    
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
                
                detailed_data.append({
                    'Block': block_id,
                    'Category': category,
                    'Component Type': component_type,
                    'Description': description,
                    'Quantity': round(quantity, 1) if unit == 'feet' else int(quantity),
                    'Unit': unit
                })
        
        # Sort by block ID, category, and component type
        detailed_data = sorted(detailed_data, key=lambda x: (x['Block'], x['Category'], x['Component Type']))
        
        return pd.DataFrame(detailed_data)

    
    def export_bom_to_excel(self, filepath: str, project_info: Optional[Dict[str, Any]] = None) -> bool:
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
                if "pos_src" in route_id or "pos_node" in route_id:
                    string_segments_pos.append(segment_length_feet)
                elif "neg_src" in route_id or "neg_node" in route_id:
                    string_segments_neg.append(segment_length_feet)
                elif "pos_dev" in route_id or "pos_main" in route_id:
                    whip_segments_pos.append(segment_length_feet)
                elif "neg_dev" in route_id or "neg_main" in route_id:
                    whip_segments_neg.append(segment_length_feet)
            
            # Process whip segments
            whip_size = getattr(block.wiring_config, 'whip_cable_size', "8 AWG")
            self._add_segment_analysis(block_quantities, whip_segments_pos, 
                                    whip_size, "Positive Whip Cable", 1)
            self._add_segment_analysis(block_quantities, whip_segments_neg, 
                                    whip_size, "Negative Whip Cable", 1)
            
            # Calculate and add total entries from segments
            self.calculate_totals_from_segments(block_quantities, whip_size, "Positive Whip Cable")
            self.calculate_totals_from_segments(block_quantities, whip_size, "Negative Whip Cable")
                    
            # Update quantities
            quantities[block_id] = block_quantities
        
        return quantities

    def _add_segment_analysis(self, block_quantities, segments, cable_size, 
                       prefix, length_increment=5):
        """Add segment analysis for a specific cable type"""
        
        if not segments:
            return
            
        # Add waste factor
        segments = [s * self.CABLE_WASTE_FACTOR for s in segments]
        
        # Round up to nearest increment
        rounded_segments = [length_increment * ((s + length_increment - 0.1) // length_increment + 1) 
                        for s in segments]
        
        # Count segments by length
        segment_counts = {}
        for length in rounded_segments:
            if length not in segment_counts:
                segment_counts[length] = 0
            segment_counts[length] += 1
            
        # Add segment counts to quantities
        for length, count in segment_counts.items():
            segment_key = f"{prefix} Segment {int(length)}ft ({cable_size})"
            block_quantities[segment_key] = {
                'description': f"{int(length)}ft {prefix} Segment ({cable_size})",
                'quantity': count,
                'unit': 'segments',
                'category': 'eBOS Segments'
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
            total_key = f"{prefix} ({cable_size})"
            block_quantities[total_key] = {
                'description': f'DC {prefix} {cable_size}',
                'quantity': round(total_length, 1),
                'unit': 'feet',
                'category': 'eBOS'
            }
            
        return total_length > 0